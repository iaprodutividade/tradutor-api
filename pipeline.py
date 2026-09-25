"""
Prototipo de validacao (Fase 1) do pipeline do Tradutor.
Extrai blocos de texto de um PDF (posicao, fonte, tamanho, cor), traduz em lote
via OpenAI gpt-5-mini preservando numeros/codigos/unidades, remove o texto
original (redaction, sem tocar em imagem/vetor) e reinsere o texto traduzido
na mesma caixa com auto-ajuste de tamanho de fonte.

Tambem trata .docx (python-docx), traduzindo paragrafos e celulas de tabela
mantendo a formatacao nativa do Word.
"""
import base64
import html
import io
import json
import os
import subprocess
import tempfile
from collections import Counter
from pathlib import Path
from typing import Callable

import fitz  # pymupdf
from docx import Document
from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image

load_dotenv(Path(__file__).parent / ".env")

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])
MODEL = "gpt-4o-mini"

# Preco por token do modelo em uso — conferido na OpenAI em 23/09/2026
# (https://openai.com/api/pricing): US$0,15 / 1M tokens de entrada,
# US$0,60 / 1M tokens de saida. Cotacao do dolar configuravel via env (o
# valor exato do extrato do cartao varia dia a dia — isso e uma
# aproximacao do custo real, nao o valor exato faturado).
PRECO_INPUT_USD_POR_1M = 0.15
PRECO_OUTPUT_USD_POR_1M = 0.60
USD_BRL_TAXA = float(os.environ.get("USD_BRL_TAXA", "5.15"))


def custo_centavos_brl(prompt_tokens: int, completion_tokens: int) -> int:
    custo_usd = (prompt_tokens / 1_000_000) * PRECO_INPUT_USD_POR_1M + (
        completion_tokens / 1_000_000
    ) * PRECO_OUTPUT_USD_POR_1M
    return round(custo_usd * USD_BRL_TAXA * 100)

SYSTEM_PROMPT = """Voce e um tradutor tecnico. Traduza cada item da lista do idioma de \
origem para o idioma de destino, mantendo o significado tecnico exato.

Regras obrigatorias:
- Preserve EXATAMENTE como estao, sem traduzir nem converter: numeros, percentuais, \
codigos (CAS, ONU, NCM), siglas, unidades de medida (mg, g, kg, km, %, degC etc.) e \
nomes de produto/marca.
- Atencao especial: um nome curto entre aspas logo depois de "nosso", "nossa", \
"chamado(a)", "conhecido(a) como" ou similar e quase sempre nome proprio (marca, \
mascote, produto) mesmo que as palavras, isoladas, tenham traducao literal obvia (ex.: \
"nosso \"Gato Mia\"" preserva "Gato Mia" tal como esta, NAO traduz pra "Cat Mia" ou \
"Cat Meow" so porque "gato" e "mia" tem traducao).
- Cada item da lista e INDEPENDENTE. A presenca de nomes de marca/produto em outros \
itens do MESMO lote nao significa que este item tambem deva ficar sem traduzir — \
traduza normalmente qualquer item que seja frase, chamada ou slogan de verdade, mesmo \
que esteja do lado de nomes de marca, telefone ou site no lote (ex.: um lote com \
"GATO MIA", "Naturalle", "Amor em cada detalhe" e um telefone: os tres primeiros sao \
nome/marca e ficam como estao, mas "Amor em cada detalhe" e slogan de verdade e DEVE \
ser traduzido).
- Nao adicione nem remova informacao. Nao resuma. Nao comente.
- Mantenha quebras de linha internas do item quando fizerem sentido.
- Se um item nao tiver texto traduzivel (so numero, so pontuacao, vazio), devolva o \
item inalterado.

Devolva APENAS um JSON com a chave "traducoes": lista de strings, na MESMA ORDEM e \
MESMA QUANTIDADE da lista de entrada."""


def _call_translate(texts: list[str], source_lang: str, target_lang: str) -> tuple[list[str], dict]:
    user_payload = {
        "idioma_origem": source_lang,
        "idioma_destino": target_lang,
        "quantidade_itens": len(texts),
        "itens": texts,
    }
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ],
        response_format={"type": "json_object"},
    )
    data = json.loads(resp.choices[0].message.content)
    usage = {
        "prompt_tokens": resp.usage.prompt_tokens if resp.usage else 0,
        "completion_tokens": resp.usage.completion_tokens if resp.usage else 0,
    }
    return data["traducoes"], usage


CLASSIFICACAO_MARCA_SYSTEM_PROMPT = """Voce recebe uma lista de textos curtos que se \
repetem identicos em varias paginas de um documento (candidatos a logo/elemento de marca, \
detectados automaticamente por repeticao). Para cada um, classifique como:

- "marca": nome de empresa, produto, marca, slogan que funciona como parte da identidade \
visual/logotipo (ex: nome do negocio, "Est. 1998" como selo), ou qualquer texto que nao \
faria sentido aparecer traduzido porque é uma marca registrada ou nome proprio.
- "conteudo": frase, chamada, slogan ou texto comum que, mesmo se repetindo em toda \
pagina como parte do design, é conteudo de verdade e deveria ser traduzido junto com o \
resto do documento (ex: um slogan/tagline motivacional, uma chamada de rodape).

Na duvida entre as duas, prefira "marca" (mais seguro nao traduzir por engano um nome \
proprio do que arriscar deixar sem traduzir um texto comum).

Devolva APENAS um JSON com a chave "classificacoes": lista de strings ("marca" ou \
"conteudo"), na MESMA ORDEM e MESMA QUANTIDADE da lista de entrada."""


def _classificar_marca_ou_conteudo(
    textos: list[str], on_uso: Callable[[dict], None] | None = None
) -> list[bool]:
    """True = proteger (é marca/nome, não traduzir), False = é conteúdo comum (traduzir
    normalmente). Achado num caso real: "GATO MIA" (nome) e "AMOR EM CADA DETALHE"
    (slogan) apareciam juntos, repetidos em toda página — os dois batiam nos mesmos
    critérios de "elemento repetido" (curto, poucas palavras), então os dois ficavam
    protegidos, mesmo o slogan sendo conteúdo de verdade que devia ser traduzido. Sem
    entender o TEXTO (não só a repetição), não dá pra diferenciar os dois casos."""
    if not textos:
        return []
    try:
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": CLASSIFICACAO_MARCA_SYSTEM_PROMPT},
                {"role": "user", "content": json.dumps({"itens": textos}, ensure_ascii=False)},
            ],
            response_format={"type": "json_object"},
        )
        if on_uso and resp.usage:
            on_uso({"prompt_tokens": resp.usage.prompt_tokens, "completion_tokens": resp.usage.completion_tokens})
        data = json.loads(resp.choices[0].message.content)
        classificacoes = data.get("classificacoes", [])
        if len(classificacoes) != len(textos):
            raise ValueError(f"esperava {len(textos)} classificacoes, recebi {len(classificacoes)}")
        return [c == "marca" for c in classificacoes]
    except Exception as e:
        # Falha na classificacao nunca deve derrubar a deteccao de logo --
        # comportamento antigo (protege tudo que repete) como fallback seguro.
        print(f"[aviso] falha ao classificar marca/conteudo, protegendo tudo por seguranca: {e}")
        return [True] * len(textos)


def translate_batch(
    texts: list[str],
    source_lang: str,
    target_lang: str,
    on_uso: Callable[[dict], None] | None = None,
) -> list[str]:
    if not texts:
        return []

    for attempt in range(2):
        out, usage = _call_translate(texts, source_lang, target_lang)
        # Toda chamada de verdade custa dinheiro, inclusive a que vai ser
        # descartada por causa de divergencia de contagem — por isso o
        # aviso ao chamador acontece aqui dentro, nao so no caminho feliz.
        if on_uso:
            on_uso(usage)
        if len(out) == len(texts):
            return out
        print(f"[aviso] tentativa {attempt + 1}: esperava {len(texts)} traducoes, recebi {len(out)} — repetindo")

    # Ainda inconsistente apos repetir em lote: traduz item por item (1 por
    # chamada). Mais lento/caro so pra esses poucos itens problematicos, mas
    # garante correspondencia 1:1 -- o fallback antigo ("out +
    # texts[len(out):]") assumia que o item que faltou era sempre o
    # ULTIMO da lista; quando o modelo funde/pula um item no MEIO, tudo
    # depois dele desalinha e texto original vaza sem traducao, em
    # silencio, mesmo em paginas que pareciam ok. Achado real na pagina 1
    # do "Gato Mia" (ver claude-sessions-log/sessions/2026-09-24_tradutor-
    # pagamento-corrigido-upload-storage-e-classificador-marca.md).
    print("[aviso] mantendo divergencia de contagem apos repetir em lote — traduzindo item por item")
    resultado = []
    for texto in texts:
        item_out, usage_item = _call_translate([texto], source_lang, target_lang)
        if on_uso:
            on_uso(usage_item)
        resultado.append(item_out[0] if item_out else texto)
    return resultado


def _parece_espacamento_artificial(texto: str) -> bool:
    """Deteta fonte de titulo com tracking largo (comum em capa/ficha
    tecnica feita no Canva) que grava espaco real entre CADA letra no PDF,
    nao so um efeito visual de kerning -- extracao normal devolve "N o s s
    o" em vez de "Nosso". Sem tratar isso antes de traduzir, o texto nao
    parece prosa e o modelo devolve inalterado (sem avisar ninguem).
    Achado real no arquivo "Gato Mia" (titulo de capa e ficha tecnica), ver
    claude-sessions-log/sessions/2026-09-25_tradutor-*.md. True quando a
    maioria dos "tokens" (separados por espaco simples) tem so 1
    caractere."""
    tokens = texto.split(" ")
    if len(tokens) < 4:
        return False
    tokens_1_char = sum(1 for t in tokens if len(t) == 1)
    return tokens_1_char / len(tokens) > 0.6


def _colapsar_espacamento_artificial(texto: str) -> str:
    """Fallback ingenuo: cola tudo, sem tentar recuperar onde ficavam as
    palavras de verdade (usado quando nao da pra checar a posicao real dos
    caracteres, ver _texto_traduzivel). Espaco DUPLO continua respeitado
    como separador de palavra de verdade, caso a mesma fonte tambem use
    espaco duplo entre palavras reais."""
    return texto.replace("  ", "\x00").replace(" ", "").replace("\x00", " ")


LIMIAR_GAP_ESPACO_REAL_PT = 1.5


def _reconstruir_espacamento_por_posicao(chars: list[dict]) -> str:
    """Parte pura (sem PDF) de _texto_traduzivel: recebe caracteres no
    formato do rawdict do PyMuPDF (cada um {"c": <char>, "bbox": (x0, y0,
    x1, y1)}) e reconstroi onde ficavam os espacos de palavra de verdade
    usando o GAP horizontal entre caracteres visiveis consecutivos.

    Limiar adaptativo: fontes diferentes tem magnitude de tracking bem
    diferente entre si (medido na pratica: uma fonte com gap ~0pt entre
    letras da mesma palavra e ~6pt no espaco de palavra; outra com gap
    ~6pt uniforme entre TODAS as letras, onde so o espaco de palavra de
    verdade destoa, ~22pt). Um limiar fixo funciona pra uma e falha pra
    outra -- compara cada gap com a MEDIANA dos gaps da propria linha
    (aproxima o tracking normal daquela fonte) em vez de um valor
    absoluto. Devolve None quando nao ha caracteres visiveis suficientes
    pra uma mediana confiavel (quem chama decide o fallback)."""
    visiveis = [c for c in chars if c["c"] != " "]
    if len(visiveis) < 4:
        return None
    gaps = [visiveis[i]["bbox"][0] - visiveis[i - 1]["bbox"][2] for i in range(1, len(visiveis))]
    gaps_ordenados = sorted(gaps)
    mediana = gaps_ordenados[len(gaps_ordenados) // 2]
    limiar = max(LIMIAR_GAP_ESPACO_REAL_PT, mediana * 2.5)
    partes = [visiveis[0]["c"]]
    for i, gap in enumerate(gaps, start=1):
        if gap > limiar:
            partes.append(" ")
        partes.append(visiveis[i]["c"])
    return "".join(partes)


def _texto_traduzivel(page: fitz.Page, texto: str, bbox: tuple) -> str:
    """Quando o texto parece ter espacamento artificial
    (_parece_espacamento_artificial), reconstroi a pontuacao real das
    palavras usando a POSICAO de cada caractere (rawdict) em vez de so
    colar tudo (ver _reconstruir_espacamento_por_posicao). Sem isso,
    colar tudo (ex.: "nossaareia" em vez de "nossa areia") as vezes o
    modelo reconstroi sozinho, as vezes nao -- achado real na ficha
    tecnica do arquivo "Gato Mia" (ver claude-sessions-log/sessions/2026-
    09-25_tradutor-*.md)."""
    if not _parece_espacamento_artificial(texto):
        return texto
    chars = []
    for b in page.get_text("rawdict", clip=fitz.Rect(bbox)).get("blocks", []):
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            for span in line["spans"]:
                chars.extend(span["chars"])
    return _reconstruir_espacamento_por_posicao(chars) or _colapsar_espacamento_artificial(texto)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def _is_bold(span: dict) -> bool:
    return "bold" in span["font"].lower() or bool(span.get("flags", 0) & 16)


def _mesclar_linhas_mesma_altura(raw_lines: list[dict]) -> list[dict]:
    """Texto justificado com espacamento largo entre palavras as vezes faz o
    PyMuPDF fragmentar uma unica linha visual em varias entradas de 'line'
    (uma por palavra/grupo), todas na mesma faixa vertical. Sem isso, cada
    palavra vira uma linha de HTML separada (<br> entre cada uma) — quebra o
    paragrafo inteiro. Funde entradas com bbox verticalmente sobreposta,
    preservando a ordem (ja vem da esquerda pra direita)."""
    mescladas: list[dict] = []
    for line in raw_lines:
        y0, y1 = line["bbox"][1], line["bbox"][3]
        if mescladas:
            atual = mescladas[-1]
            ay0, ay1 = atual["bbox"][1], atual["bbox"][3]
            sobreposicao = min(y1, ay1) - max(y0, ay0)
            altura_min = min(y1 - y0, ay1 - ay0)
            if altura_min > 0 and sobreposicao / altura_min > 0.5:
                # Garante espaco na juncao: essa funcao serve pra reunir
                # fragmentos da MESMA linha visual (letra solta por causa
                # de tracking largo), mas o limiar de sobreposicao vertical
                # tambem pode confundir duas linhas DIFERENTES e genuinas
                # (ex.: "Mais conforto" / "para as patinhas" em 2 linhas)
                # quando a fonte e grande o bastante pra bbox de uma
                # encostar na de baixo. Sem isso, o span final fica colado
                # ("Mais confortopara as patinhas") -- um espaco extra
                # nunca quebra a traducao, faltar um sempre corrompe.
                # Achado real na ficha tecnica do arquivo "Gato Mia" (ver
                # claude-sessions-log/sessions/2026-09-25_tradutor-*.md).
                if (
                    atual["spans"]
                    and line["spans"]
                    and not atual["spans"][-1]["text"].endswith((" ", "\n"))
                    and not line["spans"][0]["text"].startswith((" ", "\n"))
                ):
                    ultimo = atual["spans"][-1]
                    atual["spans"] = atual["spans"][:-1] + [{**ultimo, "text": ultimo["text"] + " "}]
                atual["spans"] = atual["spans"] + line["spans"]
                atual["bbox"] = (
                    min(atual["bbox"][0], line["bbox"][0]),
                    min(ay0, y0),
                    max(atual["bbox"][2], line["bbox"][2]),
                    max(ay1, y1),
                )
                continue
        mescladas.append({"spans": list(line["spans"]), "bbox": line["bbox"]})
    return mescladas


def _block_lines_runs(block: dict) -> list[tuple[list[list], tuple]]:
    """Por linha, agrupa spans consecutivos com o mesmo estilo (negrito) num run.
    Devolve lista de (runs, bbox) por linha — runs e a lista de [texto, negrito],
    bbox e usado depois pra decidir se a quebra pra proxima linha era so
    word-wrap do original (junta com espaco ao traduzir) ou uma quebra de
    verdade (mantem <br>)."""
    lines_runs = []
    for line in _mesclar_linhas_mesma_altura(block["lines"]):
        runs: list[list] = []
        for span in line["spans"]:
            if not span["text"]:
                continue
            bold = _is_bold(span)
            if runs and runs[-1][1] == bold:
                runs[-1][0] += span["text"]
            else:
                runs.append([span["text"], bold])
        if runs:
            lines_runs.append((runs, line["bbox"]))
    return lines_runs


def _dividir_em_subblocos(
    lines_runs: list[tuple[list[list], tuple]]
) -> list[list[tuple[list[list], tuple]]]:
    """Divide as linhas de um bloco em sub-grupos quando tudo indica que o
    bloco mistura conteudo de duas fileiras/paragrafos diferentes: a linha
    anterior termina em pontuacao final E a linha atual volta bem perto da
    margem esquerda do bloco (nao e uma continuacao com recuo, e sim o
    inicio de um texto novo). Isso acontece em documentos com layout de 2
    colunas sem tabela de verdade (tipo FISPQ/FDS) — o PyMuPDF as vezes
    agrupa o final da frase de uma fileira com o rotulo da fileira seguinte
    num unico bloco, so por estarem geometricamente proximos, sem nenhum
    espaco vertical extra que desse pra detectar isso so pela distancia
    entre linhas. Sem separar, a bbox do bloco fica alta/larga demais e
    mistura o texto errado, causando sobreposicao quando o idioma de
    destino tem tamanho bem diferente do original."""
    if not lines_runs:
        return []
    margem_esquerda = min(bbox[0] for _runs, bbox in lines_runs)
    grupos: list[list[tuple[list[list], tuple]]] = [[lines_runs[0]]]
    for i in range(1, len(lines_runs)):
        runs_anterior, _bbox_anterior = lines_runs[i - 1]
        _runs_atual, bbox_atual = lines_runs[i]
        texto_anterior = "".join(t for t, _ in runs_anterior).strip()
        termina_frase = texto_anterior.rstrip().endswith((".", ":", ";", "!", "?"))
        volta_pra_margem = abs(bbox_atual[0] - margem_esquerda) < 2.0
        if termina_frase and volta_pra_margem:
            grupos.append([])
        grupos[-1].append(lines_runs[i])
    return grupos


_MARCADORES_SOLTOS = {"•", "·", "●", "○", "▪", "‣", "◦", "-", "*"}

def _eh_linha_fina(d: dict) -> bool:
    """Retangulo fino (linha divisoria/separador/borda de tabela de
    verdade), nao um bloco de conteudo -- largura OU altura menor que
    1pt."""
    return d.get("type") in ("s", "f") and (fitz.Rect(d["rect"]).width < 1 or fitz.Rect(d["rect"]).height < 1)


def _tabela_parece_real(tbbox: fitz.Rect, celulas: list[fitz.Rect]) -> bool:
    """False quando alguma celula ocupa a maior parte da propria tabela --
    sinal de grade inventada por find_tables() (fundo decorativo ou
    alinhamento de texto/icones lido como grade), nao tabela de verdade:
    uma celula real nunca domina a tabela inteira sozinha, precisa de
    pelo menos 2 linhas/colunas pra existir. Achado real no arquivo "Gato
    Mia" (ver claude-sessions-log/sessions/2026-09-25_tradutor-*.md) --
    tanto uma celula unica cobrindo a pagina inteira quanto uma grade de
    varias celulas normais sem nenhuma linha de grade desenhada de
    verdade (ver _eh_linha_fina, usado por quem chama antes de sequer
    tentar find_tables())."""
    if not celulas:
        return False
    return not any(c.get_area() > tbbox.get_area() * 0.6 for c in celulas)


def _cor_fundo_real(pix: fitz.Pixmap, bbox: fitz.Rect) -> tuple:
    """Acha a cor de fundo REAL nessa posicao da pagina, amostrando pixel
    do render (renderizado ANTES de qualquer redacao) em vez de tentar
    adivinhar por geometria de forma (area, z-order) qual retangulo/path
    esta "por cima". A heuristica antiga (menor area contendo o centro do
    bloco) causou uma serie de bugs reais no arquivo "Gato Mia": pagina
    inteira preta (celula de tabela fantasma cobrindo um retangulo preto
    decorativo), texto branco invisivel sobre fundo branco escolhido por
    engano, e um logo (arte vetorial complexa com 1000+ curvas) apagado
    por um retangulo branco pequeno enterrado embaixo dele (opaco mas
    coberto por outros elementos — nao dava pra saber isso sem renderizar
    de verdade). Amostrar o pixel renderizado sempre acerta, nao importa
    quantas formas sobrepostas existam. Ver claude-sessions-log/sessions/
    2026-09-25_tradutor-*.md.

    Amostra pontos logo FORA da bbox (nao dentro, pra nao pegar pixel do
    proprio texto original) e usa a cor mais comum entre eles."""
    margem = 3
    meio_x = (bbox.x0 + bbox.x1) / 2
    meio_y = (bbox.y0 + bbox.y1) / 2
    candidatos = [
        (meio_x, bbox.y0 - margem),
        (meio_x, bbox.y1 + margem),
        (bbox.x0 - margem, meio_y),
        (bbox.x1 + margem, meio_y),
    ]
    cores = []
    for x, y in candidatos:
        xi, yi = int(x), int(y)
        if 0 <= xi < pix.width and 0 <= yi < pix.height:
            cores.append(tuple(c / 255 for c in pix.pixel(xi, yi)[:3]))
    if not cores:
        return (1, 1, 1)
    return Counter(cores).most_common(1)[0][0]


def _has_visible_text(lines_runs: list[tuple[list[list], tuple]]) -> bool:
    """Bloco cujos runs sao so espacos em branco nao tem nada visivel pra
    traduzir/redatar — incluir esse bloco so arrisca redatar (pintar de
    branco) por cima de imagem/vetor vizinho que a bbox encoste."""
    return any(text.strip() for line_runs, _bbox in lines_runs for text, _bold in line_runs)


def _largura_segura_x1(bloco_info: dict, outros_blocos_normais: list[dict]) -> float:
    """Acha o x1 seguro pra inserir o texto traduzido de um bloco sem invadir
    a coluna de outro bloco posicionado a direita na mesma faixa vertical.
    Em layouts de 2 colunas sem tabela de verdade (rotulo : valor lado a
    lado, comum em FISPQ/FDS), as bboxes originais as vezes ja se sobrepoem
    horizontalmente — isso e seguro no idioma de origem porque aquele texto
    especifico para antes de chegar la, mas o texto traduzido, sendo mais
    longo, pode nao parar e vazar por cima do texto do bloco vizinho."""
    bbox = fitz.Rect(bloco_info["bbox"])
    x1_seguro = bbox.x1
    for outro in outros_blocos_normais:
        if outro is bloco_info:
            continue
        obbox = fitz.Rect(outro["bbox"])
        if obbox.x0 <= bbox.x0:
            continue  # nao esta a direita deste bloco
        sobreposicao_vertical = min(bbox.y1, obbox.y1) - max(bbox.y0, obbox.y0)
        if sobreposicao_vertical <= 0:
            continue  # nao compartilha faixa vertical, nao tem risco de colisao
        if obbox.x0 < x1_seguro:
            x1_seguro = obbox.x0
    # nunca deixa a caixa vazia/absurdamente estreita — garante ao menos 10pt.
    return max(x1_seguro - 1.5, bbox.x0 + 10)


def _rect_center_inside(bbox, rect: fitz.Rect) -> bool:
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    return rect.contains(fitz.Point(cx, cy))


# Area minima (pt^2) pra um desenho vetorial contar como "conteudo visual de
# verdade" (logo, icone, ilustracao) — abaixo disso e uma linha fina de
# sublinhado/moldura, ruido demais pra alertar.
AREA_MINIMA_DESENHO_PROTEGIDO = 200
# Sobreposicao minima (pt^2) pra considerar que um bloco de texto realmente
# invade a area protegida, e nao so encosta na borda por causa de arredondamento.
SOBREPOSICAO_MINIMA_ALERTA = 4


def _avisar_colisao_com_conteudo_visual(
    page: fitz.Page, block_infos: list[dict], table_areas: list[fitz.Rect]
) -> None:
    """Antes de redatar qualquer coisa, avisa se a bbox de um bloco de texto
    normal (fora de tabela) invade uma imagem ou um desenho grande o
    suficiente pra ser conteudo de verdade. Foi assim que a logo do rodape
    e a borda da tabela foram apagadas por engano — agora fica visivel no
    log ANTES de gerar o arquivo, em vez de descoberto so comparando PDFs.
    Tabelas ficam de fora da checagem porque suas bordas ja sao redesenhadas
    de proposito depois (ver loop de `table_cell_rects` mais abaixo)."""
    protegidos = [fitz.Rect(img["bbox"]) for img in page.get_image_info()]
    for d in page.get_drawings():
        rect = fitz.Rect(d["rect"])
        if rect.get_area() < AREA_MINIMA_DESENHO_PROTEGIDO:
            continue
        if any(_rect_center_inside(tuple(rect), area) for area in table_areas):
            continue  # borda de tabela — tratada a parte, nao e bug
        fill = d.get("fill")
        if fill and all(c > 0.98 for c in fill) and not d.get("color"):
            continue  # retangulo branco de fundo — redatar em branco nao muda nada
        protegidos.append(rect)

    for info in block_infos:
        if info["tabela"]:
            continue
        bbox = fitz.Rect(info.get("bbox_redacao", info["bbox"]))
        for protegido in protegidos:
            sobreposicao = bbox & protegido
            if sobreposicao.is_empty:
                continue
            area = sobreposicao.get_area()
            if area > SOBREPOSICAO_MINIMA_ALERTA:
                print(
                    f"[ALERTA] pagina {page.number}: bloco de texto "
                    f"{tuple(round(v, 1) for v in bbox)} invade area visual protegida "
                    f"{tuple(round(v, 1) for v in protegido)} (sobreposicao {area:.1f}pt^2) "
                    f"— revisar o PDF final antes de confiar no resultado"
                )


def _merged_lines_runs(blocks: list[dict]) -> list[tuple[list[list], tuple]]:
    """Junta as linhas de varios blocos (ex.: todos os blocos dentro de uma
    celula de tabela) numa unica lista de (runs, bbox) por linha."""
    lines_runs = []
    for b in blocks:
        if b.get("type") != 0:
            continue
        lines_runs.extend(_block_lines_runs(b))
    return lines_runs


# Limiar pra decidir se um PDF tem texto real extraivel ou e so imagem
# achatada (catalogo exportado do Canva, digitalizacao, PDF "para
# impressao" que achata tudo etc.). O pipeline abaixo depende de blocos de
# texto reais (bbox) pra redatar/reinserir — sem eles nao ha nada pra
# traduzir, e o resultado sairia visualmente identico ao original, 100% no
# idioma de origem, sem nenhum erro visivel pro sistema (cobraria sem
# entregar). Caso real que motivou isso documentado em
# claude-sessions-log/sessions/2026-09-23_tradutor-pdf-imagem-*.md.
CARACTERES_MINIMOS_TEXTO_PAGINA = 30
RAZAO_MINIMA_PAGINAS_COM_TEXTO = 0.34


def analisar_texto_extraivel(doc: fitz.Document, max_paginas_checar: int = 5) -> dict:
    """Checa as primeiras `max_paginas_checar` páginas e estima se o PDF tem
    camada de texto real. Retorna eh_imagem=True quando a fração de páginas
    com texto de verdade fica abaixo do limiar — sinal forte de PDF achatado
    (imagem pura), não de documento com pouco texto por página."""
    paginas_checar = min(len(doc), max_paginas_checar)
    paginas_com_texto = 0
    for i in range(paginas_checar):
        texto = doc[i].get_text("text").strip()
        if len(texto) >= CARACTERES_MINIMOS_TEXTO_PAGINA:
            paginas_com_texto += 1
    razao = paginas_com_texto / paginas_checar if paginas_checar else 0
    return {
        "paginas_checadas": paginas_checar,
        "paginas_com_texto": paginas_com_texto,
        "eh_imagem": razao < RAZAO_MINIMA_PAGINAS_COM_TEXTO,
    }


# ---------------------------------------------------------------------------
# PDF-imagem (sem texto extraível) — OCR + inpainting + reaproveita o
# insert_htmlbox de process_pdf pra escrever o texto traduzido de volta.
#
# Validado contra um catálogo real (23 páginas, cliente Djarbas/Rosaves) que
# nenhuma ferramenta do mercado conseguia traduzir — ver
# claude-sessions-log/sessions/2026-09-23_tradutor-pdf-imagem-*.md pro
# histórico completo dos testes que levaram a essa arquitetura.
#
# Limitação conhecida, ainda não resolvida: a proteção de logo/marca
# (`areas_protegidas`) precisa ser informada manualmente por página — não
# existe detecção automática de "isto é uma logo" ainda. Sem proteção,
# texto de marca perto de um ícone gráfico pode sair corrompido pelo
# inpainting (foi exatamente o bug encontrado e corrigido nos testes).
# ---------------------------------------------------------------------------

DPI_RENDER_IMAGEM = 150
LIMIAR_SCORE_OCR = 0.5
MARGEM_MASCARA_PX = 5

_ocr_engine = None
_inpaint_model_manager = None


def _get_ocr_engine():
    global _ocr_engine
    if _ocr_engine is None:
        # PaddleOCR (motor nativo do PaddlePaddle) trava com
        # SIGSEGV/SIGABRT em ARM64 dentro de Docker — bug conhecido e
        # ainda sem solução (varios relatos independentes no GitHub do
        # PaddlePaddle, incluindo tentativas de desligar mkldnn como a
        # linha acima fazia, sem efeito). RapidOCR usa os MESMOS modelos
        # PP-OCR, mas via ONNX Runtime — mesma qualidade de
        # reconhecimento, sem o bug de arquitetura. Fica no modelo
        # "small" (padrão): testado "medium" e saiu mais lento (~50s por
        # página vs ~7s) e não mais preciso.
        from rapidocr import RapidOCR

        _ocr_engine = RapidOCR()
    return _ocr_engine


def _get_inpaint_model():
    global _inpaint_model_manager
    if _inpaint_model_manager is None:
        import torch
        from iopaint.model import models as iopaint_models
        from iopaint.model_manager import ModelManager

        # A CLI do iopaint (`iopaint run`) baixa o peso do modelo antes de
        # instanciar o ModelManager — usando a API do jeito direto (sem
        # passar pela CLI), esse passo não acontece sozinho e o
        # ModelManager nem lista "lama" como disponível (só "cv2", que não
        # precisa de download nenhum). Isso só baixa se ainda não tiver
        # os arquivos em cache; idempotente.
        if iopaint_models["lama"].is_erase_model:
            iopaint_models["lama"].download()

        _inpaint_model_manager = ModelManager(name="lama", device=torch.device("cpu"))
    return _inpaint_model_manager


def _ocr_blocos_pagina(imagem_pil) -> list[dict]:
    """Roda o OCR numa imagem de página (PIL) e devolve os blocos de texto
    com posição em pixels, descartando reconhecimentos de baixa confiança
    (ruído de fundo decorativo, geralmente)."""
    import numpy as np

    ocr = _get_ocr_engine()
    resultado = ocr(np.array(imagem_pil.convert("RGB")))
    blocos = []
    if not resultado.txts:
        return blocos
    for texto, score, poly in zip(resultado.txts, resultado.scores, resultado.boxes):
        if score < LIMIAR_SCORE_OCR or not texto.strip():
            continue
        xs = [float(p[0]) for p in poly]
        ys = [float(p[1]) for p in poly]
        blocos.append({"texto": texto, "x0": min(xs), "y0": min(ys), "x1": max(xs), "y1": max(ys)})

    # Achado testando fonte cursiva sobre foto (titulo do catalogo real):
    # o RapidOCR as vezes devolve uma caixa desproporcionalmente alta pra
    # um bloco — capturando o espaco vertical de 2 linhas mas so
    # reconhecendo o texto da primeira. Isso apaga a segunda linha (que
    # ninguem detectou, entao ninguem traduz) sem colocar nada no lugar —
    # sai pior que nao mexer. Descarta blocos com altura muito fora do
    # padrao da pagina (mais seguro deixar o texto original intocado do
    # que arriscar corromper).
    if len(blocos) >= 3:
        alturas = sorted(b["y1"] - b["y0"] for b in blocos)
        mediana = alturas[len(alturas) // 2]
        blocos = [b for b in blocos if (b["y1"] - b["y0"]) <= mediana * 1.8]

    return blocos


def _sobrepoe_1d(a0: float, a1: float, b0: float, b1: float) -> float:
    return max(0.0, min(a1, b1) - max(a0, b0))


def _gap_1d(a0: float, a1: float, b0: float, b1: float) -> float:
    """Distância real entre dois intervalos 1D — 0 se sobrepõem, positiva
    caso contrário, não importa qual dos dois vem primeiro no eixo. Um
    gap_x = b.x0 - a.x1 simples dá errado (fica negativo, "parece perto")
    quando o segundo bloco está à ESQUERDA do primeiro — foi um bug real
    encontrado ao validar isso, misturava colunas vizinhas numa só."""
    if a1 < b0:
        return b0 - a1
    if b1 < a0:
        return a0 - b1
    return 0.0


def _agrupar_em_linhas(blocos: list[dict]) -> list[dict]:
    """Funde fragmentos do OCR que estão na mesma faixa vertical E
    fisicamente próximos no eixo x numa linha só, da esquerda pra direita.
    Sem a checagem de x, duas colunas lado a lado na mesma altura viravam
    uma 'linha' só, esticada pela página inteira (mesmo bug do parágrafo
    acima, na direção horizontal)."""
    if not blocos:
        return []
    ordenados = sorted(blocos, key=lambda b: (b["y0"], b["x0"]))
    altura_media = sum(b["y1"] - b["y0"] for b in blocos) / len(blocos)
    gap_maximo_x = altura_media * 1.2

    linhas: list[dict] = []
    for b in ordenados:
        encaixou = False
        for linha in linhas:
            sobreposicao = _sobrepoe_1d(b["y0"], b["y1"], linha["y0"], linha["y1"])
            menor_altura = min(b["y1"] - b["y0"], linha["y1"] - linha["y0"])
            gap_x = _gap_1d(b["x0"], b["x1"], linha["x0"], linha["x1"])
            if sobreposicao > menor_altura * 0.4 and gap_x < gap_maximo_x:
                linha["itens"].append(b)
                linha["x0"] = min(linha["x0"], b["x0"])
                linha["y0"] = min(linha["y0"], b["y0"])
                linha["x1"] = max(linha["x1"], b["x1"])
                linha["y1"] = max(linha["y1"], b["y1"])
                encaixou = True
                break
        if not encaixou:
            linhas.append({"itens": [b], "x0": b["x0"], "y0": b["y0"], "x1": b["x1"], "y1": b["y1"]})

    for linha in linhas:
        linha["itens"].sort(key=lambda b: b["x0"])
        linha["texto"] = " ".join(i["texto"] for i in linha["itens"])
    linhas.sort(key=lambda l: l["y0"])
    return linhas


FRACAO_MAX_ALTURA_PARAGRAFO = 0.22


def _agrupar_linhas_em_paragrafos(linhas: list[dict], altura_pagina_px: float | None = None) -> list[dict]:
    """Une linhas em parágrafos/colunas: precisa estar verticalmente perto
    E ter sobreposição horizontal real com a linha vizinha (mesma coluna)
    — evita fundir colunas lado a lado, que não se sobrepõem em x. Título
    normalmente sai isolado porque o espaço até o corpo do texto é maior
    que o limiar.

    altura_pagina_px, se informado, limita a altura de um parágrafo a
    FRACAO_MAX_ALTURA_PARAGRAFO da página inteira -- sem isso, um cardápio
    ou lista de preços com pouco espaço entre seções (ex: "Entrada" /
    "Salgados" / "Jantar" um embaixo do outro, mesmo espaçamento de linha
    dentro e entre as seções) funde o documento inteiro num parágrafo só.
    Visto na prática: 37 blocos de OCR viraram 1 parágrafo cobrindo 87% da
    imagem, e o LaMa (sem contexto de fundo sobrando pra copiar) devolveu a
    área inteira em branco, apagando cor de fundo e logo do cardápio."""
    if not linhas:
        return []
    altura_media = sum(l["y1"] - l["y0"] for l in linhas) / len(linhas)
    gap_maximo_y = altura_media * 0.9
    altura_maxima_paragrafo = (
        altura_pagina_px * FRACAO_MAX_ALTURA_PARAGRAFO if altura_pagina_px else float("inf")
    )

    grupos = [dict(l, linhas=[l["texto"]]) for l in linhas]

    mudou = True
    while mudou:
        mudou = False
        grupos.sort(key=lambda g: g["y0"])
        for i in range(len(grupos)):
            if grupos[i] is None:
                continue
            for j in range(i + 1, len(grupos)):
                if grupos[j] is None:
                    continue
                a, b = grupos[i], grupos[j]
                if _gap_1d(a["y0"], a["y1"], b["y0"], b["y1"]) > gap_maximo_y:
                    continue
                largura_menor = min(a["x1"] - a["x0"], b["x1"] - b["x0"])
                if _sobrepoe_1d(a["x0"], a["x1"], b["x0"], b["x1"]) < largura_menor * 0.3:
                    continue
                nova_altura = max(a["y1"], b["y1"]) - min(a["y0"], b["y0"])
                if nova_altura > altura_maxima_paragrafo:
                    continue
                a["linhas"] += b["linhas"]
                a["x0"], a["y0"] = min(a["x0"], b["x0"]), min(a["y0"], b["y0"])
                a["x1"], a["y1"] = max(a["x1"], b["x1"]), max(a["y1"], b["y1"])
                grupos[j] = None
                mudou = True
                break
            if mudou:
                break
        grupos = [g for g in grupos if g is not None]

    resultado = [
        {"texto": " ".join(g["linhas"]), "x0": g["x0"], "y0": g["y0"], "x1": g["x1"], "y1": g["y1"]}
        for g in grupos
    ]
    resultado.sort(key=lambda b: (b["y0"], b["x0"]))
    return resultado


def _bloco_protegido(bloco: dict, areas_protegidas: list[tuple[float, float, float, float]]) -> bool:
    for x0, y0, x1, y1 in areas_protegidas:
        if bloco["x0"] >= x0 and bloco["x1"] <= x1 and bloco["y0"] >= y0 and bloco["y1"] <= y1:
            return True
    return False


def _limpar_fundo(imagem_pil, blocos_para_apagar: list[dict]):
    """Roda o LaMa pra apagar/reconstruir o fundo sob os blocos de texto
    (exceto os protegidos, já filtrados por quem chama). Devolve uma nova
    imagem PIL."""
    import numpy as np
    from PIL import Image, ImageDraw
    from iopaint.schema import InpaintRequest

    if not blocos_para_apagar:
        return imagem_pil

    mask = Image.new("L", imagem_pil.size, 0)
    draw = ImageDraw.Draw(mask)
    for b in blocos_para_apagar:
        draw.rectangle(
            [b["x0"] - MARGEM_MASCARA_PX, b["y0"] - MARGEM_MASCARA_PX, b["x1"] + MARGEM_MASCARA_PX, b["y1"] + MARGEM_MASCARA_PX],
            fill=255,
        )

    modelo = _get_inpaint_model()
    resultado_bgr = modelo(np.array(imagem_pil.convert("RGB")), np.array(mask), InpaintRequest())
    resultado_rgb = resultado_bgr[:, :, ::-1]
    return Image.fromarray(resultado_rgb)


MAX_PALAVRAS_ELEMENTO_REPETIDO = 4
MIN_CARACTERES_ELEMENTO_REPETIDO = 5  # evita casar conectivo curto ("no", "da", "em")
MARGEM_PROTECAO_PX = 20

# Abaixo desse total de caracteres reconhecidos numa página inteira
# (ex: só uma logo protegida, ou nada de OCR), não cobra a página — não é
# justo cobrar o preço cheio de uma capa que não tinha nada pra traduzir.
MIN_CARACTERES_PAGINA_COBRAVEL = 20


def _normalizar_para_comparacao(texto: str) -> str:
    """Tira espaços e caixa — o OCR não é 100% consistente entre páginas
    com o mesmo elemento gráfico (ex.: leu 'O JEITINHO BEM CASEIRO' com
    espaço numa página e 'OJEITINHO BEM CASEIRO' grudado noutra, mesma
    logo). Comparar só o texto "compactado" resolve isso."""
    return "".join(texto.lower().split())


def detectar_elementos_repetidos(
    doc: fitz.Document,
    indices: list[int],
    min_paginas: int = 2,
    on_progress: Callable[[int, int], None] | None = None,
    on_uso: Callable[[dict], None] | None = None,
) -> tuple[dict[int, list[tuple[float, float, float, float]]], dict[int, list[dict]]]:
    """Detecta blocos de texto curtos (até 4 palavras, pelo menos 5
    caracteres) cujo texto se repete em pelo menos `min_paginas` páginas
    diferentes — sinal forte de logo/marca/rodapé repetido, não conteúdo
    específico da página. Usado pra proteger automaticamente esses blocos
    da limpeza/tradução, sem precisar de coordenada manual por documento.

    Compara só o TEXTO (normalizado, sem espaço), não a posição — a
    primeira versão exigia posição parecida também, mas quebrou no teste
    real: a mesma logo aparece bem maior/centralizada na capa e menor/mais
    à esquerda numa página de conteúdo (templates de página diferentes no
    mesmo documento), então a posição normalizada variava demais (~10-20%)
    pra um limiar apertado funcionar. Risco aceito: um nome próprio
    genuinamente repetido no corpo do texto (não só na logo) também fica
    protegido — melhor deixar uma palavra sem traduzir do que corromper a
    logo de novo (bug real já visto duas vezes nos testes).

    Com menos de `min_paginas` páginas disponíveis não há como comparar —
    devolve vazio (documento de 1 página processado sozinho não tem
    proteção automática de logo ainda; limitação conhecida).

    Devolve também os blocos de OCR crus de cada página já processada
    aqui (blocos_por_pagina), pra quem chamar poder repassar pra
    process_pdf_imagem e evitar rodar o OCR de novo nas mesmas páginas —
    antes essa duplicidade era aceita como custo conhecido; com documentos
    grandes (a detecção roda nas páginas todas do PDF-imagem completo, não
    só numas poucas da prévia) o OCR em dobro passa a ser tempo de verdade
    numa VPS de CPU limitada, então vale reaproveitar.

    Nem todo elemento repetido é protegido: antes de proteger, cada grupo
    passa por _classificar_marca_ou_conteudo — repetir em toda página não
    significa ser logo (um slogan/tagline também repete e deveria ser
    traduzido, achado num caso real: "GATO MIA" + "AMOR EM CADA DETALHE")."""
    if len(indices) < min_paginas:
        return {}, {}

    candidatos = []  # (pagina, bloco, texto_normalizado)
    blocos_por_pagina: dict[int, list[dict]] = {}
    for idx_na_amostra, i in enumerate(indices):
        pix = doc[i].get_pixmap(dpi=DPI_RENDER_IMAGEM)
        imagem = Image.open(io.BytesIO(pix.tobytes("png")))
        blocos_pagina = _ocr_blocos_pagina(imagem)
        blocos_por_pagina[i] = blocos_pagina
        if on_progress:
            on_progress(idx_na_amostra + 1, len(indices))
        for b in blocos_pagina:
            if len(b["texto"].split()) > MAX_PALAVRAS_ELEMENTO_REPETIDO:
                continue
            normalizado = _normalizar_para_comparacao(b["texto"])
            if len(normalizado) < MIN_CARACTERES_ELEMENTO_REPETIDO:
                continue
            candidatos.append((i, b, normalizado))

    usados = set()
    grupos_candidatos = []  # [(texto_original, [(pagina, bloco), ...]), ...]
    for idx_a in range(len(candidatos)):
        if idx_a in usados:
            continue
        pag_a, bloco_a, texto_a = candidatos[idx_a]
        grupo = [(pag_a, bloco_a)]
        paginas_com_match = {pag_a}
        for idx_b in range(idx_a + 1, len(candidatos)):
            if idx_b in usados:
                continue
            pag_b, bloco_b, texto_b = candidatos[idx_b]
            if pag_b == pag_a or texto_a != texto_b:
                continue
            grupo.append((pag_b, bloco_b))
            paginas_com_match.add(pag_b)
            usados.add(idx_b)
        if len(paginas_com_match) >= min_paginas:
            grupos_candidatos.append((bloco_a["texto"], grupo))

    # Classifica todos os grupos candidatos numa chamada só (não um por
    # grupo) -- documento típico tem só um punhado de elementos repetidos,
    # não vale a pena uma chamada de IA por grupo.
    eh_marca = _classificar_marca_ou_conteudo([texto for texto, _ in grupos_candidatos], on_uso=on_uso)

    areas_por_pagina: dict[int, list[tuple[float, float, float, float]]] = {}
    for (_, grupo), protegido in zip(grupos_candidatos, eh_marca):
        if not protegido:
            continue
        for pag, bloco in grupo:
            areas_por_pagina.setdefault(pag, []).append(
                (
                    bloco["x0"] - MARGEM_PROTECAO_PX,
                    bloco["y0"] - MARGEM_PROTECAO_PX,
                    bloco["x1"] + MARGEM_PROTECAO_PX,
                    bloco["y1"] + MARGEM_PROTECAO_PX,
                )
            )
    return areas_por_pagina, blocos_por_pagina


def process_pdf_imagem(
    input_path: Path,
    output_path: Path,
    source_lang: str,
    target_lang: str,
    page_indices: list[int] | None = None,
    areas_protegidas_por_pagina: dict[int, list[tuple[float, float, float, float]]] | None = None,
    blocos_ocr_cache: dict[int, list[dict]] | None = None,
    on_progress: Callable[[int, int], None] | None = None,
    on_uso: Callable[[dict], None] | None = None,
):
    """Traduz um PDF sem texto extraível (imagem/foto achatada) mantendo o
    layout: OCR pra ler o texto, LaMa pra limpar o fundo, e o mesmo
    insert_htmlbox de process_pdf pra escrever a tradução de volta.

    page_indices=None processa o documento inteiro; uma lista processa só
    essas páginas (0-based) — mesmo padrão de process_pdf, útil pra prévia
    de "N primeiras páginas".

    areas_protegidas_por_pagina: {indice_da_pagina: [(x0,y0,x1,y1), ...]}
    em pixels, na resolução DPI_RENDER_IMAGEM — regiões (tipicamente logo/
    marca) que não entram na limpeza nem na tradução. Ainda não é detectado
    automaticamente; quem chama precisa informar.

    blocos_ocr_cache: {indice_da_pagina: [bloco, ...]} — blocos de OCR já
    calculados por detectar_elementos_repetidos pras mesmas páginas.
    Quando presente pra uma página, pula o OCR dessa página aqui (o custo
    mais alto do pipeline) em vez de rodar de novo.

    Devolve a lista de índices (0-based, absolutos no documento) das
    páginas processadas que não tinham texto cobrável (ex: capa só com
    desenho/foto, sem letra nenhuma pra traduzir) — usado por quem chama
    pra não cobrar essas páginas no preço final."""
    doc_original = fitz.open(input_path)
    indices = page_indices if page_indices is not None else list(range(len(doc_original)))
    total_paginas = len(indices)
    areas_protegidas_por_pagina = areas_protegidas_por_pagina or {}

    doc_saida = fitz.open()
    escala = 72 / DPI_RENDER_IMAGEM
    blocos_ocr_cache = blocos_ocr_cache or {}
    paginas_sem_texto_cobravel: list[int] = []

    for indice_na_fila, i in enumerate(indices):
        pix = doc_original[i].get_pixmap(dpi=DPI_RENDER_IMAGEM)
        imagem_original = Image.open(io.BytesIO(pix.tobytes("png")))

        blocos = blocos_ocr_cache[i] if i in blocos_ocr_cache else _ocr_blocos_pagina(imagem_original)
        areas_protegidas = areas_protegidas_por_pagina.get(i, [])

        # Protege a LINHA inteira, não o bloco isolado — um nome de empresa
        # que se repete tanto na logo quanto dentro de uma frase normal
        # ("A Rosaves é uma empresa...") batia como "elemento repetido" e
        # sumia só a palavra, quebrando a frase. Já uma linha onde TODOS os
        # blocos são repetidos (a logo de verdade, sozinha) continua
        # protegida do jeito que já era.
        linhas_todas = _agrupar_em_linhas(blocos)
        linhas = [l for l in linhas_todas if not all(_bloco_protegido(b, areas_protegidas) for b in l["itens"])]
        paragrafos = _agrupar_linhas_em_paragrafos(linhas, altura_pagina_px=imagem_original.height)

        # Grupo isolado (nunca se juntou a nenhuma linha vizinha) com texto
        # de 1-2 caracteres quase sempre é ruído do OCR lendo um pedaço de
        # desenho decorativo como se fosse letra — visto na prática: "D" e
        # "DS" com confiança >0.98, mesma confiança de palavras reais,
        # então filtrar por score não funciona. Descartar aqui evita
        # reinserir um caractere solto e sem sentido na imagem final.
        paragrafos = [p for p in paragrafos if len(p["texto"].strip()) > 2]

        # Página sem texto cobrável: soma de caracteres reais abaixo do
        # limiar (ex: só uma logo ou nada de OCR) — não é justo cobrar o
        # preço cheio de uma página que não tinha nada pra traduzir.
        caracteres_pagina = sum(len(p["texto"].strip()) for p in paragrafos)
        if caracteres_pagina < MIN_CARACTERES_PAGINA_COBRAVEL:
            paginas_sem_texto_cobravel.append(i)

        imagem_limpa = _limpar_fundo(imagem_original, paragrafos)

        largura_pt, altura_pt = imagem_limpa.width * escala, imagem_limpa.height * escala
        page = doc_saida.new_page(width=largura_pt, height=altura_pt)

        buffer_imagem = io.BytesIO()
        imagem_limpa.save(buffer_imagem, format="PNG")
        page.insert_image(page.rect, stream=buffer_imagem.getvalue())

        if paragrafos:
            traducoes = translate_batch([p["texto"] for p in paragrafos], source_lang, target_lang, on_uso=on_uso)
            for paragrafo, traduzido in zip(paragrafos, traducoes):
                altura_linha_px = (paragrafo["y1"] - paragrafo["y0"]) / max(1, paragrafo["texto"].count(" ") // 8 + 1)
                tamanho_pt = max(8, min(60, altura_linha_px * 72 / DPI_RENDER_IMAGEM * 0.85))
                rect = fitz.Rect(
                    paragrafo["x0"] * escala, paragrafo["y0"] * escala,
                    paragrafo["x1"] * escala, paragrafo["y1"] * escala,
                )
                css = f"* {{ font-family: Helvetica, Arial, sans-serif; font-size: {tamanho_pt:.1f}pt; color: #262626; }}"
                page.insert_htmlbox(rect, html.escape(traduzido), css=css, scale_low=0)

        if on_progress:
            on_progress(indice_na_fila + 1, total_paginas)

    doc_original.close()
    doc_saida.save(output_path, garbage=4, deflate=True)
    doc_saida.close()
    return paginas_sem_texto_cobravel


def process_pdf(
    input_path: Path,
    output_path: Path,
    source_lang: str,
    target_lang: str,
    page_indices: list[int] | None = None,
    on_progress: Callable[[int, int], None] | None = None,
    on_uso: Callable[[dict], None] | None = None,
):
    """page_indices=None processa o documento inteiro; uma lista processa só
    essas páginas (0-based) — usado pela prévia grátis (só a 1ª página).
    on_progress(paginas_feitas, paginas_total), se passado, é chamado depois
    de cada página terminar — usado pra barra de progresso do documento
    completo pós-pagamento. on_uso({"prompt_tokens", "completion_tokens"}),
    se passado, é chamado a cada chamada de tradução — usado pra calcular o
    custo real de IA do job."""
    doc = fitz.open(input_path)
    pages = [doc[i] for i in page_indices] if page_indices is not None else doc
    total_paginas_a_processar = len(pages)

    for indice_na_fila, page in enumerate(pages, start=1):
        # Guarda quanto conteudo visual (imagens/vetores) a pagina tinha antes
        # de mexer, pra comparar depois e pegar automaticamente qualquer bug
        # tipo "redacao apagou pedaco de imagem/borda por baixo" (ja aconteceu
        # com uma logo e com bordas de tabela).
        num_imagens_antes = len(page.get_image_info())
        num_desenhos_antes = len(page.get_drawings())

        # Linhas divisorias finas (separador de secao, sublinhado de titulo
        # etc.) tem bbox com altura/largura quase zero — a checagem de
        # colisao com conteudo visual (mais abaixo) ignora elas de proposito
        # (ruido demais senao), o que significa que uma redacao por cima
        # apaga a linha sem ninguem perceber. Podem ser desenhadas tanto como
        # traco ("s") quanto como retangulo fino preenchido ("f") — guarda os
        # dois tipos agora pra redesenhar depois de inserir o texto, do mesmo
        # jeito que a borda de tabela.
        linhas_finas = [d for d in page.get_drawings() if _eh_linha_fina(d)]

        # Render da pagina ANTES de qualquer redacao -- usado por
        # _cor_fundo_real pra saber a cor de fundo de verdade em cada
        # ponto (ver docstring da funcao pro porque nao confiar em
        # geometria de forma). dpi=72 faz 1pt = 1px, mapeando direto pras
        # coordenadas de bbox que ja usamos no resto do arquivo.
        pix_fundo = page.get_pixmap(dpi=72)

        # Detecta tabelas de verdade (pelas linhas do desenho) pra nao deixar
        # o agrupamento generico de texto misturar conteudo de celulas vizinhas.
        #
        # find_tables() pode confundir retangulos de fundo decorativos (ex:
        # uma "moldura" atras de uma foto) com grade de tabela e devolver
        # uma unica celula gigante cobrindo a pagina inteira -- uma celula
        # real nunca ocupa a maior parte da propria tabela (precisa de pelo
        # menos 2 linhas/colunas pra ser tabela de verdade). Quando isso
        # acontece, a celula gigante vira um redact_annot do tamanho da
        # pagina, redatado com a cor do fundo mais especifico por baixo
        # (podendo ser preto/escuro de um elemento decorativo) — apagando a
        # pagina inteira. Achado real numa pagina de capa (foto + titulo,
        # sem tabela nenhuma) do arquivo "Gato Mia", ver claude-sessions-
        # log/sessions/2026-09-25_tradutor-*.md. Ignora a tabela inteira
        # nesse caso (nao so a celula suspeita) — o texto cai no caminho
        # normal de blocos, mais seguro (redacao no bbox do proprio texto).
        # find_tables() tambem pode inventar uma grade plausivel de VARIAS
        # celulas normais (nenhuma delas gigante) só pelo alinhamento de
        # texto/icones da pagina, mesmo sem nenhuma linha de grade
        # desenhada de verdade — visto na pratica numa ficha tecnica cheia
        # de selos/icones (sem tabela nenhuma), gerando celulas que se
        # sobrepoem parcialmente com blocos de texto normais e duplicam
        # trecho traduzido. `linhas_finas` (abaixo) ja detecta separador/
        # borda fina de verdade; se a pagina nao tem NENHUMA, uma tabela
        # "encontrada" aqui e quase certamente essa mesma alucinacao —
        # ignora todas nesse caso. Reusa `linhas_finas` (ja calculado
        # acima) como evidencia de linha de grade de verdade.
        table_cell_rects: list[fitz.Rect] = []
        table_areas: list[fitz.Rect] = []
        for table in (page.find_tables().tables if linhas_finas else []):
            tbbox = fitz.Rect(table.bbox)
            celulas = [fitz.Rect(c) for c in table.cells if c]
            if not _tabela_parece_real(tbbox, celulas):
                continue
            table_areas.append(tbbox)
            table_cell_rects.extend(celulas)

        block_infos = []
        flat_originals = []  # todos os runs de texto da pagina, na ordem

        # Celulas de tabela: cada celula vira uma unica unidade de traducao,
        # usando a bbox real da celula (nunca ultrapassa pra celula vizinha).
        for cell_rect in table_cell_rects:
            cell_dict = page.get_text("dict", clip=cell_rect)
            cell_blocks = [b for b in cell_dict["blocks"] if b.get("type") == 0]
            if not cell_blocks:
                continue
            size = cell_blocks[0]["lines"][0]["spans"][0]["size"]
            color = cell_blocks[0]["lines"][0]["spans"][0]["color"]
            lines_runs = _merged_lines_runs(cell_blocks)
            if not lines_runs or not _has_visible_text(lines_runs):
                continue
            run_indices = []
            for line_runs, linha_bbox in lines_runs:
                for run in line_runs:
                    run_indices.append(len(flat_originals))
                    flat_originals.append(_texto_traduzivel(page, run[0], linha_bbox))
            block_infos.append(
                {
                    "bbox": tuple(cell_rect),
                    "lines_runs": lines_runs,
                    "run_indices": run_indices,
                    "size": size,
                    "color": color,
                    "tabela": True,
                }
            )

        # Blocos normais de texto, pulando qualquer bloco que caia dentro de
        # uma area de tabela (esse ja foi tratado acima, celula por celula).
        # Cada bloco pode ser dividido em varios sub-blocos (ver
        # _dividir_em_subblocos) quando parece misturar conteudo de fileiras
        # diferentes — a redacao (Fase 1) sempre usa a bbox do bloco INTEIRO
        # original (bbox_redacao), garantindo que tudo seja apagado mesmo
        # quando a insercao (Fase 2) usa a bbox mais justa de cada sub-bloco.
        blocks = [b for b in page.get_text("dict")["blocks"] if b.get("type") == 0]
        for b in blocks:
            if any(_rect_center_inside(b["bbox"], area) for area in table_areas):
                continue
            lines_runs_bloco = _block_lines_runs(b)
            if not lines_runs_bloco or not _has_visible_text(lines_runs_bloco):
                continue
            first_span = b["lines"][0]["spans"][0]
            for subgrupo in _dividir_em_subblocos(lines_runs_bloco):
                if not subgrupo or not _has_visible_text(subgrupo):
                    continue
                bbox_subgrupo = (
                    min(bbox[0] for _r, bbox in subgrupo),
                    min(bbox[1] for _r, bbox in subgrupo),
                    max(bbox[2] for _r, bbox in subgrupo),
                    max(bbox[3] for _r, bbox in subgrupo),
                )
                run_indices = []
                for line_runs, linha_bbox in subgrupo:
                    for run in line_runs:
                        run_indices.append(len(flat_originals))
                        flat_originals.append(_texto_traduzivel(page, run[0], linha_bbox))
                block_infos.append(
                    {
                        "bbox": bbox_subgrupo,
                        "bbox_redacao": b["bbox"],
                        "lines_runs": subgrupo,
                        "run_indices": run_indices,
                        "size": first_span["size"],
                        "color": first_span["color"],
                        "tabela": False,
                    }
                )

        if not block_infos:
            if on_progress:
                on_progress(indice_na_fila, total_paginas_a_processar)
            continue

        _avisar_colisao_com_conteudo_visual(page, block_infos, table_areas)

        flat_translations = translate_batch(flat_originals, source_lang, target_lang, on_uso=on_uso)

        # Fase 1: marca e aplica TODAS as redacoes da pagina de uma vez, antes de
        # inserir qualquer texto novo (evita que a redacao de um bloco vizinho
        # apague pedaco do texto ja inserido de outro bloco, quando as caixas
        # originais se encostam/leve sobreposicao de bbox).
        for info in block_infos:
            bbox = fitz.Rect(info.get("bbox_redacao", info["bbox"]))
            cor_fundo = _cor_fundo_real(pix_fundo, bbox)
            page.add_redact_annot(bbox, fill=cor_fundo)
        page.apply_redactions()

        # Fase 2: insere o texto traduzido de cada bloco.
        blocos_normais = [i for i in block_infos if not i["tabela"]]
        for info in block_infos:
            bbox_original = fitz.Rect(info["bbox"])
            if info["tabela"]:
                rect = bbox_original
            else:
                # Usa uma largura de insercao possivelmente mais estreita que
                # a bbox original, pra nao vazar em cima de um bloco vizinho
                # a direita (ver _largura_segura_x1). A redacao da Fase 1 ja
                # usou a bbox original inteira, entao isso so afeta onde o
                # texto NOVO pode ocupar, nao o que foi apagado.
                x1_seguro = _largura_segura_x1(info, blocos_normais)
                rect = fitz.Rect(bbox_original.x0, bbox_original.y0, x1_seguro, bbox_original.y1)
            color = info["color"]
            color_hex = "#{:06x}".format(color if color else 0)

            idx_iter = iter(info["run_indices"])
            html_lines: list[str] = []
            prefixo_pendente = ""
            linha_anterior_era_wrap = False
            for line_runs, _line_bbox in info["lines_runs"]:
                parts = []
                for text, bold in line_runs:
                    i = next(idx_iter)
                    translated = html.escape(flat_translations[i])
                    parts.append(f"<b>{translated}</b>" if bold else translated)
                linha_html = "".join(parts)

                # Marcador de lista (•, -, etc.) as vezes vem como uma "linha"
                # propria na extracao do PDF, separada do texto que ele
                # introduz, mesmo os dois ficando juntos no visual original.
                # Forcar quebra de linha aqui dobraria a altura necessaria e
                # o insert_htmlbox encolheria a fonte pra caber na bbox
                # (bem apertada) — em vez disso, funde o marcador com a
                # proxima linha.
                texto_puro = "".join(t for t, _ in line_runs).strip()
                if len(line_runs) == 1 and texto_puro in _MARCADORES_SOLTOS:
                    prefixo_pendente = linha_html + " "
                    linha_anterior_era_wrap = False
                    continue

                # Se a linha ANTERIOR nao termina em pontuacao final (. : ; ! ?),
                # ela quase certamente era so um ponto de quebra de largura de
                # coluna no PDF original (word-wrap no meio da frase), nao uma
                # quebra de verdade — usar largura da bbox pra decidir isso se
                # mostrou pouco confiavel neste tipo de documento (blocos as
                # vezes misturam fragmentos de linhas/colunas vizinhas, o que
                # distorce a largura "esperada" do bloco). Pontuacao final e um
                # sinal muito mais direto de fim de frase/rotulo. Quando e so
                # quebra de largura, junta com espaco em vez de <br> — assim o
                # texto traduzido (que pode ter tamanho bem diferente do
                # original) se reorganiza sozinho dentro da caixa no
                # insert_htmlbox, em vez de ficar preso exatamente nos mesmos
                # pontos de quebra do idioma de origem (o que causava texto
                # traduzido mais longo sobrepondo a linha seguinte).
                if html_lines and not prefixo_pendente and linha_anterior_era_wrap:
                    html_lines[-1] = html_lines[-1] + " " + linha_html
                else:
                    html_lines.append(prefixo_pendente + linha_html)
                prefixo_pendente = ""

                linha_anterior_era_wrap = not texto_puro.rstrip().endswith((".", ":", ";", "!", "?"))
            if prefixo_pendente:
                html_lines.append(prefixo_pendente)
            html_content = "<br>".join(html_lines)

            css = (
                f"* {{ font-family: Helvetica, Arial, sans-serif; "
                f"font-size: {info['size']}pt; color: {color_hex}; }}"
            )
            page.insert_htmlbox(rect, html_content, css=css, scale_low=0)

        # As bordas da tabela costumam ser desenhadas como retangulos finos
        # bem em cima do limite de cada celula — a redacao de texto por cima
        # apaga esses pixels junto. Redesenha a grade depois de inserir o
        # texto pra tabela ficar identica a original.
        for cell_rect in table_cell_rects:
            page.draw_rect(cell_rect, color=(0, 0, 0), width=0.75)

        # Redesenha as linhas divisorias finas capturadas no inicio, com a
        # mesma cor/espessura originais.
        for linha in linhas_finas:
            r = fitz.Rect(linha["rect"])
            cor = linha.get("color") or linha.get("fill") or (0, 0, 0)
            if r.height <= r.width:
                y_meio = (r.y0 + r.y1) / 2
                p1, p2 = (r.x0, y_meio), (r.x1, y_meio)
                espessura = linha.get("width") or r.height or 0.5
            else:
                x_meio = (r.x0 + r.x1) / 2
                p1, p2 = (x_meio, r.y0), (x_meio, r.y1)
                espessura = linha.get("width") or r.width or 0.5
            page.draw_line(p1, p2, color=cor, width=max(espessura, 0.3))

        # Confere se sobrou tudo que a pagina original tinha de visual.
        num_imagens_depois = len(page.get_image_info())
        num_desenhos_depois = len(page.get_drawings())
        if num_imagens_depois < num_imagens_antes:
            print(
                f"[ALERTA] pagina {page.number}: perdeu imagem no processamento "
                f"({num_imagens_antes} -> {num_imagens_depois})"
            )
        if num_desenhos_depois < num_desenhos_antes * 0.8:
            print(
                f"[ALERTA] pagina {page.number}: perdeu boa parte dos desenhos/bordas "
                f"({num_desenhos_antes} -> {num_desenhos_depois})"
            )

        if on_progress:
            on_progress(indice_na_fila, total_paginas_a_processar)

    # garbage=4+deflate: reduz bastante o tamanho final sem perder nada
    # (limpeza de objeto orfao + compressao de stream) -- ex. real: um
    # PDF de 9 paginas de 0,4MB saia como 7,3MB sem isso.
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()


def paginas_sem_texto_extraivel(doc: fitz.Document, indices: list[int] | None = None) -> list[int]:
    """Varre TODAS as páginas pedidas (não só uma amostra) e devolve as que
    não têm texto extraível de verdade — ao contrário de
    analisar_texto_extraivel, que amostra só as primeiras páginas pra uma
    decisão binária do documento inteiro. Usado pra achar página-imagem
    isolada (ex: capa 100% gráfica) dentro de um documento majoritariamente
    de texto, caso que a classificação binária não pega."""
    alvo = indices if indices is not None else list(range(len(doc)))
    return [i for i in alvo if len(doc[i].get_text("text").strip()) < CARACTERES_MINIMOS_TEXTO_PAGINA]


def process_pdf_misto(
    input_path: Path,
    output_path: Path,
    source_lang: str,
    target_lang: str,
    page_indices: list[int] | None = None,
    on_progress: Callable[[int, int], None] | None = None,
    on_uso: Callable[[dict], None] | None = None,
) -> list[int]:
    """Processa um PDF que pode ter página-imagem isolada (sem texto
    extraível) dentro de um documento majoritariamente de texto — caso real
    "capa 100% gráfica, resto com texto real" que analisar_texto_extraivel
    não pega, porque só amostra pra classificar o documento inteiro como um
    todo (ver claude-sessions-log/sessions/2026-09-24_tradutor-pagamento-
    corrigido-upload-storage-e-classificador-marca.md).

    Detecta, página a página, quais páginas do recorte pedido não têm texto
    extraível e roda o pipeline de OCR+inpaint (process_pdf_imagem) só
    nelas; o resto segue no pipeline de texto normal (process_pdf). Depois
    remonta tudo num único PDF, respeitando a ordem original das páginas.

    page_indices=None processa o documento inteiro; uma lista restringe —
    mesmo padrão de process_pdf/process_pdf_imagem (usado pra prévia da 1ª
    página).

    Devolve a lista de índices (0-based, absolutos no documento) das
    páginas-imagem tratadas dessa forma — só informativo/log; essas páginas
    continuam entrando na contagem/preço normal do documento de texto (não
    usa a tabela de preço do PDF-imagem, que é só pro documento inteiro sem
    texto)."""
    doc = fitz.open(input_path)
    indices_alvo = page_indices if page_indices is not None else list(range(len(doc)))
    indices_sem_texto = paginas_sem_texto_extraivel(doc, indices_alvo)
    indices_com_texto = [i for i in indices_alvo if i not in indices_sem_texto]
    doc.close()

    if not indices_sem_texto:
        process_pdf(input_path, output_path, source_lang, target_lang, page_indices=page_indices, on_progress=on_progress, on_uso=on_uso)
        return []

    total = len(indices_alvo)

    def progresso_texto(feitas: int, _total_parcial: int):
        if on_progress:
            on_progress(feitas, total)

    def progresso_imagem(feitas: int, _total_parcial: int):
        if on_progress:
            on_progress(len(indices_com_texto) + feitas, total)

    texto_tmp = output_path.with_name(output_path.stem + "_texto_tmp.pdf")
    imagem_tmp = output_path.with_name(output_path.stem + "_imagem_tmp.pdf")

    # page_indices=indices_com_texto pode vir vazio (ex: prévia da 1ª página
    # quando ela mesma é a página sem texto) — process_pdf ainda salva o
    # documento inteiro nesse caso, só sem modificar nenhuma página, o que
    # serve de base correta pra substituição abaixo.
    process_pdf(
        input_path, texto_tmp, source_lang, target_lang,
        page_indices=indices_com_texto, on_progress=progresso_texto, on_uso=on_uso,
    )
    process_pdf_imagem(
        input_path, imagem_tmp, source_lang, target_lang,
        page_indices=indices_sem_texto, on_progress=progresso_imagem, on_uso=on_uso,
    )

    doc_principal = fitz.open(texto_tmp)
    doc_substituto = fitz.open(imagem_tmp)
    for idx_na_fila, i in enumerate(sorted(indices_sem_texto)):
        doc_principal.delete_page(i)
        doc_principal.insert_pdf(doc_substituto, from_page=idx_na_fila, to_page=idx_na_fila, start_at=i)
    doc_substituto.close()
    # garbage=4+deflate: sem isso, delete_page/insert_pdf deixa objeto
    # orfao no arquivo (visto na pratica: pagina unica de 6,6MB virou
    # documento final de 183MB sem essas flags).
    doc_principal.save(output_path, garbage=4, deflate=True)
    doc_principal.close()
    texto_tmp.unlink(missing_ok=True)
    imagem_tmp.unlink(missing_ok=True)

    return indices_sem_texto


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------

def _iter_runs_text_units(doc: Document):
    """Gera (container, index) de paragrafos com texto, incluindo os de dentro de tabelas."""
    for p in doc.paragraphs:
        if p.text.strip():
            yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip():
                        yield p


def process_docx(
    input_path: Path,
    output_path: Path,
    source_lang: str,
    target_lang: str,
    on_progress: Callable[[int, int], None] | None = None,
    on_uso: Callable[[dict], None] | None = None,
):
    """on_progress(paragrafos_feitos, paragrafos_total), se passado, é
    chamado depois de cada lote traduzido — usado pra barra de progresso.
    on_uso({"prompt_tokens", "completion_tokens"}), se passado, é chamado a
    cada chamada de tradução — usado pra calcular o custo real de IA do job."""
    doc = Document(input_path)
    paragraphs = list(_iter_runs_text_units(doc))
    originals = [p.text for p in paragraphs]

    # batches de ~40 paragrafos por chamada
    BATCH = 40
    translations: list[str] = []
    for i in range(0, len(originals), BATCH):
        translations.extend(translate_batch(originals[i : i + BATCH], source_lang, target_lang, on_uso=on_uso))
        if on_progress:
            on_progress(len(translations), len(originals))

    _aplicar_traducao_em_paragrafos(paragraphs, translations)
    doc.save(output_path)


def _aplicar_traducao_em_paragrafos(paragraphs: list, translations: list[str]) -> None:
    """Concentra o texto traduzido no primeiro run de cada paragrafo
    (preserva a formatacao dele) e limpa os runs seguintes, pra nao
    duplicar formatacao/texto antigo."""
    for p, translated in zip(paragraphs, translations):
        if not p.runs:
            continue
        p.runs[0].text = translated
        for r in p.runs[1:]:
            r.text = ""


def gerar_imagem_previa_docx(doc: Document, dpi: int = 150) -> str:
    """Salva o Document num arquivo temporario, converte pra PDF via
    LibreOffice headless (precisa do pacote libreoffice-writer instalado) e
    devolve a 1a pagina como PNG em base64 — usado pra previa visual do
    DOCX (mostrar que o layout foi mantido de verdade, nao so o texto
    solto sem formatacao)."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        docx_path = tmp_path / "previa.docx"
        doc.save(docx_path)
        # O usuario do container nao tem HOME de verdade (nao roda como root),
        # entao o LibreOffice falha ao criar seu perfil padrao ("User
        # installation could not be completed", exit 77) sem apontar um
        # diretorio gravavel explicito via UserInstallation.
        perfil_lo = f"file://{tmp_path}/lo_profile"
        subprocess.run(
            [
                "soffice",
                "--headless",
                f"-env:UserInstallation={perfil_lo}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_path),
                str(docx_path),
            ],
            check=True,
            timeout=60,
            capture_output=True,
            env={**os.environ, "HOME": str(tmp_path)},
        )
        pdf_doc = fitz.open(tmp_path / "previa.pdf")
        pix = pdf_doc[0].get_pixmap(dpi=dpi)
        png_bytes = pix.tobytes("png")
        pdf_doc.close()
        return base64.b64encode(png_bytes).decode()


# ---------------------------------------------------------------------------
# CLI de teste
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    source_lang = sys.argv[3] if len(sys.argv) > 3 else "portugues"
    target_lang = sys.argv[4] if len(sys.argv) > 4 else "espanhol"

    if src.suffix.lower() == ".pdf":
        process_pdf(src, dst, source_lang, target_lang)
    elif src.suffix.lower() == ".docx":
        process_docx(src, dst, source_lang, target_lang)
    else:
        raise SystemExit(f"Tipo de arquivo nao suportado: {src.suffix}")

    print(f"OK: {dst}")
