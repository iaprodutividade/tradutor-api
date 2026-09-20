"""
Prototipo de validacao (Fase 1) do pipeline do Tradutor.
Extrai blocos de texto de um PDF (posicao, fonte, tamanho, cor), traduz em lote
via OpenAI gpt-5-mini preservando numeros/codigos/unidades, remove o texto
original (redaction, sem tocar em imagem/vetor) e reinsere o texto traduzido
na mesma caixa com auto-ajuste de tamanho de fonte.

Tambem trata .docx (python-docx), traduzindo paragrafos e celulas de tabela
mantendo a formatacao nativa do Word.
"""
import html
import json
import os
from pathlib import Path

import fitz  # pymupdf
from docx import Document
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv(Path(__file__).parent / ".env")

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])
MODEL = "gpt-4o-mini"

SYSTEM_PROMPT = """Voce e um tradutor tecnico. Traduza cada item da lista do idioma de \
origem para o idioma de destino, mantendo o significado tecnico exato.

Regras obrigatorias:
- Preserve EXATAMENTE como estao, sem traduzir nem converter: numeros, percentuais, \
codigos (CAS, ONU, NCM), siglas, unidades de medida (mg, g, kg, km, %, degC etc.) e \
nomes de produto/marca.
- Nao adicione nem remova informacao. Nao resuma. Nao comente.
- Mantenha quebras de linha internas do item quando fizerem sentido.
- Se um item nao tiver texto traduzivel (so numero, so pontuacao, vazio), devolva o \
item inalterado.

Devolva APENAS um JSON com a chave "traducoes": lista de strings, na MESMA ORDEM e \
MESMA QUANTIDADE da lista de entrada."""


def _call_translate(texts: list[str], source_lang: str, target_lang: str) -> list[str]:
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
    return data["traducoes"]


def translate_batch(texts: list[str], source_lang: str, target_lang: str) -> list[str]:
    if not texts:
        return []

    for attempt in range(2):
        out = _call_translate(texts, source_lang, target_lang)
        if len(out) == len(texts):
            return out
        print(f"[aviso] tentativa {attempt + 1}: esperava {len(texts)} traducoes, recebi {len(out)} — repetindo")

    # Ainda inconsistente apos repetir: preenche com o original pra nao
    # derrubar o pipeline inteiro por causa de 1-2 itens problematicos.
    print(f"[aviso] mantendo divergencia de contagem — completando com o texto original onde faltar")
    fixed = list(out) + texts[len(out):]
    return fixed[: len(texts)]


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def _is_bold(span: dict) -> bool:
    return "bold" in span["font"].lower() or bool(span.get("flags", 0) & 16)


def _block_lines_runs(block: dict) -> list[list[list]]:
    """Por linha, agrupa spans consecutivos com o mesmo estilo (negrito) num run.
    Devolve lista de linhas, cada linha = lista de [texto, negrito]."""
    lines_runs = []
    for line in block["lines"]:
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
            lines_runs.append(runs)
    return lines_runs


def _rect_center_inside(bbox, rect: fitz.Rect) -> bool:
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    return rect.contains(fitz.Point(cx, cy))


def _merged_lines_runs(blocks: list[dict]) -> list[list[list]]:
    """Junta as linhas de varios blocos (ex.: todos os blocos dentro de uma
    celula de tabela) numa unica lista de linhas com runs por negrito."""
    lines_runs = []
    for b in blocks:
        if b.get("type") != 0:
            continue
        lines_runs.extend(_block_lines_runs(b))
    return lines_runs


def process_pdf(input_path: Path, output_path: Path, source_lang: str, target_lang: str):
    doc = fitz.open(input_path)

    for page in doc:
        # Detecta tabelas de verdade (pelas linhas do desenho) pra nao deixar
        # o agrupamento generico de texto misturar conteudo de celulas vizinhas.
        table_cell_rects: list[fitz.Rect] = []
        table_areas: list[fitz.Rect] = []
        for table in page.find_tables().tables:
            table_areas.append(fitz.Rect(table.bbox))
            for cell_bbox in table.cells:
                if cell_bbox:
                    table_cell_rects.append(fitz.Rect(cell_bbox))

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
            if not lines_runs:
                continue
            run_indices = []
            for line_runs in lines_runs:
                for run in line_runs:
                    run_indices.append(len(flat_originals))
                    flat_originals.append(run[0])
            block_infos.append(
                {
                    "bbox": tuple(cell_rect),
                    "lines_runs": lines_runs,
                    "run_indices": run_indices,
                    "size": size,
                    "color": color,
                }
            )

        # Blocos normais de texto, pulando qualquer bloco que caia dentro de
        # uma area de tabela (esse ja foi tratado acima, celula por celula).
        blocks = [b for b in page.get_text("dict")["blocks"] if b.get("type") == 0]
        for b in blocks:
            if any(_rect_center_inside(b["bbox"], area) for area in table_areas):
                continue
            lines_runs = _block_lines_runs(b)
            if not lines_runs:
                continue
            first_span = b["lines"][0]["spans"][0]
            run_indices = []
            for line_runs in lines_runs:
                for run in line_runs:
                    run_indices.append(len(flat_originals))
                    flat_originals.append(run[0])
            block_infos.append(
                {
                    "bbox": b["bbox"],
                    "lines_runs": lines_runs,
                    "run_indices": run_indices,
                    "size": first_span["size"],
                    "color": first_span["color"],
                }
            )

        if not block_infos:
            continue

        flat_translations = translate_batch(flat_originals, source_lang, target_lang)

        # Fase 1: marca e aplica TODAS as redacoes da pagina de uma vez, antes de
        # inserir qualquer texto novo (evita que a redacao de um bloco vizinho
        # apague pedaco do texto ja inserido de outro bloco, quando as caixas
        # originais se encostam/leve sobreposicao de bbox).
        for info in block_infos:
            page.add_redact_annot(fitz.Rect(info["bbox"]), fill=(1, 1, 1))
        page.apply_redactions()

        # Fase 2: insere o texto traduzido de cada bloco.
        for info in block_infos:
            rect = fitz.Rect(info["bbox"])
            color = info["color"]
            color_hex = "#{:06x}".format(color if color else 0)

            idx_iter = iter(info["run_indices"])
            html_lines = []
            for line_runs in info["lines_runs"]:
                parts = []
                for text, bold in line_runs:
                    i = next(idx_iter)
                    translated = html.escape(flat_translations[i])
                    parts.append(f"<b>{translated}</b>" if bold else translated)
                html_lines.append("".join(parts))
            html_content = "<br>".join(html_lines)

            css = (
                f"* {{ font-family: Helvetica, Arial, sans-serif; "
                f"font-size: {info['size']}pt; color: {color_hex}; }}"
            )
            page.insert_htmlbox(rect, html_content, css=css, scale_low=0)

    doc.save(output_path)
    doc.close()


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


def process_docx(input_path: Path, output_path: Path, source_lang: str, target_lang: str):
    doc = Document(input_path)
    paragraphs = list(_iter_runs_text_units(doc))
    originals = [p.text for p in paragraphs]

    # batches de ~40 paragrafos por chamada
    BATCH = 40
    translations: list[str] = []
    for i in range(0, len(originals), BATCH):
        translations.extend(translate_batch(originals[i : i + BATCH], source_lang, target_lang))

    for p, translated in zip(paragraphs, translations):
        if not p.runs:
            continue
        # concentra o texto traduzido no primeiro run (preserva a formatacao dele)
        # e limpa os runs seguintes, pra nao duplicar formatacao/texto antigo.
        p.runs[0].text = translated
        for r in p.runs[1:]:
            r.text = ""

    doc.save(output_path)


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
