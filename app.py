"""
API do Tradutor (FastAPI). Por enquanto só o endpoint de previa gratis
(so a 1a pagina de PDF, ou os primeiros paragrafos de DOCX) — o
processamento completo pago entra numa proxima etapa, junto com o
Mercado Pago.
"""
import base64
import os
import tempfile
import time
from collections import defaultdict, deque
from pathlib import Path

import fitz  # pymupdf
import httpx
from docx import Document
from dotenv import load_dotenv
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Request, UploadFile, Header
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from pipeline import (
    _aplicar_traducao_em_paragrafos,
    analisar_texto_extraivel,
    custo_centavos_brl,
    detectar_elementos_repetidos,
    gerar_imagem_previa_docx,
    process_docx,
    process_pdf,
    process_pdf_imagem,
    translate_batch,
)

load_dotenv(Path(__file__).parent / ".env")

API_KEY = os.environ.get("TRADUTOR_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_SECRET_KEY = os.environ.get("SUPABASE_SECRET_KEY")
BUCKET_ARQUIVOS = "tradutor-arquivos"

app = FastAPI(title="Tradutor API")

MAX_TAMANHO_ARQUIVO = 15 * 1024 * 1024  # 15 MB — prévia só usa a 1ª pagina/paragrafos

# Limite de requisicoes por IP no endpoint publico de previa (evita bot
# martelar o endpoint e gerar custo de OpenAI sem controle). Em memoria —
# reseta se o container reiniciar, suficiente pra uma instancia so.
LIMITE_REQUISICOES_POR_HORA = 15
_requisicoes_por_ip: dict[str, deque] = defaultdict(deque)


def _checar_rate_limit(ip: str):
    agora = time.time()
    fila = _requisicoes_por_ip[ip]
    while fila and agora - fila[0] > 3600:
        fila.popleft()
    if len(fila) >= LIMITE_REQUISICOES_POR_HORA:
        raise HTTPException(status_code=429, detail="Muitas tentativas. Tenta de novo mais tarde.")
    fila.append(agora)

ALLOWED_ORIGINS = [
    o.strip()
    for o in os.environ.get(
        "ALLOWED_ORIGINS", "http://localhost:3000,http://127.0.0.1:3000"
    ).split(",")
    if o.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Faixas de preco por pagina, aplicadas ao documento inteiro conforme o
# total de paginas (sem degrau brusco tipo "49 paginas custa mais que 50" —
# cada faixa cobre o documento todo, nao so as paginas acima do limite).
FAIXAS_PRECO = [
    (15, 500),  # ate 15 paginas: R$5,00/pagina
    (50, 450),  # 16-50: R$4,50/pagina (10% off)
    (100, 400),  # 51-100: R$4,00/pagina (20% off)
    (200, 350),  # 101-200: R$3,50/pagina (30% off)
    (float("inf"), 300),  # 200+: R$3,00/pagina (40% off)
]
PRECO_MINIMO_CENTAVOS = 1490


def _checar_api_key(x_api_key: str | None):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="chave de API invalida")


def _preco_por_pagina_centavos(paginas: int) -> int:
    for limite, preco in FAIXAS_PRECO:
        if paginas <= limite:
            return preco
    return FAIXAS_PRECO[-1][1]


def _calcular_preco(paginas: int) -> int:
    preco_pagina = _preco_por_pagina_centavos(paginas)
    return max(PRECO_MINIMO_CENTAVOS, paginas * preco_pagina)


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


def _ip_do_cliente(request: Request) -> str:
    encaminhado = request.headers.get("x-forwarded-for")
    if encaminhado:
        return encaminhado.split(",")[0].strip()
    return request.client.host if request.client else "desconhecido"


@app.post("/preview")
async def preview(
    request: Request,
    arquivo: UploadFile = File(...),
    idioma_origem: str = Form(...),
    idioma_destino: str = Form(...),
    x_api_key: str | None = Header(default=None),
):
    _checar_api_key(x_api_key)
    _checar_rate_limit(_ip_do_cliente(request))

    sufixo = Path(arquivo.filename or "").suffix.lower()
    if sufixo not in (".pdf", ".docx", ".jpg", ".jpeg", ".png", ".webp"):
        raise HTTPException(status_code=400, detail="Só aceitamos .pdf, .docx, .jpg, .jpeg, .png ou .webp")

    conteudo = await arquivo.read()
    if len(conteudo) > MAX_TAMANHO_ARQUIVO:
        raise HTTPException(status_code=413, detail="Arquivo muito grande (limite de 15 MB).")

    with tempfile.TemporaryDirectory() as tmp:
        origem = Path(tmp) / f"origem{sufixo}"
        origem.write_bytes(conteudo)

        if sufixo == ".docx":
            return _preview_docx(origem, Path(tmp), idioma_origem, idioma_destino)
        # PDF e imagem (jpg/png/webp) passam pelo mesmo caminho -- o fitz
        # abre uma imagem crua como um "documento" de 1 página sem texto
        # extraível, então _preview_pdf já detecta eh_imagem=True sozinho e
        # cai no mesmo fluxo de PDF-imagem, sem precisar converter nada.
        return _preview_pdf(origem, Path(tmp), idioma_origem, idioma_destino)


def _preview_pdf(origem: Path, tmp: Path, idioma_origem: str, idioma_destino: str):
    doc_original = fitz.open(origem)
    total_paginas = len(doc_original)
    pix_original = doc_original[0].get_pixmap(dpi=200)
    imagem_original_b64 = base64.b64encode(pix_original.tobytes("png")).decode()
    analise_texto = analisar_texto_extraivel(doc_original)
    doc_original.close()

    if analise_texto["eh_imagem"]:
        # PDF sem camada de texto real (imagem/foto achatada) — o pipeline
        # de redação/reinserção não tem o que extrair aqui. Não roda
        # process_pdf (sairia idêntico ao original, sem avisar ninguém) e
        # devolve um tipo à parte pro frontend mostrar o aviso em vez da
        # prévia normal.
        return {
            "tipo": "pdf_sem_texto",
            "paginas_total": total_paginas,
            "imagem_original_base64": imagem_original_b64,
        }

    traduzido = tmp / "traduzido.pdf"
    process_pdf(origem, traduzido, idioma_origem, idioma_destino, page_indices=[0])

    doc_traduzido = fitz.open(traduzido)
    pix_traduzido = doc_traduzido[0].get_pixmap(dpi=200)
    imagem_traduzida_b64 = base64.b64encode(pix_traduzido.tobytes("png")).decode()

    # PDF de exemplo com só a 1ª página traduzida, pra baixar e conferir que o
    # texto é de verdade (selecionável/editável), não uma imagem achatada.
    pagina_unica = fitz.open()
    pagina_unica.insert_pdf(doc_traduzido, from_page=0, to_page=0)
    pdf_traduzido_b64 = base64.b64encode(pagina_unica.tobytes()).decode()
    pagina_unica.close()
    doc_traduzido.close()

    return {
        "tipo": "pdf",
        "paginas_total": total_paginas,
        "preco_centavos": _calcular_preco(total_paginas),
        "preco_por_pagina_centavos": _preco_por_pagina_centavos(total_paginas),
        "imagem_original_base64": imagem_original_b64,
        "imagem_traduzida_base64": imagem_traduzida_b64,
        "pdf_traduzido_base64": pdf_traduzido_b64,
    }


# Faixas de preço do PDF-imagem — mesmo padrão de FAIXAS_PRECO (documento
# inteiro numa faixa só, sem degrau brusco), mas com piso mais caro e mais
# alto que o texto em qualquer volume: o processamento é mais pesado (OCR +
# inpaint + reescrita, praticamente sequencial por página numa VPS de CPU
# limitada) e é um diferencial sem alternativa no mercado, então não faz
# sentido correr pro mesmo piso do texto (R$3,00) só porque o documento é
# grande. Decisão de 24/09/2026 com o Robson, depois de discutir a
# economia unitária.
FAIXAS_PRECO_IMAGEM = [
    (15, 1000),  # até 15 páginas: R$10,00/página
    (50, 800),  # 16-50: R$8,00/página (20% off)
    (100, 600),  # 51-100: R$6,00/página (40% off)
    (float("inf"), 500),  # 100+: R$5,00/página (50% off)
]


def _preco_por_pagina_imagem_centavos(paginas: int) -> int:
    for limite, preco in FAIXAS_PRECO_IMAGEM:
        if paginas <= limite:
            return preco
    return FAIXAS_PRECO_IMAGEM[-1][1]


def _calcular_preco_imagem(paginas: int) -> int:
    preco_pagina = _preco_por_pagina_imagem_centavos(paginas)
    return max(PRECO_MINIMO_CENTAVOS, paginas * preco_pagina)


# Quantas páginas processar de verdade pra prévia do PDF-imagem antes de
# cobrar — degressivo por tamanho do documento (decisão com o Robson em
# 24/09/2026: a fórmula antiga dava até 50% do documento de graça pra
# documentos pequenos/médios, caro demais pro processamento pesado de
# OCR+inpaint+tradução). Piso de 2 páginas sempre que o documento tiver 2+
# páginas: com só 1 página na amostra, detectar_elementos_repetidos não tem
# o que comparar entre páginas e a prévia grátis fica sem a proteção
# automática de logo — justamente o diferencial que a prévia deveria
# mostrar. Documento de exatamente 1 página segue sem proteção automática
# de logo (limitação conhecida, não resolvida — precisaria de outro sinal,
# não cross-página).
FAIXAS_PAGINAS_GRATIS_IMAGEM = [
    (20, 2),
    (30, 3),
    (40, 4),
    (50, 5),
    (float("inf"), 5),
]


def _n_paginas_previa_imagem(total_paginas: int) -> int:
    if total_paginas <= 1:
        return 1
    for limite, n_gratis in FAIXAS_PAGINAS_GRATIS_IMAGEM:
        if total_paginas <= limite:
            return min(n_gratis, total_paginas)
    return min(FAIXAS_PAGINAS_GRATIS_IMAGEM[-1][1], total_paginas)


class GerarPreviaImagemBody(BaseModel):
    job_id: str


@app.post("/gerar-previa-imagem", status_code=202)
async def gerar_previa_imagem(
    body: GerarPreviaImagemBody,
    background_tasks: BackgroundTasks,
    x_api_key: str | None = Header(default=None),
):
    """Dispara o processamento de verdade das primeiras páginas de um
    PDF-imagem em segundo plano. Assíncrono (202 + polling) igual ao
    /traduzir-completo, não porque o Robson pediu, mas porque uma chamada
    só levaria 1-2min — mais tempo do que uma função serverless do Vercel
    aguenta segurar aberta numa chamada síncrona. O arquivo já precisa
    estar no Storage (feito no upload, junto com a criação do job) e o
    job precisa existir com status "aguardando_previa_imagem"."""
    _checar_api_key(x_api_key)
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        raise HTTPException(status_code=500, detail="Supabase não configurado no backend.")

    job = _buscar_job(body.job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job não encontrado")
    if job["status"] != "aguardando_previa_imagem":
        # Idempotente: um clique duplo ou um retry nao deve disparar o
        # processamento pesado duas vezes.
        return {"ok": True, "ignorado": f"status atual é {job['status']}"}

    _atualizar_job(body.job_id, {"status": "gerando_previa_imagem"})
    background_tasks.add_task(_processar_previa_imagem, job)
    return {"ok": True}


def _processar_previa_imagem(job: dict):
    job_id = job["id"]

    def progresso(feitas: int, total: int):
        try:
            _atualizar_job(job_id, {"unidades_processadas": feitas, "unidades_total": total})
        except Exception:
            pass  # nunca derruba o processamento por causa de um update de progresso

    try:
        conteudo = _baixar_do_storage(job["arquivo_original_path"])

        with tempfile.TemporaryDirectory() as tmp:
            # Extensão real do arquivo (pode ser .jpg/.png/.webp, não só
            # .pdf) -- o fitz abre imagem crua como documento de 1 página
            # sozinho, mas precisa do arquivo salvo com a extensão certa
            # pra reconhecer o formato.
            sufixo_original = Path(job["arquivo_original_path"]).suffix.lower()
            origem = Path(tmp) / f"origem{sufixo_original}"
            origem.write_bytes(conteudo)
            saida = Path(tmp) / "previa.pdf"

            doc = fitz.open(origem)
            total_paginas = len(doc)
            indices = list(range(_n_paginas_previa_imagem(total_paginas)))
            # Progresso unificado nas duas fases (deteccao de logo + traducao
            # em si) -- sem isso a barra fica muda/parada na fase de deteccao
            # (que tambem faz OCR pagina a pagina, mesmo custo de tempo da
            # traducao) e so aparece quando a segunda fase comeca, dando a
            # falsa impressao de que nada esta acontecendo.
            roda_deteccao = len(indices) >= 2
            total_unificado = len(indices) * 2 if roda_deteccao else len(indices)
            if roda_deteccao:
                areas_protegidas, blocos_ocr_cache = detectar_elementos_repetidos(
                    doc, indices, on_progress=lambda feitas, _total: progresso(feitas, total_unificado)
                )
            else:
                areas_protegidas, blocos_ocr_cache = {}, {}
            doc.close()

            offset = len(indices) if roda_deteccao else 0
            paginas_sem_texto = process_pdf_imagem(
                origem,
                saida,
                job["idioma_origem"],
                job["idioma_destino"],
                page_indices=indices,
                areas_protegidas_por_pagina=areas_protegidas,
                blocos_ocr_cache=blocos_ocr_cache,
                on_progress=lambda feitas, _total: progresso(offset + feitas, total_unificado),
            )

            doc_previa = fitz.open(saida)
            caminhos = []
            for i, page in enumerate(doc_previa):
                pix = page.get_pixmap(dpi=150)
                caminho = f"previas/{job_id}/{i}.png"
                _subir_para_storage(caminho, pix.tobytes("png"), "image/png")
                caminhos.append(caminho)
            doc_previa.close()

            # Renderiza também as páginas originais (sem tradução) na mesma
            # resolução, pra o frontend mostrar lado a lado com a traduzida
            # (zoom/comparação) — a prévia até aqui só guardava a traduzida.
            doc_original_paginas = fitz.open(origem)
            caminhos_originais = []
            for idx_na_fila, i in enumerate(indices):
                pix = doc_original_paginas[i].get_pixmap(dpi=150)
                caminho = f"previas/{job_id}/original_{idx_na_fila}.png"
                _subir_para_storage(caminho, pix.tobytes("png"), "image/png")
                caminhos_originais.append(caminho)
            doc_original_paginas.close()

        # Páginas confirmadas sem texto cobrável dentro da amostra já
        # processada não entram no preço — o resto do documento (fora da
        # amostra) só é confirmado depois do pagamento, no processamento
        # completo (ver _processar_job_completo).
        paginas_cobradas = max(1, total_paginas - len(paginas_sem_texto))

        _atualizar_job(
            job_id,
            {
                "status": "previa_imagem_pronta",
                "previas_imagem_paths": caminhos,
                "previas_imagem_originais_paths": caminhos_originais,
                "paginas_gratis_indices": paginas_sem_texto,
                "preco_centavos": _calcular_preco_imagem(paginas_cobradas),
            },
        )
    except Exception as e:
        _atualizar_job(job_id, {"status": "erro", "erro_mensagem": str(e)[:500]})


LIMITE_PARAGRAFOS_PREVIA = 12


def _preview_docx(origem: Path, tmp: Path, idioma_origem: str, idioma_destino: str):
    doc = Document(origem)
    paragrafos = [p for p in doc.paragraphs if p.text.strip()]
    total_paragrafos = len(paragrafos)
    amostra = paragrafos[:LIMITE_PARAGRAFOS_PREVIA]
    originais = [p.text for p in amostra]
    traduzidos = translate_batch(originais, idioma_origem, idioma_destino)

    # Estimativa grosseira de "paginas" pra DOCX (nao existe pagina fixa
    # nesse formato) — usada so pra calcular o preco.
    paginas_estimadas = max(1, round(total_paragrafos / 25))

    # Imagem da 1a pagina de verdade (via LibreOffice), pra mostrar que o
    # layout do Word e mantido — so os mesmos paragrafos ja traduzidos
    # acima, sem gastar mais tradução do que a previa em texto já gasta.
    imagem_original_b64 = gerar_imagem_previa_docx(Document(origem))
    doc_traduzido = Document(origem)
    paragrafos_traduzido = [p for p in doc_traduzido.paragraphs if p.text.strip()]
    _aplicar_traducao_em_paragrafos(paragrafos_traduzido[:LIMITE_PARAGRAFOS_PREVIA], traduzidos)
    imagem_traduzida_b64 = gerar_imagem_previa_docx(doc_traduzido)

    return {
        "tipo": "docx",
        "paginas_total": paginas_estimadas,
        "preco_centavos": _calcular_preco(paginas_estimadas),
        "preco_por_pagina_centavos": _preco_por_pagina_centavos(paginas_estimadas),
        "texto_original": originais,
        "texto_traduzido": traduzidos,
        "imagem_original_base64": imagem_original_b64,
        "imagem_traduzida_base64": imagem_traduzida_b64,
        "paragrafos_restantes": max(0, total_paragrafos - len(amostra)),
    }


# Limite bem mais alto que o do endpoint /preview (multipart) de proposito:
# esse aqui recebe so o CAMINHO no Storage, o arquivo em si nunca passa pelo
# Vercel -- que tem um limite fixo de ~4,5MB por requisicao de funcao
# serverless, nao configuravel, descoberto testando um catalogo real de 11MB
# em 25/09/2026 (o proprio /preview multipart nunca teria funcionado pra
# esse arquivo, mesmo estando dentro do limite de 15MB do app). 50MB cobre
# catalogos/apresentacoes reais com folga.
MAX_TAMANHO_ARQUIVO_STORAGE = 50 * 1024 * 1024


class PreviewStorageBody(BaseModel):
    arquivo_original_path: str
    idioma_origem: str
    idioma_destino: str


@app.post("/preview-do-storage")
async def preview_do_storage(
    request: Request,
    body: PreviewStorageBody,
    x_api_key: str | None = Header(default=None),
):
    """Mesma logica do /preview, mas pro arquivo ja estar no Storage (subido
    direto do navegador via URL assinada) em vez de vir por upload multipart
    direto nessa requisicao -- ver MAX_TAMANHO_ARQUIVO_STORAGE."""
    _checar_api_key(x_api_key)
    _checar_rate_limit(_ip_do_cliente(request))
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        raise HTTPException(status_code=500, detail="Supabase não configurado no backend.")

    sufixo = Path(body.arquivo_original_path).suffix.lower()
    if sufixo not in (".pdf", ".docx", ".jpg", ".jpeg", ".png", ".webp"):
        raise HTTPException(status_code=400, detail="Só aceitamos .pdf, .docx, .jpg, .jpeg, .png ou .webp")

    conteudo = _baixar_do_storage(body.arquivo_original_path)
    if len(conteudo) > MAX_TAMANHO_ARQUIVO_STORAGE:
        raise HTTPException(status_code=413, detail="Arquivo muito grande (limite de 50 MB).")

    with tempfile.TemporaryDirectory() as tmp:
        origem = Path(tmp) / f"origem{sufixo}"
        origem.write_bytes(conteudo)

        if sufixo == ".docx":
            return _preview_docx(origem, Path(tmp), body.idioma_origem, body.idioma_destino)
        return _preview_pdf(origem, Path(tmp), body.idioma_origem, body.idioma_destino)


# ---------------------------------------------------------------------------
# Processamento completo (pos-pagamento) — chamado pelo webhook do Mercado
# Pago em tradutor-web. Le/grava direto no Supabase (Storage + tabela jobs)
# via REST, sem SDK, pra nao depender de versao de biblioteca. Roda em
# BackgroundTasks porque um documento grande pode levar minutos — bem alem do
# tempo que uma funcao serverless do Next.js/Vercel aguentaria esperar.
# ---------------------------------------------------------------------------


def _supabase_headers(content_type: str | None = None) -> dict:
    headers = {"apikey": SUPABASE_SECRET_KEY, "Authorization": f"Bearer {SUPABASE_SECRET_KEY}"}
    if content_type:
        headers["Content-Type"] = content_type
    return headers


def _buscar_job(job_id: str) -> dict | None:
    resp = httpx.get(
        f"{SUPABASE_URL}/rest/v1/jobs",
        params={"id": f"eq.{job_id}", "select": "*"},
        headers=_supabase_headers(),
        timeout=30,
    )
    resp.raise_for_status()
    linhas = resp.json()
    return linhas[0] if linhas else None


def _atualizar_job(job_id: str, campos: dict):
    resp = httpx.patch(
        f"{SUPABASE_URL}/rest/v1/jobs",
        params={"id": f"eq.{job_id}"},
        headers={**_supabase_headers("application/json"), "Prefer": "return=minimal"},
        json=campos,
        timeout=30,
    )
    resp.raise_for_status()


def _baixar_do_storage(path: str) -> bytes:
    # timeout generoso (180s) -- desde 25/09/2026 arquivos de ate 50MB passam
    # por aqui direto na chamada sincrona de /preview-do-storage (nao so em
    # BackgroundTasks como antes), entao uma rede mais lenta nao pode estourar
    # o timeout no meio de um download real.
    resp = httpx.get(f"{SUPABASE_URL}/storage/v1/object/{BUCKET_ARQUIVOS}/{path}", headers=_supabase_headers(), timeout=180)
    resp.raise_for_status()
    return resp.content


def _subir_para_storage(path: str, conteudo: bytes, content_type: str):
    # x-upsert: sem isso, o Storage responde 400 se já existir um objeto
    # nesse caminho -- acontece de verdade quando o mesmo job é reprocessado
    # (webhook do Mercado Pago pode reenviar a notificação, e mesmo com a
    # checagem de idempotência em /traduzir-completo, um reprocessamento
    # legítimo do mesmo job_id deve poder sobrescrever o resultado anterior,
    # não falhar).
    resp = httpx.post(
        f"{SUPABASE_URL}/storage/v1/object/{BUCKET_ARQUIVOS}/{path}",
        headers={**_supabase_headers(content_type), "x-upsert": "true"},
        content=conteudo,
        timeout=180,
    )
    resp.raise_for_status()


class TraduzirCompletoBody(BaseModel):
    job_id: str


@app.post("/traduzir-completo", status_code=202)
async def traduzir_completo(
    body: TraduzirCompletoBody,
    background_tasks: BackgroundTasks,
    x_api_key: str | None = Header(default=None),
):
    _checar_api_key(x_api_key)
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        raise HTTPException(status_code=500, detail="Supabase não configurado no backend (SUPABASE_URL/SUPABASE_SECRET_KEY).")

    job = _buscar_job(body.job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job não encontrado")
    if job["status"] != "pago":
        # Idempotente: o webhook pode reenviar a mesma notificação de pagamento.
        return {"ok": True, "ignorado": f"status atual é {job['status']}"}

    _atualizar_job(body.job_id, {"status": "processando"})
    background_tasks.add_task(_processar_job_completo, job)
    return {"ok": True}


def _processar_job_completo(job: dict):
    job_id = job["id"]
    sufixo = Path(job["arquivo_original_path"]).suffix.lower()
    # Extensão do arquivo de SAÍDA -- não é sempre igual à de entrada: um
    # pdf_sem_texto pode ter vindo de uma imagem crua (.jpg/.png/.webp), mas
    # o resultado processado é sempre um PDF de verdade (montado do zero via
    # fitz em process_pdf_imagem), nunca a imagem original.
    sufixo_saida = ".docx" if job["tipo_arquivo"] == "docx" else ".pdf"

    # Atualiza o progresso no Supabase pra a tela de pagamento mostrar uma
    # barra de verdade — unidades_total so e conhecido quando o processamento
    # comeca (paginas do PDF, ou paragrafos reais do DOCX, diferente da
    # estimativa grosseira usada so pra calcular o preco).
    def progresso(feitas: int, total: int):
        try:
            _atualizar_job(job_id, {"unidades_processadas": feitas, "unidades_total": total})
        except Exception:
            pass  # nunca derruba o processamento por causa de um update de progresso

    # Soma os tokens de TODAS as chamadas de traducao do job (varias por
    # pagina/lote) pra calcular o custo real de IA no final — inclusive se o
    # job falhar no meio, os tokens ja gastos ate ali continuam sendo custo
    # de verdade, entao o total acumulado e salvo nos dois caminhos (sucesso
    # e erro), nao so no sucesso.
    uso_acumulado = {"prompt_tokens": 0, "completion_tokens": 0}

    def registrar_uso(usage: dict):
        uso_acumulado["prompt_tokens"] += usage.get("prompt_tokens", 0)
        uso_acumulado["completion_tokens"] += usage.get("completion_tokens", 0)

    try:
        conteudo = _baixar_do_storage(job["arquivo_original_path"])

        with tempfile.TemporaryDirectory() as tmp:
            origem = Path(tmp) / f"origem{sufixo}"
            origem.write_bytes(conteudo)
            destino = Path(tmp) / f"traduzido{sufixo_saida}"

            paginas_gratis_completo: list[int] | None = None
            if job["tipo_arquivo"] == "pdf_sem_texto":
                # Mesmo pipeline da prévia (OCR + inpaint + reescrita),
                # mas sem o corte de páginas — page_indices=None processa
                # o documento inteiro. A detecção de logo roda nas páginas
                # todas agora (não só nas poucas da prévia): documento
                # grande pode ter template de página diferente que só
                # aparece depois da 5ª página, e mais páginas = amostra
                # melhor pra achar o que se repete de verdade.
                doc = fitz.open(origem)
                indices_completos = list(range(len(doc)))
                # Mesmo progresso unificado (deteccao + traducao) da previa --
                # ver _processar_previa_imagem pro motivo.
                roda_deteccao = len(indices_completos) >= 2
                total_unificado = len(indices_completos) * 2 if roda_deteccao else len(indices_completos)
                if roda_deteccao:
                    areas_protegidas, blocos_ocr_cache = detectar_elementos_repetidos(
                        doc,
                        indices_completos,
                        on_progress=lambda feitas, _total: progresso(feitas, total_unificado),
                        on_uso=registrar_uso,
                    )
                else:
                    areas_protegidas, blocos_ocr_cache = {}, {}
                doc.close()

                offset = len(indices_completos) if roda_deteccao else 0

                # paginas_gratis_completo aqui é o dado verdadeiro (documento
                # inteiro, não só a amostra da prévia) — sobrescreve a
                # estimativa salva no pagamento só pra registro/auditoria,
                # não gera reembolso automático (já foi cobrado via Pix).
                paginas_gratis_completo = process_pdf_imagem(
                    origem,
                    destino,
                    job["idioma_origem"],
                    job["idioma_destino"],
                    areas_protegidas_por_pagina=areas_protegidas,
                    blocos_ocr_cache=blocos_ocr_cache,
                    on_progress=lambda feitas, _total: progresso(offset + feitas, total_unificado),
                    on_uso=registrar_uso,
                )
            elif sufixo == ".pdf":
                process_pdf(
                    origem,
                    destino,
                    job["idioma_origem"],
                    job["idioma_destino"],
                    on_progress=progresso,
                    on_uso=registrar_uso,
                )
            else:
                process_docx(
                    origem,
                    destino,
                    job["idioma_origem"],
                    job["idioma_destino"],
                    on_progress=progresso,
                    on_uso=registrar_uso,
                )

            traduzido_bytes = destino.read_bytes()

        caminho_traduzido = f"traduzidos/{job_id}{sufixo_saida}"
        content_type = (
            "application/pdf"
            if sufixo_saida == ".pdf"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        _subir_para_storage(caminho_traduzido, traduzido_bytes, content_type)
        custo_centavos = custo_centavos_brl(uso_acumulado["prompt_tokens"], uso_acumulado["completion_tokens"])
        campos_finais = {"status": "pronto", "arquivo_traduzido_path": caminho_traduzido, "custo_ia_centavos": custo_centavos}
        if paginas_gratis_completo is not None:
            campos_finais["paginas_gratis_indices"] = paginas_gratis_completo
        _atualizar_job(job_id, campos_finais)
    except Exception as exc:
        custo_centavos = custo_centavos_brl(uso_acumulado["prompt_tokens"], uso_acumulado["completion_tokens"])
        _atualizar_job(
            job_id, {"status": "erro", "erro_mensagem": str(exc)[:2000], "custo_ia_centavos": custo_centavos}
        )
