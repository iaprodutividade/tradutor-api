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
    gerar_imagem_previa_docx,
    process_docx,
    process_pdf,
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
    if sufixo not in (".pdf", ".docx"):
        raise HTTPException(status_code=400, detail="Só aceitamos .pdf ou .docx")

    conteudo = await arquivo.read()
    if len(conteudo) > MAX_TAMANHO_ARQUIVO:
        raise HTTPException(status_code=413, detail="Arquivo muito grande (limite de 15 MB).")

    with tempfile.TemporaryDirectory() as tmp:
        origem = Path(tmp) / f"origem{sufixo}"
        origem.write_bytes(conteudo)

        if sufixo == ".pdf":
            return _preview_pdf(origem, Path(tmp), idioma_origem, idioma_destino)
        return _preview_docx(origem, Path(tmp), idioma_origem, idioma_destino)


def _preview_pdf(origem: Path, tmp: Path, idioma_origem: str, idioma_destino: str):
    doc_original = fitz.open(origem)
    total_paginas = len(doc_original)
    pix_original = doc_original[0].get_pixmap(dpi=200)
    imagem_original_b64 = base64.b64encode(pix_original.tobytes("png")).decode()
    doc_original.close()

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
    resp = httpx.get(f"{SUPABASE_URL}/storage/v1/object/{BUCKET_ARQUIVOS}/{path}", headers=_supabase_headers(), timeout=120)
    resp.raise_for_status()
    return resp.content


def _subir_para_storage(path: str, conteudo: bytes, content_type: str):
    resp = httpx.post(
        f"{SUPABASE_URL}/storage/v1/object/{BUCKET_ARQUIVOS}/{path}",
        headers=_supabase_headers(content_type),
        content=conteudo,
        timeout=120,
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

    # Atualiza o progresso no Supabase pra a tela de pagamento mostrar uma
    # barra de verdade — unidades_total so e conhecido quando o processamento
    # comeca (paginas do PDF, ou paragrafos reais do DOCX, diferente da
    # estimativa grosseira usada so pra calcular o preco).
    def progresso(feitas: int, total: int):
        try:
            _atualizar_job(job_id, {"unidades_processadas": feitas, "unidades_total": total})
        except Exception:
            pass  # nunca derruba o processamento por causa de um update de progresso

    try:
        conteudo = _baixar_do_storage(job["arquivo_original_path"])

        with tempfile.TemporaryDirectory() as tmp:
            origem = Path(tmp) / f"origem{sufixo}"
            origem.write_bytes(conteudo)
            destino = Path(tmp) / f"traduzido{sufixo}"

            if sufixo == ".pdf":
                process_pdf(origem, destino, job["idioma_origem"], job["idioma_destino"], on_progress=progresso)
            else:
                process_docx(origem, destino, job["idioma_origem"], job["idioma_destino"], on_progress=progresso)

            traduzido_bytes = destino.read_bytes()

        caminho_traduzido = f"traduzidos/{job_id}{sufixo}"
        content_type = (
            "application/pdf"
            if sufixo == ".pdf"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        _subir_para_storage(caminho_traduzido, traduzido_bytes, content_type)
        _atualizar_job(job_id, {"status": "pronto", "arquivo_traduzido_path": caminho_traduzido})
    except Exception as exc:
        _atualizar_job(job_id, {"status": "erro", "erro_mensagem": str(exc)[:2000]})
