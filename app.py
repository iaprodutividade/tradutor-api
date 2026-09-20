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
from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile, Header
from fastapi.middleware.cors import CORSMiddleware

from pipeline import process_pdf, translate_batch

load_dotenv(Path(__file__).parent / ".env")

API_KEY = os.environ.get("TRADUTOR_API_KEY")

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
    pix_original = doc_original[0].get_pixmap(dpi=120)
    imagem_original_b64 = base64.b64encode(pix_original.tobytes("png")).decode()
    doc_original.close()

    traduzido = tmp / "traduzido.pdf"
    process_pdf(origem, traduzido, idioma_origem, idioma_destino, page_indices=[0])

    doc_traduzido = fitz.open(traduzido)
    pix_traduzido = doc_traduzido[0].get_pixmap(dpi=120)
    imagem_traduzida_b64 = base64.b64encode(pix_traduzido.tobytes("png")).decode()
    doc_traduzido.close()

    return {
        "tipo": "pdf",
        "paginas_total": total_paginas,
        "preco_centavos": _calcular_preco(total_paginas),
        "preco_por_pagina_centavos": _preco_por_pagina_centavos(total_paginas),
        "imagem_original_base64": imagem_original_b64,
        "imagem_traduzida_base64": imagem_traduzida_b64,
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

    return {
        "tipo": "docx",
        "paginas_total": paginas_estimadas,
        "preco_centavos": _calcular_preco(paginas_estimadas),
        "preco_por_pagina_centavos": _preco_por_pagina_centavos(paginas_estimadas),
        "texto_original": originais,
        "texto_traduzido": traduzidos,
        "paragrafos_restantes": max(0, total_paragrafos - len(amostra)),
    }
