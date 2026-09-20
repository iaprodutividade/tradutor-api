"""
API do Tradutor (FastAPI). Por enquanto só o endpoint de previa gratis
(so a 1a pagina de PDF, ou os primeiros paragrafos de DOCX) — o
processamento completo pago entra numa proxima etapa, junto com o
Mercado Pago.
"""
import base64
import os
import tempfile
from pathlib import Path

import fitz  # pymupdf
from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile, Header
from fastapi.middleware.cors import CORSMiddleware

from pipeline import process_pdf, translate_batch

load_dotenv(Path(__file__).parent / ".env")

API_KEY = os.environ.get("TRADUTOR_API_KEY")

app = FastAPI(title="Tradutor API")

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

PRECO_POR_PAGINA_CENTAVOS = 500
PRECO_MINIMO_CENTAVOS = 1490


def _checar_api_key(x_api_key: str | None):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="chave de API invalida")


def _calcular_preco(paginas: int) -> int:
    return max(PRECO_MINIMO_CENTAVOS, paginas * PRECO_POR_PAGINA_CENTAVOS)


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.post("/preview")
async def preview(
    arquivo: UploadFile = File(...),
    idioma_origem: str = Form(...),
    idioma_destino: str = Form(...),
    x_api_key: str | None = Header(default=None),
):
    _checar_api_key(x_api_key)

    sufixo = Path(arquivo.filename or "").suffix.lower()
    if sufixo not in (".pdf", ".docx"):
        raise HTTPException(status_code=400, detail="Só aceitamos .pdf ou .docx")

    conteudo = await arquivo.read()

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
        "texto_original": originais,
        "texto_traduzido": traduzidos,
        "paragrafos_restantes": max(0, total_paragrafos - len(amostra)),
    }
