"""Testa o endpoint /comprimir isolado do Ghostscript de verdade
(comprimir_pdf vem sempre mockado) -- valida contrato HTTP: aceita só
PDF, rejeita arquivo absurdamente grande, devolve o header
X-Coube-No-Alvo certo, e passa adiante o erro quando a compressão falha.

Async (pytest-asyncio) porque httpx.ASGITransport -- o jeito atual de
testar um app ASGI sem subir servidor de verdade -- só tem handler async
(httpx.Client(app=...), que era sync, foi removido no httpx 0.28).
"""
from pathlib import Path

import httpx
import pytest

import app as app_module


@pytest.fixture
def client():
    transport = httpx.ASGITransport(app=app_module.app)
    return httpx.AsyncClient(transport=transport, base_url="http://test")


async def test_rejeita_arquivo_nao_pdf(client):
    resp = await client.post("/comprimir", files={"arquivo": ("foto.jpg", b"conteudo", "image/jpeg")})
    assert resp.status_code == 400


async def test_comprime_e_devolve_pdf_quando_cabe_no_alvo(client, monkeypatch):
    def fake_comprimir_pdf(input_path: Path, output_path: Path, alvo_bytes: int) -> bool:
        output_path.write_bytes(b"%PDF-1.4 comprimido")
        return True

    monkeypatch.setattr(app_module, "comprimir_pdf", fake_comprimir_pdf)

    resp = await client.post("/comprimir", files={"arquivo": ("catalogo.pdf", b"%PDF-1.4 original grande", "application/pdf")})

    assert resp.status_code == 200
    assert resp.headers["x-coube-no-alvo"] == "true"
    assert resp.content == b"%PDF-1.4 comprimido"
    assert resp.headers["content-type"] == "application/pdf"


async def test_avisa_quando_nao_coube_mas_ainda_devolve_melhor_esforco(client, monkeypatch):
    def fake_comprimir_pdf(input_path: Path, output_path: Path, alvo_bytes: int) -> bool:
        output_path.write_bytes(b"%PDF-1.4 melhor esforco")
        return False

    monkeypatch.setattr(app_module, "comprimir_pdf", fake_comprimir_pdf)

    resp = await client.post("/comprimir", files={"arquivo": ("catalogo.pdf", b"%PDF-1.4 original grande", "application/pdf")})

    assert resp.status_code == 200
    assert resp.headers["x-coube-no-alvo"] == "false"
    assert resp.content == b"%PDF-1.4 melhor esforco"


async def test_falha_de_compressao_vira_erro_500(client, monkeypatch):
    def fake_comprimir_pdf(input_path: Path, output_path: Path, alvo_bytes: int) -> bool:
        return False  # nao cria output_path -- simula Ghostscript indisponivel/quebrado

    monkeypatch.setattr(app_module, "comprimir_pdf", fake_comprimir_pdf)

    resp = await client.post("/comprimir", files={"arquivo": ("catalogo.pdf", b"%PDF-1.4 original", "application/pdf")})

    assert resp.status_code == 500


async def test_rejeita_arquivo_maior_que_limite_absoluto(client, monkeypatch):
    monkeypatch.setattr(app_module, "MAX_TAMANHO_ARQUIVO_COMPRESSAO", 10)

    resp = await client.post("/comprimir", files={"arquivo": ("catalogo.pdf", b"0" * 100, "application/pdf")})

    assert resp.status_code == 413
