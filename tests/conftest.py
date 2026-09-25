"""Fixtures compartilhadas. Nenhum teste chama a API real da OpenAI --
pipeline._call_translate vem sempre stubado (ver stub_translate), entao a
suite roda de graca, rapida e sem depender de internet/chave valida.

pipeline.py le OPENAI_API_KEY do ambiente no import (client = OpenAI(...)),
entao precisa de um valor (mesmo falso) antes do import acontecer.
"""
import os
import sys
from pathlib import Path

os.environ.setdefault("OPENAI_API_KEY", "sk-test-dummy-nao-usada-pelos-testes")
sys.path.insert(0, str(Path(__file__).parent.parent))

import fitz  # noqa: E402
import pytest  # noqa: E402

import pipeline  # noqa: E402


@pytest.fixture
def stub_translate(monkeypatch):
    """Troca pipeline._call_translate por uma versao determinista e
    gratuita: devolve cada item prefixado com 'EN:' (simula traducao sem
    chamar a API de verdade). Devolve a lista de chamadas feitas (cada
    uma a lista de textos daquela chamada), pra testes que precisam
    inspecionar o que foi mandado traduzir."""
    chamadas = []

    def fake_call_translate(texts, source_lang, target_lang):
        chamadas.append(list(texts))
        return [f"EN:{t}" for t in texts], {"prompt_tokens": 0, "completion_tokens": 0}

    monkeypatch.setattr(pipeline, "_call_translate", fake_call_translate)
    return chamadas


@pytest.fixture
def doc_vazio():
    doc = fitz.open()
    yield doc
    doc.close()


def inserir_pagina_com_texto(doc: fitz.Document, texto: str, *, fontsize=14, pos=(50, 100)) -> fitz.Page:
    """Pagina de texto real (extraivel), como qualquer PDF normal."""
    page = doc.new_page()
    page.insert_text(pos, texto, fontsize=fontsize)
    return page


def inserir_pagina_so_imagem(doc: fitz.Document, cor=(230, 120, 20)) -> fitz.Page:
    """Pagina sem nenhum texto extraivel -- so uma imagem colorida, como
    uma capa 100% grafica (o caso real do arquivo "Gato Mia")."""
    page = doc.new_page()
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 200, 200))
    pix.set_rect(pix.irect, cor)
    page.insert_image(page.rect, pixmap=pix)
    return page
