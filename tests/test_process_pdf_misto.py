"""process_pdf_misto detecta pagina-imagem isolada (sem texto extraivel)
dentro de um documento majoritariamente de texto -- caso real "capa 100%
grafica dentro de um PDF de 8 paginas com texto" (arquivo "Gato Mia"),
que a classificacao binaria antiga (analisar_texto_extraivel, que so
amostra pra decidir o documento inteiro) nao pegava. Ver claude-sessions-
log/sessions/2026-09-25_tradutor-*.md.

Esses testes chamam pipeline_imagem por baixo dos panos (OCR real via
rapidocr + inpaint via iopaint) -- mais lentos que o resto da suite, mas
ainda sem custo de IA (translate_batch stubado) nem rede.
"""
from pathlib import Path

import fitz
import pytest

from pipeline import paginas_sem_texto_extraivel, process_pdf_misto

from conftest import inserir_pagina_com_texto, inserir_pagina_so_imagem


def test_paginas_sem_texto_extraivel_varre_documento_inteiro():
    doc = fitz.open()
    inserir_pagina_com_texto(doc, "Este parágrafo tem texto real de sobra.")
    inserir_pagina_so_imagem(doc)
    inserir_pagina_com_texto(doc, "E esta página também tem bastante texto.")
    try:
        assert paginas_sem_texto_extraivel(doc) == [1]
    finally:
        doc.close()


def test_paginas_sem_texto_extraivel_respeita_recorte_de_indices():
    doc = fitz.open()
    inserir_pagina_so_imagem(doc)
    inserir_pagina_so_imagem(doc)
    try:
        assert paginas_sem_texto_extraivel(doc, indices=[1]) == [1]
    finally:
        doc.close()


@pytest.mark.slow
def test_process_pdf_misto_documento_so_texto_nao_aciona_ocr(tmp_path, stub_translate):
    # Sem pagina-imagem isolada -- devolve lista vazia e nao deveria
    # precisar do pipeline de OCR/inpaint (mais lento) pra nada.
    origem = tmp_path / "origem.pdf"
    doc = fitz.open()
    inserir_pagina_com_texto(doc, "Primeira página, com texto real de verdade.")
    inserir_pagina_com_texto(doc, "Segunda página, também com texto de verdade.")
    doc.save(origem)
    doc.close()

    destino = tmp_path / "traduzido.pdf"
    paginas_imagem = process_pdf_misto(origem, destino, "pt", "en")

    assert paginas_imagem == []
    assert destino.exists()
    saida = fitz.open(destino)
    try:
        assert len(saida) == 2
        assert "EN:" in saida[0].get_text("text")
    finally:
        saida.close()


@pytest.mark.slow
def test_process_pdf_misto_detecta_e_processa_pagina_imagem_isolada(tmp_path, stub_translate):
    origem = tmp_path / "origem.pdf"
    doc = fitz.open()
    inserir_pagina_so_imagem(doc)  # capa 100% grafica, indice 0
    inserir_pagina_com_texto(doc, "Segunda página com texto real e extraível.")
    doc.save(origem)
    doc.close()

    destino = tmp_path / "traduzido.pdf"
    paginas_imagem = process_pdf_misto(origem, destino, "pt", "en")

    assert paginas_imagem == [0]
    saida = fitz.open(destino)
    try:
        # Documento final mantem as 2 paginas, na ordem original.
        assert len(saida) == 2
    finally:
        saida.close()
