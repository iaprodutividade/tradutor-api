"""A heuristica antiga de cor de fundo (menor forma/area contendo o
centro do bloco) nao sabia distinguir retangulo visivel de forma
transparente, sobreposta ou artistica -- chegou a apagar a pagina inteira
de preto e um logo vetorial inteiro no arquivo real "Gato Mia" (ver
claude-sessions-log/sessions/2026-09-25_tradutor-*.md). Amostrar o pixel
renderizado de verdade elimina a classe inteira de bug: nao importa
quantas formas sobrepostas/transparentes/complexas existam, o pixel
mostra o que esta REALMENTE visivel.
"""
import fitz

from pipeline import _cor_fundo_real


def _pixmap_cor_solida(cor_rgb: tuple[int, int, int], w=200, h=200) -> fitz.Pixmap:
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, w, h))
    pix.set_rect(pix.irect, cor_rgb)
    return pix


def test_amostra_cor_solida():
    pix = _pixmap_cor_solida((255, 0, 0))
    cor = _cor_fundo_real(pix, fitz.Rect(50, 50, 100, 100))
    assert cor == (1.0, 0.0, 0.0)


def test_amostra_fundo_atras_de_texto_nao_pixel_do_proprio_bloco():
    # Fundo magenta em toda a pagina; um "bloco de texto" no meio (a
    # amostragem deve olhar pra FORA da bbox, nao pra um pixel qualquer
    # dentro dela que podia ser o proprio glifo).
    pix = _pixmap_cor_solida((200, 20, 120))
    bbox = fitz.Rect(80, 80, 120, 100)
    cor = _cor_fundo_real(pix, bbox)
    assert cor == (200 / 255, 20 / 255, 120 / 255)


def test_amostra_fora_da_pagina_nao_quebra():
    # bbox encostada na borda -- os pontos de amostragem fora dela podem
    # cair fora do pixmap; nao pode estourar excecao.
    pix = _pixmap_cor_solida((10, 10, 10), w=50, h=50)
    cor = _cor_fundo_real(pix, fitz.Rect(0, 0, 50, 50))
    assert cor is not None


def test_sem_amostra_valida_devolve_branco():
    # Pixmap minusculo, bbox gigante -- nenhum dos 4 pontos de amostragem
    # cai dentro do pixmap.
    pix = _pixmap_cor_solida((0, 0, 0), w=2, h=2)
    cor = _cor_fundo_real(pix, fitz.Rect(1000, 1000, 2000, 2000))
    assert cor == (1, 1, 1)
