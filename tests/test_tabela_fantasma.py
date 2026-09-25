"""find_tables() do PyMuPDF pode confundir fundo decorativo (retangulos
sobrepostos, icones alinhados) com grade de tabela, inventando uma
celula gigante cobrindo a pagina inteira -- ou ate varias celulas
"normais" sem nenhuma linha de grade desenhada de verdade. Nos dois casos
o resultado e uma redacao gigante ou texto duplicado. Achado real no
arquivo "Gato Mia" (capa e ficha tecnica), ver claude-sessions-log/
sessions/2026-09-25_tradutor-*.md.
"""
import fitz

from pipeline import _eh_linha_fina, _tabela_parece_real


def test_tabela_com_uma_celula_gigante_e_rejeitada():
    # Celula unica cobrindo praticamente a tabela inteira -- sinal claro
    # de grade fantasma (find_tables() leu 2 retangulos de fundo
    # empilhados como se fossem grade).
    tbbox = fitz.Rect(0, 0, 595, 842)
    celulas = [fitz.Rect(0, 0, 595, 842)]
    assert not _tabela_parece_real(tbbox, celulas)


def test_tabela_com_celulas_normais_e_aceita():
    # Grade real: 4 celulas, nenhuma domina a tabela inteira.
    tbbox = fitz.Rect(0, 0, 200, 200)
    celulas = [
        fitz.Rect(0, 0, 100, 100),
        fitz.Rect(100, 0, 200, 100),
        fitz.Rect(0, 100, 100, 200),
        fitz.Rect(100, 100, 200, 200),
    ]
    assert _tabela_parece_real(tbbox, celulas)


def test_tabela_sem_celulas_e_rejeitada():
    assert not _tabela_parece_real(fitz.Rect(0, 0, 100, 100), [])


def test_pagina_sem_linha_fina_nao_tem_evidencia_de_grade():
    doc = fitz.open()
    page = doc.new_page()
    # So um retangulo de fundo grande (banner colorido) -- nao e uma
    # linha/borda fina, nao deveria contar como evidencia de tabela.
    page.draw_rect(page.rect, fill=(0.9, 0.1, 0.4))
    linhas = [d for d in page.get_drawings() if _eh_linha_fina(d)]
    assert linhas == []
    doc.close()


def test_pagina_com_linha_fina_de_verdade_e_detectada():
    doc = fitz.open()
    page = doc.new_page()
    # Linha divisoria fina de verdade (borda de tabela real).
    page.draw_line((10, 10), (200, 10), width=0.5)
    linhas = [d for d in page.get_drawings() if _eh_linha_fina(d)]
    assert len(linhas) >= 1
    doc.close()
