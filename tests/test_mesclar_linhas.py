"""_mesclar_linhas_mesma_altura funde entradas de 'line' do PyMuPDF que se
sobrepoem verticalmente (>50%) -- feito pra rejuntar uma unica linha
visual que o PyMuPDF fragmentou (tracking largo). Mas o mesmo limiar
tambem confunde duas linhas DIFERENTES e genuinas (fonte grande o
bastante pra bbox de uma encostar na de baixo), colando o texto sem
espaco: "Mais conforto" + "para as patinhas" virava "Mais
confortopara as patinhas". Achado real na ficha tecnica do arquivo
"Gato Mia", ver claude-sessions-log/sessions/2026-09-25_tradutor-*.md.
"""
from pipeline import _mesclar_linhas_mesma_altura


def _span(texto, font="Helvetica"):
    return {"text": texto, "font": font, "size": 10, "color": 0, "flags": 0}


def _linha(texto, y0, y1, x0=0, x1=100):
    return {"bbox": (x0, y0, x1, y1), "spans": [_span(texto)]}


def test_funde_linhas_fragmentadas_mesma_linha_visual():
    # Simula uma linha visual fragmentada em varias entradas (tracking
    # largo) -- mesma faixa vertical exata.
    linhas = [_linha("N", 10, 20), _linha("o", 10, 20), _linha("s", 10, 20)]
    mescladas = _mesclar_linhas_mesma_altura(linhas)
    assert len(mescladas) == 1


def test_insere_espaco_ao_fundir_linhas_sem_espaco_entre_si():
    linhas = [_linha("Mais conforto", 10, 20), _linha("para as patinhas", 12, 22)]
    mescladas = _mesclar_linhas_mesma_altura(linhas)
    assert len(mescladas) == 1
    texto_final = "".join(s["text"] for s in mescladas[0]["spans"])
    assert texto_final == "Mais conforto para as patinhas"


def test_nao_duplica_espaco_quando_ja_existe():
    linhas = [_linha("Mais conforto ", 10, 20), _linha("para as patinhas", 12, 22)]
    mescladas = _mesclar_linhas_mesma_altura(linhas)
    texto_final = "".join(s["text"] for s in mescladas[0]["spans"])
    assert texto_final == "Mais conforto para as patinhas"


def test_nao_funde_linhas_sem_sobreposicao_vertical():
    linhas = [_linha("Primeira linha", 10, 20), _linha("Segunda linha", 100, 110)]
    mescladas = _mesclar_linhas_mesma_altura(linhas)
    assert len(mescladas) == 2
