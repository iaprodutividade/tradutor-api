"""Fonte de titulo com tracking largo (comum em capa/ficha tecnica feita
no Canva) grava espaco real entre CADA letra no PDF -- sem tratar isso, o
texto nao parece prosa pra IA de traducao e ela devolve inalterado, em
silencio. Achado real no arquivo "Gato Mia" (24/09/2026), ver
claude-sessions-log/sessions/2026-09-25_tradutor-*.md.
"""
from pipeline import (
    _colapsar_espacamento_artificial,
    _parece_espacamento_artificial,
    _reconstruir_espacamento_por_posicao,
)


def test_detecta_texto_letra_por_letra():
    assert _parece_espacamento_artificial("N o s s o G a t o M i a")


def test_nao_detecta_prosa_normal():
    assert not _parece_espacamento_artificial("Nosso Gato Mia é incrível")


def test_nao_detecta_texto_curto_demais():
    # menos de 4 tokens -- nao ha amostra suficiente pra decidir
    assert not _parece_espacamento_artificial("a b c")


def test_colapsar_ingenuo_gruda_tudo():
    assert _colapsar_espacamento_artificial("n o s s a a r e i a") == "nossaareia"


def test_colapsar_ingenuo_respeita_espaco_duplo():
    assert _colapsar_espacamento_artificial("n o s s a  a r e i a") == "nossa areia"


def _char(c, x0, x1, y0=0, y1=10):
    return {"c": c, "bbox": (x0, y0, x1, y1)}


def test_reconstrucao_por_posicao_gap_zero_entre_letras():
    # Padrao real da capa: letras coladas (gap 0), espaco de palavra ~6pt.
    chars = [
        _char("N", 0, 10), _char("o", 10, 20), _char("s", 20, 30), _char("s", 30, 40), _char("o", 40, 50),
        _char("G", 56, 66), _char("a", 66, 76), _char("t", 76, 86), _char("o", 86, 96),
    ]
    assert _reconstruir_espacamento_por_posicao(chars) == "Nosso Gato"


def test_reconstrucao_por_posicao_gap_uniforme_entre_letras():
    # Padrao real da ficha tecnica: gap ~6pt uniforme ATE entre letras da
    # mesma palavra, so o espaco de palavra de verdade destoa (~22pt) --
    # um limiar fixo erraria aqui (aceitaria os dois como espaco, ou
    # nenhum). Precisa do limiar ADAPTATIVO (mediana da propria linha).
    chars = [
        _char("n", 0, 10), _char("o", 16, 26), _char("s", 32, 42), _char("s", 48, 58), _char("a", 64, 74),
        _char("a", 96, 106), _char("r", 112, 122), _char("e", 128, 138), _char("i", 144, 154), _char("a", 160, 170),
    ]
    assert _reconstruir_espacamento_por_posicao(chars) == "nossa areia"


def test_reconstrucao_por_posicao_poucos_caracteres_devolve_none():
    assert _reconstruir_espacamento_por_posicao([_char("a", 0, 10), _char("b", 10, 20)]) is None


def test_reconstrucao_por_posicao_ignora_espacos_literais_do_stream():
    # Espaco literal ja presente na lista de chars (comum no rawdict) nao
    # conta pra decisao -- so a posicao dos caracteres VISIVEIS importa.
    chars = [
        _char("N", 0, 10), _char(" ", 10, 10), _char("o", 10, 20), _char("k", 20, 30),
        _char(" ", 30, 30), _char("a", 36, 46),
    ]
    assert _reconstruir_espacamento_por_posicao(chars) == "Nok a"
