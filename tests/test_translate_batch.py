"""Quando a IA devolve menos traducoes do que o pedido (mesmo depois de
repetir a chamada em lote), o fallback antigo ("out + texts[len(out):]")
assumia que o item que faltou era sempre o ULTIMO da lista. Quando o
modelo funde/pula um item no MEIO, tudo depois dele desalinhava e texto
original vazava sem traducao, em silencio -- achado real na pagina 1 do
arquivo "Gato Mia" (ver claude-sessions-log/sessions/2026-09-25_tradutor-
*.md). O fallback atual traduz item por item nesse caso, garantindo
correspondencia 1:1.
"""
import pipeline


def test_traducao_em_lote_normal(stub_translate):
    out = pipeline.translate_batch(["a", "b", "c"], "pt", "en")
    assert out == ["EN:a", "EN:b", "EN:c"]
    assert stub_translate == [["a", "b", "c"]]


def test_lista_vazia_nao_chama_api(stub_translate):
    assert pipeline.translate_batch([], "pt", "en") == []
    assert stub_translate == []


def test_divergencia_persistente_traduz_item_por_item_com_alinhamento_correto(monkeypatch):
    """Simula o cenario real: o modelo funde o item do MEIO ('b' some),
    tanto na primeira quanto na segunda tentativa em lote. O fallback tem
    que perceber a divergencia e traduzir cada item separado, sem deixar
    'c' e 'd' desalinhados com a resposta de 'b'."""
    chamadas_em_lote = []

    def fake_call_translate(texts, source_lang, target_lang):
        chamadas_em_lote.append(list(texts))
        if len(texts) > 1:
            # Funde o segundo item com o primeiro, devolvendo 1 a menos.
            fundido = [texts[0] + "+" + texts[1]] + list(texts[2:])
            return fundido, {"prompt_tokens": 0, "completion_tokens": 0}
        return [f"EN:{texts[0]}"], {"prompt_tokens": 0, "completion_tokens": 0}

    monkeypatch.setattr(pipeline, "_call_translate", fake_call_translate)

    out = pipeline.translate_batch(["a", "b", "c", "d"], "pt", "en")

    # As 2 primeiras tentativas em lote (mesma entrada, ambas divergem).
    assert chamadas_em_lote[:2] == [["a", "b", "c", "d"], ["a", "b", "c", "d"]]
    # Fallback item-por-item: cada posicao traduzida separadamente, sem
    # desalinhar 'c' e 'd' por causa do problema em 'b'.
    assert out == ["EN:a", "EN:b", "EN:c", "EN:d"]


def test_divergencia_persistente_no_ultimo_item_tambem_alinha(monkeypatch):
    def fake_call_translate(texts, source_lang, target_lang):
        if len(texts) > 1:
            return list(texts[:-1]), {"prompt_tokens": 0, "completion_tokens": 0}
        return [f"EN:{texts[0]}"], {"prompt_tokens": 0, "completion_tokens": 0}

    monkeypatch.setattr(pipeline, "_call_translate", fake_call_translate)

    out = pipeline.translate_batch(["a", "b", "c"], "pt", "en")
    assert out == ["EN:a", "EN:b", "EN:c"]
