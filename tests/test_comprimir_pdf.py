"""comprimir_pdf chama o Ghostscript real via subprocess -- esses testes
mockam a chamada (sem precisar do binário `gs` instalado) pra validar só a
LÓGICA: tenta /ebook primeiro, só cai pro /screen se ainda não coube,
para no primeiro preset que atingir o alvo. O comando gs em si (sintaxe
dos presets, comportamento real de compressão) precisa ser validado com
Ghostscript de verdade no preview da VPS antes de ir pra produção -- não
dá pra testar isso localmente sem o binário instalado.
"""
from pathlib import Path

import pipeline


def test_para_no_primeiro_preset_que_atinge_o_alvo(tmp_path, monkeypatch):
    def fake_run(cmd, **kwargs):
        preset = next(a for a in cmd if a.startswith("-dPDFSETTINGS="))
        output_path = Path(next(a.split("=", 1)[1] for a in cmd if a.startswith("-sOutputFile=")))
        # /ebook ja cabe no alvo dessa simulacao -- /screen nunca deveria
        # ser tentado.
        output_path.write_bytes(b"0" * (5 if preset == "-dPDFSETTINGS=/ebook" else 999))

        class Resultado:
            returncode = 0

        return Resultado()

    monkeypatch.setattr(pipeline.subprocess, "run", fake_run)

    origem = tmp_path / "origem.pdf"
    origem.write_bytes(b"%PDF-1.4 fake")
    destino = tmp_path / "saida.pdf"

    conseguiu = pipeline.comprimir_pdf(origem, destino, alvo_bytes=10)

    assert conseguiu is True
    assert destino.read_bytes() == b"0" * 5


def test_cai_para_screen_quando_ebook_nao_basta(tmp_path, monkeypatch):
    presets_tentados = []

    def fake_run(cmd, **kwargs):
        preset = next(a for a in cmd if a.startswith("-dPDFSETTINGS="))
        presets_tentados.append(preset)
        output_path = Path(next(a.split("=", 1)[1] for a in cmd if a.startswith("-sOutputFile=")))
        tamanho = 50 if preset == "-dPDFSETTINGS=/ebook" else 5
        output_path.write_bytes(b"0" * tamanho)

        class Resultado:
            returncode = 0

        return Resultado()

    monkeypatch.setattr(pipeline.subprocess, "run", fake_run)

    origem = tmp_path / "origem.pdf"
    origem.write_bytes(b"%PDF-1.4 fake")
    destino = tmp_path / "saida.pdf"

    conseguiu = pipeline.comprimir_pdf(origem, destino, alvo_bytes=10)

    assert conseguiu is True
    assert presets_tentados == ["-dPDFSETTINGS=/ebook", "-dPDFSETTINGS=/screen"]
    assert destino.read_bytes() == b"0" * 5


def test_devolve_false_mas_mantem_melhor_esforco_quando_nenhum_preset_basta(tmp_path, monkeypatch):
    def fake_run(cmd, **kwargs):
        preset = next(a for a in cmd if a.startswith("-dPDFSETTINGS="))
        output_path = Path(next(a.split("=", 1)[1] for a in cmd if a.startswith("-sOutputFile=")))
        tamanho = 80 if preset == "-dPDFSETTINGS=/ebook" else 40
        output_path.write_bytes(b"0" * tamanho)

        class Resultado:
            returncode = 0

        return Resultado()

    monkeypatch.setattr(pipeline.subprocess, "run", fake_run)

    origem = tmp_path / "origem.pdf"
    origem.write_bytes(b"%PDF-1.4 fake")
    destino = tmp_path / "saida.pdf"

    conseguiu = pipeline.comprimir_pdf(origem, destino, alvo_bytes=10)

    assert conseguiu is False
    # Mesmo sem atingir o alvo, o melhor resultado (/screen, mais
    # agressivo) fica salvo pra quem chama decidir o que fazer.
    assert destino.read_bytes() == b"0" * 40


def test_falha_do_ghostscript_e_tratada_sem_excecao(tmp_path, monkeypatch):
    def fake_run(cmd, **kwargs):
        class Resultado:
            returncode = 1

        return Resultado()

    monkeypatch.setattr(pipeline.subprocess, "run", fake_run)

    origem = tmp_path / "origem.pdf"
    origem.write_bytes(b"%PDF-1.4 fake")
    destino = tmp_path / "saida.pdf"

    conseguiu = pipeline.comprimir_pdf(origem, destino, alvo_bytes=10)

    assert conseguiu is False
    assert not destino.exists()
