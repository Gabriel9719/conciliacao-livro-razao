"""Valida a planilha: as fórmulas, recalculadas pelo LibreOffice, devem produzir exatamente
o mesmo status que a implementação de referência em Python."""
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "scripts"))
from construir_planilha import construir  # noqa: E402
from referencia import conciliar, normalizar  # noqa: E402


def test_normalizacao_reconhece_variacoes():
    for h in ("FORN - ID CP 123456 - NF 1", "FORN - cp-123456", "Id:CP 123456", "id cp  123456"):
        assert " CP 123456 " in normalizar(h)
    assert " CP 12345 " not in normalizar("ID CP 123456")


def test_cenarios_plantados(tmp_path):
    construir(3, tmp_path / "c.xlsx", tmp_path / "dados")
    st = conciliar(tmp_path / "dados")["status"].value_counts()
    assert set(st.index) == {"OK", "NÃO ENCONTRADO NO RAZÃO", "VALOR DIVERGENTE", "CONTA DIVERGENTE", "SEM CONTA NO HISTÓRICO"}


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice não instalado")
def test_formulas_batem_com_referencia(tmp_path):
    construir(3, tmp_path / "c.xlsx", tmp_path / "dados")
    out = tmp_path / "recalc"
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(tmp_path / "c.xlsx")],
                   check=True, capture_output=True, timeout=180)
    ws = load_workbook(out / "c.xlsx", data_only=True)["Conferencia"]
    planilha = [r[8].value for r in ws.iter_rows(min_row=2)]
    assert planilha == list(conciliar(tmp_path / "dados")["status"])
