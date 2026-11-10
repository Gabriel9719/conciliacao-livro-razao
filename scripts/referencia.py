"""Implementação de referência (Python) da regra das fórmulas, usada para validar a planilha."""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd


def normalizar(h: str) -> str:
    return " " + re.sub(r"\s+", " ", str(h).upper().replace("-", " ").replace(":", " ")).strip() + " "


def conciliar(dados: Path) -> pd.DataFrame:
    pre = pd.read_csv(dados / "pre_registros.csv", sep=";", dtype={"conta": str})
    raz = pd.read_csv(dados / "razao.csv", sep=";", dtype={"conta": str}, keep_default_na=False)
    raz["norm"] = raz["historico"].map(normalizar)
    status = []
    for p in pre.itertuples():
        hit = raz[raz["norm"].str.contains(f" {p.id.upper()} ", regex=False)]
        if hit.empty:
            status.append("NÃO ENCONTRADO NO RAZÃO"); continue
        r = hit.iloc[0]
        if r["conta"] == "":
            status.append("SEM CONTA NO HISTÓRICO")
        elif r["conta"] != p.conta:
            status.append("CONTA DIVERGENTE")
        elif abs(float(r["debito"]) - p.valor) > 0.005:
            status.append("VALOR DIVERGENTE")
        else:
            status.append("OK")
    pre["status"] = status
    return pre


if __name__ == "__main__":
    print(conciliar(Path("data"))["status"].value_counts().to_string())
