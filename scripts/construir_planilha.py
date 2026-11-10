"""Monta a ferramenta de conciliação Livro Razão × Pré-registros (100% fórmulas, sem macros)
com dados FICTÍCIOS.

Uso: python scripts/construir_planilha.py [--seed 3] [--saida Conciliacao_Livro_Razao.xlsx]
"""
from __future__ import annotations

import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill

AZUL, VERDE = "003641", "00AE9D"
AUTOR = "Gabriel Antônio Ribeiro"
CONTAS = {
    "8.1.7.03.00.01": "Despesas de água e energia",
    "8.1.7.18.00.02": "Despesas de manutenção e conservação",
    "8.1.7.33.00.01": "Despesas de serviços de vigilância",
    "8.1.7.36.00.01": "Despesas de serviços de limpeza",
    "8.1.7.48.00.01": "Despesas de material de expediente",
    "8.1.7.63.00.04": "Despesas de serviços técnicos especializados",
    "4.9.9.92.00.01": "Adiantamentos a fornecedores",
}
FORNECEDORES = ["ENERGIA CENTRO-OESTE", "LIMPA BEM SERVIÇOS", "PROTEGE VIGILÂNCIA", "PAPELARIA TOTAL",
                "MANUTENÇÃO ÔMEGA", "CONSULTORIA ÁGIL", "SANEAMENTO ESTADUAL", "TECNOLOGIA NUVEM AZUL"]
COLABORADORES = ["Ana Souza", "Bruno Lima", "Carla Dias", "Diego Rocha"]
STATUS = ["OK", "CONTA DIVERGENTE", "VALOR DIVERGENTE", "SEM CONTA NO HISTÓRICO", "NÃO ENCONTRADO NO RAZÃO"]
CORES = {"OK": "D9F2E6", "CONTA DIVERGENTE": "F8D7DA", "VALOR DIVERGENTE": "FFF2CC",
         "SEM CONTA NO HISTÓRICO": "FCE4D6", "NÃO ENCONTRADO NO RAZÃO": "E7E6E6"}


def gerar(rng: random.Random):
    pre, razao = [], []
    contas = list(CONTAS)
    for i in range(60):
        tipo = "ADT" if rng.random() < 0.15 else "CP"
        ident = f"{tipo} {rng.randint(100000, 999999)}"
        conta = "4.9.9.92.00.01" if tipo == "ADT" else rng.choice(contas[:-1])
        valor = round(rng.uniform(150, 18000), 2)
        forn = rng.choice(FORNECEDORES)
        d = date(2026, 10, 1) + timedelta(days=rng.randint(0, 29))
        pre.append([ident, d, forn, conta, valor, rng.choice(COLABORADORES)])
        cenario = rng.random()
        if cenario < 0.06:
            continue                                                # nunca contabilizado
        conta_raz, valor_raz = conta, valor
        if cenario < 0.13:
            conta_raz = rng.choice([c for c in contas if c != conta])   # classificado na conta errada
        elif cenario < 0.18:
            valor_raz = round(valor + rng.choice([-10, 0.5, 100]), 2)    # digitação
        elif cenario < 0.23:
            conta_raz = ""                                           # lançamento sem conta
        # variações reais de digitação do ID no histórico (normalizadas pelas fórmulas)
        num = ident.split()[1]
        estilo = rng.choice([f"ID {tipo} {num}", f"{tipo}-{num}", f"id {tipo.lower()}  {num}", f"Id:{tipo} {num}"])
        razao.append([d, conta_raz, CONTAS.get(conta_raz, ""), f"{forn} - {estilo} - NF {rng.randint(1000, 99999)}",
                      valor_raz, 0.0, rng.choice(["usr01", "usr02", "usr03"])])
    for _ in range(25):                                              # lançamentos sem relação (ruído)
        c = rng.choice(contas)
        razao.append([date(2026, 10, rng.randint(1, 30)), c, CONTAS[c], "TRANSFERÊNCIA ENTRE CONTAS - AJUSTE",
                      round(rng.uniform(50, 5000), 2), 0.0, "usr09"])
    rng.shuffle(razao)
    return pre, razao


def cab(ws, cols, larg):
    ws.append(cols)
    for c, w in zip(ws[1], larg):
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", wrap_text=True); ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def construir(seed: int, saida: Path, dados: Path | None = None):
    rng = random.Random(seed)
    pre, razao = gerar(rng)
    wb = Workbook()
    wb.properties.creator = AUTOR; wb.properties.lastModifiedBy = AUTOR
    wb.properties.title = "Conciliação Livro Razão × Pré-registros (dados fictícios)"

    # ---- Razão (colar aqui o extrato do razão analítico) + coluna auxiliar normalizada
    wr = wb.active; wr.title = "Razao"
    cab(wr, ["Data", "Conta contábil", "Descrição da conta", "Histórico", "Débito", "Crédito", "Usuário",
             "Histórico normalizado (fórmula)"], [11, 16, 34, 52, 12, 10, 9, 52])
    for i, r in enumerate(razao, start=2):
        wr.append(r)
        wr.cell(i, 1).number_format = "DD/MM/YYYY"
        # normalização: maiúsculas, "-" e ":" viram espaço, espaços duplos colapsam (TRIM)
        wr.cell(i, 8, f'=" "&TRIM(SUBSTITUTE(SUBSTITUTE(UPPER(D{i}),"-"," "),":"," "))&" "')
    nr = len(razao) + 1

    # ---- Pré-registros (colar aqui o relatório)
    wp = wb.create_sheet("PreRegistros")
    cab(wp, ["ID (CP/ADT)", "Data", "Fornecedor", "Conta solicitada", "Valor", "Colaborador"], [14, 11, 26, 16, 12, 16])
    for i, r in enumerate(pre, start=2):
        wp.append(r); wp.cell(i, 2).number_format = "DD/MM/YYYY"; wp.cell(i, 5).number_format = "#,##0.00"
    npre = len(pre) + 1

    # ---- Conferência (só fórmulas)
    wc = wb.create_sheet("Conferencia", 0)
    cab(wc, ["ID", "Colaborador", "Conta solicitada", "Valor solicitado", "Linha no razão", "Conta no razão",
             "Valor no razão", "Diferença", "Status"], [14, 16, 16, 14, 10, 16, 14, 12, 26])
    for i in range(2, npre + 1):
        wc[f"A{i}"] = f"=PreRegistros!A{i}"
        wc[f"B{i}"] = f"=PreRegistros!F{i}"
        wc[f"C{i}"] = f"=PreRegistros!D{i}"
        wc[f"D{i}"] = f"=PreRegistros!E{i}"
        # procura " CP 123456 " dentro do histórico normalizado (espaços evitam casar CP 12345 com CP 123456)
        wc[f"E{i}"] = f'=IFERROR(MATCH("* "&UPPER(A{i})&" *",Razao!$H$2:$H${nr},0)+1,"")'
        wc[f"F{i}"] = f'=IF(E{i}="","",INDEX(Razao!$B:$B,E{i})&"")'
        wc[f"G{i}"] = f'=IF(E{i}="","",INDEX(Razao!$E:$E,E{i}))'
        wc[f"H{i}"] = f'=IF(E{i}="","",ROUND(G{i}-D{i},2))'
        wc[f"I{i}"] = (f'=IF(E{i}="","NÃO ENCONTRADO NO RAZÃO",IF(F{i}="","SEM CONTA NO HISTÓRICO",'
                       f'IF(F{i}<>C{i},"CONTA DIVERGENTE",IF(ABS(H{i})>0.005,"VALOR DIVERGENTE","OK"))))')
        for col in "DGH":
            wc[f"{col}{i}"].number_format = "#,##0.00"
    for st, cor in CORES.items():
        wc.conditional_formatting.add(f"I2:I{npre}", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", fgColor=cor)))
    wc.auto_filter.ref = f"A1:I{npre}"

    # ---- Painel
    wd = wb.create_sheet("Painel", 0)
    wd["A1"] = "Conciliação Livro Razão × Pré-registros"; wd["A1"].font = Font(bold=True, size=16, color=AZUL)
    wd["A2"] = "Ferramenta 100% em fórmulas: cole as bases nas abas Razao e PreRegistros e tudo recalcula."
    wd["A2"].font = Font(italic=True, color="7F8C8D")
    wd["A4"], wd["B4"] = "Status", "Qtd"
    for c in (wd["A4"], wd["B4"]):
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=AZUL)
    for r, st in enumerate(STATUS, start=5):
        wd.cell(r, 1, st).fill = PatternFill("solid", fgColor=CORES[st])
        wd.cell(r, 2, f'=COUNTIF(Conferencia!$I$2:$I${npre},A{r})')
    wd["A10"], wd["B10"] = "Total", "=SUM(B5:B9)"
    wd["A11"], wd["B11"] = "% conciliado (OK)", "=IFERROR(B5/B10,0)"
    wd["B11"].number_format = "0.0%"
    for c in (wd["A10"], wd["A11"], wd["B10"], wd["B11"]):
        c.font = Font(bold=True)
    wd["D4"] = "Colaborador"
    for j, st in enumerate(STATUS, start=5):
        wd.cell(4, j, st)
    for c in wd[4][3:4 + len(STATUS) + 1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    for r, col in enumerate(COLABORADORES, start=5):
        wd.cell(r, 4, col)
        for j in range(5, 5 + len(STATUS)):
            L = wd.cell(4, j).column_letter
            wd.cell(r, j, f'=COUNTIFS(Conferencia!$B$2:$B${npre},$D{r},Conferencia!$I$2:$I${npre},{L}$4)')
    wd.column_dimensions["A"].width = 28; wd.column_dimensions["D"].width = 16
    for L in "EFGHI":
        wd.column_dimensions[L].width = 15
    wd.row_dimensions[4].height = 32

    # ---- Instruções
    wi = wb.create_sheet("Instrucoes")
    for r, t in enumerate([
        "COMO USAR",
        "1. Cole o extrato do razão analítico na aba Razao (colunas A:G). A coluna H é fórmula: arraste até o fim.",
        "2. Cole o relatório de pré-registros na aba PreRegistros (colunas A:F).",
        "3. Arraste as fórmulas da aba Conferencia até a última linha de pré-registro.",
        "4. Filtre a coluna Status e trate as divergências. O Painel resume por status e por colaborador.",
        "",
        "REGRAS",
        "- A chave (ex.: CP 123456) é procurada dentro do histórico normalizado, com espaços nas bordas.",
        "- Normalização: maiúsculas; '-' e ':' viram espaço; espaços duplos colapsam. Assim, 'cp-123456',",
        "  'Id:CP 123456' e 'ID CP  123456' são reconhecidos como a mesma chave.",
        "- Ordem do status: não encontrado > sem conta > conta divergente > valor divergente > OK.",
        "- Sem macros e sem instalação: funciona em qualquer Excel corporativo.",
    ], start=1):
        wi.cell(r, 1, t).font = Font(bold=t.isupper(), color=AZUL if t.isupper() else "000000")
    wi.column_dimensions["A"].width = 110

    wb.calculation.fullCalcOnLoad = True
    wb.save(saida)
    if dados:
        dados.mkdir(parents=True, exist_ok=True)
        with (dados / "pre_registros.csv").open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";"); w.writerow(["id", "data", "fornecedor", "conta", "valor", "colaborador"]); w.writerows(pre)
        with (dados / "razao.csv").open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";"); w.writerow(["data", "conta", "descricao", "historico", "debito", "credito", "usuario"]); w.writerows(razao)
    return pre, razao


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--seed", type=int, default=3)
    ap.add_argument("--saida", type=Path, default=Path("Conciliacao_Livro_Razao.xlsx"))
    ap.add_argument("--dados", type=Path, default=Path("data"))
    a = ap.parse_args()
    pre, raz = construir(a.seed, a.saida, a.dados)
    print(f"{len(pre)} pré-registros, {len(raz)} linhas de razão -> {a.saida}")
