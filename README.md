# Conciliação Livro Razão × Pré-registros (Excel, só fórmulas)

![excel](https://img.shields.io/badge/Excel-100%25%20f%C3%B3rmulas-217346) ![sem macros](https://img.shields.io/badge/macros-nenhuma-lightgrey) ![testes](https://img.shields.io/badge/testes-pytest%20%2B%20LibreOffice-0A9EDC)

Ferramenta em Excel que confere, lançamento a lançamento, se o que foi **solicitado** no pré-registro do contas a pagar (conta contábil e valor) é o que de fato foi **contabilizado** no Livro Razão. Ela usa só fórmulas, sem macros e sem instalação, porque o ambiente corporativo bloqueia ambos.

> **Dados fictícios**: contas, fornecedores, IDs e colaboradores foram gerados por script.

## O problema

Cada pagamento nasce com um ID (`CP 123456` ou `ADT 123456` para adiantamentos) e uma conta contábil solicitada. No razão, esse ID aparece **no meio do histórico**, digitado de várias formas: `ID CP 123456`, `CP-123456`, `id cp  123456`, `Id:CP 123456`. Uma busca simples não encontra as variações, e conferir à mão centenas de lançamentos por mês é inviável.

## Como as fórmulas resolvem

1. **Normalização** (coluna auxiliar na aba Razao):
   ```excel
   =" "&TRIM(SUBSTITUTE(SUBSTITUTE(UPPER(D2),"-"," "),":"," "))&" "
   ```
   Converte para maiúsculas, transforma `-` e `:` em espaço e colapsa os espaços duplos (`TRIM`). Os espaços nas bordas evitam que `CP 12345` case com `CP 123456`.
2. **Localização** da chave dentro do histórico, com curinga:
   ```excel
   =IFERROR(MATCH("* "&UPPER(A2)&" *",Razao!$H$2:$H$80,0)+1,"")
   ```
3. **Status**, em ordem de gravidade:

| Status | Regra |
|---|---|
| `NÃO ENCONTRADO NO RAZÃO` | ID não aparece em nenhum histórico |
| `SEM CONTA NO HISTÓRICO` | lançamento encontrado, mas sem conta contábil |
| `CONTA DIVERGENTE` | contabilizado em conta diferente da solicitada |
| `VALOR DIVERGENTE` | conta certa, valor diferente (> R$ 0,005) |
| `OK` | conta e valor conferem |

O **Painel** resume os status e monta uma matriz **colaborador × status** (`COUNTIFS`), útil para orientar a equipe sobre classificação contábil.

## Validação

Planilhas costumam quebrar sem ninguém perceber. Por isso, a regra foi reimplementada em Python (`scripts/referencia.py`) e o teste recalcula a planilha no **LibreOffice** em modo headless e compara **linha a linha** o status das fórmulas com o do Python. No exemplo, os 60 pré-registros bateram 100% (48 OK, 6 não encontrados, 3 com valor divergente, 2 com conta divergente e 1 sem conta).

```bash
pip install -r requirements.txt
python scripts/construir_planilha.py      # gera Conciliacao_Livro_Razao.xlsx + data/*.csv
python scripts/referencia.py              # mesma regra em Python
pytest -q                                 # inclui o teste com LibreOffice (pulado se não instalado)
```

## Evolução real da ferramenta

A versão usada no trabalho passou por 9 iterações, sempre restrita a fórmulas por causa das políticas de TI. Ao longo delas, ganhou o painel por colaborador e as categorias de status. A v9 trouxe a **normalização de caracteres** e o **cruzamento com o relatório de pré-registros** pela chave ID CP/ADT, que este repositório reproduz.

## Stack

Excel (MATCH com curinga, INDEX, COUNTIFS, SUBSTITUTE/TRIM, formatação condicional) · Python (openpyxl, pandas) · LibreOffice headless · pytest · GitHub Actions

---
Autor: **Gabriel Antônio Ribeiro**, Analista Financeiro · [LinkedIn](https://www.linkedin.com/in/gabriel-ant%C3%B4nio-ribeiro-017808210)
