#!/usr/bin/env python3
"""Fill the RSU IRPF template (template_rsu.xlsx) from a data.json file.

Usage:
    python fill_template.py template_rsu.xlsx output.xlsx data.json

Writes ONLY the input cells and preserves every formula. Vesting is PER-AWARD
(one row per grant per release date, rows 8-31); sales go in rows 37-48. The
calculation memo merges and sorts all events chronologically and computes the
moving weighted-average cost + capital gain on its own. Dates are formatted
yyyy-mm-dd and the workbook is flagged to recalculate on open.

data.json shape (numbers are illustrative placeholders):
{
  "saldo_inicial": {"quantidade": 1000, "preco_medio_brl": 40.0, "custo_medio_usd": 8.0},
  "vesting": [   # one entry per award per release date
    {"release_date": "2025-01-03", "award_date": "2021-10-04", "award_number": "GRANT-A",
     "qty": 41, "price_usd": 10.0},
    ...
  ],
  "venda": [{"date": "2025-08-28", "qty": 50, "price_usd": 12.0}, ...]
}

Requires: openpyxl  (pip install openpyxl)
"""
import sys
import json
import datetime as dt

import openpyxl

INPUT_SHEET = "input"
DATE_FMT = "yyyy-mm-dd"
BRL_FMT = "[$R$]#,##0.0000"
VEST = (8, 31)     # per-award vesting rows
SALE = (37, 48)    # sale rows


def _date(iso):
    y, m, d = map(int, iso.split("-"))
    return dt.datetime(y, m, d)


def fill(template, output, data):
    wb = openpyxl.load_workbook(template, data_only=False)
    ws = wb[INPUT_SHEET]

    # 1) Opening balance (row 4)
    s = data.get("saldo_inicial")
    if s:
        ws["D4"] = s["quantidade"]
        ws["E4"] = s["preco_medio_brl"]
        ws["E4"].number_format = BRL_FMT       # template ships this cell as USD format
        ws["G4"] = s["custo_medio_usd"]
        ws["B4"].number_format = DATE_FMT

    # 2) Vesting — one row per award (B release_date, D award_date, E award_number,
    #    F net qty, G price USD); C/H/I/J are formulas.
    vest = data.get("vesting", [])
    cap = VEST[1] - VEST[0] + 1
    if len(vest) > cap:
        raise SystemExit(f"{len(vest)} vesting rows but only {cap} (rows {VEST[0]}-{VEST[1]}).")
    for i, v in enumerate(vest):
        r = VEST[0] + i
        ws[f"B{r}"] = _date(v["release_date"]); ws[f"B{r}"].number_format = DATE_FMT
        ws[f"D{r}"] = _date(v["award_date"]);   ws[f"D{r}"].number_format = DATE_FMT
        ws[f"E{r}"] = v["award_number"]
        ws[f"F{r}"] = v["qty"]
        ws[f"G{r}"] = v["price_usd"]

    # 3) Sales (B date, D quantity, E unit price USD); C/F/G/H are formulas.
    sales = data.get("venda", [])
    cap = SALE[1] - SALE[0] + 1
    if len(sales) > cap:
        raise SystemExit(f"{len(sales)} sale rows but only {cap} (rows {SALE[0]}-{SALE[1]}).")
    for i, sale in enumerate(sales):
        r = SALE[0] + i
        ws[f"B{r}"] = _date(sale["date"]); ws[f"B{r}"].number_format = DATE_FMT
        ws[f"D{r}"] = sale["qty"]
        ws[f"E{r}"] = sale["price_usd"]

    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    wb.save(output)


def main(argv):
    if len(argv) != 3:
        print(__doc__)
        return 1
    template, output, data_path = argv
    with open(data_path, encoding="utf-8") as f:
        data = json.load(f)
    fill(template, output, data)
    print(f"Filled {output} from {data_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
