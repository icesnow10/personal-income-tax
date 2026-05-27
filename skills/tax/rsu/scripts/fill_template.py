#!/usr/bin/env python3
"""Fill the RSU IRPF template (template_rsu.xlsx) from a data.json file.

Usage:
    python fill_template.py template_rsu.xlsx output.xlsx data.json

Writes ONLY the input cells and preserves every formula. Dates are formatted as
yyyy-mm-dd and the workbook is flagged to recalculate on open.

data.json shape (numbers are illustrative placeholders):
{
  "saldo_inicial": {"quantidade": 1000, "preco_medio_brl": 40.0, "custo_medio_usd": 8.0},
  "vesting": [{"date": "2025-01-03", "qty": 100, "price_usd": 10.0}, ...],   # up to 4 rows (8-11)
  "venda":   [{"date": "2025-08-28", "qty": 50, "price_usd": 12.0}, ...]     # up to 10 rows (16-25)
}

Requires: openpyxl  (pip install openpyxl)
"""
import sys
import json
import datetime as dt

import openpyxl

INPUT_SHEET = "input"          # adjust if the template renames it
PTAX_SHEET = "aux_ptax_historical_data"
DATE_FMT = "yyyy-mm-dd"
BRL_FMT = "[$R$]#,##0.0000"


def _date(iso):
    y, m, d = map(int, iso.split("-"))
    return dt.datetime(y, m, d)


def fill(template, output, data):
    wb = openpyxl.load_workbook(template, data_only=False)
    ws = wb[INPUT_SHEET]

    # 1) Saldo inicial (row 4)
    s = data.get("saldo_inicial")
    if s:
        ws["D4"] = s["quantidade"]
        ws["E4"] = s["preco_medio_brl"]
        ws["E4"].number_format = BRL_FMT     # template ships this cell as USD format
        ws["G4"] = s["custo_medio_usd"]
        ws["B4"].number_format = DATE_FMT

    # 2) Vesting (rows 8-11) — write quantity, keep pre-defined date/price unless given
    for i, v in enumerate(data.get("vesting", [])[:4]):
        r = 8 + i
        if "date" in v:
            ws[f"B{r}"] = _date(v["date"]); ws[f"B{r}"].number_format = DATE_FMT
        if "price_usd" in v:
            ws[f"E{r}"] = v["price_usd"]
        ws[f"D{r}"] = v["qty"]

    # 3) Sales (rows 16-25)
    for i, s in enumerate(data.get("venda", [])[:10]):
        r = 16 + i
        ws[f"B{r}"] = _date(s["date"]); ws[f"B{r}"].number_format = DATE_FMT
        ws[f"D{r}"] = s["qty"]
        ws[f"E{r}"] = s["price_usd"]

    # 4) Format historical dates + force recalc on open
    if PTAX_SHEET in wb.sheetnames:
        h = wb[PTAX_SHEET]
        for row in h.iter_rows(min_row=3, max_row=h.max_row, max_col=2):
            if isinstance(row[1].value, dt.datetime):
                row[1].number_format = DATE_FMT
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
