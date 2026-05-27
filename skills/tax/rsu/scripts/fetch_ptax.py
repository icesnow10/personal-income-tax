#!/usr/bin/env python3
"""Fetch BCB official PTAX (compra/venda) USD/BRL for given dates.

Usage:
    python fetch_ptax.py 2025-01-03 2025-04-01 2025-08-28 ...

Prints one tab-separated line per date: <date> compra=<x> venda=<y>
PTAX is published only on business days; non-business days return "n/a".
Source: Banco Central do Brasil Olinda API (public, no auth).
"""
import sys
import json
import urllib.request

API = ("https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
       "CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='{mdy}'"
       "&$format=json&$select=cotacaoCompra,cotacaoVenda")


def ptax(date_iso):
    """date_iso = 'YYYY-MM-DD' -> (compra, venda) or (None, None)."""
    y, m, d = date_iso.split("-")
    url = API.format(mdy=f"{m}-{d}-{y}")
    with urllib.request.urlopen(url, timeout=30) as r:
        rows = json.load(r).get("value", [])
    if not rows:
        return None, None
    return rows[0]["cotacaoCompra"], rows[0]["cotacaoVenda"]


def main(argv):
    if not argv:
        print(__doc__)
        return 1
    for date_iso in argv:
        try:
            compra, venda = ptax(date_iso)
        except Exception as exc:  # network / parse errors
            print(f"{date_iso}\tERROR\t{exc}", file=sys.stderr)
            continue
        if compra is None:
            print(f"{date_iso}\tcompra=n/a\tvenda=n/a (no PTAX — non-business day?)")
        else:
            print(f"{date_iso}\tcompra={compra}\tvenda={venda}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
