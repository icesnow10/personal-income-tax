#!/usr/bin/env python3
"""Build the IRPF "Bens e Direitos" workbook from B3 data.

Inputs
  MOV.xlsx  B3 "Movimentação" export — the HISTORICAL view (all years). One sheet,
            8 columns in B3's fixed order: Entrada/Saída, Data, Movimentação, Produto,
            Instituição, Quantidade, Preço unitário, Valor da Operação.
  POS.xlsx  B3 "Posição" export at the LAST DAY of the fiscal year (31/12/YYYY).
            Multi-sheet (e.g. Acoes, BDR, Fundo de Investimento, Tesouro Direto); each
            block may have different columns.

Usage
  python build_bens_direitos.py MOV.xlsx POS.xlsx OUT.xlsx \
        [--overrides overrides.csv] [--year 2025]

Output OUT.xlsx with 5 sheets:
  movements_to_avg_price  every movement row + computed ticker, action, qty/cost
                          accumulation and the moving average price (preço médio).
  aux_mapping             the lookup tables that drive it: movement→action, ticker
                          corrections, and the corporate-action overrides (with sources).
  avg_price_summary       one row per current ticker: latest avg price + income received.
  Position                the year-end position blocks merged into one table + avg_price,
                          custo_total (avg×qty), tipo and a ready discriminação text.
  IRPF                    one row per holding with grupo / codigo / localizacao / cnpj /
                          discriminacao / valor — ready to type into the IRPF program.

Average price = ACQUISITION COST ÷ quantity (cost basis, what Bens e Direitos asks for),
NOT market price. See REFERENCE.md. Not tax advice.

Requires: pandas, openpyxl
"""
import argparse, csv, re, unicodedata, datetime as dt
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

PURPLE, LPURPLE = "FF674EA7", "FFD9D2E9"

# ---- generic B3 vocabulary: entry_movement -> (action if Credito, action if Debito) ----
CLASSIFICATION = {
 "Dividendo": ("dividend","dividend"), "Dividendo - Transferido": ("dividend","dividend"),
 "Juros Sobre Capital Próprio": ("interest_on_equity","interest_on_equity"),
 "Juros Sobre Capital Próprio - Transferido": ("interest_on_equity","interest_on_equity"),
 "Rendimento": ("yield","yield"), "Rendimento - Transferido": ("yield","yield"),
 "Amortização": ("return_of_capital","return_of_capital"),
 "Restituição de Capital": ("return_of_capital","return_of_capital"),
 "Restituição de Capital - Transferida": ("return_of_capital","return_of_capital"),
 "Atualização": ("no_action","no_action"),
 "Cessão de Direitos": ("no_action","no_action"),
 "Cessão de Direitos - Solicitada": ("no_action","no_action"),
 "Direito de Subscrição": ("no_action","no_action"),
 "Direitos de Subscrição - Não Exercido": ("no_action","no_action"),
 "Direito Sobras de Subscrição": ("no_action","no_action"),
 "Direito Sobras de Subscrição - Não Exercido": ("no_action","no_action"),
 "Transferência": ("no_action","no_action"), "Transferencia": ("no_action","no_action"),
 "TRANSFERENCIA SEM FINANCEIRO": ("no_action","no_action"),
 "VENCIMENTO": ("no_action","no_action"), "COMPRA / VENDA": ("no_action","no_action"),
 "Compra": ("purchase","sale"), "Transferência - Liquidação": ("purchase","sale"),
 "Bonificação em Ativos": ("purchase","sale"), "Desdobro": ("purchase","sale"),
 "Grupamento": ("purchase","sale"), "Fração em Ativos": ("purchase","sale"),
 "Leilão de Fração": ("purchase","sale"),
 "Direitos de Subscrição - Exercido": ("purchase","sale"),
 "Recibo de Subscrição": ("purchase","sale"),
 "Solicitação de Subscrição": ("purchase","sale"),
 "COMPRA/VENDA DEFINITIVA A TERMO": ("purchase","sale"),
 "Resgate": ("sale","sale"),
}
LOGIC = {
 "dividend":"provento: não altera quantidade nem custo",
 "interest_on_equity":"provento (JCP): não altera quantidade nem custo",
 "yield":"rendimento: não altera quantidade nem custo",
 "no_action":"evento neutro: não altera quantidade nem custo",
 "purchase":"compra: aumenta a quantidade e o custo",
 "sale":"venda: reduz a quantidade e o custo",
 "return_of_capital":"devolução de capital: reduz o custo e mantém a quantidade",
}
# IRPF Bens e Direitos: tipo -> (grupo, codigo). 105 = Brasil (localização).
IRPF_CODE = {"Ação": (3,1), "BDR": (4,4), "FII": (7,3), "RENDA FIXA": (4,2)}


def norm(s):
    if s is None: return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().lower()
    return s.strip()

def brl(x):  # 1234.5 -> "1.234,50"
    return f"{x:,.2f}".replace(",","X").replace(".",",").replace("X",".")

def detect_tipo(sheet_name):
    n = norm(sheet_name)
    if "bdr" in n: return "BDR"
    if "tesouro" in n or "renda fixa" in n: return "RENDA FIXA"
    if "fundo" in n: return "FII"
    if "acao" in n or "acoes" in n or n == "acoes": return "Ação"
    return sheet_name  # unknown block kept as-is

def getcol(d, *candidates):
    """Value of the first matching column (accent/case-insensitive)."""
    keys = {norm(k): k for k in d}
    for c in candidates:
        k = keys.get(norm(c))
        if k is not None and d[k] not in (None, ""): return d[k]
    return None


def correct_ticker(produto, renames):
    """B3 ticker correction: TRIM(LEFT(produto,6)); fold FII subscription receipts
    (XXXX12/13 -> XXXX11); then apply explicit renames (old -> current)."""
    if not isinstance(produto, str): return None
    p = produto.strip()
    if norm(p).startswith("tesouro"): code = p[:30].strip()      # treasury: keep name
    else: code = p[:6].strip()
    m = re.match(r"^([A-Z]{4})1[23]$", code)                      # FII subscription receipt
    if m: code = m.group(1) + "11"
    return renames.get(code, code)


def load_overrides(path):
    renames, resets = {}, {}     # renames[old]=new ; resets[(ticker,date)]=(qty,avg,note,src)
    rows = []
    if not path: return renames, resets, rows
    with open(path, encoding="utf-8-sig", newline="") as f:
        for d in csv.DictReader(f):
            kind = (d.get("kind") or "").strip().lower()
            if kind == "rename":
                renames[d["from_ticker"].strip()] = d["to_ticker"].strip()
            elif kind == "cost_reset":
                key = (d["ticker"].strip(), d["date"].strip())
                resets[key] = (float(d["qty"]), float(d["avg_price"]),
                               d.get("note",""), d.get("source",""))
            rows.append(d)
    return renames, resets, rows


# ============================= 1. movements -> avg price =============================
def build_movements(mov_path, renames, resets, year):
    raw = pd.read_excel(mov_path, sheet_name=0)
    raw.columns = ["entry_type","date","entry_movement","product","holder",
                   "quantity","unit_price","amount"][:len(raw.columns)]
    raw = raw.reset_index().rename(columns={"index":"_ord"})
    raw["date"] = pd.to_datetime(raw["date"], dayfirst=True, errors="coerce")
    raw["quantity"] = pd.to_numeric(raw["quantity"], errors="coerce").fillna(0.0)
    raw["unit_price"] = pd.to_numeric(raw["unit_price"], errors="coerce")
    raw["amount"] = pd.to_numeric(raw["amount"], errors="coerce").fillna(0.0)
    raw["ticker"] = raw["product"].map(lambda p: correct_ticker(p, renames))

    unknown = sorted({m for m in raw["entry_movement"].dropna().unique()
                      if m not in CLASSIFICATION})
    def classify(r):
        cred, deb = CLASSIFICATION.get(r["entry_movement"], ("no_action","no_action"))
        return cred if str(r["entry_type"]).startswith("Cred") else deb
    raw["emt"] = raw.apply(classify, axis=1)
    raw["amount_adjusted"] = raw.apply(
        lambda r: r["amount"] if str(r["entry_type"]).startswith("Cred") else -r["amount"], axis=1)

    def deltas(r):
        e = r["emt"]
        if e == "purchase":          return ( r["quantity"],  r["amount"])
        if e == "sale":              return (-r["quantity"], -r["amount"])
        if e == "return_of_capital": return ( 0.0, -r["amount"])
        return (0.0, 0.0)
    dd = raw.apply(lambda r: pd.Series(deltas(r), index=["dq","dc"]), axis=1)
    raw = pd.concat([raw, dd], axis=1)

    # custody transfers: matched in/out pairs cancel; a lone leg is a real qty change
    raw["tr_dq"] = 0.0
    trmask = raw["emt"].eq("no_action") & raw["entry_movement"].isin(["Transferência","Transferencia"])
    net = {}
    for _, r in raw[trmask].iterrows():
        k = (r["date"], r["ticker"])
        net[k] = net.get(k, 0.0) + (r["quantity"] if str(r["entry_type"]).startswith("Cred") else -r["quantity"])
    seen = set()
    for idx, r in raw[trmask].iterrows():
        k = (r["date"], r["ticker"])
        if k in seen: continue
        seen.add(k)
        if abs(net.get(k,0.0)) > 1e-9: raw.at[idx,"tr_dq"] = net[k]
    raw["dq"] = raw["dq"] + raw["tr_dq"]

    # chronological accumulation, end-of-day snapshot, cost_reset overrides applied by date
    chrono = raw.sort_values(["date","_ord"], ascending=[True, False])
    qa, ca, cyc, closed = {}, {}, {}, {}
    QA, CA, AV, CY = {}, {}, {}, {}
    reset_by_date = {}
    for (tk, d), (qty, avg, note, src) in resets.items():
        reset_by_date.setdefault(pd.Timestamp(d), {})[tk] = (qty, avg)
    for d, g in chrono.groupby("date", sort=True):
        for _, r in g.iterrows():
            tk = r["ticker"]; cy = cyc.get(tk, 1)
            q = qa.get(tk,0.0) + r["dq"]; c = ca.get(tk,0.0) + r["dc"]
            if abs(q) < 1e-9 and qa.get(tk,0.0) > 1e-9: closed[tk] = True
            if r["dq"] > 0 and qa.get(tk,0.0) <= 1e-9 and closed.get(tk): cy += 1; closed[tk] = False
            qa[tk], ca[tk], cyc[tk] = q, c, cy
        for tk, (qty, avg) in reset_by_date.get(d, {}).items():   # corporate-action reset
            qa[tk] = qty; ca[tk] = round(qty*avg, 2); cyc[tk] = cyc.get(tk,1) + 1
        for _, r in g.iterrows():
            tk = r["ticker"]; q = qa[tk]; c = ca[tk]
            QA[r["_ord"]] = round(q,4); CA[r["_ord"]] = round(c,2)
            AV[r["_ord"]] = round(c/q,6) if abs(q) > 1e-9 else None
            CY[r["_ord"]] = cyc[tk]
    raw["quantity_accumulated"] = raw["_ord"].map(QA)
    raw["amount_adjusted"] = raw["amount_adjusted"].round(2)
    raw["avg_price"] = raw["_ord"].map(AV)
    raw["cycle"] = raw["_ord"].map(CY)

    # avg price + income per ticker as of the fiscal year end
    cut = pd.Timestamp(year, 12, 31)
    income_types = {"yield","dividend","interest_on_equity","return_of_capital"}
    summary = {}
    for tk, sub in raw.groupby("ticker"):
        s = sub[sub["date"] <= cut].sort_values(["date","_ord"])
        avg = s.iloc[-1]["avg_price"] if len(s) else None
        inc = sub[sub["emt"].isin(income_types)]["amount_adjusted"].sum()
        summary[tk] = {"avg": (None if avg is None or pd.isna(avg) else float(avg)),
                       "income": round(float(inc),2)}
    return raw.sort_values("_ord"), summary, unknown


# ============================= 2. position blocks =============================
def load_position(pos_path):
    wb = load_workbook(pos_path, data_only=True)
    blocks = []
    for sh in wb.sheetnames:
        ws = wb[sh]; tipo = detect_tipo(sh)
        headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
        rows = []
        for r in range(2, ws.max_row+1):
            vals = [ws.cell(r,c).value for c in range(1, ws.max_column+1)]
            if all(v in (None,"") for v in vals): continue
            if any(str(v).strip() == "Total" for v in vals): continue
            if vals[0] in (None,""): continue
            d = {h:(v.strip() if isinstance(v,str) else v) for h,v in zip(headers,vals)}
            d["__tipo__"] = tipo
            rows.append(d)
        blocks.append((tipo, headers, rows))
    return blocks

def discriminacao(d, summary):
    tipo = d["__tipo__"]
    prod = str(getcol(d,"Produto") or "").strip()
    qty = getcol(d,"Quantidade"); qtyi = int(qty) if isinstance(qty,(int,float)) else qty
    inst = getcol(d,"Instituição","Instituicao") or ""
    tk = getcol(d,"Código de Negociação","Codigo de Negociacao") or prod
    if tipo == "RENDA FIXA":
        return f"APLICACAO EM {tk} NA CORRETORA {inst}"
    if tipo == "BDR":
        ident = getcol(d,"Código ISIN / Distribuição","Codigo ISIN / Distribuicao","Código ISIN") or ""
        label = "ISIN"
    else:
        ident = getcol(d,"CNPJ da Empresa","CNPJ do Fundo") or ""; label = "CNPJ"
    avg = (summary.get(tk) or {}).get("avg")
    custo = f"R$ {brl(avg)}" if avg is not None else "n/d"
    return (f"{tipo} {tk} // {qtyi} UNIDADES // CUSTO MEDIO: {custo} // "
            f"EMPRESA: {prod} - {label} {ident} // CUSTODIA NA CORRETORA {inst}")


# ============================= 3. write workbook =============================
def style_header(ws, ncols, fill):
    f = PatternFill("solid", fgColor=fill)
    font = Font(bold=True, color=("FFFFFFFF" if fill == PURPLE else "FF000000"))
    for j in range(1, ncols+1):
        ws.cell(1,j).fill = f; ws.cell(1,j).font = font
        ws.cell(1,j).alignment = Alignment(horizontal="center")

def write_workbook(out_path, mov, summary, blocks, ov_rows, renames):
    wb = Workbook()

    # --- movements_to_avg_price ---
    ws = wb.active; ws.title = "movements_to_avg_price"
    old = ["entry_type","date","entry_movement","product","holder","quantity","unit_price","amount"]
    new = ["ticker","entry_movement_type","quantity_accumulated","cycle","amount_adjusted","avg_price"]
    ws.append(old+new)
    style_header(ws, 8, PURPLE)
    for j in range(9, 15):
        ws.cell(1,j).fill = PatternFill("solid",fgColor=LPURPLE); ws.cell(1,j).font = Font(bold=True)
    for _, r in mov.iterrows():
        ws.append([r["entry_type"], (r["date"].to_pydatetime() if pd.notna(r["date"]) else None),
                   r["entry_movement"], r["product"], r["holder"],
                   float(r["quantity"]), (None if pd.isna(r["unit_price"]) else float(r["unit_price"])),
                   float(r["amount"]), r["ticker"], r["emt"],
                   (None if pd.isna(r["quantity_accumulated"]) else float(r["quantity_accumulated"])),
                   (None if pd.isna(r["cycle"]) else int(r["cycle"])),
                   float(r["amount_adjusted"]),
                   (None if r["avg_price"] is None or pd.isna(r["avg_price"]) else float(r["avg_price"]))])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2): row[0].number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"

    # --- aux_mapping ---
    cs = wb.create_sheet("aux_mapping")
    cs.append(["entry_movement","Credito","Debito","logic", None,
               "from_ticker","to_ticker (correction)", None,
               "override kind","ticker/from","date","qty","avg_price","note","source"])
    for col in (1,2,3,4,6,7,9,10,11,12,13,14,15):
        cs.cell(1,col).fill = PatternFill("solid",fgColor=LPURPLE); cs.cell(1,col).font = Font(bold=True)
    movs = list(CLASSIFICATION.items())
    for i,(mv,(cr,db)) in enumerate(movs, start=2):
        cs.cell(i,1,mv); cs.cell(i,2,cr); cs.cell(i,3,db); cs.cell(i,4,LOGIC[cr])
    for i,(a,b) in enumerate(renames.items(), start=2):
        cs.cell(i,6,a); cs.cell(i,7,b)
    for i,d in enumerate(ov_rows, start=2):   # echo the overrides table (with sources)
        cs.cell(i,9,d.get("kind")); cs.cell(i,10,d.get("ticker") or d.get("from_ticker"))
        cs.cell(i,11,d.get("date")); cs.cell(i,12,d.get("qty")); cs.cell(i,13,d.get("avg_price"))
        cs.cell(i,14,d.get("note")); cs.cell(i,15,d.get("source"))
    for c,w in zip(range(1,16),[42,16,16,55,3,14,18,3,12,14,12,8,10,40,40]):
        cs.column_dimensions[get_column_letter(c)].width = w

    # --- avg_price_summary (excludes fixed income / treasury) ---
    sm = wb.create_sheet("avg_price_summary")
    sm.append(["ticker","latest_avg_price","income_received"])
    style_header(sm, 3, LPURPLE)
    # equities/FII/BDR only — fixed income (treasury/CDB) carries a name with spaces
    is_code = lambda t: isinstance(t,str) and " " not in t and len(t) <= 6
    for tk in sorted(k for k in summary if summary[k]["avg"] is not None and is_code(k)):
        s = summary[tk]
        sm.append([tk, round(s["avg"],6), s["income"]])
    for c,w in zip(range(1,4),[14,18,16]): sm.column_dimensions[get_column_letter(c)].width = w

    # --- Position (blocks merged) ---
    union = []
    for _,headers,_ in blocks:
        for h in headers:
            hh = "Tipo (B3)" if h == "Tipo" else h
            if hh not in union: union.append(hh)
    front = ["tipo","Código de Negociação","Produto","Quantidade","avg_price","custo_total"]
    rest = [h for h in union if h not in ("Código de Negociação","Produto","Quantidade")]
    cols = front + rest + ["discriminacao"]
    ps = wb.create_sheet("Position"); ps.append(cols); style_header(ps, len(cols), LPURPLE)
    for tipo,headers,rows in blocks:
        for d in rows:
            tk = getcol(d,"Código de Negociação","Codigo de Negociacao")
            qty = getcol(d,"Quantidade")
            avg = (summary.get(tk) or {}).get("avg")
            custo = round(avg*qty,2) if (avg is not None and isinstance(qty,(int,float))) else None
            line = []
            for col in cols:
                if col == "tipo": line.append(tipo)
                elif col == "avg_price": line.append(avg)
                elif col == "custo_total": line.append(custo)
                elif col == "discriminacao": line.append(discriminacao(d, summary))
                elif col == "Tipo (B3)": line.append(d.get("Tipo"))
                else: line.append(d.get(col))
            ps.append(line)
    for j,col in enumerate(cols, 1):
        ps.column_dimensions[get_column_letter(j)].width = {"Produto":48,"discriminacao":120,
            "Instituição":24,"tipo":11,"Código de Negociação":16}.get(col,15)
    ps.freeze_panes = "A2"

    # --- IRPF ---
    ir = wb.create_sheet("IRPF")
    ir.append(["ticker","grupo","codigo","localizacao","cnpj","discriminacao","valor"])
    style_header(ir, 7, LPURPLE)
    for tipo,headers,rows in blocks:
        grupo, codigo = IRPF_CODE.get(tipo, ("", ""))
        for d in rows:
            tk = getcol(d,"Código de Negociação","Codigo de Negociacao") or str(getcol(d,"Produto") or "").strip()
            cnpj = getcol(d,"CNPJ da Empresa","CNPJ do Fundo") or ""
            qty = getcol(d,"Quantidade")
            if tipo == "RENDA FIXA":
                valor = getcol(d,"Valor Aplicado")
            else:
                avg = (summary.get(getcol(d,"Código de Negociação","Codigo de Negociacao")) or {}).get("avg")
                valor = round(avg*qty,2) if (avg is not None and isinstance(qty,(int,float))) else None
            ir.append([tk, grupo, codigo, 105, cnpj, discriminacao(d, summary), valor])
    for c,w in zip(range(1,8),[12,7,7,12,18,120,14]): ir.column_dimensions[get_column_letter(c)].width = w
    ir.freeze_panes = "A2"

    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description="Build IRPF Bens e Direitos workbook from B3 data.")
    ap.add_argument("movimentacao"); ap.add_argument("posicao"); ap.add_argument("saida")
    ap.add_argument("--overrides"); ap.add_argument("--year", type=int)
    a = ap.parse_args()
    renames, resets, ov_rows = load_overrides(a.overrides)
    # fiscal year: default = year of the latest movement
    peek = pd.read_excel(a.movimentacao, sheet_name=0)
    peek.columns = ["entry_type","date","entry_movement","product","holder",
                    "quantity","unit_price","amount"][:len(peek.columns)]
    year = a.year or pd.to_datetime(peek["date"], dayfirst=True, errors="coerce").dt.year.max()
    mov, summary, unknown = build_movements(a.movimentacao, renames, resets, int(year))
    blocks = load_position(a.posicao)
    write_workbook(a.saida, mov, summary, blocks, ov_rows, renames)
    print(f"OK -> {a.saida}  (fiscal year {int(year)}, {len(mov)} movement rows, "
          f"{sum(len(r) for _,_,r in blocks)} positions)")
    if unknown:
        print("WARNING: unmapped entry_movement types (treated as no_action) — add them to "
              "CLASSIFICATION:\n  " + "\n  ".join(unknown))


if __name__ == "__main__":
    main()
