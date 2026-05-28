#!/usr/bin/env python3
"""Build the IRPF "Bens e Direitos" workbook from B3 data, driven by living memory files.

Inputs
  MOV.xlsx  B3 "Movimentação" export — the HISTORICAL view (all years). One sheet, B3's fixed
            8 columns: Entrada/Saída, Data, Movimentação, Produto, Instituição, Quantidade,
            Preço unitário, Valor da Operação.
  POS.xlsx  B3 "Posição" export at the LAST DAY of the fiscal year (31/12/YYYY). Multi-sheet
            (e.g. Acoes, BDR, Fundo de Investimento, Tesouro Direto); blocks may differ.

Living memory (markdown tables — the single source of truth, pasted into aux_mapping):
  mapping_memory.md     entry_movement -> action (+ logic).   [generic; bundled fallback]
  ticker_memory.md      from_ticker -> to_ticker (renames).   [taxpayer-specific]
  overrides_memory.md   ticker,date -> qty,avg_price (cost resets for mergers). [taxpayer-specific]

Usage
  python build_bens_direitos.py MOV.xlsx POS.xlsx OUT.xlsx [--memory-dir DIR] [--year YYYY]

  --memory-dir defaults to the current folder; each memory file missing there falls back to the
  bundled copy for mapping_memory.md only (ticker/overrides default to empty, with a warning).

Average price = ACQUISITION COST ÷ quantity (cost basis, what Bens e Direitos asks for), NOT
market price. See REFERENCE.md. Not tax advice.

Requires: pandas, openpyxl
"""
import argparse, re, unicodedata
from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

PURPLE, LPURPLE = "FF674EA7", "FFD9D2E9"
SKILL_DIR = Path(__file__).resolve().parents[1]            # scripts/ -> b3/
IRPF_CODE = {"Ação": (3,1), "BDR": (4,4), "FII": (7,3), "RENDA FIXA": (4,2)}  # grupo, codigo


def norm(s):
    if s is None: return ""
    return unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().lower().strip()

def brl(x):  # 1234.5 -> "1.234,50"
    return f"{x:,.2f}".replace(",","X").replace(".",",").replace("X",".")

# --------------------------- living-memory loaders ---------------------------
def parse_md_tables(path):
    """All markdown tables in a file -> list of (header, [dict rows]). A non-pipe line ends a
    table, so a memory file may carry explanatory tables alongside the data table."""
    tables, header, cur = [], None, None
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if not s.startswith("|"):
            header = None; continue                      # table boundary
        cells = [c.strip() for c in s.strip("|").split("|")]
        if set("".join(cells)) <= set("-: "):            # separator row ---|---
            continue
        if header is None:
            header = cells; cur = []; tables.append((header, cur))
        else:
            cur.append(dict(zip(header, cells)))
    return tables

def table_with(path, key):
    """Rows of the (first) table that has column `key` — picks the data table, ignoring any
    explanatory tables in the same file."""
    if path is None: return []
    for header, rows in parse_md_tables(path):
        if key in header: return rows
    return []

def load_memory(memory_dir, warn):
    def mf(name, fallback):
        p = Path(memory_dir) / name
        if p.exists(): return p
        if fallback and (SKILL_DIR / name).exists(): return SKILL_DIR / name
        return None
    # mapping (generic; falls back to the bundled copy)
    classification, logic = {}, {}
    for d in table_with(mf("mapping_memory.md", True), "entry_movement"):
        mv = d.get("entry_movement")
        if not mv: continue
        classification[mv] = (d.get("credito","no_action"), d.get("debito","no_action"))
        logic[mv] = d.get("logic","")
    # ticker renames + cost resets (taxpayer-specific; NO fallback — empty if absent)
    renames, ren_rows = {}, []
    tpath = mf("ticker_memory.md", False)
    if tpath:
        for d in table_with(tpath, "from_ticker"):
            f, t = d.get("from_ticker"), d.get("to_ticker")
            if f and t: renames[f] = t
            ren_rows.append(d)
    else:
        warn.append("ticker_memory.md not found in --memory-dir — running with no renames "
                    f"(template: {SKILL_DIR/'ticker_memory.md'})")
    resets, ov_rows = {}, []
    opath = mf("overrides_memory.md", False)
    if opath:
        for d in table_with(opath, "ticker"):
            tk, date = d.get("ticker"), d.get("date")
            if tk and date:
                resets[(tk, date)] = (float(d["qty"]), float(d["avg_price"]),
                                      d.get("note",""), d.get("source",""))
            ov_rows.append(d)
    else:
        warn.append("overrides_memory.md not found in --memory-dir — running with no cost resets "
                    f"(template: {SKILL_DIR/'overrides_memory.md'})")
    return classification, logic, renames, resets, ren_rows, ov_rows


def correct_ticker(produto, renames):
    if not isinstance(produto, str): return None
    p = produto.strip()
    code = p[:30].strip() if norm(p).startswith("tesouro") else p[:6].strip()
    m = re.match(r"^([A-Z]{4})1[23]$", code)              # FII subscription receipt -> 11
    if m: code = m.group(1) + "11"
    return renames.get(code, code)

def detect_tipo(sheet_name):
    n = norm(sheet_name)
    if "bdr" in n: return "BDR"
    if "tesouro" in n or "renda fixa" in n: return "RENDA FIXA"
    if "fundo" in n: return "FII"
    if "aco" in n: return "Ação"
    return sheet_name

def getcol(d, *candidates):
    keys = {norm(k): k for k in d}
    for c in candidates:
        k = keys.get(norm(c))
        if k is not None and d[k] not in (None, ""): return d[k]
    return None


# =========================== 1. movements -> avg price ===========================
def build_movements(mov_path, classification, renames, resets, year):
    raw = pd.read_excel(mov_path, sheet_name=0)
    raw.columns = ["entry_type","date","entry_movement","product","holder",
                   "quantity","unit_price","amount"][:len(raw.columns)]
    raw = raw.reset_index().rename(columns={"index":"_ord"})
    raw["date"] = pd.to_datetime(raw["date"], dayfirst=True, errors="coerce")
    raw["quantity"] = pd.to_numeric(raw["quantity"], errors="coerce").fillna(0.0)
    raw["unit_price"] = pd.to_numeric(raw["unit_price"], errors="coerce")
    raw["amount"] = pd.to_numeric(raw["amount"], errors="coerce").fillna(0.0)
    raw["ticker"] = raw["product"].map(lambda p: correct_ticker(p, renames))

    unknown = sorted({m for m in raw["entry_movement"].dropna().unique() if m not in classification})
    def classify(r):
        cred, deb = classification.get(r["entry_movement"], ("no_action","no_action"))
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
    raw = pd.concat([raw, raw.apply(lambda r: pd.Series(deltas(r), index=["dq","dc"]), axis=1)], axis=1)

    # custody transfers: matched in/out pairs cancel; a lone leg is a real qty change
    raw["tr_dq"] = 0.0
    trmask = raw["emt"].eq("no_action") & raw["entry_movement"].isin(["Transferência","Transferencia"])
    net = {}
    for _, r in raw[trmask].iterrows():
        k = (r["date"], r["ticker"])
        net[k] = net.get(k,0.0) + (r["quantity"] if str(r["entry_type"]).startswith("Cred") else -r["quantity"])
    seen = set()
    for idx, r in raw[trmask].iterrows():
        k = (r["date"], r["ticker"])
        if k in seen: continue
        seen.add(k)
        if abs(net.get(k,0.0)) > 1e-9: raw.at[idx,"tr_dq"] = net[k]
    raw["dq"] = raw["dq"] + raw["tr_dq"]

    chrono = raw.sort_values(["date","_ord"], ascending=[True, False])
    qa, ca, cyc, closed = {}, {}, {}, {}
    QA, CA, AV, CY = {}, {}, {}, {}
    reset_by_date = {}
    for (tk, d), (qty, avg, *_ ) in resets.items():
        reset_by_date.setdefault(pd.Timestamp(d), {})[tk] = (qty, avg)
    for d, g in chrono.groupby("date", sort=True):
        for _, r in g.iterrows():
            tk = r["ticker"]; cy = cyc.get(tk,1)
            q = qa.get(tk,0.0) + r["dq"]; c = ca.get(tk,0.0) + r["dc"]
            if abs(q) < 1e-9 and qa.get(tk,0.0) > 1e-9: closed[tk] = True
            if r["dq"] > 0 and qa.get(tk,0.0) <= 1e-9 and closed.get(tk): cy += 1; closed[tk] = False
            qa[tk], ca[tk], cyc[tk] = q, c, cy
        for tk, (qty, avg) in reset_by_date.get(d, {}).items():
            qa[tk] = qty; ca[tk] = round(qty*avg,2); cyc[tk] = cyc.get(tk,1) + 1
        for _, r in g.iterrows():
            tk = r["ticker"]; q = qa[tk]; c = ca[tk]
            QA[r["_ord"]] = round(q,4); CA[r["_ord"]] = round(c,2)
            AV[r["_ord"]] = round(c/q,6) if abs(q) > 1e-9 else None
            CY[r["_ord"]] = cyc[tk]
    raw["quantity_accumulated"] = raw["_ord"].map(QA)
    raw["amount_adjusted"] = raw["amount_adjusted"].round(2)
    raw["avg_price"] = raw["_ord"].map(AV)
    raw["cycle"] = raw["_ord"].map(CY)

    cut = pd.Timestamp(year,12,31)
    income_types = {"yield","dividend","interest_on_equity","return_of_capital"}
    summary = {}
    for tk, sub in raw.groupby("ticker"):
        s = sub[sub["date"] <= cut].sort_values(["date","_ord"])
        avg = s.iloc[-1]["avg_price"] if len(s) else None
        inc = sub[sub["emt"].isin(income_types)]["amount_adjusted"].sum()
        summary[tk] = {"avg": (None if avg is None or pd.isna(avg) else float(avg)),
                       "income": round(float(inc),2)}
    return raw.sort_values("_ord"), summary, unknown


# =========================== 2. position blocks ===========================
def load_position(pos_path):
    wb = load_workbook(pos_path, data_only=True); blocks = []
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
            d["__tipo__"] = tipo; rows.append(d)
        blocks.append((tipo, headers, rows))
    return blocks

def discriminacao(d, summary):
    tipo = d["__tipo__"]; prod = str(getcol(d,"Produto") or "").strip()
    qty = getcol(d,"Quantidade"); qtyi = int(qty) if isinstance(qty,(int,float)) else qty
    inst = getcol(d,"Instituição","Instituicao") or ""
    tk = getcol(d,"Código de Negociação","Codigo de Negociacao") or prod
    if tipo == "RENDA FIXA":
        return f"APLICACAO EM {tk} NA CORRETORA {inst}"
    if tipo == "BDR":
        ident = getcol(d,"Código ISIN / Distribuição","Codigo ISIN / Distribuicao","Código ISIN") or ""; label = "ISIN"
    else:
        ident = getcol(d,"CNPJ da Empresa","CNPJ do Fundo") or ""; label = "CNPJ"
    avg = (summary.get(tk) or {}).get("avg")
    custo = f"R$ {brl(avg)}" if avg is not None else "n/d"
    return (f"{tipo} {tk} // {qtyi} UNIDADES // CUSTO MEDIO: {custo} // "
            f"EMPRESA: {prod} - {label} {ident} // CUSTODIA NA CORRETORA {inst}")


# =========================== 3. write workbook ===========================
def hdr(ws, n, fill):
    f = PatternFill("solid",fgColor=fill); ft = Font(bold=True, color=("FFFFFFFF" if fill==PURPLE else "FF000000"))
    for j in range(1,n+1):
        ws.cell(1,j).fill=f; ws.cell(1,j).font=ft; ws.cell(1,j).alignment=Alignment(horizontal="center")

def write_workbook(out_path, mov, summary, blocks, classification, logic, ren_rows, ov_rows):
    wb = Workbook()

    ws = wb.active; ws.title = "movements_to_avg_price"
    cols = ["entry_type","date","entry_movement","product","holder","quantity","unit_price","amount",
            "ticker","entry_movement_type","quantity_accumulated","cycle","amount_adjusted","avg_price"]
    ws.append(cols); hdr(ws, 8, PURPLE)
    for j in range(9,15):
        ws.cell(1,j).fill=PatternFill("solid",fgColor=LPURPLE); ws.cell(1,j).font=Font(bold=True)
        ws.cell(1,j).alignment=Alignment(horizontal="center")
    for _, r in mov.iterrows():
        ws.append([r["entry_type"], (r["date"].to_pydatetime() if pd.notna(r["date"]) else None),
                   r["entry_movement"], r["product"], r["holder"], float(r["quantity"]),
                   (None if pd.isna(r["unit_price"]) else float(r["unit_price"])), float(r["amount"]),
                   r["ticker"], r["emt"],
                   (None if pd.isna(r["quantity_accumulated"]) else float(r["quantity_accumulated"])),
                   (None if pd.isna(r["cycle"]) else int(r["cycle"])), float(r["amount_adjusted"]),
                   (None if r["avg_price"] is None or pd.isna(r["avg_price"]) else float(r["avg_price"]))])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2): row[0].number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"

    # aux_mapping = the three living memories pasted in (movement→action | renames | cost resets)
    cs = wb.create_sheet("aux_mapping")
    head = ["entry_movement","credito","debito","logic", None,
            "from_ticker","to_ticker","rename note","rename source", None,
            "reset ticker","reset date","qty","avg_price","reset note","reset source"]
    cs.append(head)
    for col in [1,2,3,4,6,7,8,9,11,12,13,14,15,16]:
        cs.cell(1,col).fill=PatternFill("solid",fgColor=LPURPLE); cs.cell(1,col).font=Font(bold=True)
    items = list(classification.items())
    for i,(mv,(cr,db)) in enumerate(items, start=2):
        cs.cell(i,1,mv); cs.cell(i,2,cr); cs.cell(i,3,db); cs.cell(i,4,logic.get(mv,""))
    for i,d in enumerate(ren_rows, start=2):
        cs.cell(i,6,d.get("from_ticker")); cs.cell(i,7,d.get("to_ticker"))
        cs.cell(i,8,d.get("note")); cs.cell(i,9,d.get("source"))
    for i,d in enumerate(ov_rows, start=2):
        cs.cell(i,11,d.get("ticker")); cs.cell(i,12,d.get("date")); cs.cell(i,13,d.get("qty"))
        cs.cell(i,14,d.get("avg_price")); cs.cell(i,15,d.get("note")); cs.cell(i,16,d.get("source"))
    for c,w in zip(range(1,17),[42,16,16,55,3,12,12,30,34,3,12,12,8,10,34,34]):
        cs.column_dimensions[get_column_letter(c)].width = w

    sm = wb.create_sheet("avg_price_summary")
    sm.append(["ticker","latest_avg_price","income_received"]); hdr(sm,3,LPURPLE)
    is_code = lambda t: isinstance(t,str) and " " not in t and len(t) <= 6
    for tk in sorted(k for k in summary if summary[k]["avg"] is not None and is_code(k)):
        sm.append([tk, round(summary[tk]["avg"],6), summary[tk]["income"]])
    for c,w in zip(range(1,4),[14,18,16]): sm.column_dimensions[get_column_letter(c)].width = w

    union = []
    for _,headers,_ in blocks:
        for h in headers:
            hh = "Tipo (B3)" if h == "Tipo" else h
            if hh not in union: union.append(hh)
    front = ["tipo","Código de Negociação","Produto","Quantidade","avg_price","custo_total"]
    rest = [h for h in union if h not in ("Código de Negociação","Produto","Quantidade")]
    pcols = front + rest + ["discriminacao"]
    ps = wb.create_sheet("Position"); ps.append(pcols); hdr(ps, len(pcols), LPURPLE)
    for tipo,headers,rows in blocks:
        for d in rows:
            tk = getcol(d,"Código de Negociação","Codigo de Negociacao"); qty = getcol(d,"Quantidade")
            avg = (summary.get(tk) or {}).get("avg")
            custo = round(avg*qty,2) if (avg is not None and isinstance(qty,(int,float))) else None
            line = []
            for col in pcols:
                if col == "tipo": line.append(tipo)
                elif col == "avg_price": line.append(avg)
                elif col == "custo_total": line.append(custo)
                elif col == "discriminacao": line.append(discriminacao(d, summary))
                elif col == "Tipo (B3)": line.append(d.get("Tipo"))
                else: line.append(d.get(col))
            ps.append(line)
    for j,col in enumerate(pcols,1):
        ps.column_dimensions[get_column_letter(j)].width = {"Produto":48,"discriminacao":120,
            "Instituição":24,"tipo":11,"Código de Negociação":16}.get(col,15)
    ps.freeze_panes = "A2"

    ir = wb.create_sheet("IRPF")
    ir.append(["ticker","grupo","codigo","localizacao","cnpj","discriminacao","valor"]); hdr(ir,7,LPURPLE)
    for tipo,headers,rows in blocks:
        grupo, codigo = IRPF_CODE.get(tipo, ("",""))
        for d in rows:
            tk = getcol(d,"Código de Negociação","Codigo de Negociacao") or str(getcol(d,"Produto") or "").strip()
            cnpj = getcol(d,"CNPJ da Empresa","CNPJ do Fundo") or ""; qty = getcol(d,"Quantidade")
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
    ap = argparse.ArgumentParser(description="Build IRPF Bens e Direitos workbook from B3 data + memory files.")
    ap.add_argument("movimentacao"); ap.add_argument("posicao"); ap.add_argument("saida")
    ap.add_argument("--memory-dir", default=".", help="folder with the *_memory.md files (default: current)")
    ap.add_argument("--year", type=int)
    a = ap.parse_args()
    warn = []
    classification, logic, renames, resets, ren_rows, ov_rows = load_memory(a.memory_dir, warn)
    peek = pd.read_excel(a.movimentacao, sheet_name=0)
    peek.columns = ["entry_type","date","entry_movement","product","holder",
                    "quantity","unit_price","amount"][:len(peek.columns)]
    year = a.year or int(pd.to_datetime(peek["date"], dayfirst=True, errors="coerce").dt.year.max())
    mov, summary, unknown = build_movements(a.movimentacao, classification, renames, resets, year)
    blocks = load_position(a.posicao)
    write_workbook(a.saida, mov, summary, blocks, classification, logic, ren_rows, ov_rows)
    print(f"OK -> {a.saida}  (fiscal year {year}, {len(mov)} movement rows, "
          f"{sum(len(r) for _,_,r in blocks)} positions, "
          f"{len(renames)} renames, {len(resets)} cost resets)")
    for w in warn: print("NOTE:", w)
    if unknown:
        print("WARNING: unmapped entry_movement (treated as no_action) — add to mapping_memory.md:\n  "
              + "\n  ".join(unknown))


if __name__ == "__main__":
    main()
