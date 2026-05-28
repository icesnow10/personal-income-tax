---
name: b3
description: Generate the Brazilian IRPF "Bens e Direitos" from B3 exports. Computes the average acquisition cost (preço médio) per ticker from the Movimentação history and builds the year-end position plus per-asset declaration rows for ações, FIIs, BDRs and Tesouro/renda fixa. Use when the user has B3 "Movimentação" and/or "Posição" spreadsheets, or mentions preço médio, custo de aquisição, bens e direitos, ações/FII/BDR/Tesouro Direto, or declaring B3 holdings in the IRPF (declaração de imposto de renda).
---

# B3 → IRPF "Bens e Direitos"

Turn two B3 exports into the numbers the Brazilian income-tax return needs for B3 assets:
the average acquisition cost (**preço médio**) per ticker and one ready-to-type declaration
row per holding. Works for ações, FIIs, BDRs and Tesouro Direto / renda fixa.

## Inputs the user provides

- **Movimentação** (`MOV.xlsx`) — the **historical view**: B3's "Movimentação" export covering
  **all years** since the first purchase. This is what reconstructs the cost basis. One sheet,
  B3's fixed 8 columns (Entrada/Saída, Data, Movimentação, Produto, Instituição, Quantidade,
  Preço unitário, Valor da Operação).
- **Posição** (`POS.xlsx`) — the position on the **last day of the fiscal year (31/12/YYYY)**:
  B3's "Posição" export, multi-sheet (Acoes, BDR, Fundo de Investimento, Tesouro Direto). This
  gives the year-end quantities, CNPJ/ISIN and custodian to declare.

Both are downloaded from the B3 investor area (Extratos → Movimentação / Posição).

## The rules that decide whether it's right

1. **Preço médio = custo de aquisição ÷ quantidade** (cost basis) — NOT market price. Bens e
   Direitos is declared at acquisition cost.
2. **Current ticker** = `TRIM(LEFT(Produto,6))`, folding FII subscription receipts
   (`XXXX12`/`XXXX13` → `XXXX11`) and applying explicit **renames** from `ticker_memory.md`.
3. Movement → effect comes from `mapping_memory.md`: compra/venda move qty+cost;
   **rendimento/dividendo/JCP don't touch cost**; **amortização = devolução de capital reduces
   cost**; subscrições não exercidas, transferências de custódia, atualização = neutral.
4. **Corporate actions live in `overrides_memory.md`**: desdobro/grupamento change quantity only
   (automatic); a **merger/incorporação can RESET the cost basis** to a value set in the
   *fato relevante* (the historical cost is NOT preserved) — record it as a `cost_reset`.
5. IRPF Bens e Direitos codes: **Ação → grupo 3 / código 1**, **BDR → 4 / 4**,
   **FII → 7 / 3**, **Tesouro/renda fixa → 4 / 2**; localização **105** (Brasil).

## Living memory (the source of truth)

Three markdown files drive the run and are pasted into the workbook's `aux_mapping` sheet. Edit
them as new things appear — **no code changes**:

| File | Holds | Scope |
|---|---|---|
| `mapping_memory.md` | each B3 `entry_movement` → action (purchase/sale/yield/…) | generic (B3) — bundled |
| `ticker_memory.md` | renames: old/related code → current ticker (mergers, BDR/PN→ON) | taxpayer-specific |
| `overrides_memory.md` | `cost_reset`: ticker+date → new (qty, avg_price) from a *fato relevante* | taxpayer-specific |

Keep your filled `ticker_memory.md` / `overrides_memory.md` in your **working folder** and point
`--memory-dir` at it (the bundled copies are generic templates). `mapping_memory.md` falls back
to the bundled one if not present.

## Workflow

1. **Get both exports** from B3 (Movimentação = all years; Posição = 31/12 of the fiscal year).
2. **Update the memory files** for this taxpayer's renames (`ticker_memory.md`) and corporate
   actions (`overrides_memory.md`), each with a **source link** to the fato relevante. Copy the
   bundled templates into your working folder and edit there. Leave them empty if none apply.
3. **Run** from the folder holding your memory files (or pass `--memory-dir`):
   `python scripts/build_bens_direitos.py MOV.xlsx POS.xlsx OUT.xlsx --memory-dir . --year 2025`
   (the script warns about any movement type missing from `mapping_memory.md` — add a row there).
4. **Verify** (see REFERENCE.md §validation): year-end quantities in `avg_price_summary` match
   the `Posição` quantities; spot-check that no `avg_price` looks like a market quote.

## Output — `OUT.xlsx`

| Sheet | Purpose |
|---|---|
| `movements_to_avg_price` | every movement + computed ticker, action, running qty/cost, avg |
| `aux_mapping` | the lookup tables that drove it (movement→action, renames, overrides+sources) |
| `avg_price_summary` | one row per current ticker: latest preço médio + income received |
| `Position` | year-end blocks merged + avg_price, custo_total, tipo, discriminação |
| `IRPF` | ticker / grupo / codigo / localizacao / cnpj / discriminacao / valor — type-ready |

## Important

- **Not tax advice.** Always confirm method and corporate-action treatment with an accountant.
- A **cost_reset** value (merger/incorporação) must come from the fund's **fato relevante** —
  verify it; the program/broker often shows a different number.
- Bens e Direitos asks **two situations** (31/12 prior year and 31/12 current). This builds the
  **current** value; the prior-year value comes from last year's declaration (and differs for
  any ticker touched by a 2025 event).
- Keep this skill **generic** — never commit a real taxpayer's tickers, CNPJs, values or a
  filled `ticker_memory.md` / `overrides_memory.md` (keep those in your working folder). See
  [REFERENCE.md](REFERENCE.md) for the tables, the memory-file format and the algorithm.
