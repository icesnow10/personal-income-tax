---
name: vi_br
description: Reconstruct the average acquisition cost (preço médio) of B3 renda variável (ações, FIIs, BDRs) from the Movimentação history and build the year-end Bens e Direitos rows — renda variável ONLY. Outputs b3_brazil_variable_income_avg_price_calculation.xlsx. Renda fixa is NOT handled here (see the fixed_income skill). Use when the user has B3 "Movimentação"/"Posição" exports, or mentions preço médio, custo de aquisição, bens e direitos de ações/FII/BDR.
---

# B3 → IRPF "Bens e Direitos"

Turn two B3 exports into the numbers the Brazilian income-tax return needs for B3 assets:
the average acquisition cost (**preço médio**) per ticker and one ready-to-type declaration
row per holding. Works for ações, FIIs, BDRs and Tesouro Direto / renda fixa.

> **`b3` é a fonte única da verdade do `avg_price` (preço médio / custo de aquisição).** Reconstruído
> da Movimentação consolidando **todas as corretoras** e abatendo amortizações (devolução de capital).
> Os informes de banco **não** servem para custo — costumam trazer valor de mercado, não abater
> amortização (NUIF11) ou ver só uma custódia (ALZR11). Por isso `consolidate` puxa o valor daqui e
> `completeness` nunca sobrepõe o `avg_price` do b3.

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
2. **Current ticker** = `TRIM(LEFT(Produto,6))` for equity/FII, folding FII subscription receipts
   (`XXXX12`/`XXXX13` → `XXXX11`) and applying explicit **renames** from `ticker_memory.md`. For
   renda fixa (`CDB|RDB|CRA|CRI|DEB|LCI|LCA - <código> - …`) the ticker is the **security código**
   (e.g. CDB422GUBFR, ENAT14); Tesouro keeps its name. Renda fixa is excluded from the equity-only
   sheets (avg_price_summary, income, Reconciliation).
3. Movement → effect comes from `mapping_memory.md`: compra/venda move qty+cost;
   **rendimento/dividendo/JCP don't touch cost**; **amortização = devolução de capital reduces
   cost**; subscrições não exercidas, transferências de custódia, atualização = neutral.
4. **Trust the B3 data as-is.** Corporate actions (desdobro/grupamento/amortização/conversões)
   come from the Movimentação rows themselves. Mismatches against the year-end Posição are
   surfaced by the built-in **audit** (see Workflow §4) for human review — the engine does not
   auto-override the data.
5. IRPF Bens e Direitos codes: **Ação → grupo 3 / código 1**, **BDR → 4 / 4**,
   **FII → 7 / 3**, **Tesouro/renda fixa → 4 / 2**; localização **105** (Brasil).

## Living memory (the source of truth)

These markdown files drive the run. Edit them as new things appear — **no code changes**:

| File | Holds | Scope |
|---|---|---|
| `mapping_memory.md` | each B3 `entry_movement` → action (purchase/sale/no_action/…) + provento_type | generic (B3) — bundled |
| `ticker_memory.md` | renames: old/related code → current ticker (mergers, BDR/PN→ON) | taxpayer-specific |
| `rf_memory.md` | renda-fixa product renames: prior-year `Produto` name → current name (e.g. an issuer gaining "- EM LIQUIDACAO EXTRAJUDICIAL"), so `valor_<prior>` matches across the name change | taxpayer-specific |
| `rf_value_memory.md` | Bens e Direitos value override per security código for **CRA / CRI / debêntures** (amortizing / accrued-interest papers B3 can't value): the broker informe Saldo, with source | taxpayer-specific |

Keep your filled `ticker_memory.md` / `rf_memory.md` in the taxpayer's **`memory/`** folder.
`--memory-dir` **defaults to `./memory`** (created if missing), so the memory files always live in
`memory/` — never scattered in the taxpayer root. `mapping_memory.md` falls back to the bundled one if
not present; the taxpayer-specific files are seeded empty in `memory/` on first run.

## Folder layout (taxpayer folder)

The skills share a 3-tier layout, run from the taxpayer's folder:

| Folder | Holds |
|---|---|
| `resources/` | **raw** inputs: the B3 `Movimentação`/`Posição` exports + the informe PDFs |
| `memory/` | the memory files (`ticker_memory.md`, `rf_memory.md`, `rf_value_memory.md`, `mapping_memory.md`) |
| `processed/` | **derived** artifacts: `b3_brazil_variable_income_avg_price_calculation.xlsx` (this skill's output) + the transcribed JSONs |
| (root) | the deliverables: `irpf_consolidated.xlsx` + `completeness_report.md` |

## Workflow

1. **Get both exports** from B3 (Movimentação = all years; Posição = 31/12 of the fiscal year).
2. **Update `ticker_memory.md`** with this taxpayer's renames (mergers, BDR/PN→ON), each with a
   source link. Copy the bundled template into your working folder and edit there.
3. **Run** from the taxpayer folder (see "Folder layout" below):
   `python scripts/build_bens_direitos.py resources/MOV.xlsx resources/POS.xlsx processed/b3_brazil_variable_income_avg_price_calculation.xlsx --memory-dir memory --year 2025`
   (the script warns about any movement type missing from `mapping_memory.md` — add a row there).
   Optionally pass `--posicao-anterior resources/POS_PRIOR.xlsx` (the B3 Posição at 31/12 of the **prior**
   year): it fills `valor_<prev>` for renda fixa with the authoritative applied value and corrects
   prior-year quantities for corporate actions (e.g. a grupamento), and adds a
   `reconciliation_previous` sheet. Equity/FII prior cost still comes from the movements (the
   position file has market value, not acquisition cost — total cost is preserved across events).
4. **Read the AUDIT** printed at the end (also in the `reconciliation` sheet): every quantity
   mismatch between movements and the year-end position is listed by ticker. Each one is a
   corporate action the rules don't capture (a merger that resets cost basis, an exotic
   restructure, a missing rename). Investigate and adjust the IRPF row by hand — the engine
   never auto-overrides the B3 data.

## Output — `b3_brazil_variable_income_avg_price_calculation.xlsx`

Scope: **renda variável** (ações / FII / BDR) — preço médio reconstruction. **Renda fixa**
(CDB / CRA / CRI / debênture / Tesouro) does NOT get avg_price / custo here — its Bens e Direitos
value comes from the broker informe (declared via `informes.json`). Sheet order:

| Sheet | Purpose |
|---|---|
| `movements_enriched` | every movement + ticker, action, quantity_accumulated, avg_price, custo_acumulado, and an **`obs`** column ("renda fixa — sem preço médio") |
| `aux_mapping` | the two lookup tables that drove it (movement→action+provento_type, renames) |
| `avg_price_summary` | per ticker, end-of-year accumulated_quantity / avg_price / total for each year |
| `income` | per ticker, (interest, yield, total) for each year — auditing only (rendimentos are declared from the informe, not from here) |
| `position` | year-end blocks merged + avg_price, custo_total, tipo, discriminação. **Renda fixa: no custo_total, no discriminação** |
| `position_previous` | the prior-year position as-exported + reconstructed avg_price_prev / custo_total_prev (RV only; renda fixa sem custo) — only with `--posicao-anterior` |
| `reconciliation` | year-end position qty vs movement qty per ticker (RV only) |
| `reconciliation_previous` | same at 31/12 of the prior year — only with `--posicao-anterior` |
| `irpf_bens_e_direitos_variable_income` | one row per **renda variável** ticker: grupo / codigo / localizacao / cnpj / discriminacao / valor_(prior) / valor_(current) — type-ready |

The b3 workbook no longer builds `IRPF_rendimentos_isentos` / `IRPF_rendimentos_exclusivos` (the
**informe** is the authority on dividendos/JCP/juros) nor a renda-fixa value sheet.

## Important

- **Not tax advice.** Always confirm method and corporate-action treatment with an accountant.
- Bens e Direitos asks **two situations** (31/12 prior year and 31/12 current). Both are now
  built: `valor_<current>` from the year-end position × avg, and `valor_<prior>` from the cost
  basis reconstructed at 31/12 of the prior year out of the movements (blank for assets not yet
  held a year earlier; renda fixa prior value is left blank — not tracked in the movements).
  Always reconcile the prior-year column against last year's actual declaration.
- Keep this skill **generic** — never commit a real taxpayer's tickers, CNPJs, values or a
  filled `ticker_memory.md` (keep it in your working folder). See [REFERENCE.md](REFERENCE.md)
  for the tables, the memory-file format and the algorithm.
