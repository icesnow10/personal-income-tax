# B3 → IRPF — Reference

Detailed reference for the `b3` skill. Not tax advice.

## 1. Average price (preço médio) algorithm

Built from the **Movimentação** history, chronologically, per **current ticker**:

1. **Ticker correction** — `code = TRIM(LEFT(Produto,6))` (treasury keeps its name). Fold FII
   subscription receipts `^([A-Z]{4})1[23]$ → \1 + "11"`. Then apply explicit **renames**
   (`ticker_memory.md`). Renames cover code changes the heuristic can't: fund
   incorporations (e.g. one fund's ticker becomes another's), BDR renames, PN→ON conversions.

2. **Classify each row** → an action, from `entry_movement` + Entrada/Saída (Credito/Debito):

| action | movements (Credito / Debito) | effect on the position |
|---|---|---|
| `purchase` | Compra, Transferência - Liquidação, Desdobro, Grupamento, Bonificação em Ativos, Fração em Ativos, Leilão de Fração, Direitos de Subscrição - Exercido, Recibo de Subscrição, Solicitação de Subscrição | +quantity, +cost |
| `sale` | (debito side of the above) · Resgate | −quantity, −cost |
| `return_of_capital` | Amortização, Restituição de Capital (+ Transferida) | **−cost**, quantity unchanged |
| `yield` | Rendimento (+ Transferido) | none (provento) |
| `dividend` | Dividendo (+ Transferido) | none (provento) |
| `interest_on_equity` | Juros Sobre Capital Próprio (+ Transferido) | none (provento) |
| `no_action` | Atualização, Cessão de Direitos, Direito(s) de Subscrição (não exercido/sobras), Transferência(s) de custódia, VENCIMENTO, COMPRA / VENDA (renda fixa) | none |

   Mechanics worth knowing: **Desdobro** = `purchase` with amount 0 → +qty, same cost (dilutes
   PM). **Grupamento** = `sale` with amount 0 → −qty (raises PM). **Resgate** is an **exit**
   (`sale`), never a purchase — a credit-side resgate that adds quantity is the classic bug.

3. **Accumulate** chronologically; `avg_price = cost_acc ÷ qty_acc` (end-of-day snapshot per
   day so same-date rows agree, like a `SUMIFS(... <= date)`). `amount_adjusted` per row =
   `+Valor` (Credito) or `−Valor` (Debito).

4. **Custody transfers** (`Transferência`/`Transferencia`, no price): matched in/out pairs on a
   date net to zero; a **lone** leg is a real quantity change (e.g. a single bonus cota).

5. **No cost-basis overrides.** Mergers / incorporações that reset the cost basis (e.g. fato
   relevante setting a patrimonial value) are NOT auto-applied — the engine sticks to the B3
   data. The **audit** at the end of the run lists every ticker whose accumulated quantity
   doesn't match the year-end Posição; the user investigates and adjusts the IRPF row by hand
   for those (typically a small number of corporate-action edge cases).

## 2. Living-memory files (markdown tables)

The two memories are the single source of truth; the script reads them (via `--memory-dir`,
default = current folder) and writes them into the `aux_mapping` sheet. Each is a markdown file
with **one data table** (extra explanatory tables are ignored — the data table is picked by its
key column). Edit the tables to teach the tool; never edit the code.

| File | Key column | Columns | Fallback |
|---|---|---|---|
| `mapping_memory.md` | `entry_movement` | `entry_movement, credito, debito, provento_type, logic` | bundled copy |
| `ticker_memory.md` | `from_ticker` | `from_ticker, to_ticker, note, source` | none → empty + warn |

The two axes per row in `mapping_memory.md`: **action** (`credito`/`debito` → purchase / sale /
return_of_capital / no_action) moves the position; **provento_type** (dividend /
interest_on_equity / yield / return_of_capital) labels the row as income for the
`income_received` summary. They are orthogonal — proventos have action = no_action but a
provento_type set.

`note`/`source` in `ticker_memory.md` are documentation (link to B3 / fato relevante), echoed
into `aux_mapping`. Handled **without** any entry: FII subscription receipts (12/13→11),
splits/grupamentos, amortizações (return of capital), and lone bonus cotas.

The bundled `mapping_memory.md` is generic and shared; `ticker_memory.md` ships as a template —
copy to your working folder and replace with your own (don't commit a taxpayer's real one).

## 3. Position blocks & IRPF mapping

The Posição export has one sheet per asset class; columns differ. The script detects the class
from the sheet name (bdr / tesouro|renda fixa / fundo / acao) and **unions all columns** into the
`Position` sheet (nothing dropped; B3's `Tipo` column is renamed `Tipo (B3)` to free the new
`tipo`). Per holding it adds `tipo`, `avg_price` (from §1), `custo_total = avg×qty`, and a
`discriminacao` text.

IRPF "Bens e Direitos" (codes as of recent years — confirm in the program):

| tipo | grupo | código | valor declared |
|---|---|---|---|
| Ação | 3 | 1 | custo_total (avg × qty) |
| BDR | 4 | 4 | custo_total |
| FII | 7 | 3 | custo_total |
| Tesouro / renda fixa | 4 | 2 | Valor Aplicado (from Posição) |

`localizacao = 105` (Brasil) for all. CNPJ: empresa (ações) / fundo (FII); BDR uses ISIN in the
discriminação; treasury has none. Edge cases to confirm with an accountant: **FI-Infra** funds
(may not be código 3) and the exact code for each fixed-income instrument.

### discriminação format
- ação/FII: `{tipo} {ticker} // {qty} UNIDADES // CUSTO MEDIO: R$ {avg} // EMPRESA: {produto} - CNPJ {cnpj} // CUSTODIA NA CORRETORA {corretora}`
- BDR: same, but `- ISIN {isin}` instead of CNPJ.
- renda fixa: `APLICACAO EM {produto} NA CORRETORA {corretora}`.

## 4. Audit (built-in)

At the end of each run, the script prints **AUDIT: N/M OK** comparing the year-end quantity
accumulated from movements vs the Posição quantity, per main ticker (renda fixa is excluded —
it is declared from Valor Aplicado / Valor Atualizado, not derived from movements). The same
comparison is in the workbook's `Reconciliation` sheet.

A mismatch means one of:
- a rename is missing (`ticker_memory.md` doesn't fold an old code into the current one)
- the corporate action isn't representable by the simple action rules (e.g. merger that resets
  cost basis, exotic restructure, conversion ratio ≠ 1:1)
- an unmapped `entry_movement` (the script prints `WARNING: unmapped …` separately)

The engine does **not** auto-override the data; surface and adjust the IRPF row by hand.

Additional checks: no `avg_price` should resemble a market quote — it must be the cost basis;
ending qty = opening + purchases − sales (± corporate actions) ≥ 0.

## 5. Files

```
b3/
├── SKILL.md
├── REFERENCE.md
├── mapping_memory.md                 # generic B3 movement→action+provento_type table (shared, bundled)
├── ticker_memory.md                  # template — renames (copy to your working folder, fill)
└── scripts/build_bens_direitos.py    # MOV.xlsx POS.xlsx OUT.xlsx [--memory-dir DIR] [--year Y]
```
