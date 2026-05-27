# RSU IRPF — Reference

Modular reference for the `rsu` skill. Each section stands alone.

## 1. Extraction module — broker documents

RSU equity platforms issue two document types (names vary slightly by broker). Extract text
with `pdftotext -layout file.pdf -` and match fields by value (labels are often misaligned).

### Release confirmation (one per grant per vesting date) → vesting
| Field | Meaning |
|---|---|
| Release / vesting date | the acquisition date |
| Grant / award date | when the grant was awarded |
| Grant / award number | grant id, for tracking which lot is which |
| Market value per share | closing price USD at vest |
| Gross shares (released/awarded) | total shares vested |
| Withheld / traded shares | shares sold to cover tax (negative) |
| **Net shares (issued/deposited)** | **the number to declare** = gross − withheld |

Sanity check: `gross − withheld = net`, and `gross × market value per share = gross income`.

### Trade confirmation (one per sale) → sale
Extract: trade date, quantity, unit price (before any per-order commission), principal
(= qty × price).

## 2. Opening balance (saldo inicial) module

Always confirm with the user which applies (see SKILL.md):

- **Manual**: user supplies `quantidade`, `preco_medio_brl`, `custo_medio_usd` as declared
  at the prior year-end. Use as-is.
- **Automatic**: from all prior-year release confirmations (and sales), compute the shares
  still held and their weighted-average cost:
  - `quantidade = Σ net shares vested up to and still held at 31/12 of the prior year`
  - `preco_medio_brl = Σ(net shares × MV/share × PTAX sell) ÷ quantidade`
  - `custo_medio_usd = Σ(net shares × MV/share) ÷ quantidade`

Validation: the opening quantity should equal the broker's reported holdings on 31/12.

## 3. FX module — PTAX (Central Bank of Brazil)

Use the official closing PTAX of each event date (`scripts/fetch_ptax.py`). Two rates/day:

- **PTAX Venda (sell)** → **acquisition / vesting** cost basis.
- **PTAX Compra (buy)** → **sale / alienação** proceeds.

Legal basis: IN SRF nº 118/2000 (capital gains on foreign-currency-denominated assets).
Mnemonic: **"compra na venda, venda na compra"**.

## 4. Calculation module

Per vesting lot (net):
- `valor_acao_brl = round(market_value_usd * ptax_venda, 2)`
- `valor_total_brl = net_shares * valor_acao_brl`  ← acquisition cost & vesting income

Per sale:
- `valor_acao_brl = round(sale_price_usd * ptax_compra, 2)`
- `proceeds_brl = qty_sold * valor_acao_brl`
- `gain_brl = proceeds_brl − qty_sold * average_cost_to_date_brl` (weighted moving average)

Ending position = `opening + Σ vested − Σ sold` (must be ≥ 0).

## 5. Template module (`template_rsu.xlsx`)

Open **as a Google Sheet** (calc tab uses Google-Sheets functions). Fill only the input
cells; everything else is a formula.

`input` sheet (per-award):

| Section | Input cells | Notes |
|---|---|---|
| Opening balance (row 4) | `D4` quantity, `E4` avg price R$, `G4` avg cost USD | `F4`,`H4` are formulas |
| Vesting, one row per award (rows 8–31) | `B` release_date, `D` award_date, `E` award_number, `F` net_quantity, `G` closing_price_usd | `C` event, `H` fx (ptax sell), `I`,`J` value formulas |
| Sale (rows 37–48) | `B` sale_date, `D` quantity_sold, `E` sale_price_usd | `C` event, `F` fx (ptax buy), `G`,`H` value formulas |

The `do_not_change__calculation_memo` sheet rebuilds the result independently: it merges the
vesting and sale rows, sorts them chronologically (`SORT(FILTER(...))`), then computes the
running quantity, the moving weighted-average cost (BRL and USD), and the capital gain per sale
(`proceeds − qty × avg_cost_to_date`). The `output` sheet reads the final position and the
total gain from it. Acquisition (vesting) uses **ptax_sell**; sale proceeds use **ptax_buy**.

`data.json` for `scripts/fill_template.py` (numbers are illustrative placeholders):

```json
{
  "saldo_inicial": {"quantidade": 1000, "preco_medio_brl": 40.0000, "custo_medio_usd": 8.0000},
  "vesting": [
    {"release_date": "2025-01-03", "award_date": "2021-10-04", "award_number": "GRANT-A", "qty": 40, "price_usd": 10.00},
    {"release_date": "2025-01-03", "award_date": "2022-03-14", "award_number": "GRANT-B", "qty": 12, "price_usd": 10.00},
    {"release_date": "2025-04-01", "award_date": "2022-03-14", "award_number": "GRANT-B", "qty": 12, "price_usd": 11.00}
  ],
  "venda": [
    {"date": "2025-08-28", "qty": 50, "price_usd": 12.00}
  ]
}
```

> The numbers above are illustrative placeholders, not real data.

## 6. Where each number lands in the IRPF

- **Bens e Direitos** — shares abroad (grupo 03 / código de ações no exterior): year-end
  position (quantity + accumulated cost in R$).
- **Vesting income** — rendimento tributável de fonte no exterior (carnê-leão; tax withheld
  at vest is the credit).
- **Capital gain on sales** — ganho de capital (GCAP), 15%+ on `gain_brl`. Check the current
  monthly exemption for small share sales.

Not tax advice — confirm with an accountant.
