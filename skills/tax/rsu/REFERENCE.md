# RSU IRPF — Reference

## 1. Source documents (E*TRADE / Morgan Stanley)

### getReleaseConfirmation (vesting)
One PDF per **grant** per **vesting date**. Extract with `pdftotext -layout file.pdf -`.
Relevant fields (labels may be visually misaligned in the text dump — match by value):

| Field | Meaning |
|---|---|
| `Release Date` | vesting date (e.g. `01-02-2025` = MM-DD-YYYY) |
| `Market Value Per Share` | closing price USD at vest (e.g. the `$12.345678` value; `Award Price` is `$0.00`) |
| `Award Shares` | gross shares vested |
| `Shares Traded` | shares withheld & sold to pay tax (shown negative) |
| `Shares Issued` | **NET shares deposited = the number to declare** |
| `Award Number` | grant id (ES-7283, NU4235, NU18082, …) |

Sanity check: `Award Shares − Shares Traded = Shares Issued`, and
`Award Shares × Market Value Per Share = Market Value` (the gross income at vest).

### TradeConfirmations (sale)
One PDF per sale. Extract: `Trade Date`, `Quantity`, `Price` (unit, **before** the
per-order commission, e.g. $8.99), `Principal` (= qty × price).

## 2. PTAX (BCB official USD/BRL)

Use the **closing PTAX** of each event date, fetched from the Banco Central Olinda API
(`scripts/fetch_ptax.py`). Two rates exist per day:

- **PTAX Venda** (sell) → used for **acquisition / vesting** cost basis.
- **PTAX Compra** (buy) → used for **sale / alienação** proceeds.

Legal basis: IN SRF nº 118/2000 — capital gains on assets denominated in foreign currency.
Mnemonic: **"compra na venda, venda na compra"** (you buy USD at the sell rate when you
acquire; you sell USD at the buy rate when you dispose).

## 3. Calculations

Per vesting lot (net):
- `valor_acao_brl = round(market_value_usd * ptax_venda, 2)`
- `valor_total_brl = net_shares * valor_acao_brl`  ← acquisition cost (Bens e Direitos) & income

Per sale:
- `valor_acao_brl = round(sale_price_usd * ptax_compra, 2)`
- `proceeds_brl = qty_sold * valor_acao_brl`
- `gain_brl = proceeds_brl − qty_sold * average_cost_to_date_brl` (weighted moving average)

Saldo inicial (opening position) = the shares still held at prior year-end:
- `quantidade = Σ net shares vested up to and held at 31/12 of prior year`
- `preco_medio_brl = Σ(lot cost in BRL) / quantidade`
- `custo_medio_usd = Σ(net shares × MV/share) / quantidade`

Validation: opening quantity should equal the broker statement's holdings on 31/12; and
`ending = opening + vested − sold` must be ≥ 0.

## 4. Spreadsheet template (`template_rsu.xlsx`)

Grant-Thornton-style workbook. Open **as a Google Sheet** (calc tab uses `SORT/FILTER`).
Only fill the input cells below — every other cell is a formula and must be left alone.

`input` sheet:

| Section | Input cells | Notes |
|---|---|---|
| Saldo inicial (row 4) | `D4` quantity, `E4` avg price R$, `G4` avg cost USD | `F4`,`H4` are formulas |
| Vesting (rows 8–11) | `D8:D11` net quantity | dates `B`, prices `E` are pre-defined; `F` (PTAX), `G`, `H` are formulas |
| Sale (rows 16–25) | `B` date, `D` quantity, `E` unit price USD | `C`,`F`,`G`,`H` are formulas |

The `F` (PTAX) formulas `XLOOKUP` into the `aux_ptax_historical_data` sheet — column
**E (ptax_sell)** for vesting, column **D (ptax_buy)** for sales. If you rename that sheet,
update those references.

`data.json` for `scripts/fill_template.py`:

```json
{
  "saldo_inicial": {"quantidade": 1000, "preco_medio_brl": 40.0000, "custo_medio_usd": 8.0000},
  "vesting": [
    {"date": "2025-01-03", "qty": 100, "price_usd": 10.00},
    {"date": "2025-04-01", "qty": 100, "price_usd": 11.00}
  ],
  "venda": [
    {"date": "2025-08-28", "qty": 50, "price_usd": 12.00}
  ]
}
```

> The numbers above are illustrative placeholders, not real data.

## 5. Where each number lands in the IRPF

- **Bens e Direitos** — group 03 (participações societárias) / código de ações no exterior:
  the year-end position (quantity + accumulated cost in R$).
- **Vesting income** — rendimento tributável recebido de fonte no exterior (carnê-leão,
  recolhido mensalmente; the withheld "Brazil" tax on the confirmation is the credit).
- **Capital gain on sales** — ganho de capital (GCAP), 15%+ on `gain_brl`. Note the monthly
  R$35.000 alienation exemption for ações may apply — check current rules.

Not tax advice — confirm with an accountant.
