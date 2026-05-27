---
name: rsu
description: Fill the Brazilian IRPF (income tax) declaration for Nu Holdings (Nubank) RSUs granted through E*TRADE/Morgan Stanley. Computes vesting acquisition cost, capital gains on sales, and the year-end position, converting USD to BRL with BCB PTAX rates. Use when the user mentions RSU, vesting, Nubank / Nu Holdings (NU) shares, E*TRADE or Morgan Stanley statements, getReleaseConfirmation, trade confirmations, ações no exterior, or declaring foreign stock / capital gains in the Brazilian income tax (IRPF, declaração de imposto de renda, carnê-leão).
---

# RSU (Nu Holdings) — Brazilian IRPF declaration

Turns E*TRADE/Morgan Stanley RSU documents into the numbers needed for the Brazilian
income-tax return: opening position, each vesting (income + cost basis), each sale
(capital gain), and the year-end position. A ready-to-fill spreadsheet template is bundled.

## Inputs the user provides

- **`getReleaseConfirmation` PDFs** — one per grant per vesting date → the **vestings**.
- **`TradeConfirmations` PDFs** — one per sale → the **sales**.
- **Prior year-end position** (saldo inicial) — quantity + average cost carried from last year.
- **`template_rsu.xlsx`** (bundled here) — the spreadsheet that does the math.

## The 5 hard rules (get these wrong and the tax is wrong)

1. **Quantity = NET shares** = the `Shares Issued` line (`Award Shares − Shares Traded`),
   NOT the gross `Shares Released`. The withheld shares were sold to pay tax.
2. **Vesting cost basis → PTAX *Venda* (sell)** of the vesting date.
3. **Sale proceeds → PTAX *Compra* (buy)** of the sale date. (Rule "compra na venda, venda na compra".)
4. **Vesting is ordinary income** (carnê-leão, ~27.5%); **sale is capital gain** (15%+).
5. Convert each lot with the PTAX **of its own date** (BCB official rate).

See [REFERENCE.md](REFERENCE.md) for the field-by-field mapping, the legal basis
(IN SRF 118/2000), and the spreadsheet cell layout.

## Workflow

1. **Extract vestings** from every `getReleaseConfirmation` PDF
   (`pdftotext -layout`). Per grant per date capture: release date, `Market Value Per
   Share`, `Award Shares`, `Shares Traded`, `Shares Issued`. Sum **net** shares per date.
2. **Extract sales** from every `TradeConfirmations` PDF: trade date, quantity, unit price.
3. **Fetch PTAX** for all vesting + sale dates from the BCB:
   `python scripts/fetch_ptax.py 2025-01-03 2025-04-01 ...` (prints compra + venda).
4. **Compute the saldo inicial** (if not given): weighted-average cost of all shares still
   held at the prior year-end = Σ(net shares × MV/share × PTAX venda) ÷ total shares.
5. **Fill the template**: `python scripts/fill_template.py template_rsu.xlsx out.xlsx data.json`
   (writes only input cells, preserves formulas). See REFERENCE.md for `data.json` shape.
6. **Validate**: Σ net vested shares across all years to the prior year-end must equal the
   broker's reported holdings; ending position = opening + vested − sold ≥ 0.

## Important

- The bundled template uses Google-Sheets functions (`SORT/FILTER`) on its calc tab —
  **open it as a Google Sheet** (in Excel those cells show `#NAME?`).
- PTAX is published only on **business days**; a weekend/holiday date returns empty —
  use the date the broker actually used.
- This is a tooling aid, **not tax advice**. Confirm method with an accountant, especially
  after Lei 14.754/2023 changed foreign-asset taxation.
