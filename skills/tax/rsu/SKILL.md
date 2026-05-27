---
name: rsu
description: Fill the Brazilian IRPF (income tax) declaration for RSUs of a foreign (US-listed) company held through an equity broker. Computes each vesting's acquisition cost, capital gains on sales, and the year-end position, converting USD to BRL with the Central Bank's official PTAX rates. Use when the user mentions RSU, vesting, restricted stock, stock plan release/trade confirmations, ações no exterior, or declaring foreign stock / capital gains in the Brazilian income tax (IRPF, declaração de imposto de renda, carnê-leão, ganho de capital).
---

# RSU — Brazilian IRPF declaration

Turn an equity broker's RSU documents into the numbers the Brazilian income-tax return
needs: opening position, each vesting (income + cost basis), each sale (capital gain), and
the year-end position. Works for any foreign company's RSUs and any US equity broker.

The skill is **modular** — each concern is independent, so you can run only the part you need:

| Module | What it does | Where |
|---|---|---|
| Extraction | Pull vesting & sale data from broker PDFs | [REFERENCE.md](REFERENCE.md) §1 |
| FX (PTAX) | Fetch the Central Bank USD/BRL rates per date | `scripts/fetch_ptax.py` |
| Calculation | Net shares, cost basis, capital gain, position | [REFERENCE.md](REFERENCE.md) §2–3 |
| Template fill | Write the numbers into the spreadsheet | `scripts/fill_template.py` + `template_rsu.xlsx` |

## Inputs the user provides

- **Release confirmation** documents — one per grant per vesting date → the **vestings**.
- **Trade confirmation** documents — one per sale → the **sales**.
- **Opening position (saldo inicial)** — established in step 1 below (manual or automatic).

## First, always ask about the opening balance

Before computing anything, **ask the user how to establish the prior-year opening position
(saldo inicial)**:

- **(a) Manual** — the user types the quantity and average cost already declared at the
  prior year-end. Use these as given.
- **(b) Automatic** — the user provides **all** release confirmations (and any sales) from
  prior years, and the skill computes the opening weighted-average cost itself:
  `quantity = Σ net shares still held` and
  `avg cost = Σ(net shares × MV/share × PTAX sell) ÷ quantity`.

Do not assume — confirm which path before filling the opening row.

## The 5 rules that decide whether the tax is right

1. **Quantity = NET shares received** (gross released − shares withheld for tax), not gross.
2. **Vesting cost basis → PTAX *sell* (venda)** of the vesting date.
3. **Sale proceeds → PTAX *buy* (compra)** of the sale date. ("compra na venda, venda na compra".)
4. **Vesting is ordinary income** (carnê-leão); **the sale is capital gain** (GCAP).
5. Convert each lot with the official PTAX **of its own date**.

See [REFERENCE.md](REFERENCE.md) for field mapping, legal basis (IN SRF 118/2000), the
calculation details, and the spreadsheet layout.

## Workflow

1. **Establish the opening position** — ask the user (manual vs automatic, see above).
2. **Extract vestings** from each release confirmation: release date, grant id, grant date,
   market value per share, gross shares, withheld shares, **net shares**.
3. **Extract sales** from each trade confirmation: trade date, quantity, unit price.
4. **Fetch PTAX** for every vesting + sale date:
   `python scripts/fetch_ptax.py 2025-01-03 2025-04-01 ...` (prints buy + sell per date).
5. **Compute** each lot's BRL values (and the opening average cost, if path (b)).
6. **Fill the template**: `python scripts/fill_template.py template_rsu.xlsx out.xlsx data.json`
   (writes only input cells, preserves formulas). See REFERENCE.md for the `data.json` shape.
7. **Validate**: total net shares ever vested & still held = the broker's reported holdings;
   ending position = opening + vested − sold ≥ 0.

## Important

- The template's calc tab uses Google-Sheets functions — **open it as a Google Sheet**
  (in Excel those cells show `#NAME?`).
- PTAX is published only on **business days**; a non-business date returns empty — use the
  date the broker actually used.
- This is a tooling aid, **not tax advice**. Confirm the method with an accountant,
  especially after Lei 14.754/2023 changed foreign-asset taxation.
