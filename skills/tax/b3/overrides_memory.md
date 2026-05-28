# overrides_memory — corporate-action cost resets

Living memory of corporate actions that **reset the acquisition cost** — mergers /
incorporações where the *fato relevante* defines a **new cost basis** (e.g. patrimonial value)
instead of preserving the historical cost. On the given date the ticker's position is **set** to
`(qty, qty × avg_price)`. The build script reads the table and writes it into the `aux_mapping`
sheet (with the source link, for traceability).

Handled **automatically** and therefore **not** listed here: desdobro / grupamento (quantity
only), amortização (return of capital), subscription receipts, and lone bonus cotas.

> ⚠️ **Public repo:** the row below is an **illustrative example** — replace it with your own.
> Keep your real overrides_memory.md in your working folder and run with `--memory-dir`.
> Always copy the `avg_price` from the official **fato relevante**, not from the broker screen.

| ticker | date | qty | avg_price | note | source |
|---|---|---|---|---|---|
| NEWA11 | 2025-12-12 | 20 | 83.39 | exemplo: incorporação a valor patrimonial 31/10 — CONFIRA no fato relevante | https://link-do-fato-relevante |
