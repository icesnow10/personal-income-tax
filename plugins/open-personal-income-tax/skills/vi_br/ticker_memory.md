# ticker_memory — current ticker per code

Living memory of explicit **ticker corrections** (an old or related code → the current ticker),
so the average price accumulates under one current ticker across the years. The build script
reads the table and writes it into the `aux_mapping` sheet.

The engine already folds FII **subscription receipts automatically**
(`XXXX12` / `XXXX13` → `XXXX11`) — you do **not** list those here. Only add the renames the
heuristic can't infer:

- fund **mergers / incorporações** where the ticker changes to a different code
- **BDR** renames
- **PN → ON** (or vice-versa) conversions
- any other code change for the same economic position

> ⚠️ **Public repo:** the rows below are **illustrative examples** — replace them with your own.
> Keep your real ticker_memory.md in your working folder and run the script with
> `--memory-dir <that folder>` so your tickers never land in this public repo.

| from_ticker | to_ticker | note | source |
|---|---|---|---|
| OLDA11 | NEWA11 | exemplo: FII incorporado em outro código | https://link-do-fato-relevante |
| ZZZZ33 | ZZZZ34 | exemplo: BDR renomeado | B3 |
