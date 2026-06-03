# /generate — reference

## The two sources, and who wins each field

| Field | Source of truth | Why |
|---|---|---|
| Bens e Direitos **valor** (B3 asset) | `b3_brazil_variable_income_avg_price_calculation.xlsx` (b3) | acquisition cost = preço médio × quantidade |
| Bens e Direitos **grupo / código / CNPJ** | the **informe** (`informes.json`) | official classification (e.g. NUIF11 = 07/10) |
| Bens e Direitos non-B3 item (value + codes) | the **informe** | not on the B3 at all |
| Rendimentos (isentos / exclusiva) | the **informe** | the b3 income sheets are not used here |

The script reads **only** the `IRPF_bens_e_direitos` sheet of `b3_brazil_variable_income_avg_price_calculation.xlsx`: per ticker it
takes `valor_2024`, `valor_2025`, and parses **quantidade** (`N UNIDADES`) and **preço médio**
(`CUSTO MEDIO: R$ x`) out of the discriminação for display. Everything else comes from `informes.json`.

The transcription rules (how to read each informe into `informes.json`) live in
[../read/REFERENCE.md](../read/REFERENCE.md). `/generate` only **builds**; `/completeness` **checks** —
run `/completeness` after generating (loop build → verify).

## informes.json schema (one object; every list optional)

The **same** file `/read` produces and `/completeness` consumes. `key` is the asset id (UPPERCASE for
B3, matching the b3 workbook's ticker); `source` is the PDF the value came from (tracking).

```json
{
  "bens": [
    {"key": "NUIF11", "b3": true, "grupo": 7, "codigo": 10, "cnpj": "40963403000150",
     "quantidade": 9, "valor_2025": 979.20, "source": "btg informe", "discriminacao": "<optional override>"},
    {"key": "NU-RDB", "b3": false, "grupo": 4, "codigo": 2, "localizacao": 105,
     "descr": "Saldo em RDB no Nubank", "cnpj": "30680829000143",
     "valor_2025": 88547.65, "valor_2024": 43086.29, "source": "nubank informe"}
  ],
  "isentos":   [{"codigo": 9,  "key": "BBSE3", "beneficiario": "Titular", "cnpj": "...", "fonte_pagadora": "...", "descr": "...", "valor_2025": 0.0, "source": "..."}],
  "exclusiva": [{"codigo": 10, "key": "VALE3", "beneficiario": "Titular", "cnpj": "...", "descr": "JCP ...", "valor_2025": 0.0, "source": "..."}]
}
```

### Field notes
- **`b3`** (bens): `true` → value is pulled from `b3_brazil_variable_income_avg_price_calculation.xlsx` by `key` (match is
  UPPERCASE); the JSON's `quantidade`/`valor_2025` are for `/completeness` cross-check, ignored here.
  `false` → the item is taken verbatim, including `valor_2025`/`valor_2024`.
- **`key`** of a B3 item must match the b3 workbook's `ticker` exactly (incl. `Tesouro Prefixado 2027`).
- **`source`** is carried into a `fonte` column (PDF tracking); not an IRPF field.
- **`cnpj`** raw 14 digits → formatted to `XX.XXX.XXX/XXXX-XX` and **stored as text** (leading zeros
  survive). Anything else passes through verbatim.
- **`grupo` / `codigo` / `localizacao`** are written as clean integers.
- **`localizacao`** defaults to `105` (Brasil); for exterior use the IRPF country code (e.g. `249` EUA).
- **value columns** (anything starting with `valor`) accept `1.234,56`, `1234.56`, or a number; a
  **TOTAL** row is appended per sheet.
- extra keys you add are kept (appended after the known columns).

## IRPF codes — reference only (CONFIRM each against the informe / the program)

The engine never assigns codes — it passes through what you put in the JSON, sourced from the informe.
This table is just a memory aid.

**Bens e Direitos (grupo/código, ano-base 2025):**
| ativo | grupo/código |
|---|---|
| Ação | 03 / 01 |
| BDR | 04 / 04 |
| FII / FIAGRO | 07 / 03 |
| FI-Infra incentivado (ex. NUIF11) | 07 / 10 |
| Tesouro Direto / título público | 04 / 02 |
| CDB / RDB / LCI / LCA | 04 / 02 (tributável) |
| Conta corrente / dinheiro em espécie | 06 / 01 (R$) · 06 / 02 (exterior) |
| Depósito/conta no exterior | 62 / 01 |
| Ações/ETF no exterior (custódia direta) | 07 / 99 ou 31 / … (confirmar) |
| Criptoativos | 08 / … |

**Rendimentos Isentos e Não Tributáveis (código):**
`09` lucros e dividendos (ações) · `12` rend. de caderneta de poupança/LCI/LCA/CRI/CRA · `99` outros
(**rendimento de FII / FIAgro / FI-Infra** — sem linha dedicada; NÃO use 26). Decisivo isento×exclusiva:
houve IR retido na fonte? Sem IR → isento (99); com IR → exclusiva (06).

**Rendimentos Sujeitos à Tributação Exclusiva/Definitiva (código):**
`06` rend. de aplicações financeiras (fundos comuns/come-cotas, CDB, Tesouro) · `10` JCP (juros
sobre capital próprio) · `12` ganhos em fundos. *Confirmar no programa do ano (varia entre versões).*

> Este skill **não monta** as fichas "Rendimentos Tributáveis Recebidos de PJ" nem "Rendimentos
> Recebidos no Exterior" — só Bens e Direitos, Isentos e Tributação Exclusiva.

## Algorithm
1. `load_investimentos`: ticker → {valor_2024, valor_2025, quantidade, preço médio, discriminação}.
2. `build_bens`: para cada item de `bens` do JSON — se `b3`, valor do investimentos pelo `key` +
   grupo/código/CNPJ do JSON; senão item verbatim. AUDIT lista tickers do investimentos ausentes no
   JSON.
3. `isentos` / `exclusiva`: normaliza valores/CNPJ/código (mapeia `source`→`fonte`) e escreve uma aba
   por ficha, com TOTAL.
4. Um arquivo: `irpf_consolidated.xlsx`, com as abas `bens_e_direitos`, `isentos` e
   `exclusiva_definitiva`.

Not tax advice. Every value is for human review against the informe and last year's declaration.
