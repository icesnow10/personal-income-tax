# B3 → IRPF — Reference

Detailed reference for the `b3` skill. Not tax advice.

## 1. Average price (preço médio) algorithm

Built from the **Movimentação** history, chronologically, per **current ticker**:

1. **Ticker correction** — `code = TRIM(LEFT(Produto,6))` for equity/FII (treasury keeps its name).
   Renda fixa (`CDB|RDB|CRA|CRI|CDCA|DEB|LCI|LCA|LF|LIG|LH - <código> - …`) uses the **security
   código** as the ticker (e.g. CDB422GUBFR, ENAT14, CRA025006SS). Fold FII
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
interest_on_equity / yield / return_of_capital) labels the row for the income summary. They are
orthogonal — proventos have action = no_action but a provento_type set. **Exception:**
`return_of_capital` (amortização) is devolução de capital, not income — it reduces the cost basis
and is NOT declared in any IRPF rendimento ficha (see §3 income classification).

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
| FII / FIAgro | 7 | 3 | custo_total |
| **FI-Infra** (Lei 12.431) | 7 | 10 | custo_total — "Fundos de Infraestrutura, FIDC e outros (alíquota 0%)" |
| CDB / Tesouro (tributados) | 4 | 2 | VALOR APLICADO = position quantity × acquisition unit price (never the curva). CDB: unit price = movimentação COMPRA/VENDA "Valor da Operação" ÷ qty per código. Tesouro: its own "Valor Aplicado" column |
| LCI / LCA / CRA / CRI / debênture (isentos) | 4 | 3 | declarado pelo **informe** (renda fixa não entra na ficha do b3) |

> **Renda fixa NÃO é montada pelo b3.** O b3 só reconstrói preço médio de **renda variável**
> (ações/FII/BDR) na `irpf_bens_e_direitos_variable_income`. Para CDB/CRA/CRI/debênture/Tesouro a B3 não
> tem o custo confiável (CRA/CRI/debênture embutem juros decorridos que ela não separa) — o valor de
> Bens e Direitos vem do **informe da corretora**, declarado no `informes.json` (`b3:false`).

`localizacao = 105` (Brasil) for all. CNPJ: empresa (ações) / fundo (FII); BDR uses ISIN in the
discriminação; treasury has none. **FI-Infra** is grupo 07 **código 10** (Lei 12.431, alíquota 0%
PF), NOT código 3: the script **auto-detects** it from the produto name (carries `incentivad` +
`infra`) and assigns 07/10 with a **`FI-INFRA`** label in the discriminação (instead of `FII`).
Override by hand if a fund is mis-detected. Edge case to confirm with an accountant: the exact code
for each fixed-income instrument.

### discriminação format
- ação/FII: `{tipo} {ticker} // {qty} UNIDADES // CUSTO MEDIO: R$ {avg} // EMPRESA: {produto} - CNPJ {cnpj} // CUSTODIA NA CORRETORA {corretora}`
- BDR: same, but `- ISIN {isin}` instead of CNPJ.
- renda fixa: `APLICACAO EM {produto} ({código(s)}) NA CORRETORA {corretora}`. One line per produto (Tesouros of the same name are joined). The **ticker column** shows the security código only when the produto holds a SINGLE código (e.g. ENAT14, CRA025006SS, a lone CDB); with several códigos (e.g. Banco Master) it keeps the produto name and lists every código in the discriminação.

**Auto-appended clauses (equity/FII), built from the movements for events in the fiscal year** —
all rastreável, no manual entry. Appended in this order after the base discriminação:
1. **Incorporação/conversão** — when a position came from a different fund/code via a `ticker_memory`
   rename whose origin **root differs** (e.g. IRDM11→IRIM11) and the old code's last movement is in
   the year. States the origin fund + the rename note; cost carries over (not a taxable event).
   Same-root receipts (IRIM15→IRIM11) are NOT treated as incorporation.
2. **Amortização (return of capital)** — total received in the year, per parcela (date + corretora),
   and the cost reduction (`DE R$ x PARA R$ y` only when amortization was the year's only cost change).
   States it's devolução de capital, abatida do custo, não tributável.
3. **Transferência de custódia** — a matched out-leg (Debito at origin) + in-leg (Credito at
   destination) of a `Transferência` to a DIFFERENT corretora in the year; "sem alteração de custo
   ou quantidade". Receipt codes (XXXX12/13) skipped.

### IRPF rendimentos (proventos) → ficha e código

> O b3 **não monta** as fichas de rendimento — a autoridade dos dividendos/JCP/juros é o **informe**
> da corretora/escriturador (regime de caixa, somando custódias). A aba `income` do b3 é só auditoria.
> O mapa abaixo (por `provento_type`) é referência conceitual; os valores declarados vêm do `informes.json`.

Mapping by `provento_type`:

| provento_type | natureza | ficha | código |
|---|---|---|---|
| `dividend` | dividendo de ação — isento | Rendimentos Isentos e Não Tributáveis | **09** (Lucros e dividendos) |
| `yield` | rendimento de **FII / FIAgro / FI-Infra** — isento | Rendimentos Isentos e Não Tributáveis | **99** "Outros" (sem linha dedicada; descrição "0703 - Fundos de Investimento Imobiliário") |
| `interest_on_equity` | JCP | Tributação Exclusiva/Definitiva | **10** |
| `return_of_capital` | amortização — **devolução de capital, NÃO é rendimento** | — (não vai em ficha de rendimento) | reduz o **custo** em Bens e Direitos (§1); declará-lo como renda seria duplicar |

**Teste decisivo 99 (isento) × 06 (exclusiva):** houve **IR retido na fonte**? Rendimento mensal de
FII/FIAgro/FI-Infra **não** tem IR na fonte (coluna IR do informe = "–"/zero) → **isento, código 99**.
O **código 06** ("Rendimentos de aplicações financeiras", tributação exclusiva) só vale quando o
rendimento foi efetivamente **tributado na fonte** (fundos comuns/come-cotas, CDB, Tesouro). Alguns
informes de corretora (ex.: Nubank) jogam todo provento de fundo em 06 por padrão — incorreto para
FII/FI-Infra isentos; o BTG classifica corretamente em 99.

Base legal: FII/FIAgro isentos se ≥100 cotistas, negociados em bolsa/balcão e cotista <10% (Lei
11.196/2005 + Lei 14.754/2023). FI-Infra alíquota 0% PF (Lei 12.431/2011). Amortização reduz o custo
de aquisição (devolução de capital, IN RFB 1585/2015 subseção X art. 35+).

> ⚠️ `provento_type = return_of_capital` é usado para a **redução de custo** (§1) e pode aparecer no
> resumo informativo `income`, mas **não** entra nas fichas de rendimento do IRPF (não é renda).

## 4. Audit (built-in)

At the end of each run, the script prints **AUDIT: N/M OK** comparing the year-end quantity
accumulated from movements vs the Posição quantity, per main ticker (renda fixa is excluded from
the quantity audit — it is valued as position quantity × acquisition unit price, not derived from
the movement quantity engine). The same
comparison is in the workbook's `reconciliation` sheet.

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
