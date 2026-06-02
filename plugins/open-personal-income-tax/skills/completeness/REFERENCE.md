# /completeness — reference

## The two sources

| Source | What it is | How it's read |
|---|---|---|
| **b3_source** | `IRPF_bens_e_direitos` sheet of `brazil_investments.xlsx` (the **b3** skill output) | `compare` reads it directly |
| **informe_de_rendimentos** | broker/bank PDFs (BTG, NU, Clear, Wise, BB, foreign brokers) | agent via [/read](../read/SKILL.md) → `informes.json` |

**Source of truth is per ficha.** In **Bens e Direitos**, on a VALUE divergence the b3_source figure
(consolidated cost basis across custodians, amortização abatida) prevails — the report flags the gap,
never rewrites it; the grupo/código/CNPJ come from the informe. In **Rendimentos Isentos (09/99)** and
**Tributação Exclusiva (06/10)**, the **informe** is the authority on both value and classification
(regime de caixa, somando corretoras + escriturador) — the b3 income sheets are partial and never
declared, so on a rendimento divergence the **informe wins**.

The script does NOT parse the informes into items — layouts differ too much per broker and many are
images. The agent extracts (reads PDFs and reasons); the script only matches, diffs, and writes the `.md`.

## Normalized schema (informes.json) — built by /read
The unified transcription is a single object with `bens` / `isentos` / `exclusiva` lists (the **same**
file `/generate` consumes). Full schema in [../generate/REFERENCE.md](../generate/REFERENCE.md):
```json
{
  "bens":      [{"key":"VALE3","b3":true,"grupo":3,"codigo":1,"cnpj":"33592510000154","quantidade":20,"source":"btg_787"},
                {"key":"NU-RESERVA","b3":false,"grupo":4,"codigo":2,"valor_2025":52777.03,"source":"nubank_informe"}],
  "isentos":   [{"codigo":9,"key":"VALE3","valor_2025":42.83,"source":"nubank_informe"}],
  "exclusiva": [{"codigo":10,"key":"VALE3","valor_2025":41.07,"source":"btg_787 + nubank_informe"}]
}
```
- `key`: ticker (`VALE3`), security código (`CRA02300MZT`, `ENAT14`), or a coined short id for
  non-market assets (`NU-RESERVA`, `RDB-NUCONTA`, `EUR-WISE`). UPPERCASE, no spaces.
- Use the **same key** the b3 workbook uses (its `ticker` column) for B3 assets — that's what lines
  the two views up.
- For `bens`, `valor_2025`/`quantidade` are the informe's cross-check values (omit `valor_2025` for
  ações/FII where the informe shows only quantidade). For rendimentos, `valor_2025` is the reconciled field.

How to read each informe and decide ficha/código → see [../read/REFERENCE.md](../read/REFERENCE.md).
The tips below are the **matching** angle (lining the informe up with b3_source).

## Extraction / matching tips
- Run `extract <docs_dir>` first; it prints lines that carry a money value plus a CNPJ / ticker /
  fixed-income código / asset keyword. Those are the rows to transcribe.
- **BTG** informe has a clean table: `conta código CNPJ grupo cód saldo2024 saldo2025 … rendimento`.
  Read `saldo2025` as `valor_2025`.
- **Nubank**: renda fixa transferred to another broker shows `saldo2025 = 0` (declared at the new
  broker) — don't double count. The bank fund (Reserva Imediata), the NuConta/RDB balance and the
  Conta Global (moeda estrangeira) are NU-only items to add.
- **Encrypted/garbled PDFs**: `extract` returns nothing — transcribe the visible Bens e Direitos
  values by hand.
- **Conta Global do Nubank = whitelabel da Wise** — o extrato vem com a marca **Wise** (saldo em
  moeda estrangeira). A **Wise Brasil** (CNPJ 40.571.694/0001-31) é a conta BRL e costuma ter saldo 0
  — não confundir: a moeda estrangeira está na conta global (Wise/Nubank Global), não na Wise BR.
- An informe may list assets held at **one custodian only**; the same ticker can sit at another
  broker too. b3_source (B3) is the consolidated position — a custódia-split is the common reason an
  informe value (one broker) differs from b3_source (all custodians). **b3_source wins.**

## Keying renda fixa to match b3_source
The b3 workbook labels a renda-fixa line by the **security código** only when the produto holds a
SINGLE código (e.g. `CDBC258QL7E`, `CRA02300MZT`, `ENAT14`); when a bank holds **several** CDB códigos
it uses the **produto/bank name** (`CDB - BANCO MASTER S/A …`) and lists the códigos in the
discriminação. So when building `informes.json`:
- single-código papers → key by the **código** (matches the workbook ticker).
- multi-código banks (Master, Fibra, Original…) → **sum the códigos** under ONE item keyed by the
  bank/produto exactly as the workbook shows it, else each código reports a false "FALTA em
  b3_source (B3!)".
- Tesouro: the workbook uses the produto name (`Tesouro Prefixado 2027`), the informe may say `LTN` —
  key both the same.

## Matching — código OR discriminação
An informe line may show **no código** — only the discriminação. So matching is **not** key-only: each
item yields identity tokens = its `key`, any ticker/código embedded in the `descr`, **and a normalized
produto signature** (`descr_sig`: drops `APLICACAO EM`, the `(código)`, and the `NA CORRETORA …` / `// …`
tail). Items that share **any** token are the same asset (union-find); a short signature that is a
PREFIX of a longer one is unified too. **Always put the full discriminação text in `descr`.** Papers
that are **zero everywhere** (matured/redeemed, saldo 0,00) are dropped — not flagged.

## Tolerance
- Two values are "equal" within `--tol` reais (default 0.50) OR 0.5% of the value, whichever is larger
  — absorbs rounding/curva noise.
- `looks_b3(key)` decides whether a missing item is flagged `(B3!)` (a real workbook bug to chase) or
  `(não-B3)` (expected: declare by hand). Heuristic: ticker `^[A-Z]{4}\d{1,2}$`, a fixed-income código,
  or a CDB/CRA/CRI/DEB/LCI/LCA/RDB/Tesouro prefix; class `rv`/`rf_simples`/`cra_cri_deb` counts as B3.

## Common divergence causes (for triage)
- **Custódia split** — informe shows one broker's quantity; b3_source is the consolidated B3 position.
- **Corporate action** — grupamento/desdobro/incorporação; cost should be preserved (the b3 skill
  preserves total cost; some informes don't, e.g. a grupamento left at the old per-share price).
- **Renda fixa**: declare **valor aplicado** (qty × acquisition price), not the **curva**. CRA/CRI/
  debêntures embed **juros decorridos** and amortize — only the broker informe has the exact saldo.

## Output — a single `.md`, organized by IRPF ficha
`completeness_report.md` (and a console summary). No `.xlsx`. Sections:
- **Bens e Direitos** — `grupo | código | asset | qtd b3 | qtd informe | valor b3 | valor informe | fonte (PDF) | status`.
  Código mismatch (b3 × informe) is flagged inline (e.g. NUIF11 `⚠️b3=3/inf=10`); a TOTAL-a-declarar line counts the não-B3 items to add by hand.
- **Rendimentos Isentos** (cód 09/99) and **Tributação Exclusiva** (cód 06/10) — per código, with a **SOMA por código** (so a per-item gap that reconciles at the total is visible).
- **Divergências de classificação** — same asset+value under a different código between b3 and informe (reclassificação, not missing value).
- **Resumo** + **Action items**.
Two cross-source mechanisms make the reconciliation honest: **renames** (read from the workbook's `aux_mapping`) fold the informe the same way b3 does (incorporação IRDM11→IRIM11 aglutina as linhas do informe), and the **fonte (PDF)** column traces every informe value to its source document. Not tax advice — every flag is for human review.

## Boas práticas / red flags — evitar erro de informe de banco
Nenhuma fonte sozinha é a verdade; cada uma é autoritativa para um campo. **Banco erra** — a defesa é cruzar sistematicamente (é o que este skill faz).

**Quem manda em quê:**
- **Custo / preço médio** → B3 reconstruído (todas as corretoras). Informe costuma *não abater amortização* (NUIF11: informe 979,20 × correto 961,20) e só ver a própria custódia (ALZR11: BTG 41,71 × ano todo 49,75).
- **Quantidade 31/12** → B3 Posição (oficial) + informe, cruzados.
- **Classificação** (isento×tributável, código, CNPJ, grupo) → informe do **administrador/escriturador** (NUIF11 = 07/10; BBSE3 7,41 = JCP 10, não 99). Para dividendos/JCP, o **escriturador** (Bradesco, Porto) é mais autoritativo que a corretora.
- **Renda tributada na fonte / saldos em banco / exterior** → só o informe vê (ex.: IRDM11 387,38 não consta na B3).

**Red flags (todos vistos em casos reais):**
1. Valor de **mercado** em vez de **custo** na Bens e Direitos.
2. **Amortização não abatida** do custo (informe infla o saldo).
3. Rendimento "isento" **com IR retido** (coluna IR ≠ "–") → na verdade é exclusiva/06.
4. **Período faltando** por transferência de custódia (some jan/fev de um lado).
5. **Mesmo valor em código diferente** entre fontes → reclassificação (BBSE3 99↔10).
6. **Quantidade muda sem evento** correspondente na Movimentação.
7. Valor que **só aparece no informe** e não na B3 → confirmar com o administrador (informe-only).

**Processo:**
1. B3 (Movimentação + Posição) é a espinha dorsal — consolida todas as corretoras; o evento subjacente costuma resolver a divergência.
2. Triangule ≥2 fontes; o que cair em um lado só fica sinalizado.
3. Reconcilie **por somatório**, não só linha a linha (aglutinação × reclassificação × valor faltando são coisas diferentes).
4. Amarre no **saldo do ano anterior** (abertura = fechamento anterior).
5. A **pré-preenchida** herda os mesmos erros do banco (vem da DIRF/e-Financeira) — cruze também.
6. Foco por **materialidade**; documente decisões (coluna `fonte` + `VERIFICAR`). Valores materiais em dúvida → administrador/escriturador ou contador.
