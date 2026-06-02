---
name: completeness
description: Audit the generated b3_brazil_renda_variavel_avg_price_calculation.xlsx (b3 skill output) against the taxpayer's broker/bank statements (informe_de_rendimentos), ORGANIZED BY IRPF FICHA (Bens e Direitos, Rendimentos Isentos, Tributação Exclusiva), and write a single .md report with grupo/código, quantities, values, the source PDF per value, divergences, reclassifications and per-código sum reconciliation. b3_source is the source of truth for value/custo; the informe is the authority for classification. Use when the user wants a completeness/verification check of the IRPF — e.g. "compare os informes com o brazil_investments", "falta algo?", "/completeness".
---

# /completeness — auditoria por ficha do IRPF (b3_source × informe_de_rendimentos)

Cross-check the two views and write a single **`.md`** report, organized by the **IRPF fichas**:

1. **b3_source** — the `irpf_bens_e_direitos_renda_variavel` sheet of
   `b3_brazil_renda_variavel_avg_price_calculation.xlsx` (the [b3](../b3/SKILL.md) output): **renda
   variável only** (ações/FII/BDR — preço médio). The b3 workbook no longer carries rendimentos sheets
   (authority for rendimentos is the informe) nor renda-fixa value (declared from the informe).
2. **informe_de_rendimentos** — `processed/informes.json`, the unified transcription that
   [/read](../read/SKILL.md) builds from every PDF in `resources/`. The **same** file
   [/generate](../generate/SKILL.md) consumes — one transcription, not two.

**Who is the source of truth — por ficha** (see also [REFERENCE.md](REFERENCE.md) "Boas práticas / red flags"):
- **Bens e Direitos — valor / custo / preço médio → b3_source** (consolidated, amortização abatida). The
  informe never overrides the B3 value (informe may carry market value, miss amortização, or show one
  custody only). O **grupo/código/CNPJ** desses ativos vem do **informe**.
- **Rendimentos Isentos (09/99) e Tributação Exclusiva (06/10) — valor E classificação → informe.**
  É regime de caixa, somando corretoras + escriturador; a aba de income do b3_source é **parcial** e
  **não é declarada** — numa divergência de rendimento, **vale o informe**, nunca o b3_source.
- **Itens não-B3** (banco, RDB/CDB, moeda, exterior, JCP a receber) → só existem no informe.

The b3 skill only covers **B3-custodied** assets; banco/exterior/saldos live only in the informes and
must be declared by hand — this skill flags those so nothing is missed.

## Inputs the user provides
Run from the **taxpayer folder** (layout `resources/` raw → `processed/` derived → root deliverables):
- **`processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx`** (the b3 skill output).
- **`processed/informes.json`** — the unified transcription built by [/read](../read/SKILL.md). A single
  object with `bens` / `isentos` / `exclusiva` lists (schema in [../generate/REFERENCE.md](../generate/REFERENCE.md)):
  ```json
  {
    "bens":      [{"key":"NUIF11","b3":true,"grupo":7,"codigo":10,"cnpj":"...","quantidade":9,"valor_2025":979.20,"source":"btg_FD0257"},
                  {"key":"VALE3","b3":true,"grupo":3,"codigo":1,"quantidade":20,"source":"btg_787"}],
    "isentos":   [{"codigo":9,"key":"BBSE3","valor_2025":842.87,"source":"btg_787 + nubank_informe"}],
    "exclusiva": [{"codigo":10,"key":"VALE3","valor_2025":41.07,"source":"btg_787 + nubank_informe"}]
  }
  ```
  For ações/FII the informe gives only **quantidade** (no BRL value) → put `quantidade`, omit `valor_2025`.
  `source` is the **PDF** the value came from (joined with `+`/`/` across PDFs) → the **fonte (PDF)** column.
- **`memory/escriturador_memory.md`** — living memory `ticker → escriturador` (nome, CNPJ, `tag`, source),
  que liga o papel ao escriturador para a auditoria "o informe do escriturador foi usado?". **Gere
  automaticamente da B3** (sem prints, sem depender dos informes) com
  `python scripts/fetch_escriturador.py --investimentos processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx` — ele lê os
  tickers/CNPJs do workbook do b3 (sempre presente) e busca o escriturador na API pública da B3 (campo
  `ifd` p/ FII/FI-Infra, `hasCommom` p/ ação/BDR — o mesmo de B3 > "Informações Gerais do Fundo >
  Escriturador"), preservando o `tag` já preenchido. Template/fallback manual em
  [escriturador_memory.md](escriturador_memory.md). Override do caminho com `--escriturador-memory`.

## Workflow
1. **Build the transcription** with [/read](../read/SKILL.md) (or run it standalone). Parsing aid:
   `python scripts/completeness.py extract resources/` (image/encrypted PDFs yield nothing — read by hand).
2. **Escriturador da B3** (sem prints/informes): `python scripts/fetch_escriturador.py --investimentos processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx` → popula `memory/escriturador_memory.md` consultando a API da B3 para cada ação/FII/BDR do workbook.
3. **Reconcile + fix**: `python scripts/completeness.py compare --investimentos processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx --informes processed/informes.json`
   Writes `completeness_report.md` AND **audits/edits the deliverable `irpf_consolidated.xlsx` in place**
   (it's picked up automatically beside the report; override with `--consolidado`). Any value that drifts
   from its ficha authority (Bens → b3_source custo; rendimentos → informe) is **rewritten in the
   consolidated** and stamped in an `obs_completeness` column; rows that merely diverged from the other
   source (but were already correct) get a "conferido" note. Use `--no-apply` to audit without editing.
4. **Read the `.md`** — sections per IRPF ficha:
   - **Bens e Direitos**: `grupo | código | asset | qtd b3 | qtd informe | valor b3 | valor informe | fonte (PDF) | status`. Código mismatch is flagged inline (`⚠️b3=3/inf=10`); a TOTAL-a-declarar line counts the não-B3 items to add by hand.
   - **Rendimentos Isentos** (09/99) and **Tributação Exclusiva** (06/10): per código, with a **SOMA por código** (a per-item gap that reconciles at the total is visible).
   - **Escriturador das ações/FII**: por ação/FII/FI-Infra, qual o escriturador (de `memory/escriturador_memory.md`) e se **o informe do escriturador foi usado** (o `tag` aparece no `source` do rendimento). O escriturador é a autoridade de dividendos/JCP — a corretora pode ver só parte (ex.: BBSE3 R$ 454,46 na corretora × R$ 842,87 no escriturador BB). Sinaliza escriturador desconhecido (preencher da B3) ou rendimento que veio só de corretora.
   - **Divergências de classificação**: same asset+value under a different código between b3 and informe (reclassificação — ex.: BBSE3 cód 99 no b3 × cód 10 no informe).
   - **Ajustes no `irpf_consolidated.xlsx`**: o que foi editado (valor corrigido p/ a autoridade) e o que foi conferido — espelha a coluna `obs_completeness` do arquivo.
   - **Resumo** + **Action items**.

The reconciliation applies two cross-source mechanisms automatically:
- **Renames / aglutinação** — read from the workbook's `aux_mapping`; folds the informe the same way b3
  does (incorporação IRDM11→IRIM11 aglutina as linhas do informe na cota nova, com breakdown na nota).
- **Fonte (PDF)** — every informe value traces to its source document for tracking.

## Status meanings
| status | significado | ação |
|---|---|---|
| `OK` | presente e igual nos dois lados | nada |
| `OK (informe só qtd …)` | ação/FII: qtd bate; valor é o custo do b3_source | nada |
| `DIVERGE (… — vale b3_source)` | em **Bens e Direitos**: valores diferem | revisar; **vale o b3_source** (ex.: custódia parcial) |
| `DIVERGE (… — vale o informe)` | em **rendimentos**: valores diferem | revisar; **vale o informe** (regime de caixa, corretoras + escriturador) |
| `⚠️b3=X/inf=Y` (na coluna código) | b3 e informe classificaram em código diferente | usar o código do **informe** (ex.: NUIF11 07/10) |
| `só no informe (não-B3 — declarar à mão)` | está no informe, não é B3 | declarar à mão a partir do informe |
| `FALTA em b3_source (B3!)` | ticker B3 que o b3 não pegou | **bug** — investigar o `/b3` |
| `FALTA no informe_de_rendimentos` | está no b3, sem informe que suporte | confirmar / achar o documento |
| reclassificação (seção própria) | mesmo ativo+valor em código diferente | declarar no código do **informe** |
| `✅ sim (tag)` (escriturador) | o informe do escriturador foi usado no rendimento | nada |
| `⚠️ escriturador desconhecido` | ticker fora do `escriturador_memory.md` | preencher da B3 (ações/FIIs/FI-Infra/BDRs) e re-rodar |
| `⚠️ NÃO` (escriturador) | rendimento veio só de corretora | buscar o informe do escriturador (pode faltar dividendo/JCP) |

## Important
- **Not tax advice.** A divergence is a flag for review, not an automatic correction.
- Keep this skill **generic** — never commit a taxpayer's holdings, CNPJs or values. `informes.json`
  lives in the working folder.
- See [REFERENCE.md](REFERENCE.md) for the schema, matching, the aglutinação/renames mechanism and the
  **Boas práticas / red flags** (how to avoid trusting a wrong bank informe).
