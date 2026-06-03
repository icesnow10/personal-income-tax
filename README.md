# personal-income-tax

Claude Code / agent **skills** for preparing the Brazilian personal income tax return (IRPF).

## Install

As a **Claude Code plugin** (marketplace in this repo):

```
/plugin marketplace add icesnow10/personal-income-tax
/plugin install open-personal-income-tax@personal-income-tax
```

Or drop all skills into a project with [`skills`](https://github.com/mattpocock/skills):

```bash
npx skills@latest add icesnow10/personal-income-tax
```

## Getting the inputs (onboarding)

Gather whatever applies to you into the taxpayer's `resources/` folder — only what you actually have:

- **Informes de Rendimentos** — the IRPF informe of the base year from each institution where you hold
  (or held) assets or cash: banks, brokers, fintechs (Nubank, Itaú, Bradesco, BB, Wise, Nomad, BTG, XP, …).
  This is the baseline most people need.
- **B3 exports** *(only if you have/had assets custodiados na B3 — ações, FIIs, BDRs, Tesouro, renda fixa)* —
  from the *área do investidor* ([investidor.b3.com.br](https://www.investidor.b3.com.br/) → **Extratos**):
  **Movimentação** (the full history since your first trade — what `vi_br` replays to
  rebuild the preço médio), the **Posição** at 31/12 of the base year, and ideally the **Posição** at
  31/12 of the previous year. No B3 holdings? Skip the `variable_income`/`fixed_income` steps entirely.
  Foreign stocks/ETFs (Avenue, Nomad, Interactive Brokers) aren't on the B3 — `vi_international`
  builds those straight from the broker informe.

No need to be exhaustive on the first pass: **`completeness` flags any asset still missing a supporting
informe** (and, for B3 assets, pulls the escriturador from the B3 API to check the authoritative
dividend/JCP statement was used). The report tells you what's still missing — add it and re-run.

## Pipeline (investments)

`irpf` orchestrates the whole investments declaration end to end, in five steps:

```
resources/  (broker/bank informes + B3 "Movimentação"/"Posição" exports)
     │
 [1] read ─────────────────────────►  processed/informes.json          one unified transcription of everything in resources/
     │
 [2] variable_income  (renda variável — Brasil + exterior):
     ├─ vi_br ──────────────────────►  …variable_income_avg_price….xlsx  preço médio of ações/FII/BDR on the B3
     └─ vi_international ────────────►  variable_income_international.xlsx ações/ETF abroad (from the informe — not on the B3)
     │
 [3] fixed_income ─────────────────►  processed/fixed_income.xlsx       CDB/CRA/CRI/deb/Tesouro: bens + isentos + exclusiva
     │                                                                  (value from the informe + a B3 position validation)
 [4] consolidate ──────────────────►  irpf_consolidated.xlsx           merges RV-Brasil + RV-exterior + RF + the rest → 3 fichas
     │
 [5] completeness ─────────────────►  completeness_report.md            audits + edits the consolidated to the per-ficha authority
```

Sources of truth, per ficha: **Brazilian renda variável value** → `vi_br` (preço médio);
**foreign equity value** → `vi_international` (the broker informe — no B3 to validate against);
**renda fixa value and all rendimentos** → the informes. `completeness` enforces this on the deliverable
and pulls each ação/FII escriturador straight from the B3 public API (no prints, no informe dependency).

## Folder layout (per taxpayer)

Run the pipeline from a taxpayer folder in three layers — `resources/` (raw) → `processed/` (derived)
→ root (deliverables):

```
<taxpayer>/
├─ resources/                      raw inputs YOU provide
│  ├─ MOV.xlsx                     B3 "Movimentação" (all years since the first buy)
│  ├─ POS.xlsx                     B3 "Posição" at 31/12 of the base year (POS_PRIOR.xlsx optional)
│  └─ *.pdf                        broker/bank informes (BTG, NU, Itaú, BB, Wise, Nomad, …)
├─ memory/                         editable living-memory
│  ├─ ticker_memory.md             renames / incorporações (vi_br)
│  ├─ mapping_memory.md            B3 movement → action (vi_br)
│  └─ escriturador_memory.md       ticker → escriturador (auto-generated from the B3 API)
├─ processed/                      derived (generated)
│  ├─ informes.json                unified transcription (read)
│  ├─ b3_brazil_variable_income_avg_price_calculation.xlsx  preço médio of ações/FII/BDR (vi_br)
│  ├─ variable_income_international.xlsx  ações/ETF abroad, value from the informe (vi_international)
│  └─ fixed_income.xlsx            RF bens + isentos + exclusiva + position validation (fixed_income)
├─ irpf_consolidated.xlsx          ◄ deliverable: the 3 fichas, ready to type
└─ completeness_report.md          ◄ deliverable: the audit
```

Keep your real `resources/`, `memory/` and generated files in your own working folder — **never commit
a taxpayer's data** (see `.gitignore`).

## Skills

The investments pipeline (`irpf` and everything under it) is bundled as the
**`open-personal-income-tax`** plugin. `rsu` is standalone.

Each row fills the column matching its depth in the tree: top-level orchestrators in **Skill**, their
children in the first **Children** column, grandchildren in the second.

| Skill | Children | Children | What it does |
|---|---|---|---|
| [irpf](skills/tax/irpf/SKILL.md) | | | Orchestrator — runs the full investments pipeline end to end (5 steps) from the broker/bank informes + B3 exports. |
| | [read](skills/tax/read/SKILL.md) | | Reads every broker/bank informe in `resources/` (PDFs/prints, many image/encrypted) into the unified `processed/informes.json` — one object per asset/rendimento, tagged by ficha (bens / isentos / exclusiva), with key, grupo/código, CNPJ, value and the source PDF. |
| | [variable_income](skills/tax/variable_income/SKILL.md) | | Orchestrates the whole renda-variável slice — Brazilian (B3) and foreign — and feeds both into `consolidate`. |
| | | [vi_br](skills/tax/vi_br/SKILL.md) | Reconstructs the average acquisition cost (**preço médio**) of B3 renda variável (ações, FIIs, BDRs) from the **Movimentação** history + year-end **Posição**. Driven by editable **living-memory files** (movement→action mapping, ticker renames). On-B3 renda variável only. |
| | | [vi_international](skills/tax/vi_international/SKILL.md) | Builds **foreign** stocks/ETFs (Avenue, Nomad, Interactive Brokers, …) from `informes.json` — these aren't on the B3, so there's no preço médio: the value comes straight from the broker informe. Flags foreign income for the correct ficha (carnê-leão/GCAP). |
| | [fixed_income](skills/tax/fixed_income/SKILL.md) | | Builds the **renda fixa** slice (CDB/CRA/CRI/debênture/LCI/LCA/Tesouro) from `informes.json` — bens + isentos (cód 12) + exclusiva (cód 06), value from the broker informe — plus a **B3 position validation** (every RF security held is covered by an informe). |
| | [consolidate](skills/tax/consolidate/SKILL.md) | | Merges the four sources — vi_br workbook + vi_international workbook + fixed_income workbook + `informes.json` (everything else) — into `irpf_consolidated.xlsx`, the 3 IRPF fichas. Slices owned by the workbooks aren't re-read (no double counting). |
| | [completeness](skills/tax/completeness/SKILL.md) | | Audits the consolidated against the informes **by ficha** and **edits `irpf_consolidated.xlsx` in place** to the per-ficha authority, stamping an `obs_completeness` column. Pulls each ação/FII **escriturador from the B3 API** to check the authoritative dividend/JCP informe was used. |
| [rsu](skills/tax/rsu/SKILL.md) 🧪 beta | | | IRPF declaration for RSUs of a foreign (US-listed) company at any equity broker: vesting cost basis, capital gains on sales and the year-end position, converting USD→BRL with the Central Bank's official PTAX. Bundles a ready-to-fill spreadsheet template + helper scripts. Validate carefully before relying on it. |

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
