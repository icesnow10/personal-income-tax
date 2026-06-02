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
  **Movimentação** (the full history since your first trade — what `renda_variavel` replays to rebuild the
  preço médio), the **Posição** at 31/12 of the base year, and ideally the **Posição** at 31/12 of the
  previous year. No B3 holdings? Skip the `renda_variavel`/`renda_fixa` steps entirely.

No need to be exhaustive on the first pass: **`completeness` flags any asset still missing a supporting
informe** (and, for B3 assets, pulls the escriturador from the B3 API to check the authoritative
dividend/JCP statement was used). The report tells you what's still missing — add it and re-run.

## Pipeline (investments)

`irpf` orchestrates the whole investments declaration end to end, in five steps:

```
resources/  (broker/bank informes + B3 "Movimentação"/"Posição" exports)
     │
 [1] read ───────────►  processed/informes.json                       one unified transcription of everything in resources/
     │
 [2] renda_variavel ─►  processed/b3_brazil_renda_variavel_avg_price…  preço médio of ações/FII/BDR (renda variável)
     │
 [3] renda_fixa ─────►  processed/renda_fixa.xlsx                      CDB/CRA/CRI/deb/Tesouro: bens + isentos + exclusiva
     │                                                                  (value from the informe + a B3 position validation)
 [4] consolidate ────►  irpf_consolidated.xlsx                         merges RV + RF + the rest of the informe → 3 fichas
     │
 [5] completeness ───►  completeness_report.md                         audits + edits the consolidated to the per-ficha authority
```

Sources of truth, per ficha: **renda variável value** → `renda_variavel` (preço médio); **renda fixa
value and all rendimentos** → the informes. `completeness` enforces this on the deliverable and pulls
each ação/FII escriturador straight from the B3 public API (no prints, no informe dependency).

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
│  ├─ ticker_memory.md             renames / incorporações (renda_variavel)
│  ├─ mapping_memory.md            B3 movement → action (renda_variavel)
│  └─ escriturador_memory.md       ticker → escriturador (auto-generated from the B3 API)
├─ processed/                      derived (generated)
│  ├─ informes.json                unified transcription (read)
│  ├─ b3_brazil_renda_variavel_avg_price_calculation.xlsx   preço médio of ações/FII/BDR (renda_variavel)
│  └─ renda_fixa.xlsx              RF bens + isentos + exclusiva + position validation (renda_fixa)
├─ irpf_consolidated.xlsx          ◄ deliverable: the 3 fichas, ready to type
└─ completeness_report.md          ◄ deliverable: the audit
```

Keep your real `resources/`, `memory/` and generated files in your own working folder — **never commit
a taxpayer's data** (see `.gitignore`).

## Skills

The investments pipeline (`irpf` and everything under it) is bundled as the
**`open-personal-income-tax`** plugin. `rsu` is standalone.

| Skill | Children | What it does |
|---|---|---|
| [irpf](skills/tax/irpf/SKILL.md) | `read` · `renda_variavel` · `renda_fixa` · `consolidate` · `completeness` | Orchestrator — runs the full investments pipeline end to end (5 steps) from the broker/bank informes + B3 exports. |
| [read](skills/tax/read/SKILL.md) | — | Reads every broker/bank informe in `resources/` (PDFs/prints, many image/encrypted) into the unified `processed/informes.json` — one object per asset/rendimento, tagged by ficha (bens / isentos / exclusiva), with key, grupo/código, CNPJ, value and the source PDF. |
| [renda_variavel](skills/tax/renda_variavel/SKILL.md) | — | Reconstructs the average acquisition cost (**preço médio**) of B3 renda variável (ações, FIIs, BDRs) from the **Movimentação** history + year-end **Posição**. Driven by editable **living-memory files** (movement→action mapping, ticker renames). Renda variável only. |
| [renda_fixa](skills/tax/renda_fixa/SKILL.md) | — | Builds the **renda fixa** slice (CDB/CRA/CRI/debênture/LCI/LCA/Tesouro) from `informes.json` — bens + isentos (cód 12) + exclusiva (cód 06), value from the broker informe — plus a **B3 position validation** (every RF security held is covered by an informe). |
| [consolidate](skills/tax/consolidate/SKILL.md) | — | Merges the three sources — renda_variavel workbook + renda_fixa workbook + `informes.json` (everything else) — into `irpf_consolidated.xlsx`, the 3 IRPF fichas. RF owned by renda_fixa isn't re-read (no double counting). |
| [completeness](skills/tax/completeness/SKILL.md) | — | Audits the consolidated against the informes **by ficha** and **edits `irpf_consolidated.xlsx` in place** to the per-ficha authority, stamping an `obs_completeness` column. Pulls each ação/FII **escriturador from the B3 API** to check the authoritative dividend/JCP informe was used. |
| [rsu](skills/tax/rsu/SKILL.md) | standalone · 🧪 beta | IRPF declaration for RSUs of a foreign (US-listed) company at any equity broker: vesting cost basis, capital gains on sales and the year-end position, converting USD→BRL with the Central Bank's official PTAX. Bundles a ready-to-fill spreadsheet template + helper scripts. Validate carefully before relying on it. |

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
