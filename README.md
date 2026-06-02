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

## Pipeline (investments)

`irpf` orchestrates the whole investments declaration end to end; `consolidate` is itself an
orchestrator of `read` + `generate`:

```
resources/  (B3 "Movimentação"/"Posição" exports + broker/bank informes)
     │
 [1] b3 ───────────────►  processed/brazil_investments.xlsx   preço médio / custo — source of truth for VALUE
     │
 [2] consolidate ──────►  irpf_consolidated.xlsx              the 3 fichas, ready to type
     │   ├─ read      →  processed/informes.json              one unified transcription of everything in resources/
     │   └─ generate  →  irpf_consolidated.xlsx
     │
 [3] completeness ─────►  completeness_report.md              audits b3_source × informes.json AND fixes the consolidated
```

Sources of truth, per ficha: **Bens e Direitos value** → `b3` (custo); **grupo/código/CNPJ and all
rendimentos** → the informes. `completeness` enforces this on the deliverable and pulls each ação/FII
escriturador straight from the B3 public API (no prints, no informe dependency).

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
│  ├─ ticker_memory.md             renames / incorporações (b3)
│  ├─ mapping_memory.md            B3 movement → action (b3)
│  ├─ rf_memory.md / rf_value_memory.md   renda fixa (b3)
│  └─ escriturador_memory.md       ticker → escriturador (auto-generated from the B3 API)
├─ processed/                      derived (generated)
│  ├─ brazil_investments.xlsx      b3 output — preço médio / custo
│  └─ informes.json                unified transcription (read)
├─ irpf_consolidated.xlsx          ◄ deliverable: the 3 fichas, ready to type
└─ completeness_report.md          ◄ deliverable: the audit
```

Keep your real `resources/`, `memory/` and generated files in your own working folder — **never commit
a taxpayer's data** (see `.gitignore`).

## Skills

The investments pipeline (`irpf` and everything under it) is bundled as the
**`open-personal-income-tax`** plugin. `rsu` is standalone.

| Skill | What it does |
|---|---|
| **[irpf](skills/tax/irpf/SKILL.md)** · orchestrator | Runs the full investments pipeline `b3` → `consolidate` → `completeness` from the B3 exports + broker/bank informes. |
| ├─ [b3](skills/tax/b3/SKILL.md) | IRPF "Bens e Direitos" for B3 assets (ações, FIIs, BDRs, Tesouro/renda fixa): reconstructs the average acquisition cost (preço médio) from the **Movimentação** history and the year-end **Posição**. Driven by editable **living-memory files** (movement→action mapping, ticker renames, corporate-action cost resets — each with sources). |
| ├─ [consolidate](skills/tax/consolidate/SKILL.md) · orchestrator | Orchestrates `read` + `generate`: transcribe every informe, then build the three IRPF fichas. |
| │&nbsp;&nbsp;&nbsp;├─ [read](skills/tax/read/SKILL.md) | Reads every broker/bank informe in `resources/` (PDFs/prints, many image/encrypted) into the unified `processed/informes.json` — one object per asset/rendimento, tagged by ficha (bens / isentos / exclusiva), with key, grupo/código, CNPJ, value and the source PDF. |
| │&nbsp;&nbsp;&nbsp;└─ [generate](skills/tax/generate/SKILL.md) | Deterministically merges the b3 workbook + `informes.json` into `irpf_consolidated.xlsx`. B3-asset value comes from the b3 custo; grupo/código/CNPJ and all rendimentos come from the informes. |
| └─ [completeness](skills/tax/completeness/SKILL.md) | Audits the consolidated against the informes **by ficha** and **edits `irpf_consolidated.xlsx` in place** to the per-ficha authority, stamping an `obs_completeness` column. Also pulls each ação/FII **escriturador from the B3 API** to check the authoritative dividend/JCP informe was used. |
| **[rsu](skills/tax/rsu/SKILL.md)** · standalone · 🧪 beta | IRPF declaration for RSUs of a foreign (US-listed) company at any equity broker: vesting cost basis, capital gains on sales and the year-end position, converting USD→BRL with the Central Bank's official PTAX (sell on acquisition, buy on sale). Bundles a ready-to-fill spreadsheet template + helper scripts. **Beta** — validate carefully before relying on it. |

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
