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

## Skills

The `b3 → consolidate → completeness` pipeline (orchestrated by `irpf`) is bundled as the
**`open-personal-income-tax`** plugin; `rsu` is standalone.

| Skill | Group | What it does |
|---|---|---|
| [irpf](skills/tax/irpf/SKILL.md) | tax | **Orchestrator** of the full investments pipeline end to end: `b3` → `consolidate` → `completeness`. Run it to build the whole investments declaration from the B3 exports + the broker/bank informes. |
| [b3](skills/tax/b3/SKILL.md) | tax | IRPF "Bens e Direitos" for B3 assets (ações, FIIs, BDRs, Tesouro/renda fixa): reconstructs the average acquisition cost (preço médio) from the **Movimentação** history and the year-end **Posição**. Driven by editable **living-memory files** (movement→action mapping, ticker renames, corporate-action cost resets — each with sources). Emits the movements ledger, lookup tables, an avg-price summary, the merged position and ready-to-type IRPF rows. |
| [read](skills/tax/read/SKILL.md) | tax | Reads every broker/bank informe in `resources/` (PDFs/prints, many image/encrypted) and transcribes them into the unified `processed/informes.json` consumed by `generate` and `completeness` — one object per asset/rendimento, tagged by ficha (bens / isentos / exclusiva), with key, grupo/código, CNPJ, value and the source PDF. |
| [generate](skills/tax/generate/SKILL.md) | tax | Deterministically merges the b3 workbook + `informes.json` into the final `irpf_consolidated.xlsx` (Bens e Direitos, Rendimentos Isentos, Tributação Exclusiva). Value of B3 assets comes from the b3 custo; grupo/código/CNPJ and all rendimentos come from the informes. |
| [consolidate](skills/tax/consolidate/SKILL.md) | tax | Orchestrates `read` + `generate`: transcribe every informe into `informes.json`, then build the three IRPF fichas. |
| [completeness](skills/tax/completeness/SKILL.md) | tax | Audits the consolidated against the informes **by ficha** and **edits `irpf_consolidated.xlsx` in place** to the per-ficha authority (Bens e Direitos → b3 custo; rendimentos → informe), stamping an `obs_completeness` column and reporting every adjustment in a `.md`. |
| [rsu](skills/tax/rsu/SKILL.md) | tax | IRPF declaration for RSUs of a foreign (US-listed) company at any equity broker: vesting cost basis, capital gains on sales and the year-end position, converting USD→BRL with the Central Bank's official PTAX (sell on acquisition, buy on sale). Bundles a ready-to-fill spreadsheet template + helper scripts. |

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
