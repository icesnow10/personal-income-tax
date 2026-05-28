# personal-income-tax

Claude Code / agent **skills** for preparing the Brazilian personal income tax return (IRPF).

Install all skills into a project with [`skills`](https://github.com/mattpocock/skills):

```bash
npx skills@latest add icesnow10/personal-income-tax
```

## Skills

| Skill | Group | What it does |
|---|---|---|
| [rsu](skills/tax/rsu/SKILL.md) | tax | IRPF declaration for RSUs of a foreign (US-listed) company at any equity broker: vesting cost basis, capital gains on sales and the year-end position, converting USD→BRL with the Central Bank's official PTAX (sell on acquisition, buy on sale). Bundles a ready-to-fill spreadsheet template + helper scripts. |
| [b3](skills/tax/b3/SKILL.md) | tax | IRPF "Bens e Direitos" for B3 assets (ações, FIIs, BDRs, Tesouro/renda fixa): reconstructs the average acquisition cost (preço médio) from the **Movimentação** history and the year-end **Posição**. Ticker corrections + corporate actions via a documented overrides file (renames + cost resets with sources); emits the movements ledger, lookup tables, an avg-price summary, the merged position and ready-to-type IRPF rows (grupo / código / localização / discriminação). |

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
