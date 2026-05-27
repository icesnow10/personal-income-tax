# personal-income-tax

Claude Code / agent **skills** for preparing the Brazilian personal income tax return (IRPF).

Install all skills into a project with [`skills`](https://github.com/mattpocock/skills):

```bash
npx skills@latest add icesnow10/personal-income-tax
```

## Skills

### tax

- [rsu](skills/tax/rsu/SKILL.md) — Fill the IRPF declaration for Nu Holdings (Nubank) RSUs
  from E*TRADE / Morgan Stanley documents: vesting cost basis, capital gains on sales, and
  the year-end position, converting USD→BRL with BCB PTAX (venda for acquisition, compra for
  sale). Bundles a ready-to-fill spreadsheet template and helper scripts.

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
