# personal-income-tax

Claude Code / agent **skills** for preparing the Brazilian personal income tax return (IRPF).

Install all skills into a project with [`skills`](https://github.com/mattpocock/skills):

```bash
npx skills@latest add icesnow10/personal-income-tax
```

## Skills

### tax

- [rsu](skills/tax/rsu/SKILL.md) — Fill the IRPF declaration for RSUs of a foreign (US-listed)
  company held at any equity broker: vesting cost basis, capital gains on sales, and the
  year-end position, converting USD→BRL with the Central Bank's official PTAX (sell rate for
  acquisition, buy rate for sale). Modular (extraction / FX / calculation / template) and
  bundles a ready-to-fill spreadsheet template plus helper scripts.

---

> These skills are tooling aids, **not tax advice**. Always confirm the method with a
> qualified accountant.
