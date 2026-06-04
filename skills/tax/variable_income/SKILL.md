---
name: variable_income
description: Orchestrate the renda-variável slice of the IRPF. Only the Brazilian side runs a script (vi_br — preço médio of B3 ações/FII/BDR from the Movimentação → b3 workbook). The foreign side (vi_international — ações/ETF no exterior) has NO script: it flows /read → informes.json → /consolidate. Use when the user wants the variable-income part of the IRPF (Brazilian + international), or mentions renda variável, ações/FII/BDR + ações/ETF no exterior.
---

# /variable_income — renda variável (Brasil + internacional)

```
[a] vi_br (script) ──────────────►  b3_brazil_variable_income_avg_price_calculation.xlsx
        (preço médio de ações/FII/BDR da B3, reconstruído da Movimentação)
[b] vi_international (sem script) ► itens no informes.json (b3:false, loc≠105) ─► /consolidate
        (ações/ETF no EXTERIOR — valor do informe; NÃO estão na B3, sem arquivo próprio)
```

- **[vi_br](../vi_br/SKILL.md)** — único com script: reconstrói o **preço médio** da renda variável
  custodiada na B3 (ações, FIIs, BDRs) a partir do `Movimentação` + `Posição`.
- **[vi_international](../vi_international/SKILL.md)** — **documentação, sem arquivo**: ações/ETFs no
  exterior (Avenue, Nomad, Interactive Brokers) **não estão na B3** → vêm do `informes.json` (transcritos
  pelo [/read](../read/SKILL.md)) e entram direto no [/consolidate](../consolidate/SKILL.md).

Quem é dono de quê: **B3** (ações/FII/BDR) → vi_br (workbook); **exterior** (localização ≠ 105),
**renda fixa** e o resto (contas, RDB, JCP a receber) → todos via `informes.json` → consolidate.

## Workflow
1. `python ../vi_br/scripts/build_bens_direitos.py resources/MOV.xlsx resources/POS.xlsx processed/b3_brazil_variable_income_avg_price_calculation.xlsx --memory-dir memory --year YYYY [--posicao-anterior resources/POS_PRIOR.xlsx]` — leia o AUDIT.
2. Exterior: **nada a rodar** — garanta que o [/read](../read/SKILL.md) transcreveu as ações/ETF no
   exterior no `informes.json` (`b3:false`, `localizacao` ≠ 105). O /consolidate as declara.

## Importante
- **Não é orientação fiscal.** O custo de RV-Brasil vem do preço médio (B3); o de RV-exterior vem do
  informe da corretora estrangeira (não há B3 para validar).
- Mantenha **genérico** — nunca comite dados de um contribuinte.
