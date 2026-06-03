---
name: variable_income
description: Orchestrate the renda-variável slice of the IRPF — runs variable_income_brazil (preço médio of B3 ações/FII/BDR from the Movimentação) then variable_income_international (foreign stocks/ETFs from the informes, not on the B3). Use when the user wants to build the whole variable-income part of the IRPF (Brazilian + international), or mentions renda variável, ações/FII/BDR + ações/ETF no exterior.
---

# /variable_income — renda variável (Brasil + internacional)

Orquestra os dois lados da renda variável; cada um é um skill próprio:

```
[a] variable_income_brazil ──────►  b3_brazil_variable_income_avg_price_calculation.xlsx
        (preço médio de ações/FII/BDR da B3, reconstruído da Movimentação)
[b] variable_income_international ►  variable_income_international.xlsx
        (ações/ETF no EXTERIOR — só dos informes, NÃO estão na B3)
```

- **[variable_income_brazil](../variable_income_brazil/SKILL.md)** — reconstrói o **preço médio** da
  renda variável custodiada na B3 (ações, FIIs, BDRs) a partir do `Movimentação` + `Posição`.
- **[variable_income_international](../variable_income_international/SKILL.md)** — monta a renda variável
  **no exterior** (ações/ETFs em corretoras como Avenue, Nomad, Interactive Brokers) a partir do
  `informes.json` — **não há Posição B3** para esses; o valor vem do informe.

Os dois alimentam o [/consolidate](../consolidate/SKILL.md), que junta tudo na declaração final.
Quem é dono de quê: **B3** (ações/FII/BDR) → brazil; **exterior** (localização ≠ 105) → international;
renda fixa → [/fixed_income](../fixed_income/SKILL.md); o resto (contas, RDB, JCP a receber) → consolidate.

## Workflow
1. `python ../variable_income_brazil/scripts/build_bens_direitos.py resources/MOV.xlsx resources/POS.xlsx processed/b3_brazil_variable_income_avg_price_calculation.xlsx --memory-dir memory --year YYYY [--posicao-anterior resources/POS_PRIOR.xlsx]`
2. `python ../variable_income_international/scripts/build_international.py --informes processed/informes.json --out processed/variable_income_international.xlsx`

Cada um imprime seu AUDIT/validação — leia antes de consolidar.

## Importante
- **Não é orientação fiscal.** O custo de RV-Brasil vem do preço médio (B3); o de RV-internacional vem
  do informe da corretora estrangeira (não há B3 para validar).
- Mantenha **genérico** — nunca comite dados de um contribuinte.
