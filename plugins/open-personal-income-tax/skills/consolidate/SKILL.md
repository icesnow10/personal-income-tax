---
name: consolidate
description: Build the final IRPF workbook (irpf_consolidated.xlsx) from three sources — the variable_income workbook (preço médio of ações/FII/BDR), the fixed_income workbook (RF bens + isentos + exclusiva), and processed/informes.json (everything else: dividendos/JCP de ação, contas, RDB, exterior, JCP a receber). Writes Bens e Direitos, Rendimentos Isentos e Não Tributáveis, and Tributação Exclusiva/Definitiva. RF items already owned by fixed_income are NOT re-read from informes.json (no double counting). Use when the user wants to build/montar the final IRPF fichas, or mentions gerar as fichas, montar a declaração, consolidar, /consolidate.
---

# /generate — montar as fichas do IRPF a partir do B3 + transcrição dos informes

Segunda metade (determinística) do antigo `/consolidate`, agora separado em duas etapas:
1. **[/read](../read/SKILL.md)** — o agente lê **todo** o `resources/` (PDFs/prints dos informes) e
   produz `processed/informes.json` (transcrição unificada).
2. **/generate** (este) — script determinístico que junta `processed/b3_brazil_variable_income_avg_price_calculation.xlsx` +
   `processed/informes.json` no `irpf_consolidated.xlsx`.

O **mesmo** `informes.json` é consumido pelo [/completeness](../completeness/SKILL.md) — uma
transcrição só, não duas.

## A regra que decide se está certo

1. **Bens e Direitos, ativo B3** → o **valor** (`valor_2024`/`valor_2025`) vem da `b3_brazil_variable_income_avg_price_calculation.xlsx`
   (custo de aquisição = preço médio × quantidade), **NUNCA** o valor de mercado do informe. Mas
   **grupo, código e CNPJ vêm do informe**. Ex.: **NUIF11** é FI-Infra incentivado → **07/10**, não
   FII 07/03; só o informe diz isso.
2. **Bens e Direitos, não-B3** → tudo (inclusive o valor) vem do informe; não existe na B3.
3. **Rendimentos** (isentos / exclusiva) → **todos vêm dos informes**, nunca da `b3_brazil_variable_income_avg_price_calculation.xlsx`.
4. **Trust the informe for classification.** O script não inventa código nem CNPJ — só puxa o valor do
   B3 e casa com o `informes.json`. O que estiver no B3 e faltar no JSON é apontado no **AUDIT**.

## Entradas

Rodando da **pasta do contribuinte** (layout `resources/` cru → `processed/` derivado → root final):
- `processed/b3_brazil_variable_income_avg_price_calculation.xlsx` (saída do b3 — só a aba `IRPF_bens_e_direitos` é lida).
- `processed/informes.json` (saída do `/read` — schema unificado em [REFERENCE.md](REFERENCE.md)).

## Workflow

1. **Tenha o `processed/informes.json`** pronto (rode o [/read](../read/SKILL.md) antes). Para um
   esqueleto vazio: `python scripts/generate.py --template processed/informes.json`.
2. **Monte as fichas** (saída na raiz da pasta do contribuinte):
   `python scripts/generate.py --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx --json processed/informes.json --outdir .`
3. **Leia o AUDIT** (impresso no fim): todo ativo B3 que está na `b3_brazil_variable_income_avg_price_calculation.xlsx` mas não no
   `informes.json` é listado — adicione no JSON com o grupo/código/CNPJ do informe e rode de novo.
4. **Verifique com [/completeness](../completeness/SKILL.md)** (loop build → verify): reconcilia
   `b3_source × informes.json` por ficha (divergências, reclassificações, aglutinações) antes de digitar.

## Saída — 1 arquivo (`irpf_consolidated.xlsx`)

| Aba | Ficha do IRPF | Fonte do valor |
|---|---|---|
| `bens_e_direitos` | Bens e Direitos | B3: custo da `b3_brazil_variable_income_avg_price_calculation.xlsx`; não-B3: informe |
| `isentos` | Rendimentos Isentos e Não Tributáveis | informes |
| `exclusiva_definitiva` | Tributação Exclusiva/Definitiva | informes |

## Importante

- **Não é orientação fiscal.** Um valor divergente é sinal para revisão, não correção automática.
- Mantenha o skill **genérico** — nunca comite tickers, CNPJs ou valores de um contribuinte real;
  o `informes.json` e os arquivos gerados ficam na pasta de trabalho.
- Os **códigos** (grupo/código IRPF, CNPJ) são autoridade do **informe**, não do b3. Ver
  [REFERENCE.md](REFERENCE.md) para o schema, a tabela de códigos de referência e o algoritmo.
