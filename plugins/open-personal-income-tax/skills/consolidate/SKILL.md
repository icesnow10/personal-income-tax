---
name: consolidate
description: Consolidate the broker/bank informes with the B3 output into the final IRPF fichas — an orchestrator that runs /read (transcribe every informe in resources/ into the unified processed/informes.json) then /generate (build irpf_consolidated.xlsx from b3_brazil_renda_variavel_avg_price_calculation.xlsx + informes.json). Produces Bens e Direitos, Rendimentos Isentos e Não Tributáveis, and Tributação Exclusiva/Definitiva. Use when the user wants to consolidate/merge the informes with the B3 output into the final IRPF fichas, or mentions consolidar informes, juntar tudo na declaração, montar as fichas do IRPF, /consolidate.
---

# /consolidate — read + generate (montar as fichas do IRPF)

Orquestra as duas etapas que transformam os informes + a saída do b3 nas fichas do IRPF:

```
resources/ (informes) ──[/read]──► processed/informes.json ──[/generate]──► irpf_consolidated.xlsx
                                                  ▲
                       processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx (b3) ─────────────┘
```

1. **[/read](../read/SKILL.md)** — o agente lê **todo** o `resources/` (PDFs/prints dos informes) e
   transcreve para `processed/informes.json` (schema unificado; layouts variam demais e muitos PDFs
   são imagem/cifrados, então a leitura é do agente).
2. **[/generate](../generate/SKILL.md)** — script determinístico que junta
   `processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx` + `processed/informes.json` no `irpf_consolidated.xlsx`.

O **mesmo** `processed/informes.json` é depois lido pelo [/completeness](../completeness/SKILL.md) na
verificação — uma transcrição só.

## Entradas
Rodando da **pasta do contribuinte** (layout `resources/` cru → `processed/` derivado → root final):
- **PDFs dos informes** em `resources/`.
- `processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx` (saída do [b3](../b3/SKILL.md)).

## Workflow

1. **/read** — transcreva os informes (ver [read](../read/SKILL.md) para schema e regras). Auxílio de
   parsing: `python ../completeness/scripts/completeness.py extract resources/` (cifrados leia com Read).
   Esqueleto vazio: `python ../generate/scripts/generate.py --template processed/informes.json`.
2. **/generate** — monte as fichas (saída na raiz):
   `python ../generate/scripts/generate.py --investimentos processed/b3_brazil_renda_variavel_avg_price_calculation.xlsx --json processed/informes.json --outdir .`
3. **Leia o AUDIT** do generate: ativo B3 na `b3_brazil_renda_variavel_avg_price_calculation.xlsx` ausente do `informes.json` é
   listado — adicione no JSON (volta ao /read) e re-rode o /generate.
4. **Verifique com [/completeness](../completeness/SKILL.md)** (loop build → verify) antes de digitar.

## Saída — `irpf_consolidated.xlsx`
Um arquivo com as abas `bens_e_direitos`, `isentos` e `exclusiva_definitiva` (detalhe em
[/generate](../generate/SKILL.md)).

## Importante
- **Não é orientação fiscal.** Divergência é sinal de revisão, não correção automática.
- **Fontes da verdade:** valor/custo de ativo B3 → `b3_brazil_renda_variavel_avg_price_calculation.xlsx`; classificação
  (grupo/código/CNPJ) e rendimentos → o informe.
- Mantenha **genérico** — `informes.json` e os arquivos gerados ficam na pasta do contribuinte.
