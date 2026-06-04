---
name: read
description: Read every broker/bank informe in resources/ and transcribe them into the unified processed/informes.json that both /generate and /completeness consume. The agent (not a script) reads each PDF/print — layouts vary too much and many are image/encrypted — and emits one object per asset/rendimento, tagged by IRPF ficha (bens / isentos / exclusiva), with key, grupo/código, CNPJ, value, quantidade and the source PDF. Use when the user wants to read/transcribe the informes, build informes.json, or mentions ler os informes, transcrever, /read.
---

# /read — ler os informes e transcrever para `processed/informes.json`

Primeira metade do antigo `/consolidate`, agora explícita: **ler todo o `resources/`** e produzir a
transcrição unificada que alimenta o resto do pipeline.

> **Por que é trabalho do agente, não de um script:** os informes são PDFs com layouts muito variados
> e muitos são **imagem/cifrados** (Bradesco, Itaú, Nomad, etrade) — não saem em texto. A leitura e
> normalização são do agente; o script `completeness.py extract` é só **auxílio de parsing**.

## Saída — `processed/informes.json` (unificado)

Um único objeto com três listas (uma por ficha). É o **mesmo** arquivo lido por
[/generate](../consolidate/SKILL.md) e [/completeness](../completeness/SKILL.md). Schema completo em
[../consolidate/REFERENCE.md](../consolidate/REFERENCE.md):

```json
{
  "bens": [
    {"key":"NUIF11","b3":true,"grupo":7,"codigo":10,"cnpj":"40963403000150","nome":"FI-Infra exemplo","quantidade":10,"valor_2025":1000.00,"discriminacao":"10 COTAS DO FI-INFRA NUIF11 (exemplo)","source":"corretora_informe"},
    {"key":"NU-RDB","b3":false,"grupo":4,"codigo":2,"localizacao":105,"nome":"RDB Nu Financeira","descr":"RDB Nu Financeira","cnpj":"30680829000143","valor_2025":50000.00,"valor_2024":40000.00,"discriminacao":"SALDO EM RDB (exemplo)","source":"nubank_informe"},
    {"key":"AAPL","b3":false,"grupo":3,"codigo":1,"localizacao":249,"nome":"Apple Inc","descr":"10 ações AAPL na Avenue (EUA)","cnpj":"","valor_2025":12345.67,"valor_2024":0.0,"source":"avenue_informe"}
  ],
  "isentos":   [{"codigo":9,"key":"BBSE3","beneficiario":"Titular","cnpj":"...","fonte_pagadora":"...","descr":"Dividendos BBSE3","valor_2025":842.87,"source":"btg_787 + nubank_informe"}],
  "exclusiva": [{"codigo":10,"key":"VALE3","beneficiario":"Titular","cnpj":"...","descr":"JCP VALE3","valor_2025":41.07,"source":"btg_787 + nubank_informe"}]
}
```

- **`key`**: ticker / código do título / id estável (MAIÚSCULO) — o **mesmo** que o b3 usa para ativos B3.
- **`b3`** (só em `bens`): `true` se é ativo custodiado na B3 (valor sai do `b3_brazil_variable_income_avg_price_calculation.xlsx`);
  para esses, transcreva `quantidade` (e `valor_2025` do informe quando houver, p/ cross-check). `false`
  para o que só existe no informe (contas, RDB, moeda, **exterior**, JCP a receber) — aí o valor é declarado.
- **Transcreva TODOS os campos que o informe traz** por item, em **todos os grupos** (Bens e Direitos,
  Dividendos, Rendimentos Isentos e Não Tributáveis, JCP/Tributação Exclusiva): **`grupo`, `codigo`,
  `cnpj`, `nome`** (nome do ativo/empresa/fundo), **`localizacao`** (105 Brasil; ≠105 exterior, ex.: 249
  EUA), **`ticker`/`key`** e **`discriminacao`** — o texto de *"Discriminação sugerida para a Declaração
  de Bens e Direitos"* que BTG e outras corretoras já entregam pronto (transcreva **verbatim** quando
  existir; o /consolidate prefere essa discriminação à `descr`).
- **Exterior**: posições no exterior que o informe listar (ex.: o BTG/ a corretora estrangeira traz
  ações/ETF com `localizacao` 249/EUA) **entram no `informes.json`** como `b3:false` com a sua
  `localizacao` — não as deixe de fora (furo comum: o informe as lista e a transcrição esquece). Renda variável no exterior vai em
  Bens e Direitos; o **rendimento** dela é tributável (carnê-leão/GCAP), não isento/exclusiva.
- **`source`**: o **PDF** de onde veio o valor (ref curta, ex.: `btg_787`); junte com ` + ` quando o
  valor soma mais de um PDF (regime de caixa entre corretoras). Vira a coluna de rastreio.

## Workflow

1. **Liste o que cada PDF carrega** (auxílio de parsing):
   `python ../completeness/scripts/completeness.py extract resources/`
   PDFs imagem/cifrados não saem — abra-os com a ferramenta **Read** e transcreva à mão.
2. **Para cada arquivo em `resources/`**, extraia os itens e classifique por ficha (`bens` /
   `isentos` / `exclusiva`), aplicando as regras de [REFERENCE.md](REFERENCE.md) (isento×exclusiva,
   provento por escriturador, regime de caixa, incorporações/renames, não-B3, etc.).
3. **Escreva `processed/informes.json`** com as três listas. Para um esqueleto vazio:
   `python ../consolidate/scripts/generate.py --template processed/informes.json`.
4. **Siga para [/generate](../consolidate/SKILL.md)** e depois **[/completeness](../completeness/SKILL.md)**
   (loop build → verify) — toda pendência do relatório volta para ajustar este `informes.json`.

## Importante
- **Não é orientação fiscal.** Transcreva o que o informe diz; divergências são revisão, não correção.
- **Valor de ativo B3** é cross-check aqui — a fonte da verdade do valor é o `b3_brazil_variable_income_avg_price_calculation.xlsx`
  (custo com amortização abatida e custódias consolidadas), não o saldo de mercado do informe.
- Mantenha **genérico** — `informes.json` fica na pasta do contribuinte; nunca comite tickers, CNPJs
  ou valores reais.
