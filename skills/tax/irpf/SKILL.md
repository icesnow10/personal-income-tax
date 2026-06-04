---
name: irpf
description: Orchestrate the full Brazilian IRPF investments pipeline end to end — read → vi_br → consolidate → completeness. /read (LLM) transcribes every informe into informes.json (the single downstream source: renda fixa, exterior, contas, RDB e todos os rendimentos); vi_br (script) reconstructs preço médio of B3 ações/FII/BDR; consolidate merges the b3 value + informes.json into irpf_consolidated.xlsx; completeness audits it and validates RF×B3 position. Só a extração de PDF/PNG é trabalho de LLM; o resto é script. Use when the user wants to "fazer o IRPF dos investimentos", run the whole pipeline, montar a declaração de investimentos do zero, or mentions /irpf.
---

# /irpf — pipeline completo dos investimentos no IRPF

**Princípio:** só a **extração de PDF/PNG** é trabalho de LLM (o `/read`); todo o resto é **script
determinístico** e a **única fonte a jusante é o `processed/informes.json`**. O único entregável é o
`irpf_consolidated.xlsx` (+ o workbook do vi_br como derivado de RV-Brasil).

```
resources/ (informes PDFs + B3 exports)
        │
  [1] read  (LLM) ──────────────────►  processed/informes.json   (transcrição unificada de TODO o resources/:
        │                                renda fixa, exterior, contas, RDB, JCP a receber e TODOS os rendimentos)
  [2] vi_br (script) ───────────────►  b3_brazil_variable_income_avg_price_calculation.xlsx  (só preço médio ações/FII/BDR da B3)
        │     (renda fixa e exterior NÃO têm etapa própria — vêm do informes.json no consolidate)
  [3] consolidate (script) ─────────►  irpf_consolidated.xlsx    (valor RV-Brasil do workbook + TODO o resto do informes.json → 3 fichas)
        │
  [4] completeness (LLM+script) ────►  completeness_report.md    (auditoria por ficha + valida RF×Posição B3 + edita o consolidado)
```

Skills: [read](../read/SKILL.md), [vi_br](../vi_br/SKILL.md), [consolidate](../consolidate/SKILL.md),
[completeness](../completeness/SKILL.md). Documentação (sem script, dado via informes.json):
[variable_income](../variable_income/SKILL.md) / [vi_international](../vi_international/SKILL.md) (exterior)
e [fixed_income](../fixed_income/SKILL.md) (renda fixa). O **mesmo** `processed/informes.json` alimenta
consolidate e completeness — uma transcrição só.

## Pasta do contribuinte (layout padrão)
Rode tudo da pasta do contribuinte (ex.: `contribuinte/`), num layout de 3 camadas:

| Pasta | Conteúdo |
|---|---|
| `resources/` | **cru**: os B3 exports (`Movimentação`/`Posição`) + os PDFs dos informes |
| `memory/` | os memory files (`ticker_memory.md`, `mapping_memory.md`, `escriturador_memory.md`, …) |
| `processed/` | **derivado**: `informes.json` + o workbook de RV-Brasil (vi_br) |
| (raiz) | os **entregáveis**: `irpf_consolidated.xlsx` + `completeness_report.md` |

## Entradas que o usuário fornece
- **PDFs dos informes** em `resources/` (BTG, NU, Itaú, BB, Wise, Nomad, …) — a baseline.
- **`resources/MOV.xlsx` / `POS.xlsx`** — B3 "Movimentação" (todos os anos) / "Posição" em 31/12 do
  ano-base (opcional `POS_PRIOR.xlsx`) — só se tiver ativos na B3.
- **`--year YYYY`** — o ano-base.

## Workflow

### 1. read (LLM) → `processed/informes.json`
Leia **todo** o `resources/` e transcreva (trabalho do agente; layouts variam e muitos PDFs são
imagem/cifrados). **É a única etapa de LLM** e a única fonte a jusante: transcreva por item `grupo`,
`codigo`, `cnpj`, `nome`, `localizacao`, `ticker` e a **discriminação sugerida** — em **todos os grupos**
(bens, dividendos, isentos, JCP/exclusiva), **incluindo exterior** (localização ≠ 105). Auxílio:
`python ../completeness/scripts/completeness.py extract resources/` (cifrados, leia com a ferramenta
Read). Esqueleto: `python ../consolidate/scripts/generate.py --template processed/informes.json`.

> Renda fixa, ações/ETF no exterior, contas, RDB e JCP a receber **não têm etapa própria** — são itens
> do `informes.json` e entram no consolidate direto. Veja [fixed_income](../fixed_income/SKILL.md) e
> [vi_international](../vi_international/SKILL.md) (documentação, sem script).

### 2. vi_br (script) → `b3_brazil_variable_income_avg_price_calculation.xlsx`
Só renda variável **da B3** (ações/FII/BDR). Atualize `memory/ticker_memory.md` (renames/incorporações):
```
python ../vi_br/scripts/build_bens_direitos.py resources/MOV.xlsx resources/POS.xlsx \
    processed/b3_brazil_variable_income_avg_price_calculation.xlsx --memory-dir memory --year YYYY \
    [--posicao-anterior resources/POS_PRIOR.xlsx]
```
**Leia o AUDIT** (qtd movimentação × posição): cada divergência é ação corporativa a tratar à mão.

### 3. consolidate (script) → `irpf_consolidated.xlsx`
```
python ../consolidate/scripts/generate.py \
    --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx \
    --json processed/informes.json --outdir .
```
Puxa só o **valor** de RV-Brasil do workbook (b3:true por `key`) e monta **todo o resto** direto do
`informes.json` (RF, exterior, contas, RDB, JCP a receber, dividendos/JCP) nas abas `bens_e_direitos`,
`isentos`, `exclusiva_definitiva`. **Leia o AUDIT**: ativo de RV-Brasil ausente do `informes.json` é
listado — adicione no JSON (volta ao /read) e re-rode.

### 4. completeness (LLM + script) → `completeness_report.md`  (verify + fix)
Popule o escriturador direto da B3: `python ../completeness/scripts/fetch_escriturador.py --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx`.
Então:
```
python ../completeness/scripts/completeness.py compare \
    --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx \
    --informes processed/informes.json --posicao resources/POS.xlsx
```
Confere por ficha, **valida RF × Posição B3** (`--posicao`, migrado do antigo fixed_income), sinaliza
rendimento de exterior em ficha errada **e audita/edita o `irpf_consolidated.xlsx`** para a autoridade
(Bens RV → workbook; rendimentos → informe), anotando em `obs_completeness`. Aqui o agente **pode reler
os PDFs/PNGs e buscar na web** para fechar pendências — cada correção volta para o `informes.json`.
**Leia o `.md`** e os action items. Repita até fechar.

## Saídas
| Arquivo | Skill | Conteúdo |
|---|---|---|
| `processed/informes.json` | read | transcrição unificada dos informes (fonte única a jusante) |
| `processed/b3_brazil_variable_income_avg_price_calculation.xlsx` | vi_br | preço médio de ações/FII/BDR (B3) |
| `irpf_consolidated.xlsx` | consolidate | as 3 fichas finais, prontas p/ digitar (entregável) |
| `completeness_report.md` | completeness | auditoria por ficha, RF×B3, divergências e action items |

## Importante
- **Não é orientação fiscal.** Toda divergência é sinal de revisão, nunca correção automática.
- **Fontes da verdade:** valor de RV-Brasil → o workbook do vi_br (preço médio); valor de RF, exterior e
  **todos** os rendimentos → o **informe** (`informes.json`). Os scripts respeitam isso.
- Mantenha tudo **genérico** — nunca comite `informes.json`, os `*.xlsx` ou os memory files de um contribuinte.
