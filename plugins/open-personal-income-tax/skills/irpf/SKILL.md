---
name: irpf
description: Orchestrate the full Brazilian IRPF investments pipeline end to end — read → variable_income (brazil + international) → fixed_income → consolidate → completeness. /read transcribes every informe into informes.json; vi_br reconstructs preço médio of B3 ações/FII/BDR; vi_international builds foreign stocks/ETFs from the informes; fixed_income builds the RF slice (bens/isentos/exclusiva) + a B3 position validation; consolidate merges everything into irpf_consolidated.xlsx; completeness audits it. Use when the user wants to "fazer o IRPF dos investimentos", run the whole pipeline, montar a declaração de investimentos do zero, or mentions /irpf.
---

# /irpf — pipeline completo dos investimentos no IRPF

Orquestra os skills, nesta ordem, dos informes/B3 até o relatório de conferência:

```
resources/ (informes PDFs + B3 exports)
        │
  [1] read ─────────────────────────►  processed/informes.json   (transcrição unificada de TODO o resources/)
        │
  [2] variable_income                  (renda variável — Brasil + exterior):
        ├─ vi_br ────────────────────►  b3_brazil_variable_income_avg_price_calculation.xlsx  (preço médio ações/FII/BDR da B3)
        └─ vi_international ──────────► variable_income_international.xlsx                  (ações/ETF no exterior, do informe)
        │
  [3] fixed_income ─────────────────►  processed/fixed_income.xlsx (CDB/CRA/CRI/deb/Tesouro: bens+isentos+exclusiva, do informe + validação Posição B3)
        │
  [4] consolidate ──────────────────►  irpf_consolidated.xlsx    (junta RV-Brasil + RV-exterior + RF + o resto do informe → 3 fichas)
        │
  [5] completeness ─────────────────►  completeness_report.md    (auditoria + edita o consolidado p/ a autoridade)
```

Skills: [read](../read/SKILL.md), [variable_income](../variable_income/SKILL.md) (→
[brazil](../vi_br/SKILL.md) + [international](../vi_international/SKILL.md)),
[fixed_income](../fixed_income/SKILL.md), [consolidate](../consolidate/SKILL.md),
[completeness](../completeness/SKILL.md). O **mesmo** `processed/informes.json` alimenta vi_international,
fixed_income, consolidate e completeness — uma transcrição só.

## Pasta do contribuinte (layout padrão)
Rode tudo da pasta do contribuinte (ex.: `contribuinte/`), num layout de 3 camadas:

| Pasta | Conteúdo |
|---|---|
| `resources/` | **cru**: os B3 exports (`Movimentação`/`Posição`) + os PDFs dos informes |
| `memory/` | os memory files (`ticker_memory.md`, `mapping_memory.md`, `escriturador_memory.md`, …) |
| `processed/` | **derivado**: `informes.json`, o workbook de RV e o `fixed_income.xlsx` |
| (raiz) | os **entregáveis**: `irpf_consolidated.xlsx` + `completeness_report.md` |

## Entradas que o usuário fornece
- **PDFs dos informes** em `resources/` (BTG, NU, Itaú, BB, Wise, Nomad, …) — a baseline.
- **`resources/MOV.xlsx` / `POS.xlsx`** — B3 "Movimentação" (todos os anos) / "Posição" em 31/12 do
  ano-base (opcional `POS_PRIOR.xlsx`) — só se tiver ativos na B3.
- **`--year YYYY`** — o ano-base.

## Workflow

### 1. read → `processed/informes.json`
Leia **todo** o `resources/` e transcreva (trabalho do agente; layouts variam e muitos PDFs são
imagem/cifrados). Auxílio: `python ../completeness/scripts/completeness.py extract resources/`
(cifrados, leia com a ferramenta Read). Esqueleto: `python ../consolidate/scripts/generate.py --template processed/informes.json`.

### 2. variable_income → renda variável (Brasil + exterior)
**(a) Brasil** — atualize `memory/ticker_memory.md` (renames/incorporações), então:
```
python ../vi_br/scripts/build_bens_direitos.py resources/MOV.xlsx resources/POS.xlsx \
    processed/b3_brazil_variable_income_avg_price_calculation.xlsx --memory-dir memory --year YYYY \
    [--posicao-anterior resources/POS_PRIOR.xlsx]
```
**Leia o AUDIT** (qtd movimentação × posição): cada divergência é ação corporativa a tratar à mão.
**(b) Internacional** — ações/ETF no exterior (só do informe, não estão na B3):
```
python ../vi_international/scripts/build_international.py \
    --informes processed/informes.json --out processed/variable_income_international.xlsx
```

### 3. fixed_income → `processed/fixed_income.xlsx`
```
python ../fixed_income/scripts/build_fixed_income.py --posicao resources/POS.xlsx \
    --informes processed/informes.json --out processed/fixed_income.xlsx
```
**Leia a validação**: título de RF na Posição B3 **sem item no informe** é sinalizado — confira o
informe e adicione no `informes.json` (volta ao /read). Só sinaliza, não trava.

### 4. consolidate → `irpf_consolidated.xlsx`
```
python ../consolidate/scripts/generate.py \
    --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx \
    --json processed/informes.json --renda-fixa processed/fixed_income.xlsx \
    --internacional processed/variable_income_international.xlsx --outdir .
```
Junta RV-Brasil (workbook) + RV-exterior (internacional) + RF (fixed_income) + o resto do informe
(dividendos/JCP de ação, contas, RDB, JCP a receber) nas abas `bens_e_direitos`, `isentos`,
`exclusiva_definitiva`. Os itens de RF e de exterior do `informes.json` são removidos lá (já vêm dos
workbooks — sem dupla contagem). **Leia o AUDIT**: ativo de RV-Brasil ausente do `informes.json` é
listado — adicione no JSON e re-rode.

### 5. completeness → `completeness_report.md`  (verify + fix)
Popule o escriturador direto da B3: `python ../completeness/scripts/fetch_escriturador.py --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx`.
Então:
```
python ../completeness/scripts/completeness.py compare \
    --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx --informes processed/informes.json
```
Confere por ficha **e audita/edita o `irpf_consolidated.xlsx`** para a autoridade (Bens RV → workbook;
rendimentos → informe), anotando em `obs_completeness`. **Leia o `.md`** e os action items; cada pendência
volta para a etapa certa. Repita até fechar.

## Saídas
| Arquivo | Skill | Conteúdo |
|---|---|---|
| `processed/informes.json` | read | transcrição unificada dos informes |
| `processed/b3_brazil_variable_income_avg_price_calculation.xlsx` | vi_br | preço médio de ações/FII/BDR (B3) |
| `processed/variable_income_international.xlsx` | vi_international | ações/ETF no exterior (do informe) |
| `processed/fixed_income.xlsx` | fixed_income | RF: position + bens + isentos + exclusiva (valor do informe) |
| `irpf_consolidated.xlsx` | consolidate | as 3 fichas finais, prontas p/ digitar |
| `completeness_report.md` | completeness | auditoria por ficha, divergências e action items |

## Importante
- **Não é orientação fiscal.** Toda divergência é sinal de revisão, nunca correção automática.
- **Fontes da verdade:** valor de RV → o workbook de variable_income (preço médio); valor de RF e todos
  os rendimentos → o **informe**. Os scripts respeitam isso.
- Mantenha tudo **genérico** — nunca comite `informes.json`, os `*.xlsx` ou os memory files de um contribuinte.
