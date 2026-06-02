---
name: irpf
description: Orchestrate the full Brazilian IRPF investments pipeline end to end — b3 → consolidate → completeness. Runs b3 to reconstruct preço médio / custo from the B3 exports (brazil_investments.xlsx), consolidate (which itself runs /read to transcribe every informe in resources/ into processed/informes.json, then /generate to build the final IRPF fichas irpf_consolidated.xlsx), then completeness to audit the result against the informes (completeness_report.md). Use when the user wants to "fazer o IRPF dos investimentos", run the whole pipeline, montar a declaração de investimentos do zero, or mentions /irpf.
---

# /irpf — pipeline completo dos investimentos no IRPF

Orquestra os três skills, na ordem, do export da B3 até o relatório de conferência:

```
resources/ (B3 exports + informes)
        │
        ▼
   [1] b3 ──────────────►  processed/brazil_investments.xlsx   (preço médio / custo — fonte da verdade do valor)
        │
        ▼
   [2] consolidate ─────►  irpf_consolidated.xlsx              (as 3 abas/fichas prontas p/ digitar)
        │  └─ /read     → processed/informes.json (transcrição unificada de todo o resources/)
        │  └─ /generate → irpf_consolidated.xlsx
        ▼
   [3] completeness ────►  completeness_report.md              (auditoria b3_source × informes.json)
```

Cada etapa é um skill próprio — este orquestrador só encadeia. O `consolidate` é ele mesmo um
orquestrador de `/read` + `/generate`. Veja [b3](../b3/SKILL.md), [consolidate](../consolidate/SKILL.md)
(→ [read](../read/SKILL.md) + [generate](../generate/SKILL.md)) e [completeness](../completeness/SKILL.md).

`read` e `completeness` consomem o **mesmo** `processed/informes.json` — uma transcrição só.

## Pasta do contribuinte (layout padrão)

Rode tudo da pasta do contribuinte (ex.: `contribuinte/`), num layout de 3 camadas:

| Pasta | Conteúdo |
|---|---|
| `resources/` | **cru**: os B3 exports (`Movimentação`/`Posição`) + os PDFs dos informes |
| `memory/` | os memory files do b3 (`ticker_memory.md`, `rf_memory.md`, `rf_value_memory.md`, `mapping_memory.md`) |
| `processed/` | **derivado**: `brazil_investments.xlsx` + a transcrição `informes.json` |
| (raiz) | os **entregáveis**: `irpf_consolidated.xlsx` + `completeness_report.md` |

## Entradas que o usuário fornece

- **`resources/MOV.xlsx`** — B3 "Movimentação" (todos os anos desde a primeira compra).
- **`resources/POS.xlsx`** — B3 "Posição" em 31/12 do ano-base (opcional: `resources/POS_PRIOR.xlsx`).
- **PDFs dos informes** em `resources/` (BTG, NU, Itaú, Wise, Nomad, …).
- **`--year YYYY`** — o ano-base da declaração.

## Workflow

### 1. b3 → `processed/brazil_investments.xlsx`
Atualize o `memory/ticker_memory.md` do contribuinte (renames/incorporações), então:
```
python ../b3/scripts/build_bens_direitos.py resources/MOV.xlsx resources/POS.xlsx \
    processed/brazil_investments.xlsx --memory-dir memory --year YYYY \
    [--posicao-anterior resources/POS_PRIOR.xlsx]
```
**Leia o AUDIT** impresso no fim — cada divergência de quantidade movimentação×posição é uma ação
corporativa a tratar à mão antes de seguir.

### 2. consolidate → `irpf_consolidated.xlsx`  (= /read + /generate)
O [consolidate](../consolidate/SKILL.md) orquestra as duas sub-etapas:
- **/read** — leia **todo** o `resources/` e transcreva para `processed/informes.json` (trabalho do
  agente; layouts variam demais e muitos PDFs são imagem/cifrados). Auxílio de parsing:
  `python ../completeness/scripts/completeness.py extract resources/` (cifrados leia com a ferramenta Read).
  Esqueleto vazio: `python ../generate/scripts/generate.py --template processed/informes.json`.
- **/generate** — monte as fichas (saída na raiz):
  ```
  python ../generate/scripts/generate.py --investimentos processed/brazil_investments.xlsx \
      --json processed/informes.json --outdir .
  ```
  Gera **um** arquivo com as abas `bens_e_direitos`, `isentos` e `exclusiva_definitiva`.
  **Leia o AUDIT**: ativo B3 na `brazil_investments.xlsx` ausente do `informes.json` é listado —
  adicione no JSON (volta ao /read) e re-rode o /generate.

### 3. completeness → `completeness_report.md`  (verify + fix)
```
python ../completeness/scripts/completeness.py compare \
    --investimentos processed/brazil_investments.xlsx --informes processed/informes.json
```
Confere `b3_source × informes.json` por ficha **e audita/edita o `irpf_consolidated.xlsx`** (achado ao
lado do relatório): valor fora da autoridade da ficha (Bens e Direitos → b3_source; rendimentos →
informe) é **corrigido no consolidado** e anotado na coluna `obs_completeness`; o `.md` ganha a seção
**Ajustes** com o que foi editado/conferido. Trata também classificação (vale o informe), aglutinações
e itens não-B3 a declarar à mão. Use `--no-apply` para auditar sem editar. **Leia o `.md`** e os action
items.

### 4. Loop build → verify
Cada pendência do `completeness_report.md` volta para a etapa certa: divergência de classificação ou
ativo não-B3 → ajuste o `processed/informes.json` (/read) e re-rode o /generate; `FALTA em b3_source
(B3!)` → investigue o `/b3` (etapa 1). Repita até o relatório fechar.

## Saídas

| Arquivo | Skill | Conteúdo |
|---|---|---|
| `processed/brazil_investments.xlsx` | b3 | preço médio/custo por ticker + abas IRPF auxiliares |
| `processed/informes.json` | consolidate (/read) | transcrição unificada dos informes (bens/isentos/exclusiva) |
| `irpf_consolidated.xlsx` | consolidate (/generate) | as 3 abas/fichas finais, prontas p/ digitar no programa |
| `completeness_report.md` | completeness | auditoria por ficha, divergências e action items |

## Importante
- **Não é orientação fiscal.** Toda divergência é sinal de revisão, nunca correção automática.
- **Fontes da verdade:** valor/custo → `b3_source`; classificação (grupo/código/CNPJ) e rendimentos →
  o informe. Os scripts respeitam isso; não sobreponha à mão sem rastrear ao PDF.
- Mantenha tudo **genérico** — `informes.json`, os `*.xlsx` e os memory files preenchidos ficam na
  pasta de trabalho; nunca comite tickers, CNPJs ou valores de um contribuinte.
