---
name: vi_international
description: Build the FOREIGN variable-income slice of the IRPF (ações/ETFs no exterior, held at Avenue, Nomad, Interactive Brokers, etrade, …) from the informes transcription. These are NOT on the B3, so there is no preço médio and no Posição — the value comes from the broker informe. Reads processed/informes.json (the foreign equity/ETF items, b3:false + localização ≠ 105) and writes bens_e_direitos (+ a flagged rendimentos_exterior sheet, since foreign income goes in OTHER fichas). Use when the user has foreign stocks/ETFs, or mentions ações/ETF no exterior, Avenue/Nomad/Interactive Brokers, renda variável internacional, /vi_international.
---

# /vi_international — renda variável no exterior (a partir dos informes)

Ações/ETFs **no exterior** (Avenue, Nomad, Interactive Brokers, etrade, …) **não estão na B3** — não
há `Movimentação`/`Posição` nem preço médio a reconstruir. O valor de Bens e Direitos vem do **informe
da corretora estrangeira**. Este skill só pega esses itens do `informes.json` e monta a fatia.

## Entrada
- **`processed/informes.json`** — a transcrição do [/read](../read/SKILL.md). Daqui vêm as posições e o
  valor (a autoridade é o informe da corretora estrangeira).

## Saída — `processed/variable_income_international.xlsx`
| Sheet | Conteúdo |
|---|---|
| `bens_e_direitos` | ações/ETFs no exterior (valor do informe, **localização ≠ 105**) |
| `rendimentos_exterior` | rendimento estrangeiro que o informe trouxe — **sinalizado**, porque dividendo/ganho no exterior vai em OUTRA ficha (Rendimentos Tributáveis Recebidos de PF/Exterior · carnê-leão · GCAP), **não** em isento/exclusiva |

## Escopo (de quem é o quê)
Dono dos itens `b3:false`, **estrangeiros** (localização ≠ 105) e de **renda variável** (grupo 3/7/31 —
ação/fundo-ETF). **NÃO** é dono de caixa no exterior (grupo 06), cripto (grupo 08) nem RSU
(skill `rsu` próprio) — esses ficam com o [/consolidate](../consolidate/SKILL.md) ou seu skill.

## Workflow
1. Rode o [/read](../read/SKILL.md) primeiro (gera `informes.json`).
2. `python scripts/build_international.py --informes processed/informes.json --out processed/variable_income_international.xlsx`
3. **Bens** entram na declaração final via [/consolidate](../consolidate/SKILL.md). **Rendimento no
   exterior** (dividendos/ganhos) é TRIBUTÁVEL — não vai nas 3 fichas; trate na ficha correta à mão
   (a aba `rendimentos_exterior` lista o que apareceu, para conferência).

## Importante
- **Não é orientação fiscal.** O custo vem do informe estrangeiro; conversão moeda→BRL pela PTAX do
  BCB (compra de 31/12 p/ saldo). Offshore Lei 14.754/2023 tem regra própria — confirme com contador.
- Mantenha **genérico** — nunca comite dados de um contribuinte.
