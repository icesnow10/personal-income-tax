---
name: renda_fixa
description: Build the IRPF renda-fixa slice (CDB, CRA, CRI, debênture, LCI, LCA, Tesouro) from the informes transcription plus a B3 position validation. Reads processed/informes.json (the RF items — value from the broker informe) and the B3 Posição export (to validate that every renda-fixa security held on the B3 is covered by an informe), and writes an Excel with position + bens_e_direitos + isentos (cód 12) + exclusiva (cód 06), renda fixa only. Use when the user wants to build the renda-fixa part of the IRPF, validate RF positions against the informes, or mentions renda fixa, CDB/CRA/CRI/debênture/Tesouro, /renda_fixa.
---

# /renda_fixa — a fatia de renda fixa do IRPF (a partir dos informes + validação na B3)

Renda fixa da B3 (**CDB / CRA / CRI / debênture / LCI / LCA / Tesouro**) **não tem preço médio
reconstruível** — o custo embute juros decorridos que a B3 não separa e o principal amortiza. O valor
de Bens e Direitos é o **saldo do informe da corretora**. Este skill não reconstrói nada: ele pega os
itens de RF que o [/read](../read/SKILL.md) já transcreveu (`informes.json`) e usa a **Posição B3**
só para **validar** que todo título de RF na custódia tem um item no informe.

## Entradas
- **`processed/informes.json`** — a transcrição unificada (do /read). Daqui vêm os valores e
  rendimentos de RF (a autoridade do valor é o informe).
- **B3 Posição** (`resources/POS.xlsx`) — só para listar os títulos de RF e suas quantidades (validação).

## Saída — `processed/renda_fixa.xlsx`
| Sheet | Conteúdo |
|---|---|
| `position` | os títulos de RF da Posição B3 (código, produto, quantidade, corretora) + coluna `validacao` = se há item no informe |
| `bens_e_direitos` | bens de RF (04/02 tributado · 04/03 isento) — **valor do informe** |
| `isentos` | rendimentos isentos de RF — **código 12** (juros de CRA/CRI/LCI/LCA/debênture incentivada) |
| `exclusiva` | rendimentos exclusiva de RF — **código 06** (CDB/Tesouro/aplicações) |

## Escopo (de quem é o quê)
Este skill é dono **só da renda fixa da B3** — os títulos que aparecem na Posição B3. **NÃO** é dono de
RDB, conta de pagamento, fundo come-cotas, moeda/exterior (esses não estão na Posição B3 → ficam com o
[/consolidate](../consolidate/SKILL.md)). Dividendos/JCP de ação também não (são renda variável).

## Workflow
1. Rode o [/read](../read/SKILL.md) primeiro (gera `informes.json`).
2. `python scripts/build_renda_fixa.py --posicao resources/POS.xlsx --informes processed/informes.json --out processed/renda_fixa.xlsx`
3. **Leia a validação**: cada título da Posição B3 **sem item no informe** é sinalizado (`⚠️ FALTA no
   informe`) — confira o informe da corretora e adicione no `informes.json` (volta ao /read). A validação
   **só sinaliza**, não trava o pipeline.
4. O [/consolidate](../consolidate/SKILL.md) junta este excel + o workbook de renda variável + o resto do
   informe na declaração final (sem duplicar — os itens de RF daqui são removidos do informes.json lá).

## Importante
- **Não é orientação fiscal.** O valor de RF vem do informe; divergência de quantidade é sinal de revisão.
- **CRA/CRI/debênture** embutem juros decorridos — por isso o valor **tem** que vir do informe, nunca de
  uma reconstrução. CDB/LCI/LCA/Tesouro idem (valor aplicado do informe).
- Mantenha **genérico** — nunca comite o `informes.json`/`renda_fixa.xlsx` de um contribuinte.
