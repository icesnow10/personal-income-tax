# /read — reference (transcribing the informes into `informes.json`)

`/read` **builds** the transcription; `/completeness` **checks** it against the b3 output. Fill the
JSON honestly from the informes and let `/completeness` flag divergences (loop build → verify).

## Regras ao montar o `informes.json`
- **Valor de ativo B3** sempre do `b3_brazil_variable_income_avg_price_calculation.xlsx` (custo já com amortização abatida e custódias
  consolidadas) — **nunca** o saldo do informe (que pode não abater amortização: NUIF11 informe 979,20 ×
  correto 961,20; ou ver só uma corretora: ALZR11 41,71 × ano todo 49,75). No JSON, para B3, `quantidade`
  e `valor_2025` (do informe) são só para o cross-check do `/completeness`.
- **Proventos**: o **escriturador/administrador** (Bradesco, Porto, gestora do fundo) é mais autoritativo
  que a corretora. Some o que foi **creditado no ano** entre corretoras (regime de caixa) e registre as
  parcelas no `source` (ex.: `btg_787 + nubank_informe`).
- **Isento (99) × Exclusiva (06/10)**: teste = houve **IR retido na fonte**? Sem IR → isento; com IR →
  exclusiva. Cuidado com "Rendimento" de **ação** (ex.: BBSE3 7,41): é JCP/10, não isento/99.
- **Valores informe-only** (não constam na B3, ex.: IRDM11 387,38 tributado na fonte; RDB; come-cotas;
  contas; moeda; exterior) entram pelo informe como `b3:false` (bens) ou direto nas listas de rendimento.
- **Incorporação/rename** (IRDM11→IRIM11): o valor B3 já vem aglutinado no ticker novo; transcreva os
  rendimentos por CNPJ pagador com o `key` do papel — o `completeness` reconcilia a soma via os renames
  do `aux_mapping` do workbook.
- Confira sempre contra a **declaração do ano anterior** (abertura = fechamento anterior). A
  **pré-preenchida** herda erros do informe — cruze, não copie cego.

## Reading the informes (per source)
- Use o `extract` do skill `completeness` para listar linhas candidatas; PDFs imagem/cifrados
  (Bradesco, Itaú, etrade, Nomad) **não** saem no texto — transcreva à mão lendo o PDF com a ferramenta Read.
- **BTG**: tabela limpa com `código CNPJ grupo cód saldo2024 saldo2025 … rendimento`. Para B3, use o
  **grupo/cód** daqui (autoridade), mas lembre que o **valor** do B3 vem da `b3_brazil_variable_income_avg_price_calculation.xlsx`.
- **Nubank/Wise Conta Global**: saldo em moeda estrangeira (EUR/USD em 31/12) → Bens e Direitos
  não-B3 (depósito no exterior, `localizacao` = código do país). A Wise BR (CNPJ 40.571.694/0001-31)
  costuma ter saldo 0.
- **Renda fixa transferida** (saldo 0 no informe de origem): declarar no destino, não duplicar.

## Mapping each informe line to a ficha
| O que é | ficha | código típico |
|---|---|---|
| Ação / BDR / FII / FI-Infra / Tesouro / RF (posição em 31/12) | `bens` (b3 ou não) | grupo/cód do informe |
| Conta corrente, RDB/CDB, moeda, cripto, conta no exterior, JCP a receber | `bens` (`b3:false`) | 06/01, 04/02, 08/03, 62/01, 99/07 … |
| Dividendos de ação | `isentos` | 09 |
| Rendimento de FII / FIAgro / FI-Infra (sem IR) | `isentos` | 99 |
| JCP (juros sobre capital próprio) | `exclusiva` | 10 |
| Rendimento de RDB/CDB/Tesouro/fundo come-cotas (com IR na fonte) | `exclusiva` | 06 |

Schema completo dos campos em [../consolidate/REFERENCE.md](../consolidate/REFERENCE.md). Not tax advice.
