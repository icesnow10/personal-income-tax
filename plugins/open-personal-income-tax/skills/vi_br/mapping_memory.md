# mapping_memory — B3 movement → action

Living memory: how each B3 **Movimentação** type affects the position. The build script reads
the table below and writes it verbatim into the workbook's `aux_mapping` sheet. When the script
warns about an **unmapped** movement type, add a row here (do not edit the code).

This file is **generic** (B3's vocabulary, not taxpayer-specific) and is safe to share.

Two orthogonal axes per row:

- **action** (`credito` / `debito`) — how the row moves the **position** (quantity + cost):

| action | effect |
|---|---|
| `purchase` | + quantity, + cost |
| `sale` | − quantity; − cost pelo CUSTO MÉDIO vigente (quantidade × preço médio), NÃO pelo valor recebido (o preço de venda só importa para ganho de capital). Saída total zera o custo |
| `rebase` | DEFINE a quantidade para o valor da linha (não soma), custo total preservado — grupamento: a B3 credita a nova quantidade total |
| `return_of_capital` | − cost, quantity unchanged (signed by entry_type so a Credito+Debito pair cancels) |
| `snapshot` | does not move quantity/cost; the row's `quantity` is treated as B3's authoritative position on that date — the engine anchors the ticker's qty to it IFF no purchase/sale of that ticker happens after the latest snapshot row (B3 typically emits these after a merger / conversion / restructure that the action rules can't decode) |
| `no_action` | none |

- **provento_type** — does NOT touch the position, only labels the row as **income** for the
  `income_received` summary. Empty for non-income rows.

| provento_type | description |
|---|---|
| `dividend` | dividendo |
| `interest_on_equity` | juros sobre capital próprio (JCP) |
| `yield` | rendimento (FII / FI-Infra / juros de renda fixa) |
| `return_of_capital` | devolução de capital (também ajusta o custo via `action`) |

| entry_movement | credito | debito | provento_type | logic |
|---|---|---|---|---|
| Compra | purchase | sale |  | compra: aumenta a quantidade e o custo |
| Transferência - Liquidação | purchase | sale |  | liquidação de negócio: compra(+) / venda(−) quantidade e custo |
| COMPRA/VENDA DEFINITIVA A TERMO | purchase | sale |  | compra/venda a termo: move quantidade e custo |
| Bonificação em Ativos | purchase | sale |  | bonificação: aumenta a quantidade pelo custo informado (dilui o preço médio) |
| Desdobro | purchase | sale |  | desdobramento: a B3 credita as cotas ADICIONAIS (delta) com custo zero — soma à quantidade e dilui o preço médio |
| Grupamento | rebase | no_action |  | grupamento: a B3 credita a NOVA quantidade total (ex.: 23 cotas a 10:1 -> credita 2,3) — define a quantidade, custo total preservado, eleva o preço médio |
| Fração em Ativos | purchase | sale |  | ajuste de fração na quantidade, sem alterar o custo total |
| Leilão de Fração | no_action | no_action |  | venda da fração leiloada: registra só o caixa (a quantidade saiu via Fração em Ativos) |
| Direitos de Subscrição - Exercido | no_action | no_action |  | passo de processo (direito consumido): a quantidade chega via Recibo de Subscrição |
| Recibo de Subscrição | purchase | no_action |  | crédito: cota recebida na subscrição (+qtd). débito: recibo consumido na conversão para a cota principal (neutro) |
| Solicitação de Subscrição | no_action | no_action |  | passo de processo (solicitação): a quantidade chega via Recibo de Subscrição |
| Resgate | sale | sale |  | redenção de cotas: saída da posição (reduz a quantidade e o custo) |
| Amortização | return_of_capital | return_of_capital | return_of_capital | devolução de capital: reduz o custo e mantém a quantidade |
| AMORTIZAÇÃO | return_of_capital | return_of_capital | return_of_capital | amortização (variante em maiúsculas, renda fixa): devolução de capital, reduz o custo |
| Restituição de Capital | return_of_capital | return_of_capital | return_of_capital | devolução de capital: reduz o custo e mantém a quantidade |
| Restituição de Capital - Transferida | return_of_capital | return_of_capital | return_of_capital | devolução de capital (par Credito+Debito cancela: bookkeeping de transferência) |
| Rendimento | no_action | no_action | yield | rendimento: não altera quantidade nem custo (contabilizado como provento) |
| Rendimento - Transferido | no_action | no_action | yield | rendimento (transferido): mesmo tratamento de Rendimento |
| Dividendo | no_action | no_action | dividend | provento: não altera quantidade nem custo (contabilizado como provento) |
| Dividendo - Transferido | no_action | no_action | dividend | dividendo (transferido): mesmo tratamento de Dividendo |
| Juros Sobre Capital Próprio | no_action | no_action | interest_on_equity | JCP: não altera quantidade nem custo (contabilizado como provento) |
| Juros Sobre Capital Próprio - Transferido | no_action | no_action | interest_on_equity | JCP (transferido): mesmo tratamento de JCP |
| PAGAMENTO DE JUROS | no_action | no_action | yield | juros de renda fixa (debênture/CRA): não altera quantidade nem custo |
| Atualização | snapshot | snapshot |  | atualização de posição: sem efeito no custo ou quantidade |
| Cessão de Direitos | no_action | no_action |  | cessão de direitos de subscrição: não afeta a posição principal |
| Cessão de Direitos - Solicitada | no_action | no_action |  | evento neutro: não altera quantidade nem custo |
| Direito de Subscrição | no_action | no_action |  | evento neutro: não altera quantidade nem custo |
| Direitos de Subscrição - Não Exercido | no_action | no_action |  | evento neutro: não altera quantidade nem custo |
| Direito Sobras de Subscrição | no_action | no_action |  | evento neutro: não altera quantidade nem custo |
| Direito Sobras de Subscrição - Não Exercido | no_action | no_action |  | evento neutro: não altera quantidade nem custo |
| Transferência | no_action | no_action |  | transferência de custódia: neutra (pares entrada/saída se anulam; sobra avulsa conta como cota real) |
| Transferencia | no_action | no_action |  | transferência de custódia: neutra |
| TRANSFERENCIA SEM FINANCEIRO | no_action | no_action |  | transferência sem financeiro: neutra |
| VENCIMENTO | no_action | no_action |  | vencimento de renda fixa: fora deste controle |
| COMPRA / VENDA | no_action | no_action |  | renda fixa (CDB): fora do controle de ações/FII |
| COMPRA/VENDA | no_action | no_action |  | renda fixa (CRA/CDB/debênture): fora do controle de ações/FII |
| COMPRA/VENDA DEFINITIVA/CESSAO | no_action | no_action |  | renda fixa (CRA/debênture): fora do controle de ações/FII |
