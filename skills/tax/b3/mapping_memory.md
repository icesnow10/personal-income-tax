# mapping_memory — B3 movement → action

Living memory: how each B3 **Movimentação** type affects the position. The build script reads
the table below and writes it verbatim into the workbook's `aux_mapping` sheet. When the script
warns about an **unmapped** movement type, add a row here (do not edit the code).

This file is **generic** (B3's vocabulary, not taxpayer-specific) and is safe to share.

Actions and their effect on the position:

| action | effect |
|---|---|
| `purchase` | + quantity, + cost |
| `sale` | − quantity, − cost |
| `return_of_capital` | − cost, quantity unchanged (devolução de capital) |
| `yield` / `dividend` / `interest_on_equity` | none — provento, doesn't touch cost |
| `no_action` | none — neutral event |

`credito` / `debito` = the action to apply depending on the row's Entrada/Saída.

| entry_movement | credito | debito | logic |
|---|---|---|---|
| Compra | purchase | sale | compra: aumenta a quantidade e o custo |
| Transferência - Liquidação | purchase | sale | liquidação de negócio: compra(+) / venda(−) quantidade e custo |
| COMPRA/VENDA DEFINITIVA A TERMO | purchase | sale | compra/venda a termo: move quantidade e custo |
| Bonificação em Ativos | purchase | sale | bonificação: aumenta a quantidade pelo custo informado (dilui o preço médio) |
| Desdobro | purchase | sale | desdobramento: aumenta a quantidade com custo zero (dilui o preço médio) |
| Grupamento | purchase | sale | grupamento: reduz a quantidade com custo zero (eleva o preço médio) |
| Fração em Ativos | purchase | sale | ajuste de fração na quantidade, sem alterar o custo total |
| Leilão de Fração | purchase | sale | venda da fração leiloada: reduz a quantidade |
| Direitos de Subscrição - Exercido | purchase | sale | exercício de direito: compra de cotas (aumenta a quantidade e o custo) |
| Recibo de Subscrição | purchase | sale | recibo do exercício: vira a cota principal (aumenta a quantidade e o custo) |
| Solicitação de Subscrição | purchase | sale | pagamento da subscrição: compra (aumenta a quantidade e o custo) |
| Resgate | sale | sale | redenção de cotas: saída da posição (reduz a quantidade e o custo) |
| Amortização | return_of_capital | return_of_capital | devolução de capital: reduz o custo e mantém a quantidade |
| Restituição de Capital | return_of_capital | return_of_capital | devolução de capital: reduz o custo e mantém a quantidade |
| Restituição de Capital - Transferida | return_of_capital | return_of_capital | devolução de capital: reduz o custo e mantém a quantidade |
| Rendimento | yield | yield | rendimento: não altera quantidade nem custo |
| Rendimento - Transferido | yield | yield | rendimento: não altera quantidade nem custo |
| Dividendo | dividend | dividend | provento: não altera quantidade nem custo |
| Dividendo - Transferido | dividend | dividend | provento: não altera quantidade nem custo |
| Juros Sobre Capital Próprio | interest_on_equity | interest_on_equity | provento (JCP): não altera quantidade nem custo |
| Juros Sobre Capital Próprio - Transferido | interest_on_equity | interest_on_equity | provento (JCP): não altera quantidade nem custo |
| Atualização | no_action | no_action | atualização de posição: sem efeito no custo ou quantidade |
| Cessão de Direitos | no_action | no_action | cessão de direitos de subscrição: não afeta a posição principal |
| Cessão de Direitos - Solicitada | no_action | no_action | evento neutro: não altera quantidade nem custo |
| Direito de Subscrição | no_action | no_action | evento neutro: não altera quantidade nem custo |
| Direitos de Subscrição - Não Exercido | no_action | no_action | evento neutro: não altera quantidade nem custo |
| Direito Sobras de Subscrição | no_action | no_action | evento neutro: não altera quantidade nem custo |
| Direito Sobras de Subscrição - Não Exercido | no_action | no_action | evento neutro: não altera quantidade nem custo |
| Transferência | no_action | no_action | transferência de custódia: neutra (pares entrada/saída se anulam; sobra avulsa conta como cota real) |
| Transferencia | no_action | no_action | transferência de custódia: neutra |
| TRANSFERENCIA SEM FINANCEIRO | no_action | no_action | transferência sem financeiro: neutra |
| VENCIMENTO | no_action | no_action | vencimento de renda fixa: fora deste controle |
| COMPRA / VENDA | no_action | no_action | renda fixa (CDB): fora do controle de ações/FII |
