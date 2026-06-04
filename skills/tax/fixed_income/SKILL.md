---
name: fixed_income
description: Documentação da fatia de renda fixa do IRPF (CDB, CRA, CRI, debênture, LCI, LCA, Tesouro). NÃO gera arquivo próprio — a renda fixa flui /read → processed/informes.json → /consolidate como qualquer outro bem do informe. A validação "todo título de RF na Posição B3 tem item no informe" vive agora no /completeness (--posicao). Use quando o usuário perguntar como a renda fixa entra no IRPF, ou mencionar renda fixa, CDB/CRA/CRI/debênture/Tesouro, /fixed_income.
---

# /fixed_income — renda fixa no IRPF (documentação; sem arquivo próprio)

Renda fixa da B3 (**CDB / CRA / CRI / debênture / LCI / LCA / Tesouro**) **não tem preço médio
reconstruível** — o custo embute juros decorridos que a B3 não separa e o principal amortiza. Por isso
o valor de Bens e Direitos é o **saldo do informe da corretora**, transcrito pelo [/read](../read/SKILL.md).

> **Este skill não roda script nem gera `.xlsx`.** A renda fixa é só mais um conjunto de itens do
> `informes.json` (grupo 04: 04/02 tributado, 04/03 isento) e entra na declaração final pelo
> [/consolidate](../consolidate/SKILL.md), direto do `informes.json` — sem arquivo intermediário e sem
> dupla contagem. O único entregável é o `irpf_consolidated.xlsx`.

## Como a renda fixa flui (sem etapa própria)
```
/read ──► processed/informes.json ──► /consolidate ──► irpf_consolidated.xlsx
 (LLM lê os PDFs)   (RF = grupo 04)     (script junta)     (Bens + Isentos + Exclusiva)
```
- **Bens e Direitos**: cada título de RF é um item `bens` `b3:false` (04/02 CDB/Tesouro tributado;
  04/03 CRA/CRI/LCI/LCA/debênture incentivada isento) — **valor do informe**.
- **Isentos (cód 12)**: juros de CRA/CRI/LCI/LCA/debênture incentivada — do informe.
- **Exclusiva (cód 06)**: rendimento de CDB/Tesouro/aplicações tributado na fonte — do informe.

## Validação Posição B3 × informe → agora no /completeness
A checagem "todo título de renda fixa na **Posição B3** tem um item correspondente no `informes.json`"
(sinal de transcrição faltando no /read) migrou para o [/completeness](../completeness/SKILL.md):
```
python ../completeness/scripts/completeness.py compare \
    --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx \
    --informes processed/informes.json --posicao resources/POS.xlsx
```
A seção **"Renda fixa — Posição B3 × informe"** do relatório lista cada título de RF na Posição sem item
no informe (`⚠️ FALTA no informe`). Confira o informe da corretora e adicione no `informes.json` (volta
ao /read). Só sinaliza, não trava.

## Importante
- **Não é orientação fiscal.** O valor de RF vem do informe; divergência de quantidade é sinal de revisão.
- **CRA/CRI/debênture** embutem juros decorridos — o valor **tem** que vir do informe, nunca de uma
  reconstrução. CDB/LCI/LCA/Tesouro idem (valor aplicado do informe).
- Mantenha **genérico** — nunca comite o `informes.json` de um contribuinte.
