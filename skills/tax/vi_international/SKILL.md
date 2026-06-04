---
name: vi_international
description: Documentação da renda variável no exterior (ações/ETFs em Avenue, Nomad, Interactive Brokers, etrade, …). NÃO gera arquivo próprio — como não estão na B3, não há preço médio nem Posição; o valor vem do informe da corretora estrangeira e flui /read → processed/informes.json → /consolidate. Use quando o usuário tiver ações/ETF no exterior, ou mencionar renda variável internacional, Avenue/Nomad/Interactive Brokers, /vi_international.
---

# /vi_international — renda variável no exterior (documentação; sem arquivo próprio)

Ações/ETFs **no exterior** (Avenue, Nomad, Interactive Brokers, etrade, …) **não estão na B3** — não há
`Movimentação`/`Posição` nem preço médio a reconstruir. O valor de Bens e Direitos vem do **informe da
corretora estrangeira**, transcrito pelo [/read](../read/SKILL.md).

> **Este skill não roda script nem gera `.xlsx`.** Os ativos no exterior são só itens do `informes.json`
> (`b3:false`, `localizacao` ≠ 105) e entram na declaração final pelo [/consolidate](../consolidate/SKILL.md),
> direto do `informes.json`. O único entregável é o `irpf_consolidated.xlsx`.

## Como o exterior flui (sem etapa própria)
```
/read ──► processed/informes.json ──► /consolidate ──► irpf_consolidated.xlsx
 (LLM lê o informe estrangeiro)   (b3:false, loc≠105)   (Bens e Direitos)
```
- **Bens e Direitos**: cada ação/ETF no exterior é um item `bens` `b3:false` com a sua `localizacao`
  (código do país, ex.: 249 EUA) e grupo/código do informe (3/1, 7/99 ou 31/…, confirmar) — **valor do
  informe** (conversão moeda→BRL pela PTAX de compra de 31/12).
- **Rendimento no exterior** (dividendos/ganhos) é **TRIBUTÁVEL** — **não** vai nas 3 fichas (isento/
  exclusiva); trate na ficha correta (Tributáveis Recebidos do Exterior / carnê-leão / GCAP). O
  [/completeness](../completeness/SKILL.md) **sinaliza** se um rendimento cujo `key` é um bem do exterior
  cair em isento/exclusiva ("Exterior — rendimento em ficha errada?").

## Importante
- **Não é orientação fiscal.** O custo vem do informe estrangeiro; offshore Lei 14.754/2023 tem regra
  própria — confirme com contador.
- **Lembre o /read** de transcrever as posições no exterior que o informe trouxer (furo comum: o
  informe da corretora as lista e a transcrição esquece): elas precisam estar no `informes.json`
  para o /consolidate declará-las.
- Mantenha **genérico** — nunca comite dados de um contribuinte.
