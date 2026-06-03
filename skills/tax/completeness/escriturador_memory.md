# escriturador_memory — ação/FII → escriturador (instituição escrituradora)

O **escriturador** mantém o registro do papel e emite o **informe autoritativo de dividendos / JCP /
rendimentos** — mais confiável que a corretora, que pode enxergar só parte (ex.: BBSE3 mostrou
R$ 454,46 na corretora × **R$ 842,87** no escriturador BB). O `/completeness` usa esta tabela para,
por ação/FII em Bens e Direitos, conferir se **alguma fonte do rendimento é o informe do escriturador**
(o `tag` aparece no `source` do rendimento no `informes.json`).

> **Gere este arquivo automaticamente da B3** (recomendado — sem prints, sem depender dos informes):
> `python scripts/fetch_escriturador.py --investimentos processed/b3_brazil_variable_income_avg_price_calculation.xlsx`
> Ele lê cada ação/FII/BDR do workbook do b3 e consulta a API pública da B3 (campo `ifd` p/ FII/FI-Infra,
> `hasCommom` p/ ação/BDR). O resto deste arquivo é o **formato + fallback manual**.

**Prioridade da fonte do escriturador** (use a mais alta disponível e registre em `source`):
1. **O próprio informe do escriturador** — Bradesco "Ativos Escriturais", BB Seguros, Porto/Itaú etc.
   (é o documento que o escriturador emite; mais autoritativo que qualquer terceiro).
2. **B3** — a página oficial do ativo (abaixo), seção **Escriturador**. As páginas são SPA (não saem
   em texto puro); abra no navegador. É a fonte de cadastro oficial.
3. **Regulamento / CVM-FNET** do fundo (para FII/FI-Infra).
4. **Busca web** (statusinvest, clubefii, material do fundo) — **só como último recurso** e sempre
   marque `source` com **"⚠️ via web — CONFIRMAR na B3"**; snippets erram/desatualizam (ex.: CNPJ divergente).

Preencha a partir da B3 (abra o ativo → seção **Escriturador**):
- **Ações:**         https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/empresas-listadas.htm
- **FIIs:**          https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/fundos-de-investimentos/fii/fiis-listados/
- **FI-Infra:**      https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/fundos-de-investimentos/fi-infra/fi-infra-listados/
- **BDR patroc.:**   https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/bdrs/bdrs-patrocinados/bdrs-patrocinados-listados/
- **BDR não patr.:** https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/bdrs/bdrs-nao-patrocinados/bdrs-nao-patrocinados-listados/

Colunas: `ticker | escriturador | cnpj | tag | source`
- **tag** — o trecho **minúsculo** que aparece no `source` do informe quando o documento DAQUELE
  escriturador foi usado (ex.: `bradesco`, `bb`, `porto`). É o que liga a memória ao informe.
- **source** — o link da B3 (ou outra referência) de onde veio o escriturador.

> ⚠️ **Repo público:** as linhas abaixo são **exemplos** — substitua pelas suas e mantenha o seu
> `escriturador_memory.md` na pasta de trabalho (`memory/`), passando `--escriturador-memory` se preciso.

| ticker | escriturador | cnpj | tag | source |
|---|---|---|---|---|
| EXEMPLO3 | Banco Exemplo S.A. | 00.000.000/0001-00 | exemplo | https://www.b3.com.br/... |
| EXEMP11 | Escriturador Exemplo de FII | 00.000.000/0001-00 | exemplo | https://www.b3.com.br/... |
