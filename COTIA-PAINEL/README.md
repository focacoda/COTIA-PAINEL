# Cotia 2026 — Painel Eleitoral

Painel offline de comparação de votos entre dois candidatos nas seções eleitorais de Cotia (zonas 227 e 286).

## Conteúdo

| Arquivo | Descrição |
|---|---|
| `index.html` | Painel interativo. HTML/CSS/JS num único arquivo, dados embutidos, sem dependências externas. ~1,79 MB. |
| `dados/cotia_dados.json` | Base completa: 558 seções, 81 escolas, 1.812 candidatos, 93.657 registros de voto. |
| `dados/cruzamento_equipe_votos.csv` | Cruzamento seção × votos (formato simples). |
| `dados/cruzamento_equipe_votos_detalhado.csv` | Mesmo cruzamento, com o nome da equipe em cada seção. |

O `index.html` é autossuficiente: os dados embutidos são byte a byte idênticos a `dados/cotia_dados.json`. Funciona aberto direto do disco, sem servidor. Pode ser hospedado em qualquer static host renomeando para `index.html`.

## Uso

Abra `index.html` no navegador. Selecione o candidato A e o candidato B, e filtre por zona, escola, bairro ou seção. Os resultados podem ser agrupados por escola, bairro ou seção e exportados em CSV.

Instalável como app no celular (PWA). O painel detecta quando aberto pelo navegador interno do WhatsApp e mostra instruções para abrir no Safari.

## Números de referência

Votação real em Cotia, todas as 558 seções:

| Candidato | Cargo | Votos |
|---|---|---|
| Átila Jacomussi | Deputado Estadual | 1.874 |
| Alexandre Pereira | Deputado Federal | 1.826 |

Total de votos válidos no município: 149.926 por cargo.

## Qualidade dos dados — leia antes de citar os números da equipe

A planilha de declaração das equipes (`ZONA_SECAO_EQUIPE.xlsx`) contém **278 linhas**, mas apenas **234 atribuições únicas** de equipe × seção (44 linhas são repetições). Essas 234 atribuições cobrem **158 seções distintas**, das quais **41 são compartilhadas** por mais de uma equipe e **117 são exclusivas**.

### 11 seções declaradas não existem na base do TSE

| Seção | Equipe | Problema |
|---|---|---|
| Z227 / S113 | TIME MQ BARRIGA | ausente na base |
| Z227 / S253 | TIME GU KIKO | ausente na base |
| Z227 / S352 | TIME GU NALDO | ausente na base |
| Z227 / S359 | TIME GU RODRIGO | ausente na base |
| Z227 / S438 | TIME DN MOISES | ausente na base |
| Z227 / S485 | TIME GU RODRIGO | ausente na base |
| Z227 / S743 | TIME DN MOISES | **acima da seção máxima da zona (682)** |
| Z227 / S832 | TIME GU BIRO | **acima da seção máxima da zona (682)** |
| Z286 / S237 | TIME RON LANGO | ausente na base |
| Z315 / S304 | TIME GU NALDO | **zona 315 não pertence a Cotia** |
| Z388 / S294 | TIME DN MOISES | **zona 388 não pertence a Cotia** |

Todas as 11 são seções exclusivas, ou seja, o erro é de digitação do número da seção, não de sobreposição entre equipes.

**Consequência:** das 158 seções declaradas, **147 são válidas** e **106 são exclusivas válidas** (não 117). Nove equipes têm seções exclusivas superestimadas:

| Equipe | Exclusivas declaradas | Exclusivas válidas |
|---|---|---|
| TIME DN MOISES | 12 | 9 |
| TIME GU NALDO | 11 | 9 |
| TIME GU RODRIGO | 12 | 10 |
| TIME MQ BARRIGA | 7 | 6 |
| TIME GU KIKO | 5 | 4 |
| TIME GU BIRO | 5 | 4 |
| TIME RON LANGO | 16 | 15 |

As demais (MQ FELIPE, GU ADALTO_FELIPE, FB MARIVALDO, GU AUGUSTO, FB DANILO PÃO, RON MARCIO) não são afetadas. O ranking por votos totais **não muda**, porque seções inexistentes não têm votos.

## Pendências conhecidas no painel

- O botão **Limpar** do filtro de bairros não filtra nada: zera o conjunto e o painel volta a exibir Cotia inteira, em vez de nenhum bairro.
- **Exportar CSV** trata apenas o agrupamento por seção; escolher "por bairro" gera o cabeçalho de escola.
- O rótulo de cargos do KPI é fixo no HTML (`812 fed · 1000 est`) e não acompanha a base.

Uma versão de apresentação (slides) chegou a ser gerada com placeholders não renderizados na tela; ela não faz parte desta pasta.
