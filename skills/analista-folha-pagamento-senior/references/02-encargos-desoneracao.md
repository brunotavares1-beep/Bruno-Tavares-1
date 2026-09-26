# Encargos sobre a folha e desoneração (CPRB), 2025-2028

Base legal e premissas usadas por `scripts/analisar_folha.py`. Revise periodicamente: a reoneração é gradual e o cronograma pode mudar por nova legislação. Confirme a alíquota vigente antes de fechar número para decisão (preço de venda, orçamento, negociação).

## O regime da empreiteira

A construção civil (CNAE 41 a 43) está entre os setores que podem optar pela Contribuição Previdenciária sobre a Receita Bruta (**CPRB**, Lei 12.546/2011), em substituição à CPP (INSS patronal) de 20% sobre a folha. A alíquota cheia de CPRB para construção civil é **4,5% sobre a receita bruta**.

Isso só se aplica à receita da atividade desonerada. Se a empreiteira tiver receita fora do escopo de construção civil (CNAE diferente), essa parcela **não** entra na base da CPRB.

## Cronograma de reoneração gradual

Lei 14.973/2024, publicada em 16/09/2024. A partir de 2025 há retorno gradual ao INSS patronal pleno, em paralelo à redução progressiva da CPRB:

| Ano | CPRB aplicável (% da alíquota cheia de 4,5%) | CPP sobre a folha |
|---|---|---|
| 2025 | 80% → CPRB efetiva ≈ 3,60% da receita bruta | fontes secundárias divergem entre 5% e 10% — **confirme na lei** |
| 2026 | 60% → CPRB efetiva ≈ 2,70% da receita bruta | 10% sobre a folha |
| 2027 | 40% → CPRB efetiva ≈ 1,80% da receita bruta | 15% sobre a folha |
| 2028 | 0% (CPRB extinta) | 20% sobre a folha (INSS patronal integral) |

Durante a transição a empresa paga **os dois ao mesmo tempo**: CPRB sobre a receita bruta E uma CPP básica sobre a folha. Não é mais "ou um, ou outro", como era antes da Lei 14.973/2024. O 13º salário permanece com tratamento apenas de CPRB, sem a CPP adicional, durante a transição.

**Fonte primária: Lei 14.973/2024.** Confirme o texto no Planalto antes de fechar número de alto impacto; fontes secundárias (portais de contabilidade) divergem, principalmente sobre 2025. Em dúvida, acione `analista-tributario-senior` ou o contador.

## Demais encargos — sem isenção para empresa desonerada

A desoneração substitui **apenas** a CPP. Continuam devidos integralmente:

| Encargo | Alíquota | Observação |
|---|---|---|
| FGTS | 8% mensal | + multa rescisória de 40% em dispensa sem justa causa |
| RAT/SAT | 1%, 2% ou 3% por grau de risco do CNAE | **Sem valor assumido.** O script exige `--rat`. Construção civil (CNAE 41 a 43) costuma ser grau de risco 3, mas confirme o CNAE do estabelecimento — supor custa 1 ponto percentual da base do INSS por erro |
| FAP | Multiplicador de 0,5 a 2,0 sobre o RAT | Publicado anualmente em janeiro pelo INSS/MTE, específico por CNPJ. Nunca suponha 1,0 sem confirmar |
| Terceiros / Sistema S | ~5,8% (usual do FPAS 507) | SENAI, SESI, SESC, SENAC, SEBRAE, INCRA, salário-educação. Varia pelo FPAS do estabelecimento |

## O que o script calcula e o que não calcula

**Calcula:** FGTS (valor apurado no holerite quando disponível, senão 8% sobre a base), CPP híbrida do ano, RAT × FAP, Terceiros, e a CPRB em R$ **somente se** a receita bruta for informada.

**Não calcula:** a provisão de rescisão (multa de 40% + aviso prévio + férias e 13º proporcionais), que depende da rotatividade real da empresa. Peça a taxa de rotatividade antes de estimar; não use percentual genérico.

`--rat`, `--fap` e `--terceiros` são **obrigatórios**: o script não tem valor default para nenhum dos três e recusa rodar sem eles. Os três dependem do CNAE, do FAP publicado e do FPAS do estabelecimento — qualquer valor assumido produz custo patronal errado com aparência de certo. A aba `Premissas` registra os três valores informados.

## Reforma Tributária não muda isso

EC 132/2023 + LC 214/2025 incidem sobre o **consumo** (IBS/CBS), não sobre a folha. Há alíquota reduzida para construção civil em obras e incorporações, mas isso é IBS/CBS na venda — não afeta FGTS, CPRB ou CPP sobre a folha.
