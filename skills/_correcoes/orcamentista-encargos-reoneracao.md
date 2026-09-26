# Correção — `orcamentista-master-civil` desconhece a reoneração da folha

## O defeito

O arquivo `references/encargos-sociais-clt.md` trata a desoneração como regra
binária, com este texto:

> `Desonerado` = a empresa troca os 20% de INSS patronal do Grupo A pela CPRB
> (Contribuição Previdenciária sobre a Receita Bruta)

| Regime | Encargo sobre o salário nominal |
|---|---|
| Sem desoneração da folha | ~130,6% |
| Com desoneração (CPRB, Lei 12.546/2011) | ~100,4% |

Esse "ou um, ou outro" **deixou de valer em 2025**. A Lei 14.973/2024 instituiu
reoneração gradual: a empresa desonerada passou a pagar **os dois ao mesmo
tempo** — CPRB reduzida sobre a receita bruta E uma CPP crescente sobre a folha.

| Ano | CPRB aplicável (da alíquota cheia de 4,5%) | CPP sobre a folha |
|---|---|---|
| 2025 | 80% | fontes secundárias divergem entre 5% e 10% — confirmar na lei |
| 2026 | 60% | **10%** |
| 2027 | 40% | 15% |
| 2028 | 0% (CPRB extinta) | 20% (INSS patronal integral) |

A skill não menciona a Lei 14.973/2024 em nenhum dos seus 8 arquivos. Ela
precifica obra com a regra de 2024.

## Por que isso é caro

Essa skill monta o **preço de venda**. Se o encargo de mão de obra entra
subestimado, o preço sai abaixo do custo e a margem só aparece como perda no
fechamento da obra.

Ordem de grandeza do erro em 2026: a CPP de 10% sobre a folha **não está** nos
~100,4%. O encargo real do regime desonerado é hoje da ordem de **~110%**, não
~100,4%.

| MDO nominal na obra | Encargo a ~100,4% | Encargo a ~110% | Custo não coberto |
|---|---|---|---|
| R$ 200.000,00 | R$ 200.800,00 | R$ 220.000,00 | R$ 19.200,00 |
| R$ 500.000,00 | R$ 502.000,00 | R$ 550.000,00 | R$ 48.000,00 |
| R$ 1.000.000,00 | R$ 1.004.000,00 | R$ 1.100.000,00 | R$ 96.000,00 |

**Atenção ao número:** os ~110% são estimativa de ordem de grandeza, obtida
somando os 10 pontos da CPP ao percentual do estudo CBIC. O valor exato é
provavelmente **maior**, porque o Grupo D da metodologia SINAPI/TCPO faz o
Grupo A incidir novamente sobre as verbas dos Grupos B e C — e a CPP entra no
Grupo A. Fechar o percentual correto exige refazer a composição CBIC com a CPP
dentro, o que depende do estudo original. Até que isso seja feito, trate como
premissa e confirme com o contador.

## Substituir no arquivo `references/encargos-sociais-clt.md`

Trocar a seção "Percentual de referência" inteira por:

```markdown
## Percentual de referência (CBIC — "Encargos Previdenciários e Trabalhistas no
Setor da Construção Civil"; trate como `premissa`, sempre confirmar se há
versão mais recente)

| Regime | Total de encargos sobre o salário nominal |
|---|---|
| Sem desoneração da folha | **~130,6%** |
| Com desoneração, regra anterior a 2025 | ~100,4% — **NÃO USE MAIS** |
| Com desoneração, em reoneração gradual (2025 a 2028) | **~100,4% + a CPP do ano** — ver tabela abaixo |

**ATENÇÃO — a desoneração não é mais binária.** Até 2024, `desonerado`
significava trocar os 20% de INSS patronal pela CPRB. A **Lei 14.973/2024**
acabou com isso: desde 2025 a empresa desonerada paga os dois ao mesmo tempo —
CPRB reduzida sobre a receita bruta E uma CPP crescente sobre a folha.

| Ano | CPRB aplicável (da alíquota cheia de 4,5% da construção civil) | CPP sobre a folha |
|---|---|---|
| 2025 | 80% → ~3,60% da receita bruta | fontes secundárias divergem (5% ou 10%) — **confirmar na lei** |
| 2026 | 60% → ~2,70% da receita bruta | **10%** |
| 2027 | 40% → ~1,80% da receita bruta | 15% |
| 2028 | 0% — CPRB extinta | 20% (INSS patronal integral) |

Consequências para o orçamento:

1. **O percentual de encargo do regime desonerado sobe todo ano** até 2028,
   quando os dois regimes convergem. Um orçamento de obra que atravessa a
   virada do ano tem dois percentuais diferentes de encargo.
2. **A CPRB não é encargo de folha** — incide sobre a receita bruta. Ela entra
   no preço junto com os demais tributos sobre a venda, não na composição de
   custo de mão de obra. Não a lance nos dois lugares: é dupla contagem.
3. **A CPP entra no Grupo A** e, portanto, reincide sobre os Grupos B e C pelo
   Grupo D. Somar 10 pontos ao percentual antigo é aproximação por baixo; o
   correto é refazer a composição com a CPP dentro.
4. **Nunca valide um % de encargo sem perguntar a competência do orçamento.**
   "Estamos desonerados" já não define o percentual — só o ano define.

Fonte primária: Lei 14.973/2024. Confirme o texto no Planalto antes de fechar
preço. Em dúvida sobre o enquadramento, acione `analista-tributario-senior`.
```

## Ainda no mesmo arquivo

Na linha do Grupo A, onde está:

> INSS patronal (20%), GIIL-RAT/seguro acidente (~3%), FGTS (8%), Sistema S
> (SESI/SENAI/SEBRAE/INCRA, ~5,3%), salário-educação (~2,5%)

Acrescentar depois de "INSS patronal (20%)":

> — ou, no regime desonerado, a CPP do ano conforme o cronograma da Lei
> 14.973/2024, que substitui parcialmente esses 20%

## Verificação depois de aplicar

Pergunte à skill: *"qual o % de encargo de CLT numa obra nossa em 2026?"*

- **Errado:** responde ~100,4% porque estamos desonerados.
- **Certo:** responde que o regime está em reoneração, que ~100,4% é a regra
  anterior a 2025, que falta somar a CPP de 10% de 2026, e que o percentual
  exato precisa da composição refeita.
