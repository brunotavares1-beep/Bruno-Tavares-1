# Registro da sessão — varredura e limpeza do ambiente

Data: 26/09/2026 · Repositório: `brunotavares1-beep/Bruno-Tavares-1` · Branch `claude/skill-efficiency-jl6wi0` · PR #1

Este documento substitui o histórico do chat. Reflete o **estado final**, não as conclusões intermediárias — inclusive as que eu corrigi no meio do caminho.

---

## 1. Decisões executadas

| # | Decisão | Estado |
|---|---|---|
| 1 | Perfil pessoal reescrito | **Pronto** — `_descricoes/PERFIL.txt`, pendente colar em claude.ai |
| 2 | 14 unidades de negócio cadastradas no perfil, com regra de rateio do CSC | Pronto |
| 3 | Lucro Real e Excel 365 como blocos de regra no perfil | Pronto |
| 4 | ERP deixa de ser o eixo do perfil | Pronto |
| 5 | Skills de folha excluídas do ambiente e do repositório | **Feito** — serão refeitas |
| 6 | Skills de ponto/catraca a excluir | Pendente em claude.ai (4 skills) |
| 7 | Lixeira de skills e arquivos temporários limpos | Feito |

## 2. O que mudou no perfil

Cortadas ~55 linhas de "Áreas de atuação", que eram enumeração de escopo sem efeito no comportamento. Em lugar delas entraram:

- Contexto: regime tributário (Lucro Real), Excel (365 no Windows), 14 unidades, 2 praças
- Sistema de rótulos: confirmado · premissa · pendência · inconsistência
- Regra de dado faltante: pedir tudo de uma vez, em lista numerada
- Proporcionalidade: estrutura de 4 blocos só para análise e decisão
- Calibragem de crítica: discordar na primeira linha, não concordar por educação
- Bloco de Lucro Real: PIS/COFINS não cumulativo, margem líquida de tributos, compensação de prejuízo limitada a 30%, retenções como antecipação de caixa
- Bloco de Excel 365: funções modernas liberadas, Power Query acima de ~50 mil linhas, proibição de funções voláteis em base grande
- Regra de dados sensíveis (CPF, salário, nome)
- Roteamento de skills — atualizado para não apontar para skill inexistente

Resultado: de 710 para ~1.000 palavras, com **zero** linha de enumeração de escopo. De 9 para 43 regras executáveis.

## 3. Achado pendente — o único que custa dinheiro

### `orcamentista-master-civil` precifica com a regra revogada da desoneração

`references/encargos-sociais-clt.md` trata a desoneração como binária: "a empresa troca os 20% de INSS patronal pela CPRB", com encargo de ~100,4%. Isso **deixou de valer em 2025**: a Lei 14.973/2024 instituiu reoneração gradual e a empresa desonerada paga os dois ao mesmo tempo — CPRB reduzida sobre a receita bruta E CPP crescente sobre a folha (10% em 2026, 15% em 2027, 20% em 2028).

A skill não menciona a Lei 14.973/2024 em nenhum dos seus 8 arquivos. Como ela monta o **preço de venda**, o encargo de mão de obra entra subestimado.

| MDO nominal na obra | Encargo a ~100,4% | Encargo a ~110% | Custo não coberto |
|---|---|---|---|
| R$ 200.000,00 | R$ 200.800,00 | R$ 220.000,00 | R$ 19.200,00 |
| R$ 500.000,00 | R$ 502.000,00 | R$ 550.000,00 | R$ 48.000,00 |
| R$ 1.000.000,00 | R$ 1.004.000,00 | R$ 1.100.000,00 | R$ 96.000,00 |

**Os ~110% são estimativa de ordem de grandeza** — somei os 10 pontos da CPP ao percentual do CBIC. O valor real é provavelmente maior, porque o Grupo D da metodologia SINAPI/TCPO faz o Grupo A reincidir sobre os Grupos B e C, e a CPP entra no Grupo A. Fechar o percentual exige refazer a composição CBIC. Trate como premissa e confirme com o contador.

Patch pronto em `orcamentista-encargos-reoneracao.md`, neste mesmo diretório.

**Agravante:** a `analista-folha-pagamento-senior`, excluída nesta sessão, era a única skill que conhecia o cronograma da reoneração. O ambiente ficou sem nenhuma que saiba da Lei 14.973/2024, e com uma que precifica pela regra antiga.

## 4. Onde eu errei nesta sessão

Registro para não se repetir:

1. **Chamei as 3 skills de folha de "duplicadas" na primeira varredura.** Estava errado: eram um pipeline com três entradas distintas (holerite individual em PDF, relatório Resumo da Folha do RM, planilha já consolidada). O que eu vi foi sobreposição de **gatilho**, não de função. Corrigi antes da exclusão, mas o diagnóstico inicial estava ruim.
2. **Defendi default de 3% para o RAT.** Você endureceu para "limpe RAT de tudo" e estava certo: rótulo de premissa não impede o número de ser usado numa decisão de preço. A correção certa era tornar o parâmetro obrigatório, não escolher um valor melhor.

## 5. Estado final do ambiente

| Item | Antes | Depois |
|---|---|---|
| Skills instaladas | 45 | 42 (menos as 3 de folha) |
| Skills no repositório | 0 | 0 (as refatorações foram descartadas) |
| Repositório: pasta `skills/` | — | 28 KB: só este registro, o patch e o perfil |
| Lixeira de skills | 176 KB | limpa |
| Referências mortas no perfil | 3 | 0 |

### Ainda a excluir em claude.ai (ponto/catraca)

```
conciliacao-ponto-aw
matrizes-ponto-catraca
validar-ponto-catraca
extrair-apontamento-semanal-aw
```

### Conectores sem uso aparente

Adobe Experience Manager (sem autorização), Canva, Dropbox, Monday, Intercom, Linear, Cloudflare, AccuWeather, Google Drive, Microsoft Learn, Notion, Slack. Cada um consome orçamento de ferramentas e atrasa o início da sessão. Desconecte os que não usa.

## 6. O que persiste e o que é efêmero

**Persiste:** o repositório no GitHub, as skills e o perfil em claude.ai.

**Efêmero:** este container. Ele é reciclado por inatividade, com tudo que não estiver commitado. Por isso a limpeza de disco aqui é cosmética — o que importa é o que está versionado.

## 7. Material recuperável do histórico do git

As skills de folha excluídas, com os scripts de extração de PDF do RM, continuam no commit `8a69df3`:

```bash
git show 8a69df3 --stat
git checkout 8a69df3 -- skills/_recuperadas/consolidar-holerites-aw
```

Vale carregar para as novas skills de folha:

- `extrair_holerites.py` e `extrair_resumo_folha.py` — já resolviam o layout de PDF do RM
- O padrão de estrutura: `SKILL.md` curto + `references/` sob demanda + `scripts/` executável, **nunca** script transcrito dentro do SKILL.md
- Descrição começando pelo arquivo de entrada, que é o que distingue uma skill de folha da outra
