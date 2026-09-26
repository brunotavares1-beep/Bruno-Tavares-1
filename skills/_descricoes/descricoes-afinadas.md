# Descrições afinadas — 5 skills de folha e ponto

Objetivo: acabar com o acionamento errado sem excluir nenhuma skill. A regra
aplicada nas cinco é a mesma: **a primeira frase declara qual arquivo entra**
(é o que distingue as skills entre si), depois vêm os gatilhos, e a última
frase é a fronteira negativa apontando para a skill correta.

Substituir apenas o campo `description` no frontmatter de cada SKILL.md, em
claude.ai. Nada mais muda.

---

## 1. consolidar-holerites-aw

Discriminador: **holerite individual, um por funcionário, em PDF.**

```yaml
description: >-
  ENTRADA: holerites / Recibos de Pagamento INDIVIDUAIS em PDF do TOTVS RM —
  um documento por funcionário por competência, tipicamente em pastas ou ZIPs
  mensais. SAÍDA: uma única planilha Excel consolidada (abas Resumo, Rubricas,
  Dicionario_Rubricas e Erros), que é a base de entrada da skill
  analista-folha-pagamento-senior. Acione ao anexar pasta, ZIP ou vários PDFs
  de holerite / recibo de pagamento / contracheque e pedir para extrair,
  consolidar, "jogar numa planilha só", montar a base de folha ou juntar vários
  meses — mesmo sem dizer "holerite". Diferencia automaticamente um Recibo de
  Pagamento de outro arquivo na mesma pasta (Comunicado de Consignado, Espelho
  de Ponto) e nunca inventa dado: o que não for lido com segurança vai para a
  aba Erros. NÃO use para o relatório "Resumo da Folha" do RM, que traz totais
  da empresa e não linhas por funcionário — esse é resumo-folha-pagamento. NÃO
  use para auditar uma base já consolidada — essa é analista-folha-pagamento-senior.
```

---

## 2. resumo-folha-pagamento

Discriminador: **relatório de totais da empresa, não documento por funcionário.**

```yaml
description: >-
  ENTRADA: o relatório "RESUMO DA FOLHA" ou "Resumo Mensal" do TOTVS RM em PDF —
  um documento por competência com os TOTAIS da empresa (vencimentos, descontos,
  informativa, encargos patronais, líquido), NÃO um documento por funcionário.
  Aceita 1 PDF ou lote (vários PDFs, pastas ou ZIPs num único Excel). SAÍDA:
  Excel de 6 abas no layout executivo aprovado — Resumo com CUSTO TOTAL DA FOLHA
  e headcount por competência, mais Consolidado, Proventos, Descontos,
  Informativa e Encargos, com conferência automática contra os totais impressos
  no PDF (diferença tem de ser 0,00). Acione ao anexar esse relatório ou pedir
  custo total da folha, planilhar a folha, consolidar ou comparar competências
  ("quanto custou a folha", "monta a base de jan a jun"). Nunca inventa dado.
  NÃO use para holerites individuais — essa é consolidar-holerites-aw. NÃO use
  para auditoria de rubricas e custo por função — essa é
  analista-folha-pagamento-senior.
```

---

## 3. analista-folha-pagamento-senior

Discriminador: **entra Excel já consolidado, não PDF.**

```yaml
description: >-
  ENTRADA: a planilha Excel JÁ CONSOLIDADA de holerites (abas Resumo e Rubricas,
  formato de consolidar-holerites-aw) — não PDF. Auditor sênior de folha de
  empreiteira de construção civil, 100% CLT, desonerada (CPRB) com reoneração da
  Lei 14.973/2024. SAÍDA: Excel de auditoria com Resumo Executivo,
  Rubricas Atípicas, Média Salarial por Função, Custo Empresa e Premissas. Aponta
  rubrica fora do padrão histórico (produção, horas extras, adicional noturno,
  insalubridade, periculosidade), compara salário por função com o mercado, estima
  o custo total empresa (FGTS + CPRB/CPP + RAT x FAP + Terceiros) e propõe ação
  para desvio recorrente (dimensionamento de equipe, empreitada x CLT). Acione ao
  pedir para auditar, revisar ou analisar folha, rubricas, horas extras, produção,
  custo de mão de obra, encargos ou média salarial por função, e em "o que subiu
  na folha" ou "vale terceirizar essa equipe". PDF não consolidado: rode
  consolidar-holerites-aw antes. Relatório "Resumo da Folha" do RM:
  resumo-folha-pagamento.
```

---

## 4. matrizes-ponto-catraca

Discriminador: **UMA fonte — só o log da catraca.**

```yaml
description: >-
  ENTRADA: UMA fonte apenas — o log bruto da catraca / controle de acesso (uma
  linha por batida, coluna EVENTO = Entrada/Saida). Não exige lista de ativos nem
  apontamento dos coordenadores. SAÍDA: Excel executivo com quatro matrizes e um
  ranking — PRESENÇA funcionário x dia, ENTRADAS E SAÍDAS com horários, HORAS de
  permanência, HORAS EXCEDENTES sobre a jornada de 44h, e MAIORES DESVIOS POR
  UNIDADE DE NEGÓCIO (NOME FORNECEDOR). Acione quando o usuário jogar SÓ a
  exportação da catraca e pedir matriz de presença, entradas e saídas, horas de
  permanência, horas excedentes ou maiores desvios por unidade / fornecedor /
  obra — inclusive sem dizer "matriz" ("monta a presença da catraca", "quem ficou
  mais horas", "hora excedente por unidade"). Dia com entrada sem saída vira
  PENDÊNCIA: nunca inventa saída. PT-BR, formato BR, rotula confirmado x premissa
  x pendência. Com lista de ativos E apontamento das equipes, use
  conciliacao-ponto-aw.
```

---

## 5. conciliacao-ponto-aw

Discriminador: **TRÊS fontes — catraca + ativos + apontamento.**

```yaml
description: >-
  ENTRADA: TRÊS fontes ao mesmo tempo — CATRACA (log de acessos, EVENTO =
  Entrada/Saida, idealmente com MATRÍCULA/ID), lista de ATIVOS (ID, Colaborador,
  Unidade de Negócio) e APONTAMENTOS dos coordenadores (um arquivo por equipe, uma
  aba por colaborador). É o fechamento semanal de ponto da AW. SAÍDA: Excel com
  matriz de presença, horas extras e domingos por unidade de negócio, o CONFRONTO
  do apontamento contra a catraca (presente sem batida, batida sem apontamento,
  extra, domingo) e a CHECAGEM de cobertura (todo ativo aparece na catraca e no
  apontamento da sua equipe). Acione quando o usuário trouxer essas três fontes e
  pedir conciliação, fechamento semanal de ponto, matriz de presença, horas
  extras, domingos, "conferir o apontamento contra a catraca" ou "ver se todo
  ativo está batendo ponto". Vínculo por matrícula/ID, senão por nome. Furo de
  saída vira pendência: nunca inventa saída. Só a catraca, sem ativos nem
  apontamento: use matrizes-ponto-catraca.
```

---

## Tabela de decisão resultante

| O que o usuário anexa | Skill |
|---|---|
| PDFs de holerite, um por funcionário | `consolidar-holerites-aw` |
| PDF "Resumo da Folha" do RM, totais da empresa | `resumo-folha-pagamento` |
| Excel já consolidado (abas Resumo + Rubricas) | `analista-folha-pagamento-senior` |
| Apontamentos das unidades, cada um com layout próprio | `extrair-apontamento-semanal-aw` |
| Só o log da catraca | `matrizes-ponto-catraca` |
| Catraca + ativos + apontamento | `conciliacao-ponto-aw` |

## Por que isso funciona melhor que excluir

O acionamento errado vinha de todas as descrições começarem pela função
("extrai", "audita", "consolida") e só mencionarem o arquivo de entrada no meio
do texto. Como a função é parecida nas cinco, o gatilho ficava ambíguo. Passando
o arquivo de entrada para a primeira frase, o discriminador aparece antes de
qualquer coisa — e a fronteira negativa no fim redireciona o caso limítrofe para
a skill certa em vez de deixar a escolha no acaso.
