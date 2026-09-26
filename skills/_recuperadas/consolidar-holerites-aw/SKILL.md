---
name: consolidar-holerites-aw
description: >-
  Extrai dados de Recibos de Pagamento (holerites) em PDF gerados pelo TOTVS RM,
  mesmo em grande volume e distribuídos em várias pastas/ZIPs mensais, e
  consolida tudo em uma única planilha Excel (Resumo por funcionário/competência,
  Rubricas detalhadas, Dicionário de Rubricas e Erros). Acione sempre que o
  usuário anexar uma pasta zipada ou várias pastas mensais de holerites ou
  recibos de pagamento da AW Empreiteira e pedir para extrair, consolidar,
  jogar numa planilha só, montar a base de folha, ou comparar/juntar holerites
  de vários meses, mesmo sem citar a palavra holerite explicitamente (por
  exemplo recibo de pagamento, folha mensal, contracheque). A skill diferencia
  automaticamente um Recibo de Pagamento de qualquer outro arquivo incluído por
  engano na mesma pasta, como Comunicado de Empréstimo Consignado/Crédito do
  Trabalhador ou Cartão/Espelho de Ponto, e nunca inventa dado: o que não é
  lido com segurança vai para a aba Erros.
---

# Consolidar Holerites AW — PDF para Excel

## O que esta skill faz

Recebe uma ou mais pastas/ZIPs (tipicamente 1 por competência/mês) contendo
Recibos de Pagamento em PDF no layout padrão TOTVS RM da AW Empreiteira, e
gera **uma única planilha Excel** com:

- **Resumo** — 1 linha por funcionário/competência (dados cadastrais + totais do recibo).
- **Rubricas** — 1 linha por lançamento (código, descrição, referência, vencimento, desconto, natureza).
- **Dicionario_Rubricas** — catálogo dos códigos de rubrica observados (descrição mais frequente, natureza, nº de ocorrências), para conferência/padronização.
- **Erros** — arquivos que não puderam ser lidos, ou que não são Recibo de Pagamento (ex.: Comunicado de Empréstimo Consignado, Cartão/Espelho de Ponto). Por padrão, **não entram** no Resumo/Rubricas.
- **Leia-me** — explicação das abas e estatísticas gerais.

O motor de extração está em `scripts/extrair_holerites.py`. Ele já resolve os
problemas típicos deste layout: cada recibo vem em 2 vias impressas na mesma
página (a skill usa só a 1ª via), com texto de "Declaro ter recebido..." e
"Assinatura do Funcionário" impresso girado 90° sobre as colunas (a skill
ignora texto rotacionado), e os valores de Referência/Vencimentos/Descontos
são classificados pela posição horizontal (x0) de cada número em relação ao
cabeçalho da tabela — não pela ordem em que o texto aparece — o que evita
trocar coluna quando um valor "pula" de linha na extração bruta.

## Como executar

1. Reúna os inputs do usuário: pode ser 1 ZIP com várias subpastas mensais
   dentro, ou vários ZIPs (1 por mês), ou pastas já descompactadas — a skill
   aceita qualquer combinação, inclusive ZIP dentro de ZIP.
2. Rode o script, um `--input` por pasta/zip anexado:

```bash
python scripts/extrair_holerites.py \
  --input "/mnt/user-data/uploads/Janeiro_2026.zip" \
  --input "/mnt/user-data/uploads/Fevereiro_2026.zip" \
  --input "/mnt/user-data/uploads/Marco_2026.zip" \
  ... (um --input por competência) \
  --output "/mnt/user-data/outputs/base_holerites_AW.xlsx"
```

3. Leia a saída do script no terminal: ele imprime quantos PDFs foram
   encontrados, quantos recibos foram processados com sucesso, e a lista de
   arquivos que caíram em Erros (com o motivo). **Sempre confira essa lista
   com o usuário antes de considerar a base pronta** — principalmente se o
   número de arquivos em Erros for maior do que o esperado (pode indicar um
   layout de holerite diferente, não coberto ainda).
4. Apresente o arquivo `.xlsx` gerado ao usuário com `present_files`.

## Classificação: o que é (e o que não é) um holerite

Um arquivo é tratado como Recibo de Pagamento apenas se o texto contiver
simultaneamente "Total de Vencimentos", "Total de Descontos", "Valor Líquido"
e "CNPJ". Isso é deliberadamente restritivo — é melhor um arquivo ir parar em
Erros para revisão manual do que ser processado errado e sujar a base.

Arquivos conhecidos que devem ser excluídos automaticamente (e já têm
mensagem própria no script):
- **Comunicado de Empréstimo Consignado / Crédito do Trabalhador** (carta do
  MTE/DET sobre desconto de empréstimo consignado — não é holerite).
- **Cartão/Espelho de Ponto**.

Qualquer outro tipo de arquivo cai na mensagem genérica de Erros. Se o
usuário mostrar um novo tipo de arquivo indevido que a skill não reconheceu
com uma mensagem específica, adicione uma nova regra em `classificar_arquivo()`
no script — não tente resolver isso "no olho" analisando caso a caso.

## Se a extração falhar em algum recibo

O script já é resiliente a variações pequenas (nome com "via 1"/"via 2"
duplicadas, competências diferentes, funcionários com ou sem rubricas de
empréstimo consignado etc.), porque ele não fixa a ordem/posição das
rubricas — só usa o cabeçalho fixo da tabela e a posição das colunas. Se um
recibo cair em Erros com uma mensagem tipo "Não localizei...", normalmente é
porque:

- O PDF é uma imagem escaneada (sem texto real) — nesse caso é necessário
  OCR antes; use pytesseract conforme a skill `pdf` do sistema, ou avise o
  usuário que aquele PDF específico precisa ser reenviado com texto
  pesquisável.
- O layout daquela competência mudou algum rótulo (ex.: "Sal. Contr. INSS"
  virou outro texto) — abra o PDF com `pdftotext -layout` para comparar e
  ajustar o regex correspondente no script.

Nunca preencha esses campos "chutando" um valor razoável — deixe o recibo em
Erros e avise o usuário.

## Formato de saída

A estrutura de colunas de cada aba é fixa (ver `RESUMO_COLS`, `RUBRICAS_COLS`,
`DICIONARIO_COLS`, `ERROS_COLS` no topo do script) e reflete o modelo já
validado pelo usuário. Não renomeie nem reordene colunas sem confirmar antes
— outras planilhas do usuário (tabela dinâmica, Power Query) podem depender
desses nomes exatos.
