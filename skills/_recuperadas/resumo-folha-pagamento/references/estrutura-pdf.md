# Estrutura do PDF "RESUMO DA FOLHA" (TOTVS RM)

## Cabeçalho repetido (uma vez por página — filtrar sempre)
```
Empresa: <código> - <razão social> Página: N/M
CNPJ: <cnpj> Emissão: <data>
Cálculo: <tipo> Hora: <hora>
Competência: MM/AAAA
Complemento de cálculo: Todos
RESUMO DA FOLHA               <- só na página 1, mas pode faltar em algumas exportações
Rubrica Nome da Rubrica Nº Empregados/Contribuintes Valor informado Valor Calculado
Sistema licenciado para <contabilidade>   <- rodapé de toda página
```

## Marcadores de seção
`PROVENTOS`, `DESCONTOS`, `INFORMATIVA` aparecem uma vez cada como linha isolada
abrindo o bloco — **mas o marcador `DESCONTOS` (ou o que estiver ativo) também
pode reaparecer sozinho no fim de uma página**, como uma espécie de rodapé de
continuação, antes do `Sistema licenciado...`. Isso é inofensivo: o parser
apenas reafirma a seção corrente, não reseta nem duplica nada. Não tente
"consumir" esse marcador extra de forma especial.

## Linha de rubrica
```
<código> <nome da rubrica, pode conter números/%/º/Nº> <nº empregados> <valor informado> <valor calculado>[*]
```
- O parser NUNCA tenta separar `nome` do resto pela esquerda (nome pode ter
  números embutidos, ex. `DESC. EMP. CRED. TRAB Nº 2760430138`). Ele ancora
  pela **direita**: os 3 últimos tokens da linha devem casar com
  `inteiro`, `valor monetário`, `valor monetário` — só então a linha é aceita
  como rubrica. Se algum dia o RM mudar o layout (ex. adicionar uma coluna),
  esse casamento vai falhar e a linha cai em Pendências — é o comportamento
  correto, não um bug a "corrigir na hora".
- `*` colado no fim do valor calculado (sem espaço, ex. `908,19*`) marca
  rubrica informativa. No exemplo testado, isso só ocorre dentro do próprio
  bloco INFORMATIVA (redundante com a seção), mas o parser também detecta
  `*` em PROVENTOS/DESCONTOS caso apareça — nesses casos a rubrica é somada
  normalmente no bloco, só é marcada como "Informativa" na coluna de saída.

## Bloco de encargos (INSS/FGTS/PIS/ISS) — após o `Total:` da INFORMATIVA
Vem em **pares por linha**, formato `Label: valor Label: valor`:
```
Salário contribuição empregados: 800.824,13 Base do FGTS: 815.907,65
Empresa: 83.452,84 Valor FGTS Rescisório: 37.066,66
```
Atenção: a palavra "Empresa" aparece aqui como **label do INSS patronal**
(coluna esquerda do bloco INSS), não confundir com o "Empresa: <razão social>"
do cabeçalho de página — o parser só entra nesse modo depois de já ter visto
`Líquido Geral:` ou `Total: Resumo Geral Mensal e Complementar`, então não há
ambiguidade real, mas fique atento se o layout mudar.

## Bloco IRRF
Precedido pela linha `IRRF conforme competência do cálculo IRRF conforme
competência do pagamento`. Cada linha seguinte tem o MESMO label duas vezes
(uma para "cálculo", uma para "pagamento") — o parser usa a **posição** (1º
par = cálculo, 2º par = pagamento), não o texto do label, para separar as
duas colunas.

## Bloco Situações (headcount)
Depois da linha `Situações`. Mesmo formato de pares, mas os valores são
**inteiros simples** (não monetários) — ex. `No. Empregados: 215 Demitido: 7`.

## Linhas que o parser não reconhece (vão para Pendências)
Isso acontece quando:
- o PDF vier de uma versão do RM com colunas extras ou nomes de campo diferentes;
- a rubrica tiver nome vazio ou algum caractere quebrando a extração de texto do PDF (raro, mas pode ocorrer em PDFs escaneados — se for o caso, use OCR antes, via skill `pdf`);
- o bloco de encargos trouxer um label novo que não seja `Label: valor` (ex. texto livre).

Sempre trate pendência como sinal para revisar manualmente — nunca estime o
valor "pelo padrão dos outros meses" para preencher uma pendência.

## Validação cruzada de totais — regra do asterisco
Rubricas com `*` colado ao valor calculado dentro de PROVENTOS/DESCONTOS não
entram no `Total:` impresso. Prova (competência 02/2026, AW Empreiteira):
- Soma de todas as 188 linhas de DESCONTOS: R$ 641.848,23
- Soma sem as linhas `*` (9176 = 8.721,14 e 9177 = 758,36): R$ 632.368,73
- `Total:` impresso: R$ 632.368,73 → bate centavo a centavo; Líquido Geral
  (974.898,34 − 632.368,73 = 342.529,61) também bate.

`843 INSS EMPREGADOR` NÃO é excluída — compõe o total normalmente. (A versão
anterior da skill atribuía a divergência a 843/9752; estava errado.)

Se aparecer divergência mesmo excluindo `*`, é sinal de layout novo ou linha
perdida: vá para Pendências e investigue; nunca "feche a conta" excluindo
rubricas por palpite.

## Armadilha corrigida: "Empresa:" no bloco de encargos
O cabeçalho de página começa com `Empresa: 50 - RAZÃO SOCIAL`, e o bloco INSS
tem a linha `Empresa: 76.553,35 Valor FGTS Rescisório: 15.597,68`. O filtro
de cabeçalho antigo (`startswith("Empresa:")`) descartava essa linha e perdia
o INSS patronal e o FGTS rescisório. Hoje o filtro usa regex estrito
(`Empresa: <código> - <texto>`). Se criar novo filtro, teste contra essa linha.

Linhas de pares `Label: valor` com texto que sobra (não casado) vão inteiras
para Pendências, além de registrar os pares lidos.
