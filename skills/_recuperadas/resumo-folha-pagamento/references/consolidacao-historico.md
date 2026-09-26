# Lote / histórico de competências

## Como o script trata o lote
- Entradas: PDFs, pastas (recursivo) e ZIPs, misturáveis numa mesma chamada.
- PDF que não tem Competência + Total de PROVENTOS é ignorado e listado
  (ex.: holerite, cartão de ponto enviado por engano).
- Ordena por (ano, mês) — "12/2025" vem antes de "01/2026".
- Bloqueia: empresas diferentes no mesmo lote; competência repetida.
- Remonta o Excel do zero a cada execução (sem append).

## Onde cada coisa fica
- Resumo: uma coluna por competência, na mesma ordem das linhas do
  Consolidado (B2 = Consolidado!A2, C2 = Consolidado!A3...). Colunas da aba
  Encargos seguem a mesma ordem — as referências da seção 4 são diretas
  por coluna, então não reordene manualmente as competências numa aba só.
- Proventos/Descontos/Informativa: formato longo, coluna A = competência.

## Leitura da variação
- Decomponha alta/queda em efeito headcount x efeito valor médio.
- Rescisão, Provisão/Estorno e Diferença/Ajuste são voláteis e pouco operacionais.
- Entre anos, a CPP da reoneração sobe (5%→10%→15%→20%): parte da alta do
  custo é regime, não operação — separe isso na análise.
