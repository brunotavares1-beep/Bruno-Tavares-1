# DP específico de construção civil e empreiteira

O que diferencia a folha de obra da folha de escritório — e onde mora a autuação.

## CNO e matrícula da obra
Obra própria ou por empreitada total exige **CNO (Cadastro Nacional de Obras)**, que substituiu a matrícula CEI. Amarre o CNO ao centro de custo da obra: sem isso, o recolhimento e o eSocial não conversam com o controle gerencial e a baixa da obra trava.

## Retenção de 11% — cessão de mão de obra e empreitada (art. 31, Lei 8.212/91)
Serviço prestado mediante cessão de mão de obra ou empreitada sofre **retenção de 11% sobre o valor bruto dos serviços** na nota, recolhida pelo tomador. Pontos que quebram na prática:
- A retenção é do prestador e compensável — trate como adiantamento de INSS, não como custo.
- Base pode ser reduzida quando material e equipamento estão discriminados em contrato e na nota, dentro dos limites da IN vigente. **Discriminação vaga = retenção sobre o valor cheio.**
- Alíquota sobe para 15%/12%/9% quando há aposentadoria especial na atividade cedida.
- Empreitada total de obra tem tratamento distinto da cessão de mão de obra — classifique o contrato antes de aplicar.
- Confirme sempre a IN vigente antes de fechar o percentual e a base.

## CCT / SINDUSCON
Piso salarial por função, adicional de produtividade, cesta básica, seguro de vida obrigatório, café da manhã, prêmio de assiduidade, regras de banco de horas e de compensação de sábado — tudo varia por sindicato e por base territorial. **Obra em município diferente pode ter CCT diferente.** Nunca aplique a CCT de uma obra em outra sem confirmar a base territorial e a vigência.

## Jornada e ponto em canteiro
Compensação de sábado, banco de horas (exige previsão em CCT ou acordo escrito), turnos, horas in itinere quando previstas em norma coletiva, DSR sobre horas extras, e o efeito de chuva/paralisação. A folha só é confiável se o ponto foi conciliado antes — use `conciliacao-ponto-aw` e `matrizes-ponto-catraca` como etapa anterior ao fechamento.

## Insalubridade e periculosidade
Grau (10%, 20% ou 40% para insalubridade; 30% para periculosidade) definido por laudo técnico vigente, não por costume da obra. Base de cálculo da insalubridade é tema com controvérsia — declare qual base está sendo usada e rotule como interpretação a validar. Periculosidade e insalubridade não se acumulam: o empregado opta.

## Segurança e obrigações acessórias que geram passivo
- SESMT e CIPA dimensionados por grau de risco e número de empregados.
- PCMSO/PGR e exames (admissional, periódico, demissional, de mudança de função) em dia — ASO vencido inviabiliza demissão regular.
- PPP e LTCAT alimentando o eSocial (S-2240) para exposição a agentes nocivos.
- Cota de aprendizes e cota de PCD: base de cálculo e funções que podem ser excluídas são objeto de autuação frequente em construção civil. Levante o cálculo da cota junto com o headcount.
- CAT em até 1 dia útil (imediata em caso de óbito).

## Rotatividade
Turnover alto é a marca do setor e distorce tudo: provisão de rescisão, Fenc, custo por m² e curva de aprendizado da equipe. Sempre calcule a provisão de 40% com o turnover histórico real da empresa, não com média genérica.
