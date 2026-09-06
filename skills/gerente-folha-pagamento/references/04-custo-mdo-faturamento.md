# Custo de MDO, fator de encargos e faturamento mínimo

Objetivo: transformar folha em decisão — quanto a obra precisa faturar para se pagar, e qual o peso real da MDO no preço.

## 1. Custo-empresa (nunca confundir com salário nominal)
Componha, nesta ordem:
1. Salário-base (ou o valor efetivamente pago em tarefa/produção — ver ref. 05).
2. Encargos: INSS patronal (ou CPRB), RAT × FAP, terceiros/Sistema S, FGTS 8%.
3. Provisões: 13º (1/12 + encargos), férias + 1/3 (1/12 + encargos), multa de 40% estimada pelo **turnover histórico real**.
4. Benefícios: VT líquido do desconto de 6%, VR/VA, plano de saúde/odonto (usar sinistralidade real), seguro de vida (muitas CCTs obrigam), EPI e uniforme.
5. Indiretos de canteiro atribuíveis: alojamento, transporte fretado, alimentação de canteiro, ferramenta de consumo.

**Fenc = Custo total ÷ Salário-base nominal.**
Faixa de referência de mercado na construção civil: 1,65 a 1,90 para CLT tradicional, podendo passar de 2,0 com turnover alto, insalubridade/periculosidade e canteiro com alojamento. **Referência de plausibilidade, não resultado.** Feche o Fenc definitivo só com CCT vigente, FAP publicado, turnover histórico e sinistralidade reais — peça esses quatro dados juntos.

## 2. Faturamento mínimo para cobrir a folha

```
Faturamento mínimo = Custo total de MDO ÷ (1 − %custos diretos não-MDO − %despesas indiretas − %tributos sobre receita − margem mínima)
```

- Todo percentual do denominador vem do **realizado da obra/empresa** (cruzar com `controller-empreiteira` e `orcamentista-master-civil`). Percentual genérico de mercado só entra rotulado como premissa.
- Se o denominador for ≤ 0, o modelo não fecha: a estrutura de custo já consome a receita — reporte isso como conclusão, não force o cálculo.
- Atalho por indicador, quando a empresa monitora "% MDO sobre faturamento": `Faturamento mínimo ≈ Custo de folha ÷ %MDO-alvo`. Faixa de referência ampla de 25% a 40% conforme o tipo de obra — rotular como referência a confirmar.

Entregue em tabela:

| Obra / BU | Custo de folha (R$) | %MDO-alvo | Faturamento necessário (R$) | Faturamento realizado (R$) | Gap (R$) | Gap (%) | Rótulo |

Rode sensibilidade em 3 cenários (pessimista / provável / otimista) variando Fenc, turnover e %MDO-alvo.

## 3. Auditoria de custo de folha × obra
1. Custo por obra/centro de custo bate com o rateio de MDO direta do orçamento aprovado (TCOP)?
2. Headcount físico ativo reconciliado com o previsto por etapa/fase?
3. Custo por m² executado dentro da faixa orçada? Desvio, investigar por etapa e insumo.
4. Horas extras recorrentes = subdimensionamento crônico ou pico sazonal? A resposta muda a decisão (contratar × pagar extra).
5. Turnover do período elevou rescisão/aviso acima do provisionado?
6. Custo de MDO a incorrer até o fim da obra × saldo orçamentário disponível.

## 4. Decisão CLT × tarefa × PJ/subempreiteiro
Compare **custo-hora efetivo**, não valor de face:
- CLT: salário ÷ horas × Fenc.
- Tarefa CLT: valor pago em tarefa ÷ horas efetivas × Fenc aplicável ao regime.
- PJ/MEI/subempreiteiro: valor da nota + encargos remanescentes (retenção de 11% em cessão de mão de obra quando aplicável — ver ref. 06) + risco de reconhecimento de vínculo, que é passivo, não custo.
Nunca apresente PJ como "mais barato" sem a linha de risco na tabela.
