---
name: "resumo-folha-pagamento"
description: Extrai o "RESUMO DA FOLHA" do TOTVS RM (PDF) para Excel no layout executivo aprovado — 1ª aba "Resumo" com vencimentos, descontos, informativa, encargos patronais, líquido, CUSTO TOTAL DA FOLHA e headcount por competência — mais 5 abas de base (Consolidado, Proventos, Descontos, Informativa, Encargos INSS-FGTS-IRRF). Funciona para 1 PDF ou em LOTE (vários PDFs, pastas ou ZIPs num único Excel), garantindo 100% dos registros do PDF com conferência automática contra os totais impressos. Acione SEMPRE que o usuário anexar PDF "Resumo da Folha"/"Resumo Mensal" do RM ou pedir extrair, "plnilhar", custo total da folha, consolidar ou comparar competências — mesmo sem citar "resumo da folha" ("joga essa folha numa planilha", "monta a base de jan a jun", "quanto custou a folha"). Nunca inventa dado. Não usar para holerites individuais (consolidar-holerites-aw) nem auditoria sem esse relatório (analista-folha-pagamento-senior).
---

# Resumo da Folha (TOTVS RM) → Excel de Custo Total

## Entrega padrão (6 abas, nesta ordem)
1. **Resumo** — layout aprovado pelo usuário (arquivo "Exemplo - Extração"). Linhas fixas por categoria, uma coluna por competência + "Total período". Tudo em fórmula viva (SUMIFS/INDEX-MATCH nas abas de base):
   1. Vencimentos · 2. Descontos · 3. Informativa · 4. Encargos patronais · 5. Fechamento (Líquido e **CUSTO TOTAL DA FOLHA**) · 6. Headcount · 7. Conferência com os totais impressos no PDF (diferença tem de ser 0,00).
2. **Consolidado** — 1 linha por competência (A=Competência, B..E totais impressos, F custo total, G/H/I Trabalhando/Admissões/Demitido — **não mover essas colunas**, a aba Resumo depende delas), metadados, contagem de linhas, status, e o bloco **PENDÊNCIAS** no rodapé.
3. **Proventos** · 4. **Descontos** · 5. **Informativa** — todas as rubricas, todas as competências (A Competência, D Categoria, G Valor Calculado, H `*` informativa).
6. **Encargos INSS-FGTS-IRRF** — formato largo (1 coluna por competência): INSS/FGTS/PIS/ISS, IRRF "cálculo", IRRF "pagamento" e Situações (headcount).

Não crie abas extras: o usuário quer a 1ª aba no layout aprovado + no máximo 5.

## Passo 1 — Extrair (sempre via script)
A extração é determinística (regex sobre o texto do PDF). Rode o script mesmo que o texto do PDF esteja visível na conversa — é o que garante que nada é inventado nem esquecido.

```bash
python3 scripts/extrair_resumo_folha.py <pdf|pasta|zip> [mais...] \
  --out /mnt/user-data/outputs/Resumo_Folha_<empresa>_<periodo>.xlsx
```
- Lote: passe vários PDFs, uma pasta ou um ZIP — o script expande tudo, ignora PDFs que não são "Resumo da Folha" (e lista quais), ordena por competência (ano/mês, não alfabético) e monta um único Excel.
- O script **para com erro** se o lote misturar empresas (rode 1× por empresa) ou tiver competência repetida (somaria em dobro). Informe o usuário, não contorne.
- Sempre remonta do zero. Para incluir um mês novo, rode de novo com todos os PDFs.
- Categorias: `assets/categorias_rubricas.csv` (carregado por padrão).

## Passo 2 — Validar antes de entregar (obrigatório)
1. Recalcule: `python3 /mnt/skills/public/xlsx/scripts/recalc.py <xlsx>` → `total_errors` = 0.
2. Leia o terminal: para cada competência os 3 blocos devem estar **OK** e o Líquido Geral **OK**.
3. Abra a aba Resumo (data_only) e confira que a seção 7 tem diferença 0,00 em todas as colunas.
4. **Pendências > 0 ou linha vermelha "NÃO CLASSIFICADO" é bloqueante**: resolva (ver abaixo) ou avise explicitamente o usuário. Nunca entregue calado.

Regra dos totais (confirmada no PDF 02/2026, bate centavo a centavo): rubricas com `*` no valor calculado dentro de PROVENTOS/DESCONTOS (ex.: 9176/9177 DEPENDENTE IRRF) **não compõem o `Total:` impresso** — ficam na aba de base marcadas "SIM" na coluna H e a aba Resumo as exclui dos SUMIFS. O bloco INFORMATIVA soma inteiro.

## Passo 3 — Custo Total da Folha
`CUSTO TOTAL = Total Vencimentos + Total Encargos Empresa`, onde encargos = FGTS mensal + FGTS aprendiz + FGTS rescisório (8% + multa 40%) + FGTS rescisório mês anterior + **INSS patronal empregados ("Empresa")** + **INSS patronal contribuintes/pró-labore** + RAT/FAP + Terceiros — todos lidos do bloco "INSS FGTS, PIS e ISS" do PDF.

- **Reoneração (Lei 14.973/2024):** a empresa (desonerada/CPRB) passou a recolher CPP sobre a folha: 5% em 2025, 10% em 2026, 15% em 2027, 20% em 2028. O valor da linha "Empresa" do PDF é essa CPP e **entra no custo** (em 02/2026: R$ 76.553,35 ≈ 10% da base). A regra antiga "nunca somar INSS patronal" está superada.
- A CPRB (sobre receita) não aparece no Resumo da Folha — não está no custo; se o usuário quiser custo previdenciário total, precisa vir da apuração da CPRB.
- `843 INSS EMPREGADOR` (bloco Descontos) é desconto de 11% sobre o pró-labore do contribuinte, não custo patronal — fica em Descontos.

## Rubrica nova (NÃO CLASSIFICADO)
Aparece em linha vermelha extra na aba Resumo (o valor nunca some do total). Para resolver:
1. Classifique seguindo o padrão do CSV (ex.: "MEDIA ... FERIAS" → Férias; "DESC EMP CRED TRAB" → Empréstimo Consignado) e use apenas categorias que já existem no layout.
2. Adicione a linha em `assets/categorias_rubricas.csv` e rode o script de novo.
3. Liste para o usuário os códigos classificados por inferência.

## Análise (quando pedida, além da extração)
- Compare por categoria, não por código (consignado muda de código por contrato).
- Separe efeito volume (headcount/nº empregados) de efeito valor médio.
- Categorias voláteis (Rescisão, Provisão/Estorno, Diferença/Ajuste) geram ruído não operacional — sinalize.
- Ao comparar custo entre anos, lembre que a alíquota de CPP sobe a cada ano da reoneração.

## Arquivos
- `scripts/extrair_resumo_folha.py` — parser + gerador do Excel (única forma de extrair).
- `assets/categorias_rubricas.csv` — rubrica → categoria (editável).
- `references/estrutura-pdf.md` — layout do PDF e casos de borda de parsing. Leia se houver pendência ou divergência.
- `references/consolidacao-historico.md` — regras do lote/histórico e de leitura de variação.
