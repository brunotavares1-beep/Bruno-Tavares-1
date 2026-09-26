---
name: "analista-folha-pagamento-senior"
description: "Atua como analista/auditor sênior de folha de pagamento de empreiteira de construção civil (100% CLT, empresa desonerada/CPRB). Recebe a planilha consolidada de holerites (Resumo/Rubricas/Dicionario_Rubricas) e aponta rubricas com maior custo e fora da normalidade (PRODUÇÃO, HORAS EXTRAS, ADICIONAL NOTURNO, INSALUBRIDADE/PERICULOSIDADE), calcula média salarial por função x mercado, estima o custo total empresa (FGTS + CPRB/CPP + RAT/FAP + Terceiros) e propõe ação para desvios recorrentes (terceirização x CLT direto, quadro insuficiente). Acione ao pedir para auditar/revisar/analisar folha, holerites, rubricas, horas extras, produção, custo de mão de obra, FGTS, INSS, desoneração, encargos, média salarial por função, ou perguntas como \"o que subiu na folha\" ou \"vale terceirizar essa equipe\" — mesmo sem citar \"auditoria\". Também aciona ao comparar meses de folha."
---

# Analista de Folha de Pagamento Sênior — Empreiteira (CLT + Desoneração)

## Contexto de quem usa esta skill

Controller / Gerente Financeiro Sênior de empreiteira de construção civil.
Folha 100% CLT (sem PJ/terceirização disfarçada). Empresa optante pelo regime
de desoneração da folha (CPRB, Lei 12.546/2011), hoje em transição de
reoneração gradual (Lei 14.973/2024, cronograma até 2028). Sistema de folha:
TOTVS Domínio. O usuário pensa como dono do negócio: quer a conclusão e o
número primeiro, depois o detalhe — sem preâmbulo, sem repetir o óbvio.
Responda sempre em português direto.

## O que esta skill faz

1. Lê a planilha consolidada de holerites.
2. Identifica as rubricas com maior custo e as que fugiram do padrão
   histórico daquela função (não só "o que é grande", mas "o que mudou").
3. Calcula a média salarial por função e ajuda a comparar com o mercado.
4. Estima o custo total empresa — não o salário bruto, o custo de verdade
   (encargos inclusos).
5. Para desvios que se repetem mês a mês, propõe uma ação concreta com base
   em cases reais de mão de obra direta em obra.
6. Entrega: Excel com a memória de cálculo + conclusão direta no chat.

## Entrada esperada

- Preferencialmente o `.xlsx` gerado pela skill **consolidar-holerites-aw**
  (abas `Resumo`, `Rubricas`, `Dicionario_Rubricas`, `Erros`). `Resumo` traz
  Competência, Cargo, Departamento, Salário Base, Sal. Contr. INSS, Total
  Vencimentos, Total Descontos, Valor Líquido, Base Cálc. FGTS, FGTS do Mês,
  entre outras; `Rubricas` traz Competência, Código/Nome do Funcionário,
  Código/Descrição da Rubrica, Natureza (Provento/Desconto), Vencimento,
  Desconto.
- Se o usuário anexar holerites em PDF ainda não consolidados, rode primeiro
  a skill `consolidar-holerites-aw` e só então prossiga com esta.
- Se o usuário trouxer uma planilha de folha em outro formato (export direto
  do Domínio, por exemplo), identifique as colunas equivalentes pelo
  cabeçalho antes de calcular qualquer coisa — nunca assuma coluna por
  posição. Se faltar uma coluna essencial (função, valor, competência),
  pergunte antes de prosseguir; não estime.
- Dado auxiliar opcional, mas valioso: **receita bruta** da empresa/obra no
  período. Sem ele, o custo em R$ da CPRB não pode ser calculado — a skill
  informa só o percentual aplicável e sinaliza a lacuna, nunca inventa a
  receita.

## Passo a passo

### 0. Preparar o motor de cálculo (uma vez por ambiente novo)

Esta skill depende de um script Python determinístico para a parte de
cálculo (agregações, outliers estatísticos, custo empresa) — não refaça
essas contas "no olho" linha por linha, isso é o que garante um número
auditável e reproduzível. Se o arquivo `scripts/analisar_folha.py` ainda
não existir no diretório de trabalho atual, crie-o agora com exatamente o
conteúdo abaixo (copie literalmente, não resuma nem "melhore" ao digitar):

```python
#!/usr/bin/env python3
"""
analisar_folha.py — Motor de cálculo da skill analista-folha-pagamento-senior.

Lê a planilha consolidada de holerites (abas "Resumo" e "Rubricas", no
formato gerado pela skill consolidar-holerites-aw) e gera um Excel de
auditoria com:
  - Resumo Executivo
  - Rubricas Atípicas (desvio estatístico por função + rubrica)
  - Média Salarial por Função
  - Custo Empresa (FGTS + CPRB/CPP híbrido de desoneração + RAT/FAP + Terceiros)
  - Premissas (memória de cálculo — leia antes de comentar qualquer número)

Não inventa dado: coluna ausente ou insuficiente gera aviso explícito na aba
Premissas, nunca um número "chutado". Ver o Apêndice B deste SKILL.md para a
fonte de cada alíquota usada abaixo.
"""
import argparse
import re
import sys
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Constantes de encargos — ver Apêndice B
# ---------------------------------------------------------------------------

CPRB_ALIQUOTA_CHEIA = 0.045  # alíquota cheia de CPRB s/ receita bruta, construção civil (Lei 12.546/2011)
FGTS_PCT = 0.08
TERCEIROS_PCT = 0.058  # Sistema S / terceiros, aproximado
RAT_PADRAO = 0.02  # ponto médio da faixa legal 1%-3%, ajustado pelo FAP informado

# Cronograma de reoneração gradual (Lei 14.973/2024): percentual da CPRB
# cheia ainda aplicável sobre a receita bruta, e CPP básica sobre a folha,
# por ano. 2025 tem fontes divergentes sobre a CPP (5% x 10%) — sinalizado
# no aviso; confirme antes de fechar número para esse ano específico.
CRONOGRAMA_DESONERACAO = {
    2025: {"pct_cprb_aplicavel": 0.80, "cpp_pct_folha": 0.05,
           "aviso": "Fontes consultadas divergem sobre a CPP de 2025 (5% vs 10%). "
                    "Confirme a alíquota efetiva antes de fechar o número para esse ano."},
    2026: {"pct_cprb_aplicavel": 0.60, "cpp_pct_folha": 0.10, "aviso": None},
    2027: {"pct_cprb_aplicavel": 0.40, "cpp_pct_folha": 0.15, "aviso": None},
    2028: {"pct_cprb_aplicavel": 0.00, "cpp_pct_folha": 0.20,
           "aviso": "Cronograma de reoneração encerrado: CPRB extinta, CPP volta "
                    "a 20% cheio sobre a folha (equivalente ao INSS patronal integral)."},
}

RESUMO_REQUIRED = [
    "Competência", "Cargo", "Salário Base", "Sal. Contr. INSS",
    "Total Vencimentos", "Base Cálc. FGTS", "FGTS do Mês",
    "Nome do Funcionário", "Código Funcionário",
]
RUBRICAS_REQUIRED = [
    "Competência", "Código Funcionário", "Nome do Funcionário",
    "Descrição Rubrica", "Natureza", "Vencimento", "Desconto",
]

MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


# ---------------------------------------------------------------------------
# Carga e normalização
# ---------------------------------------------------------------------------

def parse_competencia(valor):
    """Extrai (ano, mes, chave_ordenavel) de um valor de Competência.
    Aceita 'MM/AAAA', 'AAAA-MM', datas do pandas/Excel, ou nome de mês em
    PT-BR ('Janeiro/2026'). Se não conseguir interpretar, devolve
    (None, None, str(valor)) — quem chamar deve tratar como não-cronológico
    em vez de arriscar uma ordenação errada."""
    if pd.isna(valor):
        return None, None, ""
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.year, valor.month, f"{valor.year:04d}-{valor.month:02d}"

    s = str(valor).strip()

    m = re.match(r"^(\d{1,2})[/\-](\d{4})$", s)
    if m:
        mes, ano = int(m.group(1)), int(m.group(2))
        if 1 <= mes <= 12:
            return ano, mes, f"{ano:04d}-{mes:02d}"

    m = re.match(r"^(\d{4})[/\-](\d{1,2})$", s)
    if m:
        ano, mes = int(m.group(1)), int(m.group(2))
        if 1 <= mes <= 12:
            return ano, mes, f"{ano:04d}-{mes:02d}"

    m = re.match(r"^([A-Za-zçÇ]+)[/\- ](\d{4})$", s)
    if m:
        nome, ano = m.group(1).lower()[:3], int(m.group(2))
        if nome in MESES_PT:
            mes = MESES_PT[nome]
            return ano, mes, f"{ano:04d}-{mes:02d}"

    return None, None, s


def carregar_base(caminho):
    try:
        resumo = pd.read_excel(caminho, sheet_name="Resumo")
    except Exception as e:
        sys.exit(f"[erro] Não consegui ler a aba 'Resumo' de {caminho}: {e}")
    try:
        rubricas = pd.read_excel(caminho, sheet_name="Rubricas")
    except Exception as e:
        sys.exit(f"[erro] Não consegui ler a aba 'Rubricas' de {caminho}: {e}")

    faltando_resumo = [c for c in RESUMO_REQUIRED if c not in resumo.columns]
    faltando_rubricas = [c for c in RUBRICAS_REQUIRED if c not in rubricas.columns]
    if faltando_resumo:
        sys.exit(
            f"[erro] Aba 'Resumo' sem as colunas esperadas: {faltando_resumo}. "
            "Confirme se o arquivo veio da skill consolidar-holerites-aw, ou "
            "ajuste os nomes das colunas antes de rodar — não estime o que falta."
        )
    if faltando_rubricas:
        sys.exit(
            f"[erro] Aba 'Rubricas' sem as colunas esperadas: {faltando_rubricas}."
        )

    resumo["Código Funcionário"] = resumo["Código Funcionário"].astype(str).str.strip()
    rubricas["Código Funcionário"] = rubricas["Código Funcionário"].astype(str).str.strip()
    resumo["Cargo"] = resumo["Cargo"].fillna("(Cargo não identificado)").astype(str).str.strip()

    for df in (resumo, rubricas):
        anos, meses, chaves = [], [], []
        for v in df["Competência"]:
            a, m, k = parse_competencia(v)
            anos.append(a)
            meses.append(m)
            chaves.append(k)
        df["_ano"] = anos
        df["_mes"] = meses
        df["_competencia_chave"] = chaves

    return resumo, rubricas


# ---------------------------------------------------------------------------
# 1) Rubricas atípicas
# ---------------------------------------------------------------------------

def calcular_rubricas_atipicas(resumo, rubricas):
    chave_resumo = resumo[["Código Funcionário", "_competencia_chave", "Cargo"]].drop_duplicates()
    rub = rubricas.merge(chave_resumo, on=["Código Funcionário", "_competencia_chave"], how="left")
    rub["Cargo"] = rub["Cargo"].fillna("(Cargo não identificado)")

    rub["Natureza_norm"] = rub["Natureza"].astype(str).str.strip().str.lower()
    proventos = rub[rub["Natureza_norm"].str.startswith("prov", na=False)].copy()
    proventos["Valor"] = pd.to_numeric(proventos["Vencimento"], errors="coerce").fillna(0.0)

    agrupado = (
        proventos.groupby(["Cargo", "Descrição Rubrica", "_competencia_chave"])["Valor"]
        .sum()
        .reset_index()
    )

    linhas = []
    for (cargo, rubrica_desc), grupo in agrupado.groupby(["Cargo", "Descrição Rubrica"]):
        grupo = grupo.sort_values("_competencia_chave")
        valores = grupo["Valor"].tolist()
        competencias = grupo["_competencia_chave"].tolist()
        if not valores:
            continue

        ultimo_valor = valores[-1]
        ultima_competencia = competencias[-1]
        n_meses = len(valores)

        atipico, motivo = False, ""
        if n_meses >= 4:
            historico = valores[:-1]
            media_hist = float(np.mean(historico))
            desvio_hist = float(np.std(historico, ddof=1)) if len(historico) >= 2 else 0.0
            if desvio_hist > 0:
                z = (ultimo_valor - media_hist) / desvio_hist
                if abs(z) > 2:
                    atipico = True
                    motivo = f"{z:+.1f} desvios-padrão da média histórica (R$ {media_hist:,.2f})"
        if not atipico and n_meses >= 2:
            anterior = valores[-2]
            if anterior != 0:
                variacao = (ultimo_valor - anterior) / abs(anterior)
                if abs(variacao) > 0.30:
                    atipico = True
                    motivo = (f"variação de {variacao:+.0%} sobre o mês anterior "
                              f"(R$ {anterior:,.2f}); histórico curto ({n_meses} mês(es))")

        linhas.append({
            "Cargo": cargo,
            "Descrição Rubrica": rubrica_desc,
            "Última Competência": ultima_competencia,
            "Valor no Mês (R$)": round(ultimo_valor, 2),
            "Meses no Histórico": n_meses,
            "Atípico": "Sim" if atipico else "Não",
            "Motivo": motivo,
        })

    df = pd.DataFrame(linhas)
    if not df.empty:
        df = df.sort_values(["Atípico", "Valor no Mês (R$)"], ascending=[False, False])
    return df


# ---------------------------------------------------------------------------
# 2) Média salarial por função
# ---------------------------------------------------------------------------

def calcular_media_salarial(resumo):
    ultima_competencia = resumo["_competencia_chave"].max()
    base = resumo[resumo["_competencia_chave"] == ultima_competencia].copy()
    if base.empty:
        base = resumo.copy()

    for col in ["Salário Base", "Total Vencimentos"]:
        base[col] = pd.to_numeric(base[col], errors="coerce")

    agg = base.groupby("Cargo").agg(
        Funcionários=("Código Funcionário", "nunique"),
        Salário_Base_Médio=("Salário Base", "mean"),
        Salário_Base_Mediana=("Salário Base", "median"),
        Salário_Base_DesvPad=("Salário Base", "std"),
        Total_Vencimentos_Médio=("Total Vencimentos", "mean"),
    ).reset_index()

    agg = agg.rename(columns={
        "Salário_Base_Médio": "Salário Base — Média (R$)",
        "Salário_Base_Mediana": "Salário Base — Mediana (R$)",
        "Salário_Base_DesvPad": "Salário Base — Desvio-Padrão (R$)",
        "Total_Vencimentos_Médio": "Total Vencimentos — Média (R$)",
    })
    for c in agg.columns:
        if c not in ("Cargo", "Funcionários"):
            agg[c] = agg[c].round(2)
    agg = agg.sort_values("Total Vencimentos — Média (R$)", ascending=False)
    agg["Competência de Referência"] = ultima_competencia
    return agg


# ---------------------------------------------------------------------------
# 3) Custo empresa
# ---------------------------------------------------------------------------

def calcular_custo_empresa(resumo, receita_bruta, fap, ano_cprb):
    ultima_competencia = resumo["_competencia_chave"].max()
    base = resumo[resumo["_competencia_chave"] == ultima_competencia].copy()
    if base.empty:
        base = resumo.copy()

    for col in ["Total Vencimentos", "Sal. Contr. INSS", "Base Cálc. FGTS", "FGTS do Mês"]:
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0.0)

    n_func = base["Código Funcionário"].nunique()
    folha_bruta = base["Total Vencimentos"].sum()
    base_inss = base["Sal. Contr. INSS"].sum()
    base_fgts = base["Base Cálc. FGTS"].sum()
    fgts_apurado = base["FGTS do Mês"].sum()
    fgts_calculado_8pct = base_fgts * FGTS_PCT

    schedule = CRONOGRAMA_DESONERACAO.get(ano_cprb)
    if schedule is None:
        aviso_schedule = (
            f"Ano {ano_cprb} fora do cronograma de reoneração conhecido (2025-2028). "
            "Confirme a alíquota vigente antes de usar este número — nenhuma CPP/CPRB foi calculada."
        )
        cpp_pct_folha, pct_cprb_aplicavel = 0.0, 0.0
    else:
        cpp_pct_folha = schedule["cpp_pct_folha"]
        pct_cprb_aplicavel = schedule["pct_cprb_aplicavel"]
        aviso_schedule = schedule["aviso"]

    cpp = base_inss * cpp_pct_folha
    rat = base_inss * RAT_PADRAO * fap
    terceiros = base_inss * TERCEIROS_PCT

    cprb_reais = None
    if receita_bruta is not None:
        cprb_reais = receita_bruta * CPRB_ALIQUOTA_CHEIA * pct_cprb_aplicavel

    fgts_usado = fgts_apurado if fgts_apurado > 0 else fgts_calculado_8pct
    custo_sobre_folha = fgts_usado + cpp + rat + terceiros
    custo_total_mes = folha_bruta + custo_sobre_folha
    custo_medio_por_cabeca = (custo_total_mes / n_func) if n_func else None

    resumo_empresa = {
        "Competência de Referência": ultima_competencia,
        "Nº Funcionários": n_func,
        "Folha Bruta (R$)": round(folha_bruta, 2),
        "FGTS (R$)": round(fgts_usado, 2),
        f"CPP {cpp_pct_folha:.0%} sobre a folha (R$)": round(cpp, 2),
        "RAT ajustado pelo FAP (R$)": round(rat, 2),
        "Terceiros / Sistema S (R$)": round(terceiros, 2),
        "Custo Total Folha + Encargos (R$)": round(custo_total_mes, 2),
        "Custo Médio por Funcionário (R$)": (
            round(custo_medio_por_cabeca, 2) if custo_medio_por_cabeca else None
        ),
        "CPRB sobre Receita Bruta (R$)": (
            round(cprb_reais, 2) if cprb_reais is not None else "N/D — informe --receita-bruta"
        ),
        "Alíquota CPRB aplicável no ano": (
            f"{CPRB_ALIQUOTA_CHEIA * pct_cprb_aplicavel:.2%} da receita bruta" if schedule else "N/D"
        ),
    }

    por_funcao = base.groupby("Cargo").agg(
        Funcionários=("Código Funcionário", "nunique"),
        Salário_Bruto=("Total Vencimentos", "sum"),
    ).reset_index()
    if folha_bruta > 0:
        por_funcao["Encargos Rateados (R$)"] = (por_funcao["Salário_Bruto"] / folha_bruta) * custo_sobre_folha
    else:
        por_funcao["Encargos Rateados (R$)"] = 0.0
    por_funcao["Custo Total da Função (R$)"] = por_funcao["Salário_Bruto"] + por_funcao["Encargos Rateados (R$)"]
    por_funcao["Custo Médio por Cabeça (R$)"] = por_funcao["Custo Total da Função (R$)"] / por_funcao["Funcionários"]
    por_funcao = por_funcao.rename(columns={"Salário_Bruto": "Salário Bruto (R$)"})
    for c in ["Salário Bruto (R$)", "Encargos Rateados (R$)", "Custo Total da Função (R$)", "Custo Médio por Cabeça (R$)"]:
        por_funcao[c] = por_funcao[c].round(2)
    por_funcao = por_funcao.sort_values("Custo Total da Função (R$)", ascending=False)

    premissas = [
        f"Competência de referência: {ultima_competencia}",
        f"Ano usado para o cronograma de desoneração: {ano_cprb}",
        (f"CPP sobre a folha: {cpp_pct_folha:.0%} | CPRB aplicável: {pct_cprb_aplicavel:.0%} da alíquota "
         f"cheia de {CPRB_ALIQUOTA_CHEIA:.1%} sobre a receita bruta (construção civil, Lei 12.546/2011)"),
        f"FAP usado no cálculo do RAT: {fap} (padrão 1.0 se não informado)",
        ("RAT: ponto médio da faixa legal (1% a 3%) ajustado pelo FAP — confirme o RAT real da "
         "atividade da empresa (CNAE + FAP publicado) para maior precisão"),
        f"Terceiros/Sistema S: {TERCEIROS_PCT:.1%} (aproximado; não há isenção para empresa desonerada)",
        ("FGTS: usado o valor já apurado no holerite (coluna 'FGTS do Mês') quando disponível; "
         "senão, 8% sobre a Base Cálc. FGTS"),
        ("Provisão de multa rescisória (40% do FGTS + aviso prévio + férias/13º proporcionais) NÃO foi "
         "calculada automaticamente — depende da rotatividade real da empresa. Peça a taxa de "
         "rotatividade ao usuário antes de estimar esse número; não invente um percentual genérico."),
        (f"Receita bruta informada: {'Sim — R$ ' + format(receita_bruta, ',.2f') if receita_bruta is not None else 'Não — CPRB em R$ não calculada, apenas o percentual aplicável'}"),
    ]
    if aviso_schedule:
        premissas.append(f"AVISO: {aviso_schedule}")

    return resumo_empresa, por_funcao, premissas


# ---------------------------------------------------------------------------
# Escrita do Excel de saída
# ---------------------------------------------------------------------------

def autofit(ws, max_width=60):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[letter].width = min(max(length + 2, 10), max_width)


def escrever_df(wb, nome_aba, df, nome_tabela):
    ws = wb.create_sheet(nome_aba)
    if df is None or df.empty:
        ws.append(["(sem dados suficientes para esta aba)"])
        return ws
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for _, row in df.iterrows():
        ws.append(list(row))
    n_rows, n_cols = len(df) + 1, len(df.columns)
    ref = f"A1:{get_column_letter(n_cols)}{n_rows}"
    tabela = Table(displayName=nome_tabela, ref=ref)
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabela)
    ws.freeze_panes = "A2"
    autofit(ws)
    return ws


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="Excel consolidado (abas Resumo e Rubricas).")
    ap.add_argument("--output", required=True, help="Caminho do Excel de auditoria a ser gerado.")
    ap.add_argument("--receita-bruta", type=float, default=None,
                     help="Receita bruta do período, para calcular a CPRB em R$ (opcional).")
    ap.add_argument("--fap", type=float, default=1.0, help="Fator Acidentário de Prevenção (padrão 1.0).")
    ap.add_argument("--ano-cprb", type=int, default=None,
                     help="Ano de referência do cronograma de desoneração "
                          "(padrão: ano da última competência da base).")
    args = ap.parse_args()

    resumo, rubricas = carregar_base(args.input)

    anos_validos = [a for a in resumo["_ano"].tolist() if a is not None]
    ano_cprb = args.ano_cprb or (max(anos_validos) if anos_validos else datetime.now().year)

    print(f"[info] {resumo['Código Funcionário'].nunique()} funcionário(s), "
          f"{resumo['_competencia_chave'].nunique()} competência(s) na base.")

    df_atipicas = calcular_rubricas_atipicas(resumo, rubricas)
    df_media_salarial = calcular_media_salarial(resumo)
    resumo_empresa, df_por_funcao, premissas = calcular_custo_empresa(
        resumo, args.receita_bruta, args.fap, ano_cprb
    )

    n_atipicas = int((df_atipicas["Atípico"] == "Sim").sum()) if not df_atipicas.empty else 0
    print(f"[info] {n_atipicas} combinação(ões) função+rubrica sinalizada(s) como atípica.")
    print(f"[info] Custo total (folha + encargos) no período: R$ {resumo_empresa['Custo Total Folha + Encargos (R$)']:,.2f}")

    wb = Workbook()
    wb.remove(wb.active)

    # --- Resumo Executivo ---
    ws = wb.create_sheet("Resumo Executivo")
    ws.append(["Auditoria de Folha de Pagamento — Resumo Executivo"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")])
    for k, v in resumo_empresa.items():
        ws.append([k, v])
    ws.append([])
    ws.append([f"Top rubricas atípicas (competência {resumo_empresa['Competência de Referência']})"])
    ws[ws.max_row][0].font = Font(bold=True)
    ws.append(["Cargo", "Rubrica", "Valor no Mês (R$)", "Motivo"])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    if not df_atipicas.empty:
        top = df_atipicas[df_atipicas["Atípico"] == "Sim"].head(10)
        for _, row in top.iterrows():
            ws.append([row["Cargo"], row["Descrição Rubrica"], row["Valor no Mês (R$)"], row["Motivo"]])
    autofit(ws, max_width=80)

    # --- Rubricas Atípicas ---
    escrever_df(wb, "Rubricas Atípicas", df_atipicas, "tbAtipicas")

    # --- Média Salarial por Função ---
    escrever_df(wb, "Média Salarial por Função", df_media_salarial, "tbMediaSalarial")

    # --- Custo Empresa ---
    ws = wb.create_sheet("Custo Empresa")
    ws.append(["Custo Empresa Consolidado"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    for k, v in resumo_empresa.items():
        ws.append([k, v])
    ws.append([])
    ws.append(["Custo por Função"])
    ws[ws.max_row][0].font = Font(bold=True)
    header_row = ws.max_row + 1
    for col_i, col_name in enumerate(df_por_funcao.columns, start=1):
        ws.cell(row=header_row, column=col_i, value=col_name).font = Font(bold=True)
    for _, row in df_por_funcao.iterrows():
        ws.append(list(row))
    autofit(ws, max_width=60)

    # --- Premissas ---
    ws = wb.create_sheet("Premissas")
    ws.append(["Premissas e Memória de Cálculo — leia antes de comentar qualquer número"])
    ws["A1"].font = Font(bold=True, size=12)
    ws.append([])
    for p in premissas:
        ws.append([p])
        ws[ws.max_row][0].alignment = Alignment(wrap_text=True, vertical="top")
    autofit(ws, max_width=110)

    wb.save(args.output)
    print(f"[info] Excel de auditoria salvo em: {args.output}")


if __name__ == "__main__":
    main()
```

### 1. Rodar o motor de cálculo

```bash
python scripts/analisar_folha.py \
  --input "/caminho/para/base_holerites_AW.xlsx" \
  --output "/caminho/para/Auditoria_Folha_<competencia>.xlsx" \
  [--receita-bruta 1234567.89] \
  [--fap 1.0] \
  [--ano-cprb 2026]
```

Todas as premissas assumidas (ano do cronograma de desoneração, FAP usado,
se a receita bruta foi informada) ficam registradas na aba `Premissas` da
saída — leia essa aba antes de comentar qualquer número com o usuário, e
repita as premissas relevantes na sua resposta em texto.

### 2. Interpretar os outliers de rubrica

O script sinaliza como atípica qualquer combinação função+rubrica cujo valor
do mês esteja a mais de 2 desvios-padrão da própria média histórica (quando
há pelo menos 3 meses de histórico), ou com variação mês a mês acima de 30%
quando o histórico ainda é curto. Isso é ponto de partida, não veredito —
antes de concluir algo, contextualize com o que cada rubrica costuma
significar em obra (ver **Apêndice A**, ao final deste documento):

- **PRODUÇÃO / PRÊMIO** alto e crescente pode ser aumento real de
  produtividade (bom sinal) ou meta mal calibrada no orçamento da obra,
  "vazando" margem — antes de cravar recomendação, cruze com o orçamento da
  obra (skill `orcamentista-master-civil`, se disponível).
- **HORAS EXTRAS** alta pode ser sazonal (fechamento de etapa, prazo
  apertado — normal) ou estrutural (recorrente por 3+ meses na mesma
  função — sinal de quadro insuficiente). Só trate como problema o padrão
  estrutural; hora extra pontual em pico de obra é esperado.
- **ADICIONAL NOTURNO** fora do padrão costuma indicar turno extra não
  planejado — questão de cronograma de obra, não de RH.
- **INSALUBRIDADE / PERICULOSIDADE** atípica pode ser pagamento sem laudo
  técnico vigente (risco trabalhista) ou mudança de função sem atualização
  cadastral — confirme com o usuário antes de tratar como custo indevido,
  nunca afirme isso sozinho a partir da planilha.

### 3. Média salarial por função x mercado

O script agrupa por Cargo e calcula média, mediana e desvio-padrão do
salário base e do total (base + variáveis). Para comparar com mercado, siga
essa ordem (ver **Apêndice C** para como consultar cada fonte):

1. **Piso da CCT** do sindicato regional da obra (é o piso legal mínimo, não
   "o mercado" — nunca confunda os dois).
2. **RAIS/CAGED e Salario.com.br** para a mediana de mercado factual da
   região.
3. Nunca diga "está acima/abaixo do mercado" sem citar contra qual fonte
   comparou.

Se o usuário não disser a região/estado da obra, pergunte antes de buscar
convenção coletiva — não existe piso nacional único na construção civil,
cada sindicato regional tem o seu.

### 4. Custo total empresa (não é só o salário)

O script calcula, por funcionário, por função e em nível de empresa:
salário bruto (base + variáveis) + FGTS (8%, usando o valor já apurado no
holerite quando disponível) + encargo previdenciário do regime híbrido de
desoneração vigente (CPRB sobre a receita bruta + CPP sobre a folha,
conforme o ano — ver **Apêndice B** para o cronograma completo 2025-2028) +
RAT ajustado pelo FAP + Terceiros/Sistema S (~5,8%).

Sempre explicite as premissas usadas (ano do cronograma, FAP, se a receita
bruta foi informada) — nunca apresente o custo empresa como número fechado
sem a memória de cálculo ao lado. A provisão de rescisão (multa de 40% do
FGTS, aviso prévio, férias/13º proporcionais) depende da rotatividade real
da empresa — se o usuário não informar a taxa de rotatividade, mostre a
fórmula e peça o dado em vez de estimar um percentual genérico.

### 5. Recorrência atípica → ação recomendada

Para desvios que se repetem 3 ou mais meses seguidos na mesma função e
rubrica, proponha uma ação concreta, não genérica — use o **Apêndice D**
como ponto de partida (revisão de dimensionamento de equipe, terceirização
de mão de obra por empreitada, revisão do preço de venda que embute a
produção, saneamento de rubrica sem lastro). Compare sempre em R$: custo
marginal de 1 CLT adicional (salário + todos os encargos do passo 4) x
custo da hora extra recorrente x custo estimado de terceirização daquela
frente. Não recomende sem números ao lado, e deixe claro que é uma hipótese
a validar com o usuário — nunca um veredito automático do sistema.

## Formato de entrega

Excel gerado por `scripts/analisar_folha.py` (abas: `Resumo Executivo`,
`Rubricas Atípicas`, `Média Salarial por Função`, `Custo Empresa`,
`Premissas`) + conclusão direta no chat, no padrão de resposta do usuário
(PT-BR direto, resposta/recomendação primeiro, sem repetir o básico): o que
mais desviou, o custo empresa total do período, e a recomendação
prioritária.

## Nunca faça

- Nunca invente rubrica, valor, alíquota, região ou fonte de mercado — o que
  não está na base ou não foi pesquisado fica marcado como "a confirmar",
  nunca preenchido "no olho".
- Nunca aplique a alíquota de desoneração sem considerar se a atividade da
  empresa está de fato dentro do regime (construção civil, CNAE 41 a 43) —
  se parte da receita vier de atividade fora desse escopo, avise que a CPRB
  não se aplica a essa parcela.
- Nunca recomende terceirização, corte de quadro ou desligamento como
  conclusão automática de um outlier estatístico — é hipótese a validar com
  o usuário, com números ao lado, não um veredito do sistema.

## Combine com

- `consolidar-holerites-aw` — gera a base de entrada desta skill.
- `matrizes-ponto-catraca` / `conciliacao-ponto-aw` / `validar-ponto-catraca`
  — quando a hora extra apontada na folha precisar ser cruzada com o ponto
  ou a catraca real da obra.
- `calculo-processo-trabalhista` — se um desvio (insalubridade, hora extra
  não paga) escalar para risco de reclamação trabalhista.
- `analista-tributario-senior` — para aprofundar o regime tributário além do
  que esta skill assume por padrão (CPRB híbrido de construção civil).
- `orcamentista-master-civil` — para ligar a rubrica de Produção ao
  orçamento da obra (mão de obra direta x preço de venda).
- `controller-empreiteira` — para levar o resultado desta análise ao DRE
  gerencial e ao fluxo de caixa consolidado.

---

## Apêndice A — Rubricas típicas de folha em obra (TOTVS Domínio)

Este catálogo não substitui a aba `Dicionario_Rubricas` da base do usuário
(ela reflete os códigos reais em uso na empresa) — é um guia para interpretar
o que cada rubrica costuma significar em uma empreiteira de construção
civil, e o que costuma explicar uma variação atípica nela. Sempre confira a
descrição real da rubrica na base antes de aplicar a leitura abaixo — nomes
podem variar entre empresas e entre versões do Domínio.

**Estrutura de códigos do Domínio**: o Domínio organiza rubricas por faixas
de código (padrão observado em manuais e no portal de suporte do
fornecedor, pode variar por implantação): faixas 1–200/250 para proventos e
descontos de uso corrente, e faixas 800–999 / 8000–9756 para rubricas de
eSocial, retenções especiais e eventos de fechamento. Duas rubricas de
referência bem documentadas: **150 — Horas Extras 50%** (dia útil, valor da
hora normal + 50%) e **200 — Horas Extras 100%** (domingo/feriado ou dobra).
O **DSR sobre horas extras** normalmente aparece como reflexo automático
vinculado a essas rubricas, não como lançamento manual — se aparecer
isolado e grande, vale conferir se não há duplicidade. O **adicional
noturno** costuma ser configurado como uma cópia da lógica de horas extras,
com base de cálculo própria — por isso pode aparecer sob nomes variados
dependendo da parametrização de cada empresa.

| Rubrica | O que costuma significar | O que uma alta atípica costuma indicar |
|---|---|---|
| Produção / Prêmio de Produção | Remuneração variável ligada a meta de produtividade (comum em armação, alvenaria, forma) | Produtividade real subindo (bom) OU meta de orçamento mal calibrada, "vazando" margem — cruzar com o orçamento antes de concluir |
| Horas Extras 50% / 100% | Jornada além da 44h semanal, ou trabalho em domingo/feriado | Pico pontual de obra (normal) OU quadro insuficiente se recorrente 3+ meses na mesma função |
| DSR sobre Horas Extras | Reflexo automático das horas extras no descanso semanal remunerado | Segue o padrão das horas extras; se subir sozinho sem elas subirem, investigar erro de cálculo |
| Adicional Noturno | Trabalho entre 22h e 5h | Turno extra não planejado — normalmente questão de cronograma de obra, não de RH |
| Insalubridade | Exposição a agente insalubre (ruído, poeira, calor etc.), % sobre salário mínimo ou salário base conforme grau | Pagamento sem laudo técnico (PPRA/PGR) vigente é risco trabalhista; mudança de função sem atualização cadastral também explica oscilação |
| Periculosidade | Atividade de risco (eletricidade, altura, inflamáveis), 30% sobre o salário base | Mesmo cuidado da insalubridade — confirmar laudo e enquadramento antes de tratar como indevido |
| Vale-Transporte / Vale-Alimentação | Benefícios fixos ou por dia trabalhado | Alta atípica geralmente é erro de lançamento (duplicidade) ou mudança de política, raramente "custo de obra" |
| Faltas / Atrasos (desconto) | Ausência não justificada | Queda atípica de líquido pode ser isso, não necessariamente rubrica de provento subindo |
| Consignado / Empréstimo | Desconto de crédito consignado do funcionário | Não é custo da empresa — não deve entrar na análise de custo empresa, é passagem de valor de terceiro |

Ao ver uma rubrica sinalizada como atípica pelo script, primeiro identifique
a que categoria da tabela acima ela pertence (pela descrição real na base,
não pelo nome desta tabela), aplique a pergunta certa para aquela categoria
antes de escrever a conclusão, e nunca trate "rubrica grande" e "rubrica
atípica" como sinônimos — o que importa é o desvio em relação ao próprio
histórico daquela função.

## Apêndice B — Encargos sobre a folha e desoneração (CPRB), 2025-2028

Base legal e premissas usadas pelo script acima. Revise periodicamente — a
reoneração é gradual e o cronograma pode ser alterado por nova legislação;
confirme a alíquota vigente antes de fechar números para fins de decisão
(preço de venda, orçamento, negociação).

**O regime da empreiteira**: a construção civil (CNAE 41 a 43) está entre os
setores que podem optar pela Contribuição Previdenciária sobre a Receita
Bruta (CPRB, Lei 12.546/2011), em substituição à CPP (INSS patronal) de 20%
sobre a folha. A alíquota cheia de CPRB para construção civil é **4,5% sobre
a receita bruta**. Isso só se aplica à receita da atividade desonerada — se
a empreiteira tiver receita fora do escopo de construção civil (CNAE
diferente), essa parcela não entra na base da CPRB.

**Cronograma de reoneração gradual** (Lei 14.973/2024, publicada
16/09/2024): a partir de 2025 há um retorno gradual ao INSS patronal pleno,
em paralelo com a redução progressiva da CPRB:

| Ano | CPRB aplicável (% da alíquota cheia de 4,5%) | CPP sobre a folha |
|---|---|---|
| 2025 | 80% → CPRB efetiva ≈ 3,6% da receita bruta | fontes divergem entre 5% e 10% — **confirme antes de usar** |
| 2026 | 60% → CPRB efetiva ≈ 2,7% da receita bruta | 10% sobre a folha |
| 2027 | 40% → CPRB efetiva ≈ 1,8% da receita bruta | 15% sobre a folha |
| 2028 | 0% (CPRB extinta) | 20% sobre a folha (volta ao INSS patronal integral) |

Ou seja, hoje (2026) a empresa paga os dois ao mesmo tempo: CPRB sobre a
receita bruta E uma CPP básica sobre a folha — não é mais "ou um, ou outro"
como era antes da Lei 14.973/2024. O 13º salário permanece com tratamento
apenas de CPRB (sem a CPP adicional) durante a transição. Fontes
consultadas: Contabeis.com.br ("Desoneração da folha de 2024 a 2027" e
"Desoneração da folha em 2026: reoneração gradual"), Sienge.com.br (blog),
Planning.com.br — nenhuma é fonte primária oficial; para decisão de alto
impacto, confirme com o contador/tributarista ou a skill
`analista-tributario-senior`.

**Reforma Tributária (EC 132/2023 + LC 214/2025) — não muda isso**: incide
sobre o consumo (IBS/CBS), não sobre a folha. Para construção civil há
alíquota reduzida (~50% da padrão, ~13,25% efetivo) em obras/incorporações,
mas isso é IBS/CBS na venda, não afeta FGTS/CPRB/CPP sobre a folha.

**Demais encargos — sem isenção para empresa desonerada**: a desoneração
substitui apenas a CPP. Continuam devidos integralmente: **FGTS** (8% +
multa rescisória de 40% em demissão sem justa causa), **RAT/SAT** (1% a 3%
sobre a folha, multiplicado pelo **FAP** publicado anualmente em janeiro
pelo INSS/MTE, específico por CNPJ), e **Terceiros/Sistema S** (SENAI,
SESI, SESC, SENAC, SEBRAE, INCRA, salário-educação, ~5,8%).

O script calcula FGTS, CPP híbrida do ano, RAT aproximado (2% x FAP) e
Terceiros aproximado sempre; calcula a CPRB em R$ apenas se a receita bruta
for informada. **Não calcula** a provisão de rescisão nem o RAT exato por
CNAE — peça esses dados ao usuário antes de tratar o custo empresa como
número fechado para uma decisão de precificação ou corte.

## Apêndice C — Fontes para comparar salário com o mercado

Não existe piso salarial nacional único para construção civil no Brasil —
cada sindicato regional (SindusCon patronal + sindicato laboral do
estado/região) negocia sua própria Convenção Coletiva de Trabalho (CCT).
Sempre confirme a região/estado da obra antes de buscar a CCT — comparar com
a CCT errada invalida a conclusão.

**Ordem recomendada de consulta**:

1. **Piso da CCT do sindicato regional** — piso legal mínimo por
   categoria/função, reajustado anualmente (normalmente maio, varia por
   região; ex.: CCT SindusCon-SP/Sintracon-SP). É o piso, não "o mercado" —
   nunca apresente um como o outro.
2. **RAIS / Novo CAGED** (Ministério do Trabalho e Emprego, dados.gov.br) —
   dados oficiais de admissões, desligamentos e salário médio por CBO e
   CNAE. Estatisticamente mais robusto, mas com defasagem de divulgação.
3. **Salario.com.br** — tabela salarial por cargo e região, bom para
   cross-check rápido, não como única fonte.
4. **Catho, Vagas.com, Glassdoor** — complementares, amostra pequena e
   viesada para vagas anunciadas online; pouco confiáveis para funções
   operacionais de obra (servente, armador, meio-oficial); melhores para
   cargos administrativos/técnicos (engenheiro, mestre de obras, técnico de
   segurança).

**Como aplicar**: busque a CCT vigente do sindicato regional da obra em
questão, registre a data de vigência usada, e compare o **salário base**
(fixo) com o piso da CCT e com a mediana de RAIS/CAGED/Salario.com.br —
nunca compare o "Total Vencimentos" (que inclui variáveis) com o piso de
CCT, isso distorce a comparação. Sempre cite a fonte e a data da consulta.
Se o usuário não informar a região, pergunte antes de pesquisar.

## Apêndice D — Cases de mão de obra direta: da rubrica atípica à ação

Padrões recorrentes em empreiteiras, não regras automáticas. Sempre
apresente a comparação numérica junto com a hipótese, e deixe claro que é
uma hipótese a validar com o usuário.

**Case 1 — Hora extra estrutural (não sazonal)**: sinal = horas extras
recorrentes por 3+ meses na mesma função, sem correlação com evento pontual
de obra. Costuma significar quadro dimensionado abaixo da carga real de
trabalho. Comparar: custo mensal médio da hora extra + reflexos, contra o
custo marginal de 1 CLT adicional (salário de mercado + todos os encargos
do Apêndice B). Se a hora extra se aproxima ou supera esse custo de forma
consistente, a contratação costuma ser mais barata — mas confirme
disponibilidade de mão de obra na região e horizonte da obra. Alternativa:
terceirização por empreitada daquela frente (Case 3), quando a atividade é
bem delimitável e não é núcleo da operação.

**Case 2 — Produção/Prêmio alto e crescente**: pode ser ganho real de
produtividade (bom, mas checar se o preço de venda está capturando isso) ou
meta de orçamento subestimada, "vazando" margem. Cruze o crescimento da
rubrica com o avanço físico/medição da obra no mesmo período (skill
`orcamentista-master-civil`, se disponível, ou pergunte ao usuário). Se a
produção cresce mais rápido que o avanço físico, sinalize para revisão da
meta no próximo orçamento, não como correção retroativa.

**Case 3 — Terceirização de mão de obra direta por empreitada**: costuma
valer a pena considerar para atividades bem delimitáveis e mensuráveis por
produção (armação, alvenaria, forma, revestimento, pintura), com volume
intermitente entre obras. Compare custo total CLT direto por
m²/unidade de produção (salário + variáveis + todos os encargos, dividido
pela produção física do período) x preço de empreitada de mercado por
m²/unidade (pesquisa com fornecedores da região). Tende a comparar melhor
quando há hora extra estrutural (Case 1), ociosidade entre obras, ou a
atividade não é competência central da empreiteira. **Atenção**:
terceirização de mão de obra tem regras específicas de responsabilidade
solidária/subsidiária trabalhista e previdenciária — sinalize esse risco
sempre que recomendar essa via; não é decisão só financeira.

**Case 4 — Insalubridade/Periculosidade sem lastro aparente**: rubrica
aparecendo, mudando de grau, ou variando sem mudança de função/frente
visível na base. Risco trabalhista (pagamento sem laudo técnico vigente que
respalde o enquadramento) é mais provável do que "custo indevido a
cortar" — não trate como economia a capturar sem confirmar com o usuário se
há laudo (PPRA/PGR) e se o enquadramento está correto. Cortar sem essa
confirmação pode gerar passivo trabalhista maior que a economia gerada.

**Case 5 — Rotatividade alta em função de baixa especialização**: mesma
função aparecendo repetidamente em rescisões, com FGTS de multa recorrente.
Custo de rotatividade (rescisão + aviso prévio + ramp-up do substituto +
curva de aprendizado) pode superar o custo de reter a equipe, ou pode
indicar que aquela frente é boa candidata a terceirização por empreitada.
Peça ao usuário a taxa de rotatividade real antes de quantificar esse
custo — o script não estima isso sozinho.

