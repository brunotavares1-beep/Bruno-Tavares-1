#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calculadora de folha - DP construcao civil.
Aritmetica deterministica: INSS progressivo, IRRF, encargos patronais,
custo-empresa/Fenc, ponto de equilibrio de faturamento e producao por tarefa.

REGRA: nao ha tabela embutida. Tudo vem de tabelas.json, que precisa ser
preenchido com a tabela VIGENTE confirmada, com data_base e fonte.
"""
import argparse, json, os, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
PADRAO = os.path.join(AQUI, "tabelas.json")


def num(v, casas=2):
    s = f"{v:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def brl(v):
    return "R$ " + num(v)


def pct(v):
    return num(v * 100) + "%"


def carrega(caminho, exigir=("inss_empregado",)):
    with open(caminho, encoding="utf-8") as f:
        t = json.load(f)
    if t.get("status") != "PREENCHIDA":
        sys.exit(
            "PENDENCIA: tabelas.json nao preenchido.\n"
            "Preencha as faixas vigentes, marque status='PREENCHIDA' e informe "
            "data_base e fonte. Nao estime valores."
        )
    if not t.get("data_base"):
        sys.exit("PENDENCIA: informe a data_base da tabela em tabelas.json.")
    return t


def inss_empregado(base, tab):
    teto = tab["inss_empregado"]["teto_salario_contribuicao"]
    faixas = tab["inss_empregado"]["faixas"]
    if teto is None or any(f["ate"] is None or f["aliquota"] is None for f in faixas):
        sys.exit("PENDENCIA: faixas/teto de INSS nao preenchidos em tabelas.json.")
    b = min(base, teto)
    total, anterior, detalhe = 0.0, 0.0, []
    for f in faixas:
        limite = min(f["ate"], b)
        if limite > anterior:
            parcela = (limite - anterior) * f["aliquota"]
            detalhe.append((anterior, limite, f["aliquota"], parcela))
            total += parcela
        anterior = f["ate"]
        if anterior >= b:
            break
    return total, detalhe


def irrf(base_bruta, inss, dependentes, pensao, tab, regime="progressivo"):
    ded_dep = tab["irrf"]["deducao_por_dependente"]
    simpl = tab["irrf"]["desconto_simplificado"]
    faixas = tab["irrf"]["faixas"]
    if any(f["ate"] is None for f in faixas[:-1]):
        sys.exit("PENDENCIA: faixas de IRRF nao preenchidas em tabelas.json.")
    if regime == "simplificado":
        if simpl is None:
            sys.exit("PENDENCIA: desconto_simplificado nao preenchido.")
        base = max(base_bruta - simpl, 0.0)
    else:
        if ded_dep is None and dependentes:
            sys.exit("PENDENCIA: deducao_por_dependente nao preenchida.")
        base = max(base_bruta - inss - (ded_dep or 0) * dependentes - pensao, 0.0)
    for f in faixas:
        teto = f["ate"] if f["ate"] is not None else float("inf")
        if base <= teto:
            return max(base * (f["aliquota"] or 0) - (f["parcela_deduzir"] or 0), 0.0), base
    return 0.0, base


def encargos_patronais(base, tab):
    e = tab["empresa"]
    for k in ("rat", "fap", "terceiros"):
        if e.get(k) is None:
            sys.exit(f"PENDENCIA: '{k}' nao preenchido em tabelas.json (empresa).")
    if e.get("cprb_ativa"):
        if e.get("cprb_aliquota") is None:
            sys.exit("PENDENCIA: cprb_ativa=true exige cprb_aliquota.")
        patronal = 0.0
        obs = "CPRB ativa: patronal substituida por aliquota sobre receita bruta (calcular fora da folha)."
    else:
        patronal = base * e["inss_patronal"]
        obs = ""
    rat = base * e["rat"] * e["fap"]
    terceiros = base * e["terceiros"]
    fgts = base * e["fgts"]
    return {
        "INSS patronal": patronal,
        "RAT x FAP": rat,
        "Terceiros/Sistema S": terceiros,
        "FGTS": fgts,
        "_total": patronal + rat + terceiros + fgts,
        "_obs": obs,
    }


def cmd_liquido(a, tab):
    base = a.base if a.base is not None else a.salario
    ins, det = inss_empregado(base, tab)
    ir, base_ir = irrf(base, ins, a.dependentes, a.pensao, tab, a.regime_ir)
    print(f"Tabela data-base: {tab['data_base']} | fonte: {tab.get('fonte') or 'PENDENCIA'}")
    print(f"Base de calculo........ {brl(base)}")
    for ini, fim, al, val in det:
        print(f"  faixa {brl(ini)} a {brl(fim)} @ {pct(al)} = {brl(val)}")
    print(f"INSS empregado......... {brl(ins)}  ({pct(ins/base if base else 0)} efetivo)")
    print(f"Base IRRF ({a.regime_ir})... {brl(base_ir)}")
    print(f"IRRF................... {brl(ir)}")
    outros = a.outros_descontos
    print(f"Outros descontos....... {brl(outros)}")
    liq = base - ins - ir - outros
    print(f"LIQUIDO................ {brl(liq)}")
    if base and (ins + ir + outros) / base > 0.70:
        print("ALERTA: descontos acima de 70% da remuneracao - verificar limite legal.")


def cmd_custo(a, tab):
    sal = a.salario
    base = sal + a.variaveis
    enc = encargos_patronais(base, tab)
    prov_13 = base / 12
    prov_fer = (base / 12) * (4 / 3)
    prov_base = prov_13 + prov_fer
    enc_prov = encargos_patronais(prov_base, tab)["_total"]
    multa40 = base * tab["empresa"]["fgts"] * 0.40 * a.turnover
    beneficios = a.beneficios
    total = base + enc["_total"] + prov_base + enc_prov + multa40 + beneficios
    linhas = [
        ("Salario + variaveis", base),
        *[(k, v) for k, v in enc.items() if not k.startswith("_")],
        ("Provisao 13o (1/12)", prov_13),
        ("Provisao ferias + 1/3", prov_fer),
        ("Encargos sobre provisoes", enc_prov),
        (f"Multa 40% FGTS (turnover {pct(a.turnover)})", multa40),
        ("Beneficios/indiretos", beneficios),
    ]
    print(f"Tabela data-base: {tab['data_base']}")
    print(f"{'Componente':<40}{'Valor':>16}{'% s/ salario':>16}")
    for nome, val in linhas:
        print(f"{nome:<40}{brl(val):>16}{pct(val/sal) if sal else '-':>16}")
    print("-" * 72)
    print(f"{'CUSTO-EMPRESA TOTAL':<40}{brl(total):>16}{pct(total/sal) if sal else '-':>16}")
    if sal:
        fenc = total / sal
        print("\nFator de encargos (Fenc) = " + num(fenc, 4))
        if not 1.65 <= fenc <= 2.10:
            print("ATENCAO: Fenc fora da faixa de plausibilidade 1,65-2,10 da construcao civil. Revisar componentes.")
    if enc["_obs"]:
        print(enc["_obs"])
    print(f"Custo-hora ({a.horas_mes:.0f}h/mes) = {brl(total / a.horas_mes)}")


def cmd_equilibrio(a, _tab=None):
    den = 1 - a.custos_diretos - a.indiretos - a.tributos - a.margem
    print(f"Custo total de MDO..... {brl(a.custo_folha)}")
    print(f"Denominador............ 1 - {pct(a.custos_diretos)} - {pct(a.indiretos)} - {pct(a.tributos)} - {pct(a.margem)} = {pct(den)}")
    if den <= 0:
        print("RESULTADO: modelo nao fecha. A estrutura de custo consome 100% da receita.")
        print("Conclusao: nao existe faturamento que cubra a folha com essas premissas - revisar margem/estrutura.")
        return
    fat = a.custo_folha / den
    print(f"FATURAMENTO MINIMO..... {brl(fat)}")
    if a.faturamento_realizado is not None:
        gap = a.faturamento_realizado - fat
        print(f"Realizado.............. {brl(a.faturamento_realizado)}")
        print(f"Gap.................... {brl(gap)} ({pct(gap/fat)})")
        print("Situacao: " + ("COBRE a folha" if gap >= 0 else "NAO cobre a folha"))


def cmd_tarefa(a, _tab=None):
    if a.valor_unitario <= 0:
        sys.exit("PENDENCIA: valor unitario da tarefa deve ser > 0.")
    m2 = a.valor_alvo / a.valor_unitario
    print(f"Valor-alvo............. {brl(a.valor_alvo)}")
    print(f"Valor unitario......... {brl(a.valor_unitario)}/m2  [confirmar: bruto ou liquido de encargos]")
    print("m2 necessarios......... " + num(m2))
    print("[confirmar: metrica de m2 de PAREDE (uma face) ou de PLACA (duas faces)]")
    if a.dias:
        dia = m2 / a.dias
        print("m2/dia necessarios..... " + num(dia))
        if a.produtividade:
            equipe = dia / a.produtividade
            print("Equipe necessaria...... " + num(equipe) + " homens")
            print("(produtividade de referencia informada: " + num(a.produtividade)
                  + " m2/homem-dia - confirmar com historico da obra)")


def main():
    p = argparse.ArgumentParser(description="Calculadora de folha - DP construcao civil")
    p.add_argument("--tabelas", default=PADRAO)
    sub = p.add_subparsers(dest="cmd", required=True)

    l = sub.add_parser("liquido", help="INSS + IRRF + liquido do colaborador")
    l.add_argument("--salario", type=float, required=True)
    l.add_argument("--base", type=float, default=None, help="base de calculo se != salario")
    l.add_argument("--dependentes", type=int, default=0)
    l.add_argument("--pensao", type=float, default=0.0)
    l.add_argument("--outros-descontos", type=float, default=0.0)
    l.add_argument("--regime-ir", choices=["progressivo", "simplificado"], default="progressivo")

    c = sub.add_parser("custo", help="custo-empresa e fator de encargos (Fenc)")
    c.add_argument("--salario", type=float, required=True)
    c.add_argument("--variaveis", type=float, default=0.0, help="HE, adicionais, producao")
    c.add_argument("--beneficios", type=float, default=0.0, help="VT liquido, VR, saude, EPI, alojamento")
    c.add_argument("--turnover", type=float, default=0.0, help="taxa mensal de desligamento, ex 0.05")
    c.add_argument("--horas-mes", type=float, default=220.0)

    e = sub.add_parser("equilibrio", help="faturamento minimo para cobrir a folha")
    e.add_argument("--custo-folha", type=float, required=True)
    e.add_argument("--custos-diretos", type=float, required=True, help="% nao-MDO sobre receita, ex 0.30")
    e.add_argument("--indiretos", type=float, required=True)
    e.add_argument("--tributos", type=float, required=True)
    e.add_argument("--margem", type=float, required=True)
    e.add_argument("--faturamento-realizado", type=float, default=None)

    t = sub.add_parser("tarefa", help="m2 necessarios e produtividade por tarefa")
    t.add_argument("--valor-alvo", type=float, required=True)
    t.add_argument("--valor-unitario", type=float, required=True)
    t.add_argument("--dias", type=float, default=None, help="dias uteis disponiveis")
    t.add_argument("--produtividade", type=float, default=None, help="m2/homem-dia de referencia")

    a = p.parse_args()
    if a.cmd in ("liquido", "custo"):
        tab = carrega(a.tabelas)
        (cmd_liquido if a.cmd == "liquido" else cmd_custo)(a, tab)
    elif a.cmd == "equilibrio":
        cmd_equilibrio(a)
    else:
        cmd_tarefa(a)


if __name__ == "__main__":
    main()
