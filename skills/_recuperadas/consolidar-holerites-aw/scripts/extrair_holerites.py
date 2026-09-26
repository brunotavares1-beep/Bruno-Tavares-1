#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrai dados de Recibos de Pagamento (holerites) em PDF, gerados pelo TOTVS RM,
a partir de uma ou mais pastas/ZIPs mensais, e consolida tudo em uma única
planilha Excel (Resumo + Rubricas + Dicionario_Rubricas + Erros).

Uso:
    python extrair_holerites.py --input pasta1.zip --input pasta2 --output saida.xlsx
    python extrair_holerites.py --input pasta_com_8_zips/ --output saida.xlsx --incluir-invalidos

O script NUNCA inventa dado: se um recibo não puder ser lido com segurança
(nome/CPF ausente ou totais ausentes), ele vai para a aba "Erros" e, por
padrão, NÃO entra em Resumo/Rubricas.
"""

import argparse
import os
import re
import sys
import zipfile
import shutil
import tempfile
from collections import Counter, defaultdict

import pdfplumber
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]
MESES_RE = "|".join(MESES)

NUM_RE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")

RESUMO_COLS = [
    "Competência", "Empresa", "CNPJ", "Filial (Descrição)", "Código Funcionário",
    "Nome do Funcionário", "CPF", "Cargo", "CBO", "Departamento", "Filial (Código)",
    "Admissão", "Tipo Contratação", "Salário Base", "Sal. Contr. INSS",
    "Total Vencimentos", "Total Descontos", "Valor Líquido", "Base Cálc. FGTS",
    "FGTS do Mês", "Base Cálc. IRRF", "Faixa IRRF (%)", "Arquivo Origem",
]
RUBRICAS_COLS = [
    "Competência", "Código Funcionário", "Nome do Funcionário", "Código Rubrica",
    "Descrição Rubrica", "Natureza", "Referência", "Vencimento", "Desconto",
    "Arquivo Origem",
]
DICIONARIO_COLS = [
    "Código Rubrica", "Descrição (mais frequente)", "Natureza Observada",
    "Nº Ocorrências", "Descrições Alternativas Vistas",
]
ERROS_COLS = ["Arquivo", "Erro"]


def to_float(s):
    if s is None:
        return None
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# 1) Localizar PDFs (descompactando ZIPs, inclusive aninhados)
# ---------------------------------------------------------------------------

def extrair_zips_recursivo(root_dir):
    """Descompacta todo .zip encontrado dentro de root_dir, recursivamente,
    apagando o .zip original após extrair (para não reprocessá-lo como PDF)."""
    mudou = True
    while mudou:
        mudou = False
        for dirpath, _, filenames in os.walk(root_dir):
            for fn in filenames:
                if fn.lower().endswith(".zip"):
                    zpath = os.path.join(dirpath, fn)
                    destino = os.path.join(dirpath, os.path.splitext(fn)[0])
                    os.makedirs(destino, exist_ok=True)
                    try:
                        with zipfile.ZipFile(zpath) as z:
                            z.extractall(destino)
                    except zipfile.BadZipFile:
                        pass
                    os.remove(zpath)
                    mudou = True


def coletar_pdfs(input_paths, workdir):
    """Copia/descompacta todas as entradas de input_paths para workdir e
    retorna a lista de caminhos de PDF encontrados (ordem estável)."""
    for p in input_paths:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            destino = os.path.join(workdir, os.path.basename(p.rstrip("/")))
            shutil.copytree(p, destino, dirs_exist_ok=True)
        elif zipfile.is_zipfile(p):
            destino = os.path.join(workdir, os.path.splitext(os.path.basename(p))[0])
            os.makedirs(destino, exist_ok=True)
            with zipfile.ZipFile(p) as z:
                z.extractall(destino)
        elif p.lower().endswith(".pdf"):
            shutil.copy2(p, workdir)
        else:
            print(f"[aviso] entrada ignorada (não é pasta, zip nem pdf): {p}")

    extrair_zips_recursivo(workdir)

    pdfs = []
    for dirpath, _, filenames in os.walk(workdir):
        for fn in sorted(filenames):
            if fn.lower().endswith(".pdf"):
                pdfs.append(os.path.join(dirpath, fn))
    return sorted(pdfs)


# ---------------------------------------------------------------------------
# 2) Reconstrução de linhas "limpas" de cada página (só texto não rotacionado)
# ---------------------------------------------------------------------------

def linhas_da_pagina(page, tol=3.0):
    """Agrupa palavras em linhas por proximidade vertical, ignorando texto
    rotacionado (o aviso 'Declaro ter recebido...' / 'Assinatura do
    Funcionário' impresso girado, que atrapalha a leitura em colunas)."""
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False,
                                extra_attrs=["upright"])
    words = [w for w in words if w.get("upright", True)]
    words.sort(key=lambda w: (w["top"], w["x0"]))

    linhas = []  # cada item: {"top": float, "words": [...]}
    for w in words:
        colocado = False
        for linha in linhas:
            if abs(linha["top"] - w["top"]) <= tol:
                linha["words"].append(w)
                n = len(linha["words"])
                linha["top"] = (linha["top"] * (n - 1) + w["top"]) / n
                colocado = True
                break
        if not colocado:
            linhas.append({"top": w["top"], "words": [w]})

    linhas.sort(key=lambda l: l["top"])
    for linha in linhas:
        linha["words"].sort(key=lambda w: w["x0"])
        linha["text"] = " ".join(w["text"] for w in linha["words"])
    return linhas


# ---------------------------------------------------------------------------
# 3) Classificação: é um Recibo de Pagamento (holerite)?
# ---------------------------------------------------------------------------

def classificar_arquivo(texto_completo):
    """Retorna (eh_holerite: bool, motivo_se_nao: str|None)."""
    marcadores_holerite = ["Total de Vencimentos", "Total de Descontos", "Valor Líquido"]
    if all(m in texto_completo for m in marcadores_holerite) and "CNPJ" in texto_completo:
        return True, None

    if "Comunicado sobre Desconto de Empréstimo" in texto_completo or \
       "Crédito do Trabalhador" in texto_completo:
        return False, ("Arquivo não é um Recibo de Pagamento — parece ser um Comunicado de "
                        "Desconto de Empréstimo Consignado (Crédito do Trabalhador). "
                        "Verifique se foi incluído por engano na pasta.")

    if re.search(r"Cart[aã]o de Ponto|Espelho de Ponto", texto_completo, re.IGNORECASE):
        return False, ("Arquivo não é um Recibo de Pagamento — parece ser um Cartão/Espelho "
                        "de Ponto. Verifique se foi incluído por engano na pasta.")

    return False, ("Arquivo não reconhecido como Recibo de Pagamento — não contém os totais "
                    "esperados (Total de Vencimentos / Total de Descontos / Valor Líquido).")


# ---------------------------------------------------------------------------
# 4) Parsing de um recibo já classificado como holerite
# ---------------------------------------------------------------------------

def classificar_coluna(x0, col_ref, col_venc, col_desc):
    """Decide se um valor numérico pertence à coluna Referência, Vencimentos
    ou Descontos, pela proximidade ao x0 de cada cabeçalho de coluna."""
    dists = {
        "referencia": abs(x0 - col_ref),
        "vencimento": abs(x0 - col_venc),
        "desconto": abs(x0 - col_desc),
    }
    return min(dists, key=dists.get)


def parse_holerite(pdf_path):
    """Extrai (resumo_dict, [rubrica_dict, ...]) de um PDF de Recibo de
    Pagamento. Lança ValueError com uma mensagem clara se campos essenciais
    não puderem ser lidos com segurança."""
    nome_arquivo = os.path.basename(pdf_path)

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        linhas = linhas_da_pagina(page)

        # localizar cabeçalho da tabela de rubricas para pegar posição das colunas
        header_idx = None
        col_ref = col_venc = col_desc = None
        for i, linha in enumerate(linhas):
            palavras = {w["text"]: w["x0"] for w in linha["words"]}
            if "Referência" in palavras and "Vencimentos" in palavras and "Descontos" in palavras:
                header_idx = i
                col_ref = palavras["Referência"]
                col_venc = palavras["Vencimentos"]
                col_desc = palavras["Descontos"]
                break
        if header_idx is None:
            raise ValueError("Não localizei o cabeçalho da tabela de rubricas "
                              "(Código/Descrição/Referência/Vencimentos/Descontos).")

        texto_ate_header = "\n".join(l["text"] for l in linhas[:header_idx])

        # --- Empresa / CNPJ / Filial (Descrição) ---
        m = re.search(r"CNPJ:\s*([\d./-]+).*?CC:\s*(.+?)\s+Folha Mensal", texto_ate_header)
        if not m:
            raise ValueError("Não localizei CNPJ / Filial (CC:) / 'Folha Mensal' no cabeçalho.")
        cnpj = m.group(1).strip()
        filial_desc = m.group(2).strip()
        empresa = linhas[0]["text"].strip()

        # --- Filial (Código) / Tipo Contratação / Competência ---
        m = re.search(
            rf"^(\d+)\s*-\s*.+?\s+([A-Za-zÀ-ÿçÇãÃéÉíÍóÓúÚ]+)\s+((?:{MESES_RE})\s+de\s+\d{{4}})$",
            texto_ate_header, re.MULTILINE,
        )
        if not m:
            raise ValueError("Não localizei a linha 'Filial - Empresa  Tipo Contratação  Mês de Ano'.")
        filial_codigo = m.group(1).strip()
        tipo_contratacao = m.group(2).strip()
        competencia = m.group(3).strip()

        # --- Código Funcionário / Nome / CPF / CBO / Departamento ---
        m = re.search(
            r"^(\d+)\s+(.+?)\s+(\d{3}\.\d{3}\.\d{3}-\d{2})\s+(\d+)\s+(\d+)\s+(\d+)$",
            texto_ate_header, re.MULTILINE,
        )
        if not m:
            raise ValueError("Não localizei a linha de dados do funcionário "
                              "(Código / Nome / CPF / CBO / Departamento / Filial).")
        cod_func = m.group(1).strip()
        nome_func = m.group(2).strip()
        cpf = m.group(3).strip()
        cbo = m.group(4).strip()
        departamento = m.group(5).strip()

        # --- Cargo / Admissão ---
        m = re.search(r"^(.+?)\s+Admiss[aã]o:\s+(\d{2}/\d{2}/\d{4})$", texto_ate_header, re.MULTILINE)
        if not m:
            raise ValueError("Não localizei Cargo / Admissão.")
        cargo = m.group(1).strip()
        admissao = m.group(2).strip()

        if not nome_func or not cpf:
            raise ValueError("Nome do Funcionário ou CPF vieram vazios após o parsing.")

        # --- Rubricas (até a próxima linha 'Total de Vencimentos ...') ---
        rubricas = []
        i = header_idx + 1
        fim_idx = None
        while i < len(linhas):
            texto = linhas[i]["text"].strip()
            if texto.startswith("Total de Vencimentos"):
                fim_idx = i
                break
            m_cod = re.match(r"^(\d+)\s+(.*)$", texto)
            if m_cod:
                codigo_rubrica = m_cod.group(1)
                resto = m_cod.group(2)
                valores = [w for w in linhas[i]["words"] if NUM_RE.match(w["text"])]
                descricao = resto
                for v in valores:
                    descricao = descricao.replace(v["text"], "").strip()
                descricao = re.sub(r"\s{2,}", " ", descricao).strip(" -")

                referencia = vencimento = desconto = None
                for v in valores:
                    coluna = classificar_coluna(v["x0"], col_ref, col_venc, col_desc)
                    val = to_float(v["text"])
                    if coluna == "referencia":
                        referencia = val
                    elif coluna == "vencimento":
                        vencimento = val
                    elif coluna == "desconto":
                        desconto = val

                if vencimento is not None:
                    natureza = "Provento"
                elif desconto is not None:
                    natureza = "Desconto"
                else:
                    natureza = "Indefinido"

                rubricas.append({
                    "Competência": competencia,
                    "Código Funcionário": cod_func,
                    "Nome do Funcionário": nome_func,
                    "Código Rubrica": codigo_rubrica,
                    "Descrição Rubrica": descricao,
                    "Natureza": natureza,
                    "Referência": referencia,
                    "Vencimento": vencimento,
                    "Desconto": desconto,
                    "Arquivo Origem": nome_arquivo,
                })
            i += 1

        if fim_idx is None:
            raise ValueError("Não localizei a linha 'Total de Vencimentos / Total de Descontos'.")

        # valores totais (linha logo após o rótulo, ignorando linhas em branco)
        j = fim_idx + 1
        while j < len(linhas) and not linhas[j]["text"].strip():
            j += 1
        m = re.match(r"^([\d.,]+)\s+([\d.,]+)$", linhas[j]["text"].strip())
        if not m:
            raise ValueError("Não consegui ler os valores de Total de Vencimentos / Total de Descontos.")
        total_venc = to_float(m.group(1))
        total_desc = to_float(m.group(2))

        # Valor Líquido
        valor_liquido = None
        for k in range(j, min(j + 6, len(linhas))):
            m = re.search(r"Valor\s+L[ií]quido\s+([\d.,]+)", linhas[k]["text"])
            if m:
                valor_liquido = to_float(m.group(1))
                break
        if valor_liquido is None:
            raise ValueError("Não localizei o Valor Líquido.")

        # Linha de rodapé: Salário Base / Sal. Contr. INSS / Base Cálc. FGTS / FGTS do Mês / Base Cálc. IRRF / Faixa IRRF
        salario_base = sal_contr_inss = base_fgts = fgts_mes = base_irrf = faixa_irrf = None
        for k in range(j, min(j + 10, len(linhas))):
            if "Salário" in linhas[k]["text"] and "Base" in linhas[k]["text"]:
                # valores ficam na(s) próxima(s) linha(s) não vazia(s)
                m2 = k + 1
                while m2 < len(linhas) and not linhas[m2]["text"].strip():
                    m2 += 1
                nums = re.findall(r"[\d.]+,\d{2}", linhas[m2]["text"])
                if len(nums) >= 6:
                    (salario_base, sal_contr_inss, base_fgts,
                     fgts_mes, base_irrf, faixa_irrf) = [to_float(n) for n in nums[:6]]
                break

        if salario_base is None:
            raise ValueError("Não localizei a linha de Salário Base / INSS / FGTS / IRRF.")

    resumo = {
        "Competência": competencia,
        "Empresa": empresa,
        "CNPJ": cnpj,
        "Filial (Descrição)": filial_desc,
        "Código Funcionário": cod_func,
        "Nome do Funcionário": nome_func,
        "CPF": cpf,
        "Cargo": cargo,
        "CBO": cbo,
        "Departamento": departamento,
        "Filial (Código)": filial_codigo,
        "Admissão": admissao,
        "Tipo Contratação": tipo_contratacao,
        "Salário Base": salario_base,
        "Sal. Contr. INSS": sal_contr_inss,
        "Total Vencimentos": total_venc,
        "Total Descontos": total_desc,
        "Valor Líquido": valor_liquido,
        "Base Cálc. FGTS": base_fgts,
        "FGTS do Mês": fgts_mes,
        "Base Cálc. IRRF": base_irrf,
        "Faixa IRRF (%)": faixa_irrf,
        "Arquivo Origem": nome_arquivo,
    }
    return resumo, rubricas


# ---------------------------------------------------------------------------
# 5) Orquestração
# ---------------------------------------------------------------------------

def processar(pdf_paths, incluir_invalidos=False):
    resumo_rows, rubricas_rows, erros_rows = [], [], []

    for pdf_path in pdf_paths:
        nome_arquivo = os.path.basename(pdf_path)
        try:
            with pdfplumber.open(pdf_path) as pdf:
                texto_completo = "\n".join((p.extract_text() or "") for p in pdf.pages)
        except Exception as e:
            erros_rows.append({"Arquivo": nome_arquivo, "Erro": f"Falha ao abrir o PDF: {e}"})
            continue

        eh_holerite, motivo = classificar_arquivo(texto_completo)
        if not eh_holerite:
            sufixo = "" if incluir_invalidos else " [EXCLUÍDO da base — use --incluir-invalidos para trazê-lo mesmo incompleto]"
            erros_rows.append({"Arquivo": nome_arquivo, "Erro": motivo + sufixo})
            continue

        try:
            resumo, rubricas = parse_holerite(pdf_path)
        except Exception as e:
            erros_rows.append({
                "Arquivo": nome_arquivo,
                "Erro": f"Recibo de Pagamento reconhecido, mas falhou ao extrair os dados: {e}",
            })
            continue

        resumo_rows.append(resumo)
        rubricas_rows.extend(rubricas)

    # Dicionário de Rubricas
    agregada = defaultdict(lambda: {"desc": Counter(), "nat": Counter()})
    for r in rubricas_rows:
        agregada[r["Código Rubrica"]]["desc"][r["Descrição Rubrica"]] += 1
        agregada[r["Código Rubrica"]]["nat"][r["Natureza"]] += 1

    dicionario_rows = []
    for codigo, info in sorted(agregada.items(), key=lambda kv: int(kv[0]) if kv[0].isdigit() else 0):
        desc_mais_freq, _ = info["desc"].most_common(1)[0]
        nat_mais_freq, _ = info["nat"].most_common(1)[0]
        total_ocorrencias = sum(info["desc"].values())
        alternativas = [d for d in info["desc"] if d != desc_mais_freq]
        dicionario_rows.append({
            "Código Rubrica": codigo,
            "Descrição (mais frequente)": desc_mais_freq,
            "Natureza Observada": nat_mais_freq,
            "Nº Ocorrências": total_ocorrencias,
            "Descrições Alternativas Vistas": "; ".join(alternativas) if alternativas else None,
        })

    return resumo_rows, rubricas_rows, dicionario_rows, erros_rows


# ---------------------------------------------------------------------------
# 6) Geração do Excel
# ---------------------------------------------------------------------------

def autofit(ws, max_width=60):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[letter].width = min(max(length + 2, 10), max_width)


def escrever_aba_tabela(wb, nome_aba, colunas, linhas, nome_tabela, style="TableStyleMedium2"):
    ws = wb.create_sheet(nome_aba)
    ws.append(colunas)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in linhas:
        ws.append([row.get(c) for c in colunas])

    n_rows = len(linhas) + 1
    n_cols = len(colunas)
    if n_rows > 1:
        ref = f"A1:{get_column_letter(n_cols)}{n_rows}"
        tabela = Table(displayName=nome_tabela, ref=ref)
        tabela.tableStyleInfo = TableStyleInfo(
            name=style, showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(tabela)
    ws.freeze_panes = "A2"
    autofit(ws)
    return ws


def gerar_excel(resumo_rows, rubricas_rows, dicionario_rows, erros_rows, output_path,
                 fontes_processadas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Leia-me ---
    ws = wb.create_sheet("Leia-me")
    competencias = sorted(set(r["Competência"] for r in resumo_rows))
    linhas_texto = [
        "BASE DE HOLERITES — AW EMPREITEIRA",
        "",
        "Estrutura do arquivo:",
        "• Resumo: 1 linha por funcionário/competência (dados cadastrais e totais do recibo).",
        "• Rubricas: 1 linha por lançamento (código, descrição, referência, vencimento, desconto). "
        "Use esta aba para tabela dinâmica, Power BI, PROCX e SOMASES.",
        "• Dicionario_Rubricas: catálogo dos códigos de rubrica observados, para conferência e "
        "padronização de descrições ao longo dos recibos.",
        "• Erros: arquivos que o script não conseguiu ler, ou que não são Recibo de Pagamento "
        "(ex.: Comunicado de Empréstimo Consignado, Cartão/Espelho de Ponto). Por padrão, esses "
        "arquivos NÃO entram no Resumo/Rubricas — ficam só listados aqui, para não sujar a base.",
        "",
        f"Competências incluídas: {', '.join(competencias) if competencias else '(nenhuma)'}",
        f"Recibos processados com sucesso: {len(resumo_rows)}",
        f"Arquivos em Erros/Excluídos: {len(erros_rows)}",
        f"Total de arquivos PDF analisados: {fontes_processadas}",
    ]
    for t in linhas_texto:
        ws.append([t])
    ws.column_dimensions["A"].width = 110
    ws["A1"].font = Font(bold=True, size=13)

    escrever_aba_tabela(wb, "Resumo", RESUMO_COLS, resumo_rows, "tbResumo")
    escrever_aba_tabela(wb, "Rubricas", RUBRICAS_COLS, rubricas_rows, "tbRubricas")
    escrever_aba_tabela(wb, "Dicionario_Rubricas", DICIONARIO_COLS, dicionario_rows, "tbDicionario")

    ws = wb.create_sheet("Erros")
    ws.append(ERROS_COLS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in erros_rows:
        ws.append([row.get(c) for c in ERROS_COLS])
    autofit(ws, max_width=100)
    for row in ws.iter_rows(min_row=2, max_col=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    wb.move_sheet("Leia-me", offset=-len(wb.sheetnames))
    wb.save(output_path)


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", action="append", required=True,
                     help="Pasta, .zip ou .pdf. Pode repetir para várias competências.")
    ap.add_argument("--output", required=True, help="Caminho do .xlsx de saída.")
    ap.add_argument("--incluir-invalidos", action="store_true",
                     help="Mesmo assim, tenta listar em Erros arquivos que não são holerite "
                          "(padrão: já lista; a flag só muda o texto da observação).")
    args = ap.parse_args()

    with tempfile.TemporaryDirectory() as workdir:
        pdfs = coletar_pdfs(args.input, workdir)
        print(f"[info] {len(pdfs)} arquivo(s) PDF encontrado(s) para análise.")
        resumo_rows, rubricas_rows, dicionario_rows, erros_rows = processar(
            pdfs, incluir_invalidos=args.incluir_invalidos)

    gerar_excel(resumo_rows, rubricas_rows, dicionario_rows, erros_rows,
                args.output, fontes_processadas=len(pdfs))

    print(f"[ok] Recibos processados: {len(resumo_rows)}")
    print(f"[ok] Lançamentos (Rubricas): {len(rubricas_rows)}")
    print(f"[ok] Códigos de rubrica no dicionário: {len(dicionario_rows)}")
    print(f"[ok] Arquivos em Erros/Excluídos: {len(erros_rows)}")
    for e in erros_rows:
        print(f"     - {e['Arquivo']}: {e['Erro']}")
    print(f"[ok] Planilha gerada em: {args.output}")


if __name__ == "__main__":
    main()
