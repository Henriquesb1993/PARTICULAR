# -*- coding: utf-8 -*-
"""
Gerador da planilha de Conciliacao Cadastral e Financeira - Convenio Medico Hapvida.

Le a base do Hapvida (CSV exportado do convenio) e monta um workbook .xlsx
totalmente orientado a formulas: ao substituir as bases nas abas HAPVIDA e
SISTEMA, todas as demais abas se recalculam sozinhas.

Uso:  python3 gerar_planilha.py [caminho_do_csv] [caminho_de_saida]
"""
import csv
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

# --------------------------------------------------------------------------
# 1. PARAMETROS DE NEGOCIO  (fonte: especificacao do RH, itens 9 a 12)
# --------------------------------------------------------------------------
VLR_TITULAR = 151.73          # 100% custeado pela empresa
VLR_AGREGADO = 544.95         # por agregado, custeado pelo colaborador
COMBO = [                     # faixa pela QUANTIDADE TOTAL de dependentes
    (0, 0.00),
    (1, 239.93),
    (2, 446.62),
    (3, 579.83),
    (4, 663.57),
    (5, 846.86),              # 5 ou mais
]
IDADE_LIMITE_DEP = 25         # acima disso o dependente vira agregado

# Capacidades das areas de dados (linhas de formula pre-preenchidas)
N_HAP = 9700                  # base atual: 9.181 linhas
N_SIS = 10000
N_TIT = 6200                  # base atual: 5.861 titulares
N_CONC_B = 3000               # bloco "somente sistema" da CONCILIACAO
N_ALERTA = 3000

FONTE = "Arial"

# --------------------------------------------------------------------------
# 2. PALETA / ESTILOS
# --------------------------------------------------------------------------
C_NAVY   = "1F3A5F"
C_NAVY_D = "12263A"
C_TEAL   = "0E6F6E"
C_CINZA  = "F2F4F7"
C_CINZA2 = "E4E8EE"
C_BORDA  = "B7C0CD"
C_AZUL_TXT = "0000FF"     # entradas digitaveis
C_VERDE  = "1E8449"; F_VERDE  = "D5F5E3"
C_AMAR   = "9A7D0A"; F_AMAR   = "FCF3CF"
C_VERM   = "B03A2E"; F_VERM   = "FADBD8"
C_LARANJ = "AF601A"; F_LARANJ = "FDEBD0"
C_AZUL   = "21618C"; F_AZUL   = "D6EAF8"

MOEDA = '"R$" #,##0.00'
MOEDA0 = '"R$" #,##0'
INT = '#,##0'
DATA = 'DD/MM/YYYY'

thin = Side(style="thin", color=C_BORDA)
BORDA = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(**kw):
    kw.setdefault("name", FONTE)
    return Font(**kw)


def fill(hexcolor):
    return PatternFill("solid", fgColor=hexcolor)


def header_row(ws, row, labels, start_col=1, bg=C_NAVY, h=30):
    """Escreve uma linha de cabecalho formatada."""
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = f(bold=True, color="FFFFFF", size=9)
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA
    ws.row_dimensions[row].height = h


def titulo(ws, cell, texto, span=1, size=14, bg=C_NAVY_D, color="FFFFFF"):
    ws[cell] = texto
    ws[cell].font = f(bold=True, size=size, color=color)
    ws[cell].fill = fill(bg)
    ws[cell].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    r = ws[cell].row
    c0 = ws[cell].column
    if span > 1:
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + span - 1)
        for i in range(1, span):
            ws.cell(row=r, column=c0 + i).fill = fill(bg)
    ws.row_dimensions[r].height = max(22, size + 10)


def larguras(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


# --------------------------------------------------------------------------
# 3. NORMALIZACAO DE NOME (formula viva, para casar as duas bases)
# --------------------------------------------------------------------------
_ACENTOS = [("Á", "A"), ("À", "A"), ("Â", "A"), ("Ã", "A"), ("Ä", "A"),
            ("É", "E"), ("Ê", "E"), ("Í", "I"), ("Ó", "O"), ("Ô", "O"),
            ("Õ", "O"), ("Ú", "U"), ("Ü", "U"), ("Ç", "C"), (".", ""), ("-", " ")]


def data_formula(ref):
    """Converte `ref` em data serial aceitando data real OU texto DD/MM/AAAA.

    DATA.VALOR/DATEVALUE depende do idioma do Excel: em locale en-US
    "30/09/2002" vira erro e "05/09/2002" vira 9 de maio. Aqui a data e
    montada explicitamente a partir das duas barras, entao o resultado e o
    mesmo em qualquer maquina.
    """
    return ('IF({r}="","",IF(ISNUMBER({r}),{r},IFERROR(DATE('
            'VALUE(MID({r},FIND("/",{r},FIND("/",{r})+1)+1,4)),'
            'VALUE(MID({r},FIND("/",{r})+1,FIND("/",{r},FIND("/",{r})+1)-FIND("/",{r})-1)),'
            'VALUE(LEFT({r},FIND("/",{r})-1))),"")))').format(r=ref)


def norm_formula(ref):
    """Devolve a formula que normaliza o nome da celula `ref` (sem o '=')."""
    expr = "UPPER(TRIM({}))".format(ref)
    for a, b in _ACENTOS:
        expr = 'SUBSTITUTE({},"{}","{}")'.format(expr, a, b)
    return "TRIM({})".format(expr)


# --------------------------------------------------------------------------
# 4. LEITURA DA BASE HAPVIDA
# --------------------------------------------------------------------------
def parse_data(s):
    s = (s or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def parse_num(s):
    s = (s or "").strip().replace(".", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def ler_hapvida(caminho):
    linhas = []
    with open(caminho, encoding="utf-8-sig", newline="") as fh:
        r = csv.reader(fh, delimiter=";")
        cab = None
        for row in r:
            if not any((c or "").strip() for c in row):
                continue
            vals = [(c or "").strip() for c in row]
            if cab is None:
                if vals[0].upper() == "CONTRATO":
                    cab = vals
                continue
            linhas.append(dict(zip(cab, vals)))
    return linhas


# ==========================================================================
# ABA: PARAMETROS
# ==========================================================================
def aba_parametros(wb):
    ws = wb.create_sheet("PARÂMETROS")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 42, "C": 18, "D": 4, "E": 30, "F": 18, "G": 60})

    titulo(ws, "B2", "PARÂMETROS DO CONVÊNIO — HAPVIDA", span=6, size=15)
    ws["B3"] = "Células em AZUL são editáveis. Todo o restante da pasta lê estes valores."
    ws["B3"].font = f(italic=True, size=9, color="555555")

    def par(row, label, valor, fmt=None, nota=None):
        ws.cell(row=row, column=2, value=label).font = f(bold=True, size=10)
        c = ws.cell(row=row, column=3, value=valor)
        c.font = f(color=C_AZUL_TXT, bold=True, size=10)
        c.fill = fill("FFFFCC")
        c.border = BORDA
        c.alignment = Alignment(horizontal="center")
        if fmt:
            c.number_format = fmt
        if nota:
            ws.cell(row=row, column=5, value=nota).font = f(size=9, color="555555")
        return c

    titulo(ws, "B5", "DATA-BASE E REGRAS", span=6, size=11, bg=C_TEAL)
    par(6, "Data-base para cálculo das idades", "=TODAY()", DATA,
        "Troque por uma data fixa se quiser congelar a análise.")
    par(7, "Idade limite do dependente (anos)", IDADE_LIMITE_DEP, INT,
        "Acima disso o dependente passa a ser tratado como AGREGADO.")
    par(8, "Aplicar a regra de idade ao CÔNJUGE?", "NÃO", None,
        "Padrão NÃO: cônjuge é dependente vitalício, não vira agregado por idade.")

    titulo(ws, "B10", "VALORES DO CONTRATO", span=6, size=11, bg=C_TEAL)
    par(11, "Valor do TITULAR (mensal)", VLR_TITULAR, MOEDA, "100% custeado pela EMPRESA.")
    par(12, "Valor de CADA AGREGADO (mensal)", VLR_AGREGADO, MOEDA,
        "Responsabilidade do COLABORADOR. Soma-se por agregado.")

    titulo(ws, "B14", "COMBO FAMILIAR — VALOR PELA FAIXA DE DEPENDENTES", span=6, size=11, bg=C_TEAL)
    header_row(ws, 15, ["Qtd. de dependentes", "Valor do Combo"], start_col=2, h=24)
    for i, (q, v) in enumerate(COMBO):
        r = 16 + i
        c1 = ws.cell(row=r, column=2, value=q)
        c1.font = f(color=C_AZUL_TXT, bold=True, size=10)
        c1.alignment = Alignment(horizontal="center")
        c1.border = BORDA
        c1.fill = fill("FFFFCC")
        c2 = ws.cell(row=r, column=3, value=v)
        c2.number_format = MOEDA
        c2.font = f(color=C_AZUL_TXT, bold=True, size=10)
        c2.border = BORDA
        c2.fill = fill("FFFFCC")
    ws.cell(row=16 + len(COMBO), column=2,
            value="A última faixa (5) vale para 5 OU MAIS dependentes. "
                  "O valor NUNCA é multiplicado pela quantidade.").font = f(italic=True, size=9, color=C_VERM)
    ws.merge_cells(start_row=16 + len(COMBO), start_column=2, end_row=16 + len(COMBO), end_column=6)

    ws["C6"].comment = Comment(
        "Fonte: informado pelo RH. Padrão =HOJE(). Toda idade da pasta é recalculada "
        "a partir da data de nascimento com esta data-base — a coluna 'idade' que vem "
        "no arquivo do Hapvida é apenas um retrato da data da extração.", "Conciliação")
    ws["C11"].comment = Comment("Fonte: especificação do RH — contrato Combo Familiar Hapvida.", "Conciliação")

    titulo(ws, "B24", "RATEIO", span=6, size=11, bg=C_TEAL)
    ws["B25"] = "Custeado pela EMPRESA"
    ws["C25"] = "Titular"
    ws["B26"] = "Custeado pelo COLABORADOR"
    ws["C26"] = "Combo Familiar + Agregados"
    for r in (25, 26):
        ws.cell(row=r, column=2).font = f(bold=True, size=10)
        ws.cell(row=r, column=3).font = f(size=10)
    ws.merge_cells("C25:E25")
    ws.merge_cells("C26:E26")
    return ws


# ==========================================================================
# ABA: HAPVIDA  (base + motor de cálculo)
# ==========================================================================
HAP_COLS = [
    ("CONTRATO", 22), ("RE (matrícula)", 12), ("BENEFICIÁRIO", 40), ("NOME DA MÃE", 34),
    ("NASCIMENTO", 13), ("INÍCIO", 12), ("IDADE (arquivo)", 10), ("PARENTESCO", 14),
    ("PLANO", 10), ("COBRADO (arquivo)", 14),
    # ---- calculadas ----
    ("Dt. nascimento", 13), ("Idade calculada", 11), ("Tipo", 13), ("Nome normalizado", 34),
    ("Chave RE+Nasc.", 20), ("Chave RE+Nome", 30), ("Idx sistema", 11),
    ("RE existe no sistema?", 12), ("Elegível a agregado?", 12),
    ("Chave duplicidade", 30), ("Duplicado no Hapvida?", 12),
    ("Chave titular", 12), ("1º titular do RE?", 12), ("_cnt", 8), ("Seq. titular", 10),
    ("Chave RE+Tipo", 22), ("Divergência de dados?", 12),
]


def aba_hapvida(wb, dados):
    ws = wb.create_sheet("HAPVIDA")
    header_row(ws, 1, [c[0] for c in HAP_COLS], h=34)
    for i, (_, w) in enumerate(HAP_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    # ---- dados brutos (A:J) ----
    for i, d in enumerate(dados):
        r = i + 2
        ws.cell(row=r, column=1, value=d.get("CONTRATO", ""))
        re_ = d.get("matricula", "").strip()
        ws.cell(row=r, column=2, value=int(re_) if re_.isdigit() else re_)
        ws.cell(row=r, column=3, value=d.get("beneficiario", ""))
        ws.cell(row=r, column=4, value=d.get("nome_mae", ""))
        cn = ws.cell(row=r, column=5, value=parse_data(d.get("nascimento", "")))
        cn.number_format = DATA
        ci = ws.cell(row=r, column=6, value=parse_data(d.get("inicio", "")))
        ci.number_format = DATA
        idd = d.get("idade", "").strip()
        ws.cell(row=r, column=7, value=int(idd) if idd.isdigit() else idd)
        ws.cell(row=r, column=8, value=d.get("parentesco", ""))
        ws.cell(row=r, column=9, value=d.get("plano", ""))
        cc = ws.cell(row=r, column=10, value=parse_num(d.get("cobrado", "")))
        cc.number_format = MOEDA

    # ---- colunas calculadas (K:U) ----
    for r in range(2, N_HAP + 2):
        ws.cell(row=r, column=11, value='=' + data_formula("$E{}".format(r))).number_format = DATA
        ws.cell(row=r, column=12, value='=IF($K{r}="","",DATEDIF($K{r},DATA_BASE,"Y"))'.format(r=r))
        ws.cell(row=r, column=13, value='=IF($B{r}="","",IF(UPPER($H{r})="TITULAR","TITULAR",IF(UPPER($H{r})="AGREGADO","AGREGADO","DEPENDENTE")))'.format(r=r))
        ws.cell(row=r, column=14, value='=IF($C{r}="","",{n})'.format(r=r, n=norm_formula("$C{}".format(r))))
        ws.cell(row=r, column=15, value='=IF($B{r}="","",$B{r}&"|"&IF($K{r}="","?",TEXT($K{r},"ddmmyyyy")))'.format(r=r))
        ws.cell(row=r, column=16, value='=IF($B{r}="","",$B{r}&"|"&$N{r})'.format(r=r))
        ws.cell(row=r, column=17, value=(
            '=IF($B{r}="","",IF(COUNTA(SIS_RE)=0,"",'
            'IFERROR(MATCH($O{r},SIS_CHDT,0),IFERROR(MATCH($P{r},SIS_CHNM,0),""))))').format(r=r))
        ws.cell(row=r, column=18, value=(
            '=IF($B{r}="","",IF(COUNTA(SIS_RE)=0,"NÃO",IF(COUNTIF(SIS_RE,$B{r})>0,"SIM","NÃO")))').format(r=r))
        ws.cell(row=r, column=19, value='=IF($M{r}="","",IF(AND($M{r}="DEPENDENTE",N($L{r})>IDADE_LIMITE,OR(LEFT(UPPER($H{r}),4)<>"CONJ",REGRA_CONJUGE="SIM")),"SIM","NÃO"))'.format(r=r))
        ws.cell(row=r, column=20, value='=IF($B{r}="","",$O{r}&"|"&$N{r})'.format(r=r))
        ws.cell(row=r, column=21, value='=IF($T{r}="","",IF(MATCH($T{r},HAP_CHDUP,0)=ROW()-1,"NÃO","SIM"))'.format(r=r))
        # 1 linha por RE: so a PRIMEIRA linha de titular de cada RE alimenta a aba TITULARES.
        ws.cell(row=r, column=22, value='=IF($M{r}<>"TITULAR","",$B{r}&"")'.format(r=r))
        ws.cell(row=r, column=23, value='=IF($V{r}="","NÃO",IF(MATCH($V{r},HAP_CHTIT,0)=ROW()-1,"SIM","NÃO"))'.format(r=r))
        novo = 'IF($W{r}="SIM",1,0)'.format(r=r)
        if r == 2:
            ws.cell(row=r, column=24, value='=' + novo)
        else:
            ws.cell(row=r, column=24, value='=$X{p}+{n}'.format(p=r - 1, n=novo))
        ws.cell(row=r, column=25, value='=IF($W{r}="SIM",$X{r},"")'.format(r=r))
        ws.cell(row=r, column=26, value='=IF($B{r}="","",$B{r}&"|"&$M{r})'.format(r=r))
        ws.cell(row=r, column=27, value=(
            '=IF($Q{r}="","NÃO",IF(OR($N{r}<>INDEX(SIS_NOMENORM,$Q{r}),'
            'AND(N($K{r})>0,N(INDEX(SIS_DTN,$Q{r}))>0,N($K{r})<>N(INDEX(SIS_DTN,$Q{r}))),'
            'LEFT(UPPER($H{r}),4)<>LEFT(UPPER(INDEX(SIS_GRAU,$Q{r})),4)),"SIM","NÃO"))').format(r=r))

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:W{}".format(N_HAP + 1)
    ws.column_dimensions.group("X", "Z", hidden=True)
    ws.auto_filter.ref = "A1:AA{}".format(N_HAP + 1)
    return ws


# ==========================================================================
# ABA: SISTEMA  (base da empresa - para colar)
# ==========================================================================
SIS_COLS = [
    ("RE (matrícula)", 12), ("COLABORADOR (titular)", 38), ("NOME DA PESSOA", 38),
    ("GRAU / CONDIÇÃO", 18), ("NASCIMENTO", 13), ("TIPO (opcional)", 15),
    ("UNIDADE", 22), ("SITUAÇÃO", 16),
    # ---- calculadas ----
    ("Dt. nascimento", 13), ("Idade calculada", 11), ("Tipo", 13), ("Nome normalizado", 34),
    ("Chave RE+Nasc.", 20), ("Chave RE+Nome", 30), ("Idx Hapvida", 11),
    ("Está no Hapvida?", 12), ("_seq", 8), ("Elegível a agregado?", 12),
    ("RE existe no Hapvida?", 12), ("Chave RE+Tipo", 22),
]


def aba_sistema(wb):
    ws = wb.create_sheet("SISTEMA")
    titulo(ws, "A1", "BASE DO SISTEMA DA EMPRESA — cole aqui (colunas A a H, a partir da linha 4). "
                     "As colunas I a T são calculadas: não digite nelas.", span=20, size=11, bg=C_VERM)
    ws["A2"] = ("Uma linha por PESSOA: o próprio colaborador (GRAU = TITULAR) e cada dependente/agregado. "
                "RE é o identificador do cruzamento. NASCIMENTO em DD/MM/AAAA. "
                "TIPO é opcional — se ficar em branco, é deduzido do GRAU/CONDIÇÃO.")
    ws["A2"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("A2:R2")
    ws.row_dimensions[2].height = 26

    header_row(ws, 3, [c[0] for c in SIS_COLS], h=34)
    for i, (_, w) in enumerate(SIS_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    # linha 4 = exemplo (formato esperado), removivel
    exemplo = [11364, "COLABORADOR EXEMPLO", "ISABELLE EXEMPLO DA SILVA", "FILHO(A)",
               dt.date(2016, 5, 12), "DEPENDENTE", "MATRIZ", "ATIVO"]
    for i, v in enumerate(exemplo):
        c = ws.cell(row=4, column=i + 1, value=v)
        c.font = f(italic=True, size=9, color="808080")
        if i == 4:
            c.number_format = DATA
    ws.cell(row=4, column=1).comment = Comment(
        "LINHA DE EXEMPLO — mostra o formato esperado. Apague-a ao colar a base real.", "Conciliação")

    first = 4
    last = N_SIS + first - 1
    for r in range(first, last + 1):
        ws.cell(row=r, column=9, value='=' + data_formula("$E{}".format(r))).number_format = DATA
        ws.cell(row=r, column=10, value='=IF($I{r}="","",DATEDIF($I{r},DATA_BASE,"Y"))'.format(r=r))
        ws.cell(row=r, column=11, value='=IF($A{r}="","",IF($F{r}<>"",UPPER(TRIM($F{r})),IF(OR(UPPER($D{r})="TITULAR",UPPER($D{r})="COLABORADOR",UPPER($D{r})="SEGURADO"),"TITULAR",IF(UPPER($D{r})="AGREGADO","AGREGADO","DEPENDENTE"))))'.format(r=r))
        ws.cell(row=r, column=12, value='=IF($C{r}="","",{n})'.format(r=r, n=norm_formula("$C{}".format(r))))
        ws.cell(row=r, column=13, value='=IF($A{r}="","",$A{r}&"|"&IF($I{r}="","?",TEXT($I{r},"ddmmyyyy")))'.format(r=r))
        ws.cell(row=r, column=14, value='=IF($A{r}="","",$A{r}&"|"&$L{r})'.format(r=r))
        ws.cell(row=r, column=15, value='=IF($A{r}="","",IFERROR(MATCH($M{r},HAP_CHDT,0),IFERROR(MATCH($N{r},HAP_CHNM,0),"")))'.format(r=r))
        ws.cell(row=r, column=16, value='=IF($A{r}="","",IF(ISNUMBER($O{r}),"SIM","NÃO"))'.format(r=r))
        if r == first:
            ws.cell(row=r, column=17, value='=IF($P{r}="NÃO",1,0)'.format(r=r))
        else:
            ws.cell(row=r, column=17, value='=$Q{p}+IF($P{r}="NÃO",1,0)'.format(p=r - 1, r=r))
        ws.cell(row=r, column=18, value='=IF($K{r}="","",IF(AND($K{r}="DEPENDENTE",N($J{r})>IDADE_LIMITE,OR(LEFT(UPPER($D{r}),4)<>"CONJ",REGRA_CONJUGE="SIM")),"SIM","NÃO"))'.format(r=r))
        ws.cell(row=r, column=19, value='=IF($A{r}="","",IF(COUNTIF(HAP_RE,$A{r})>0,"SIM","NÃO"))'.format(r=r))
        ws.cell(row=r, column=20, value='=IF($A{r}="","",$A{r}&"|"&$K{r})'.format(r=r))

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = "A3:S{}".format(last)
    ws.column_dimensions.group("Q", "Q", hidden=True)
    ws.column_dimensions.group("T", "T", hidden=True)
    return ws


# ==========================================================================
# NOMES DEFINIDOS
# ==========================================================================
def col_ref(sheet, col, r1, r2):
    return "'{s}'!${c}${a}:${c}${b}".format(s=sheet, c=col, a=r1, b=r2)


def definir_nomes(wb):
    add = lambda n, ref: wb.defined_names.add(DefinedName(n, attr_text=ref))

    add("DATA_BASE", "'PARÂMETROS'!$C$6")
    add("IDADE_LIMITE", "'PARÂMETROS'!$C$7")
    add("REGRA_CONJUGE", "'PARÂMETROS'!$C$8")
    add("VLR_TITULAR", "'PARÂMETROS'!$C$11")
    add("VLR_AGREGADO", "'PARÂMETROS'!$C$12")
    add("COMBO_QTD", "'PARÂMETROS'!$B$16:$B${}".format(15 + len(COMBO)))
    add("COMBO_VLR", "'PARÂMETROS'!$C$16:$C${}".format(15 + len(COMBO)))

    h1, h2 = 2, N_HAP + 1
    for nome, col in [("HAP_CONTRATO", "A"), ("HAP_RE", "B"), ("HAP_NOME", "C"),
                      ("HAP_NASC", "E"), ("HAP_INICIO", "F"), ("HAP_PAREN", "H"),
                      ("HAP_COB", "J"), ("HAP_DTN", "K"), ("HAP_IDADE", "L"),
                      ("HAP_TIPO", "M"), ("HAP_NOMENORM", "N"), ("HAP_CHDT", "O"),
                      ("HAP_CHNM", "P"), ("HAP_IDXSIS", "Q"), ("HAP_RENOSIS", "R"),
                      ("HAP_ELEG", "S"), ("HAP_CHDUP", "T"),
                      ("HAP_DUP", "U"), ("HAP_CHTIT", "V"),
                      ("HAP_1TIT", "W"), ("HAP_SEQTIT", "Y"),
                      ("HAP_CHORD", "Z"), ("HAP_DIVDADOS", "AA")]:
        add(nome, col_ref("HAPVIDA", col, h1, h2))

    s1, s2 = 4, N_SIS + 3
    for nome, col in [("SIS_RE", "A"), ("SIS_COLAB", "B"), ("SIS_NOME", "C"),
                      ("SIS_GRAU", "D"), ("SIS_UNID", "G"), ("SIS_SIT", "H"),
                      ("SIS_DTN", "I"), ("SIS_IDADE", "J"), ("SIS_TIPO", "K"),
                      ("SIS_NOMENORM", "L"), ("SIS_CHDT", "M"), ("SIS_CHNM", "N"),
                      ("SIS_IDXHAP", "O"), ("SIS_NOHAP", "P"), ("SIS_SEQ", "Q"),
                      ("SIS_ELEG", "R"), ("SIS_RENOHAP", "S"), ("SIS_CHORD", "T")]:
        add(nome, col_ref("SISTEMA", col, s1, s2))

    t1, t2 = 2, N_TIT + 1
    mapa_tit = [("TIT_RE", "A"), ("TIT_NOME", "B"), ("TIT_CONTRATO", "C"), ("TIT_INICIO", "D"),
                ("TIT_NOSIS", "E"), ("TIT_VINCULO", "F"), ("TIT_QDEP", "G"), ("TIT_QAGR", "H"),
                ("TIT_QDEPSIS", "I"), ("TIT_QAGRSIS", "J"), ("TIT_DEPOK", "K"),
                ("TIT_DEPSOHAP", "L"), ("TIT_DEPINC", "M"), ("TIT_AGRINC", "N"),
                ("TIT_DEPELEG", "O"), ("TIT_VTIT", "P"), ("TIT_VCOMBO", "Q"),
                ("TIT_VAGR", "R"), ("TIT_TOTAL", "S"), ("TIT_COBRADO", "T"),
                ("TIT_DIF", "U"), ("TIT_EMPRESA", "V"), ("TIT_COLAB", "W"),
                ("TIT_SITUACAO", "X"), ("TIT_ALERTA", "Y"), ("TIT_DIVDADOS", "Z"),
                ("TIT_SEQALERTA", "AB")]
    for nome, col in mapa_tit:
        add(nome, col_ref("TITULARES", col, t1, t2))


# ==========================================================================
# ABA: TITULARES  (um registro por RE - espinha dorsal da pasta)
# ==========================================================================
TIT_COLS = [
    ("RE", 10), ("COLABORADOR (titular)", 38), ("CONTRATO", 22), ("INÍCIO NO PLANO", 13),
    ("RE existe no sistema?", 11), ("VÍNCULO / CLASSIFICAÇÃO", 30),
    ("Dep. HAPVIDA", 10), ("Agreg. HAPVIDA", 10), ("Dep. SISTEMA", 10), ("Agreg. SISTEMA", 10),
    ("Dep. nas 2 bases", 10), ("Dep. só HAPVIDA", 10),
    ("A incluir como DEPENDENTE", 12), ("A incluir como AGREGADO", 12),
    ("Dep. HAPVIDA acima da idade", 12),
    ("Valor TITULAR", 14), ("Valor COMBO", 14), ("Valor AGREGADOS", 14), ("TOTAL calculado", 15),
    ("Cobrado no arquivo", 15), ("Diferença de cobrança", 15),
    ("Custo EMPRESA", 14), ("Custo COLABORADOR", 15),
    ("SITUAÇÃO CADASTRAL", 40), ("ALERTAS", 60),
    ("Pessoas com divergência de dados", 12), ("_idx", 8), ("_seqal", 8),
]


def aba_titulares(wb):
    ws = wb.create_sheet("TITULARES")
    header_row(ws, 1, [c[0] for c in TIT_COLS], h=42)
    for i, (_, w) in enumerate(TIT_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    for r in range(2, N_TIT + 2):
        g = lambda c: "${}{}".format(c, r)
        ws.cell(row=r, column=27, value='=IFERROR(MATCH(ROW()-1,HAP_SEQTIT,0),"")')
        ws.cell(row=r, column=26, value='=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_DIVDADOS,"SIM"))'.format(a=g("A")))
        ws.cell(row=r, column=1, value='=IFERROR(INDEX(HAP_RE,{z}),"")'.format(z=g("AA")))
        ws.cell(row=r, column=2, value='=IFERROR(INDEX(HAP_NOME,{z}),"")'.format(z=g("AA")))
        ws.cell(row=r, column=3, value='=IFERROR(INDEX(HAP_CONTRATO,{z}),"")'.format(z=g("AA")))
        c = ws.cell(row=r, column=4, value='=IFERROR(INDEX(HAP_INICIO,{z}),"")'.format(z=g("AA")))
        c.number_format = DATA
        ws.cell(row=r, column=5, value=(
            '=IF({a}="","",IF(COUNTA(SIS_RE)=0,"NÃO",IF(COUNTIF(SIS_RE,{a})>0,"SIM","NÃO")))').format(a=g("A")))
        ws.cell(row=r, column=6, value=(
            '=IF({a}="","",IF({e}="SIM","FUNCIONÁRIO (LOCALIZADO NO SISTEMA)",'
            'IF(ISNUMBER(SEARCH("PJ",{c})),"PJ",'
            'IF(ISNUMBER(SEARCH("ACORDO",{c})),"ACORDO COM A EMPRESA",'
            'IF(ISNUMBER(SEARCH("SINDICATO",{c})),"SINDICATO",'
            '"NÃO LOCALIZADO – VERIFICAR")))))').format(a=g("A"), e=g("E"), c=g("C")))
        ws.cell(row=r, column=7, value='=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO"))'.format(a=g("A")))
        ws.cell(row=r, column=8, value='=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO"))'.format(a=g("A")))
        ws.cell(row=r, column=9, value='=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"DEPENDENTE"))'.format(a=g("A")))
        ws.cell(row=r, column=10, value='=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"AGREGADO"))'.format(a=g("A")))
        ws.cell(row=r, column=11, value='=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_TIPO,"DEPENDENTE",HAP_IDXSIS,">0",HAP_DUP,"NÃO"))'.format(a=g("A")))
        ws.cell(row=r, column=12, value='=IF({a}="","",IF({e}<>"SIM",0,{g}-{k}))'.format(
            a=g("A"), e=g("E"), g=g("G"), k=g("K")))
        ws.cell(row=r, column=13, value='=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO",SIS_ELEG,"NÃO"))'.format(a=g("A")))
        ws.cell(row=r, column=14, value=(
            '=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"AGREGADO",SIS_NOHAP,"NÃO")'
            '+COUNTIFS(SIS_RE,{a},SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO",SIS_ELEG,"SIM"))').format(a=g("A")))
        ws.cell(row=r, column=15, value='=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_ELEG,"SIM",HAP_DUP,"NÃO"))'.format(a=g("A")))
        ws.cell(row=r, column=16, value='=IF({a}="","",VLR_TITULAR)'.format(a=g("A"))).number_format = MOEDA
        ws.cell(row=r, column=17, value='=IF({a}="","",INDEX(COMBO_VLR,MATCH(MIN({g},5),COMBO_QTD,0)))'.format(a=g("A"), g=g("G"))).number_format = MOEDA
        ws.cell(row=r, column=18, value='=IF({a}="","",{h}*VLR_AGREGADO)'.format(a=g("A"), h=g("H"))).number_format = MOEDA
        ws.cell(row=r, column=19, value='=IF({a}="","",{p}+{q}+{rr})'.format(a=g("A"), p=g("P"), q=g("Q"), rr=g("R"))).number_format = MOEDA
        ws.cell(row=r, column=20, value='=IF({a}="","",SUMIFS(HAP_COB,HAP_RE,{a}))'.format(a=g("A"))).number_format = MOEDA
        ws.cell(row=r, column=21, value='=IF({a}="","",{t}-{s})'.format(a=g("A"), t=g("T"), s=g("S"))).number_format = MOEDA
        ws.cell(row=r, column=22, value='=IF({a}="","",{p})'.format(a=g("A"), p=g("P"))).number_format = MOEDA
        ws.cell(row=r, column=23, value='=IF({a}="","",{q}+{rr})'.format(a=g("A"), q=g("Q"), rr=g("R"))).number_format = MOEDA
        so_agr = ('IF({e}<>"SIM",0,COUNTIFS(HAP_RE,{a},HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")'
                  '-COUNTIFS(HAP_RE,{a},HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO",HAP_IDXSIS,">0"))').format(
                      a=g("A"), e=g("E"))
        ws.cell(row=r, column=24, value=(
            '=IF({a}="","",IF({e}<>"SIM","🔵 RE NÃO LOCALIZADO NA BASE DA EMPRESA",'
            'IF(({l}+{sa}+{m}+{n})=0,IF({z}>0,"⚠️ DIVERGÊNCIA DE DADOS","🟢 CADASTRO OK"),'
            'IF({m}+{n}>0,IF({l}+{sa}>0,"⚠️ DIVERGÊNCIA NOS DOIS SENTIDOS","🟡 INCLUSÃO NECESSÁRIA"),'
            '"🔴 DIVERGÊNCIA (consta só no Hapvida)"))))').format(
                a=g("A"), e=g("E"), l=g("L"), m=g("M"), n=g("N"), z=g("Z"), sa=so_agr))
        ws.cell(row=r, column=25, value=(
            '=IF({a}="","",TRIM('
            'IF(ABS({u})>0.01,"Cobrança diverge do contrato em "&TEXT({u},"R$ #,##0.00")&". ","")&'
            'IF({o}>0,{o}&" dependente(s) acima da idade limite. ","")&'
            'IF({m}+{n}>0,{m}+{n}&" pessoa(s) do sistema a incluir no Hapvida. ","")&'
            'IF({l}>0,{l}&" dependente(s) constam só no Hapvida. ","")&'
            'IF({z}>0,{z}&" pessoa(s) com divergência de dados (nome, nascimento ou parentesco). ","")&'
            'IF({sa}>0,{sa}&" agregado(s) constam só no Hapvida. ","")&'
            'IF(COUNTIFS(HAP_RE,{a},HAP_DUP,"SIM")>0,COUNTIFS(HAP_RE,{a},HAP_DUP,"SIM")'
            '&" registro(s) DUPLICADO(S) no arquivo do Hapvida. ","")&'
            'IF(COUNTIFS(HAP_RE,{a},HAP_TIPO,"TITULAR",HAP_DUP,"NÃO")>1,'
            '"ATENÇÃO: matrícula compartilhada por "&COUNTIFS(HAP_RE,{a},HAP_TIPO,"TITULAR",HAP_DUP,"NÃO")'
            '&" titulares DIFERENTES — os dependentes e os valores desta linha somam os dois. Verificar. ","")))').format(
                a=g("A"), u=g("U"), o=g("O"), m=g("M"), n=g("N"), l=g("L"), z=g("Z"), sa=so_agr))
        if r == 2:
            ws.cell(row=r, column=28, value='=IF(AND($A2<>"",$Y2<>""),1,0)')
        else:
            ws.cell(row=r, column=28, value='=$AB{p}+IF(AND($A{r}<>"",$Y{r}<>""),1,0)'.format(p=r - 1, r=r))

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:Z{}".format(N_TIT + 1)
    ws.column_dimensions.group("AA", "AB", hidden=True)
    return ws


# ==========================================================================
# ABA: CONCILIACAO
# ==========================================================================
CONC_COLS = [
    ("RE", 10), ("COLABORADOR", 36), ("CONTRATO", 20), ("DEPENDENTE / AGREGADO", 38),
    ("NATUREZA", 26), ("GRAU/CONDIÇÃO HAPVIDA", 18), ("NASCIMENTO HAPVIDA", 14),
    ("IDADE HAPVIDA", 9), ("VALOR HAPVIDA", 13), ("ESTÁ NO HAPVIDA?", 11),
    ("GRAU/CONDIÇÃO SISTEMA", 18), ("NASCIMENTO SISTEMA", 14), ("IDADE SISTEMA", 9),
    ("ESTÁ NO SISTEMA?", 11), ("SITUAÇÃO", 38), ("OBSERVAÇÃO", 70),
    ("_q", 8), ("_r", 8), ("_chdiv", 8),
]

STATUS_CORES = [
    ("CADASTRO OK", F_VERDE, C_VERDE),
    ("INCLUSÃO NECESSÁRIA", F_AMAR, C_AMAR),
    ("DIVERGÊNCIA DE DADOS", F_LARANJ, C_LARANJ),
    ("DIVERGÊNCIA NOS DOIS", F_LARANJ, C_LARANJ),
    ("só no Hapvida", F_VERM, C_VERM),
    ("NÃO LOCALIZADO", F_AZUL, C_AZUL),
    ("SEM CONVÊNIO", C_CINZA2, "555555"),
]


def pinta_status(ws, rng):
    """Formatacao condicional por palavra-chave (evita LEFT() sobre emoji)."""
    primeira = rng.split(":")[0].replace("$", "")
    for chave, bg, fg in STATUS_CORES:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=['ISNUMBER(SEARCH("{}",{}))'.format(chave, primeira)],
            fill=fill(bg), font=Font(name=FONTE, color=fg, bold=True, size=9),
            stopIfTrue=True))


def aba_conciliacao(wb):
    ws = wb.create_sheet("CONCILIAÇÃO")
    header_row(ws, 1, [c[0] for c in CONC_COLS], h=42)
    for i, (_, w) in enumerate(CONC_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    ini_b = N_HAP + 2
    fim = N_HAP + 1 + N_CONC_B

    for r in range(2, fim + 1):
        g = lambda c: "${}{}".format(c, r)
        if r < ini_b:
            ws.cell(row=r, column=17, value=r - 1)                     # indice na base HAPVIDA
            ws.cell(row=r, column=18, value='=IFERROR(INDEX(HAP_IDXSIS,{q}),"")'.format(q=g("Q")))
        else:
            ws.cell(row=r, column=18, value='=IFERROR(MATCH({j},SIS_SEQ,0),"")'.format(j=r - N_HAP - 1))

        ws.cell(row=r, column=1, value=(
            '=IFERROR(IF({q}<>"",IF(INDEX(HAP_CHDT,{q})="","",INDEX(HAP_RE,{q})),'
            'IF({rr}="","",INDEX(SIS_RE,{rr}))),"")').format(q=g("Q"), rr=g("R")))
        ws.cell(row=r, column=2, value=(
            '=IF({a}="","",IFERROR(INDEX(TIT_NOME,MATCH({a},TIT_RE,0)),'
            'IFERROR(INDEX(SIS_COLAB,{rr}),"(titular não consta no Hapvida)")))').format(a=g("A"), rr=g("R")))
        ws.cell(row=r, column=3, value=(
            '=IF({a}="","",IF({q}<>"",INDEX(HAP_CONTRATO,{q}),'
            'IFERROR(INDEX(TIT_CONTRATO,MATCH({a},TIT_RE,0)),"—")))').format(a=g("A"), q=g("Q")))
        ws.cell(row=r, column=4, value=(
            '=IF({a}="","",IF({q}<>"",INDEX(HAP_NOME,{q}),INDEX(SIS_NOME,{rr})))').format(
                a=g("A"), q=g("Q"), rr=g("R")))
        ws.cell(row=r, column=5, value=(
            '=IF({a}="","",IF({q}<>"",'
            'IF(INDEX(HAP_TIPO,{q})="AGREGADO","🟠 AGREGADO",'
            'IF(INDEX(HAP_ELEG,{q})="SIM","⚠️ DEPENDENTE ACIMA DA IDADE",INDEX(HAP_TIPO,{q}))),'
            'IF(INDEX(SIS_TIPO,{rr})="AGREGADO","🟠 AGREGADO",'
            'IF(INDEX(SIS_ELEG,{rr})="SIM","⚠️ DEPENDENTE ACIMA DA IDADE",INDEX(SIS_TIPO,{rr})))))').format(
                a=g("A"), q=g("Q"), rr=g("R")))
        ws.cell(row=r, column=6, value='=IF({a}="","",IF({q}<>"",INDEX(HAP_PAREN,{q}),""))'.format(a=g("A"), q=g("Q")))
        c = ws.cell(row=r, column=7, value='=IF({a}="","",IF({q}<>"",INDEX(HAP_DTN,{q}),""))'.format(a=g("A"), q=g("Q")))
        c.number_format = DATA
        ws.cell(row=r, column=8, value='=IF({a}="","",IF({q}<>"",INDEX(HAP_IDADE,{q}),""))'.format(a=g("A"), q=g("Q")))
        ws.cell(row=r, column=9, value='=IF({a}="","",IF({q}<>"",INDEX(HAP_COB,{q}),0))'.format(a=g("A"), q=g("Q"))).number_format = MOEDA
        ws.cell(row=r, column=10, value='=IF({a}="","",IF({q}<>"","SIM","NÃO"))'.format(a=g("A"), q=g("Q")))
        ws.cell(row=r, column=11, value='=IF({a}="","",IF({rr}="","",INDEX(SIS_GRAU,{rr})))'.format(a=g("A"), rr=g("R")))
        c = ws.cell(row=r, column=12, value='=IF({a}="","",IF({rr}="","",INDEX(SIS_DTN,{rr})))'.format(a=g("A"), rr=g("R")))
        c.number_format = DATA
        ws.cell(row=r, column=13, value='=IF({a}="","",IF({rr}="","",INDEX(SIS_IDADE,{rr})))'.format(a=g("A"), rr=g("R")))
        ws.cell(row=r, column=14, value='=IF({a}="","",IF({rr}="","NÃO","SIM"))'.format(a=g("A"), rr=g("R")))
        ws.cell(row=r, column=16, value=(
            '=IF({a}="","",TRIM(IF(AND({j}="SIM",{n}="SIM"),'
            'IF(INDEX(HAP_NOMENORM,{q})<>INDEX(SIS_NOMENORM,{rr}),"Nome grafado diferente no sistema: "&INDEX(SIS_NOME,{rr})&". ","")'
            '&IF(AND(N({gg})>0,N({l})>0,N({gg})<>N({l})),"Nascimento diferente — sistema: "&TEXT({l},"dd/mm/yyyy")&". ","")'
            '&IF(AND({k}<>"",LEFT(UPPER({ff}),4)<>LEFT(UPPER({k}),4)),"Grau/condição diferente — sistema: "&{k}&". ",""),'
            'IF({j}="NÃO","Consta no sistema da empresa e não no Hapvida. "'
            '&IF(INDEX(SIS_RENOHAP,{rr})="NÃO","O RE não possui nenhum registro no Hapvida. ","")'
            '&IF(INDEX(SIS_ELEG,{rr})="SIM","Acima da idade limite: entra como AGREGADO ("&TEXT(VLR_AGREGADO,"R$ #,##0.00")&"). ",""),'
            'IF(INDEX(HAP_RENOSIS,{q})="NÃO","RE não localizado na base da empresa. Contrato informado: "&{c}&". ",'
            '"Consta no Hapvida e não no cadastro do sistema da empresa. "))'
            '&IF({q}="","",IF(INDEX(HAP_DUP,{q})="SIM","REGISTRO DUPLICADO no arquivo do Hapvida. ",""))))').format(
                a=g("A"), j=g("J"), n=g("N"), q=g("Q"), rr=g("R"), gg=g("G"),
                l=g("L"), k=g("K"), ff=g("F"), c=g("C")))
        ws.cell(row=r, column=15, value=(
            '=IF({a}="","",IF(AND({j}="SIM",{n}="SIM"),'
            'IF(INDEX(HAP_DIVDADOS,{q})="SIM","⚠️ DIVERGÊNCIA DE DADOS","🟢 CADASTRO OK"),'
            'IF({j}="NÃO",IF(INDEX(SIS_TIPO,{rr})="TITULAR","⚪ TITULAR DO SISTEMA SEM CONVÊNIO",'
            'IF(INDEX(SIS_RENOHAP,{rr})="NÃO","⚪ RE SEM CONVÊNIO NO HAPVIDA","🟡 INCLUSÃO NECESSÁRIA")),'
            'IF(INDEX(HAP_RENOSIS,{q})="NÃO","🔵 RE NÃO LOCALIZADO NA BASE DA EMPRESA",'
            '"🔴 DIVERGÊNCIA (consta só no Hapvida)"))))').format(
                a=g("A"), j=g("J"), n=g("N"), p=g("P"), q=g("Q"), rr=g("R")))

    for r in range(2, fim + 1):
        ws.cell(row=r, column=19, value=(
            '=IF($A{r}="","",$A{r}&IF(ISNUMBER(SEARCH("CADASTRO OK",$O{r})),"|OK","|DIV"))').format(r=r))
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = "A1:P{}".format(fim)
    pinta_status(ws, "O2:O{}".format(fim))
    ws.column_dimensions.group("Q", "S", hidden=True)
    wb.defined_names.add(DefinedName("CONC_CHDIV", attr_text="'CONCILIAÇÃO'!$S$2:$S${}".format(fim)))
    wb.defined_names.add(DefinedName("CONC_SIT", attr_text="'CONCILIAÇÃO'!$O$2:$O${}".format(fim)))
    wb.defined_names.add(DefinedName("CONC_NAT", attr_text="'CONCILIAÇÃO'!$E$2:$E${}".format(fim)))
    wb.defined_names.add(DefinedName("CONC_RE", attr_text="'CONCILIAÇÃO'!$A$2:$A${}".format(fim)))
    wb.defined_names.add(DefinedName("CONC_PESSOA", attr_text="'CONCILIAÇÃO'!$D$2:$D${}".format(fim)))
    wb.defined_names.add(DefinedName("CONC_OBS", attr_text="'CONCILIAÇÃO'!$P$2:$P${}".format(fim)))
    wb.defined_names.add(DefinedName("CONC_COLAB", attr_text="'CONCILIAÇÃO'!$B$2:$B${}".format(fim)))
    return ws


# ==========================================================================
# ABA: SIMULACAO DE IMPACTO
# ==========================================================================
SIM_COLS = [
    ("RE", 10), ("Colaborador", 36), ("Dependentes Atuais", 11), ("Dependentes Sistema", 11),
    ("A Incluir", 10), ("Agregados", 10), ("Valor Atual", 14), ("Novo Valor", 14),
    ("Aumento Mensal", 14), ("Aumento Anual", 15),
    ("A incluir como AGREGADO", 12), ("Novo nº de dependentes", 11),
    ("Novo nº de agregados", 11), ("COMO FICA", 74),
]


def aba_simulacao(wb):
    ws = wb.create_sheet("SIMULAÇÃO DE IMPACTO")
    titulo(ws, "A1", "SIMULAÇÃO — quanto custaria regularizar no Hapvida tudo o que já consta "
                     "no sistema da empresa", span=14, size=13)
    ws["A2"] = ("Dependente com até a idade limite entra no COMBO FAMILIAR (o valor muda de faixa, "
                "nunca é multiplicado). Acima da idade limite entra como AGREGADO, somando "
                "o valor cheio por pessoa. O aumento é de responsabilidade do colaborador.")
    ws["A2"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("A2:N2")
    ws.row_dimensions[2].height = 26

    header_row(ws, 3, [c[0] for c in SIM_COLS], h=44)
    for i, (_, w) in enumerate(SIM_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    for r in range(4, N_TIT + 4):
        k = r - 3
        g = lambda c: "${}{}".format(c, r)
        ws.cell(row=r, column=1, value='=IFERROR(IF(INDEX(TIT_RE,{k})="","",INDEX(TIT_RE,{k})),"")'.format(k=k))
        ws.cell(row=r, column=2, value='=IF({a}="","",INDEX(TIT_NOME,{k}))'.format(a=g("A"), k=k))
        ws.cell(row=r, column=3, value='=IF({a}="","",INDEX(TIT_QDEP,{k}))'.format(a=g("A"), k=k))
        ws.cell(row=r, column=4, value='=IF({a}="","",INDEX(TIT_QDEPSIS,{k}))'.format(a=g("A"), k=k))
        ws.cell(row=r, column=5, value='=IF({a}="","",INDEX(TIT_DEPINC,{k}))'.format(a=g("A"), k=k))
        ws.cell(row=r, column=6, value='=IF({a}="","",INDEX(TIT_QAGR,{k}))'.format(a=g("A"), k=k))
        ws.cell(row=r, column=7, value='=IF({a}="","",INDEX(TIT_TOTAL,{k}))'.format(a=g("A"), k=k)).number_format = MOEDA
        ws.cell(row=r, column=11, value='=IF({a}="","",INDEX(TIT_AGRINC,{k}))'.format(a=g("A"), k=k))
        ws.cell(row=r, column=12, value='=IF({a}="","",{c}+{e})'.format(a=g("A"), c=g("C"), e=g("E")))
        ws.cell(row=r, column=13, value='=IF({a}="","",{ff}+{kk})'.format(a=g("A"), ff=g("F"), kk=g("K")))
        ws.cell(row=r, column=8, value=(
            '=IF({a}="","",VLR_TITULAR+INDEX(COMBO_VLR,MATCH(MIN({l},5),COMBO_QTD,0))'
            '+{m}*VLR_AGREGADO)').format(a=g("A"), l=g("L"), m=g("M"))).number_format = MOEDA
        ws.cell(row=r, column=9, value='=IF({a}="","",{h}-{gg})'.format(a=g("A"), h=g("H"), gg=g("G"))).number_format = MOEDA
        ws.cell(row=r, column=10, value='=IF({a}="","",{i}*12)'.format(a=g("A"), i=g("I"))).number_format = MOEDA
        ws.cell(row=r, column=14, value=(
            '=IF({a}="","",TRIM('
            'IF({e}>0,{e}&" dependente(s) a incluir: o Combo passa da faixa de "&{c}&" para "&{l}&" dependentes. ","")'
            '&IF({kk}>0,{kk}&" pessoa(s) entram como AGREGADO a "&TEXT(VLR_AGREGADO,"R$ #,##0.00")&" cada. ","")'
            '&IF({i}=0,"Nada a incluir — sem impacto financeiro.","")))').format(
                a=g("A"), e=g("E"), c=g("C"), l=g("L"), kk=g("K"), i=g("I")))

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = "A3:N{}".format(N_TIT + 3)
    wb.defined_names.add(DefinedName("SIM_AUM_MES", attr_text="'SIMULAÇÃO DE IMPACTO'!$I$4:$I${}".format(N_TIT + 3)))
    wb.defined_names.add(DefinedName("SIM_AUM_ANO", attr_text="'SIMULAÇÃO DE IMPACTO'!$J$4:$J${}".format(N_TIT + 3)))
    wb.defined_names.add(DefinedName("SIM_NOVO", attr_text="'SIMULAÇÃO DE IMPACTO'!$H$4:$H${}".format(N_TIT + 3)))
    return ws


# ==========================================================================
# ABA: RESUMO POR CONTRATO
# ==========================================================================
def aba_resumo_contrato(wb, contratos):
    ws = wb.create_sheet("RESUMO POR CONTRATO")
    ws.sheet_view.showGridLines = False
    titulo(ws, "B2", "RESUMO POR CONTRATO", span=10, size=15)
    ws["B3"] = ("Os nomes de contrato na coluna B são editáveis (azul). Acrescente novas linhas de "
                "contrato nas linhas em branco reservadas — a linha OUTROS mostra tudo o que não "
                "foi listado, para nada ficar de fora.")
    ws["B3"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("B3:K3")
    ws.row_dimensions[3].height = 26

    cols = ["CONTRATO", "Titulares", "Dependentes", "Agregados", "Valor Titulares",
            "Valor Dependentes (Combo)", "Valor Agregados", "TOTAL",
            "Custo EMPRESA", "Custo COLABORADOR"]
    header_row(ws, 5, cols, start_col=2, h=40)
    larguras(ws, {"A": 3, "B": 30, "C": 11, "D": 12, "E": 11,
                  "F": 15, "G": 18, "H": 15, "I": 16, "J": 15, "K": 17})

    n_slots = len(contratos) + 6
    r0 = 6
    for i in range(n_slots):
        r = r0 + i
        nome = contratos[i] if i < len(contratos) else None
        c = ws.cell(row=r, column=2, value=nome)
        c.font = f(color=C_AZUL_TXT, bold=True, size=10)
        c.fill = fill("FFFFCC")
        c.border = BORDA
        a = "$B{}".format(r)
        specs = [
            (3, '=IF({a}="","",COUNTIFS(TIT_CONTRATO,{a}))', INT),
            (4, '=IF({a}="","",SUMIFS(TIT_QDEP,TIT_CONTRATO,{a}))', INT),
            (5, '=IF({a}="","",SUMIFS(TIT_QAGR,TIT_CONTRATO,{a}))', INT),
            (6, '=IF({a}="","",SUMIFS(TIT_VTIT,TIT_CONTRATO,{a}))', MOEDA),
            (7, '=IF({a}="","",SUMIFS(TIT_VCOMBO,TIT_CONTRATO,{a}))', MOEDA),
            (8, '=IF({a}="","",SUMIFS(TIT_VAGR,TIT_CONTRATO,{a}))', MOEDA),
            (9, '=IF({a}="","",SUMIFS(TIT_TOTAL,TIT_CONTRATO,{a}))', MOEDA),
            (10, '=IF({a}="","",SUMIFS(TIT_EMPRESA,TIT_CONTRATO,{a}))', MOEDA),
            (11, '=IF({a}="","",SUMIFS(TIT_COLAB,TIT_CONTRATO,{a}))', MOEDA),
        ]
        for col, fml, fmt in specs:
            cc = ws.cell(row=r, column=col, value=fml.format(a=a))
            cc.number_format = fmt
            cc.border = BORDA
            cc.font = f(size=10)

    r_out = r0 + n_slots
    r_tot = r_out + 1
    ws.cell(row=r_out, column=2, value="OUTROS (contratos não listados acima)")
    ws.cell(row=r_tot, column=2, value="TOTAL GERAL")
    totais = [
        (3, '=SUMPRODUCT(--(TIT_RE<>""))', INT),
        (4, '=SUM(TIT_QDEP)', INT), (5, '=SUM(TIT_QAGR)', INT),
        (6, '=SUM(TIT_VTIT)', MOEDA), (7, '=SUM(TIT_VCOMBO)', MOEDA),
        (8, '=SUM(TIT_VAGR)', MOEDA), (9, '=SUM(TIT_TOTAL)', MOEDA),
        (10, '=SUM(TIT_EMPRESA)', MOEDA), (11, '=SUM(TIT_COLAB)', MOEDA),
    ]
    for col, fml, fmt in totais:
        L = CL(col)
        co = ws.cell(row=r_out, column=col,
                     value='={t}-SUM({L}{a}:{L}{b})'.format(t=fml[1:], L=L, a=r0, b=r0 + n_slots - 1))
        co.number_format = fmt
        co.border = BORDA
        co.font = f(size=10, italic=True)
        co.fill = fill(C_CINZA)
        ct = ws.cell(row=r_tot, column=col, value=fml)
        ct.number_format = fmt
        ct.border = BORDA
        ct.font = f(bold=True, size=10, color="FFFFFF")
        ct.fill = fill(C_NAVY)
    for r, bg, cor in ((r_out, C_CINZA, "555555"), (r_tot, C_NAVY, "FFFFFF")):
        cb = ws.cell(row=r, column=2)
        cb.fill = fill(bg)
        cb.font = f(bold=True, size=10, color=cor)
        cb.border = BORDA

    ws.cell(row=r_tot + 2, column=2,
            value="Custo EMPRESA = valor do titular. Custo COLABORADOR = Combo Familiar + agregados.").font = \
        f(italic=True, size=9, color="555555")
    ws.auto_filter.ref = "B5:K{}".format(r0 + n_slots - 1)
    ws.freeze_panes = "C6"
    return ws


# ==========================================================================
# ABA: RESUMO GERENCIAL (dashboard)
# ==========================================================================
def bloco(ws, r0, c0, tit, itens, largura=4, bg=C_TEAL):
    """Escreve um bloco titulo + linhas de rotulo/valor. itens = (rotulo, formula, fmt)."""
    cel = "{}{}".format(CL(c0), r0)
    titulo(ws, cel, tit, span=largura, size=11, bg=bg)
    r = r0 + 1
    for rot, fml, fmt in itens:
        cr = ws.cell(row=r, column=c0, value=rot)
        cr.font = f(size=10)
        cr.border = BORDA
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + largura - 2)
        for i in range(1, largura - 1):
            ws.cell(row=r, column=c0 + i).border = BORDA
        cv = ws.cell(row=r, column=c0 + largura - 1, value=fml)
        cv.number_format = fmt
        cv.font = f(bold=True, size=10)
        cv.border = BORDA
        cv.alignment = Alignment(horizontal="right")
        if r % 2 == 0:
            for i in range(largura):
                if not ws.cell(row=r, column=c0 + i).fill.fgColor.rgb or \
                   ws.cell(row=r, column=c0 + i).fill.patternType is None:
                    ws.cell(row=r, column=c0 + i).fill = fill(C_CINZA)
        r += 1
    return r


def card(ws, row, col, rotulo, formula, fmt, cor=C_NAVY):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 2)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 2)
    c1 = ws.cell(row=row, column=col, value=rotulo)
    c1.font = f(bold=True, size=9, color="FFFFFF")
    c1.fill = fill(cor)
    c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c2 = ws.cell(row=row + 1, column=col, value=formula)
    c2.number_format = fmt
    c2.font = f(bold=True, size=18, color=cor)
    c2.fill = fill("FFFFFF")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.border = Border(left=Side("medium", color=cor), right=Side("medium", color=cor),
                       bottom=Side("medium", color=cor))
    ws.row_dimensions[row].height = 30
    ws.row_dimensions[row + 1].height = 34


def aba_resumo(wb):
    ws = wb.create_sheet("RESUMO GERENCIAL")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 40, "C": 8, "D": 8, "E": 17, "F": 4,
                  "G": 40, "H": 8, "I": 8, "J": 17, "K": 3})

    titulo(ws, "B2", "CONVÊNIO MÉDICO HAPVIDA — RESUMO GERENCIAL", span=9, size=16)
    ws["B3"] = '="Idades calculadas na data-base "&TEXT(DATA_BASE,"dd/mm/yyyy")&'\
               '"  |  Titular "&TEXT(VLR_TITULAR,"R$ #,##0.00")&" (empresa)  |  '\
               'Agregado "&TEXT(VLR_AGREGADO,"R$ #,##0.00")&" (colaborador)  |  '\
               'Combo Familiar pela faixa de dependentes"'
    ws["B3"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("B3:J3")

    card(ws, 5, 2, "TITULARES NO HAPVIDA", '=SUMPRODUCT(--(TIT_RE<>""))', INT, C_NAVY)
    card(ws, 5, 5, "DEPENDENTES NO HAPVIDA", '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")', INT, C_TEAL)
    card(ws, 5, 8, "AGREGADOS NO HAPVIDA", '=COUNTIFS(HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")', INT, C_LARANJ)
    card(ws, 8, 2, "CUSTO MENSAL TOTAL", '=SUM(TIT_TOTAL)', MOEDA0, C_NAVY)
    card(ws, 8, 5, "CUSTO MENSAL DA EMPRESA", '=SUM(TIT_EMPRESA)', MOEDA0, C_VERDE)
    card(ws, 8, 8, "IMPACTO MENSAL SE REGULARIZAR", '=SUM(SIM_AUM_MES)', MOEDA0, C_VERM)

    r_esq = bloco(ws, 11, 2, "COLABORADORES", [
        ("Total de titulares no Hapvida", '=SUMPRODUCT(--(TIT_RE<>""))', INT),
        ("Localizados na base da empresa", '=COUNTIF(TIT_NOSIS,"SIM")', INT),
        ("PJ (pelo contrato)", '=COUNTIF(TIT_VINCULO,"PJ")', INT),
        ("Acordo com a empresa (pelo contrato)", '=COUNTIF(TIT_VINCULO,"ACORDO COM A EMPRESA")', INT),
        ("Sindicato (pelo contrato)", '=COUNTIF(TIT_VINCULO,"SINDICATO")', INT),
        ("NÃO LOCALIZADO – VERIFICAR", '=COUNTIF(TIT_VINCULO,"NÃO LOCALIZADO*")', INT),
        ("Com alguma divergência cadastral", '=COUNTIF(TIT_SITUACAO,"*DIVERGÊNCIA*")+COUNTIF(TIT_SITUACAO,"*INCLUSÃO*")', INT),
        ("Cadastro OK nas duas bases", '=COUNTIF(TIT_SITUACAO,"*CADASTRO OK*")', INT),
        ("Registros DUPLICADOS no arquivo do Hapvida", '=COUNTIF(HAP_DUP,"SIM")', INT),
        ("Matrículas compartilhadas por 2+ titulares",
         '=COUNTIFS(HAP_TIPO,"TITULAR",HAP_1TIT,"NÃO",HAP_DUP,"NÃO")', INT),
        ("REs no Hapvida sem linha de TITULAR", '=COUNTIF(CONC_COLAB,"*titular não consta*")', INT),
    ])

    r_dir = bloco(ws, 11, 7, "DEPENDENTES", [
        ("Total de dependentes no Hapvida", '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")', INT),
        ("Total de dependentes no sistema", '=COUNTIF(SIS_TIPO,"DEPENDENTE")', INT),
        ("Encontrados nas duas bases", '=SUM(TIT_DEPOK)', INT),
        ("Somente no Hapvida", '=SUM(TIT_DEPSOHAP)', INT),
        ("Somente no sistema da empresa", '=COUNTIFS(SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO")', INT),
        ("A incluir no Hapvida como DEPENDENTE", '=SUM(TIT_DEPINC)', INT),
        ("Acima da idade limite no Hapvida", '=COUNTIFS(HAP_ELEG,"SIM",HAP_DUP,"NÃO")', INT),
        ("Colaboradores com dependentes", '=COUNTIF(TIT_QDEP,">0")', INT),
        ("Vinculados a um titular (entram no Combo)", '=SUM(TIT_QDEP)', INT),
        ("Sem titular vinculado no arquivo",
         '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")-SUM(TIT_QDEP)', INT),
    ])

    r_esq = bloco(ws, r_esq + 1, 2, "AGREGADOS", [
        ("Total de agregados no Hapvida", '=COUNTIFS(HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")', INT),
        ("Total de agregados no sistema", '=COUNTIF(SIS_TIPO,"AGREGADO")', INT),
        ("A incluir no Hapvida como AGREGADO", '=SUM(TIT_AGRINC)', INT),
        ("Colaboradores com agregados", '=COUNTIF(TIT_QAGR,">0")', INT),
        ("Sem titular vinculado no arquivo",
         '=COUNTIFS(HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")-SUM(TIT_QAGR)', INT),
        ("Custo mensal dos agregados", '=SUM(TIT_VAGR)', MOEDA),
    ])

    faixas = [("0 a 5 anos", 0, 5), ("6 a 10 anos", 6, 10), ("11 a 17 anos", 11, 17),
              ("18 a 23 anos", 18, 23), ("24 a 25 anos", 24, 25)]
    itens_f = [(rot, '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_IDADE,">={a}",HAP_IDADE,"<={b}")'.format(a=a, b=b), INT)
               for rot, a, b in faixas]
    itens_f.append(("Acima de 25 anos (possíveis agregados)",
                    '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_IDADE,">"&IDADE_LIMITE)', INT))
    itens_f.append(("→ dos quais são CÔNJUGE (não viram agregado)",
                    '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_IDADE,">"&IDADE_LIMITE,HAP_PAREN,"CONJUGE*")', INT))
    itens_f.append(("→ efetivamente elegíveis a agregado", '=COUNTIFS(HAP_ELEG,"SIM",HAP_DUP,"NÃO")', INT))
    r_dir = bloco(ws, r_dir + 1, 7, "FAIXA ETÁRIA DOS DEPENDENTES (HAPVIDA)", itens_f)

    r_esq = bloco(ws, r_esq + 1, 2, "FINANCEIRO (situação atual)", [
        ("Valor total dos TITULARES", '=SUM(TIT_VTIT)', MOEDA),
        ("Valor total do COMBO FAMILIAR", '=SUM(TIT_VCOMBO)', MOEDA),
        ("Valor total dos AGREGADOS", '=SUM(TIT_VAGR)', MOEDA),
        ("VALOR TOTAL DO CONVÊNIO", '=SUM(TIT_TOTAL)', MOEDA),
        ("Custeado pela EMPRESA", '=SUM(TIT_EMPRESA)', MOEDA),
        ("Custeado pelos COLABORADORES", '=SUM(TIT_COLAB)', MOEDA),
        ("Total cobrado no arquivo do Hapvida", '=SUM(TIT_COBRADO)', MOEDA),
        ("DIFERENÇA (cobrado − calculado)", '=SUM(TIT_DIF)', MOEDA),
        ("Titulares com diferença de cobrança", '=COUNTIF(TIT_DIF,">0.01")+COUNTIF(TIT_DIF,"<-0.01")', INT),
    ])

    r_dir = bloco(ws, r_dir + 1, 7, "SIMULAÇÃO — SE TUDO FOR REGULARIZADO", [
        ("Colaboradores impactados", '=COUNTIF(SIM_AUM_MES,">0")', INT),
        ("Pessoas a incluir (dependentes)", '=SUM(TIT_DEPINC)', INT),
        ("Pessoas a incluir (agregados)", '=SUM(TIT_AGRINC)', INT),
        ("Custo mensal HOJE", '=SUM(TIT_TOTAL)', MOEDA),
        ("Custo mensal APÓS regularização", '=SUM(SIM_NOVO)', MOEDA),
        ("AUMENTO MENSAL", '=SUM(SIM_AUM_MES)', MOEDA),
        ("AUMENTO ANUAL", '=SUM(SIM_AUM_ANO)', MOEDA),
    ])

    r_esq = bloco(ws, max(r_esq, r_dir) + 1, 2, "CONCILIAÇÃO — PESSOAS POR SITUAÇÃO", [
        ("🟢 CADASTRO OK", '=COUNTIF(CONC_SIT,"*CADASTRO OK*")', INT),
        ("🟡 INCLUSÃO NECESSÁRIA", '=COUNTIF(CONC_SIT,"*INCLUSÃO NECESSÁRIA*")', INT),
        ("🔴 DIVERGÊNCIA (só no Hapvida)", '=COUNTIF(CONC_SIT,"*só no Hapvida*")', INT),
        ("⚠️ DIVERGÊNCIA DE DADOS", '=COUNTIF(CONC_SIT,"*DIVERGÊNCIA DE DADOS*")', INT),
        ("🔵 RE NÃO LOCALIZADO NA BASE DA EMPRESA", '=COUNTIF(CONC_SIT,"*NÃO LOCALIZADO*")', INT),
        ("⚪ Sem convênio / titular sem plano", '=COUNTIF(CONC_SIT,"*SEM CONVÊNIO*")', INT),
        ("🟠 AGREGADOS (natureza)", '=COUNTIF(CONC_NAT,"*AGREGADO*")', INT),
        ("TOTAL DE PESSOAS CONCILIADAS", '=SUMPRODUCT(--(CONC_RE<>""))', INT),
    ])

    rn = max(r_esq, r_dir) + 1
    ws.cell(row=rn, column=2, value=("Como ler: a ausência de um RE na base da empresa NÃO é tratada como erro — "
                 "o vínculo é classificado pelo CONTRATO do Hapvida (PJ, Acordo, Sindicato) e, "
                 "quando não é possível identificar, fica como NÃO LOCALIZADO – VERIFICAR."))
    ws.cell(row=rn, column=2).font = f(italic=True, size=9, color=C_VERM)
    ws.cell(row=rn, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=rn, start_column=2, end_row=rn, end_column=10)
    ws.row_dimensions[rn].height = 28
    wb = ws.parent
    wb.defined_names.add(DefinedName("RG_DEP_BASE", attr_text="'RESUMO GERENCIAL'!$E$6"))
    return ws


# ==========================================================================
# ABA: ALERTAS
# ==========================================================================
def aba_alertas(wb):
    ws = wb.create_sheet("ALERTAS")
    titulo(ws, "A1", "ALERTAS — apenas os REs que exigem alguma ação", span=9, size=13, bg=C_VERM)
    ws["A2"] = ("Lista viva: cada RE com divergência cadastral, diferença de cobrança ou dependente "
                "acima da idade limite aparece aqui automaticamente.")
    ws["A2"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("A2:E2")

    cols = [("RE", 10), ("COLABORADOR", 36), ("CONTRATO", 20), ("VÍNCULO", 30),
            ("SITUAÇÃO CADASTRAL", 38), ("ALERTA", 66), ("Cobrado no arquivo", 15),
            ("Total calculado", 15), ("Diferença", 14)]
    header_row(ws, 3, [c[0] for c in cols], h=34)
    for i, (_, w) in enumerate(cols):
        ws.column_dimensions[CL(i + 1)].width = w

    for i in range(N_ALERTA):
        r = 4 + i
        k = i + 1
        idx = '$J{}'.format(r)
        ws.cell(row=r, column=10, value='=IFERROR(MATCH({k},TIT_SEQALERTA,0),"")'.format(k=k))
        ws.cell(row=r, column=1, value='=IF({i}="","",INDEX(TIT_RE,{i}))'.format(i=idx))
        ws.cell(row=r, column=2, value='=IF({i}="","",INDEX(TIT_NOME,{i}))'.format(i=idx))
        ws.cell(row=r, column=3, value='=IF({i}="","",INDEX(TIT_CONTRATO,{i}))'.format(i=idx))
        ws.cell(row=r, column=4, value='=IF({i}="","",INDEX(TIT_VINCULO,{i}))'.format(i=idx))
        ws.cell(row=r, column=5, value='=IF({i}="","",INDEX(TIT_SITUACAO,{i}))'.format(i=idx))
        ws.cell(row=r, column=6, value='=IF({i}="","",INDEX(TIT_ALERTA,{i}))'.format(i=idx))
        ws.cell(row=r, column=7, value='=IF({i}="","",INDEX(TIT_COBRADO,{i}))'.format(i=idx)).number_format = MOEDA
        ws.cell(row=r, column=8, value='=IF({i}="","",INDEX(TIT_TOTAL,{i}))'.format(i=idx)).number_format = MOEDA
        ws.cell(row=r, column=9, value='=IF({i}="","",INDEX(TIT_DIF,{i}))'.format(i=idx)).number_format = MOEDA

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = "A3:I{}".format(N_ALERTA + 3)
    pinta_status(ws, "E4:E{}".format(N_ALERTA + 3))
    ws.column_dimensions.group("J", "J", hidden=True)
    ws["F2"] = ('=IF(MAX(TIT_SEQALERTA)<={n},"Exibindo os "&MAX(TIT_SEQALERTA)&" REs com alerta (capacidade desta aba: {n}).",'
                '"⚠️ ATENÇÃO: há "&MAX(TIT_SEQALERTA)&" REs com alerta e esta aba exibe apenas {n}. '
                'Copie a última linha para baixo para ver o restante.")').format(n=N_ALERTA)
    ws["F2"].font = f(bold=True, size=9, color=C_VERM)
    ws["F2"].alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells("F2:I2")
    return ws


# ==========================================================================
# ABA: FICHA DO COLABORADOR
# ==========================================================================
def _proxima(rng, chave, m_atual, m_ant):
    """Formula que acha a PROXIMA ocorrencia de `chave` em `rng` depois de `m_ant`.

    Encadeamento OFFSET+CORRESP: uma unica varredura por linha exibida, sem
    formulas matriciais. Funciona em qualquer versao do Excel, no LibreOffice
    e no Google Sheets (AGREGAR/AGGREGATE nao e universal).
    """
    if m_ant is None:
        return '=IFERROR(MATCH({c},{r},0),"")'.format(c=chave, r=rng)
    return ('=IF(N({p})=0,"",IFERROR(MATCH({c},OFFSET({r},{p},0,ROWS({r})-{p},1),0)+{p},""))'
            .format(c=chave, r=rng, p=m_ant))


def _borda(ws, r, c1, c2, bg=None):
    for c in range(c1, c2 + 1):
        cel = ws.cell(row=r, column=c)
        cel.border = BORDA
        if bg:
            cel.fill = fill(bg)


def _tabela_ficha(ws, r_head, labels_spans, n_linhas, campos, chave, rng_ref):
    """Monta um bloco-tabela da ficha. campos = lista de (c1, c2, formula_tpl, fmt, align)."""
    for (c1, c2, texto) in labels_spans:
        ws.merge_cells(start_row=r_head, start_column=c1, end_row=r_head, end_column=c2)
        cel = ws.cell(row=r_head, column=c1, value=texto)
        cel.font = f(bold=True, color="FFFFFF", size=9)
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c in range(c1, c2 + 1):
            ws.cell(row=r_head, column=c).fill = fill(C_NAVY)
            ws.cell(row=r_head, column=c).border = BORDA
    ws.row_dimensions[r_head].height = 26

    for i in range(n_linhas):
        r = r_head + 1 + i
        ant = None if i == 0 else "$M{}".format(r - 1)
        ws.cell(row=r, column=13, value=_proxima(rng_ref, chave, "$M{}".format(r), ant))
        for (c1, c2, tpl, fmt, al) in campos:
            if c2 > c1:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
            cel = ws.cell(row=r, column=c1, value=tpl.format(m="$M{}".format(r), k=i + 1))
            if fmt:
                cel.number_format = fmt
            cel.font = f(size=9)
            cel.alignment = Alignment(horizontal=al, vertical="center")
        _borda(ws, r, 2, 10, C_CINZA if i % 2 else None)
        ws.row_dimensions[r].height = 17
    return r_head + n_linhas


def aba_ficha(wb):
    ws = wb.create_sheet("FICHA DO COLABORADOR")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 30, "C": 22, "D": 20, "E": 13, "F": 8,
                  "G": 22, "H": 20, "I": 15, "J": 15, "K": 3, "L": 3, "M": 9})

    titulo(ws, "B2", "FICHA DO COLABORADOR — CONVÊNIO HAPVIDA", span=9, size=16)

    ws["B4"] = "🔎 INFORME O RE:"
    ws["B4"].font = f(bold=True, size=13, color=C_NAVY_D)
    inp = ws["C4"]
    inp.value = 4721
    inp.font = f(bold=True, size=16, color=C_AZUL_TXT)
    inp.fill = fill("FFFF99")
    inp.alignment = Alignment(horizontal="center", vertical="center")
    inp.border = Border(left=Side("medium", color=C_NAVY), right=Side("medium", color=C_NAVY),
                        top=Side("medium", color=C_NAVY), bottom=Side("medium", color=C_NAVY))
    inp.comment = Comment("Digite aqui o RE do colaborador. Toda a ficha se atualiza sozinha.",
                          "Conciliação")
    ws.row_dimensions[4].height = 30
    ws.merge_cells("E4:J4")
    ws["E4"] = ('=IF(RE_FICHA="","← Digite o RE do colaborador nesta célula amarela",'
                'IF(FICHA_IDX="","⚠️ RE não encontrado na base do Hapvida",'
                'INDEX(TIT_NOME,FICHA_IDX)&"   —   "&INDEX(TIT_VINCULO,FICHA_IDX)))')
    ws["E4"].font = f(bold=True, size=12, color=C_TEAL)
    ws["E4"].alignment = Alignment(vertical="center", indent=1)

    ws["M3"] = '=IF($C$4="","",IFERROR(--$C$4,$C$4))'
    ws["M4"] = '=IFERROR(MATCH(RE_FICHA,TIT_RE,0),"")'
    wb.defined_names.add(DefinedName("RE_FICHA", attr_text="'FICHA DO COLABORADOR'!$M$3"))
    wb.defined_names.add(DefinedName("FICHA_IDX", attr_text="'FICHA DO COLABORADOR'!$M$4"))

    def val(nome, fmt=None):
        return '=IF(FICHA_IDX="","—",INDEX({},FICHA_IDX))'.format(nome), fmt

    # ---------------- TITULAR ----------------
    titulo(ws, "B6", "TITULAR", span=9, size=11, bg=C_TEAL)
    sis_tit = 'MATCH(RE_FICHA&"|TITULAR",SIS_CHORD,0)' 
    pares = [
        ("RE", '=IF(FICHA_IDX="","—",INDEX(TIT_RE,FICHA_IDX))', None,
         "Nome do colaborador", '=IF(FICHA_IDX="","—",INDEX(TIT_NOME,FICHA_IDX))', None),
        ("Contrato", '=IF(FICHA_IDX="","—",INDEX(TIT_CONTRATO,FICHA_IDX))', None,
         "Vínculo / tipo de base", '=IF(FICHA_IDX="","—",INDEX(TIT_VINCULO,FICHA_IDX))', None),
        ("Início no plano", '=IF(FICHA_IDX="","—",INDEX(TIT_INICIO,FICHA_IDX))', DATA,
         "Consta na base da empresa?", '=IF(FICHA_IDX="","—",INDEX(TIT_NOSIS,FICHA_IDX))', None),
        ("Unidade (sistema da empresa)", '=IFERROR(INDEX(SIS_UNID,{}),"—")'.format(sis_tit), None,
         "Situação (sistema da empresa)", '=IFERROR(INDEX(SIS_SIT,{}),"—")'.format(sis_tit), None),
        ("Qtd. de dependentes (Hapvida)", '=IF(FICHA_IDX="","—",INDEX(TIT_QDEP,FICHA_IDX))', INT,
         "Qtd. de agregados (Hapvida)", '=IF(FICHA_IDX="","—",INDEX(TIT_QAGR,FICHA_IDX))', INT),
        ("Qtd. de dependentes (sistema)", '=IF(FICHA_IDX="","—",INDEX(TIT_QDEPSIS,FICHA_IDX))', INT,
         "Qtd. de agregados (sistema)", '=IF(FICHA_IDX="","—",INDEX(TIT_QAGRSIS,FICHA_IDX))', INT),
    ]
    for i, (l1, f1, fmt1, l2, f2, fmt2) in enumerate(pares):
        r = 7 + i
        ws.cell(row=r, column=2, value=l1).font = f(size=9, color="333333")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        c = ws.cell(row=r, column=3, value=f1)
        c.font = f(bold=True, size=10)
        if fmt1:
            c.number_format = fmt1
        ws.cell(row=r, column=7, value=l2).font = f(size=9, color="333333")
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
        c = ws.cell(row=r, column=8, value=f2)
        c.font = f(bold=True, size=10)
        if fmt2:
            c.number_format = fmt2
        _borda(ws, r, 2, 10, C_CINZA if i % 2 else None)

    sit_dep = ('=IF({m}="","",IF(INDEX(HAP_IDXSIS,{m})<>"","🟢 Também no sistema",'
               'IF(INDEX(HAP_RENOSIS,{m})="NÃO","🔵 RE não localizado no sistema","🔴 Só no Hapvida"))'
               '&IF(INDEX(HAP_ELEG,{m})="SIM"  ," | ⚠️ acima da idade limite",""))')

    # ---------------- DEPENDENTES HAPVIDA ----------------
    titulo(ws, "B14", "DEPENDENTES — HAPVIDA", span=9, size=11, bg=C_TEAL)
    _tabela_ficha(
        ws, 15,
        [(2, 3, "Nome"), (4, 4, "Grau"), (5, 5, "Nascimento"), (6, 6, "Idade"),
         (7, 8, "Situação"), (9, 10, "Valor no Combo")],
        12,
        [(2, 3, '=IF({m}="","",INDEX(HAP_NOME,{m}))', None, "left"),
         (4, 4, '=IF({m}="","",INDEX(HAP_PAREN,{m}))', None, "center"),
         (5, 5, '=IF({m}="","",INDEX(HAP_DTN,{m}))', DATA, "center"),
         (6, 6, '=IF({m}="","",INDEX(HAP_IDADE,{m}))', INT, "center"),
         (7, 8, sit_dep, None, "left"),
         (9, 10, '=IF({m}="","",INDEX(COMBO_VLR,MATCH(MIN({k},5),COMBO_QTD,0))'
                 '-INDEX(COMBO_VLR,MATCH(MIN({k}-1,5),COMBO_QTD,0)))', MOEDA, "right")],
        'RE_FICHA&"|DEPENDENTE"', "HAP_CHORD")
    ws["B28"] = ('A coluna "Valor no Combo" mostra quanto cada dependente acrescenta à faixa — '
                 'a soma dá exatamente o valor do Combo Familiar. O Combo nunca é '
                 'valor unitário × quantidade.')
    ws["B28"].font = f(italic=True, size=8, color=C_VERM)
    ws.merge_cells("B28:J28")

    # ---------------- AGREGADOS HAPVIDA ----------------
    titulo(ws, "B30", "AGREGADOS — HAPVIDA (responsabilidade do colaborador)", span=9, size=11, bg=C_LARANJ)
    _tabela_ficha(
        ws, 31,
        [(2, 3, "Nome"), (4, 4, "Condição"), (5, 5, "Nascimento"), (6, 6, "Idade"),
         (7, 8, "Situação"), (9, 10, "Valor")],
        6,
        [(2, 3, '=IF({m}="","",INDEX(HAP_NOME,{m}))', None, "left"),
         (4, 4, '=IF({m}="","",INDEX(HAP_PAREN,{m}))', None, "center"),
         (5, 5, '=IF({m}="","",INDEX(HAP_DTN,{m}))', DATA, "center"),
         (6, 6, '=IF({m}="","",INDEX(HAP_IDADE,{m}))', INT, "center"),
         (7, 8, '=IF({m}="","",IF(INDEX(HAP_IDXSIS,{m})<>"","🟢 Também no sistema",'
                'IF(INDEX(HAP_RENOSIS,{m})="NÃO","🔵 RE não localizado no sistema","🔴 Só no Hapvida")))',
          None, "left"),
         (9, 10, '=IF({m}="","",VLR_AGREGADO)', MOEDA, "right")],
        'RE_FICHA&"|AGREGADO"', "HAP_CHORD")

    # ---------------- SISTEMA DA EMPRESA ----------------
    titulo(ws, "B39", "CADASTRO NO SISTEMA DA EMPRESA", span=9, size=11, bg=C_TEAL)
    _tabela_ficha(
        ws, 40,
        [(2, 3, "Nome"), (4, 4, "Grau / Condição"), (5, 5, "Nascimento"), (6, 6, "Idade"),
         (7, 8, "Está no Hapvida?"), (9, 10, "Situação")],
        14,
        [(2, 3, '=IF({m}="","",INDEX(SIS_NOME,{m}))', None, "left"),
         (4, 4, '=IF({m}="","",INDEX(SIS_GRAU,{m}))', None, "center"),
         (5, 5, '=IF({m}="","",INDEX(SIS_DTN,{m}))', DATA, "center"),
         (6, 6, '=IF({m}="","",INDEX(SIS_IDADE,{m}))', INT, "center"),
         (7, 8, '=IF({m}="","",IF(INDEX(SIS_NOHAP,{m})="SIM","🟢 SIM","🟡 NÃO — incluir"))', None, "left"),
         (9, 10, '=IF({m}="","",INDEX(SIS_TIPO,{m})&IF(INDEX(SIS_ELEG,{m})="SIM"," (acima da idade → agregado)",""))',
          None, "left")],
        "RE_FICHA", "SIS_RE")
    ws["B55"] = ("Este bloco fica vazio enquanto a aba SISTEMA não for preenchida com a base da empresa.")
    ws["B55"].font = f(italic=True, size=8, color="555555")
    ws.merge_cells("B55:J55")

    # ---------------- DIVERGENCIAS ----------------
    titulo(ws, "B57", "⚠️ DIVERGÊNCIAS CADASTRAIS", span=9, size=11, bg=C_VERM)
    _tabela_ficha(
        ws, 58,
        [(2, 3, "Pessoa"), (4, 5, "Natureza"), (6, 7, "Situação"), (8, 10, "O que verificar")],
        12,
        [(2, 3, '=IF({m}="","",INDEX(CONC_PESSOA,{m}))', None, "left"),
         (4, 5, '=IF({m}="","",INDEX(CONC_NAT,{m}))', None, "left"),
         (6, 7, '=IF({m}="","",INDEX(CONC_SIT,{m}))', None, "left"),
         (8, 10, '=IF({m}="","",INDEX(CONC_OBS,{m}))', None, "left")],
        'RE_FICHA&"|DIV"', "CONC_CHDIV")
    pinta_status(ws, "F59:F70")
    return ws


def ficha_custo_e_simulacao(ws):
    """Blocos de custo atual e cenario atual x apos regularizacao da FICHA."""
    def linha(r, label, formula, fmt=MOEDA, destaque=False):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        cl = ws.cell(row=r, column=2, value=label)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
        cv = ws.cell(row=r, column=8, value=formula)
        cv.number_format = fmt
        cv.alignment = Alignment(horizontal="right", indent=1)
        if destaque:
            cl.font = f(bold=True, size=11, color="FFFFFF")
            cv.font = f(bold=True, size=12, color="FFFFFF")
            _borda(ws, r, 2, 10, C_NAVY)
        else:
            cl.font = f(size=10)
            cv.font = f(bold=True, size=10)
            _borda(ws, r, 2, 10)
        cl.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 18

    titulo(ws, "B72", "💰 CUSTO ATUAL", span=9, size=11, bg=C_TEAL)
    v = lambda n: '=IF(FICHA_IDX="",0,INDEX({},FICHA_IDX))'.format(n)
    linha(73, "Titular — 100% custeado pela EMPRESA", v("TIT_VTIT"))
    linha(74, '="Combo Familiar ("&IF(FICHA_IDX="",0,INDEX(TIT_QDEP,FICHA_IDX))&'
              '" dependente(s)) — responsabilidade do COLABORADOR"', v("TIT_VCOMBO"))
    linha(75, '="Agregados ("&IF(FICHA_IDX="",0,INDEX(TIT_QAGR,FICHA_IDX))&" × "&'
              'TEXT(VLR_AGREGADO,"R$ #,##0.00")&") — responsabilidade do COLABORADOR"', v("TIT_VAGR"))
    linha(76, "TOTAL DO CONVÊNIO", v("TIT_TOTAL"), destaque=True)
    linha(77, "➤ CUSTO DA EMPRESA", v("TIT_EMPRESA"))
    linha(78, "➤ CUSTO DO COLABORADOR", v("TIT_COLAB"))
    linha(79, "Valor efetivamente cobrado no arquivo do Hapvida", v("TIT_COBRADO"))
    linha(80, "Diferença (cobrado − calculado pelo contrato)", v("TIT_DIF"))
    ws.cell(row=77, column=2).font = f(bold=True, size=10, color=C_VERDE)
    ws.cell(row=78, column=2).font = f(bold=True, size=10, color=C_LARANJ)

    # ------- cenario -------
    titulo(ws, "B82", "📊 CENÁRIO ATUAL × CENÁRIO APÓS REGULARIZAÇÃO", span=9, size=11, bg=C_TEAL)
    for c1, c2, txt in ((2, 5, "Informação"), (6, 7, "Atual"), (8, 10, "Após Inclusão")):
        ws.merge_cells(start_row=83, start_column=c1, end_row=83, end_column=c2)
        cel = ws.cell(row=83, column=c1, value=txt)
        cel.font = f(bold=True, color="FFFFFF", size=9)
        cel.alignment = Alignment(horizontal="center", vertical="center")
        _borda(ws, 83, c1, c2, C_NAVY)
    ws.row_dimensions[83].height = 22

    linhas = [
        ("Dependentes", '=IF(FICHA_IDX="",0,INDEX(TIT_QDEP,FICHA_IDX))',
         '=$F$84+IF(FICHA_IDX="",0,INDEX(TIT_DEPINC,FICHA_IDX))', INT),
        ("Agregados", '=IF(FICHA_IDX="",0,INDEX(TIT_QAGR,FICHA_IDX))',
         '=$F$85+IF(FICHA_IDX="",0,INDEX(TIT_AGRINC,FICHA_IDX))', INT),
        ("Titular (custeado pela empresa)", '=IF(FICHA_IDX="",0,VLR_TITULAR)',
         '=IF(FICHA_IDX="",0,VLR_TITULAR)', MOEDA),
        ("Combo Familiar", '=IF(FICHA_IDX="",0,INDEX(TIT_VCOMBO,FICHA_IDX))',
         '=INDEX(COMBO_VLR,MATCH(MIN($H$84,5),COMBO_QTD,0))', MOEDA),
        ("Agregados", '=IF(FICHA_IDX="",0,INDEX(TIT_VAGR,FICHA_IDX))',
         '=$H$85*VLR_AGREGADO', MOEDA),
        ("TOTAL", "=SUM($F$86:$F$88)", "=SUM($H$86:$H$88)", MOEDA),
        ("Impacto mensal", '="—"', "=$H$89-$F$89", MOEDA),
        ("Impacto anual", '="—"', "=($H$89-$F$89)*12", MOEDA),
    ]
    for i, (rot, at, dep, fmt) in enumerate(linhas):
        r = 84 + i
        forte = rot in ("TOTAL", "Impacto mensal", "Impacto anual")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        cl = ws.cell(row=r, column=2, value=rot)
        cl.font = f(bold=forte, size=10)
        cl.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        ca = ws.cell(row=r, column=6, value=at)
        ca.number_format = fmt
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
        cd = ws.cell(row=r, column=8, value=dep)
        cd.number_format = fmt
        for c in (ca, cd):
            c.font = f(bold=True, size=10)
            c.alignment = Alignment(horizontal="right", indent=1)
        bg = None
        if rot == "TOTAL":
            bg = C_CINZA2
        if rot.startswith("Impacto"):
            bg = F_VERM
            cd.font = f(bold=True, size=11, color=C_VERM)
        _borda(ws, r, 2, 10, bg)
        ws.row_dimensions[r].height = 18

    ws["B93"] = ("Regra aplicada: dependente com até a idade limite entra no COMBO FAMILIAR "
                 "(muda a faixa, nunca multiplica o valor). Acima da idade limite entra como "
                 "AGREGADO, somando o valor cheio por pessoa. Todo o aumento é de "
                 "responsabilidade do colaborador.")
    ws["B93"].font = f(italic=True, size=9, color=C_VERM)
    ws.merge_cells("B93:J94")
    ws["B93"].alignment = Alignment(wrap_text=True, vertical="top")

    ws.print_area = "A1:K94"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.column_dimensions.group("L", "M", hidden=True)


# ==========================================================================
# ABA: INSTRUCOES
# ==========================================================================
def aba_instrucoes(wb, n_hap_dados, n_titulares, contratos):
    ws = wb.create_sheet("INSTRUÇÕES")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 34, "C": 100})
    titulo(ws, "B2", "CONCILIAÇÃO CADASTRAL E FINANCEIRA — CONVÊNIO HAPVIDA", span=2, size=16)
    ws["B3"] = "Ferramenta de Gestão de Benefícios / RH — leia esta página antes de usar."
    ws["B3"].font = f(italic=True, size=10, color="555555")

    secoes = [
        ("COMO ATUALIZAR AS BASES", [
            ("Aba HAPVIDA",
             "Cole a exportação do convênio nas colunas A a J, a partir da linha 2. "
             "As colunas K a AA são o motor de cálculo — não digite nelas. "
             "Já vêm {} linhas de fórmula prontas.".format(N_HAP)),
            ("Aba SISTEMA",
             "Cole a base do sistema da empresa nas colunas A a H, a partir da linha 4 "
             "(a linha 4 traz um exemplo do formato — apague-a). Uma linha por pessoa: "
             "o próprio colaborador (GRAU = TITULAR) e cada dependente/agregado. "
             "Já vêm {} linhas de fórmula prontas.".format(N_SIS)),
            ("Se a base crescer",
             "Se a nova base tiver mais linhas do que a capacidade acima, selecione a última "
             "linha de fórmulas e arraste para baixo. Depois amplie os intervalos em "
             "Fórmulas ▸ Gerenciador de Nomes."),
            ("Recalcular", "Tudo é fórmula viva: ao colar as bases, todas as abas se atualizam. "
                           "Se precisar forçar, pressione F9 (ou Ctrl+Alt+F9)."),
        ]),
        ("O QUE CADA ABA FAZ", [
            ("RESUMO GERENCIAL", "Painel para a diretoria: colaboradores, dependentes, agregados, "
                                 "faixa etária, financeiro e impacto da regularização."),
            ("FICHA DO COLABORADOR", "Digite um RE na célula amarela e a ficha inteira se monta: "
                                     "titular, dependentes, agregados, cadastro do sistema, "
                                     "divergências, custo atual e simulação."),
            ("CONCILIAÇÃO", "Uma linha por PESSOA, cruzando as duas bases, com a situação "
                            "classificada e a observação do que verificar. Use os filtros do cabeçalho."),
            ("TITULARES", "Uma linha por RE — a espinha dorsal. Quantidades, valores, situação e alertas."),
            ("SIMULAÇÃO DE IMPACTO", "Quanto custaria incluir no Hapvida tudo o que já consta no "
                                     "sistema da empresa, com aumento mensal e anual por colaborador."),
            ("RESUMO POR CONTRATO", "Titulares, dependentes, agregados e valores agrupados por contrato."),
            ("ALERTAS", "Só os REs que exigem ação — divergência cadastral, diferença de cobrança "
                        "ou dependente acima da idade limite."),
            ("PARÂMETROS", "Valores do contrato e regras. É o único lugar onde se mudam preços e a "
                           "idade limite — tudo o mais lê daqui."),
        ]),
        ("CLASSIFICAÇÃO DAS SITUAÇÕES", [
            ("🟢 CADASTRO OK", "Pessoa cadastrada corretamente no Hapvida e no sistema da empresa."),
            ("🟡 INCLUSÃO NECESSÁRIA", "Está no sistema da empresa e não está no Hapvida."),
            ("🔴 DIVERGÊNCIA", "Está no Hapvida e não está no sistema da empresa (com o RE localizado)."),
            ("🟠 AGREGADO", "Natureza da pessoa — tratada financeiramente à parte do Combo Familiar."),
            ("⚠️ DIVERGÊNCIA DE DADOS", "Mesma pessoa nas duas bases, mas com nome, grau de "
                                        "parentesco ou data de nascimento diferentes."),
            ("🔵 RE NÃO LOCALIZADO", "O RE não existe na base da empresa. NÃO é tratado como erro: "
                                     "o vínculo é classificado pelo CONTRATO do Hapvida (PJ, Acordo, "
                                     "Sindicato) e, quando não dá para identificar, fica como "
                                     "NÃO LOCALIZADO – VERIFICAR."),
            ("⚪ SEM CONVÊNIO", "Pessoa do sistema da empresa cujo RE não tem nenhum registro no Hapvida."),
        ]),
        ("REGRAS FINANCEIRAS APLICADAS", [
            ("Titular", "R$ 151,73 por titular — 100% custeado pela EMPRESA."),
            ("Combo Familiar", "Definido pela FAIXA da quantidade total de dependentes "
                               "(1→R$ 239,93; 2→R$ 446,62; 3→R$ 579,83; 4→R$ 663,57; 5 ou mais→R$ 846,86). "
                               "NUNCA é valor unitário × quantidade. Responsabilidade do COLABORADOR."),
            ("Agregados", "R$ 544,95 por agregado, somados individualmente e sempre separados do "
                          "Combo. Responsabilidade do COLABORADOR."),
            ("Regra de idade", "Dependente acima de 25 anos é sinalizado como elegível a AGREGADO. "
                               "Por padrão a regra NÃO se aplica ao CÔNJUGE (que é dependente "
                               "independentemente da idade). Isso é configurável em PARÂMETROS."),
            ("Idades", "Recalculadas a partir da data de nascimento com a data-base de PARÂMETROS. "
                       "A coluna 'idade' que vem no arquivo do Hapvida é apenas o retrato da data "
                       "da extração e por isso não é usada nos cálculos."),
        ]),
        ("CHAVE DE CRUZAMENTO", [
            ("Identificador principal", "RE (matrícula) do colaborador."),
            ("Identificação da pessoa", "1º) RE + data de nascimento. 2º) RE + nome normalizado "
                                        "(maiúsculas, sem acentos e sem pontuação). Quem casa pela "
                                        "data mas não pelo nome aparece como ⚠️ DIVERGÊNCIA DE DADOS, "
                                        "com o nome do sistema na observação."),
        ]),
        ("ESTADO DESTA CÓPIA", [
            ("Base HAPVIDA carregada", "{:,} linhas · {:,} titulares · contratos: {}".format(
                n_hap_dados, n_titulares, ", ".join(contratos)).replace(",", ".")),
            ("Base SISTEMA", "VAZIA — não foi fornecida. Enquanto não for colada, todo mundo do "
                             "Hapvida aparece como 'RE não localizado na base da empresa' e a "
                             "simulação fica zerada. Isso é esperado: cole a base e os números aparecem."),
        ]),
        ("FILTROS, TABELA DINÂMICA E SEGMENTAÇÃO", [
            ("Já pronto", "Todas as abas de lista têm filtro no cabeçalho e painéis congelados."),
            ("Tabela dinâmica", "Para montar uma, selecione uma célula da aba TITULARES ou "
                                "CONCILIAÇÃO e vá em Inserir ▸ Tabela Dinâmica. Em seguida "
                                "Análise de Tabela Dinâmica ▸ Inserir Segmentação de Dados para "
                                "filtrar por CONTRATO ou SITUAÇÃO com um clique."),
            ("Por que fórmulas e não Power Query", "As fórmulas recalculam sozinhas ao colar as "
                                                   "bases, sem depender de atualizar consulta, e "
                                                   "abrem em qualquer versão do Excel (2016 em "
                                                   "diante), no LibreOffice e no Google Sheets."),
        ]),
    ]

    r = 5
    for tit, itens in secoes:
        titulo(ws, "B{}".format(r), tit, span=2, size=11, bg=C_TEAL)
        r += 1
        for rot, txt in itens:
            cr = ws.cell(row=r, column=2, value=rot)
            cr.font = f(bold=True, size=9)
            cr.alignment = Alignment(vertical="top", wrap_text=True)
            cr.border = BORDA
            ct = ws.cell(row=r, column=3, value=txt)
            ct.font = f(size=9)
            ct.alignment = Alignment(vertical="top", wrap_text=True)
            ct.border = BORDA
            ws.row_dimensions[r].height = max(15, 12 * (1 + len(txt) // 95))
            r += 1
        r += 1
    return ws


# ==========================================================================
# MAIN
# ==========================================================================
ORDEM = ["INSTRUÇÕES", "RESUMO GERENCIAL", "FICHA DO COLABORADOR", "CONCILIAÇÃO",
         "TITULARES", "SIMULAÇÃO DE IMPACTO", "RESUMO POR CONTRATO", "ALERTAS",
         "HAPVIDA", "SISTEMA", "PARÂMETROS"]
CORES_ABA = {"INSTRUÇÕES": "808080", "RESUMO GERENCIAL": C_NAVY, "FICHA DO COLABORADOR": C_TEAL,
             "CONCILIAÇÃO": C_NAVY, "TITULARES": C_NAVY, "SIMULAÇÃO DE IMPACTO": C_LARANJ,
             "RESUMO POR CONTRATO": C_TEAL, "ALERTAS": C_VERM,
             "HAPVIDA": "9AA5B1", "SISTEMA": "9AA5B1", "PARÂMETROS": "9AA5B1"}


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    csv_in = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, "dados", "base_hapvida.csv")
    saida = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        base, "Conciliacao_Hapvida_x_Sistema.xlsx")

    dados = ler_hapvida(csv_in)
    if not dados:
        raise SystemExit("Nenhuma linha lida de {}".format(csv_in))

    vistos, contratos = set(), []
    for d in dados:
        c = d.get("CONTRATO", "").strip()
        if c and c not in vistos:
            vistos.add(c)
            contratos.append(c)
    contratos.sort()
    n_tit = sum(1 for d in dados if d.get("parentesco", "").upper() == "TITULAR")

    print("Lidas {} linhas · {} titulares · contratos: {}".format(
        len(dados), n_tit, ", ".join(contratos)))
    if len(dados) > N_HAP:
        raise SystemExit("A base tem {} linhas e a capacidade configurada é {}. "
                         "Aumente N_HAP.".format(len(dados), N_HAP))

    wb = Workbook()
    wb.remove(wb.active)
    wb._named_styles["Normal"].font = Font(name=FONTE, size=10)

    aba_parametros(wb)
    print("  · PARÂMETROS")
    aba_hapvida(wb, dados)
    print("  · HAPVIDA")
    aba_sistema(wb)
    print("  · SISTEMA")
    aba_titulares(wb)
    print("  · TITULARES")
    aba_conciliacao(wb)
    print("  · CONCILIAÇÃO")
    aba_simulacao(wb)
    print("  · SIMULAÇÃO DE IMPACTO")
    aba_resumo_contrato(wb, contratos)
    print("  · RESUMO POR CONTRATO")
    aba_resumo(wb)
    print("  · RESUMO GERENCIAL")
    aba_alertas(wb)
    print("  · ALERTAS")
    ws_ficha = aba_ficha(wb)
    ficha_custo_e_simulacao(ws_ficha)
    print("  · FICHA DO COLABORADOR")
    aba_instrucoes(wb, len(dados), n_tit, contratos)
    print("  · INSTRUÇÕES")

    definir_nomes(wb)

    wb._sheets.sort(key=lambda s: ORDEM.index(s.title) if s.title in ORDEM else 99)
    for nome, cor in CORES_ABA.items():
        if nome in wb.sheetnames:
            wb[nome].sheet_properties.tabColor = cor
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True

    wb.save(saida)
    print("Gravado: {} ({:.1f} MB)".format(saida, os.path.getsize(saida) / 1e6))
    return saida


if __name__ == "__main__":
    main()
