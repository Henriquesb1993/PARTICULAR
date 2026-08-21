# -*- coding: utf-8 -*-
"""
Gerador da planilha de Conciliacao Cadastral e Financeira - Convenio Medico Hapvida.

Le as duas bases (aba HAPVIDA e aba SISTEMA JB do arquivo de origem) e monta um
workbook .xlsx totalmente orientado a formulas: ao substituir as bases, todas as
demais abas se recalculam sozinhas.

Uso:  python3 gerar_planilha.py [arquivo_origem.xlsx] [saida.xlsx]
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

from compactar import compactar

# --------------------------------------------------------------------------
# 1. PARAMETROS DE NEGOCIO  (fonte: especificacao do RH)
# --------------------------------------------------------------------------
VLR_TITULAR = 151.73          # 100% custeado pela empresa
VLR_AGREGADO = 544.95         # por agregado, custeado pelo colaborador
COMBO = [(0, 0.00), (1, 239.93), (2, 446.62), (3, 579.83), (4, 663.57), (5, 846.86)]
IDADE_LIMITE_DEP = 25

# Relacao/parentesco -> tipo e sujeicao a regra de idade. Editavel em PARAMETROS.
MAPA_RELACAO = [
    ("TITULAR", "TITULAR", "NÃO"),
    ("FILHO(A)", "DEPENDENTE", "SIM"),
    ("FILHO UNIVERS", "DEPENDENTE", "SIM"),
    ("ENTEADO", "DEPENDENTE", "SIM"),
    ("TUTELADO", "DEPENDENTE", "SIM"),
    ("MENOR POBRE", "DEPENDENTE", "SIM"),
    ("OUTROS", "DEPENDENTE", "SIM"),
    ("CONJUGE", "DEPENDENTE", "NÃO"),
    ("CÔNJUGE", "DEPENDENTE", "NÃO"),
    ("COMPANHEIRO(A)", "DEPENDENTE", "NÃO"),
    ("AGREGADO", "AGREGADO", "NÃO"),
    ("AGREGADO(A)", "AGREGADO", "NÃO"),
    ("EX CONJUGE", "NÃO ELEGÍVEL", "NÃO"),
]
N_MAPA = len(MAPA_RELACAO) + 7          # linhas em branco para o RH acrescentar

# Capacidades (linhas de formula pre-preenchidas)
N_HAP = 9400                  # base atual: 9.181 linhas
N_SIS = 16600                 # base atual: 16.275 linhas
N_TIT = 5950                  # base atual: 5.860 titulares
N_CONC_B = 8000               # bloco "somente sistema" (atual: 7.773)
N_ALERTA = 5950
N_FICHA = 20                  # linhas do quadro unico da ficha (maior familia: 14)

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
C_AZUL_TXT = "0000FF"
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
    r, c0 = ws[cell].row, ws[cell].column
    if span > 1:
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + span - 1)
        for i in range(1, span):
            ws.cell(row=r, column=c0 + i).fill = fill(bg)
    ws.row_dimensions[r].height = max(22, size + 10)


def larguras(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


def _borda(ws, r, c1, c2, bg=None):
    for c in range(c1, c2 + 1):
        cel = ws.cell(row=r, column=c)
        cel.border = BORDA
        if bg:
            cel.fill = fill(bg)


# --------------------------------------------------------------------------
# 3. FORMULAS AUXILIARES REUTILIZADAS
# --------------------------------------------------------------------------
_ACENTOS = [("Á", "A"), ("À", "A"), ("Â", "A"), ("Ã", "A"), ("É", "E"), ("Ê", "E"),
            ("Í", "I"), ("Ó", "O"), ("Ô", "O"), ("Õ", "O"), ("Ú", "U"), ("Ç", "C"),
            (".", ""), ("-", " ")]


def norm_formula(ref):
    """Nome padronizado para comparacao: maiusculas, sem acento e sem pontuacao.

    Usada apenas nas colunas auxiliares — os nomes originais das bases nunca
    sao alterados.
    """
    expr = "UPPER(TRIM({}))".format(ref)
    for a, b in _ACENTOS:
        expr = 'SUBSTITUTE({},"{}","{}")'.format(expr, a, b)
    return "TRIM({})".format(expr)


def primeiro_nome(ref):
    return 'LEFT({r},FIND(" ",{r}&" ")-1)'.format(r=ref)


def ultimo_nome(ref):
    return 'TRIM(RIGHT(SUBSTITUTE({r}," ",REPT(" ",60)),60))'.format(r=ref)


def data_formula(ref):
    """Converte `ref` em data serial aceitando data real, numero de serie OU
    texto DD/MM/AAAA. Nao usa DATA.VALOR, que depende do idioma do Excel."""
    return ('IF({r}="","",IF(ISNUMBER({r}),{r},IFERROR(DATE(VALUE(MID({r},7,4)),'
            'VALUE(MID({r},4,2)),VALUE(LEFT({r},2))),"")))').format(r=ref)


def tipo_formula(rel_ref):
    """TIPO da pessoa a partir da relacao/parentesco, pela tabela de PARAMETROS."""
    return ('IF({r}="","",IFERROR(INDEX(MAPA_TIPO,MATCH(UPPER(TRIM({r})),MAPA_REL,0)),"DEPENDENTE"))'
            .format(r=rel_ref))


def regra_idade_formula(rel_ref):
    return ('IFERROR(INDEX(MAPA_IDADE,MATCH(UPPER(TRIM({r})),MAPA_REL,0)),"SIM")'.format(r=rel_ref))


def cascata(ch_dt, ch_nm, ch_pu, pref_pn, r_dt, r_nm, r_pu):
    """Cascata de correspondencia: 1) nascimento 2) nome completo
    3) primeiro+ultimo sobrenome 4) so primeiro nome (curinga)."""
    return ('IFERROR(MATCH({cd},{rd},0),IFERROR(MATCH({cn},{rn},0),'
            'IFERROR(MATCH({cp},{rp},0),IFERROR(MATCH({pp},{rp},0),""))))').format(
                cd=ch_dt, rd=r_dt, cn=ch_nm, rn=r_nm, cp=ch_pu, pp=pref_pn, rp=r_pu)


# --------------------------------------------------------------------------
# 4. LEITURA DAS BASES DE ORIGEM
# --------------------------------------------------------------------------
def ler_origem(caminho):
    """Devolve (linhas_hapvida, linhas_sistema) como listas de listas."""
    wb = load_workbook(caminho, data_only=True, read_only=True)
    nomes = {n.strip().upper(): n for n in wb.sheetnames}

    def acha(*cands):
        for c in cands:
            for k, n in nomes.items():
                if k.startswith(c):
                    return wb[n]
        return None

    wh = acha("HAPVIDA")
    wsis = acha("SISTEMA", "JB")
    if wh is None:
        raise SystemExit("Aba HAPVIDA não encontrada em {}".format(caminho))
    if wsis is None:
        raise SystemExit("Aba SISTEMA não encontrada em {}".format(caminho))

    hap = []
    for row in wh.iter_rows(min_row=2, max_col=10, values_only=True):
        if row[1] in (None, ""):
            continue
        hap.append(list(row))
    sis = []
    for row in wsis.iter_rows(min_row=2, max_col=9, values_only=True):
        if row[0] in (None, ""):
            continue
        sis.append(list(row))
    wb.close()
    return hap, sis


def _lim(v):
    return v.strip() if isinstance(v, str) else v


# ==========================================================================
# ABA: PARAMETROS
# ==========================================================================
def aba_parametros(wb):
    ws = wb.create_sheet("PARÂMETROS")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 34, "C": 18, "D": 20, "E": 3, "F": 62})

    titulo(ws, "B2", "PARÂMETROS DO CONVÊNIO — HAPVIDA", span=5, size=15)
    ws["B3"] = "Células em AZUL são editáveis. Toda a pasta lê estes valores."
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
            ws.cell(row=row, column=6, value=nota).font = f(size=9, color="555555")
        return c

    titulo(ws, "B5", "DATA-BASE E REGRA DE IDADE", span=5, size=11, bg=C_TEAL)
    par(6, "Data-base para cálculo das idades", "=TODAY()", DATA,
        "Troque por uma data fixa para congelar a análise.")
    par(7, "Idade limite do dependente (anos)", IDADE_LIMITE_DEP, INT,
        "Acima disso o dependente é tratado como AGREGADO — mas só nas relações "
        "marcadas com SIM na tabela abaixo.")

    titulo(ws, "B9", "VALORES DO CONTRATO", span=5, size=11, bg=C_TEAL)
    par(10, "Valor do TITULAR (mensal)", VLR_TITULAR, MOEDA, "100% custeado pela EMPRESA.")
    par(11, "Valor de CADA AGREGADO (mensal)", VLR_AGREGADO, MOEDA,
        "Responsabilidade do COLABORADOR. Soma-se por agregado.")
    ws["C10"].comment = Comment("Fonte: especificação do RH — contrato Combo Familiar Hapvida.",
                                "Conciliação")

    titulo(ws, "B13", "COMBO FAMILIAR — VALOR PELA FAIXA DE DEPENDENTES", span=5, size=11, bg=C_TEAL)
    header_row(ws, 14, ["Qtd. de dependentes", "Valor do Combo"], start_col=2, h=24)
    for i, (q, v) in enumerate(COMBO):
        r = 15 + i
        for col, val, fmt in ((2, q, INT), (3, v, MOEDA)):
            c = ws.cell(row=r, column=col, value=val)
            c.number_format = fmt
            c.font = f(color=C_AZUL_TXT, bold=True, size=10)
            c.fill = fill("FFFFCC")
            c.border = BORDA
            c.alignment = Alignment(horizontal="center")
    r = 15 + len(COMBO)
    ws.cell(row=r, column=2, value="A faixa 5 vale para 5 OU MAIS dependentes. O valor NUNCA é "
                                   "multiplicado pela quantidade.").font = f(italic=True, size=9, color=C_VERM)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)

    titulo(ws, "B24", "RELAÇÃO / PARENTESCO → COMO A PESSOA É TRATADA", span=5, size=11, bg=C_TEAL)
    ws["B25"] = ("Vale para as duas bases. O que não estiver listado entra como DEPENDENTE "
                 "sujeito à regra de idade. Acrescente novas relações nas linhas em branco.")
    ws["B25"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("B25:F25")
    header_row(ws, 26, ["RELAÇÃO (como vem nas bases)", "TIPO",
                        "Sujeito à regra de idade?"], start_col=2, h=32)
    for i in range(N_MAPA):
        r = 27 + i
        vals = MAPA_RELACAO[i] if i < len(MAPA_RELACAO) else ("", "", "")
        for col, val in ((2, vals[0]), (3, vals[1]), (4, vals[2])):
            c = ws.cell(row=r, column=col, value=val or None)
            c.font = f(color=C_AZUL_TXT, bold=True, size=9)
            c.fill = fill("FFFFCC")
            c.border = BORDA
            c.alignment = Alignment(horizontal="center")
    ws.cell(row=27 + N_MAPA, column=2,
            value="TIPO aceita: TITULAR, DEPENDENTE, AGREGADO ou NÃO ELEGÍVEL. "
                  "NÃO ELEGÍVEL fica de fora dos cálculos e da simulação "
                  "(é o caso de EX CONJUGE).").font = f(italic=True, size=9, color=C_VERM)
    ws.merge_cells(start_row=27 + N_MAPA, start_column=2, end_row=27 + N_MAPA, end_column=6)
    ws["D27"].comment = Comment(
        "NÃO no cônjuge é proposital: cônjuge é dependente independentemente da idade. "
        "Marcar SIM aqui converteria mais de mil cônjuges em agregados.", "Conciliação")

    titulo(ws, "B49", "RATEIO", span=5, size=11, bg=C_TEAL)
    for i, (a, b) in enumerate((("Custeado pela EMPRESA", "Titular"),
                                ("Custeado pelo COLABORADOR", "Combo Familiar + Agregados"))):
        r = 50 + i
        ws.cell(row=r, column=2, value=a).font = f(bold=True, size=10)
        ws.cell(row=r, column=3, value=b).font = f(size=10)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    return ws


# ==========================================================================
# ABA: HAPVIDA
# ==========================================================================
HAP_COLS = [
    ("CONTRATO", 20), ("RE (matrícula)", 11), ("BENEFICIÁRIO", 38), ("NOME DA MÃE", 32),
    ("NASCIMENTO", 12), ("INÍCIO", 11), ("IDADE (arquivo)", 9), ("PARENTESCO", 13),
    ("PLANO", 9), ("COBRADO (arquivo)", 13),
    # ---- calculadas ----
    ("Dt. nascimento", 12), ("Idade calculada", 10), ("Tipo", 13),
    ("Nome padronizado", 32), ("Chave RE+Nasc.", 18), ("Chave RE+Nome", 30),
    ("Chave RE+1º+último", 26), ("Idx sistema", 10), ("Correspondência", 26),
    ("Está no sistema?", 11), ("RE existe no sistema?", 11), ("Elegível a agregado?", 11),
    ("Chave duplicidade", 26), ("Duplicado?", 10), ("Chave titular", 11),
    ("1º titular do RE?", 10), ("_cnt", 8), ("Seq. titular", 9),
    ("Divergência de dados?", 11),
]


def aba_hapvida(wb, dados):
    ws = wb.create_sheet("HAPVIDA")
    header_row(ws, 1, [c[0] for c in HAP_COLS], h=40)
    for i, (_, w) in enumerate(HAP_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    for i, d in enumerate(dados):
        r = i + 2
        for c in range(10):
            v = _lim(d[c]) if c < len(d) else None
            cel = ws.cell(row=r, column=c + 1, value=v)
            if c in (4, 5):
                cel.number_format = DATA
            elif c == 9:
                cel.number_format = MOEDA

    for r in range(2, N_HAP + 2):
        g = lambda c: "${}{}".format(c, r)
        put = lambda col, v: ws.cell(row=r, column=col, value=v)
        put(11, "=" + data_formula(g("E"))).number_format = DATA
        put(12, '=IF({k}="","",DATEDIF({k},DATA_BASE,"Y"))'.format(k=g("K")))
        put(13, "=" + tipo_formula(g("H")))
        put(14, '=IF({c}="","",{n})'.format(c=g("C"), n=norm_formula(g("C"))))
        put(15, '=IF({b}="","",{b}&"|"&IF({k}="","?"&ROW(),TEXT({k},"ddmmyyyy")))'.format(b=g("B"), k=g("K")))
        put(16, '=IF({b}="","",{b}&"|"&{n})'.format(b=g("B"), n=g("N")))
        put(17, '=IF({b}="","",{b}&"|"&{p}&"|"&{u})'.format(
            b=g("B"), p=primeiro_nome(g("N")), u=ultimo_nome(g("N"))))
        put(18, '=IF({b}="","",IF(COUNTA(SIS_RE)=0,"",{c}))'.format(
            b=g("B"), c=cascata(g("O"), g("P"), g("Q"),
                                '{b}&"|"&{p}&"|*"'.format(b=g("B"), p=primeiro_nome(g("N"))),
                                "SIS_CHDT", "SIS_CHNM", "SIS_CHPU")))
        put(19, ('=IF({r}="","— não encontrado",'
                 'IF(AND(N({k})>0,N({k})=N(INDEX(SIS_DTN,{r}))),"FORTE (nascimento)",'
                 'IF({n}=INDEX(SIS_NOMEPAD,{r}),"FORTE (nome completo)",'
                 'IF({q}=INDEX(SIS_CHPU,{r}),"PROVÁVEL (nome abreviado)",'
                 '"VERIFICAR (só 1º nome)"))))').format(r=g("R"), k=g("K"), n=g("N"), q=g("Q")))
        put(20, '=IF({b}="","",IF({r}="","NÃO",IF(LEFT({s},9)="VERIFICAR","VERIFICAR","SIM")))'.format(
            b=g("B"), r=g("R"), s=g("S")))
        put(21, '=IF({b}="","",IF(COUNTA(SIS_RE)=0,"NÃO",IF(COUNTIF(SIS_RE,{b})>0,"SIM","NÃO")))'.format(b=g("B")))
        put(22, '=IF({m}<>"DEPENDENTE","NÃO",IF(AND(N({l})>IDADE_LIMITE,{g}="SIM"),"SIM","NÃO"))'.format(
            m=g("M"), l=g("L"), g=regra_idade_formula(g("H"))))
        put(23, '=IF({b}="","",{o}&"|"&{n})'.format(b=g("B"), o=g("O"), n=g("N")))
        put(24, '=IF({w}="","",IF(MATCH({w},HAP_CHDUP,0)=ROW()-1,"NÃO","SIM"))'.format(w=g("W")))
        put(25, '=IF({m}<>"TITULAR","",{b}&"")'.format(m=g("M"), b=g("B")))
        put(26, '=IF({y}="","NÃO",IF(MATCH({y},HAP_CHTIT,0)=ROW()-1,"SIM","NÃO"))'.format(y=g("Y")))
        novo = 'IF({z}="SIM",1,0)'.format(z=g("Z"))
        put(27, ("=" + novo) if r == 2 else '=$AA{p}+{n}'.format(p=r - 1, n=novo))
        put(28, '=IF({z}="SIM",$AA{r},"")'.format(z=g("Z"), r=r))
        put(29, ('=IF({r}="","NÃO",IF(OR({n}<>INDEX(SIS_NOMEPAD,{r}),'
                 'AND(N({k})>0,N(INDEX(SIS_DTN,{r}))>0,N({k})<>N(INDEX(SIS_DTN,{r}))),'
                 '{m}<>INDEX(SIS_TIPO,{r})),"SIM","NÃO"))').format(
                     r=g("R"), n=g("N"), k=g("K"), m=g("M")))

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:V{}".format(N_HAP + 1)
    ws.column_dimensions.group("W", "AC", hidden=True)
    return ws


# ==========================================================================
# ABA: SISTEMA  (layout do export do sistema da empresa)
# ==========================================================================
SIS_COLS = [
    ("RE", 11), ("NOME (colaborador)", 34), ("FUNCAO", 22), ("ESTADO CIVIL", 14),
    ("DT NASCTO (titular)", 13), ("NOME DEPENDENTE", 34), ("DT NASCTO", 12),
    ("IDADE (arquivo)", 9), ("RELACAO", 16),
    # ---- calculadas ----
    ("Pessoa da linha", 34), ("Dt. nascimento", 12), ("Idade calculada", 10), ("Tipo", 13),
    ("Nome padronizado", 32), ("Chave RE+Nasc.", 18), ("Chave RE+Nome", 30),
    ("Chave RE+1º+último", 26), ("Idx Hapvida", 10),
    ("Está no Hapvida?", 13), ("RE existe no Hapvida?", 11), ("Elegível a agregado?", 11),
    ("_seq", 8),
]
SIS_INI = 4


def aba_sistema(wb, dados):
    ws = wb.create_sheet("SISTEMA")
    titulo(ws, "A1", "BASE DO SISTEMA DA EMPRESA — cole aqui as colunas A a I, a partir da "
                     "linha 4 (sem o cabeçalho). As colunas J a V são calculadas: não digite nelas.",
           span=22, size=11, bg=C_VERM)
    ws["A2"] = ("Layout do export: uma linha por PESSOA. Na linha do próprio colaborador a "
                "RELACAO é TITULAR e NOME DEPENDENTE fica vazio — o nome sai da coluna B. "
                "Nas demais linhas a pessoa é o NOME DEPENDENTE. A data de nascimento da "
                "pessoa é sempre a coluna G.")
    ws["A2"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("A2:V2")
    ws.row_dimensions[2].height = 26

    header_row(ws, 3, [c[0] for c in SIS_COLS], h=40)
    for i, (_, w) in enumerate(SIS_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    for i, d in enumerate(dados):
        r = SIS_INI + i
        for c in range(9):
            v = _lim(d[c]) if c < len(d) else None
            cel = ws.cell(row=r, column=c + 1, value=v)
            if c in (4, 6):
                cel.number_format = DATA

    last = N_SIS + SIS_INI - 1
    for r in range(SIS_INI, last + 1):
        g = lambda c: "${}{}".format(c, r)
        put = lambda col, v: ws.cell(row=r, column=col, value=v)
        put(10, '=IF({a}="","",IF(TRIM({f}&"")="",TRIM({b}&""),TRIM({f})))'.format(
            a=g("A"), f=g("F"), b=g("B")))
        put(11, "=" + data_formula(g("G"))).number_format = DATA
        put(12, '=IF({k}="","",DATEDIF({k},DATA_BASE,"Y"))'.format(k=g("K")))
        put(13, "=" + tipo_formula(g("I")))
        put(14, '=IF({j}="","",{n})'.format(j=g("J"), n=norm_formula(g("J"))))
        put(15, '=IF({a}="","",{a}&"|"&IF({k}="","?"&ROW(),TEXT({k},"ddmmyyyy")))'.format(a=g("A"), k=g("K")))
        put(16, '=IF({a}="","",{a}&"|"&{n})'.format(a=g("A"), n=g("N")))
        put(17, '=IF({a}="","",{a}&"|"&{p}&"|"&{u})'.format(
            a=g("A"), p=primeiro_nome(g("N")), u=ultimo_nome(g("N"))))
        put(18, '=IF({a}="","",{c})'.format(
            a=g("A"), c=cascata(g("O"), g("P"), g("Q"),
                                '{a}&"|"&{p}&"|*"'.format(a=g("A"), p=primeiro_nome(g("N"))),
                                "HAP_CHDT", "HAP_CHNM", "HAP_CHPU")))
        put(19, ('=IF({a}="","",IF({r}="","NÃO",IF(OR(AND(N({k})>0,N({k})=N(INDEX(HAP_DTN,{r}))),'
                 '{n}=INDEX(HAP_NOMEPAD,{r}),{q}=INDEX(HAP_CHPU,{r})),"SIM","VERIFICAR")))').format(
                     a=g("A"), r=g("R"), k=g("K"), n=g("N"), q=g("Q")))
        put(20, '=IF({a}="","",IF(COUNTIF(HAP_RE,{a})>0,"SIM","NÃO"))'.format(a=g("A")))
        put(21, '=IF({m}<>"DEPENDENTE","NÃO",IF(AND(N({l})>IDADE_LIMITE,{g}="SIM"),"SIM","NÃO"))'.format(
            m=g("M"), l=g("L"), g=regra_idade_formula(g("I"))))
        novo = 'IF({s}="NÃO",1,0)'.format(s=g("S"))
        put(22, ("=" + novo) if r == SIS_INI else '=$V{p}+{n}'.format(p=r - 1, n=novo))

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = "A3:U{}".format(last)
    ws.column_dimensions.group("V", "V", hidden=True)
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
    add("VLR_TITULAR", "'PARÂMETROS'!$C$10")
    add("VLR_AGREGADO", "'PARÂMETROS'!$C$11")
    add("COMBO_QTD", "'PARÂMETROS'!$B$15:$B${}".format(14 + len(COMBO)))
    add("COMBO_VLR", "'PARÂMETROS'!$C$15:$C${}".format(14 + len(COMBO)))
    add("MAPA_REL", "'PARÂMETROS'!$B$27:$B${}".format(26 + N_MAPA))
    add("MAPA_TIPO", "'PARÂMETROS'!$C$27:$C${}".format(26 + N_MAPA))
    add("MAPA_IDADE", "'PARÂMETROS'!$D$27:$D${}".format(26 + N_MAPA))

    h1, h2 = 2, N_HAP + 1
    for nome, col in [("HAP_CONTRATO", "A"), ("HAP_RE", "B"), ("HAP_NOME", "C"),
                      ("HAP_INICIO", "F"), ("HAP_PAREN", "H"), ("HAP_COB", "J"),
                      ("HAP_DTN", "K"), ("HAP_IDADE", "L"), ("HAP_TIPO", "M"),
                      ("HAP_NOMEPAD", "N"), ("HAP_CHDT", "O"), ("HAP_CHNM", "P"),
                      ("HAP_CHPU", "Q"), ("HAP_IDXSIS", "R"), ("HAP_CORRESP", "S"),
                      ("HAP_NOSIS", "T"), ("HAP_RENOSIS", "U"), ("HAP_ELEG", "V"),
                      ("HAP_CHDUP", "W"), ("HAP_DUP", "X"), ("HAP_CHTIT", "Y"),
                      ("HAP_1TIT", "Z"), ("HAP_SEQTIT", "AB"), ("HAP_DIVDADOS", "AC")]:
        add(nome, col_ref("HAPVIDA", col, h1, h2))

    s1, s2 = SIS_INI, N_SIS + SIS_INI - 1
    for nome, col in [("SIS_RE", "A"), ("SIS_COLAB", "B"), ("SIS_FUNCAO", "C"),
                      ("SIS_ECIVIL", "D"), ("SIS_REL", "I"), ("SIS_PESSOA", "J"),
                      ("SIS_DTN", "K"), ("SIS_IDADE", "L"), ("SIS_TIPO", "M"),
                      ("SIS_NOMEPAD", "N"), ("SIS_CHDT", "O"), ("SIS_CHNM", "P"),
                      ("SIS_CHPU", "Q"), ("SIS_IDXHAP", "R"), ("SIS_NOHAP", "S"),
                      ("SIS_RENOHAP", "T"), ("SIS_ELEG", "U"), ("SIS_SEQ", "V")]:
        add(nome, col_ref("SISTEMA", col, s1, s2))

    t1, t2 = 2, N_TIT + 1
    for nome, col in [("TIT_RE", "A"), ("TIT_NOME", "B"), ("TIT_CONTRATO", "C"),
                      ("TIT_INICIO", "D"), ("TIT_NOSIS", "E"), ("TIT_VINCULO", "F"),
                      ("TIT_QDEP", "G"), ("TIT_QAGR", "H"), ("TIT_QDEPSIS", "I"),
                      ("TIT_QAGRSIS", "J"), ("TIT_DEPOK", "K"), ("TIT_DEPSOHAP", "L"),
                      ("TIT_DEPINC", "M"), ("TIT_AGRINC", "N"), ("TIT_DEPELEG", "O"),
                      ("TIT_VTIT", "P"), ("TIT_VCOMBO", "Q"), ("TIT_VAGR", "R"),
                      ("TIT_TOTAL", "S"), ("TIT_COBRADO", "T"), ("TIT_DIF", "U"),
                      ("TIT_EMPRESA", "V"), ("TIT_COLAB", "W"), ("TIT_SITUACAO", "X"),
                      ("TIT_ALERTA", "Y"), ("TIT_DIVDADOS", "Z"), ("TIT_VERIF", "AA"),
                      ("TIT_SEQALERTA", "AD")]:
        add(nome, col_ref("TITULARES", col, t1, t2))


# ==========================================================================
# ABA: TITULARES  (um registro por RE)
# ==========================================================================
TIT_COLS = [
    ("RE", 10), ("COLABORADOR (titular)", 36), ("CONTRATO", 20), ("INÍCIO NO PLANO", 12),
    ("RE existe no sistema?", 11), ("VÍNCULO / CLASSIFICAÇÃO", 30),
    ("Dep. HAPVIDA", 10), ("Agreg. HAPVIDA", 10), ("Dep. SISTEMA", 10), ("Agreg. SISTEMA", 10),
    ("Dep. nas 2 bases", 10), ("Dep. só HAPVIDA", 10),
    ("A incluir como DEPENDENTE", 11), ("A incluir como AGREGADO", 11),
    ("Dep. HAPVIDA acima da idade", 11),
    ("Valor TITULAR", 13), ("Valor COMBO", 13), ("Valor AGREGADOS", 13), ("TOTAL calculado", 14),
    ("Cobrado no arquivo", 14), ("Diferença de cobrança", 14),
    ("Custo EMPRESA", 13), ("Custo COLABORADOR", 14),
    ("SITUAÇÃO CADASTRAL", 40), ("ALERTAS", 62),
    ("Divergências de dados", 11), ("A verificar correspondência", 11),
    ("_soagr", 8), ("_idx", 8), ("_seqal", 8),
]


def aba_titulares(wb):
    ws = wb.create_sheet("TITULARES")
    header_row(ws, 1, [c[0] for c in TIT_COLS], h=44)
    for i, (_, w) in enumerate(TIT_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    for r in range(2, N_TIT + 2):
        g = lambda c: "${}{}".format(c, r)
        put = lambda col, v: ws.cell(row=r, column=col, value=v)
        a, e = g("A"), g("E")
        put(28, ('=IF({e}<>"SIM",0,COUNTIFS(HAP_RE,{a},HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")'
                 '-COUNTIFS(HAP_RE,{a},HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO",HAP_NOSIS,"SIM"))').format(a=a, e=e))
        so_agr = g("AB")
        put(29, '=IFERROR(MATCH(ROW()-1,HAP_SEQTIT,0),"")')
        idx = g("AC")
        put(1, '=IFERROR(INDEX(HAP_RE,{i}),"")'.format(i=idx))
        put(2, '=IFERROR(INDEX(HAP_NOME,{i}),"")'.format(i=idx))
        put(3, '=IFERROR(INDEX(HAP_CONTRATO,{i}),"")'.format(i=idx))
        put(4, '=IFERROR(INDEX(HAP_INICIO,{i}),"")'.format(i=idx)).number_format = DATA
        put(5, '=IF({a}="","",IF(COUNTA(SIS_RE)=0,"NÃO",IF(COUNTIF(SIS_RE,{a})>0,"SIM","NÃO")))'.format(a=a))
        put(6, ('=IF({a}="","",IF({e}="SIM","FUNCIONÁRIO (LOCALIZADO NO SISTEMA)",'
                'IF(ISNUMBER(SEARCH("PJ",{c})),"PJ",'
                'IF(ISNUMBER(SEARCH("ACORDO",{c})),"ACORDO COM A EMPRESA",'
                'IF(ISNUMBER(SEARCH("SINDICATO",{c})),"SINDICATO",'
                '"NÃO LOCALIZADO – VERIFICAR")))))').format(a=a, e=e, c=g("C")))
        put(7, '=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO"))'.format(a=a))
        put(8, '=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO"))'.format(a=a))
        put(9, '=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"DEPENDENTE"))'.format(a=a))
        put(10, '=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"AGREGADO"))'.format(a=a))
        put(11, '=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_NOSIS,"SIM"))'.format(a=a))
        put(12, '=IF({a}="","",IF({e}<>"SIM",0,COUNTIFS(HAP_RE,{a},HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_NOSIS,"NÃO")))'.format(a=a, e=e))
        put(13, '=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO",SIS_ELEG,"NÃO"))'.format(a=a))
        put(14, ('=IF({a}="","",COUNTIFS(SIS_RE,{a},SIS_TIPO,"AGREGADO",SIS_NOHAP,"NÃO")'
                 '+COUNTIFS(SIS_RE,{a},SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO",SIS_ELEG,"SIM"))').format(a=a))
        put(15, '=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_ELEG,"SIM",HAP_DUP,"NÃO"))'.format(a=a))
        put(16, '=IF({a}="","",VLR_TITULAR)'.format(a=a)).number_format = MOEDA
        put(17, '=IF({a}="","",INDEX(COMBO_VLR,MATCH(MIN({g},5),COMBO_QTD,0)))'.format(a=a, g=g("G"))).number_format = MOEDA
        put(18, '=IF({a}="","",{h}*VLR_AGREGADO)'.format(a=a, h=g("H"))).number_format = MOEDA
        put(19, '=IF({a}="","",{p}+{q}+{r})'.format(a=a, p=g("P"), q=g("Q"), r=g("R"))).number_format = MOEDA
        put(20, '=IF({a}="","",SUMIFS(HAP_COB,HAP_RE,{a}))'.format(a=a)).number_format = MOEDA
        put(21, '=IF({a}="","",{t}-{s})'.format(a=a, t=g("T"), s=g("S"))).number_format = MOEDA
        put(22, '=IF({a}="","",{p})'.format(a=a, p=g("P"))).number_format = MOEDA
        put(23, '=IF({a}="","",{q}+{r})'.format(a=a, q=g("Q"), r=g("R"))).number_format = MOEDA
        put(26, '=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_DIVDADOS,"SIM",HAP_DUP,"NÃO"))'.format(a=a))
        put(27, ('=IF({a}="","",COUNTIFS(HAP_RE,{a},HAP_NOSIS,"VERIFICAR",HAP_DUP,"NÃO")'
                 '+COUNTIFS(SIS_RE,{a},SIS_NOHAP,"VERIFICAR"))').format(a=a))
        put(24, ('=IF({a}="","",IF({e}<>"SIM","🔵 RE NÃO LOCALIZADO NA BASE DA EMPRESA",'
                 'IF({m}+{n}>0,IF({l}+{sa}>0,"⚠️ DIVERGÊNCIA NOS DOIS SENTIDOS","🟡 INCLUSÃO NECESSÁRIA"),'
                 'IF({l}+{sa}>0,"🔴 CADASTRO SOMENTE HAPVIDA",'
                 'IF({z}>0,"⚠️ DIVERGÊNCIA DE DADOS",'
                 'IF({v}>0,"🟡 VERIFICAR CORRESPONDÊNCIA","🟢 CADASTRO OK"))))))').format(
                     a=a, e=e, m=g("M"), n=g("N"), l=g("L"), sa=so_agr, z=g("Z"), v=g("AA")))
        put(25, ('=IF({a}="","",TRIM('
                 'IF({m}+{n}>0,{m}+{n}&" a incluir ("&{m}&" dep + "&{n}&" agr). ","")&'
                 'IF({l}+{sa}>0,{l}+{sa}&" só no Hapvida. ","")&'
                 'IF({v}>0,{v}&" a verificar nome. ","")&'
                 'IF({z}>0,{z}&" com divergência de dados. ","")&'
                 'IF({o}>0,{o}&" dep acima da idade. ","")&'
                 'IF(ABS({u})>0.01,"Cobrança diverge em "&TEXT({u},"R$ #,##0.00")&". ","")&'
                 'IF(COUNTIFS(HAP_RE,{a},HAP_DUP,"SIM")>0,"Registro duplicado. ","")&'
                 'IF(COUNTIFS(HAP_RE,{a},HAP_TIPO,"TITULAR",HAP_DUP,"NÃO")>1,"Matrícula compartilhada. ","")))').format(
                     a=a, m=g("M"), n=g("N"), l=g("L"), sa=so_agr, v=g("AA"),
                     z=g("Z"), o=g("O"), u=g("U")))
        novo = 'IF(AND({a}<>"",{y}<>""),1,0)'.format(a=a, y=g("Y"))
        put(30, ("=" + novo) if r == 2 else '=$AD{p}+{n}'.format(p=r - 1, n=novo))

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:AA{}".format(N_TIT + 1)
    ws.column_dimensions.group("AB", "AD", hidden=True)
    return ws


# ==========================================================================
# ABA: CONCILIACAO
# ==========================================================================
CONC_COLS = [
    ("RE", 10), ("COLABORADOR", 34), ("CONTRATO", 18), ("PESSOA", 36), ("CONDIÇÃO", 26),
    ("GRAU/CONDIÇÃO HAPVIDA", 16), ("NASCIMENTO HAPVIDA", 13), ("IDADE HAPVIDA", 9),
    ("VALOR HAPVIDA", 12), ("ESTÁ NO HAPVIDA?", 13),
    ("GRAU/CONDIÇÃO SISTEMA", 16), ("NASCIMENTO SISTEMA", 13), ("IDADE SISTEMA", 9),
    ("ESTÁ NO SISTEMA?", 13), ("SITUAÇÃO", 38), ("OBSERVAÇÃO", 74),
    ("_q", 8), ("_r", 8), ("_t", 8), ("_ns", 8),
]

STATUS_CORES = [
    ("CADASTRO OK", F_VERDE, C_VERDE),
    ("INCLUSÃO NECESSÁRIA", F_AMAR, C_AMAR),
    ("VERIFICAR", F_AMAR, C_AMAR),
    ("DIVERGÊNCIA DE DADOS", F_LARANJ, C_LARANJ),
    ("DIVERGÊNCIA NOS DOIS", F_LARANJ, C_LARANJ),
    ("SOMENTE HAPVIDA", F_VERM, C_VERM),
    ("NÃO LOCALIZADO", F_AZUL, C_AZUL),
    ("SEM CONVÊNIO", C_CINZA2, "555555"),
    ("NÃO ELEGÍVEL", C_CINZA2, "555555"),
]


def pinta_status(ws, rng):
    """Formatacao condicional por palavra-chave (evita LEFT() sobre emoji)."""
    primeira = rng.split(":")[0].replace("$", "")
    for chave, bg, fg in STATUS_CORES:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=['ISNUMBER(SEARCH("{}",{}))'.format(chave, primeira)],
            fill=fill(bg), font=Font(name=FONTE, color=fg, bold=True, size=9), stopIfTrue=True))


def pinta_sim_nao(ws, rng):
    for chave, bg, fg in (("SIM", F_VERDE, C_VERDE), ("VERIFICAR", F_AMAR, C_AMAR),
                          ("NÃO", F_VERM, C_VERM)):
        primeira = rng.split(":")[0].replace("$", "")
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=['ISNUMBER(SEARCH("{}",{}))'.format(chave, primeira)],
            fill=fill(bg), font=Font(name=FONTE, color=fg, bold=True, size=9), stopIfTrue=True))


def aba_conciliacao(wb):
    ws = wb.create_sheet("CONCILIAÇÃO")
    header_row(ws, 1, [c[0] for c in CONC_COLS], h=44)
    for i, (_, w) in enumerate(CONC_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    ini_b = N_HAP + 2
    fim = N_HAP + 1 + N_CONC_B

    for r in range(2, fim + 1):
        g = lambda c: "${}{}".format(c, r)
        put = lambda col, v: ws.cell(row=r, column=col, value=v)
        a, q, rr, t = g("A"), g("Q"), g("R"), g("S")
        if r < ini_b:
            put(17, r - 1)
            put(18, '=IFERROR(INDEX(HAP_IDXSIS,{q}),"")'.format(q=q))
        else:
            put(18, '=IFERROR(MATCH({j},SIS_SEQ,0),"")'.format(j=r - N_HAP - 1))
        put(19, '=IF({a}="","",IFERROR(MATCH({a},TIT_RE,0),""))'.format(a=a))
        put(20, '=IF({q}="","",INDEX(HAP_NOSIS,{q}))'.format(q=q))

        put(1, ('=IFERROR(IF({q}<>"",IF(INDEX(HAP_CHDT,{q})="","",INDEX(HAP_RE,{q})),'
                'IF({r}="","",INDEX(SIS_RE,{r}))),"")').format(q=q, r=rr))
        put(2, ('=IF({a}="","",IF({t}<>"",INDEX(TIT_NOME,{t}),'
                'IF({q}<>"",INDEX(HAP_NOME,{q}),TRIM(INDEX(SIS_COLAB,{r})&""))))').format(
                    a=a, t=t, q=q, r=rr))
        put(3, '=IF({a}="","",IF({q}<>"",INDEX(HAP_CONTRATO,{q}),IF({t}<>"",INDEX(TIT_CONTRATO,{t}),"—")))'.format(
            a=a, q=q, t=t))
        put(4, '=IF({a}="","",IF({q}<>"",INDEX(HAP_NOME,{q}),INDEX(SIS_PESSOA,{r})))'.format(a=a, q=q, r=rr))
        put(5, ('=IF({a}="","",IF(IF({q}<>"",INDEX(HAP_TIPO,{q}),INDEX(SIS_TIPO,{r}))="AGREGADO","🟠 AGREGADO",'
                'IF(IF({q}<>"",INDEX(HAP_ELEG,{q}),INDEX(SIS_ELEG,{r}))="SIM","⚠️ DEPENDENTE ACIMA DA IDADE",'
                'IF({q}<>"",INDEX(HAP_TIPO,{q}),INDEX(SIS_TIPO,{r})))))').format(a=a, q=q, r=rr))
        put(6, '=IF({a}="","",IF({q}<>"",INDEX(HAP_PAREN,{q}),""))'.format(a=a, q=q))
        put(7, '=IF({a}="","",IF({q}<>"",INDEX(HAP_DTN,{q}),""))'.format(a=a, q=q)).number_format = DATA
        put(8, '=IF({a}="","",IF({q}<>"",INDEX(HAP_IDADE,{q}),""))'.format(a=a, q=q))
        put(9, '=IF({a}="","",IF({q}<>"",INDEX(HAP_COB,{q}),0))'.format(a=a, q=q)).number_format = MOEDA
        put(10, '=IF({a}="","",IF({q}<>"","🟢 SIM","🔴 NÃO"))'.format(a=a, q=q))
        put(11, '=IF({a}="","",IF({r}="","",INDEX(SIS_REL,{r})))'.format(a=a, r=rr))
        put(12, '=IF({a}="","",IF({r}="","",INDEX(SIS_DTN,{r})))'.format(a=a, r=rr)).number_format = DATA
        put(13, '=IF({a}="","",IF({r}="","",INDEX(SIS_IDADE,{r})))'.format(a=a, r=rr))
        put(14, ('=IF({a}="","",IF({q}="","🟢 SIM",IF(INDEX(HAP_NOSIS,{q})="SIM","🟢 SIM",'
                 'IF(INDEX(HAP_NOSIS,{q})="VERIFICAR","🟡 VERIFICAR","🔴 NÃO"))))').format(a=a, q=q))
        put(15, ('=IF({a}="","",IF({q}="",'
                 'IF(INDEX(SIS_TIPO,{r})="NÃO ELEGÍVEL","⚪ NÃO ELEGÍVEL PARA O CONVÊNIO",'
                 'IF(INDEX(SIS_TIPO,{r})="TITULAR","⚪ TITULAR DO SISTEMA SEM CONVÊNIO",'
                 'IF(INDEX(SIS_RENOHAP,{r})="NÃO","⚪ RE SEM CONVÊNIO NO HAPVIDA","🟡 INCLUSÃO NECESSÁRIA"))),'
                 'IF({t}="SIM",IF(INDEX(HAP_DIVDADOS,{q})="SIM","⚠️ DIVERGÊNCIA DE DADOS","🟢 CADASTRO OK"),'
                 'IF({t}="VERIFICAR","🟡 VERIFICAR CORRESPONDÊNCIA",'
                 'IF(INDEX(HAP_RENOSIS,{q})="NÃO","🔵 RE NÃO LOCALIZADO NA BASE DA EMPRESA",'
                 '"🔴 CADASTRO SOMENTE HAPVIDA")))))').format(a=a, q=q, r=rr, t=g("T")))
        put(16, ('=IF({a}="","",IF({q}="","Só no sistema da empresa"'
                 '&IF(INDEX(SIS_ELEG,{r})="SIM"," · acima da idade → entra como AGREGADO","")'
                 '&IF(INDEX(SIS_TIPO,{r})="NÃO ELEGÍVEL"," · relação não elegível ao convênio","")'
                 '&IF(INDEX(SIS_RENOHAP,{r})="NÃO"," · RE sem convênio no Hapvida",""),'
                 'IF({t}="VERIFICAR","Confere só o 1º nome com "&INDEX(SIS_PESSOA,{r})&" — confirmar",'
                 'IF({t}="NÃO",IF(INDEX(HAP_RENOSIS,{q})="NÃO","RE fora da base da empresa · contrato "&{c},'
                 '"Não localizado no cadastro do sistema"),'
                 'IF(INDEX(HAP_DIVDADOS,{q})="SIM","No sistema: "&INDEX(SIS_PESSOA,{r})&" · "&INDEX(SIS_REL,{r}),'
                 'INDEX(HAP_CORRESP,{q}))))'
                 '&IF(INDEX(HAP_DUP,{q})="SIM"," · REGISTRO DUPLICADO","")))').format(
                     a=a, q=q, r=rr, t=g("T"), c=g("C")))

    ws.freeze_panes = "E2"
    ws.auto_filter.ref = "A1:P{}".format(fim)
    pinta_status(ws, "O2:O{}".format(fim))
    pinta_sim_nao(ws, "J2:J{}".format(fim))
    pinta_sim_nao(ws, "N2:N{}".format(fim))
    ws.column_dimensions.group("Q", "T", hidden=True)
    for nome, col in [("CONC_RE", "A"), ("CONC_COLAB", "B"), ("CONC_PESSOA", "D"),
                      ("CONC_COND", "E"), ("CONC_GRAUH", "F"), ("CONC_NASCH", "G"),
                      ("CONC_NOHAP", "J"), ("CONC_GRAUS", "K"), ("CONC_NASCS", "L"),
                      ("CONC_NOSIS", "N"), ("CONC_SIT", "O"), ("CONC_OBS", "P")]:
        wb.defined_names.add(DefinedName(nome, attr_text=col_ref("CONCILIAÇÃO", col, 2, fim)))
    return ws


# ==========================================================================
# ABA: SIMULACAO DE IMPACTO
# ==========================================================================
SIM_COLS = [
    ("RE", 10), ("Colaborador", 34), ("Dependentes Atuais", 11), ("Dependentes Sistema", 11),
    ("A Incluir", 10), ("Agregados", 10), ("Valor Atual", 13), ("Novo Valor", 13),
    ("Aumento Mensal", 13), ("Aumento Anual", 14),
    ("A incluir como AGREGADO", 11), ("Novo nº de dependentes", 11),
    ("Novo nº de agregados", 11), ("COMO FICA", 76),
]


def aba_simulacao(wb):
    ws = wb.create_sheet("SIMULAÇÃO DE IMPACTO")
    titulo(ws, "A1", "SIMULAÇÃO — quanto custaria regularizar no Hapvida tudo o que já consta "
                     "no sistema da empresa", span=14, size=13)
    ws["A2"] = ("Pessoa até a idade limite entra no COMBO FAMILIAR (muda a faixa, nunca multiplica "
                "o valor). Acima da idade limite entra como AGREGADO, somando o valor cheio. "
                "Quem está marcado como VERIFICAR ou NÃO ELEGÍVEL fica de fora — não se cria "
                "custo em cima de dúvida.")
    ws["A2"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("A2:N2")
    ws.row_dimensions[2].height = 26
    header_row(ws, 3, [c[0] for c in SIM_COLS], h=44)
    for i, (_, w) in enumerate(SIM_COLS):
        ws.column_dimensions[CL(i + 1)].width = w

    for r in range(4, N_TIT + 4):
        k = r - 3
        g = lambda c: "${}{}".format(c, r)
        put = lambda col, v: ws.cell(row=r, column=col, value=v)
        a = g("A")
        put(1, '=IFERROR(IF(INDEX(TIT_RE,{k})="","",INDEX(TIT_RE,{k})),"")'.format(k=k))
        for col, nome in ((2, "TIT_NOME"), (3, "TIT_QDEP"), (4, "TIT_QDEPSIS"),
                          (5, "TIT_DEPINC"), (6, "TIT_QAGR"), (11, "TIT_AGRINC")):
            put(col, '=IF({a}="","",INDEX({n},{k}))'.format(a=a, n=nome, k=k))
        put(7, '=IF({a}="","",INDEX(TIT_TOTAL,{k}))'.format(a=a, k=k)).number_format = MOEDA
        put(12, '=IF({a}="","",{c}+{e})'.format(a=a, c=g("C"), e=g("E")))
        put(13, '=IF({a}="","",{f}+{k})'.format(a=a, f=g("F"), k=g("K")))
        put(8, ('=IF({a}="","",VLR_TITULAR+INDEX(COMBO_VLR,MATCH(MIN({l},5),COMBO_QTD,0))'
                '+{m}*VLR_AGREGADO)').format(a=a, l=g("L"), m=g("M"))).number_format = MOEDA
        put(9, '=IF({a}="","",{h}-{g})'.format(a=a, h=g("H"), g=g("G"))).number_format = MOEDA
        put(10, '=IF({a}="","",{i}*12)'.format(a=a, i=g("I"))).number_format = MOEDA
        put(14, ('=IF({a}="","",TRIM('
                 'IF({e}>0,{e}&" no Combo: faixa de "&{c}&" → "&{l}&" dependentes. ","")&'
                 'IF({k}>0,{k}&" como AGREGADO a "&TEXT(VLR_AGREGADO,"R$ #,##0.00")&" cada. ","")&'
                 'IF({i}=0,"Nada a incluir — sem impacto financeiro.","")))').format(
                     a=a, e=g("E"), c=g("C"), l=g("L"), k=g("K"), i=g("I")))

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = "A3:N{}".format(N_TIT + 3)
    for nome, col in (("SIM_AUM_MES", "I"), ("SIM_AUM_ANO", "J"), ("SIM_NOVO", "H")):
        wb.defined_names.add(DefinedName(
            nome, attr_text=col_ref("SIMULAÇÃO DE IMPACTO", col, 4, N_TIT + 3)))
    return ws


# ==========================================================================
# ABA: RESUMO POR CONTRATO
# ==========================================================================
def aba_resumo_contrato(wb, contratos):
    ws = wb.create_sheet("RESUMO POR CONTRATO")
    ws.sheet_view.showGridLines = False
    titulo(ws, "B2", "RESUMO POR CONTRATO", span=10, size=15)
    ws["B3"] = ("Os nomes de contrato (coluna B, em azul) são editáveis. Acrescente novos nas "
                "linhas reservadas — a linha OUTROS mostra tudo o que não foi listado, "
                "para nada ficar de fora.")
    ws["B3"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("B3:K3")
    ws.row_dimensions[3].height = 26

    header_row(ws, 5, ["CONTRATO", "Titulares", "Dependentes", "Agregados", "Valor Titulares",
                       "Valor Dependentes (Combo)", "Valor Agregados", "TOTAL",
                       "Custo EMPRESA", "Custo COLABORADOR"], start_col=2, h=40)
    larguras(ws, {"A": 3, "B": 30, "C": 11, "D": 12, "E": 11, "F": 15, "G": 18, "H": 15,
                  "I": 16, "J": 15, "K": 17})

    n_slots = len(contratos) + 6
    r0 = 6
    specs = [(3, "COUNTIFS(TIT_CONTRATO,{a})", INT), (4, "SUMIFS(TIT_QDEP,TIT_CONTRATO,{a})", INT),
             (5, "SUMIFS(TIT_QAGR,TIT_CONTRATO,{a})", INT), (6, "SUMIFS(TIT_VTIT,TIT_CONTRATO,{a})", MOEDA),
             (7, "SUMIFS(TIT_VCOMBO,TIT_CONTRATO,{a})", MOEDA), (8, "SUMIFS(TIT_VAGR,TIT_CONTRATO,{a})", MOEDA),
             (9, "SUMIFS(TIT_TOTAL,TIT_CONTRATO,{a})", MOEDA), (10, "SUMIFS(TIT_EMPRESA,TIT_CONTRATO,{a})", MOEDA),
             (11, "SUMIFS(TIT_COLAB,TIT_CONTRATO,{a})", MOEDA)]
    for i in range(n_slots):
        r = r0 + i
        c = ws.cell(row=r, column=2, value=contratos[i] if i < len(contratos) else None)
        c.font = f(color=C_AZUL_TXT, bold=True, size=10)
        c.fill = fill("FFFFCC")
        c.border = BORDA
        a = "$B{}".format(r)
        for col, fml, fmt in specs:
            cc = ws.cell(row=r, column=col, value='=IF({a}="","",{f})'.format(a=a, f=fml.format(a=a)))
            cc.number_format = fmt
            cc.border = BORDA
            cc.font = f(size=10)

    r_out, r_tot = r0 + n_slots, r0 + n_slots + 1
    ws.cell(row=r_out, column=2, value="OUTROS (contratos não listados acima)")
    ws.cell(row=r_tot, column=2, value="TOTAL GERAL")
    totais = [(3, 'SUMPRODUCT(--(TIT_RE<>""))', INT), (4, "SUM(TIT_QDEP)", INT),
              (5, "SUM(TIT_QAGR)", INT), (6, "SUM(TIT_VTIT)", MOEDA), (7, "SUM(TIT_VCOMBO)", MOEDA),
              (8, "SUM(TIT_VAGR)", MOEDA), (9, "SUM(TIT_TOTAL)", MOEDA),
              (10, "SUM(TIT_EMPRESA)", MOEDA), (11, "SUM(TIT_COLAB)", MOEDA)]
    for col, fml, fmt in totais:
        L = CL(col)
        co = ws.cell(row=r_out, column=col,
                     value="={t}-SUM({L}{a}:{L}{b})".format(t=fml, L=L, a=r0, b=r0 + n_slots - 1))
        co.number_format = fmt; co.border = BORDA
        co.font = f(size=10, italic=True); co.fill = fill(C_CINZA)
        ct = ws.cell(row=r_tot, column=col, value="=" + fml)
        ct.number_format = fmt; ct.border = BORDA
        ct.font = f(bold=True, size=10, color="FFFFFF"); ct.fill = fill(C_NAVY)
    for r, bg, cor in ((r_out, C_CINZA, "555555"), (r_tot, C_NAVY, "FFFFFF")):
        cb = ws.cell(row=r, column=2)
        cb.fill = fill(bg); cb.font = f(bold=True, size=10, color=cor); cb.border = BORDA
    ws.cell(row=r_tot + 2, column=2,
            value="Custo EMPRESA = valor do titular. Custo COLABORADOR = Combo Familiar + "
                  "agregados.").font = f(italic=True, size=9, color="555555")
    ws.auto_filter.ref = "B5:K{}".format(r0 + n_slots - 1)
    ws.freeze_panes = "C6"
    return ws


# ==========================================================================
# ABA: RESUMO GERENCIAL
# ==========================================================================
def bloco(ws, r0, c0, tit, itens, largura=4, bg=C_TEAL):
    titulo(ws, "{}{}".format(CL(c0), r0), tit, span=largura, size=11, bg=bg)
    r = r0 + 1
    for rot, fml, fmt in itens:
        cr = ws.cell(row=r, column=c0, value=rot)
        cr.font = f(size=10)
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + largura - 2)
        cv = ws.cell(row=r, column=c0 + largura - 1, value=fml)
        cv.number_format = fmt
        cv.font = f(bold=True, size=10)
        cv.alignment = Alignment(horizontal="right")
        _borda(ws, r, c0, c0 + largura - 1, C_CINZA if r % 2 == 0 else None)
        r += 1
    return r


def card(ws, row, col, rotulo, formula, fmt, cor=C_NAVY):
    for rr in (row, row + 1):
        ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + 2)
    c1 = ws.cell(row=row, column=col, value=rotulo)
    c1.font = f(bold=True, size=9, color="FFFFFF"); c1.fill = fill(cor)
    c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c2 = ws.cell(row=row + 1, column=col, value=formula)
    c2.number_format = fmt; c2.font = f(bold=True, size=18, color=cor); c2.fill = fill("FFFFFF")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.border = Border(left=Side("medium", color=cor), right=Side("medium", color=cor),
                       bottom=Side("medium", color=cor))
    ws.row_dimensions[row].height = 30
    ws.row_dimensions[row + 1].height = 34


def aba_resumo(wb):
    ws = wb.create_sheet("RESUMO GERENCIAL")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 42, "C": 8, "D": 8, "E": 17, "F": 4,
                  "G": 42, "H": 8, "I": 8, "J": 17, "K": 3})

    titulo(ws, "B2", "CONVÊNIO MÉDICO HAPVIDA — RESUMO GERENCIAL", span=9, size=16)
    ws["B3"] = ('="Idades calculadas na data-base "&TEXT(DATA_BASE,"dd/mm/yyyy")&'
                '"  |  Titular "&TEXT(VLR_TITULAR,"R$ #,##0.00")&" (empresa)  |  '
                'Agregado "&TEXT(VLR_AGREGADO,"R$ #,##0.00")&" (colaborador)  |  '
                'Combo Familiar pela faixa de dependentes"')
    ws["B3"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("B3:J3")

    card(ws, 5, 2, "TITULARES NO HAPVIDA", '=SUMPRODUCT(--(TIT_RE<>""))', INT, C_NAVY)
    card(ws, 5, 5, "DEPENDENTES NO HAPVIDA", '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")', INT, C_TEAL)
    card(ws, 5, 8, "AGREGADOS NO HAPVIDA", '=COUNTIFS(HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")', INT, C_LARANJ)
    card(ws, 8, 2, "CUSTO MENSAL TOTAL", "=SUM(TIT_TOTAL)", MOEDA0, C_NAVY)
    card(ws, 8, 5, "CUSTO MENSAL DA EMPRESA", "=SUM(TIT_EMPRESA)", MOEDA0, C_VERDE)
    card(ws, 8, 8, "IMPACTO MENSAL SE REGULARIZAR", "=SUM(SIM_AUM_MES)", MOEDA0, C_VERM)

    r_esq = bloco(ws, 11, 2, "COLABORADORES", [
        ("Total de titulares no Hapvida", '=SUMPRODUCT(--(TIT_RE<>""))', INT),
        ("Localizados na base da empresa", '=COUNTIF(TIT_NOSIS,"SIM")', INT),
        ("PJ (pelo contrato)", '=COUNTIF(TIT_VINCULO,"PJ")', INT),
        ("Acordo com a empresa (pelo contrato)", '=COUNTIF(TIT_VINCULO,"ACORDO COM A EMPRESA")', INT),
        ("Sindicato (pelo contrato)", '=COUNTIF(TIT_VINCULO,"SINDICATO")', INT),
        ("NÃO LOCALIZADO – VERIFICAR", '=COUNTIF(TIT_VINCULO,"NÃO LOCALIZADO*")', INT),
        ("Com alguma diferença a tratar", '=COUNTIF(TIT_SITUACAO,"*DIVERGÊNCIA*")+COUNTIF(TIT_SITUACAO,"*INCLUSÃO*")+COUNTIF(TIT_SITUACAO,"*SOMENTE HAPVIDA*")+COUNTIF(TIT_SITUACAO,"*VERIFICAR CORRESP*")', INT),
        ("Cadastro OK nas duas bases", '=COUNTIF(TIT_SITUACAO,"*CADASTRO OK*")', INT),
        ("Colaboradores no sistema SEM convênio", '=COUNTIF(CONC_SIT,"*TITULAR DO SISTEMA SEM CONVÊNIO*")', INT),
        ("Registros DUPLICADOS no arquivo do Hapvida", '=COUNTIF(HAP_DUP,"SIM")', INT),
        ("Matrículas compartilhadas por 2+ titulares", '=COUNTIFS(HAP_TIPO,"TITULAR",HAP_1TIT,"NÃO",HAP_DUP,"NÃO")', INT),
    ])
    r_dir = bloco(ws, 11, 7, "DEPENDENTES", [
        ("Total de dependentes no Hapvida", '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")', INT),
        ("Total de dependentes no sistema", '=COUNTIF(SIS_TIPO,"DEPENDENTE")', INT),
        ("Encontrados nas duas bases", "=SUM(TIT_DEPOK)", INT),
        ("Somente no Hapvida", "=SUM(TIT_DEPSOHAP)", INT),
        ("Somente no sistema da empresa", '=COUNTIFS(SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO")', INT),
        ("A incluir no Hapvida como DEPENDENTE", "=SUM(TIT_DEPINC)", INT),
        ("Acima da idade limite no Hapvida", '=COUNTIFS(HAP_ELEG,"SIM",HAP_DUP,"NÃO")', INT),
        ("Colaboradores com dependentes", '=COUNTIF(TIT_QDEP,">0")', INT),
        ("Vinculados a um titular (entram no Combo)", "=SUM(TIT_QDEP)", INT),
        ("Sem titular vinculado no arquivo", '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")-SUM(TIT_QDEP)', INT),
    ])
    r_esq = bloco(ws, r_esq + 1, 2, "AGREGADOS", [
        ("Total de agregados no Hapvida", '=COUNTIFS(HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")', INT),
        ("Total de agregados no sistema", '=COUNTIF(SIS_TIPO,"AGREGADO")', INT),
        ("A incluir no Hapvida como AGREGADO", "=SUM(TIT_AGRINC)", INT),
        ("→ por serem agregados no sistema", '=COUNTIFS(SIS_TIPO,"AGREGADO",SIS_NOHAP,"NÃO")', INT),
        ("→ por estarem acima da idade limite", '=COUNTIFS(SIS_TIPO,"DEPENDENTE",SIS_NOHAP,"NÃO",SIS_ELEG,"SIM")', INT),
        ("Colaboradores com agregados", '=COUNTIF(TIT_QAGR,">0")', INT),
        ("Custo mensal dos agregados", "=SUM(TIT_VAGR)", MOEDA),
    ])
    faixas = [("0 a 5 anos", 0, 5), ("6 a 10 anos", 6, 10), ("11 a 17 anos", 11, 17),
              ("18 a 23 anos", 18, 23), ("24 a 25 anos", 24, 25)]
    itens_f = [(rot, '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_IDADE,">={a}",HAP_IDADE,"<={b}")'.format(a=a, b=b), INT)
               for rot, a, b in faixas]
    itens_f += [("Acima de 25 anos (possíveis agregados)",
                 '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_IDADE,">"&IDADE_LIMITE)', INT),
                ("→ dos quais são cônjuge (não viram agregado)",
                 '=COUNTIFS(HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO",HAP_IDADE,">"&IDADE_LIMITE)-COUNTIFS(HAP_ELEG,"SIM",HAP_DUP,"NÃO")', INT),
                ("→ efetivamente elegíveis a agregado", '=COUNTIFS(HAP_ELEG,"SIM",HAP_DUP,"NÃO")', INT)]
    r_dir = bloco(ws, r_dir + 1, 7, "FAIXA ETÁRIA DOS DEPENDENTES (HAPVIDA)", itens_f)

    r_esq = bloco(ws, r_esq + 1, 2, "FINANCEIRO (situação atual)", [
        ("Valor total dos TITULARES", "=SUM(TIT_VTIT)", MOEDA),
        ("Valor total do COMBO FAMILIAR", "=SUM(TIT_VCOMBO)", MOEDA),
        ("Valor total dos AGREGADOS", "=SUM(TIT_VAGR)", MOEDA),
        ("VALOR TOTAL DO CONVÊNIO", "=SUM(TIT_TOTAL)", MOEDA),
        ("Custeado pela EMPRESA", "=SUM(TIT_EMPRESA)", MOEDA),
        ("Custeado pelos COLABORADORES", "=SUM(TIT_COLAB)", MOEDA),
        ("Total cobrado no arquivo do Hapvida", "=SUM(TIT_COBRADO)", MOEDA),
        ("DIFERENÇA (cobrado − calculado)", "=SUM(TIT_DIF)", MOEDA),
        ("Titulares com diferença de cobrança", '=COUNTIF(TIT_DIF,">0.01")+COUNTIF(TIT_DIF,"<-0.01")', INT),
    ])
    r_dir = bloco(ws, r_dir + 1, 7, "QUALIDADE DO CRUZAMENTO DE NOMES", [
        ("Correspondência forte (data de nascimento)", '=COUNTIF(HAP_CORRESP,"FORTE (nascimento)*")', INT),
        ("Correspondência forte (nome completo)", '=COUNTIF(HAP_CORRESP,"FORTE (nome completo)*")', INT),
        ("Provável (nome abreviado no sistema)", '=COUNTIF(HAP_CORRESP,"PROVÁVEL*")', INT),
        ("A VERIFICAR (só o 1º nome confere)", '=COUNTIF(HAP_NOSIS,"VERIFICAR")+COUNTIF(SIS_NOHAP,"VERIFICAR")', INT),
        ("Pessoas do Hapvida sem correspondência", '=COUNTIF(HAP_NOSIS,"NÃO")', INT),
        ("Pessoas do sistema sem correspondência", '=COUNTIF(SIS_NOHAP,"NÃO")', INT),
        ("Relações NÃO ELEGÍVEIS no sistema", '=COUNTIF(SIS_TIPO,"NÃO ELEGÍVEL")', INT),
    ])
    r_dir = bloco(ws, r_dir + 1, 7, "SIMULAÇÃO — SE TUDO FOR REGULARIZADO", [
        ("Colaboradores impactados", '=COUNTIF(SIM_AUM_MES,">0")', INT),
        ("Pessoas a incluir (dependentes)", "=SUM(TIT_DEPINC)", INT),
        ("Pessoas a incluir (agregados)", "=SUM(TIT_AGRINC)", INT),
        ("Custo mensal HOJE", "=SUM(TIT_TOTAL)", MOEDA),
        ("Custo mensal APÓS regularização", "=SUM(SIM_NOVO)", MOEDA),
        ("AUMENTO MENSAL", "=SUM(SIM_AUM_MES)", MOEDA),
        ("AUMENTO ANUAL", "=SUM(SIM_AUM_ANO)", MOEDA),
    ])
    r_esq = bloco(ws, max(r_esq, r_dir) + 1, 2, "CONCILIAÇÃO — PESSOAS POR SITUAÇÃO", [
        ("🟢 CADASTRO OK", '=COUNTIF(CONC_SIT,"*CADASTRO OK*")', INT),
        ("🟡 INCLUSÃO NECESSÁRIA", '=COUNTIF(CONC_SIT,"*INCLUSÃO NECESSÁRIA*")', INT),
        ("🔴 CADASTRO SOMENTE HAPVIDA", '=COUNTIF(CONC_SIT,"*SOMENTE HAPVIDA*")', INT),
        ("🟡 VERIFICAR CORRESPONDÊNCIA", '=COUNTIF(CONC_SIT,"*VERIFICAR CORRESPONDÊNCIA*")', INT),
        ("⚠️ DIVERGÊNCIA DE DADOS", '=COUNTIF(CONC_SIT,"*DIVERGÊNCIA DE DADOS*")', INT),
        ("🔵 RE NÃO LOCALIZADO NA BASE DA EMPRESA", '=COUNTIF(CONC_SIT,"*NÃO LOCALIZADO*")', INT),
        ("⚪ Sem convênio / não elegível", '=COUNTIF(CONC_SIT,"*SEM CONVÊNIO*")+COUNTIF(CONC_SIT,"*NÃO ELEGÍVEL*")', INT),
        ("🟠 AGREGADOS (condição)", '=COUNTIF(CONC_COND,"*AGREGADO*")', INT),
        ("TOTAL DE PESSOAS CONCILIADAS", '=SUMPRODUCT(--(CONC_RE<>""))', INT),
    ])
    rn = max(r_esq, r_dir) + 1
    ws.cell(row=rn, column=2, value=(
        "Como ler: a ausência de um RE na base da empresa NÃO é erro — o vínculo é classificado "
        "pelo CONTRATO do Hapvida (PJ, Acordo, Sindicato). E ninguém é dado como não cadastrado "
        "só por diferença de escrita do nome: quando só o primeiro nome confere, a pessoa entra "
        "como VERIFICAR CORRESPONDÊNCIA e fica de fora da simulação."))
    ws.cell(row=rn, column=2).font = f(italic=True, size=9, color=C_VERM)
    ws.cell(row=rn, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=rn, start_column=2, end_row=rn + 1, end_column=10)
    ws.row_dimensions[rn].height = 28
    return ws


# ==========================================================================
# ABA: ALERTAS
# ==========================================================================
def aba_alertas(wb):
    ws = wb.create_sheet("ALERTAS")
    titulo(ws, "A1", "ALERTAS — apenas os REs que exigem alguma ação", span=9, size=13, bg=C_VERM)
    ws["A2"] = "Lista viva: cada RE com algo a tratar aparece aqui automaticamente."
    ws["A2"].font = f(italic=True, size=9, color="555555")
    ws.merge_cells("A2:E2")
    ws["F2"] = ('=IF(MAX(TIT_SEQALERTA)<={n},"Exibindo os "&MAX(TIT_SEQALERTA)&" REs com alerta '
                '(capacidade desta aba: {n}).","⚠️ Há "&MAX(TIT_SEQALERTA)&" REs com alerta e esta '
                'aba exibe apenas {n}. Copie a última linha para baixo.")').format(n=N_ALERTA)
    ws["F2"].font = f(bold=True, size=9, color=C_VERM)
    ws["F2"].alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells("F2:I2")

    cols = [("RE", 10), ("COLABORADOR", 34), ("CONTRATO", 18), ("VÍNCULO", 30),
            ("SITUAÇÃO CADASTRAL", 38), ("ALERTA", 68), ("Cobrado no arquivo", 14),
            ("Total calculado", 14), ("Diferença", 13)]
    header_row(ws, 3, [c[0] for c in cols], h=34)
    for i, (_, w) in enumerate(cols):
        ws.column_dimensions[CL(i + 1)].width = w

    for i in range(N_ALERTA):
        r = 4 + i
        idx = "$J{}".format(r)
        ws.cell(row=r, column=10, value='=IFERROR(MATCH({k},TIT_SEQALERTA,0),"")'.format(k=i + 1))
        for col, nome, fmt in ((1, "TIT_RE", None), (2, "TIT_NOME", None), (3, "TIT_CONTRATO", None),
                               (4, "TIT_VINCULO", None), (5, "TIT_SITUACAO", None),
                               (6, "TIT_ALERTA", None), (7, "TIT_COBRADO", MOEDA),
                               (8, "TIT_TOTAL", MOEDA), (9, "TIT_DIF", MOEDA)):
            c = ws.cell(row=r, column=col, value='=IF({i}="","",INDEX({n},{i}))'.format(i=idx, n=nome))
            if fmt:
                c.number_format = fmt

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = "A3:I{}".format(N_ALERTA + 3)
    pinta_status(ws, "E4:E{}".format(N_ALERTA + 3))
    ws.column_dimensions.group("J", "J", hidden=True)
    return ws


# ==========================================================================
# ABA: FICHA DO COLABORADOR
# ==========================================================================
def _proxima(rng, chave, m_ant):
    """Proxima ocorrencia de `chave` em `rng` depois de `m_ant`.

    Encadeamento DESLOC+CORRESP: uma varredura por linha exibida, sem formula
    matricial. Funciona em qualquer versao do Excel, no LibreOffice e no
    Google Sheets (AGREGAR/AGGREGATE nao e universal)."""
    if m_ant is None:
        return '=IFERROR(MATCH({c},{r},0),"")'.format(c=chave, r=rng)
    return ('=IF(N({p})=0,"",IFERROR(MATCH({c},OFFSET({r},{p},0,ROWS({r})-{p},1),0)+{p},""))'
            .format(c=chave, r=rng, p=m_ant))


def aba_ficha(wb):
    ws = wb.create_sheet("FICHA DO COLABORADOR")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 34, "C": 18, "D": 12, "E": 7, "F": 17,
                  "G": 15, "H": 15, "I": 34, "J": 3, "K": 3, "L": 9})

    titulo(ws, "B2", "FICHA DO COLABORADOR — CONVÊNIO HAPVIDA", span=8, size=16)
    ws["B4"] = "🔎 INFORME O RE:"
    ws["B4"].font = f(bold=True, size=13, color=C_NAVY_D)
    inp = ws["C4"]
    inp.value = 4
    inp.font = f(bold=True, size=16, color=C_AZUL_TXT)
    inp.fill = fill("FFFF99")
    inp.alignment = Alignment(horizontal="center", vertical="center")
    md = Side("medium", color=C_NAVY)
    inp.border = Border(left=md, right=md, top=md, bottom=md)
    inp.comment = Comment("Digite aqui o RE do colaborador. A ficha inteira se atualiza sozinha.",
                          "Conciliação")
    ws.row_dimensions[4].height = 30
    ws.merge_cells("E4:I4")
    ws["E4"] = ('=IF(RE_FICHA="","← Digite o RE do colaborador nesta célula amarela",'
                'IF(FICHA_IDX<>"",INDEX(TIT_NOME,FICHA_IDX)&"   —   "&INDEX(TIT_VINCULO,FICHA_IDX),'
                'IFERROR(TRIM(INDEX(SIS_COLAB,MATCH(RE_FICHA,SIS_RE,0)))&"   —   SEM CONVÊNIO NO HAPVIDA",'
                '"⚠️ RE não encontrado em nenhuma das duas bases")))')
    ws["E4"].font = f(bold=True, size=12, color=C_TEAL)
    ws["E4"].alignment = Alignment(vertical="center", indent=1)

    ws["L3"] = '=IF($C$4="","",IFERROR(--$C$4,$C$4))'
    ws["L4"] = '=IFERROR(MATCH(RE_FICHA,TIT_RE,0),"")'
    wb.defined_names.add(DefinedName("RE_FICHA", attr_text="'FICHA DO COLABORADOR'!$L$3"))
    wb.defined_names.add(DefinedName("FICHA_IDX", attr_text="'FICHA DO COLABORADOR'!$L$4"))

    tit = lambda n, fmt=None: ('=IF(FICHA_IDX="","—",INDEX({},FICHA_IDX))'.format(n), fmt)
    sis1 = 'IFERROR(INDEX({},MATCH(RE_FICHA,SIS_RE,0)),"—")'

    titulo(ws, "B6", "TITULAR", span=8, size=11, bg=C_TEAL)
    pares = [
        ("RE", '=IF(RE_FICHA="","—",RE_FICHA)', None,
         "Nome do colaborador", '=IF(FICHA_IDX="",IFERROR(TRIM(INDEX(SIS_COLAB,MATCH(RE_FICHA,SIS_RE,0))),"—"),INDEX(TIT_NOME,FICHA_IDX))', None),
        ("Contrato", tit("TIT_CONTRATO")[0], None,
         "Vínculo / tipo de base", '=IF(FICHA_IDX="","SEM CONVÊNIO NO HAPVIDA",INDEX(TIT_VINCULO,FICHA_IDX))', None),
        ("Início no plano", tit("TIT_INICIO")[0], DATA,
         "Consta na base da empresa?", '=IF(COUNTIF(SIS_RE,RE_FICHA)>0,"🟢 SIM","🔴 NÃO")', None),
        ("Função (sistema)", '=' + sis1.format("SIS_FUNCAO"), None,
         "Estado civil (sistema)", '=' + sis1.format("SIS_ECIVIL"), None),
        ("Dependentes (Hapvida / Sistema)",
         '=IF(RE_FICHA="","—",COUNTIFS(HAP_RE,RE_FICHA,HAP_TIPO,"DEPENDENTE",HAP_DUP,"NÃO")&"  /  "&COUNTIFS(SIS_RE,RE_FICHA,SIS_TIPO,"DEPENDENTE"))', None,
         "Agregados (Hapvida / Sistema)",
         '=IF(RE_FICHA="","—",COUNTIFS(HAP_RE,RE_FICHA,HAP_TIPO,"AGREGADO",HAP_DUP,"NÃO")&"  /  "&COUNTIFS(SIS_RE,RE_FICHA,SIS_TIPO,"AGREGADO"))', None),
    ]
    for i, (l1, f1, fmt1, l2, f2, fmt2) in enumerate(pares):
        r = 7 + i
        ws.cell(row=r, column=2, value=l1).font = f(size=9, color="333333")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        c = ws.cell(row=r, column=3, value=f1)
        c.font = f(bold=True, size=10)
        if fmt1:
            c.number_format = fmt1
        ws.cell(row=r, column=6, value=l2).font = f(size=9, color="333333")
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        c = ws.cell(row=r, column=7, value=f2)
        c.font = f(bold=True, size=10)
        if fmt2:
            c.number_format = fmt2
        _borda(ws, r, 2, 9, C_CINZA if i % 2 else None)

    # ---------------- QUADRO UNICO ----------------
    titulo(ws, "B13", "DEPENDENTES E AGREGADOS — COMPOSIÇÃO FAMILIAR COMPLETA", span=8,
           size=11, bg=C_TEAL)
    cab = ["Nome", "Grau / Condição", "Nascimento", "Idade", "Tipo",
           "Está no Hapvida?", "Está no Sistema?", "Situação"]
    header_row(ws, 14, cab, start_col=2, h=30)
    for i in range(N_FICHA):
        r = 15 + i
        m = "$L{}".format(r)
        ws.cell(row=r, column=12, value=_proxima("CONC_RE", "RE_FICHA",
                                                 None if i == 0 else "$L{}".format(r - 1)))
        vals = [
            (2, '=IF({m}="","",INDEX(CONC_PESSOA,{m}))'.format(m=m), None, "left"),
            (3, '=IF({m}="","",IF(INDEX(CONC_GRAUH,{m})<>"",INDEX(CONC_GRAUH,{m}),INDEX(CONC_GRAUS,{m})))'.format(m=m), None, "center"),
            (4, '=IF({m}="","",IF(N(INDEX(CONC_NASCH,{m}))>0,INDEX(CONC_NASCH,{m}),INDEX(CONC_NASCS,{m})))'.format(m=m), DATA, "center"),
            (5, '=IF(N($D{r})=0,"",DATEDIF($D{r},DATA_BASE,"Y"))'.format(r=r), INT, "center"),
            (6, '=IF({m}="","",INDEX(CONC_COND,{m}))'.format(m=m), None, "left"),
            (7, '=IF({m}="","",INDEX(CONC_NOHAP,{m}))'.format(m=m), None, "center"),
            (8, '=IF({m}="","",INDEX(CONC_NOSIS,{m}))'.format(m=m), None, "center"),
            (9, '=IF({m}="","",INDEX(CONC_SIT,{m}))'.format(m=m), None, "left"),
        ]
        for col, fml, fmt, al in vals:
            c = ws.cell(row=r, column=col, value=fml)
            if fmt:
                c.number_format = fmt
            c.font = f(size=9, bold=(col == 7))
            c.alignment = Alignment(horizontal=al, vertical="center")
        _borda(ws, r, 2, 9, C_CINZA if i % 2 else None)
        ws.row_dimensions[r].height = 17
    fim_q = 14 + N_FICHA
    pinta_sim_nao(ws, "G15:G{}".format(fim_q))
    pinta_sim_nao(ws, "H15:H{}".format(fim_q))
    pinta_status(ws, "I15:I{}".format(fim_q))
    ws.cell(row=fim_q + 1, column=2, value=(
        "Um único quadro com toda a composição familiar: o que está no Hapvida, o que está no "
        "sistema da empresa e o que está nos dois. 🟡 VERIFICAR = o primeiro nome confere mas o "
        "restante não — conferir antes de tratar como não cadastrado."))
    ws.cell(row=fim_q + 1, column=2).font = f(italic=True, size=8, color=C_VERM)
    ws.cell(row=fim_q + 1, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=fim_q + 1, start_column=2, end_row=fim_q + 1, end_column=9)
    ws.row_dimensions[fim_q + 1].height = 24
    return ws, fim_q + 2


def ficha_custo_e_simulacao(ws, r0):
    """Blocos de custo atual, cenario e impacto da FICHA, a partir da linha r0."""
    v = lambda n: '=IF(FICHA_IDX="",0,INDEX({},FICHA_IDX))'.format(n)

    def linha(r, label, formula, fmt=MOEDA, destaque=False):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        cl = ws.cell(row=r, column=2, value=label)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        cv = ws.cell(row=r, column=7, value=formula)
        cv.number_format = fmt
        cv.alignment = Alignment(horizontal="right", indent=1)
        if destaque:
            cl.font = f(bold=True, size=11, color="FFFFFF")
            cv.font = f(bold=True, size=12, color="FFFFFF")
            _borda(ws, r, 2, 9, C_NAVY)
        else:
            cl.font = f(size=10)
            cv.font = f(bold=True, size=10)
            _borda(ws, r, 2, 9)
        cl.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 18

    rc = r0 + 1
    titulo(ws, "B{}".format(rc), "💰 CUSTO ATUAL", span=8, size=11, bg=C_TEAL)
    linha(rc + 1, "Titular — 100% custeado pela EMPRESA", v("TIT_VTIT"))
    linha(rc + 2, '="Combo Familiar ("&IF(FICHA_IDX="",0,INDEX(TIT_QDEP,FICHA_IDX))&'
                  '" dependente(s)) — responsabilidade do COLABORADOR"', v("TIT_VCOMBO"))
    linha(rc + 3, '="Agregados ("&IF(FICHA_IDX="",0,INDEX(TIT_QAGR,FICHA_IDX))&" × "&'
                  'TEXT(VLR_AGREGADO,"R$ #,##0.00")&") — responsabilidade do COLABORADOR"', v("TIT_VAGR"))
    linha(rc + 4, "TOTAL DO CONVÊNIO", v("TIT_TOTAL"), destaque=True)
    linha(rc + 5, "➤ CUSTO DA EMPRESA", v("TIT_EMPRESA"))
    linha(rc + 6, "➤ CUSTO DO COLABORADOR", v("TIT_COLAB"))
    linha(rc + 7, "Valor efetivamente cobrado no arquivo do Hapvida", v("TIT_COBRADO"))
    linha(rc + 8, "Diferença (cobrado − calculado pelo contrato)", v("TIT_DIF"))
    ws.cell(row=rc + 5, column=2).font = f(bold=True, size=10, color=C_VERDE)
    ws.cell(row=rc + 6, column=2).font = f(bold=True, size=10, color=C_LARANJ)

    rs = rc + 10
    titulo(ws, "B{}".format(rs), "📊 CENÁRIO ATUAL × CENÁRIO APÓS REGULARIZAÇÃO", span=8,
           size=11, bg=C_TEAL)
    for c1, c2, txt in ((2, 4, "Informação"), (5, 6, "Atual"), (7, 9, "Após inclusão")):
        ws.merge_cells(start_row=rs + 1, start_column=c1, end_row=rs + 1, end_column=c2)
        cel = ws.cell(row=rs + 1, column=c1, value=txt)
        cel.font = f(bold=True, color="FFFFFF", size=9)
        cel.alignment = Alignment(horizontal="center", vertical="center")
        _borda(ws, rs + 1, c1, c2, C_NAVY)
    ws.row_dimensions[rs + 1].height = 22

    b = rs + 2
    linhas = [
        ("Dependentes", '=IF(FICHA_IDX="",0,INDEX(TIT_QDEP,FICHA_IDX))',
         '=$E${r}+IF(FICHA_IDX="",0,INDEX(TIT_DEPINC,FICHA_IDX))'.format(r=b), INT),
        ("Agregados", '=IF(FICHA_IDX="",0,INDEX(TIT_QAGR,FICHA_IDX))',
         '=$E${r}+IF(FICHA_IDX="",0,INDEX(TIT_AGRINC,FICHA_IDX))'.format(r=b + 1), INT),
        ("Titular (custeado pela empresa)", '=IF(FICHA_IDX="",0,VLR_TITULAR)',
         '=IF(FICHA_IDX="",0,VLR_TITULAR)', MOEDA),
        ("Combo Familiar", '=IF(FICHA_IDX="",0,INDEX(TIT_VCOMBO,FICHA_IDX))',
         '=IF(FICHA_IDX="",0,INDEX(COMBO_VLR,MATCH(MIN($G${r},5),COMBO_QTD,0)))'.format(r=b), MOEDA),
        ("Agregados", '=IF(FICHA_IDX="",0,INDEX(TIT_VAGR,FICHA_IDX))',
         '=$G${r}*VLR_AGREGADO'.format(r=b + 1), MOEDA),
        ("TOTAL", "=SUM($E${a}:$E${z})".format(a=b + 2, z=b + 4),
         "=SUM($G${a}:$G${z})".format(a=b + 2, z=b + 4), MOEDA),
    ]
    for i, (rot, at, dep, fmt) in enumerate(linhas):
        r = b + i
        forte = rot == "TOTAL"
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        cl = ws.cell(row=r, column=2, value=rot)
        cl.font = f(bold=forte, size=10)
        cl.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        for col, val in ((5, at), (7, dep)):
            c = ws.cell(row=r, column=col, value=val)
            c.number_format = fmt
            c.font = f(bold=True, size=10)
            c.alignment = Alignment(horizontal="right", indent=1)
        _borda(ws, r, 2, 9, C_CINZA2 if forte else None)
        ws.row_dimensions[r].height = 18

    ri = b + 7
    titulo(ws, "B{}".format(ri), "💵 IMPACTO DA REGULARIZAÇÃO", span=8, size=11, bg=C_VERM)
    tot = b + 5
    for i, (rot, fml) in enumerate((
            ("AUMENTO MENSAL", "=$G${t}-$E${t}".format(t=tot)),
            ("IMPACTO ANUAL (mensal × 12)", "=($G${t}-$E${t})*12".format(t=tot)))):
        r = ri + 1 + i
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        cl = ws.cell(row=r, column=2, value=rot)
        cl.font = f(bold=True, size=11)
        cl.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        cv = ws.cell(row=r, column=7, value=fml)
        cv.number_format = MOEDA
        cv.font = f(bold=True, size=13, color=C_VERM)
        cv.alignment = Alignment(horizontal="right", indent=1)
        _borda(ws, r, 2, 9, F_VERM)
        ws.row_dimensions[r].height = 22
    rf = ri + 3
    ws.merge_cells(start_row=rf, start_column=2, end_row=rf + 1, end_column=9)
    ws.cell(row=rf, column=2, value=(
        '=IF(FICHA_IDX="","Selecione um RE que tenha convênio no Hapvida.",'
        'IF($G${t}-$E${t}=0,"Nada a incluir: todas as pessoas do sistema já constam no Hapvida '
        '(ou estão marcadas como VERIFICAR / NÃO ELEGÍVEL). Sem impacto financeiro.",'
        '"Todo o aumento é de responsabilidade do COLABORADOR — o Combo Familiar e os agregados '
        'não são custeados pela empresa. Quem tem mais de "&IDADE_LIMITE&" anos entra como '
        'AGREGADO a "&TEXT(VLR_AGREGADO,"R$ #,##0.00")&", e não muda a faixa do Combo."))').format(t=tot))
    ws.cell(row=rf, column=2).font = f(italic=True, size=9, color=C_VERM)
    ws.cell(row=rf, column=2).alignment = Alignment(wrap_text=True, vertical="top")

    ws.print_area = "A1:J{}".format(rf + 1)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.column_dimensions.group("K", "L", hidden=True)


# ==========================================================================
# ABA: INSTRUCOES
# ==========================================================================
def aba_instrucoes(wb, n_hap, n_tit, n_sis, contratos, relacoes):
    ws = wb.create_sheet("INSTRUÇÕES")
    ws.sheet_view.showGridLines = False
    larguras(ws, {"A": 3, "B": 32, "C": 104})
    titulo(ws, "B2", "CONCILIAÇÃO CADASTRAL E FINANCEIRA — CONVÊNIO HAPVIDA", span=2, size=16)
    ws["B3"] = "Ferramenta de Gestão de Benefícios / RH — leia esta página antes de usar."
    ws["B3"].font = f(italic=True, size=10, color="555555")

    secoes = [
        ("COMO ATUALIZAR AS BASES", [
            ("Aba HAPVIDA", "Cole a exportação do convênio nas colunas A a J, a partir da linha 2. "
                            "As colunas K a AC são o motor de cálculo — não digite nelas. "
                            "Capacidade: {} linhas.".format(N_HAP)),
            ("Aba SISTEMA", "Cole o export do sistema da empresa nas colunas A a I, a partir da "
                            "linha 4 (sem o cabeçalho). Layout: uma linha por pessoa; na linha do "
                            "próprio colaborador a RELACAO é TITULAR e NOME DEPENDENTE fica vazio. "
                            "As colunas J a V são calculadas. Capacidade: {} linhas.".format(N_SIS)),
            ("Se a base crescer", "Selecione a última linha de fórmulas e arraste para baixo. "
                                  "Depois amplie os intervalos em Fórmulas ▸ Gerenciador de Nomes."),
            ("Recalcular", "Tudo é fórmula viva: ao colar as bases, todas as abas se atualizam. "
                           "Para forçar, pressione F9 (ou Ctrl+Alt+F9)."),
        ]),
        ("COMO OS NOMES SÃO COMPARADOS", [
            ("O problema", "O mesmo dependente pode estar escrito de formas diferentes nas duas "
                           "bases — MARIA SILVA no sistema e MARIA APARECIDA DA SILVA no Hapvida. "
                           "Comparar só o nome completo apontaria essa pessoa como não cadastrada."),
            ("A solução", "Dentro de cada RE, a busca é feita em quatro níveis, nesta ordem: "
                          "1) data de nascimento; 2) nome completo padronizado; "
                          "3) primeiro nome + último sobrenome; 4) só o primeiro nome."),
            ("Nome padronizado", "Coluna auxiliar em maiúsculas, sem acentos e sem pontuação. "
                                 "Os nomes originais das duas bases NUNCA são alterados."),
            ("Grau de confiança", "Níveis 1 e 2 = correspondência FORTE. Nível 3 = PROVÁVEL "
                                  "(nome abreviado). Nível 4 = 🟡 VERIFICAR CORRESPONDÊNCIA."),
            ("Por que isso importa", "Quem cai no nível 4 NÃO é tratado como não cadastrado e "
                                     "NÃO entra na simulação — não se cria custo em cima de dúvida. "
                                     "Fica sinalizado para conferência humana."),
            ("Limite conhecido", "A busca devolve a primeira pessoa que casa. Se duas pessoas do "
                                 "mesmo RE tiverem o mesmo primeiro nome e nada mais as separar, "
                                 "ambas caem em VERIFICAR — que é exatamente o sinal de conferir."),
        ]),
        ("O QUE CADA ABA FAZ", [
            ("RESUMO GERENCIAL", "Painel para a diretoria: colaboradores, dependentes, agregados, "
                                 "faixa etária, financeiro, qualidade do cruzamento e impacto."),
            ("FICHA DO COLABORADOR", "Digite um RE e a ficha se monta: titular, um único quadro "
                                     "com toda a composição familiar, custo atual e simulação."),
            ("CONCILIAÇÃO", "Uma linha por PESSOA, cruzando as duas bases, com situação e "
                            "observação do que verificar. Use os filtros do cabeçalho."),
            ("TITULARES", "Uma linha por RE: quantidades, valores, situação e alertas."),
            ("SIMULAÇÃO DE IMPACTO", "Quanto custaria incluir no Hapvida tudo o que já consta no "
                                     "sistema, com aumento mensal e anual por colaborador."),
            ("RESUMO POR CONTRATO", "Titulares, dependentes, agregados e valores por contrato."),
            ("ALERTAS", "Só os REs que exigem ação."),
            ("PARÂMETROS", "Valores do contrato, idade limite e a tabela que diz como cada "
                           "RELAÇÃO é tratada. É o único lugar onde se mudam as regras."),
        ]),
        ("SITUAÇÕES", [
            ("🟢 CADASTRO OK", "Cadastrada corretamente no Hapvida e no sistema da empresa."),
            ("🟡 INCLUSÃO NECESSÁRIA", "Está no sistema da empresa e não está no Hapvida."),
            ("🔴 CADASTRO SOMENTE HAPVIDA", "Está no Hapvida e não foi localizada no sistema "
                                            "(com o RE existindo na base da empresa)."),
            ("🟠 AGREGADO", "Condição da pessoa — tratada financeiramente à parte do Combo."),
            ("⚠️ DIVERGÊNCIA DE DADOS", "Mesma pessoa nas duas bases, com nome, parentesco ou "
                                        "data de nascimento diferentes."),
            ("🟡 VERIFICAR CORRESPONDÊNCIA", "Só o primeiro nome confere. Conferir antes de tratar "
                                             "como não cadastrada."),
            ("🔵 RE NÃO LOCALIZADO", "O RE não existe na base da empresa. NÃO é erro: o vínculo sai "
                                     "do CONTRATO do Hapvida (PJ, Acordo, Sindicato) e, sem "
                                     "identificação, fica NÃO LOCALIZADO – VERIFICAR."),
            ("⚪ SEM CONVÊNIO / NÃO ELEGÍVEL", "Pessoa do sistema cujo RE não tem registro no "
                                               "Hapvida, ou cuja relação não dá direito ao convênio."),
        ]),
        ("REGRAS FINANCEIRAS", [
            ("Titular", "R$ 151,73 por titular — 100% custeado pela EMPRESA."),
            ("Combo Familiar", "Definido pela FAIXA da quantidade total de dependentes "
                               "(1→239,93; 2→446,62; 3→579,83; 4→663,57; 5 ou mais→846,86). "
                               "NUNCA é valor unitário × quantidade. Custo do COLABORADOR."),
            ("Agregados", "R$ 544,95 por agregado, somados individualmente e sempre separados do "
                          "Combo. Custo do COLABORADOR."),
            ("Regra de idade", "Dependente acima de {} anos é tratado como AGREGADO — mas só nas "
                               "relações marcadas com SIM na tabela de PARÂMETROS. Cônjuge e "
                               "companheiro(a) vêm com NÃO: são dependentes independentemente da "
                               "idade.".format(IDADE_LIMITE_DEP)),
            ("Relações não elegíveis", "EX CONJUGE vem marcado como NÃO ELEGÍVEL: fica fora dos "
                                       "cálculos e da simulação. Para incluir, troque o TIPO na "
                                       "tabela de PARÂMETROS."),
            ("Idades", "Recalculadas da data de nascimento com a data-base de PARÂMETROS. A coluna "
                       "de idade que vem nos arquivos é apenas o retrato da data da extração."),
        ]),
        ("ESTADO DESTA CÓPIA", [
            ("Base HAPVIDA", "{:,} linhas · {:,} titulares · contratos: {}".format(
                n_hap, n_tit, ", ".join(contratos)).replace(",", ".")),
            ("Base SISTEMA", "{:,} linhas · relações encontradas: {}".format(
                n_sis, ", ".join(relacoes)).replace(",", ".")),
            ("Conferência", "Todas as relações acima estão mapeadas em PARÂMETROS. Se uma nova "
                            "aparecer numa atualização futura e não estiver na tabela, ela entra "
                            "como DEPENDENTE sujeito à regra de idade — acrescente-a à tabela."),
        ]),
        ("FILTROS E TABELA DINÂMICA", [
            ("Já pronto", "Todas as abas de lista têm filtro no cabeçalho e painéis congelados."),
            ("Tabela dinâmica", "Selecione uma célula de TITULARES ou CONCILIAÇÃO e vá em "
                                "Inserir ▸ Tabela Dinâmica. Depois Análise ▸ Inserir Segmentação "
                                "de Dados para filtrar por CONTRATO ou SITUAÇÃO com um clique."),
            ("Por que fórmulas e não Power Query", "Recalculam sozinhas ao colar as bases, sem "
                                                   "depender de atualizar consulta, e abrem em "
                                                   "qualquer Excel (2016 em diante), no LibreOffice "
                                                   "e no Google Sheets."),
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
            ws.row_dimensions[r].height = max(15, 12 * (1 + len(txt) // 100))
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
    origem = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, "dados", "bases_origem.xlsx")
    saida = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        base, "Conciliacao_Hapvida_x_Sistema.xlsx")

    hap, sis = ler_origem(origem)
    if len(hap) > N_HAP:
        raise SystemExit("HAPVIDA tem {} linhas; capacidade é {}.".format(len(hap), N_HAP))
    if len(sis) > N_SIS:
        raise SystemExit("SISTEMA tem {} linhas; capacidade é {}.".format(len(sis), N_SIS))

    contratos, vistos = [], set()
    for d in hap:
        c = (d[0] or "").strip() if isinstance(d[0], str) else d[0]
        if c and c not in vistos:
            vistos.add(c); contratos.append(c)
    contratos.sort()
    relacoes, vr = [], set()
    for d in sis:
        v = (str(d[8]) or "").strip()
        if v and v not in vr:
            vr.add(v); relacoes.append(v)
    relacoes.sort()
    n_tit = sum(1 for d in hap if str(d[7]).strip().upper() == "TITULAR")
    nao_map = [x for x in relacoes if x.upper() not in {m[0] for m in MAPA_RELACAO}]

    print("HAPVIDA: {} linhas · {} titulares · contratos: {}".format(len(hap), n_tit, contratos))
    print("SISTEMA: {} linhas · relações: {}".format(len(sis), relacoes))
    if nao_map:
        print("  ATENÇÃO — relações fora do mapa de PARÂMETROS:", nao_map)

    wb = Workbook()
    wb.remove(wb.active)
    wb._named_styles["Normal"].font = Font(name=FONTE, size=10)

    for nome, fn in (("PARÂMETROS", lambda: aba_parametros(wb)),
                     ("HAPVIDA", lambda: aba_hapvida(wb, hap)),
                     ("SISTEMA", lambda: aba_sistema(wb, sis)),
                     ("TITULARES", lambda: aba_titulares(wb)),
                     ("CONCILIAÇÃO", lambda: aba_conciliacao(wb)),
                     ("SIMULAÇÃO DE IMPACTO", lambda: aba_simulacao(wb)),
                     ("RESUMO POR CONTRATO", lambda: aba_resumo_contrato(wb, contratos)),
                     ("RESUMO GERENCIAL", lambda: aba_resumo(wb)),
                     ("ALERTAS", lambda: aba_alertas(wb))):
        fn(); print("  ·", nome)
    ws_ficha, r_apos = aba_ficha(wb)
    ficha_custo_e_simulacao(ws_ficha, r_apos)
    print("  · FICHA DO COLABORADOR")
    aba_instrucoes(wb, len(hap), n_tit, len(sis), contratos, relacoes)
    print("  · INSTRUÇÕES")

    definir_nomes(wb)
    wb._sheets.sort(key=lambda s: ORDEM.index(s.title) if s.title in ORDEM else 99)
    for nome, cor in CORES_ABA.items():
        if nome in wb.sheetnames:
            wb[nome].sheet_properties.tabColor = cor
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(saida)
    compactar(saida)          # colunas repetidas viram sharedFormula
    print("Gravado: {} ({:.1f} MB)".format(saida, os.path.getsize(saida) / 1e6))
    return saida


if __name__ == "__main__":
    main()
