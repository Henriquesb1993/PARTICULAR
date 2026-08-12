#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Planilha comparativa de orcamentos fotovoltaicos - Henrique dos Santos Peite
UC Enel SP 0118487990 / 100242856108 | B1 residencial MONOFASICO
Telhado declarado Sudeste 20 graus | BYD hibrido a ser carregado em casa
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

OUT = "/home/user/PARTICULAR/solar/Comparativo_Orcamentos_Solar_Henrique.xlsx"

F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10)
GREEN = Font(name=F, size=10, color="008000")
BOLD = Font(name=F, size=10, bold=True)
BOLDW = Font(name=F, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=F, size=14, bold=True, color="1F3864")
SUB = Font(name=F, size=11, bold=True, color="1F3864")
SMALL = Font(name=F, size=8, italic=True, color="595959")
RED = Font(name=F, size=10, bold=True, color="C00000")
REDS = Font(name=F, size=9, color="C00000")

YEL = PatternFill("solid", fgColor="FFFF00")
HDR = PatternFill("solid", fgColor="1F3864")
BAND = PatternFill("solid", fgColor="D9E1F2")
OK_F = PatternFill("solid", fgColor="C6EFCE")
BAD_F = PatternFill("solid", fgColor="FFC7CE")
WARN_F = PatternFill("solid", fgColor="FFEB9C")

thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

MONEY = 'R$ #,##0.00'
MONEY0 = 'R$ #,##0'
NUM2 = '#,##0.00'
NUM0 = '#,##0'
PCT = '0.0%'
WP = 'R$ 0.000'

meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

wb = openpyxl.Workbook()


def sheet(name):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    return ws


def title(ws, txt, sub=None):
    ws["A1"] = txt
    ws["A1"].font = TITLE
    if sub:
        ws["A2"] = sub
        ws["A2"].font = SMALL


def hdr_row(ws, row, values, start=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start + i, value=v)
        c.font = BOLDW
        c.fill = HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def inp(ws, cell, val, fmt=None, note=None):
    ws[cell] = val
    ws[cell].font = BLUE
    ws[cell].fill = YEL
    ws[cell].border = BOX
    if fmt:
        ws[cell].number_format = fmt
    if note:
        ws[cell].comment = Comment(note, "Analise tecnica")


def frm(ws, cell, f, fmt=None, font=BLACK):
    ws[cell] = f
    ws[cell].font = font
    if fmt:
        ws[cell].number_format = fmt


# =====================================================================
# LEIA_ME
# =====================================================================
ws = sheet("LEIA_ME")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 125
title(ws, "COMPARATIVO DE ORCAMENTOS - ENERGIA SOLAR + CARREGAMENTO DE BYD HIBRIDO")
ws["B2"] = ("Cliente: Henrique dos Santos Peite  |  UC Enel SP 0118487990 / 100242856108  |  Classe B1 residencial, "
            "fornecimento MONOFASICO  |  Local da instalacao informado: Av. Cel. Sezefredo Fagundes, 8482 C2")
ws["B2"].font = SMALL

blocos = [
    ("SUB", "1. OS QUATRO ORCAMENTOS ANALISADOS"),
    ("T", "A) TecSolarSP - 12 modulos JA Solar 615 W = 7,38 kWp + Huawei 6 kW  ->  cotacao Belenergy WEB-006506779 = R$ 18.228,70"),
    ("T", "B) TecSolarSP - 12 modulos Astronergy 605 W = 7,26 kWp + Huawei 6 kW  ->  cotacao Belenergy WEB-006512519 = R$ 17.307,90"),
    ("T", "C) SunWash (Yuri) - 10 modulos TCL 620 W = 6,20 kWp + inversor AUXSOL 5 kW  ->  R$ 14.220,34 INSTALADO"),
    ("T", "D) SunWash (Yuri) - 10 modulos TCL 620 W = 6,20 kWp + inversor HUAWEI 5 kW  ->  R$ 15.220,34 INSTALADO"),
    ("T", "     (a SunWash tambem anunciou kits Intelbras de 'ate 500 / 700 / 1000 kWh/mes' por R$ 12.459,81 / 16.190,48 / 20.777,81"),
    ("T", "      + R$ 1.800 de acrescimo regional. Sem quantidade nem potencia de modulo informadas - descartados da comparacao.)"),
    ("T", "E) William - 10 modulos Maxeon 535 W = 5,35 kWp + Solis 4 kW  ->  R$ 18.848,91 INSTALADO (kit 11.158,91 + eng. 790 + instal. 6.900)"),
    ("T", "F) Sun Coast (Alexandre) - 8 modulos RONMA 585 W = 4,68 kWp + AUXSOL 3 kW 1 MPPT  ->  R$ 12.387,64 INSTALADO"),
    ("T", "     Proposta 40436 de 11/08/2026, valida ate 20/08/2026. CNPJ 47.864.839/0001-31. Unica com garantias declaradas."),
    ("T", "G) SunWash (Yuri) - 10 modulos de 650 W = 6,50 kWp + 4 MICROINVERSORES de 2,25 kW  ->  R$ 16.780,54 INSTALADO"),
    ("T", "     Unica com 18x SEM JUROS pelo mesmo valor a vista (R$ 932,25/mes). Unica com 1 MPPT por modulo."),
    ("T", "H) Sfero (indicacao do Renato) - 10 modulos Longi/DAH 590 W = 5,90 kWp + 2 microinversores Solis/Growatt de 2,5 kW"),
    ("T", "     -> R$ 18.733,00 INSTALADO. A UNICA com marcas Tier 1 nos DOIS equipamentos, seguro All Risk e garantia de performance em contrato."),
    ("T", "I) Sun Coast v2 (Alexandre) - 11 modulos TCL 620 W = 6,82 kWp + AUXSOL 7,5 kW 2 MPPT  ->  R$ 16.890,00 INSTALADO"),
    ("T", "     Substitui a proposta F: corrigiu a relacao DC/AC de 1,56 para 0,91 e passou de 1 para 2 MPPTs. Melhor R$/Wp de todos."),
    ("T", "J) AD BioSolar (Adriano Melo) - 8 modulos JA Solar 615 W = 4,92 kWp + HUAWEI SUN2000-5KTL-L1  ->  R$ 17.442,00 INSTALADO"),
    ("T", "     A UNICA que informou o MODELO EXATO do inversor, e a de melhores garantias: 30 anos de eficiencia no modulo e 10 anos no inversor."),
    ("T", "     Em contrapartida e a mais cara por Wp entre as turnkey e cobre so 77% do consumo com o carro."),
    ("T", "K) Bruno G4 (11) 96861-7830 - nada recebido ainda. Coluna reservada na aba COMPARATIVO."),
    ("", ""),
    ("SUB", "2. AS TRES DESCOBERTAS MAIS IMPORTANTES"),
    ("R", "(1) SEU PADRAO E MONOFASICO, E TODOS OS INVERSORES COTADOS SAO DE 220 V."),
    ("T", "     Na Enel SP o monofasico entrega 127 V. Inversor de 220 V exige DUAS fases. Ou seja: TODOS os quatro orcamentos"),
    ("T", "     dependem de mudanca do padrao de entrada para bifasico - com custo, projeto e aprovacao da Enel. Somente o"),
    ("T", "     William menciona isso (observacao 2 da proposta dele), e nenhum dos quatro colocou o custo no preco."),
    ("T", "     Consequencia extra: ao virar bifasico, o custo de disponibilidade sobe de 30 para 50 kWh/mes para sempre."),
    ("R", "(2) A GERACAO PROMETIDA PELA TECSOLARSP FOI CALCULADA COMO SE O TELHADO FOSSE VOLTADO AO NORTE."),
    ("T", "     Eles declaram telhado SUDESTE e prometem 10.554 kWh/ano para 7,38 kWp, o que da 1.430 kWh por kWp ao ano."),
    ("T", "     Esse valor e exatamente o que um telhado NORTE produziria (1.421 na aba GERACAO). No Sudeste sao ~1.264."),
    ("T", "     A SunWash escreve isso abertamente no anuncio: 'CALCULO CONSIDERANDO NORTE DO TELHADO'. Desconto de ~11%."),
    ("R", "(3) NENHUM DOS QUATRO CONSIDEROU O CARREGAMENTO DO BYD."),
    ("T", "     Todos dimensionaram pela conta antiga (media de 381 kWh/mes). Com o carro o consumo vai para a faixa de"),
    ("T", "     580 a 680 kWh/mes. A proposta do William (5,35 kWp) fica subdimensionada e o inversor Solis de 4 kW nao"),
    ("T", "     permite ampliar depois. A Sun Coast (4,68 kWp) e a menor de todas e cobre so 73% do consumo com o carro."),
    ("T", "     As de 6,20 e 7,38 kWp por coincidencia ficam adequadas."),
    ("", ""),
    ("SUB", "2c. CUIDADO COM O PAYBACK COMO CRITERIO UNICO"),
    ("T", "Sistema pequeno se paga rapido e depois te deixa pagando conta de luz para sempre. O payback ignora a energia que"),
    ("T", "voce NAO compensou. Use a linha 61 do COMPARATIVO (LUCRO LIQUIDO EM 20 ANOS) para decidir, nao a linha 57 (payback)."),
    ("T", "Exemplo real deste caso: Sun Coast e SunWash tem payback quase igual, mas a SunWash rende cerca de R$ 48 mil MAIS"),
    ("T", "em 20 anos, porque o sistema e 32% maior e cobre o carro."),
    ("", ""),
    ("SUB", "2b. DIFERENCA DE PRECO ENTRE OS FORNECEDORES - O DADO MAIS IMPORTANTE PARA NEGOCIAR"),
    ("T", "SunWash cobra R$ 2,29 a R$ 2,46 por Wp INSTALADO. William cobrou R$ 3,52 por Wp INSTALADO."),
    ("T", "A TecSolarSP paga R$ 2,47 por Wp so de MATERIAL - o preco de venda dela ainda nao foi informado."),
    ("T", "Voce agora sabe o custo de material do integrador. Use isso para negociar."),
    ("", ""),
    ("SUB", "3. LEGENDA DE CORES"),
    ("Y", "CELULA AMARELA COM NUMERO AZUL = voce pode/deve editar. Todo o resto e formula e recalcula sozinho."),
    ("T", "Texto preto = formula.   Texto verde = valor puxado de outra aba.   Fundo vermelho = alerta tecnico."),
    ("", ""),
    ("SUB", "4. O QUE JA E DADO REAL E O QUE AINDA E ESTIMATIVA"),
    ("T", "DADO REAL (extraido da conta que veio na proposta do William): historico de 12 meses, media de 381,3 kWh/mes,"),
    ("T", "fatura de jul/2026 de R$ 572,70 para 540 kWh, COSIP de R$ 21,11, tarifa efetiva de R$ 1,0215/kWh, classe B1 monofasica."),
    ("R", "ESTIMATIVA A CONFIRMAR: (a) todos os dados do BYD - modelo, km/dia, horario e potencia do carregador;"),
    ("R", "(b) potencia, quantidade de modulos e inversor dos kits SunWash - o anuncio nao informa;"),
    ("R", "(c) preco de venda instalado da TecSolarSP - os 2 PDFs sao o CUSTO DE MATERIAL deles, nao o seu preco;"),
    ("R", "(d) custo da adequacao do padrao monofasico -> bifasico; (e) TUSD Fio B da Enel SP; (f) tipo real da telha."),
    ("", ""),
    ("SUB", "5. ROTEIRO DAS ABAS"),
    ("T", "ENTRADAS ........... dados da conta (REAIS), premissas tecnicas e dados do BYD (a preencher)"),
    ("T", "CONTA .............. analise da fatura e do historico de 12 meses, media, extremos e tendencia"),
    ("T", "GERACAO ............ irradiacao mes a mes e geracao por kWp: telhado Sudeste x Norte"),
    ("T", "BYD ................ consumo do carro em kWh/mes, potencia x energia do carregador e os 3 cenarios de consumo"),
    ("T", "DIMENSIONAMENTO .... o kWp que VOCE precisa e as opcoes economica / ideal / com expansao"),
    ("T", "COMPARATIVO ........ quadro lado a lado dos 5 orcamentos. ABA PRINCIPAL"),
    ("T", "BOM_DETALHADO ...... lista de materiais item por item das 3 cotacoes Belenergy"),
    ("T", "ECONOMIA ........... economia do ano 1 pelas regras da Lei 14.300 (Fio B, disponibilidade, COSIP)"),
    ("T", "PROJECAO ........... fluxo de caixa de 20 anos e payback de cada opcao"),
    ("T", "SIMULACAO_MENSAL ... consumo x geracao x banco de creditos mes a mes"),
    ("T", "CHECKLIST .......... o que exigir por escrito antes de assinar"),
    ("T", "NOVOS_ORCAMENTOS ... formulario para o Bruno G4 e proximos vendedores"),
    ("", ""),
    ("SUB", "6. AVISO TECNICO"),
    ("T", "Toda conclusao aqui e 'aparentemente adequada pelos dados fornecidos'. Nada esta 'confirmado apos inspecao tecnica'."),
    ("T", "Estrutura do telhado, quadro eletrico, aterramento, padrao de entrada e area util exigem visita de profissional"),
    ("T", "habilitado com CREA ou CFT e ART emitida. Nao foram enviadas fotos do telhado, portanto a area disponivel e a"),
    ("T", "quantidade real de placas que cabem NAO foram avaliadas nesta analise."),
]
r = 4
for kind, txt in blocos:
    c = ws.cell(row=r, column=2, value=txt)
    if kind == "SUB":
        c.font = SUB
    elif kind == "R":
        c.font = RED
    elif kind == "Y":
        c.font = BLUE
        c.fill = YEL
    else:
        c.font = BLACK
    r += 1

# =====================================================================
# ENTRADAS
# =====================================================================
ws = sheet("ENTRADAS")
for col, w in [("A", 3), ("B", 50), ("C", 17), ("D", 12), ("E", 70)]:
    ws.column_dimensions[col].width = w
title(ws, "ENTRADAS - DADOS DE PARTIDA",
      "Edite apenas as celulas AMARELAS. Os dados da conta ja sao REAIS (extraidos da fatura). Os dados do BYD ainda sao estimativa.")


def sect(row, txt):
    ws.cell(row=row, column=2, value=txt).font = SUB
    for c in range(2, 6):
        ws.cell(row=row, column=c).fill = BAND


def line(row, label, cell, val, fmt, unit, note, formula=False, is_input=True):
    ws.cell(row=row, column=2, value=label).font = BLACK
    if formula:
        frm(ws, cell, val, fmt)
    elif is_input:
        inp(ws, cell, val, fmt)
    else:
        ws[cell] = val
        ws[cell].font = BLACK
        if fmt:
            ws[cell].number_format = fmt
    ws.cell(row=row, column=4, value=unit).font = SMALL
    cn = ws.cell(row=row, column=5, value=note)
    cn.font = REDS if note.startswith("ALERTA") or note.startswith("ESTIMATIVA") else SMALL


sect(4, "A) LOCAL E TELHADO")
line(5, "Endereco da unidade consumidora", "C5", "Av Cel Sezefredo Fagundes 8482 CS 2", None, "", "CONFIRMADO na fatura")
line(6, "Bairro / Cidade / CEP", "C6", "Jardim das Pedras - SAO PAULO/SP - 02367-075", None, "", "ALERTA: a peca da TecSolarSP diz 'Franco da Rocha'. A conta diz SAO PAULO capital. Corrigir antes da homologacao")
line(7, "Distribuidora", "C7", "Enel SP (Eletropaulo Metropolitana)", None, "", "CONFIRMADO na fatura - CNPJ 61.xxx.xxx/xxxx-93")
line(8, "Orientacao do telhado (azimute)", "C8", "Sudeste (SE)", None, "", "Declarado pela TecSolarSP. Penaliza a geracao no hemisferio sul")
line(9, "Inclinacao do telhado", "C9", 20, NUM0, "graus", "Declarado pela TecSolarSP")
line(10, "Tipo de telha", "C10", "A CONFIRMAR", None, "", "ALERTA: TecSolarSP cotou FIBROCIMENTO e William cotou telha COLONIAL. Um dos dois nao viu o telhado")
line(11, "Sombreamento", "C11", "NAO AVALIADO", None, "", "Nao foram enviadas fotos do telhado - area util e sombras nao foram verificadas")

sect(13, "B) CONTA DE ENERGIA  (TODOS OS VALORES CONFIRMADOS NA FATURA DE JUL/2026)")
line(14, "Classe / subgrupo / modalidade", "C14", "B1 Residencial Convencional", None, "", "CONFIRMADO: 'B - B1 - CONVENCIONAL - Residencial'")
line(15, "Numero de fases do padrao ATUAL", "C15", 1, NUM0, "fases", "CONFIRMADO na fatura: MONOFASICO (127 V na Enel SP)")
line(16, "Numero de fases DEPOIS do projeto", "C16", 2, NUM0, "fases", "ALERTA: todos os inversores cotados sao de 220 V e exigem 2 fases. Este e o valor usado nos calculos")
line(17, "Custo de disponibilidade apos o projeto", "C17", "=IF(C16=1,30,IF(C16=2,50,100))", NUM0, "kWh/mes",
     "Minimo faturado todo mes. Sobe de 30 para 50 kWh ao virar bifasico - custo permanente", formula=True)
line(18, "TARIFA EFETIVA TOTAL com tributos", "C18", 1.02147, MONEY, "R$/kWh", "CONFIRMADO: R$ 551,59 de energia / 540 kWh. Inclui TUSD + TE + bandeira amarela + PIS/COFINS + ICMS 18%")
line(19, "TUSD com tributos", "C19", 0.59643, MONEY, "R$/kWh", "CONFIRMADO na fatura (tarifa sem tributos: R$ 0,46394)")
line(20, "TE - Energia com tributos", "C20", 0.40085, MONEY, "R$/kWh", "CONFIRMADO na fatura (tarifa sem tributos: R$ 0,31182)")
line(21, "Adicional de bandeira (amarela em jul/26)", "C21", 0.02419, MONEY, "R$/kWh", "CONFIRMADO: R$ 13,06 / 540 kWh. Em bandeira verde este valor e zero")
line(22, "Parcela da TUSD que corresponde ao Fio B", "C22", 0.60, PCT, "", "ESTIMATIVA - a fatura nao detalha o Fio B. Pedir a planilha tarifaria da Enel/ANEEL. Faixa usual: 55% a 70% da TUSD")
line(23, "TUSD Fio B com tributos", "C23", "=C19*C22", MONEY, "R$/kWh", "Base de cobranca sobre a energia compensada (Lei 14.300)", formula=True)
line(24, "% do Fio B faturado em 2026", "C24", 0.60, PCT, "", "Lei 14.300/2022: 60% em 2026, 75% em 2027, 90% em 2028, 100% de 2029 em diante")
line(25, "CUSTO EFETIVO DO FIO B HOJE", "C25", "=C23*C24", MONEY, "R$/kWh", "E o que voce paga por cada kWh compensado com credito", formula=True)
line(26, "COSIP / iluminacao publica", "C26", 21.11, MONEY, "R$/mes", "CONFIRMADO na fatura. NAO e compensada pela geracao solar")
line(27, "Inflacao energetica anual", "C27", 0.07, PCT, "%/ano", "Premissa. A Enel aplicou +8,8% para residencial em 04/07/2026 (Res. Homologatoria 3.596/26)")
line(28, "Custo da adequacao do padrao (mono -> bifasico)", "C28", 2500.0, MONEY, "R$",
     "ESTIMATIVA - nenhum orcamento incluiu. Pedir orcamento de eletricista + taxa Enel. Aplicado a TODAS as opcoes")

sect(30, "C) PREMISSAS TECNICAS DO SISTEMA FV")
line(31, "Performance Ratio (PR)", "C31", 0.79, PCT, "", "Perdas de ~21%: temperatura, sujeira, cabos, inversor, mismatch")
line(32, "PR do cenario conservador", "C32", 0.72, PCT, "", "Usado na linha de sensibilidade do COMPARATIVO. William usou 0,70 na proposta dele")
line(33, "Degradacao no 1o ano", "C33", 0.02, PCT, "", "Tipico de modulo novo")
line(34, "Degradacao dos anos seguintes", "C34", 0.0045, PCT, "%/ano", "N-type TOPCon ~0,40%/ano. PERC ~0,55%/ano - o modulo do William e PERC")
line(35, "Manutencao / limpeza (O&M)", "C35", 300.0, MONEY, "R$/ano", "Premissa: 2 limpezas por ano")
line(36, "Provisao anual para troca do inversor", "C36", 150.0, MONEY, "R$/ano", "Inversor dura ~12 a 15 anos, o modulo dura 25+")
line(37, "Autoconsumo instantaneo", "C37", 0.30, PCT, "", "Parcela usada no mesmo instante da geracao. Sobe muito se o BYD carregar de dia")

sect(39, "D) BYD HIBRIDO  >>> DADOS NAO INFORMADOS - TUDO AQUI E ESTIMATIVA <<<")
line(40, "Modelo do veiculo", "C40", "BYD Song Pro DM-i GL", None, "", "INFORMADO pelo cliente")
line(41, "Capacidade da bateria", "C41", 12.9, NUM2, "kWh", "ESTIMATIVA para o Song Pro DM-i - CONFIRMAR no manual, pois a versao GL pode diferir")
line(42, "Consumo eletrico do veiculo", "C42", 16.0, NUM2, "kWh/100km", "ESTIMATIVA de uso real do Song Pro em modo eletrico. Faixa plausivel: 14 a 18 conforme trajeto e uso de ar-condicionado")
line(43, "Distancia percorrida por dia", "C43", 80, NUM0, "km/dia", "INFORMADO pelo cliente")
line(44, "Dias de carregamento por mes", "C44", 30, NUM0, "dias/mes", "INFORMADO pelo cliente: carrega todos os dias")
line(45, "Eficiencia de carregamento (tomada -> bateria)", "C45", 0.86, PCT, "", "Perdas do carregador AC, do BMS e termicas. Faixa tecnica de 85% a 90%")
line(46, "Potencia do carregador (kW = POTENCIA)", "C46", 3.3, NUM2, "kW", "ESTIMATIVA: carregador AC embarcado do Song Pro DM-i. CONFIRMAR - define disjuntor e cabo")
line(47, "Objetivo declarado", "C47", "RODAR TUDO NO ELETRICO", None, "", "INFORMADO pelo cliente. Exige DUAS cargas por dia - ver bloco 7 da aba BYD")
line(48, "Margem de crescimento futuro do consumo", "C48", 0.15, PCT, "", "Premissa para o Cenario 3. A conta ja mostra alta de 22% em 6 meses e 39% ano a ano")
line(49, "Reserva da bateria mantida para o modo hibrido", "C49", 0.25, PCT, "", "Os DM-i nao esvaziam a bateria: guardam uma reserva para operar como hibrido. ESTIMATIVA de 25%")

sect(51, "E) BASE DE DIMENSIONAMENTO (calculada da conta real)")
line(52, "Base residencial PURA - 10 meses antes do degrau", "C52", "=CONTA!C77*30.4", NUM0, "kWh/mes",
     "352 kWh/mes. E a base USADA nos cenarios, partindo do principio de que o degrau de maio ja e carga do carro", formula=True)
line(53, "Alternativa: consumo atual medido (com o degrau)", "C53", "=CONTA!C78*30.4", NUM0, "kWh/mes",
     "468 kWh/mes. Use esta se o degrau for OUTRA carga nova e nao o carro - ver bloco 6 da aba BYD", formula=True)
line(54, "Custo marginal por modulo adicional (instalado)", "C54", 900.0, MONEY, "R$/modulo",
     "ESTIMATIVA: modulo + estrutura + cabo + mao de obra, SEM inversor. Usado nas colunas de simulacao G12 e G14. PEDIR O VALOR REAL AO YURI")

sect(56, "F) PADRAO REAL DE CARREGAMENTO (informado pelo cliente)")
line(57, "Reposicao media por carga", "C57", 0.50, PCT, "da bateria", "INFORMADO: recarrega em media 50% da bateria por sessao")
line(58, "Cargas por dia", "C58", 1, NUM0, "cargas/dia", "INFORMADO: uma carga por dia")

sect(60, "G) COMBUSTIVEL (para comparar eletrico x gasolina)")
line(61, "Consumo em modo hibrido", "C61", 16.0, NUM2, "km/litro", "ESTIMATIVA para o Song Pro DM-i rodando como hibrido. Faixa plausivel: 15 a 20 km/l")
line(62, "Preco da gasolina", "C62", 6.20, MONEY, "R$/litro", "ESTIMATIVA. Ajuste para o preco que voce paga")

# =====================================================================
# CONTA
# =====================================================================
ws = sheet("CONTA")
for col, w in [("A", 3), ("B", 11), ("C", 15), ("D", 9), ("E", 13), ("F", 3), ("G", 42), ("H", 16), ("I", 52)]:
    ws.column_dimensions[col].width = w
title(ws, "ANALISE DA CONTA DE ENERGIA - DADOS REAIS DA FATURA",
      "Fonte: fatura Enel SP no 599307231012, ref. 07/2026, UC 0118487990 / 100242856108, NF 085726810 emitida em 30/07/2026.")

ws.cell(row=4, column=2, value="HISTORICO DE 13 MESES (impresso na fatura)").font = SUB
for c in range(2, 6):
    ws.cell(row=4, column=c).fill = BAND
hdr_row(ws, 5, ["Mes/Ano", "Consumo (kWh)", "Dias", "Media (kWh/dia)"], start=2)
hist = [("Jul/25", 389, 32), ("Ago/25", 364, 31), ("Set/25", 290, 29), ("Out/25", 375, 32),
        ("Nov/25", 337, 30), ("Dez/25", 323, 29), ("Jan/26", 372, 32), ("Fev/26", 348, 30),
        ("Mar/26", 354, 29), ("Abr/26", 394, 32), ("Mai/26", 457, 30), ("Jun/26", 422, 29),
        ("Jul/26", 540, 33)]
for i, (m, v, d) in enumerate(hist):
    r = 6 + i
    ws.cell(row=r, column=2, value=m).font = BLACK
    inp(ws, f"C{r}", v, NUM0)
    inp(ws, f"D{r}", d, NUM0)
    frm(ws, f"E{r}", f"=IFERROR(C{r}/D{r},0)", NUM2)
    ws.cell(row=r, column=2).border = BOX
    ws.cell(row=r, column=5).border = BOX
ws["C18"].fill = BAD_F
ws["C8"].fill = OK_F
ws.cell(row=19, column=2, value="Linha 6 = Jul/25 (mes mais antigo). Linha 18 = Jul/26 (mes atual). Verde = menor consumo. Vermelho = maior consumo.").font = SMALL

ws.cell(row=21, column=2, value="ESTATISTICAS DOS 12 MESES MAIS RECENTES (Ago/25 a Jul/26)").font = SUB
for c in range(2, 6):
    ws.cell(row=21, column=c).fill = BAND
stats = [
    (22, "TOTAL ANUAL", "=SUM(C7:C18)", NUM0, "kWh/ano"),
    (23, "CONSUMO MEDIO MENSAL  =  soma / 12", "=C22/12", NUM2, "kWh/mes"),
    (24, "Total de dias faturados", "=SUM(D7:D18)", NUM0, "dias"),
    (25, "Consumo medio diario", "=IFERROR(C22/C24,0)", NUM2, "kWh/dia"),
    (26, "MAIOR CONSUMO (Jul/26)", "=MAX(C7:C18)", NUM0, "kWh"),
    (27, "MENOR CONSUMO (Set/25)", "=MIN(C7:C18)", NUM0, "kWh"),
    (28, "Amplitude maior / menor", "=IFERROR(C26/C27-1,0)", PCT, ""),
    (29, "Consumo do mes atual (Jul/26)", "=C18", NUM0, "kWh"),
    (30, "Mes atual x media dos 12 meses", "=IFERROR(C29/C23-1,0)", PCT, ""),
    (31, "Media dos 6 meses mais antigos (Ago/25 a Jan/26)", "=AVERAGE(C7:C12)", NUM2, "kWh/mes"),
    (32, "Media dos 6 meses mais recentes (Fev/26 a Jul/26)", "=AVERAGE(C13:C18)", NUM2, "kWh/mes"),
    (33, "TENDENCIA - recentes x antigos", "=IFERROR(C32/C31-1,0)", PCT, "semestre contra semestre"),
    (34, "MESMO MES, ANO ANTERIOR: Jul/25 x Jul/26", "=IFERROR(C18/C6-1,0)", PCT, "compara so julho com julho"),
]
for r, lab, f, fmt, unit in stats:
    ws.cell(row=r, column=2, value=lab).font = BOLD if lab.isupper() else BLACK
    frm(ws, f"C{r}", f, fmt, BOLD if lab.isupper() else BLACK)
    ws.cell(row=r, column=4, value=unit).font = SMALL
ws["C23"].fill = OK_F
ws["C33"].fill = BAD_F
ws["C34"].fill = BAD_F

ws.cell(row=4, column=7, value="DECOMPOSICAO DA FATURA DE JUL/2026 (540 kWh)").font = SUB
for c in range(7, 10):
    ws.cell(row=4, column=c).fill = BAND
itens = [
    (5, "USO DO SISTEMA DE DISTRIBUICAO (TUSD)", 322.07, MONEY, "540 kWh x R$ 0,59643 com tributos (tarifa base R$ 0,46394)"),
    (6, "ENERGIA (TE)", 216.46, MONEY, "540 kWh x R$ 0,40085 com tributos (tarifa base R$ 0,31182)"),
    (7, "ADICIONAL DE BANDEIRA AMARELA", 13.06, MONEY, "Bandeira aplicada no mes: AMARELA. Em bandeira verde este item e zero"),
    (8, "Subtotal de faturamento (energia)", 551.59, MONEY, "Base de calculo do ICMS"),
    (9, "COSIP - iluminacao publica municipal", 21.11, MONEY, "Cobrada pela Prefeitura. A geracao solar NAO reduz este item"),
    (10, "TOTAL DA FATURA", 572.70, MONEY, "Vencimento 14/08/2026"),
]
for r, lab, v, fmt, note in itens:
    ws.cell(row=r, column=7, value=lab).font = BOLD if r in (8, 10) else BLACK
    inp(ws, f"H{r}", v, fmt)
    if r in (8, 10):
        ws[f"H{r}"].fill = WARN_F
    ws.cell(row=r, column=9, value=note).font = SMALL
ws.cell(row=11, column=7, value="TARIFA EFETIVA  =  551,59 / 540").font = BOLD
frm(ws, "H11", "=IFERROR(H8/540,0)", MONEY, BOLD)
ws["H11"].fill = OK_F
ws.cell(row=11, column=9, value="R$ 1,0215/kWh. E a tarifa usada em toda a planilha.").font = SMALL

ws.cell(row=13, column=7, value="TRIBUTOS EMBUTIDOS NA FATURA").font = SUB
trib = [
    (14, "PIS/PASEP - base R$ 452,31 a 0,92%", 4.14, "Incide antes do ICMS"),
    (15, "COFINS - base R$ 452,31 a 4,23%", 19.12, "Incide antes do ICMS"),
    (16, "ICMS - base R$ 551,59 a 18%", 99.28, "Maior tributo da conta"),
    (17, "TOTAL DE TRIBUTOS", 122.54, "21,4% do valor total da fatura"),
]
for r, lab, v, note in trib:
    ws.cell(row=r, column=7, value=lab).font = BOLD if r == 17 else BLACK
    inp(ws, f"H{r}", v, MONEY)
    if r == 17:
        ws[f"H{r}"].fill = WARN_F
    ws.cell(row=r, column=9, value=note).font = SMALL

ws.cell(row=19, column=7, value="DADOS DA UNIDADE CONSUMIDORA").font = SUB
uc = [
    (20, "Distribuidora", "Enel SP - Eletropaulo Metropolitana"),
    (21, "Unidade consumidora", "0118487990 / 100242856108"),
    (22, "Numero do cliente", "24186384"),
    (23, "Classe / subgrupo / modalidade", "B - B1 - Convencional - Residencial"),
    (24, "TIPO DE LIGACAO", "MONOFASICO"),
    (25, "Bandeira tarifaria do mes", "AMARELA"),
    (26, "Periodo de leitura", "26/06/2026 a 29/07/2026 - 33 dias"),
    (27, "Proxima leitura prevista", "27/08/2026"),
    (28, "Medidor", "UQM9HDN25C01281527"),
    (29, "Leitura anterior / atual", "3.458 -> 3.998 (constante 1,0)"),
    (30, "Reajuste tarifario vigente", "+8,8% residencial em 04/07/2026 (Res. Hom. ANEEL 3.596/26)"),
    (31, "Custo de disponibilidade hoje (monofasico)", "30 kWh/mes"),
    (32, "Custo de disponibilidade se virar bifasico", "50 kWh/mes"),
    (33, "DEC / FEC do conjunto (ano)", "DEC 7,46 h  /  FEC 4,25 interrupcoes"),
]
for r, lab, v in uc:
    ws.cell(row=r, column=7, value=lab).font = BLACK
    c = ws.cell(row=r, column=8, value=v)
    c.font = RED if r in (24, 25) else BLACK
    c.alignment = Alignment(horizontal="left")

ws.cell(row=36, column=2, value="TABELA EM ORDEM DE CALENDARIO (usada na aba SIMULACAO_MENSAL)").font = SUB
for c in range(2, 6):
    ws.cell(row=36, column=c).fill = BAND
hdr_row(ws, 37, ["Mes", "Consumo (kWh)", "Referencia"], start=2)
cal_src = [("Jan", "C12", "Jan/26"), ("Fev", "C13", "Fev/26"), ("Mar", "C14", "Mar/26"),
           ("Abr", "C15", "Abr/26"), ("Mai", "C16", "Mai/26"), ("Jun", "C17", "Jun/26"),
           ("Jul", "C18", "Jul/26"), ("Ago", "C7", "Ago/25"), ("Set", "C8", "Set/25"),
           ("Out", "C9", "Out/25"), ("Nov", "C10", "Nov/25"), ("Dez", "C11", "Dez/25")]
for i, (m, ref, lab) in enumerate(cal_src):
    r = 38 + i
    ws.cell(row=r, column=2, value=m).font = BLACK
    frm(ws, f"C{r}", f"={ref}", NUM0, GREEN)
    ws.cell(row=r, column=4, value=lab).font = SMALL
    for cc in range(2, 5):
        ws.cell(row=r, column=cc).border = BOX
frm(ws, "C50", "=SUM(C38:C49)", NUM0, BOLD)
ws.cell(row=50, column=2, value="TOTAL").font = BOLD
ws["C50"].fill = BAND

ws.cell(row=52, column=2, value="GASTO ATUAL COM ENERGIA").font = SUB
for c in range(2, 6):
    ws.cell(row=52, column=c).fill = BAND
gasto = [
    (53, "Gasto anual atual (sem BYD, sem solar)", "=C22*H11+H9*12", MONEY),
    (54, "Gasto medio mensal atual", "=C53/12", MONEY),
    (55, "Gasto do mes de maior consumo (Jul/26)", "=H10", MONEY),
]
for r, lab, f, fmt in gasto:
    ws.cell(row=r, column=2, value=lab).font = BOLD
    frm(ws, f"C{r}", f, fmt, BOLD)
    ws[f"C{r}"].fill = WARN_F

obs = [
    "CONCLUSOES DA ANALISE DA CONTA:",
    "1. Media de 381,3 kWh/mes nos 12 meses mais recentes (4.576 kWh no ano). Este e o CENARIO 1 e NAO deve ser a base do projeto.",
    "2. TENDENCIA DE ALTA CONFIRMADA POR DOIS CAMINHOS DIFERENTES:",
    "   - semestre contra semestre: os 6 meses mais recentes estao 22,0% acima dos 6 mais antigos;",
    "   - julho contra julho: Jul/25 foram 389 kWh e Jul/26 foram 540 kWh, ou seja +38,8% no MESMO mes do ano anterior.",
    "   O segundo numero elimina o efeito de estacao: o consumo esta crescendo de verdade, nao e so inverno.",
    "3. PERGUNTA IMPORTANTE: o BYD ja esta sendo carregado em casa? A curva Abr-Mai-Jun-Jul (394 / 457 / 422 / 540) e o salto",
    "   de 39% ano a ano sao compativeis com a entrada de uma carga nova. Se o carro ja carrega em casa, o consumo futuro sera",
    "   ainda maior do que o Cenario 2 desta planilha, e o dimensionamento precisa subir.",
    "4. LIGACAO MONOFASICA CONFIRMADA NA FATURA. E o achado mais critico: os quatro orcamentos usam inversor de 220 V, que",
    "   exige duas fases. Nenhum incluiu o custo da troca de padrao. Ao virar bifasico, o minimo faturado sobe de 30 para 50 kWh/mes.",
    "5. Bandeira AMARELA no mes, somando R$ 13,06. Em bandeira verde a tarifa efetiva cai para cerca de R$ 0,997/kWh.",
    "6. A Enel reajustou a tarifa residencial em +8,8% em 04/07/2026. A premissa de 7%/ano de inflacao energetica e conservadora.",
    "7. Tributos representam R$ 122,54, ou 21,4% da fatura. A energia solar reduz a base de calculo, e por isso reduz tributos tambem.",
    "8. O endereco na fatura e SAO PAULO/SP, Jardim das Pedras, CEP 02367-075 - e nao Franco da Rocha, como diz a peca da TecSolarSP.",
    "   Isso nao muda quase nada na irradiacao, mas muda a agencia de homologacao e o municipio da COSIP. Corrigir na proposta.",
]
r = 57
for t in obs:
    c = ws.cell(row=r, column=2, value=t)
    c.font = SUB if t.endswith(":") else (RED if t.startswith(("2.", "3.", "4.")) else BLACK)
    r += 1

ws.cell(row=76, column=2, value="ANALISE FORENSE: JA EXISTE CARGA NOVA NA CONTA? (o medidor nao separa o carro, mas o padrao diario denuncia)").font = SUB
for c in range(2, 6):
    ws.cell(row=76, column=c).fill = BAND
fo = [
    (77, "Media diaria dos 10 meses estaveis (Jul/25 a Abr/26)", "=AVERAGE(E6:E15)", NUM2, "kWh/dia", "Periodo em que o consumo diario ficou entre 10,0 e 12,3"),
    (78, "Media diaria dos 3 ultimos meses (Mai a Jul/26)", "=AVERAGE(E16:E18)", NUM2, "kWh/dia", "Periodo apos o degrau"),
    (79, "DEGRAU DETECTADO", "=C78-C77", NUM2, "kWh/dia", "Carga nova que apareceu por volta de maio de 2026"),
    (80, "DEGRAU EM BASE MENSAL", "=C79*30.4", NUM0, "kWh/mes", "E este o valor que provavelmente e o carro - ou uma carga nova qualquer"),
    (81, "Aumento percentual", "=IFERROR(C78/C77-1,0)", PCT, "", ""),
    (82, "CONTROLE DE SAZONALIDADE - Jul/25", "=E6", NUM2, "kWh/dia", "Julho de 2025 tambem foi inverno, com chuveiro eletrico ligado"),
    (83, "CONTROLE DE SAZONALIDADE - Jul/26", "=E18", NUM2, "kWh/dia", "Mesmo mes, mesma estacao, um ano depois"),
    (84, "Diferenca julho contra julho", "=C83-C82", NUM2, "kWh/dia", "ESTE e o numero mais confiavel: compara estacao com estacao"),
    (85, "Diferenca julho contra julho em base mensal", "=C84*30.4", NUM0, "kWh/mes", "Prova que o aumento NAO e sazonal"),
    (86, "Energia da rede por km rodado", "=ENTRADAS!C42/100/ENTRADAS!C45", NUM2, "kWh/km", "Consumo do veiculo dividido pela eficiencia de carregamento"),
    (87, "SE O DEGRAU FOR O BYD, ISSO EQUIVALE A", "=IFERROR(C80/C86,0)", NUM0, "km/mes", "Quilometragem eletrica compativel com o degrau medido"),
    (88, "  o mesmo em km por dia", "=IFERROR(C87/30.4,0)", NUM0, "km/dia", ""),
    (89, "Pelo controle julho x julho, equivale a", "=IFERROR(C85/C86,0)", NUM0, "km/mes", "Segunda estimativa, pelo caminho independente"),
]
for r, lab, f, fmt, unit, note in fo:
    ws.cell(row=r, column=2, value=lab).font = BOLD if lab.isupper() else BLACK
    frm(ws, f"C{r}", f, fmt, BOLD if lab.isupper() else BLACK)
    if lab.isupper():
        ws[f"C{r}"].fill = WARN_F
    ws.cell(row=r, column=4, value=unit).font = SMALL
    ws.cell(row=r, column=5, value=note).font = SMALL
ws["C80"].fill = BAD_F
ws["C87"].fill = OK_F
for i, t in enumerate([
    "COMO ESTA ANALISE FUNCIONA E PARA QUE SERVE:",
    "O medidor da Enel registra apenas o TOTAL da casa - nao existe forma de ler na conta quanto foi para o carro.",
    "Mas o consumo DIARIO (coluna E) revela um degrau claro: ficou entre 10,0 e 12,3 kWh/dia por dez meses seguidos,",
    "e saltou para 14,5 a 16,4 kWh/dia a partir de maio de 2026. Normalizar por dia e essencial porque os ciclos de",
    "faturamento variam de 29 a 33 dias - comparar kWh do mes direto induz a erro.",
    "",
    "O teste decisivo e julho contra julho: mesma estacao, mesmo uso de chuveiro eletrico, um ano de diferenca.",
    "Deu +4,2 kWh/dia. Isso praticamente elimina a hipotese de que o aumento seja apenas inverno.",
    "",
    "RESSALVAS HONESTAS: (a) nao ha dados de maio e junho de 2025 para comparar esses dois meses ano a ano;",
    "(b) o degrau pode ser outra carga nova - ar-condicionado, freezer, mais gente em casa - e nao o carro;",
    "(c) o mes de Nov/25 foi faturado como MED (media estimada), nao leitura real, o que adiciona um pouco de ruido;",
    "(d) a conversao para quilometragem usa o consumo estimado de 17 kWh/100 km, que ainda nao foi confirmado.",
    "",
    "POR ISSO A PERGUNTA IMPORTA TANTO: se o carro JA esta carregando desde maio, o consumo futuro e proximo dos",
    "468 kWh/mes de hoje e o sistema pode ser MENOR. Se o degrau for outra coisa e o carro ainda vai entrar, o",
    "consumo sobe para a faixa de 590 a 765 kWh/mes. A diferenca entre os dois casos e de 2 a 3 placas.",
]):
    c = ws.cell(row=91 + i, column=2, value=t)
    c.font = SUB if t.endswith(":") else (RED if t.startswith(("RESSALVAS", "POR ISSO")) else SMALL)

# =====================================================================
# GERACAO
# =====================================================================
ws = sheet("GERACAO")
for col, w in [("A", 3), ("B", 8), ("C", 8), ("D", 13), ("E", 12), ("F", 13), ("G", 12), ("H", 13), ("I", 14), ("J", 14), ("K", 13)]:
    ws.column_dimensions[col].width = w
title(ws, "IRRADIACAO E GERACAO POR kWp - REGIAO METROPOLITANA NORTE DE SAO PAULO",
      "HSP estimada para a regiao. CONFIRMAR no CRESESB ou PVGIS com o CEP exato. Fatores de transposicao calculados para latitude -23,3 graus.")
hdr_row(ws, 4, ["Mes", "Dias", "HSP horizontal (kWh/m2.dia)", "Fator SE 20 graus", "HSP plano SE",
                "Fator Norte 20 graus", "HSP plano Norte", "Geracao/kWp SE (kWh/mes)",
                "Geracao/kWp Norte (kWh/mes)", "Perda por ser SE"], start=2)
ws.row_dimensions[4].height = 42
hsp = [5.30, 5.40, 4.90, 4.40, 3.70, 3.50, 3.65, 4.55, 4.45, 4.90, 5.35, 5.45]
dias = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
f_se = [0.97, 0.96, 0.95, 0.94, 0.92, 0.91, 0.91, 0.93, 0.95, 0.96, 0.97, 0.97]
f_n = [1.02, 1.04, 1.07, 1.11, 1.14, 1.16, 1.15, 1.11, 1.06, 1.02, 1.00, 1.00]
for i in range(12):
    r = 5 + i
    ws.cell(row=r, column=2, value=meses[i]).font = BLACK
    ws.cell(row=r, column=3, value=dias[i]).font = BLACK
    inp(ws, f"D{r}", hsp[i], NUM2)
    inp(ws, f"E{r}", f_se[i], NUM2)
    frm(ws, f"F{r}", f"=D{r}*E{r}", NUM2)
    inp(ws, f"G{r}", f_n[i], NUM2)
    frm(ws, f"H{r}", f"=D{r}*G{r}", NUM2)
    frm(ws, f"I{r}", f"=F{r}*C{r}*ENTRADAS!$C$31", NUM0)
    frm(ws, f"J{r}", f"=H{r}*C{r}*ENTRADAS!$C$31", NUM0)
    frm(ws, f"K{r}", f"=IFERROR(I{r}/J{r}-1,0)", PCT)
    for cc in range(2, 12):
        ws.cell(row=r, column=cc).border = BOX
frm(ws, "C17", "=SUM(C5:C16)", NUM0, BOLD)
frm(ws, "D17", "=AVERAGE(D5:D16)", NUM2, BOLD)
frm(ws, "F17", "=AVERAGE(F5:F16)", NUM2, BOLD)
frm(ws, "H17", "=AVERAGE(H5:H16)", NUM2, BOLD)
frm(ws, "I17", "=SUM(I5:I16)", NUM0, BOLD)
frm(ws, "J17", "=SUM(J5:J16)", NUM0, BOLD)
frm(ws, "K17", "=IFERROR(I17/J17-1,0)", PCT, RED)
ws.cell(row=17, column=2, value="ANO").font = BOLD
for cc in range(2, 12):
    ws.cell(row=17, column=cc).fill = BAND
    ws.cell(row=17, column=cc).border = BOX
frm(ws, "I18", "=I17/12", NUM0, BOLD)
ws.cell(row=18, column=2, value="MEDIA MES").font = BOLD
frm(ws, "J18", "=J17/12", NUM0, BOLD)

notas = [
    "A COLUNA I E A CHAVE DE TODA A PLANILHA: quantos kWh cada 1 kWp gera por mes no SEU telhado (Sudeste, 20 graus).",
    "A coluna J mostra o que geraria num telhado Norte. A coluna K e a perda por estar voltado ao Sudeste.",
    "",
    "COMPARACAO DIRETA COM O QUE OS VENDEDORES PROMETERAM (celula I17 x J17):",
    "  TecSolarSP: 10.554 kWh/ano / 7,38 kWp = 1.430 kWh por kWp. Praticamente identico ao valor NORTE (J17).",
    "     Eles declaram o telhado como SUDESTE na propria peca publicitaria. A promessa nao fecha com a orientacao declarada.",
    "  SunWash: o anuncio diz textualmente 'CALCULO CONSIDERANDO NORTE DO TELHADO' e vale para Ribeirao Preto ate 80 km.",
    "     Ribeirao Preto tem irradiacao maior que Sao Paulo. Sao dois desvios somados a favor do vendedor.",
    "  William: usou HSP fixa de 4,5 h/dia e 30% de perdas (PR 0,70), chegando a 1.150 kWh por kWp. E o mais CONSERVADOR",
    "     dos tres - abaixo do valor calculado aqui. A proposta dele erra por seguranca, nao por exagero.",
    "",
    "CONSEQUENCIA PRATICA DO TELHADO SUDESTE: a geracao se concentra na MANHA e cai a tarde.",
    "  Isso favorece carregar o BYD entre 8h e 14h. Desfavorece consumo concentrado no fim da tarde e na noite.",
    "",
    "SENSIBILIDADE: com PR de 0,72 (telhado mais sujo, mais quente ou com sombra parcial) a geracao cai cerca de 9%.",
    "  A aba COMPARATIVO mostra as duas linhas - realista e conservadora - para voce ver o pior caso.",
]
r = 20
for t in notas:
    c = ws.cell(row=r, column=2, value=t)
    c.font = SUB if t.endswith(":") else SMALL
    r += 1

# =====================================================================
# BYD
# =====================================================================
ws = sheet("BYD")
for col, w in [("A", 3), ("B", 50), ("C", 16), ("D", 13), ("E", 13), ("F", 15), ("G", 17), ("H", 56)]:
    ws.column_dimensions[col].width = w
title(ws, "CONSUMO DO BYD HIBRIDO, CARREGADOR E CENARIOS DE CONSUMO",
      "TODOS os dados do veiculo sao ESTIMATIVA - voce ainda nao informou modelo, km/dia nem horario. kW = potencia. kWh = energia.")


def brow(r, label, f, fmt, unit, note, big=False):
    ws.cell(row=r, column=2, value=label).font = BOLD if big else BLACK
    fnt = BOLD if big else (GREEN if f.startswith("=ENTRADAS") else BLACK)
    frm(ws, f"C{r}", f, fmt, fnt)
    if big:
        ws[f"C{r}"].fill = WARN_F
        ws[f"C{r}"].border = BOX
    ws.cell(row=r, column=4, value=unit).font = SMALL
    ws.cell(row=r, column=8, value=note).font = SMALL


ws.cell(row=4, column=2, value="1) ENERGIA QUE O CARRO PRECISA (na bateria)").font = SUB
for c in range(2, 9):
    ws.cell(row=4, column=c).fill = BAND
brow(5, "Distancia por dia", "=ENTRADAS!C43", NUM0, "km/dia", "ESTIMATIVA - o dado que mais muda o resultado")
brow(6, "Consumo eletrico do veiculo", "=ENTRADAS!C42", NUM2, "kWh/100km", "ESTIMATIVA de uso real em modo eletrico")
brow(7, "Energia na bateria por dia = MENOR entre o que voce roda e o que voce recarrega",
     "=MIN(C5*C6/100,ENTRADAS!C41*ENTRADAS!C57*ENTRADAS!C58)", NUM2, "kWh/dia",
     "Roda 80 km = demanda 12,80 kWh/dia. Uma carga de 50% repoe 6,45 kWh/dia. Vale o MENOR: o carro so usa da bateria o que foi reposto")
brow(8, "Dias de carregamento por mes", "=ENTRADAS!C44", NUM0, "dias/mes", "ESTIMATIVA")
brow(9, "ENERGIA NECESSARIA PARA O VEICULO POR MES", "=C7*C8", NUM0, "kWh/mes", "Energia UTIL na bateria, ANTES das perdas de carregamento", big=True)

ws.cell(row=11, column=2, value="2) ENERGIA QUE SAI DA REDE (o que o medidor registra)").font = SUB
for c in range(2, 9):
    ws.cell(row=11, column=c).fill = BAND
brow(12, "Eficiencia de carregamento (tomada -> bateria)", "=ENTRADAS!C45", PCT, "", "86% adotado: 14% se perde no carregador AC, no BMS e em calor")
brow(13, "ENERGIA EFETIVAMENTE CONSUMIDA DA REDE PARA CARREGAR", "=C9/C12", NUM0, "kWh/mes", "ESTE e o numero que entra na conta de luz e no dimensionamento", big=True)
brow(14, "Perda de carregamento (energia paga que nao roda)", "=C13-C9", NUM0, "kWh/mes", "Desperdicio inevitavel do processo de carga")
brow(15, "Custo do carro na conta de luz SEM solar", "=C13*ENTRADAS!C18", MONEY, "R$/mes", "Tarifa real de R$ 1,0215/kWh")
brow(16, "Custo anual do carro SEM solar", "=C15*12", MONEY, "R$/ano", "")
brow(17, "Custo do carro na conta COM solar (so o Fio B)", "=C13*ENTRADAS!C25", MONEY, "R$/mes", "Compensado por credito: paga-se apenas 60% da TUSD Fio B em 2026")
brow(18, "ECONOMIA MENSAL SO NO CARREGAMENTO DO BYD", "=C15-C17", MONEY, "R$/mes", "Este e o item de maior retorno do projeto", big=True)
brow(19, "Economia anual so no carregamento do BYD", "=C18*12", MONEY, "R$/ano", "")
brow(20, "Custo por km em modo eletrico SEM solar", "=IFERROR(C15/(C5*C8),0)", MONEY, "R$/km", "Comparar com o custo por km na gasolina")
brow(21, "Custo por km em modo eletrico COM solar", "=IFERROR(C17/(C5*C8),0)", MONEY, "R$/km", "Rodar praticamente de graca")

ws.cell(row=23, column=2, value="3) POTENCIA x ENERGIA - CARREGADOR (nao confundir kW com kWh)").font = SUB
for c in range(2, 9):
    ws.cell(row=23, column=c).fill = BAND
brow(24, "Potencia do carregador", "=ENTRADAS!C46", NUM2, "kW", "POTENCIA: define corrente, bitola do cabo e disjuntor")
brow(25, "Tempo de carga por dia  =  energia do dia / potencia", "=IFERROR(C13/C8/C24,0)", NUM2, "horas/dia", "ENERGIA (kWh) dividida por POTENCIA (kW) = TEMPO (h)")
brow(26, "Tempo para carga completa 0-100%", "=IFERROR(ENTRADAS!C41/ENTRADAS!C45/C24,0)", NUM2, "horas", "Bateria inteira partindo de vazia")
brow(27, "Corrente do carregador em 220 V", "=IFERROR(C24*1000/220,0)", NUM2, "A", "Corrente nominal de operacao")
brow(28, "Disjuntor do circuito (125% da corrente - NBR 5410)", "=IFERROR(C27*1.25,0)", NUM2, "A", "Carga continua: circuito dimensionado a 125%")
brow(29, "Energia entregue em 1 hora de carregador ligado", "=C24*1", NUM2, "kWh", "EXEMPLO DA LOGICA: 7,4 kW x 3 h = 22,2 kWh")

ws.cell(row=31, column=2, value="4) CENARIOS DE CONSUMO").font = SUB
for c in range(2, 9):
    ws.cell(row=31, column=c).fill = BAND
hdr_row(ws, 32, ["Cenario", "Residencia (kWh/mes)", "BYD (kWh/mes)", "Margem (kWh/mes)",
                 "CONSUMO TOTAL (kWh/mes)", "Consumo anual (kWh)"], start=2)
cen = [(33, "1 - So a residencia (base pura, sem o carro)", "=ENTRADAS!C52", "=0", "=0"),
       (34, "2 - RESIDENCIA + BYD tudo no eletrico", "=ENTRADAS!C52", "=C13", "=0"),
       (35, "3 - RESIDENCIA + BYD + CRESCIMENTO", "=ENTRADAS!C52", "=C13", "=(C35+D35)*ENTRADAS!C48")]
for r, lab, res, byd, marg in cen:
    ws.cell(row=r, column=2, value=lab).font = BOLD if r == 35 else BLACK
    frm(ws, f"C{r}", res, NUM0, GREEN)
    frm(ws, f"D{r}", byd, NUM0)
    frm(ws, f"E{r}", marg, NUM0)
    frm(ws, f"F{r}", f"=C{r}+D{r}+E{r}", NUM0, BOLD)
    frm(ws, f"G{r}", f"=F{r}*12", NUM0)
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BOX
    ws[f"F{r}"].fill = WARN_F
ws["F35"].fill = OK_F
frm(ws, "C37", '="O BYD aumenta o consumo da residencia em "&TEXT(IFERROR(D34/C34,0),"0.0%")&" - de "&TEXT(C34,"0")&" para "&TEXT(F34,"0")&" kWh/mes."', None, RED)
frm(ws, "C38", '="O Cenario 3 ("&TEXT(F35,"0")&" kWh/mes) e o que deve orientar o dimensionamento, NAO a media de "&TEXT(C33,"0")&" kWh/mes da conta."', None, RED)
ws.cell(row=40, column=2, value="5) DIA x NOITE - ONDE CARREGAR COMPENSA MAIS").font = SUB
for c in range(2, 9):
    ws.cell(row=40, column=c).fill = BAND
hdr_row(ws, 41, ["Estrategia", "Como funciona", "Custo por kWh", "Custo mensal do carro", "Vantagem", "Desvantagem"], start=2)
estr = [
    (42, "Carregar de DIA (8h as 14h)", "Autoconsumo instantaneo: a energia vai do painel direto para o carro",
     "=0", "=0", "Nao passa pelo medidor: nao paga Fio B, nao paga nada",
     "Precisa que o carro esteja em casa de manha. Telhado SE gera mais de manha - combina bem"),
    (43, "Carregar a NOITE", "A geracao do dia vira credito e a noite o credito e usado",
     "=ENTRADAS!C25", "=C13*ENTRADAS!C25", "Comodidade: chega, pluga e dorme",
     "Paga o Fio B sobre tudo o que foi compensado. Em 2026 sao 60% da TUSD, em 2029 sera 100%"),
    (44, "Carregar SEM energia solar", "Tarifa cheia da Enel", "=ENTRADAS!C18", "=C13*ENTRADAS!C18",
     "Nenhuma", "Custo maximo"),
]
for r, e1, e2, e3, e4, e5, e6 in estr:
    ws.cell(row=r, column=2, value=e1).font = BOLD
    ws.cell(row=r, column=3, value=e2).font = BLACK
    frm(ws, f"D{r}", e3, MONEY)
    frm(ws, f"E{r}", e4, MONEY)
    ws.cell(row=r, column=6, value=e5).font = BLACK
    ws.cell(row=r, column=7, value=e6).font = BLACK
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BOX
        ws.cell(row=r, column=cc).alignment = Alignment(wrap_text=True, vertical="top")
ws["E42"].fill = OK_F
ws["E44"].fill = BAD_F
ws.cell(row=45, column=2, value="Ganho anual de carregar de DIA em vez de a NOITE:").font = BOLD
frm(ws, "C46", "=(E43-E42)*12", MONEY, BOLD)
ws["C46"].fill = OK_F
ws.cell(row=46, column=4, value="R$/ano. Parece pouco hoje porque o Fio B esta em 60%. Em 2029 sera 100% e esse ganho quase dobra.").font = SMALL
ws.cell(row=47, column=2, value="RECOMENDACAO: programar o carregamento do veiculo para iniciar entre 9h e 10h nos dias em que o carro ficar em casa.").font = BLACK
ws.cell(row=48, column=2, value="O app do BYD permite agendar horario de carga. Vale a pena usar, e o ganho cresce todo ano com a regra do Fio B.").font = BLACK

ws.cell(row=51, column=2, value="6) QUAL VAI SER O MEU CONSUMO? OS DOIS CAMINHOS POSSIVEIS").font = SUB
for c in range(2, 9):
    ws.cell(row=51, column=c).fill = BAND
ref = [
    (52, "Base residencial PURA (10 meses estaveis, antes do degrau)", "=CONTA!C77*30.4", NUM0, "kWh/mes"),
    (53, "Consumo ATUAL medido (ultimos 3 meses, ja com o degrau)", "=CONTA!C78*30.4", NUM0, "kWh/mes"),
    (54, "Degrau detectado na conta", "=CONTA!C80", NUM0, "kWh/mes"),
    (55, "Energia da rede por km rodado", "=CONTA!C86", NUM2, "kWh/km"),
]
for r, lab, f, fmt, unit in ref:
    ws.cell(row=r, column=2, value=lab).font = BLACK
    frm(ws, f"C{r}", f, fmt, GREEN)
    ws.cell(row=r, column=4, value=unit).font = SMALL

hdr_row(ws, 57, ["Quilometragem do BYD", "km/dia aprox.", "kWh/mes da rede so do carro",
                 "CASO A: o degrau JA e o carro -> total", "kWp necessario (+10%)",
                 "CASO B: o carro ainda vai entrar -> total", "kWp necessario (+10%)"], start=2)
ws.row_dimensions[57].height = 46
for i, km in enumerate([400, 600, 800, 1000, 1200, 1500]):
    r = 58 + i
    inp(ws, f"B{r}", km, NUM0)
    frm(ws, f"C{r}", f"=B{r}/30.4", NUM0)
    frm(ws, f"D{r}", f"=B{r}*$C$55", NUM0)
    frm(ws, f"E{r}", f"=$C$52+D{r}", NUM0)
    frm(ws, f"F{r}", f"=E{r}*12/GERACAO!$I$17*1.1", NUM2)
    frm(ws, f"G{r}", f"=$C$53+D{r}", NUM0)
    frm(ws, f"H{r}", f"=G{r}*12/GERACAO!$I$17*1.1", NUM2)
    for cc in range(2, 9):
        ws.cell(row=r, column=cc).border = BOX
    ws[f"F{r}"].fill = OK_F
    ws[f"H{r}"].fill = WARN_F
for i, t in enumerate([
    "CASO A (coluna E): o degrau de maio/26 JA e o carregamento do carro. Entao a residencia pura consome 352 kWh/mes",
    "   e o total futuro e 352 + o carro na quilometragem que voce realmente rodar. NAO some o degrau duas vezes.",
    "CASO B (coluna G): o degrau foi outra carga (ar-condicionado, mais gente em casa) e o carro ainda vai entrar.",
    "   Entao a base e o consumo atual de 468 kWh/mes e o carro soma por cima.",
    "",
    "A linha de 600 km/mes e a que corresponde ao degrau medido na conta. Se for esse o seu caso e o carro ja carrega,",
    "voce precisa de cerca de 4,9 kWp - e nao dos 6,2 a 6,8 kWp que os orcamentos estao oferecendo.",
    "Isso vale de 2 a 3 placas de diferenca, ou algo entre R$ 3 mil e R$ 5 mil.",
    "",
    "COMO RESOLVER SEM ADIVINHAR: o app do BYD mostra a quilometragem rodada e o consumo em modo eletrico.",
    "Olhe o total de km do mes e quanto rodou em eletrico. Com esse numero a coluna certa desta tabela aparece sozinha.",
]):
    c = ws.cell(row=65 + i, column=2, value=t)
    c.font = RED if t.startswith(("A linha de 600", "COMO RESOLVER")) else BLACK

ws.cell(row=78, column=2, value="7) LIMITE FISICO DA BATERIA - DA PARA RODAR 80 km/dia SO NO ELETRICO?").font = SUB
for c in range(2, 9):
    ws.cell(row=78, column=c).fill = BAND
lim = [
    (79, "Bateria nominal do Song Pro DM-i", "=ENTRADAS!C41", NUM2, "kWh", "ESTIMATIVA a confirmar no manual"),
    (80, "Reserva mantida para o modo hibrido", "=ENTRADAS!C49", PCT, "", "O DM-i nao esvazia a bateria - guarda reserva para rodar como hibrido"),
    (81, "Energia UTIL por carga (na bateria)", "=C79*(1-C80)", NUM2, "kWh", "E o que sobra para rodar em modo eletrico"),
    (82, "Energia da REDE por carga completa", "=IFERROR(C81/ENTRADAS!C45,0)", NUM2, "kWh", "Inclui as perdas de carregamento"),
    (83, "KM COBERTOS POR UMA CARGA", "=IFERROR(C81/(ENTRADAS!C42/100),0)", NUM0, "km", "Autonomia eletrica real por carga"),
    (84, "km/dia que voce roda", "=ENTRADAS!C43", NUM0, "km/dia", "Informado: 80 km/dia"),
    (85, "CARGAS NECESSARIAS POR DIA", "=IFERROR(C84/C83,0)", NUM2, "cargas/dia", "Acima de 1 significa que uma carga em casa NAO cobre o dia"),
    (86, "km/dia que ficam na GASOLINA com uma carga so", "=MAX(0,C84-C83)", NUM0, "km/dia", "Energia que o sistema solar NAO substitui"),
    (87, "Tempo total de carga por dia no carregador de 3,3 kW", "=IFERROR(C84*ENTRADAS!C42/100/ENTRADAS!C45/ENTRADAS!C46,0)", NUM2, "horas/dia", "Somando as duas sessoes. Nao e o gargalo - o gargalo e o carro estar plugado duas vezes"),
]
for r, lab, f, fmt, unit, note in lim:
    ws.cell(row=r, column=2, value=lab).font = BOLD if lab.isupper() else BLACK
    frm(ws, f"C{r}", f, fmt, BOLD if lab.isupper() else BLACK)
    if lab.isupper():
        ws[f"C{r}"].fill = WARN_F
    ws.cell(row=r, column=4, value=unit).font = SMALL
    ws.cell(row=r, column=8, value=note).font = SMALL
frm(ws, "C88", '=IF(C85>1,"NAO COBRE COM UMA CARGA - sao necessarias "&TEXT(C85,"0.0")&" cargas por dia para rodar tudo no eletrico","Uma carga por dia cobre a quilometragem")', None, RED)
ws["C85"].fill = BAD_F
ws["C83"].fill = OK_F

ws.cell(row=90, column=2, value="8) ESTRATEGIAS DE CARREGAMENTO E O SISTEMA QUE CADA UMA EXIGE").font = SUB
for c in range(2, 9):
    ws.cell(row=90, column=c).fill = BAND
hdr_row(ws, 91, ["Estrategia de carregamento", "kWh/dia da rede", "kWh/mes da rede",
                 "Consumo TOTAL (base 352 + carro)", "kWp necessario (+10%)", "Modulos de 650 W"], start=2)
ws.row_dimensions[91].height = 44
estr = [
    (92, "1 - Uma carga por dia em casa (limite da bateria)", "=C82", "cobre ~60 dos 80 km; o resto na gasolina"),
    (93, "2 - DUAS cargas por dia: 80 km TODO ELETRICO", "=ENTRADAS!C43*ENTRADAS!C42/100/ENTRADAS!C45", "E O SEU OBJETIVO. Exige o carro plugado duas vezes ao dia"),
    (94, "3 - Carga parcial de hoje (o degrau da conta)", "=IFERROR(CONTA!C80/ENTRADAS!C44,0)", "E o que a conta mostra acontecendo hoje: so ~19 km/dia de carga"),
    (95, "4 - SEU PADRAO REAL: 1 carga/dia repondo 50%, 30 dias", "=C13/ENTRADAS!C44", "E O CENARIO ADOTADO NA PLANILHA. Cobre ~40 dos 80 km/dia; o resto vai na gasolina"),
]
for r, lab, f, note in estr:
    ws.cell(row=r, column=2, value=lab).font = BOLD if r == 93 else BLACK
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
    frm(ws, f"C{r}", f, NUM2)
    frm(ws, f"D{r}", f"=C{r}*ENTRADAS!C44", NUM0)
    frm(ws, f"E{r}", f"=ENTRADAS!C52+D{r}", NUM0)
    frm(ws, f"F{r}", f"=IFERROR(E{r}*12/GERACAO!$I$17*1.1,0)", NUM2)
    frm(ws, f"G{r}", f"=IFERROR(ROUNDUP(F{r}*1000/650,0),0)", NUM0)
    ws.cell(row=r, column=8, value=note).font = REDS if r == 93 else SMALL
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BOX
for cl in ["E", "F", "G"]:
    ws[f"{cl}93"].fill = WARN_F
    ws[f"{cl}95"].fill = OK_F

for i, t in enumerate([
    "O ACHADO MAIS IMPORTANTE DESTE BLOCO:",
    "O Song Pro DM-i tem bateria de 12,9 kWh. Descontando a reserva que o carro guarda para operar como hibrido,",
    "sobram cerca de 9,7 kWh uteis por carga, o que rende aproximadamente 60 km de autonomia eletrica.",
    "Voce roda 80 km/dia. Portanto UMA carga em casa NAO cobre o seu dia - faltam uns 20 km, que vao na gasolina.",
    "",
    "Para rodar TUDO no eletrico voce precisa de DUAS sessoes de carga por dia. Isso significa uma destas rotinas:",
    "  (a) carregar de madrugada e novamente quando voltar para casa no meio do dia; ou",
    "  (b) carregar em casa e uma segunda vez no trabalho ou em eletroposto.",
    "O tempo nao e o problema: 4,5 h somadas no carregador de 3,3 kW. O problema e o carro estar plugado duas vezes.",
    "SE A SUA ROTINA E SAIR E SO VOLTAR NO FIM DO DIA, os 80 km todo eletrico NAO sao possiveis carregando so em casa,",
    "por mais placas que voce instale. Nesse caso o cenario realista e a Estrategia 1, com ~20 km/dia na gasolina.",
    "",
    "CONSEQUENCIA PARA O DIMENSIONAMENTO: a Estrategia 2 exige de 12 a 13 modulos de 650 W (cerca de 7,7 a 8,5 kWp).",
    "TODOS os dez orcamentos da planilha estao subdimensionados para esse objetivo - o maior deles tem 7,38 kWp.",
    "",
    "OPORTUNIDADE: como voce vai carregar duas vezes ao dia, uma das sessoes pode ficar entre 9h e 15h, em cima da",
    "geracao solar. Isso e autoconsumo instantaneo - nao passa pelo medidor e nao paga Fio B. A premissa de 30% de",
    "autoconsumo da aba ENTRADAS fica conservadora: com carga diurna do carro ela pode chegar a 45% ou 50%,",
    "e a economia sobe. Vale carregar de dia sempre que o carro estiver em casa.",
]):
    c = ws.cell(row=96 + i, column=2, value=t)
    c.font = SUB if t.endswith(":") else (RED if t.startswith(("SE A SUA ROTINA", "CONSEQUENCIA", "OPORTUNIDADE")) else BLACK)

ws.cell(row=120, column=2, value="9) ELETRICO x GASOLINA - VALE A PENA CARREGAR DUAS VEZES POR DIA?").font = SUB
for c in range(2, 9):
    ws.cell(row=120, column=c).fill = BAND
gas = [
    (121, "Km rodados por mes", "=C5*C8", NUM0, "km/mes", "80 km/dia x 30 dias"),
    (122, "Km rodados em modo ELETRICO por mes", "=IFERROR(C9/(C6/100),0)", NUM0, "km/mes", "Limitado pela energia reposta nas cargas"),
    (123, "Km rodados na GASOLINA por mes", "=MAX(0,C121-C122)", NUM0, "km/mes", "O que a bateria nao cobriu"),
    (124, "Percentual do trajeto em modo eletrico", "=IFERROR(C122/C121,0)", PCT, "", "Com uma carga de 50% por dia fica em torno da metade"),
    (125, "Litros de gasolina por mes", "=IFERROR(C123/ENTRADAS!C61,0)", NUM2, "litros", "ESTIMATIVA: 16 km/l em modo hibrido"),
    (126, "CUSTO MENSAL DE GASOLINA", "=C125*ENTRADAS!C62", MONEY, "R$/mes", "ESTIMATIVA. Este custo o sistema solar NAO reduz"),
    (128, "Energia da rede se carregasse 2x/dia (tudo eletrico)", "=C121*C6/100/ENTRADAS!C45", NUM0, "kWh/mes", "Cobre os 2.400 km no eletrico"),
    (129, "Energia ADICIONAL necessaria (2x/dia menos 1x/dia)", "=C128-C13", NUM0, "kWh/mes", "Quanto mais a casa consumiria"),
    (130, "Custo dessa energia adicional COM energia solar", "=C129*ENTRADAS!C25", MONEY, "R$/mes", "Compensada por credito: paga-se apenas o Fio B"),
    (131, "GANHO LIQUIDO DE CARREGAR 2x/DIA", "=C126-C130", MONEY, "R$/mes", "Gasolina economizada menos a energia extra"),
    (132, "Ganho anual", "=C131*12", MONEY, "R$/ano", ""),
    (134, "kWp adicionais para suportar o carregamento 2x/dia", "=IFERROR(C129*12/GERACAO!$I$17,0)", NUM2, "kWp", "Potencia extra que o sistema precisaria ter"),
    (135, "Modulos de 650 W adicionais", "=IFERROR(ROUNDUP(C134*1000/650,0),0)", NUM0, "modulos", ""),
    (136, "Custo desses modulos adicionais", "=C135*ENTRADAS!C54", MONEY, "R$", "Ao custo marginal estimado de R$ 900 por modulo"),
    (137, "PAYBACK DO INCREMENTO, EM MESES", "=IFERROR(C136/C131,0)", NUM2, "meses", "Quanto tempo os modulos extras levam para se pagar SO na gasolina economizada"),
]
for r, lab, f, fmt, unit, note in gas:
    ws.cell(row=r, column=2, value=lab).font = BOLD if lab.isupper() else BLACK
    frm(ws, f"C{r}", f, fmt, BOLD if lab.isupper() else BLACK)
    if lab.isupper():
        ws[f"C{r}"].fill = WARN_F
    ws.cell(row=r, column=4, value=unit).font = SMALL
    ws.cell(row=r, column=8, value=note).font = SMALL
ws["C131"].fill = OK_F
ws["C137"].fill = OK_F
for i, t in enumerate([
    "O QUE ESTE BLOCO MOSTRA:",
    "Carregando uma vez por dia e repondo 50% da bateria, cerca de metade dos seus 2.400 km/mes roda no eletrico",
    "e a outra metade na gasolina. O sistema solar nao toca nessa metade - combustivel nao entra na compensacao de energia.",
    "",
    "Se voce conseguir carregar DUAS vezes por dia, a energia extra da rede custa muito pouco quando ha energia solar",
    "(paga-se so o Fio B), enquanto a gasolina economizada e um valor grande. O bloco calcula os dois lados e diz",
    "em quantos MESES os modulos adicionais se pagam apenas com a gasolina que voce deixaria de comprar.",
    "",
    "ATENCAO AS ESTIMATIVAS DESTE BLOCO: o consumo de 16 km/l em modo hibrido e o preco de R$ 6,20/litro sao premissas",
    "minhas, editaveis em ENTRADAS C61 e C62. O resultado e sensivel a elas - confira com o consumo real do seu carro.",
    "E lembre: carregar 2x/dia exige o carro plugado duas vezes, o que depende da sua rotina, nao do sistema solar.",
]):
    c = ws.cell(row=139 + i, column=2, value=t)
    c.font = SUB if t.endswith(":") else (RED if t.startswith("ATENCAO") else BLACK)

# =====================================================================
# DIMENSIONAMENTO
# =====================================================================
ws = sheet("DIMENSIONAMENTO")
for col, w in [("A", 3), ("B", 46), ("C", 16), ("D", 16), ("E", 16), ("F", 60)]:
    ws.column_dimensions[col].width = w
title(ws, "DIMENSIONAMENTO CORRETO - QUANTO VOCE REALMENTE PRECISA",
      "Calculado a partir do consumo projetado, da irradiacao local, das perdas do sistema e da margem tecnica. Nao por regra de bolo.")
ws.cell(row=4, column=2, value="A) MEMORIA DE CALCULO").font = SUB
for c in range(2, 7):
    ws.cell(row=4, column=c).fill = BAND
mem = [
    (5, "Geracao anual por kWp no seu telhado SE", "=GERACAO!I17", NUM0, "kWh por kWp por ano"),
    (6, "Geracao mensal por kWp no seu telhado SE", "=C5/12", NUM2, "kWh por kWp por mes"),
    (7, "Formula usada", '="kWp necessario = consumo anual / "&TEXT(C5,"0")&" kWh por kWp"', None, ""),
]
for r, lab, f, fmt, unit in mem:
    ws.cell(row=r, column=2, value=lab).font = BLACK
    frm(ws, f"C{r}", f, fmt, GREEN if f.startswith("=GERACAO") else BLACK)
    ws.cell(row=r, column=4, value=unit).font = SMALL

ws.cell(row=9, column=2, value="B) kWp NECESSARIO EM CADA CENARIO").font = SUB
for c in range(2, 7):
    ws.cell(row=9, column=c).fill = BAND
hdr_row(ws, 10, ["Cenario", "Consumo (kWh/mes)", "Consumo (kWh/ano)", "kWp necessario", "Modulos de 615 W"], start=2)
for i, (lab, ref) in enumerate([("1 - Atual (so residencia)", "BYD!F33"),
                                ("2 - Residencia + BYD", "BYD!F34"),
                                ("3 - Residencia + BYD + crescimento", "BYD!F35")]):
    r = 11 + i
    ws.cell(row=r, column=2, value=lab).font = BOLD if i == 2 else BLACK
    frm(ws, f"C{r}", f"={ref}", NUM0, GREEN)
    frm(ws, f"D{r}", f"=C{r}*12", NUM0)
    frm(ws, f"E{r}", f"=IFERROR(D{r}/$C$5,0)", NUM2, BOLD if i == 2 else BLACK)
    frm(ws, f"F{r}", f"=IFERROR(ROUNDUP(E{r}*1000/615,0),0)", NUM0)
    for cc in range(2, 7):
        ws.cell(row=r, column=cc).border = BOX
ws["E13"].fill = OK_F

ws.cell(row=15, column=2, value="C) AS TRES OPCOES DE PROJETO (com modulos de 615 W)").font = SUB
for c in range(2, 7):
    ws.cell(row=15, column=c).fill = BAND
hdr_row(ws, 16, ["", "OPCAO A - ECONOMICO", "OPCAO B - IDEAL", "OPCAO C - COM EXPANSAO"], start=2)
opc = [
    (17, "Meta de atendimento", '="100% do Cenario 2"', '="105% do Cenario 3"', '="Cenario 3 + folga para dobrar a km do carro"'),
    (18, "kWp alvo", "=E12", "=E13*1.05", "=E13*1.25"),
    (19, "Quantidade de modulos de 615 W", "=ROUNDUP(C18*1000/615,0)", "=ROUNDUP(D18*1000/615,0)", "=ROUNDUP(E18*1000/615,0)"),
    (20, "Potencia instalada real (kWp)", "=C19*615/1000", "=D19*615/1000", "=E19*615/1000"),
    (21, "Inversor recomendado (kW)", "=5", "=6", "=7.5"),
    (22, "Relacao DC/AC", "=IFERROR(C20/C21,0)", "=IFERROR(D20/D21,0)", "=IFERROR(E20/E21,0)"),
    (23, "Geracao anual estimada (kWh)", "=C20*$C$5", "=D20*$C$5", "=E20*$C$5"),
    (24, "Geracao media mensal (kWh)", "=C23/12", "=D23/12", "=E23/12"),
    (25, "% do Cenario 2 atendido", "=IFERROR(C23/$D$12,0)", "=IFERROR(D23/$D$12,0)", "=IFERROR(E23/$D$12,0)"),
    (26, "% do Cenario 3 atendido", "=IFERROR(C23/$D$13,0)", "=IFERROR(D23/$D$13,0)", "=IFERROR(E23/$D$13,0)"),
    (27, "Excedente anual desperdicado (kWh)", "=MAX(0,C23-$D$13)", "=MAX(0,D23-$D$13)", "=MAX(0,E23-$D$13)"),
]
fmts = {18: NUM2, 19: NUM0, 20: NUM2, 21: NUM2, 22: '0.00"x"', 23: NUM0, 24: NUM0, 25: PCT, 26: PCT, 27: NUM0}
for r, lab, a, b, cc_ in opc:
    ws.cell(row=r, column=2, value=lab).font = BLACK
    for j, f in enumerate([a, b, cc_]):
        cell = f"{get_column_letter(3+j)}{r}"
        frm(ws, cell, f, fmts.get(r))
        ws[cell].border = BOX
        ws[cell].alignment = Alignment(horizontal="center", wrap_text=True)
for cl in ["C", "D", "E"]:
    ws[f"{cl}26"].fill = WARN_F
    ws[f"{cl}20"].fill = BAND
ws["D26"].fill = OK_F

rec = [
    "D) RECOMENDACAO DE DIMENSIONAMENTO",
    "A OPCAO B e a recomendada. Motivos tecnicos:",
    "1. Ela atende o Cenario 3, que e o unico realista: residencia + carro + margem. A Opcao A cobre o carro de hoje mas",
    "   nao sobra nada para o crescimento que a propria conta ja mostra (+22% em 6 meses).",
    "2. Ela nao joga dinheiro fora. Passar do Cenario 3 gera excedente que vira credito e expira em 60 meses - a linha 27",
    "   mostra esse desperdicio. Cada kWh gerado acima do consumo tem retorno praticamente zero.",
    "3. O inversor de 6 kW da Opcao B ainda aceita cerca de 8,1 kWp (limite DC/AC de 1,35), ou seja, permite acrescentar",
    "   modulos depois sem trocar o inversor e sem refazer a homologacao. Isso e a expansao futura barata.",
    "4. A Opcao C so se justifica se voce souber HOJE que a quilometragem do carro vai crescer muito, ou se pretende",
    "   instalar ar-condicionado, piscina aquecida ou um segundo veiculo eletrico.",
    "",
    "ATENCAO: os numeros acima dependem de km/dia do BYD, que ainda e ESTIMATIVA. Se voce rodar 20 km/dia em vez de 40,",
    "o kWp necessario cai cerca de 1,3 kWp (2 modulos). Se rodar 60 km/dia, sobe cerca de 1,3 kWp. Informe a km real.",
]
r = 30
for t in rec:
    c = ws.cell(row=r, column=2, value=t)
    c.font = SUB if t.startswith("D)") else (RED if t.startswith("ATENCAO") else BLACK)
    r += 1

# =====================================================================
# COMPARATIVO
# =====================================================================
OPT = [
    dict(L="A) TecSolarSP\nJA Solar 615W\nHuawei 6 kW", emp="TecSolarSP", nome="nao informado (razao social: Edijane Gomes dos Santos)", fone="(11) 97365-9192",
         cnpj="29.775.600/0001-03", doc="Cotacao Belenergy WEB-006506779", tipo="MATERIAL apenas",
         qty=12, wp=615, mod="JA Solar bifacial N-type", inv="Huawei 6 kW mono 220V 2 MPPT", ac=6.0,
         mppt="2", volt="220 V (2 fases)", telha="Fibrocimento", gmod="nao informada", ginv="nao informada",
         gserv="nao informada", excl="nao informadas", price=18228.70, extra=0, homol=790,
         instal="=6900/10*{c}{QTY}", promise=10554, mat=True, area=None, pag="PIX (cotacao de material)"),
    dict(L="B) TecSolarSP\nAstronergy 605W\nHuawei 6 kW", emp="TecSolarSP", nome="nao informado (razao social: Edijane Gomes dos Santos)", fone="(11) 97365-9192",
         cnpj="29.775.600/0001-03", doc="Cotacao Belenergy WEB-006512519", tipo="MATERIAL apenas",
         qty=12, wp=605, mod="Astronergy bifacial N-type", inv="Huawei 6 kW mono 220V 2 MPPT", ac=6.0,
         mppt="2", volt="220 V (2 fases)", telha="Fibrocimento", gmod="nao informada", ginv="nao informada",
         gserv="nao informada", excl="nao informadas", price=17307.90, extra=0, homol=790,
         instal="=6900/10*{c}{QTY}", promise=0, mat=True, area=None, pag="PIX (cotacao de material)"),
    dict(L="C) SunWash\nTCL 620W\nAUXSOL 5 kW", emp="SunWash Energia Solar e Ar-Condicionado", nome="Yuri", fone="(16) 99353-4346",
         cnpj="NAO INFORMADO", doc="WhatsApp 10/08/2026", tipo="INSTALADO (turnkey)",
         qty=10, wp=620, mod="TCL 620 W (tecnologia a confirmar)", inv="AUXSOL 5 kW (modelo nao informado)",
         ac=5.0, mppt="a confirmar", volt="220 V (a confirmar)", telha="nao informada", gmod="nao informada",
         ginv="nao informada", gserv="nao informada", excl="nao informadas", price=14220.34, extra=1800,
         homol=0, instal=0, promise=8400, mat=False, area=None, pag="a vista ou 21x COM juros"),
    dict(L="D) SunWash\nTCL 620W\nHUAWEI 5 kW", emp="SunWash Energia Solar e Ar-Condicionado", nome="Yuri", fone="(16) 99353-4346",
         cnpj="NAO INFORMADO", doc="WhatsApp 10/08/2026", tipo="INSTALADO (turnkey)",
         qty=10, wp=620, mod="TCL 620 W (tecnologia a confirmar)", inv="HUAWEI 5 kW (modelo nao informado)",
         ac=5.0, mppt="a confirmar", volt="220 V (a confirmar)", telha="nao informada", gmod="nao informada",
         ginv="nao informada", gserv="nao informada", excl="nao informadas", price=15220.34, extra=1800,
         homol=0, instal=0, promise=8400, mat=False, area=None, pag="a vista ou 21x COM juros"),
    dict(L="E) William\nMaxeon 535W\nSolis 4 kW", emp="nao informada na proposta", nome="William", fone="(11) 99636-5333",
         cnpj="NAO INFORMADO", doc="Proposta tecnica de 7 paginas", tipo="INSTALADO (turnkey)",
         qty=10, wp=535, mod="Maxeon bifacial PERC", inv="Solis 4 kW mono 220V 2 MPPT", ac=4.0,
         mppt="2", volt="220 V (2 fases)", telha="Colonial", gmod="nao informada", ginv="nao informada",
         gserv="nao informada", excl="nao informadas", price=11158.91, extra=0, homol=790, instal=6900,
         promise=6151, mat=True, area=None, pag="PIX no kit; restante nao informado"),
    dict(L="F) Sun Coast v1\nRONMA 585W\nAUXSOL 3 kW\n(SUPERADA)", emp="Sun Coast Materiais Eletricos e Energia Solar", nome="Alexandre", fone="(11) 91544-8947",
         cnpj="47.864.839/0001-31", doc="Proposta 40436 de 11/08/2026 (val. 20/08)", tipo="INSTALADO (turnkey)",
         qty=8, wp=585, mod="RONMA SOLAR bifacial N-type 144 cel", inv="AUXSOL 3 kW mono 220V 1 MPPT",
         ac=3.0, mppt="1", volt="220 V (2 fases)", telha="Colonial",
         gmod="25 anos eficiencia / 15 anos defeito", ginv="10 a 25 anos", gserv="12 meses",
         excl="padrao de entrada, obras civis, telhado, eletrodutos, ART estrutural",
         price=12387.64, extra=0, homol=0, instal=0, promise=6367.50, mat=False, area=21.12,
         pag="a vista, cartao 12x/21x COM juros ou financiamento 36x"),
    dict(L="G) SunWash\nMICRO 10x650W\n4x 2,25 kW", emp="SunWash Energia Solar e Ar-Condicionado", nome="Yuri", fone="(16) 99353-4346",
         cnpj="NAO INFORMADO", doc="Proposta 6695731 de 12/08/2026 - VALIDADE 17/08 (5 dias)", tipo="INSTALADO (turnkey)",
         qty=10, wp=650, mod="ZNShine 650W qt.10 (pag.5) x RISEN 660W qt.16 (tabela) - CONTRADICAO", inv="4x DEYE SUN-S225G4-EU-Q0 2,25 kW = 9,0 kW AC",
         ac=9.0, mppt="4 por micro = 16 canais (10 usados, 6 livres)", volt="220 V (a confirmar)",
         telha="estrutura apropriada ao telhado (tipo nao especificado)",
         gmod="12 anos defeito / 25 anos eficiencia",
         ginv="12 anos",
         gserv="vistoria, projeto eletrico, ART, licencas, gestao de obra, frete e documentacao",
         excl="alvenaria, reforco estrutural e alteracoes na rede pedidas pela concessionaria",
         price=16780.54, extra=1800, homol=0, instal=0, promise=8400, mat=False, area=25.0,
         padrao_incluso=True,
         pag="18x de R$ 932,25 SEM JUROS (= preco a vista)"),
    dict(L="H) Sfero (Renato)\nLongi/DAH 590W\n2x micro 2,5 kW", emp="Sfero", nome="nao informado (indicacao do Renato)", fone="(11) 99705-1311 (Renato)",
         cnpj="NAO INFORMADO na proposta", doc="Proposta 0664/2026 de 11/08/2026 (val. 10 dias)",
         tipo="INSTALADO (turnkey)", qty=10, wp=590, mod="Longi ou DAH Solar (Tier 1 de verdade)",
         inv="2x microinversor 2,5 kW = 5,0 kW AC - Solis ou Growatt", ac=5.0,
         mppt="1 por modulo (10 MPPTs)", volt="220 V (a confirmar)", telha="Ceramico / colonial",
         gmod="12 anos produto / 25 anos eficiencia", ginv="7 a 10 anos",
         gserv="1 ano instalacao + SEGURO ALL RISK 1 ano + GARANTIA DE PERFORMANCE EM CONTRATO",
         excl="obras civis, reformas no telhado/laje e no PADRAO DE ENTRADA",
         price=18733.00, extra=0, homol=0, instal=0, promise=7329.96, mat=False, area=28.0,
         pag="3x: 60% entrada / 20% montado / 20% ligado; ou Santander 60x R$ 553 ou 72x R$ 507"),
    dict(L="I) Sun Coast v2\nTCL 620W x11\nAUXSOL 7,5 kW", emp="Sun Coast Materiais Eletricos e Energia Solar", nome="Alexandre", fone="(11) 91544-8947",
         cnpj="47.864.839/0001-31", doc="Proposta 40438 de 11/08/2026 (val. 20/08)",
         tipo="INSTALADO (turnkey)", qty=11, wp=620, mod="TCL Solar bifacial N-type 132 cel",
         inv="AUXSOL 7,5 kW mono 220V 2 MPPT (100% over)", ac=7.5, mppt="2", volt="220 V (2 fases)",
         telha="Colonial", gmod="25 anos eficiencia / 15 anos defeito", ginv="10 a 25 anos",
         gserv="12 meses", excl="padrao de entrada, obras civis, telhado, eletrodutos, ART estrutural",
         price=16890.00, extra=0, homol=0, instal=0, promise=9279.13, mat=False, area=30.58,
         pag="a vista, cartao 12x/21x COM juros ou financiamento 36x ate 120x"),
    dict(L="J) AD BioSolar\nJA Solar 615W x10\nHuawei SUN2000-6KTL-L1", emp="AD BioSolar", nome="Adriano Melo (proposta) / Carlos (negociacao)",
         fone="(11) 97350-3103 / Carlos (11) 91408-8892",
         cnpj="64.313.573/0001-58", doc="Proposta #01237 de 12/08/2026 (val. 27/08) - substitui a #01225",
         tipo="INSTALADO (turnkey)", qty=10, wp=615, mod="JA Solar bifacial N-type 132 cel - TIER 1",
         inv="HUAWEI SUN2000-6KTL-L1 6.000 W (modelo EXATO informado)", ac=6.0, mppt="2 (a confirmar)",
         volt="220 V (2 fases)", telha="nao informada", gmod="15 anos defeito / 30 anos eficiencia",
         ginv="10 anos", gserv="vistoria tecnica, ART, licencas, gestao de obra e frete inclusos",
         excl="alvenaria, reforco estrutural e alteracoes na rede pedidas pela concessionaria",
         price=16200.00, extra=0, homol=0, instal=0, promise=7981.26, mat=False, area=32.41,
         pag="R$ 16.200 NEGOCIADO (o PDF traz R$ 21.289 - EXIGIR O DESCONTO POR ESCRITO)"),
    dict(L="SIMULACAO G12\nMICRO 12x650W\n4x Deye 2,25 kW", emp="SunWash - SIMULACAO, preco a confirmar",
         nome="Yuri", fone="(16) 99353-4346", cnpj="NAO INFORMADO",
         doc="SIMULACAO: opcao G escalada para 12 modulos", tipo="INSTALADO (turnkey)",
         qty=12, wp=650, mod="ZNShine ZXNR-MD132-650 N-type (Tier 1 BNEF)",
         inv="4x DEYE SUN-S225G4-EU-Q0 = 9,0 kW AC (12 dos 16 canais)", ac=9.0,
         mppt="4 por micro = 16 canais (12 usados, 4 livres)", volt="220 V (a confirmar)",
         telha="nao informada", gmod="exigir por escrito", ginv="exigir por escrito (Deye pratica 10 anos)",
         gserv="exigir por escrito", excl="exigir por escrito",
         price="={GPRICE}+2*ENTRADAS!$C$54", extra=1800, homol=0, instal=0, promise=0, mat=False, area=None, padrao_incluso=True,
         pag="pedir preco e parcelamento para esta configuracao"),
    dict(L="SIMULACAO G14\nMICRO 14x650W\n4x Deye 2,25 kW", emp="SunWash - SIMULACAO, preco a confirmar",
         nome="Yuri", fone="(16) 99353-4346", cnpj="NAO INFORMADO",
         doc="SIMULACAO: opcao G escalada para 14 modulos", tipo="INSTALADO (turnkey)",
         qty=14, wp=650, mod="ZNShine ZXNR-MD132-650 N-type (Tier 1 BNEF)",
         inv="4x DEYE SUN-S225G4-EU-Q0 = 9,0 kW AC (14 dos 16 canais)", ac=9.0,
         mppt="4 por micro = 16 canais (14 usados, 2 livres)", volt="220 V (a confirmar)",
         telha="nao informada", gmod="exigir por escrito", ginv="exigir por escrito (Deye pratica 10 anos)",
         gserv="exigir por escrito", excl="exigir por escrito",
         price="={GPRICE}+4*ENTRADAS!$C$54", extra=1800, homol=0, instal=0, promise=0, mat=False, area=None, padrao_incluso=True,
         pag="pedir preco e parcelamento para esta configuracao"),
    dict(L="M) Bruno G4\n(a receber)", emp="G4 (razao social a confirmar)", nome="Bruno", fone="(11) 96861-7830", cnpj="", doc="Nada recebido",
         tipo="?", qty="", wp="", mod="", inv="", ac="", mppt="", volt="", telha="", gmod="", ginv="",
         gserv="", excl="", price="", extra="", homol="", instal="", promise="", mat=False, area=None, pag=""),
]
NOPT = len(OPT)
CL = [get_column_letter(3 + i) for i in range(NOPT)]   # C..I
NOTECOL = 3 + NOPT                                     # J

ws = sheet("COMPARATIVO")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 42
for cl in CL:
    ws.column_dimensions[cl].width = 17
ws.column_dimensions[get_column_letter(NOTECOL)].width = 62
title(ws, "COMPARATIVO DOS ORCAMENTOS - QUADRO PRINCIPAL",
      "Contatos nas linhas 5 a 7. Colunas A e B: TecSolarSP, preco de MATERIAL. As demais: preco INSTALADO. A linha do PRECO TOTAL INSTALADO e a unica comparavel entre todos.")
hdr_row(ws, 4, ["ITEM"] + [o["L"] for o in OPT] + ["OBSERVACOES / ALERTAS"], start=2)
ws.row_dimensions[4].height = 52

RW = {}
_cur = [4]


def _nxt(key=None):
    _cur[0] += 1
    if key:
        RW[key] = _cur[0]
    return _cur[0]


def crow(label, vals, fmt=None, note="", bold=False, fill=None, inputs=False, textrow=False, key=None):
    r = _nxt(key)
    c = ws.cell(row=r, column=2, value=label)
    c.font = BOLD if bold else BLACK
    c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, v in enumerate(vals):
        cell = ws.cell(row=r, column=3 + i)
        cell.value = v
        isf = isinstance(v, str) and v.startswith("=")
        if isf:
            cell.font = BOLD if bold else BLACK
        elif inputs and v != "":
            cell.font = BLUE
            cell.fill = YEL
        else:
            cell.font = BOLD if bold else BLACK
        if fmt and not textrow:
            cell.number_format = fmt
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        if fill and not inputs:
            cell.fill = fill
        if isinstance(v, str) and ("NAO INFORMADO" in v or "nao informada" in v or v.startswith("padrao de entrada")):
            cell.font = REDS
    cn = ws.cell(row=r, column=NOTECOL, value=note)
    cn.font = REDS if note.startswith("ALERTA") else SMALL
    cn.alignment = Alignment(wrap_text=True, vertical="top")
    return r


def sec(txt):
    _nxt()
    r = _nxt()
    ws.cell(row=r, column=2, value=txt).font = SUB
    for c in range(2, NOTECOL + 1):
        ws.cell(row=r, column=c).fill = BAND
    return r


def F_(f):
    return [f.replace("{c}", cl) if o["qty"] != "" else "" for cl, o in zip(CL, OPT)]


crow("EMPRESA / razao social", [o["emp"] for o in OPT], textrow=True, bold=True, key="emp",
     note="ALERTA: SunWash, William e Sfero nao informaram razao social. Sem empresa identificada nao existe garantia exigivel.")
crow("NOME DO RESPONSAVEL", [o["nome"] for o in OPT], textrow=True, bold=True, key="nome",
     note="Quem falou com voce. Peca sempre que a proposta saia com nome e CNPJ da empresa, nao apenas o contato pessoal.")
crow("TELEFONE DE CONTATO", [o["fone"] for o in OPT], textrow=True, bold=True, key="fone",
     note="A Sfero chegou por indicacao do Renato - o telefone e dele, nao da empresa. Pedir o contato direto da Sfero.")
crow("CNPJ informado na proposta", [o["cnpj"] for o in OPT], textrow=True, key="cnpj",
     note="ALERTA: so Sun Coast, AD BioSolar e TecSolarSP (via cotacao do distribuidor) tem CNPJ identificado.")
crow("Documento recebido", [o["doc"] for o in OPT], textrow=True, key="doc",
     note="ALERTA: orcamento por WhatsApp nao e proposta comercial. Sfero, Sun Coast, AD BioSolar e William entregaram documento formal.")
crow("TIPO DE PRECO", [o["tipo"] for o in OPT], textrow=True, bold=True, fill=BAD_F, key="tipo",
     note="ALERTA: comparar MATERIAL com INSTALADO e o erro mais comum. As linhas de homologacao e instalacao corrigem isso.")
crow("Quantidade de modulos", [o["qty"] for o in OPT], NUM0, inputs=True, key="qty", note="Informada por cada fornecedor.")
crow("Potencia por modulo (W)", [o["wp"] for o in OPT], NUM0, inputs=True, key="wp",
     note="A SunWash escreveu '620kw' e '650w' - erro de digitacao para watts.")
crow("Fabricante / tecnologia do modulo", [o["mod"] for o in OPT], textrow=True, key="mod",
     note="ALERTA: Maxeon do William e PERC, uma geracao atras. TCL e RONMA tem pouquissima rede de assistencia no Brasil. JA Solar, Longi e DAH sao Tier 1 com representacao aqui.")
crow("POTENCIA DC (kWp)", F_("={c}%d*{c}%d/1000" % (RW["qty"], RW["wp"])), NUM2, bold=True, fill=BAND, key="kwp",
     note="Quantidade x potencia do modulo / 1000.")
crow("Inversor", [o["inv"] for o in OPT], textrow=True, key="inv",
     note="ALERTA: AUXSOL tem baixa penetracao no Brasil. Huawei, Solis e Growatt tem assistencia estabelecida. So a AD BioSolar informou o MODELO EXATO.")
crow("Potencia AC do inversor (kW)", [o["ac"] for o in OPT], NUM2, inputs=True, key="ac",
     note="Exigir o codigo exato do modelo, como fez a AD BioSolar (SUN2000-5KTL-L1).")
crow("Numero de MPPTs", [o["mppt"] for o in OPT], textrow=True, key="mppt",
     note="ALERTA: Sun Coast v1 usava 1 MPPT UNICO para 8 modulos. As opcoes com microinversor dao 1 MPPT POR MODULO: sombra em um modulo nao afeta os outros e permite usar mais de uma agua do telhado.")
crow("RELACAO DC/AC", F_("=IFERROR({c}%d/{c}%d,0)" % (RW["kwp"], RW["ac"])), '0.00"x"', bold=True, key="dcac",
     note="ALERTA: faixa saudavel de 1,10 a 1,35. Sun Coast v1 em 1,56 = clipping e risco de perder garantia. Abaixo de 1,10 voce pagou por inversor que nao vai usar - aceitavel se for folga proposital para expansao.")
crow("Tensao de operacao", [o["volt"] for o in OPT], textrow=True, fill=BAD_F, key="volt",
     note="ALERTA MAIOR: sua ligacao e MONOFASICA (127 V na Enel SP). TODOS exigem mudanca para bifasico.")
crow("Tipo de telha que o orcamento considerou", [o["telha"] for o in OPT], textrow=True, key="telha",
     note="ALERTA: William, Sun Coast e Sfero dizem CERAMICA/COLONIAL; a TecSolarSP diz FIBROCIMENTO. Tres contra um. Se for ceramica, a estrutura das duas cotacoes da TecSolarSP esta errada. Uma foto resolve.")
crow("Area de telhado necessaria (m2)",
     [(o["area"] if o["area"] else (f"={cl}{RW['qty']}*{cl}{RW['wp']}*0.00452" if o["qty"] != "" else "")) for cl, o in zip(CL, OPT)],
     NUM2, key="area",
     note="Valores em azul foram DECLARADOS pelo fornecedor; os demais estimados a 0,00452 m2 por watt. Sem fotos do telhado NAO foi verificado se essa area existe na agua Sudeste.")
crow("Garantia declarada do modulo", [o["gmod"] for o in OPT], textrow=True, key="gmod",
     note="AD BioSolar tem a melhor: 30 anos de eficiencia e 15 de defeito. TecSolarSP, SunWash e William nao declararam nada.")
crow("Garantia declarada do inversor", [o["ginv"] for o in OPT], textrow=True, key="ginv",
     note="AD BioSolar declara 10 anos para um modelo Huawei nomeado - e a unica verificavel. '10 a 25 anos' da Sun Coast e vago.")
crow("Garantia de servico / seguro", [o["gserv"] for o in OPT], textrow=True, key="gserv",
     note="So a Sfero oferece seguro All Risk e garantia de performance em contrato. Isso vale dinheiro em risco.")
crow("Exclusoes declaradas na proposta", [o["excl"] for o in OPT], textrow=True, key="excl",
     note="ALERTA: Sun Coast, Sfero e AD BioSolar excluem por escrito a adequacao do padrao. E honesto declarar, mas o custo e seu. Quem nao declarou exclusao tambem nao incluiu - so nao avisou.")
crow("Condicao de pagamento declarada", [o["pag"] for o in OPT], textrow=True, key="pag",
     note="ATENCAO: 18x SEM JUROS pelo valor a vista (SunWash micro) e vantagem real. Ja o financiamento Sfero em 72x custa cerca de R$ 17,8 mil de juros sobre R$ 18,7 mil - quase dobra o preco.")

sec("PRECOS")
crow("Preco BASE informado (R$)",
     [(o["price"].replace("{GPRICE}", f"I{_cur[0]+1}") if isinstance(o["price"], str) and o["price"] else o["price"]) for o in OPT],
     MONEY, inputs=True, key="price",
     note="A, B e E: valor do KIT/material. As demais: projeto completo, ja com instalacao. ALERTA na coluna J: o PDF da AD BioSolar traz R$ 21.289,00 e o valor de R$ 16.200,00 foi negociado verbalmente - um desconto de R$ 5.089 (24%) que PRECISA constar em proposta assinada, senao pode evaporar na hora de fechar.")
crow("Acrescimo regional a confirmar (R$)", [o["extra"] for o in OPT], MONEY, inputs=True, key="extra",
     note="ALERTA: a SunWash citou +R$ 1.800 pela sua regiao no kit Intelbras e nao repetiu nos demais. Mantido por PRUDENCIA. Se nao incidir, zere e o payback melhora.")
crow("Homologacao / engenharia / ART (R$)", [o["homol"] for o in OPT], MONEY, inputs=True, key="homol",
     note="E: R$ 790 declarado pelo William. A e B: mesmo valor como ESTIMATIVA - a TecSolarSP nunca informou o preco de venda.")
crow("Instalacao / mao de obra (R$)",
     [(o["instal"].replace("{QTY}", str(RW["qty"])).replace("{c}", cl) if isinstance(o["instal"], str) and o["instal"] else o["instal"]) for cl, o in zip(CL, OPT)],
     MONEY, inputs=True, key="instal",
     note="ALERTA: E = R$ 6.900 declarado pelo William para 10 modulos. A e B usam esse valor escalado por modulo, como ESTIMATIVA.")
crow("Adequacao do padrao mono -> bifasico (R$)",
     [(0 if o.get("padrao_incluso") else (f"=ENTRADAS!$C$28" if o["qty"] != "" else "")) for o in OPT], MONEY, key="padrao",
     note="ALERTA: so a proposta do Yuri (colunas G, e as simulacoes) lista 'PADRAO DE ENTRADA' e 'TROCA DE DISJUNTOR DO PADRAO' como inclusos - por isso zero nelas. MAS a pagina 8 da mesma proposta diz que a adequacao de entrada e responsabilidade do cliente. CONTRADICAO A RESOLVER POR ESCRITO: vale R$ 2.500.")
crow("PRECO TOTAL INSTALADO (R$)", F_("=SUM({c}%d:{c}%d)" % (RW["price"], RW["padrao"])), MONEY, bold=True, fill=WARN_F, key="total",
     note="O UNICO numero comparavel entre fornecedores.")
crow("R$ por Wp instalado", F_("=IFERROR({c}%d/({c}%d*1000),0)" % (RW["total"], RW["kwp"])), WP, bold=True, key="wpinst",
     note="Metrica de mercado, valida so entre precos de mesmo escopo.")
crow("R$ por Wp somente de material",
     [(f"=IFERROR({cl}{RW['price']}/({cl}{RW['kwp']}*1000),0)" if o["mat"] else "n/d") if o["qty"] != "" else "" for cl, o in zip(CL, OPT)],
     WP, key="wpmat",
     note="Mede a margem. A TecSolarSP paga R$ 2,47/Wp de material: use isso para negociar o preco de venda dela.")

sec("GERACAO - REALIDADE x PROMESSA")
crow("Geracao anual realista - telhado SE (kWh)", F_("={c}%d*GERACAO!$I$17" % RW["kwp"]), NUM0, key="ger",
     note="Irradiacao de Sao Paulo, telhado Sudeste 20 graus, PR de 79%.")
crow("Geracao media mensal realista (kWh/mes)", F_("={c}%d/12" % RW["ger"]), NUM0, key="germes")
crow("Geracao no cenario CONSERVADOR (PR 72%)", F_("={c}%d/ENTRADAS!$C$31*ENTRADAS!$C$32" % RW["ger"]), NUM0, key="gercons",
     note="Pior caso: telhado mais sujo, mais quente ou com sombra parcial.")
crow("Geracao PROMETIDA pelo vendedor (kWh/ano)", [o["promise"] for o in OPT], NUM0, inputs=True, key="promise",
     note="A: 10.554. C, D e G: 700 kWh/mes x 12. E: 6.151. F: 6.367,50. H: 7.329,96. I: 9.279,13. J: 7.981,26. B: nao informada.")
crow("DESVIO DA PROMESSA", F_("=IFERROR({c}%d/{c}%d-1,0)" % (RW["ger"], RW["promise"])), PCT, bold=True, fill=BAD_F, key="desvio",
     note="ALERTA: negativo = o vendedor promete mais do que o telhado entrega. Positivo = foi conservador. Sfero e a unica positiva.")
crow("Geracao se o telhado fosse NORTE (kWh)", F_("={c}%d*GERACAO!$J$17" % RW["kwp"]), NUM0, key="gernorte",
     note="Compare com a linha da promessa: e daqui que saem os numeros otimistas.")

sec("ATENDIMENTO DO CONSUMO")
crow("% atendido SEM o BYD (Cenario 1)", F_("=IFERROR({c}%d/(BYD!$F$33*12),0)" % RW["ger"]), PCT, key="pct1",
     note="Todos parecem otimos aqui - e por isso que os vendedores acham que acertaram o tamanho.")
crow("% atendido COM o BYD (Cenario 2)", F_("=IFERROR({c}%d/(BYD!$F$34*12),0)" % RW["ger"]), PCT, key="pct2")
crow("% atendido no CENARIO 3", F_("=IFERROR({c}%d/(BYD!$F$35*12),0)" % RW["ger"]), PCT, bold=True, fill=WARN_F, key="pct3",
     note="ESTE e o criterio de dimensionamento. Alvo: 95% a 110%. Abaixo de 90% = subdimensionado.")
crow("Deficit / excedente anual (kWh)", F_("={c}%d-BYD!$G$35" % RW["ger"]), NUM0, key="deficit",
     note="Negativo = continua comprando energia todo mes. Muito positivo = credito que expira em 60 meses.")
crow("Modulos que faltam para o Cenario 3",
     F_("=IFERROR(ROUNDUP(MAX(0,BYD!$G$35-{c}%d)/GERACAO!$I$17*1000/{c}%d,0),0)" % (RW["ger"], RW["wp"])), NUM0, bold=True, key="faltam")
crow("Modulos atribuiveis SO ao BYD",
     F_("=IFERROR(ROUNDUP(BYD!$C$13*12/GERACAO!$I$17*1000/{c}%d,0),0)" % RW["wp"]), NUM0, bold=True, fill=WARN_F, key="sobyd",
     note="Placas que existem no projeto apenas para compensar o carregamento do carro.")
crow("Potencia AC a declarar na homologacao e corrente em 220 V",
     F_('={c}%d&" kW = "&TEXT({c}%d*1000/220,"0")&" A"' % (RW["ac"], RW["ac"])), None, key="homologac",
     note="ALERTA POUCO OBVIO: a potencia declarada a Enel e a do INVERSOR (AC), nao a dos modulos. Quanto maior, maior o disjuntor e o padrao exigidos - e mais caro fica o item de adequacao do padrao.")

sec("RESULTADO FINANCEIRO")
crow("Conta mensal HOJE (sem solar, sem BYD)", F_("=CONTA!$C$54"), MONEY, key="conta_hoje", note="Situacao atual real: R$ 410,63.")
crow("Conta mensal SEM solar mas COM o BYD", F_("=ECONOMIA!{c}8/12"), MONEY, key="conta_sem", note="O que a conta vira se o carro chegar e nao houver solar.")
crow("Conta mensal COM solar e COM o BYD", F_("=ECONOMIA!{c}19"), MONEY, bold=True, key="conta_com",
     note="ALERTA: nao vai a zero. Sobram disponibilidade, Fio B e COSIP.")
crow("Economia liquida ano 1 (R$/ano)", F_("=ECONOMIA!{c}26"), MONEY, key="econ_ano")
crow("Economia liquida ano 1 (R$/mes)", F_("={c}%d/12" % RW["econ_ano"]), MONEY, bold=True, key="econ_mes")
crow("PAYBACK SIMPLES (anos)", F_("=IFERROR({c}%d/{c}%d,0)" % (RW["total"], RW["econ_ano"])), NUM2, bold=True, fill=OK_F, key="payback",
     note="Investimento total dividido pela economia do ano 1. CUIDADO: payback bom nao significa melhor negocio - ver a linha do lucro em 20 anos.")
PRJ, _k = [], 0
for _o in OPT:
    if _o["qty"] != "":
        PRJ.append(get_column_letter(5 + 3 * _k)); _k += 1
    else:
        PRJ.append(None)
crow("Economia acumulada em 5 anos (R$)", [f"=PROJECAO!{c}10" if c else "" for c in PRJ], MONEY0, key="acum5")
crow("Economia acumulada em 10 anos (R$)", [f"=PROJECAO!{c}15" if c else "" for c in PRJ], MONEY0, key="acum10")
crow("Economia acumulada em 20 anos (R$)", [f"=PROJECAO!{c}25" if c else "" for c in PRJ], MONEY0, key="acum20")
crow("LUCRO LIQUIDO EM 20 ANOS (R$)", F_("={c}%d-{c}%d" % (RW["acum20"], RW["total"])), MONEY0, bold=True, fill=OK_F, key="lucro20",
     note="Economia de 20 anos menos o investimento. E O MELHOR CRITERIO DE DECISAO FINANCEIRA - melhor que o payback.")
crow("Retorno sobre o investimento em 20 anos", F_("=IFERROR({c}%d/{c}%d,0)" % (RW["lucro20"], RW["total"])), PCT, key="roi")

sec("EXPANSAO FUTURA")
crow("kWp maximo com o mesmo inversor (DC/AC 1,35)", F_("={c}%d*1.35" % RW["ac"]), NUM2, key="kwpmax")
crow("Modulos que ainda cabem no inversor",
     F_("=MAX(0,ROUNDDOWN(({c}%d-{c}%d)*1000/{c}%d,0))" % (RW["kwpmax"], RW["kwp"], RW["wp"])), NUM0, bold=True, key="modextra",
     note="ALERTA: os inversores de 5 kW da SunWash (C e D) nao aceitam nenhum modulo a mais. Sun Coast v2 (+5) e a opcao com microinversores tem folga real - mas nos micros o limite verdadeiro e o numero de CANAIS livres, nao a potencia.")
crow("LIMITE REAL de expansao (canais e estrutura)",
     ["1 modulo (limite de potencia)", "1 modulo (limite de potencia)", "0 - inversor no limite",
      "0 - inversor no limite", "0 - inversor no limite", "0 - inversor sobrecarregado",
      "6 MODULOS - canais livres nos 4 micros Deye", "1 modulo; para 12 precisa de um 3o micro",
      "5 modulos", "2 a 3 modulos", "4 modulos (canais livres)", "2 modulos (canais livres)", ""],
     textrow=True, key="limreal",
     note="A linha acima calcula pela POTENCIA do inversor. Em microinversor o limite verdadeiro e o numero de CANAIS livres. Os 4 micros Deye SUN-S225G4-EU-Q0 tem 4 MPPTs cada = 16 canais para 10 modulos, ou seja 6 canais livres: da para ir a 10,4 kWp sem comprar inversor nenhum.")
crow("Geracao extra possivel sem trocar inversor (kWh/ano)",
     F_("={c}%d*{c}%d/1000*GERACAO!$I$17" % (RW["modextra"], RW["wp"])), NUM0, key="gerextra")
crow("Situacao do inversor",
     F_('=IF({c}%d>1.35,"JA SOBRECARREGADO",IF({c}%d<1.1,"SUBAPROVEITADO","OK"))' % (RW["dcac"], RW["dcac"])), None, bold=True, key="sitinv",
     note="'JA SOBRECARREGADO' = inversor pequeno para os modulos desde o dia 1 (perde geracao e garantia). 'SUBAPROVEITADO' = folga; boa se proposital para expansao, ruim se foi venda a mais.")

_r0 = _nxt() + 1
ws.cell(row=_r0, column=2, value="DECISAO ESPECIFICA: AUXSOL x HUAWEI NA PROPOSTA DA SUNWASH").font = SUB
for c in range(2, NOTECOL + 1):
    ws.cell(row=_r0, column=c).fill = BAND
ws.cell(row=_r0 + 1, column=2, value="Diferenca de preco (Huawei - AUXSOL)").font = BOLD
frm(ws, f"C{_r0+1}", "=F%d-E%d" % (RW["price"], RW["price"]), MONEY, BOLD)
ws[f"C{_r0+1}"].fill = WARN_F
ws.cell(row=_r0 + 2, column=2, value="Diferenca no payback, em anos").font = BLACK
frm(ws, f"C{_r0+2}", "=IFERROR(F%d-E%d,0)" % (RW["payback"], RW["payback"]), NUM2)
ws.cell(row=_r0 + 3, column=2, value="Diferenca no lucro de 20 anos").font = BLACK
frm(ws, f"C{_r0+3}", "=F%d-E%d" % (RW["lucro20"], RW["lucro20"]), MONEY0)
for i, t in enumerate([
    "RECOMENDACAO: entre as duas da SunWash com inversor string, pagar os R$ 1.000 e ficar com o HUAWEI.",
    "O inversor e o unico componente que quase certamente sera trocado dentro da vida do sistema. Huawei tem garantia de fabrica mais longa,",
    "assistencia estabelecida no Brasil, monitoramento maduro e valor de revenda. AUXSOL tem penetracao pequena: se a marca sair do mercado,",
    "voce fica com o sistema parado esperando peca. O atraso de cerca de 2 meses no payback nao paga esse risco.",
    "",
    "OBSERVACAO SOBRE PAYBACK x LUCRO: sistema pequeno se paga rapido e depois te deixa pagando conta para sempre.",
    "Payback ignora a energia que voce NAO compensou. Por isso o criterio correto e o LUCRO LIQUIDO EM 20 ANOS.",
]):
    c = ws.cell(row=_r0 + 5 + i, column=2, value=t)
    c.font = RED if t.startswith(("RECOMENDACAO", "OBSERVACAO")) else BLACK

# =====================================================================
# BOM_DETALHADO
# =====================================================================
ws = sheet("BOM_DETALHADO")
for col, w in [("A", 3), ("B", 44), ("C", 20), ("D", 20), ("E", 20), ("F", 58)]:
    ws.column_dimensions[col].width = w
title(ws, "LISTA DE MATERIAIS - AS TRES COTACOES BELENERGY LADO A LADO",
      "Os tres orcamentos vem do MESMO distribuidor (Belenus/BelEnergy). Isso permite comparar escopo item por item.")
hdr_row(ws, 4, ["ITEM", "A) TecSolarSP\nJA Solar 615W\nWEB-006506779", "B) TecSolarSP\nAstronergy 605W\nWEB-006512519",
                "E) William\nMaxeon 535W\nWEB-006493952", "ANALISE TECNICA"], start=2)
ws.row_dimensions[4].height = 50
bom = [
    ("Modulo fotovoltaico", "12 pc JA Solar 615 W bifacial N-type", "12 pc Astronergy 605 W bifacial N-type",
     "10 pc Maxeon 535 W bifacial PERC", "A e B usam N-type TOPCon (geracao atual). E usa PERC, tecnologia anterior: menor eficiencia por m2 e degradacao maior."),
    ("Inversor", "Huawei 6 kW mono 220 V 2 MPPT", "Huawei 6 kW mono 220 V 2 MPPT", "Solis 4 kW mono 220 V 2 MPPT",
     "TODOS de 220 V - incompativeis com o padrao monofasico atual sem adequacao. O Solis de 4 kW limita a expansao futura a zero."),
    ("Estrutura - suporte de telhado", "18 pc suporte pe em L FIBROCIMENTO", "18 pc suporte pe em L FIBROCIMENTO",
     "10 pc gancho telha COLONIAL c/ prolongador", "CONTRADICAO GRAVE: TecSolarSP cotou FIBROCIMENTO; William e Sun Coast cotaram TELHA COLONIAL. Dois contra um. Se for colonial, a estrutura das duas cotacoes da TecSolarSP esta errada. Resolver com uma foto antes de comprar."),
    ("Haste / fixacao", "18 pc haste 10x250 mm inox (madeira)", "20 pc haste 10x250 mm inox", "incluso no gancho",
     "Fixacao por perfuracao exige vedacao com arruela EPDM. Exigir garantia contra infiltracao."),
    ("Perfil de fixacao", "24 pc de 2,70 m", "24 pc de 2,70 m", "10 pc de 2,28 m",
     "Compativel com a quantidade de modulos de cada projeto."),
    ("Grampos final / intermediario", "6 jg final + 6 jg intermediario", "6 jg final + 6 jg intermediario",
     "2 jg final + 10 jg intermediario", "Sem observacao relevante."),
    ("Garra de aterramento", "6 jg", "3 jg", "2 jg",
     "ALERTA: quantidade cai muito em B e E. Aterramento de moldura e estrutura precisa ser continuo em todas as fileiras."),
    ("Cabo solar CC", "30 m vermelho + 30 m preto de 6 mm2", "30 m vermelho + 30 m preto de 4 mm2",
     "25 m vermelho + 25 m preto de 4 mm2", "Bitola de 4 mm2 atende em corrente, mas verificar queda de tensao no percurso real telhado-inversor."),
    ("Cabo de aterramento / equipotencializacao", "20 m de 4 mm2 verde", "NAO INCLUIDO", "10 m de 6 mm2 verde",
     "ALERTA: a cotacao B nao tem condutor de equipotencializacao. Item de seguranca, nao e opcional."),
    ("String box CC", "1 pc 2E/2S 1000 V Clamper", "1 pc 2E/2S 1000 V Clamper", "1 pc 2-4E/2-4S 1000 V",
     "Verificar em todas se contem DPS CC e chave seccionadora sob carga."),
    ("DPS do lado CA", "NAO INCLUIDO", "NAO INCLUIDO", "3 pc Clamper classe II 275 V 15 kA",
     "ALERTA: so o William incluiu protecao contra surto no lado CA. Nas duas da TecSolarSP falta - e obrigatorio pela NBR 5410."),
    ("Cabo CA", "100 m flex 6 mm2 preto", "100 m flex 6 mm2 preto", "nao listado separadamente",
     "Falta o condutor neutro/terra e o disjuntor CA dedicado em A e B. Confirmar se entram na instalacao."),
    ("Disjuntor CA dedicado", "NAO INCLUIDO", "NAO INCLUIDO", "NAO LISTADO",
     "ALERTA: nenhuma das tres lista o disjuntor de protecao do inversor. Confirmar por escrito."),
    ("Validacao da estrutura pelo distribuidor", "OK - 6 linhas x 2 modulos, paisagem",
     "NAO RECOMENDADA pela BelEnergy", "nao mencionado",
     "ALERTA: a propria cotacao B avisa que a estrutura escolhida nao foi validada pelo distribuidor. Risco mecanico transferido ao integrador."),
    ("Validade da cotacao", "3 dias (emitida 08/08/2026)", "3 dias (emitida 10/08/2026)", "PIX (WEB-006493952)",
     "Cotacoes de material valem 3 dias. Os precos provavelmente ja mudaram - pedir revalidacao antes de fechar."),
    ("Garantia do distribuidor", "1 ano Belenus", "1 ano Belenus", "1 ano Belenus",
     "A Belenus cobre 1 ano. Todo o resto (25 a 30 anos do modulo, 5 a 10 do inversor) e garantia do FABRICANTE, nao do integrador."),
]
r = 5
for item, a, b, e, obs in bom:
    ws.cell(row=r, column=2, value=item).font = BLACK
    for j, v in enumerate([a, b, e]):
        c = ws.cell(row=r, column=3 + j, value=v)
        c.font = RED if ("NAO INCLUIDO" in v or "NAO RECOMENDADA" in v) else BLACK
        c.alignment = Alignment(wrap_text=True, vertical="top")
    co = ws.cell(row=r, column=6, value=obs)
    co.font = REDS if obs.startswith("ALERTA") or obs.startswith("CONTRADICAO") else SMALL
    co.alignment = Alignment(wrap_text=True, vertical="top")
    for cc in range(2, 7):
        ws.cell(row=r, column=cc).border = BOX
    r += 1

r += 1
ws.cell(row=r, column=2, value="FECHAMENTO FINANCEIRO DAS COTACOES DE MATERIAL").font = SUB
base = r + 1
hdr_row(ws, base, ["Componente", "A) JA Solar 615W", "B) Astronergy 605W", "E) William Maxeon 535W"], start=2)
fin = [("Total de produtos", 17344.82, 16588.37, 10565.21),
       ("Frete", 732.44, 719.53, 593.70),
       ("Seguro", 151.44, 0.00, 0.00),
       ("ICMS / IPI / ST / DIFAL", 0.00, 0.00, 0.00)]
for i, (lab, a, b, e) in enumerate(fin):
    rr = base + 1 + i
    ws.cell(row=rr, column=2, value=lab).font = BLACK
    inp(ws, f"C{rr}", a, MONEY)
    inp(ws, f"D{rr}", b, MONEY)
    inp(ws, f"E{rr}", e, MONEY)
rt = base + 1 + len(fin)
ws.cell(row=rt, column=2, value="TOTAL DO KIT (material)").font = BOLD
for cl in ["C", "D", "E"]:
    frm(ws, f"{cl}{rt}", f"=SUM({cl}{base+1}:{cl}{rt-1})", MONEY, BOLD)
    ws[f"{cl}{rt}"].fill = WARN_F
    ws[f"{cl}{rt}"].border = BOX
rr = rt + 2
ws.cell(row=rr, column=2, value="Potencia do sistema (kWp)").font = BLACK
for cl, src in [("C", f"C{RW['kwp']}"), ("D", f"D{RW['kwp']}"), ("E", f"G{RW['kwp']}")]:
    frm(ws, f"{cl}{rr}", f"=COMPARATIVO!{src}", NUM2, GREEN)
rr += 1
ws.cell(row=rr, column=2, value="Custo do material por Wp").font = BOLD
for cl in ["C", "D", "E"]:
    frm(ws, f"{cl}{rr}", f"=IFERROR({cl}{rt}/({cl}{rr-1}*1000),0)", WP, BOLD)
    ws[f"{cl}{rr}"].fill = BAND
rr += 2
ws.cell(row=rr, column=2, value="Mao de obra + engenharia declarada pelo William").font = BLACK
ws.cell(row=rr, column=3, value="R$ 7.690,00 para 10 modulos = R$ 769 por modulo, ou R$ 1,44 por Wp").font = BLACK
rr += 1
ws.cell(row=rr, column=2, value="Margem que o William cobrou sobre o kit").font = BOLD
frm(ws, f"C{rr}", f"=IFERROR(7690/E{rt},0)", PCT, BOLD)
ws[f"C{rr}"].fill = BAND
ws.cell(row=rr, column=4, value="Referencia de mercado para negociar com a TecSolarSP: eles pagam R$ 18.228,70 de material.").font = SMALL
rr += 1
ws.cell(row=rr, column=2, value="Preco de venda justo estimado para a opcao A").font = BOLD
frm(ws, f"C{rr}", f"=C{rt}+790+6900/10*12", MONEY, BOLD)
ws[f"C{rr}"].fill = OK_F
ws.cell(row=rr, column=4, value="Se a TecSolarSP pedir muito acima disso, ha espaco para negociar - agora voce sabe o custo deles.").font = SMALL
rr += 2
ws.cell(row=rr, column=2, value="A opcao A (615 W) vale os R$ 920,80 a mais que a opcao B (605 W)?").font = SUB
rr += 1
ws.cell(row=rr, column=2, value="Diferenca de preco do material").font = BLACK
frm(ws, f"C{rr}", f"=C{rt}-D{rt}", MONEY)
rr += 1
ws.cell(row=rr, column=2, value="Geracao anual extra da opcao A (kWh)").font = BLACK
frm(ws, f"C{rr}", f"=COMPARATIVO!C{RW['ger']}-COMPARATIVO!D{RW['ger']}", NUM0)
rr += 1
ws.cell(row=rr, column=2, value="Valor dessa geracao extra por ano").font = BLACK
frm(ws, f"C{rr}", f"=C{rr-1}*ENTRADAS!C18", MONEY)
rr += 1
ws.cell(row=rr, column=2, value="Anos para a geracao extra pagar a diferenca").font = BOLD
frm(ws, f"C{rr}", f"=IFERROR(C{rr-3}/C{rr-1},0)", NUM2, BOLD)
ws[f"C{rr}"].fill = WARN_F
ws.cell(row=rr, column=4, value="Se der acima de ~6 anos, a diferenca NAO se paga pela geracao: decida por marca, garantia e completude do kit (a opcao A tem aterramento completo, a B nao).").font = SMALL

# =====================================================================
# ECONOMIA
# =====================================================================
ws = sheet("ECONOMIA")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 48
for cl in CL:
    ws.column_dimensions[cl].width = 15
ws.column_dimensions[get_column_letter(NOTECOL)].width = 56
title(ws, "MODELO DE ECONOMIA - ANO 1, PELAS REGRAS DA LEI 14.300/2022",
      "Base: Cenario 3 (residencia + BYD + margem). A conta NAO vai a zero: permanecem custo de disponibilidade, Fio B e COSIP.")
hdr_row(ws, 4, ["LINHA DE CALCULO"] + [o["L"].replace("\n", " ") for o in OPT] + ["EXPLICACAO"], start=2)
ws.row_dimensions[4].height = 34
eco = [
    (5, "Consumo anual - Cenario 3 (kWh)", "=BYD!$G$35", NUM0, "Residencia + BYD + margem de crescimento"),
    (6, "Geracao anual do sistema (kWh)", "=COMPARATIVO!{c}%d" % RW["ger"], NUM0, "Geracao realista no telhado Sudeste"),
    (7, "Tarifa efetiva (R$/kWh)", "=ENTRADAS!$C$18", MONEY, "R$ 1,0215 - calculado da fatura real de jul/2026"),
    (8, "CONTA ANUAL SEM SOLAR, COM O BYD (R$)", "={c}5*{c}7+ENTRADAS!$C$26*12", MONEY, "O cenario que voce enfrenta se nao instalar nada e o carro chegar"),
    (9, "Autoconsumo instantaneo (kWh)", "=MIN({c}5,{c}6)*ENTRADAS!$C$37", NUM0, "Usado no instante da geracao: nao passa pelo medidor, nao paga nada"),
    (10, "Energia injetada na rede (kWh)", "={c}6-{c}9", NUM0, "Vira credito de energia"),
    (11, "Consumo minimo faturavel - disponibilidade (kWh)", "=ENTRADAS!$C$17*12", NUM0, "50 kWh/mes apos virar bifasico. Nunca compensavel"),
    (12, "Energia compensada com creditos (kWh)", "=MAX(0,MIN({c}10,{c}5-{c}9-{c}11))", NUM0, "Limitada pelo consumo restante apos autoconsumo e disponibilidade"),
    (13, "Excedente que nao sera aproveitado (kWh)", "={c}10-{c}12", NUM0, "Credito com prazo de 60 meses. Sobredimensionar e desperdicio"),
    (14, "Energia ainda comprada da rede (kWh)", "=MAX({c}11,{c}5-{c}9-{c}12)", NUM0, "O que continua sendo faturado"),
    (15, "Custo da energia comprada (R$)", "={c}14*{c}7", MONEY, ""),
    (16, "Custo do Fio B sobre a energia compensada (R$)", "={c}12*ENTRADAS!$C$25", MONEY, "Lei 14.300: 60% da TUSD Fio B em 2026, 100% a partir de 2029"),
    (17, "COSIP / iluminacao publica (R$)", "=ENTRADAS!$C$26*12", MONEY, "R$ 21,11 por mes. A geracao nao afeta"),
    (18, "CONTA ANUAL COM SOLAR (R$)", "={c}15+{c}16+{c}17", MONEY, "O que voce continuara pagando por ano"),
    (19, "CONTA MEDIA MENSAL COM SOLAR (R$)", "={c}18/12", MONEY, "A prova de que nao existe conta de R$ 0"),
    (20, "Reducao da conta", "=IFERROR(1-{c}18/{c}8,0)", PCT, ""),
    (21, "Economia bruta anual (R$)", "={c}8-{c}18", MONEY, ""),
    (22, "Manutencao e limpeza (R$/ano)", "=ENTRADAS!$C$35", MONEY, "Custo recorrente real"),
    (23, "Provisao para troca do inversor (R$/ano)", "=ENTRADAS!$C$36", MONEY, "O inversor nao dura os 25 anos do modulo"),
    (24, "Reservado", "=0", MONEY, "Degradacao ja tratada no fator de geracao da aba PROJECAO"),
    (25, "Subtotal de custos recorrentes (R$)", "={c}22+{c}23", MONEY, ""),
    (26, "ECONOMIA LIQUIDA NO ANO 1 (R$)", "={c}21-{c}25", MONEY, "Base do payback e da projecao de 20 anos"),
]
for r, label, f, fmt, note in eco:
    c = ws.cell(row=r, column=2, value=label)
    c.font = BOLD if label.isupper() else BLACK
    for cl, o in zip(CL, OPT):
        if o["qty"] == "":
            ws[f"{cl}{r}"] = ""
            ws[f"{cl}{r}"].border = BOX
            continue
        frm(ws, f"{cl}{r}", f.replace("{c}", cl), fmt, BOLD if label.isupper() else BLACK)
        ws[f"{cl}{r}"].border = BOX
        if label.isupper():
            ws[f"{cl}{r}"].fill = WARN_F
    cn = ws.cell(row=r, column=NOTECOL, value=note)
    cn.font = SMALL
    cn.alignment = Alignment(wrap_text=True, vertical="top")
for cl, o in zip(CL, OPT):
    if o["qty"] != "":
        ws[f"{cl}26"].fill = OK_F
        ws[f"{cl}8"].fill = BAD_F

ws.cell(row=28, column=2, value="SIMULACAO SEM CARRO x COM CARRO (impacto isolado do BYD)").font = SUB
for c in range(2, NOTECOL + 1):
    ws.cell(row=28, column=c).fill = BAND
sim = [
    (29, "Consumo anual SEM o BYD (kWh)", "=BYD!$F$33*12", NUM0, "Media real da conta x 12"),
    (30, "Consumo anual COM o BYD (kWh)", "=BYD!$F$34*12", NUM0, "Cenario 2"),
    (31, "Acrescimo de consumo causado pelo BYD (kWh/ano)", "=C30-C29", NUM0, "Inclui as perdas de carregamento"),
    (32, "Acrescimo percentual sobre o consumo atual", "=IFERROR(C30/C29-1,0)", PCT, "Por isso a media da conta nao serve de base"),
    (33, "Conta anual SEM solar e SEM o BYD (R$)", "=C29*ENTRADAS!C18+ENTRADAS!C26*12", MONEY, "Sua situacao de hoje"),
    (34, "Conta anual SEM solar e COM o BYD (R$)", "=C30*ENTRADAS!C18+ENTRADAS!C26*12", MONEY, "O carro sozinho encarece a conta neste valor"),
    (35, "Custo anual so do BYD sem energia solar (R$)", "=C31*ENTRADAS!C18", MONEY, "Na tarifa cheia"),
    (36, "Custo anual so do BYD com energia solar (R$)", "=C31*ENTRADAS!C25", MONEY, "Compensado: paga-se apenas o Fio B"),
    (37, "ECONOMIA ANUAL SO NO CARREGAMENTO DO BYD (R$)", "=C35-C36", MONEY, "O melhor retorno marginal do projeto"),
    (38, "Numero de placas de 615 W dedicadas ao BYD", "=IFERROR(ROUNDUP(C31/GERACAO!I17*1000/615,0),0)", NUM0, "Placas que existem por causa do carro"),
]
for r, lab, f, fmt, note in sim:
    ws.cell(row=r, column=2, value=lab).font = BOLD if lab.isupper() else BLACK
    frm(ws, f"C{r}", f, fmt, BOLD if lab.isupper() else BLACK)
    if lab.isupper():
        ws[f"C{r}"].fill = OK_F
    ws.cell(row=r, column=NOTECOL, value=note).font = SMALL

# =====================================================================
# PROJECAO
# =====================================================================
ws = sheet("PROJECAO")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 7
ws.column_dimensions["C"].width = 12
CALC = [(cl, o) for cl, o in zip(CL, OPT) if o["qty"] != ""]
heads = ["Ano", "Fator geracao"]
grupos = []
col = 4
for cl, o in CALC:
    nm = o["L"].split("\n")[0]
    heads += [f"{nm}\nEconomia", f"{nm}\nAcumulado", f"{nm}\nFluxo"]
    grupos.append((get_column_letter(col), get_column_letter(col + 1), get_column_letter(col + 2),
                   f"ECONOMIA!{cl}$26", f"COMPARATIVO!{cl}${RW['total']}"))
    col += 3
for i in range(4, col):
    ws.column_dimensions[get_column_letter(i)].width = 14
title(ws, "PROJECAO DE 20 ANOS, ECONOMIA ACUMULADA E PAYBACK",
      "Economia do ano 1 corrigida por inflacao energetica e degradacao, menos O&M. 'Fluxo' = acumulado menos o investimento; fica positivo no ano do payback.")
hdr_row(ws, 4, heads, start=2)
ws.row_dimensions[4].height = 46
for i in range(20):
    r = 6 + i
    ano = i + 1
    ws.cell(row=r, column=2, value=ano).font = BLACK
    ws.cell(row=r, column=2).number_format = '0'
    frm(ws, f"C{r}", f"=(1-ENTRADAS!$C$33)*(1-ENTRADAS!$C$34)^({ano}-1)", PCT)
    for (ce, ca, cf, eco_ref, inv_ref) in grupos:
        frm(ws, f"{ce}{r}", f"={eco_ref}*C{r}/$C$6*(1+ENTRADAS!$C$27)^({ano}-1)", MONEY0)
        frm(ws, f"{ca}{r}", f"={ce}{r}" if i == 0 else f"={ca}{r-1}+{ce}{r}", MONEY0)
        frm(ws, f"{cf}{r}", f"={ca}{r}-{inv_ref}", MONEY0)
    for cc in range(2, col):
        ws.cell(row=r, column=cc).border = BOX
        if ano in (5, 10, 20):
            ws.cell(row=r, column=cc).fill = BAND
for i, t in enumerate([
    "Anos 5, 10 e 20 destacados. As colunas 'Fluxo' viram positivas no ano do payback ja corrigido por inflacao e degradacao.",
    "Nao inclui: troca do inversor no ano ~13 (provisionada em R$ 150/ano na aba ECONOMIA), reajustes de bandeira tarifaria,",
    "nem a mudanca da regra do Fio B (60% em 2026 -> 100% em 2029), que reduz a economia dos anos seguintes.",
    "Por isso o payback real tende a ser um pouco pior que o calculado, e o carregamento diurno do BYD fica cada vez mais vantajoso.",
]):
    c = ws.cell(row=27 + i, column=2, value=t)
    c.font = RED if i == 3 else SMALL

# =====================================================================
# SIMULACAO_MENSAL
# =====================================================================
ws = sheet("SIMULACAO_MENSAL")
for col, w in [("A", 3), ("B", 8), ("C", 15), ("D", 12), ("E", 14), ("F", 14), ("G", 13), ("H", 17), ("I", 42)]:
    ws.column_dimensions[col].width = w
title(ws, "SIMULACAO MES A MES - CONSUMO x GERACAO x BANCO DE CREDITOS",
      "Consumo da residencia = historico REAL da conta. Geracao com sazonalidade real (nao e media achatada).")
ws["B4"] = "Potencia simulada (kWp):"
ws["B4"].font = BOLD
inp(ws, "C4", 6.20, NUM2)
ws["D4"] = "<- edite para testar outra potencia (6,20 = SunWash / 7,38 = TecSolarSP / 5,35 = William)"
ws["D4"].font = SMALL
hdr_row(ws, 6, ["Mes", "Residencia (kWh)", "BYD (kWh)", "Consumo total (kWh)", "Geracao solar (kWh)",
                "Saldo do mes (kWh)", "Banco de creditos (kWh)", "Situacao"], start=2)
for i in range(12):
    r = 7 + i
    ws.cell(row=r, column=2, value=meses[i]).font = BLACK
    frm(ws, f"C{r}", f"=CONTA!C{38+i}", NUM0, GREEN)
    frm(ws, f"D{r}", "=BYD!$C$13", NUM0, GREEN)
    frm(ws, f"E{r}", f"=C{r}+D{r}", NUM0)
    frm(ws, f"F{r}", f"=$C$4*GERACAO!I{5+i}", NUM0)
    frm(ws, f"G{r}", f"=F{r}-E{r}", NUM0)
    frm(ws, f"H{r}", f"=MAX(0,G{r})" if i == 0 else f"=MAX(0,H{r-1}+G{r})", NUM0)
    frm(ws, f"I{r}", f'=IF(G{r}>=0,"Sobra - acumula credito","Falta - usa credito ou compra da rede")')
    for cc in range(2, 10):
        ws.cell(row=r, column=cc).border = BOX
ws.cell(row=19, column=2, value="ANO").font = BOLD
for cl in ["C", "D", "E", "F", "G"]:
    frm(ws, f"{cl}19", f"=SUM({cl}7:{cl}18)", NUM0, BOLD)
    ws[f"{cl}19"].fill = BAND
notas2 = [
    "COMO LER ESTA ABA:",
    "Os meses de maio a julho ficam negativos - inverno somado a telhado Sudeste, que perde mais justamente no inverno.",
    "Isso e normal e esperado. O excedente de outubro a marco forma o banco de creditos que cobre o inverno.",
    "O que importa e o SALDO ANUAL da coluna G, linha 19, ficar proximo de zero ou levemente positivo.",
    "Saldo anual muito positivo = voce pagou por placas que geram credito que vai expirar em 60 meses sem ser usado.",
    "Saldo anual negativo = voce vai continuar comprando energia da Enel todo mes, com tarifa cheia.",
    "",
    "OBSERVACAO SOBRE O CONSUMO DA RESIDENCIA: os valores vem do historico real, que ja mostra tendencia de alta de ~22%.",
    "Se o consumo continuar subindo, o deficit dos meses de inverno sera maior do que o mostrado aqui.",
]
r = 21
for t in notas2:
    c = ws.cell(row=r, column=2, value=t)
    c.font = SUB if t.endswith(":") else SMALL
    r += 1

# =====================================================================
# CHECKLIST
# =====================================================================
ws = sheet("CHECKLIST")
for col, w in [("A", 3), ("B", 5), ("C", 100), ("D", 14), ("E", 36)]:
    ws.column_dimensions[col].width = w
title(ws, "CHECKLIST - O QUE EXIGIR ANTES DE ASSINAR QUALQUER CONTRATO",
      "Itens em vermelho sao alertas ja identificados nos documentos recebidos. Preencha STATUS e a resposta de cada fornecedor.")
hdr_row(ws, 4, ["#", "PONTO A VERIFICAR", "STATUS", "RESPOSTA DO FORNECEDOR"], start=2)
chk = [
    ("A", "1) BLOQUEADORES - RESOLVER ANTES DE QUALQUER COMPRA"),
    ("R", "PADRAO MONOFASICO x INVERSOR DE 220 V. A conta indica fornecimento monofasico (127 V na Enel SP) e os quatro orcamentos usam inversor de 220 V, que precisa de 2 fases. Exigir de cada fornecedor: quem faz o pedido de alteracao de padrao na Enel, quanto custa, quanto tempo leva e se esta no preco."),
    ("R", "TIPO DA TELHA CONTRADITORIO. TecSolarSP cotou estrutura para FIBROCIMENTO e William cotou gancho para telha COLONIAL. Sao estruturas totalmente diferentes. Descobrir qual e a telha real - se um deles errou, nao visitou o telhado."),
    ("R", "GERACAO CALCULADA COMO SE O TELHADO FOSSE NORTE. TecSolarSP promete 1.430 kWh/kWp em telhado que ela mesma declara Sudeste. SunWash escreve 'calculo considerando norte do telhado'. Exigir simulacao com o azimute e a inclinacao REAIS, em PVsol, PVsyst ou Solergo, com relatorio em PDF."),
    ("R", "NENHUM ORCAMENTO CONSIDEROU O BYD. Informar a todos que havera carregamento diario do veiculo e pedir redimensionamento. A proposta do William (5,35 kWp) fica subdimensionada e o inversor Solis de 4 kW nao permite ampliar depois."),
    ("A", "1-sexies) SOBRE A AD BIOSOLAR (opcao J) - E A VERIFICACAO MAIS IMPORTANTE DE TODAS"),
    ("T", "PONTOS FORTES: unica que informou o MODELO EXATO do inversor (Huawei SUN2000-5KTL-L1); modulo JA Solar Tier 1 com 30 anos de garantia de eficiencia e 15 de defeito; inversor com 10 anos de garantia; CNPJ na proposta; vistoria tecnica, ART, licencas, gestao de obra e frete no escopo; unica a informar o peso por m2 (8,33 kg/m2), dado que importa em telhado ceramico."),
    ("R", "AGORA DA PARA FECHAR A CONTA QUE ESTAVA ABERTA: baixe a ficha tecnica do Huawei SUN2000-5KTL-L1 e a do modulo JA Solar de 615 W e compare a CORRENTE MAXIMA DE ENTRADA POR MPPT do inversor com a CORRENTE DE OPERACAO (Impp) do modulo. Modulos de 615 W nesse formato tem Impp na faixa de 16 A. Varias versoes da linha SUN2000-L1 aceitam de 12,5 a 13,5 A por MPPT. Se a corrente do modulo for maior que a do MPPT, o inversor NAO consegue extrair toda a potencia nas melhores horas do dia - perda permanente que nenhuma simulacao mostra."),
    ("R", "ESSA MESMA VERIFICACAO VALE PARA AS OPCOES A, B, C, D e I, que usam modulos de 605 a 620 W. Exigir de cada fornecedor o PROJETO DE STRING por escrito: quantos modulos em serie por MPPT, tensao minima e maxima nas temperaturas extremas, e a corrente por MPPT comparada ao limite do inversor. Fornecedor que nao souber responder isso nao fez projeto - fez lista de compras."),
    ("T", "AD BIOSOLAR CORRIGIU A PROPOSTA (#01237 de 12/08): subiu de 8 para 10 modulos e de 4,92 para 6,15 kWp, trocou o inversor do SUN2000-5KTL-L1 para o 6KTL-L1 de 6 kW, e passou a dimensionar com 650 kWh/mes em vez de 460. Com o valor negociado de R$ 16.200 ela sai a R$ 2,63/Wp - deixou de ser a mais cara e passou a ser uma das mais baratas, com o melhor equipamento do conjunto."),
    ("R", "O DESCONTO DE R$ 5.089 PRECISA ESTAR NO PAPEL. O PDF #01237 traz R$ 21.289,00; o valor de R$ 16.200,00 veio da negociacao. Sao 24% de desconto. Exigir proposta reemitida com o valor final, ou o desconto pode nao existir na assinatura."),
    ("R", "VERIFICACAO OBRIGATORIA NA AD BIOSOLAR: 10 modulos JA Solar de 615 W em um SUN2000-6KTL-L1 de 2 MPPTs dao 5 modulos por string. Confirmar na ficha tecnica do inversor a corrente maxima de entrada por MPPT e comparar com o Impp do modulo (~16 A). Se a do inversor for menor, perde-se geracao nas melhores horas. Pedir tambem a tensao da string nas temperaturas extremas."),
    ("R", "NUMEROS DE MARKETING DA AD BIOSOLAR: 'Economia total em 25 anos R$ 419.713,53', 'ROI 24,06 vezes' e 'TIR 35,81%' usam reajuste de energia de 10% ao ano composto por 25 anos. Ignore. Alem disso a 'conta COM sistema de R$ 75,31/mes' vale para 460 kWh/mes SEM o carro - com o BYD o meu modelo da cerca de R$ 259/mes nessa potencia."),
    ("R", "TRES VENDEDORES USARAM TRES CONSUMOS DIFERENTES: William usou 381 kWh/mes (a media real da conta), AD BioSolar usou 460 e Sfero usou 600. Nenhum deles explicou de onde veio o numero. Padronize: mande para TODOS o mesmo dado - 381 kWh/mes de historico MAIS 206 kWh/mes do BYD - e peca que refacam. So assim os orcamentos ficam comparaveis."),
    ("", ""),
    ("R", "CORRENTE DO MODULO x CORRENTE MAXIMA DO MPPT. Modulos de 605 a 615 W tem corrente de operacao na faixa de 14 a 17 A. Inversores residenciais da linha Huawei SUN2000-L1 costumam aceitar de 12,5 a 13,5 A por MPPT. Se a corrente do modulo ultrapassar a do inversor, perde-se geracao nas melhores horas. Exigir as duas fichas tecnicas e o calculo da string."),
    ("R", "STRING CURTA DE 6 MODULOS. A tensao de operacao fica na faixa de 200 a 225 V. Em dia quente a tensao cai e pode sair da faixa de potencia plena do MPPT. Exigir o projeto de string com temperatura minima e maxima."),
    ("R", "FALTA PROTECAO CA nas duas cotacoes da TecSolarSP: sem DPS CA e sem disjuntor dedicado do inversor. So o William incluiu DPS (3 pecas Clamper classe II). Item obrigatorio pela NBR 5410."),
    ("R", "FALTA CABO DE EQUIPOTENCIALIZACAO na cotacao Astronergy 605 W, e as garras de aterramento cairam de 6 para 3 jogos. Item de seguranca."),
    ("R", "ESTRUTURA NAO VALIDADA na cotacao Astronergy: a propria BelEnergy escreveu que 'as estruturas deste orcamento nao foram recomendadas'."),
    ("R", "PROPOSTA SEM CNPJ. A proposta do William nao traz razao social nem CNPJ. Sem isso nao existe garantia exigivel. Pedir o documento com identificacao da empresa."),
    ("R", "'INVERSOR HIBRIDO' NAO COMPROVADO. A peca publicitaria da TecSolarSP anuncia inversor HIBRIDO 6 kW, mas a cotacao traz 'INVHW-MO-220V-6KW'. Exigir o codigo completo do modelo e a confirmacao por escrito de que aceita bateria - hibrido e nao hibrido tem preco e funcao diferentes."),
    ("A", "1-bis) PERGUNTAS ESPECIFICAS DA PROPOSTA COM MICROINVERSORES (opcao G)"),
    ("T", "RESPONDIDO: os microinversores sao DEYE SUN-S225G4-EU-Q0, 2,25 kW monofasico. A Deye e um dos maiores fabricantes do mundo e tem rede de assistencia estabelecida no Brasil. Isso resolve a principal duvida que havia sobre esta proposta."),
    ("T", "CONSEQUENCIA 1 - CORRENTE: o sufixo Q0 indica 4 MPPTs por micro, e a linha G4 da Deye foi projetada para modulos de 600 a 700 W, com corrente por MPPT na casa de 20 A. Isso e folgado para os ~16 A de um modulo de 650 W. Ou seja: nesta proposta o risco de corte de corrente por MPPT NAO existe - diferente das propostas com inversor string e modulos de 605 a 620 W, onde a verificacao continua obrigatoria. CONFIRMAR na ficha tecnica do SUN-S225G4-EU-Q0."),
    ("T", "CONSEQUENCIA 2 - EXPANSAO: 4 micros x 4 MPPTs = 16 canais para 10 modulos = 6 CANAIS LIVRES. Da para chegar a 16 modulos (10,4 kWp) sem comprar inversor, com DC/AC de 1,16. E a unica proposta do conjunto com folga real de expansao - e voce tem um consumo subindo 39% ao ano."),
    ("R", "PERGUNTA QUE SOBROU: QUANTO CUSTA COM 3 MICROS EM VEZ DE 4? Tres micros dao 12 canais (10 usados, 2 livres) e 6,75 kW AC em vez de 9,0 kW. Duas vantagens: preco menor, e potencia homologada de 31 A em vez de 41 A, o que pode reduzir o custo do padrao de entrada. Vale so se voce NAO pretende passar de ~7,8 kWp no futuro."),
    ("R", "EXIGIR A GARANTIA DA DEYE POR ESCRITO. A Deye pratica 10 anos de fabrica nos microinversores, com extensao opcional, mas o Yuri nao declarou nada. Sem estar no papel, nao vale."),
    ("T", "RESPONDIDO: o modulo e ZNShine ZXNR-MD132-650, N-type, 650 W. A ZNShine e fabricante listada como Tier 1 pela Bloomberg NEF e atua desde os anos 1990. Com isso, os DOIS equipamentos desta proposta estao identificados e sao de fabricantes credenciados."),
    ("T", "ESCLARECIMENTO IMPORTANTE SOBRE 'TIER 1': o selo Tier 1 da Bloomberg NEF mede BANCABILIDADE, ou seja, que bancos financiaram projetos com aquele fabricante. NAO e certificado de qualidade nem de eficiencia, e nao substitui o registro INMETRO. Exigir o INMETRO do modelo ZXNR-MD132-650 do mesmo jeito."),
    ("T", "ESCLARECIMENTO SOBRE 'MONO': significa monocristalino, e nao monofacial. Se o modulo e bifacial ou nao precisa ser lido na ficha tecnica. Em telhado ceramico com modulo assentado proximo a telha o ganho bifacial e pequeno (1% a 2%), portanto isso quase nao muda a estimativa de geracao desta planilha."),
    ("R", "AGORA SO FALTA PAPEL NESTA PROPOSTA: razao social e CNPJ da SunWash, garantias de modulo/inversor/instalacao por escrito em anos, escopo item a item, esclarecimento dos R$ 1.800 da regiao, e o tipo de telha considerado. Nenhuma pendencia tecnica restou."),
    ("R", "PONTO A PESAR COM HONESTIDADE: a Deye domina o mercado brasileiro em inversor string e hibrido, mas a linha de MICROINVERSORES dela e mais recente no Brasil e tem menos historico de campo aqui do que os micros da Hoymiles ou APsystems. Nao e impedimento - e um dado a considerar, e reforca a exigencia da garantia por escrito."),
    ("R", "QUAL A MARCA DO MODULO DE 650 W? O orcamento nao informa. 650 W e formato grande (cerca de 2,6 m2 por placa): confirmar que cabe na agua Sudeste e que o peso por m2 e compativel com o telhado."),
    ("R", "QUANTOS MODULOS CADA MICROINVERSOR ACEITA? 4 micros de 2,25 kW dao 9 kW AC para 6,5 kWp de modulos (relacao 0,72). Se cada micro aceita 4 modulos, sao 16 canais para 10 modulos = 6 canais livres, o que e uma excelente folga de expansao. Se aceita menos, voce esta pagando por um micro desnecessario. Essa unica resposta muda a avaliacao da proposta."),
    ("R", "CONFIRMAR QUE O PARCELAMENTO EM 18x E REALMENTE SEM JUROS pelo valor a vista (18 x 932,25 = 16.780,54 confere). Pedir por escrito, e confirmar se o preco a vista nao teria desconto adicional."),
    ("R", "MICROINVERSOR EXIGE MANUTENCAO NO TELHADO. Perguntar como e feito o acesso, o prazo de troca em garantia e se o monitoramento e por modulo (deveria ser)."),
    ("T", "Vantagens tecnicas reais do microinversor no SEU caso: sombreamento nunca foi avaliado e o micro isola o problema por modulo; permite dividir os modulos em mais de uma agua do telhado (poderia recuperar parte dos 11% perdidos pela orientacao Sudeste); nao ha tensao alta de corrente continua no telhado; falha de um equipamento nao para o sistema todo."),
    ("", ""),
    ("A", "1-quater) PERGUNTAS DA PROPOSTA SFERO (opcao H)"),
    ("R", "MARCA COM BARRA NAO E MARCA. A proposta diz 'Longi/ DAH Solar' e 'Solis/ Growatt', ou seja, um OU outro. Exigir por escrito qual modelo exato sera entregue ANTES de assinar - a diferenca entre Longi e DAH, e entre Solis e Growatt, e real em preco e em suporte."),
    ("R", "5,90 kWp COBRE SO 92% DO CENARIO 3. Pedir a versao com 12 modulos e 3 microinversores (7,08 kWp, DC/AC 0,94), que levaria a cobertura para ~110%, e comparar o preco. Com 2 micros de 2,5 kW nao da para so acrescentar modulos: 12 x 590 W daria DC/AC de 1,42, acima do limite."),
    ("R", "60% DE ENTRADA E RISCO ALTO. Sao R$ 11.240 pagos antes de qualquer equipamento chegar ao telhado. Negociar para 30-40% de entrada, ou exigir garantia contratual e nota fiscal do material no ato do pagamento."),
    ("R", "A SFERO DIMENSIONOU COM 600 kWh/mes, e nao com os 381 kWh/mes reais da sua conta. Coincidentemente 600 e proximo do consumo COM o BYD (587). VOCE INFORMOU O CARRO PARA ELES? Se sim, e a unica empresa que considerou. Se nao, foi sorte - e vale confirmar de onde saiu esse numero."),
    ("T", "PONTOS FORTES REAIS DA SFERO, que nenhuma outra tem: modulo Longi ou DAH (Tier 1 com rede no Brasil) E microinversor Solis ou Growatt (marcas estabelecidas); seguro All Risk de 1 ano cobrindo danos, furto e responsabilidade civil; garantia de performance e resultado em contrato desde o 1o ano; string box com protecoes e monitoramento WiFi ja no escopo; relacao DC/AC de 1,18, a melhor de todas."),
    ("", ""),
    ("A", "1-quinquies) SOBRE A SUN COAST v2 (opcao I)"),
    ("T", "A Sun Coast CORRIGIU os dois defeitos apontados: a relacao DC/AC saiu de 1,56 para 0,91 e o inversor passou de 1 para 2 MPPTs, com 11 modulos em vez de 8. O preco marginal do acrescimo foi de R$ 2,10 por Wp, abaixo do preco medio dela - ou seja, foi um upgrade barato. Hoje tem o melhor R$/Wp de todos."),
    ("R", "MAS AS MARCAS CONTINUAM SENDO TCL (modulo) E AUXSOL (inversor), ambas com pouquissima rede de assistencia no Brasil. Garantia de '10 a 25 anos' no inversor so vale se existir empresa para honrar em 2040. Exigir INMETRO, o representante no Brasil com CNPJ e o prazo de atendimento."),
    ("R", "OS NUMEROS DE MARKETING SEGUEM INFLADOS: 'Retorno do investimento R$ 368.037,69' e 'Rendimento 40,69% ao ano' sobre um investimento de R$ 16.890. Meu calculo da cerca de R$ 219 mil de lucro em 20 anos, que ja e otimo. Ignore esses numeros e use a aba COMPARATIVO."),
    ("R", "AUXSOL DE 7,5 kW SIGNIFICA 7,5 kW DECLARADOS NA HOMOLOGACAO = 34 A em 220 V. Isso pode exigir padrao de entrada maior (e mais caro) do que 6,82 kWp precisaria. Perguntar se um inversor de 6 kW nao serviria, com preco menor."),
    ("", ""),
    ("A", "1-septies) AS DUAS CONTRADICOES DA PROPOSTA FORMAL DO YURI (proposta 6695731)"),
    ("R", "CONTRADICAO 1 - PADRAO DE ENTRADA. A pagina 5 lista como inclusos 'MAO DE OBRA + ENGENHARIA + HOMOLOGACAO + PADRAO DE ENTRADA + MATERIAL ELETRICO' e ainda 'TROCA DE DISJUNTOR DO PADRAO'. Mas a pagina 8, item 5, diz: 'E de responsabilidade do cliente todo o material e obra que tenha que ser realizada por motivos de adequacao de entrada'. Uma pagina inclui e a outra exclui. ISSO VALE CERCA DE R$ 2.500 - a planilha adotou como INCLUSO conforme a pagina 5, mas EXIJA a confirmacao por escrito. Se for excluido, some R$ 2.500 ao preco."),
    ("R", "CONTRADICAO 2 - MODULO. A secao de caracteristicas tecnicas diz marca ZSHINE, 650 Wp, quantidade 10, totalizando os 6,50 kWp declarados e 25 m2 de area. Mas a tabela de equipamentos lista 'MODULOS BIFACIAL 660W RISEN' com quantidade 16. Sao tres divergencias ao mesmo tempo: marca (ZShine x Risen), potencia (650 x 660 W) e quantidade (10 x 16). Dezesseis modulos de 660 W dariam 10,56 kWp e precisariam de uns 43 m2, nao 25. Provavelmente a tabela ficou com sobra de outro orcamento - mas o CONTRATO e o que vale, entao exija a proposta corrigida."),
    ("T", "SE POR ACASO FOREM MESMO 16 MODULOS pelo mesmo preco, seria 10,56 kWp a R$ 1,59/Wp - imbativel, e caberia exatamente nos 16 canais dos 4 micros Deye (DC/AC 1,17). Vale perguntar, mas nao contar com isso."),
    ("T", "GANHOS REAIS DESTA PROPOSTA FORMAL em relacao ao orcamento de WhatsApp: garantias agora declaradas (modulo 12 anos de defeito e 25 de eficiencia, inversor 12 anos), escopo de servicos listado (vistoria tecnica, projeto eletrico, ART, licencas junto a concessionaria, gestao e fiscalizacao de obra, frete e documentacao), e a troca do disjuntor do padrao no escopo."),
    ("R", "VALIDADE DE APENAS 5 DIAS: emitida em 12/08/2026 e valida ate 17/08/2026. Se for decidir, e nesta semana - ou peca revalidacao."),
    ("R", "AINDA FALTAM NESTA PROPOSTA: razao social e CNPJ da SunWash (o PDF traz apenas 'YURI' e um WhatsApp), o tipo de telha considerado, o registro INMETRO do modulo e do micro, e o esclarecimento sobre os R$ 1.800 do acrescimo regional."),
    ("T", "OBSERVACAO SOBRE OS NUMEROS FINANCEIROS DELES: usaram consumo de 502 kWh/mes e tarifa de R$ 0,882/kWh (R$ 442,67 / 502 kWh). Sua tarifa real e R$ 1,0215. A tabela de anos vai de 2023 a 2030, o que indica planilha-modelo nao atualizada. Ignore a analise financeira deles e use a aba COMPARATIVO."),
    ("", ""),
    ("A", "1-ter) DEMAIS BLOQUEADORES"),
    ("R", "SUN COAST: INVERSOR DE 3 kW PARA 4,68 kWp = RELACAO DC/AC DE 1,56. A faixa saudavel e 1,10 a 1,35. Nessa configuracao o inversor corta potencia (clipping) nas melhores horas do dia e a maioria dos fabricantes considera fora da garantia. Exigir do Alexandre o limite de sobrecarga do modelo AUXSOL cotado, por escrito, e a versao com inversor de 4 ou 5 kW."),
    ("R", "SUN COAST: INVERSOR COM 1 MPPT UNICO. Os 8 modulos ficam numa unica string. Isso exige que todos estejam na MESMA agua, com mesma orientacao e inclinacao, e faz com que sombra em um unico modulo derrube a producao da string inteira. Se o telhado tiver mais de uma agua util, essa configuracao nao serve."),
    ("R", "SUN COAST EXCLUI POR ESCRITO a adequacao do padrao de entrada, obras no telhado, eletrodutos e eletrocalhas, e o projeto estrutural com ART. E honesto declarar, mas o custo e seu: sao os R$ 2.500 (estimados) do padrao mais o que aparecer no telhado. Somar isso ao preco antes de comparar."),
    ("R", "MODULO RONMA SOLAR - MARCA A INVESTIGAR. Praticamente sem presenca no Brasil. Exigir registro INMETRO do modelo exato, ficha tecnica e o nome/CNPJ de quem honra os 25 anos de garantia de eficiencia no Brasil. Garantia de 25 anos so vale se existir empresa para honrar."),
    ("R", "TELHA: AGORA SAO DOIS CONTRA UM. William e Sun Coast cotaram TELHA COLONIAL; a TecSolarSP cotou FIBROCIMENTO. Se a telha for colonial, toda a estrutura das duas cotacoes da TecSolarSP esta errada (18 suportes pe em L de fibrocimento e hastes para madeira). Tirar uma foto do telhado resolve isso em 1 minuto."),
    ("R", "MODULO TCL 620 W - MARCA A INVESTIGAR. A TCL/TCL Zhonghuan e gigante em wafer de silicio, mas a marca de MODULO tem pouca rede de assistencia no Brasil. Exigir: registro INMETRO do modelo exato, ficha tecnica, termos de garantia de produto e de performance, e QUEM honra a garantia no Brasil (importador ou representante com CNPJ)."),
    ("R", "INVERSOR AUXSOL - MARCA DE BAIXA PENETRACAO. O inversor e o componente que mais falha e o que sera trocado antes dos modulos. Por R$ 1.000 a mais a SunWash oferece HUAWEI, com garantia de fabrica maior, assistencia estabelecida e monitoramento maduro. RECOMENDACAO: pagar a diferenca e ficar com o Huawei."),
    ("R", "INVERSOR DE 5 kW COM 6,20 kWp NAO DEIXA FOLGA DE EXPANSAO. O limite pratico de DC/AC de 1,35 da 6,75 kWp: nao cabe nem um modulo de 620 W a mais. Se a quilometragem do BYD crescer, sera preciso trocar o inversor e refazer a homologacao. Pedir a versao com inversor de 6 kW e comparar o preco."),
    ("R", "SUNWASH PROMETE 700 kWh/mes PARA 6,20 kWp = 1.355 kWh por kWp ao ano. No seu telhado Sudeste o calculo da ~1.264. Diferenca de ~7%. Menos exagerado que a TecSolarSP, mas ainda otimista - exigir a simulacao com o azimute real."),
    ("R", "O ACRESCIMO DE R$ 1.800 PELA REGIAO NAO FOI REPETIDO NO ORCAMENTO TCL. Confirmar por escrito se o preco de R$ 14.220,34 / R$ 15.220,34 e o valor FINAL entregue em Sao Paulo capital, ou se o acrescimo ainda incide."),
    ("R", "'ESTRUTURA, PROTECOES, CABOS E HOMOLOGACAO INCLUSOS' precisa virar lista item a item. Exigir: modelo da estrutura e para qual tipo de telha, DPS CC e CA, disjuntores, string box, bitola dos cabos, ART e quem protocola na Enel."),
    ("R", "ASSISTENCIA TECNICA A 340 km. A SunWash anuncia atendimento para Ribeirao Preto ate 80 km e cobra R$ 1.800 para atender a sua regiao. Perguntar o prazo de atendimento em caso de falha e se ha equipe propria em Sao Paulo."),
    ("R", "COTACOES VENCIDAS. As cotacoes Belenergy valem 3 dias e foram emitidas em 08/08 e 10/08/2026. Pedir revalidacao de preco antes de fechar."),
    ("", ""),
    ("A", "2) DOCUMENTOS A EXIGIR DE CADA FORNECEDOR"),
    ("T", "Proposta comercial formal com razao social, CNPJ, escopo completo, prazo de execucao e forma de pagamento."),
    ("T", "Relatorio de simulacao de geracao com azimute, inclinacao e sombreamento reais do telhado."),
    ("T", "Ficha tecnica e registro INMETRO do modulo e do inversor."),
    ("T", "Garantias por escrito: modulo (produto e performance), inversor, estrutura e mao de obra."),
    ("T", "Responsabilidade pela homologacao na Enel, com ART emitida por profissional com CREA ou CFT."),
    ("T", "Quem paga a adequacao do padrao de entrada e o aumento de carga, se a Enel exigir."),
    ("T", "Garantia contra infiltracao no telhado apos a perfuracao para fixacao da estrutura."),
    ("T", "Prazo de atendimento tecnico e distancia da equipe."),
    ("T", "Acesso do proprietario ao monitoramento (app) e se ha mensalidade."),
    ("", ""),
    ("A", "3) INSTALACAO ELETRICA E CARREGADOR DO BYD"),
    ("T", "Levantamento do padrao atual: numero de fases (hoje MONOFASICO), corrente do disjuntor geral e carga instalada."),
    ("T", "Calculo de demanda somando residencia + inversor + carregador do veiculo, para saber se e preciso aumento de carga na Enel."),
    ("T", "Circuito EXCLUSIVO para o carregador, com disjuntor dedicado e condutor dimensionado a 125% da corrente (NBR 5410, carga continua)."),
    ("T", "DPS no quadro do lado CA e DR/DDR conforme exigencia do fabricante do carregador."),
    ("T", "Medicao da resistencia de aterramento. Inversor e carregador dependem de aterramento adequado."),
    ("T", "Espaco livre no quadro de distribuicao para os novos disjuntores."),
    ("T", "Definir o horario de carregamento. Telhado Sudeste gera mais de manha: carregar entre 9h e 14h aproveita a geracao direta e nao paga Fio B."),
    ("", ""),
    ("A", "4) SO SE CONFIRMA COM INSPECAO PRESENCIAL - NAO ACEITE APENAS 'OK' NO WHATSAPP"),
    ("T", "Capacidade estrutural do telhado e das tercas para receber de 10 a 14 modulos."),
    ("T", "Area util real da agua Sudeste e quantas placas realmente cabem. NAO foram enviadas fotos do telhado: esta analise nao avaliou area disponivel, sombras, chamines, caixa d'agua nem obstaculos."),
    ("T", "Sombreamento ao longo do dia e do ano (arvores, caixa d'agua, muros, vizinhos)."),
    ("T", "Distancia real do telhado ao inversor e do inversor ao quadro, para definir a bitola dos cabos."),
    ("T", "Estado do padrao de entrada, do ramal e do medidor."),
]
r = 5
n = 0
for kind, txt in chk:
    if kind == "A":
        ws.cell(row=r, column=3, value=txt).font = SUB
        for cc in range(2, 6):
            ws.cell(row=r, column=cc).fill = BAND
    elif kind == "":
        pass
    else:
        n += 1
        ws.cell(row=r, column=2, value=n).font = BLACK
        c = ws.cell(row=r, column=3, value=txt)
        c.font = REDS if kind == "R" else BLACK
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if kind == "R":
            c.fill = BAD_F
        inp(ws, f"D{r}", "Pendente", None)
        ws.cell(row=r, column=5).border = BOX
        ws.cell(row=r, column=5).fill = YEL
    r += 1

# =====================================================================
# NOVOS_ORCAMENTOS
# =====================================================================
ws = sheet("NOVOS_ORCAMENTOS")
for col, w in [("A", 3), ("B", 44), ("C", 24), ("D", 22), ("E", 22), ("F", 52)]:
    ws.column_dimensions[col].width = w
title(ws, "FORMULARIO PARA OS PROXIMOS ORCAMENTOS",
      "Preencha uma coluna por vendedor. Depois copie os valores para as colunas correspondentes da aba COMPARATIVO.")
hdr_row(ws, 4, ["DADO A COLETAR", "Bruno G4", "Vendedor 6", "Vendedor 7", "POR QUE IMPORTA"], start=2)
campos = [
    ("Empresa / razao social / CNPJ", "Sem CNPJ nao existe garantia exigivel."),
    ("Contato", "Bruno G4: (11) 96861-7830"),
    ("Data da proposta e validade", "Cotacoes de material valem poucos dias."),
    ("O PRECO E MATERIAL OU INSTALADO?", "Pergunta mais importante da lista. Nunca compare os dois tipos."),
    ("Quantidade de modulos", "Base do dimensionamento."),
    ("Potencia por modulo (W)", "Quantidade x potencia / 1000 = kWp."),
    ("Fabricante, modelo e tecnologia do modulo", "N-type TOPCon degrada menos que PERC."),
    ("Potencia total DC (kWp)", "Confira se fecha com quantidade x potencia."),
    ("Fabricante e modelo exato do inversor", "Exigir o codigo, nao 'hibrido 6 kW'."),
    ("Potencia AC do inversor (kW)", "Limita a injecao e a expansao futura."),
    ("Corrente maxima por MPPT (A)", "Precisa ser maior que a corrente do modulo."),
    ("Corrente de operacao do modulo - Impp (A)", "Se for maior que a do MPPT, perde geracao."),
    ("Numero de MPPTs", "2 ou mais ajuda em telhado com aguas diferentes."),
    ("O inversor aceita bateria? Qual modelo?", "Se cobra 'hibrido', tem que entregar hibrido."),
    ("Geracao prometida (kWh/mes e kWh/ano)", "Compare com a coluna I da aba GERACAO."),
    ("A simulacao usou qual azimute e inclinacao?", "Se usou Norte, nao vale para o seu telhado Sudeste."),
    ("Software da simulacao e relatorio em PDF", "PVsol, PVsyst ou Solergo. Recusar planilha caseira."),
    ("Considerou o carregamento diario do BYD?", "Se nao perguntou do carro, dimensionou pela conta antiga."),
    ("Preco a vista (R$)", "Base do payback."),
    ("Acrescimos (frete, regiao, estrutura)", "SunWash cobra +R$ 1.800 pela regiao."),
    ("Inclui projeto, ART e homologacao na Enel?", "Sem isso o sistema nao compensa energia legalmente."),
    ("Inclui protecao CA: disjuntor dedicado + DPS?", "Faltou nas duas cotacoes da TecSolarSP."),
    ("Inclui aterramento e equipotencializacao?", "Faltou na cotacao Astronergy."),
    ("INCLUI A ADEQUACAO DO PADRAO MONO -> BIFASICO?", "Nenhum dos quatro incluiu. Pode custar milhares."),
    ("Inclui o carregador do veiculo e o circuito dele?", "Circuito dedicado com disjuntor e DR proprios."),
    ("Garantia do modulo: produto / performance (anos)", "Tipico: 12 a 15 anos produto, 25 a 30 performance."),
    ("Garantia do inversor (anos)", "Tipico: 5 a 10 anos, extensivel."),
    ("Garantia da instalacao e contra infiltracao", "Telhado perfurado e o maior risco pos-obra."),
    ("Prazo de execucao (dias)", "Inclui o tempo de homologacao na Enel."),
    ("Distancia da equipe de assistencia (km)", "Afeta o tempo de resposta em falha."),
    ("Tipo de telha que o orcamento considerou", "TecSolarSP disse fibrocimento, William disse colonial."),
]
r = 5
for lab, why in campos:
    c = ws.cell(row=r, column=2, value=lab)
    c.font = BOLD if lab.isupper() else BLACK
    c.border = BOX
    for cc in range(3, 6):
        inp(ws, f"{get_column_letter(cc)}{r}", "", None)
    cw = ws.cell(row=r, column=6, value=why)
    cw.font = SMALL
    cw.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1

del wb["Sheet"]
wb.active = 0
wb.save(OUT)
print("Salvo:", OUT)
