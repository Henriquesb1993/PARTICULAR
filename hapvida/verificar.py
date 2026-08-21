# -*- coding: utf-8 -*-
"""Confere a planilha gerada contra um recalculo independente feito em Python."""
import collections
import datetime as dt
import sys

from openpyxl import load_workbook

VT, VA = 151.73, 544.95
COMBO = {0: 0.0, 1: 239.93, 2: 446.62, 3: 579.83, 4: 663.57, 5: 846.86}
LIM = 25
EPOCH = dt.date(1899, 12, 30)
MAPA = {"TITULAR": ("TITULAR", False), "FILHO(A)": ("DEPENDENTE", True),
        "FILHO UNIVERS": ("DEPENDENTE", True), "ENTEADO": ("DEPENDENTE", True),
        "TUTELADO": ("DEPENDENTE", True), "MENOR POBRE": ("DEPENDENTE", True),
        "OUTROS": ("DEPENDENTE", True), "CONJUGE": ("DEPENDENTE", False),
        "CÔNJUGE": ("DEPENDENTE", False), "COMPANHEIRO(A)": ("DEPENDENTE", False),
        "AGREGADO": ("AGREGADO", False), "AGREGADO(A)": ("AGREGADO", False),
        "EX CONJUGE": ("NÃO ELEGÍVEL", False)}
AC = [("Á","A"),("À","A"),("Â","A"),("Ã","A"),("Ä","A"),("É","E"),("Ê","E"),("Í","I"),
      ("Ó","O"),("Ô","O"),("Õ","O"),("Ú","U"),("Ü","U"),("Ç","C"),(".",""),("-"," ")]


def ser(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        return EPOCH + dt.timedelta(days=int(v))
    return None


def norm(x):
    x = (str(x) if x is not None else "").upper().strip()
    for a, b in AC:
        x = x.replace(a, b)
    return " ".join(x.split())


def idade(d, base):
    return base.year - d.year - ((base.month, base.day) < (d.month, d.day)) if d else None


pn = lambda n: n.split(" ")[0] if n else ""
un = lambda n: n.split(" ")[-1] if n else ""


def casar(origem, destino):
    """Cascata de 4 níveis; devolve (indice, nivel) na ordem das linhas de destino."""
    i_dt, i_nm, i_pu = {}, {}, {}
    pref = []
    for j, x in enumerate(destino):
        i_dt.setdefault(x["chdt"], j)
        i_nm.setdefault(x["chnm"], j)
        i_pu.setdefault(x["chpu"], j)
        pref.append(x["chpu"])
    for x in origem:
        j = i_dt.get(x["chdt"])
        if j is None:
            j = i_nm.get(x["chnm"])
        if j is None:
            j = i_pu.get(x["chpu"])
        if j is None:
            p = "{}|{}|".format(x["re"], pn(x["nome"]))
            j = next((k for k, v in enumerate(pref) if v.startswith(p)), None)
        if j is None:
            x["idx"], x["corr"] = None, "— não encontrado"
            continue
        d = destino[j]
        x["idx"] = j
        if x["dt"] and x["dt"] == d["dt"]:
            x["corr"] = "FORTE (nascimento)"
        elif x["nome"] == d["nome"]:
            x["corr"] = "FORTE (nome completo)"
        elif x["chpu"] == d["chpu"]:
            x["corr"] = "PROVÁVEL (nome abreviado)"
        else:
            x["corr"] = "VERIFICAR (só 1º nome)"
        x["achou"] = "VERIFICAR" if x["corr"].startswith("VERIFICAR") else "SIM"


def carregar(origem, base, lim_h=None, lim_s=None):
    wb = load_workbook(origem, data_only=True, read_only=True)
    nomes = {n.strip().upper(): n for n in wb.sheetnames}
    wh = wb[next(n for k, n in nomes.items() if k.startswith("HAPVIDA"))]
    wsis = wb[next(n for k, n in nomes.items() if k.startswith("SISTEMA") or k.startswith("JB"))]

    HAP, SIS = [], []
    for row in wh.iter_rows(min_row=2, max_col=10, values_only=True):
        if row[1] in (None, ""):
            continue
        d = ser(row[4]); rel = str(row[7]).strip().upper()
        t, ag = MAPA.get(rel, ("DEPENDENTE", True))
        HAP.append(dict(re=str(row[1]).strip(), nome=norm(row[2]), dt=d, rel=rel, tipo=t,
                        idade=idade(d, base), regra=ag, cob=row[9] or 0,
                        contrato=(row[0] or "").strip(), nome_orig=(row[2] or "").strip()))
    for row in wsis.iter_rows(min_row=2, max_col=9, values_only=True):
        if row[0] in (None, ""):
            continue
        dep = str(row[5] or "").strip(); tit = str(row[1] or "").strip()
        d = ser(row[6]); rel = str(row[8]).strip().upper()
        t, ag = MAPA.get(rel, ("DEPENDENTE", True))
        SIS.append(dict(re=str(row[0]).strip(), nome=norm(dep or tit), dt=d, rel=rel, tipo=t,
                        idade=idade(d, base), regra=ag))
    wb.close()
    if lim_h:
        HAP = HAP[:lim_h]
    if lim_s:
        SIS = SIS[:lim_s]
    for i, x in enumerate(HAP + SIS):
        pass
    for lst, off in ((HAP, 2), (SIS, 4)):
        for i, x in enumerate(lst):
            x["chdt"] = "{}|{}".format(x["re"], x["dt"].strftime("%d%m%Y") if x["dt"] else "?%d" % (i + off))
            x["chnm"] = "{}|{}".format(x["re"], x["nome"])
            x["chpu"] = "{}|{}|{}".format(x["re"], pn(x["nome"]), un(x["nome"]))
            x["eleg"] = (x["tipo"] == "DEPENDENTE" and x["regra"] and (x["idade"] or 0) > LIM)
            x["achou"] = "NÃO"
    casar(HAP, SIS)
    casar(SIS, HAP)
    vis = set()
    for x in HAP:
        k = (x["chdt"], x["nome"])
        x["dup"] = k in vis
        vis.add(k)
    return HAP, SIS


def main(origem, xlsx, lim_h=None, lim_s=None):
    wv = load_workbook(xlsx, data_only=True)
    base = wv["PARÂMETROS"]["C6"].value
    base = base.date() if isinstance(base, dt.datetime) else base
    print("Data-base:", base)
    HAP, SIS = carregar(origem, base, lim_h, lim_s)

    for x in HAP:
        j = x["idx"]
        d = SIS[j] if j is not None else None
        x["div"] = "NÃO" if d is None else (
            "SIM" if (x["nome"] != d["nome"]
                      or (x["dt"] and d["dt"] and x["dt"] != d["dt"])
                      or x["tipo"] != d["tipo"]) else "NÃO")

    sre = collections.Counter(x["re"] for x in SIS)
    porRE = collections.defaultdict(list)
    for x in HAP:
        porRE[x["re"]].append(x)
    sporRE = collections.defaultdict(list)
    for x in SIS:
        sporRE[x["re"]].append(x)

    modelo = {}
    for x in HAP:                                  # 1 linha por RE: 1º titular nao duplicado
        if x["tipo"] == "TITULAR" and not x["dup"] and x["re"] not in modelo:
            rs = porRE[x["re"]]
            ss = sporRE.get(x["re"], [])
            nosis = x["re"] in sre
            dep = [p for p in rs if p["tipo"] == "DEPENDENTE" and not p["dup"]]
            agr = [p for p in rs if p["tipo"] == "AGREGADO" and not p["dup"]]
            modelo[x["re"]] = dict(
                nome=x["nome_orig"], contrato=x["contrato"], nosis="SIM" if nosis else "NÃO",
                ndep=len(dep), nagr=len(agr),
                ndep_sis=sum(1 for p in ss if p["tipo"] == "DEPENDENTE"),
                nagr_sis=sum(1 for p in ss if p["tipo"] == "AGREGADO"),
                depok=sum(1 for p in dep if p["achou"] == "SIM"),
                depsohap=(sum(1 for p in dep if p["achou"] == "NÃO") if nosis else 0),
                depinc=sum(1 for p in ss if p["tipo"] == "DEPENDENTE" and p["achou"] == "NÃO" and not p["eleg"]),
                agrinc=sum(1 for p in ss if (p["tipo"] == "AGREGADO" and p["achou"] == "NÃO")
                           or (p["tipo"] == "DEPENDENTE" and p["achou"] == "NÃO" and p["eleg"])),
                eleg=sum(1 for p in dep if p["eleg"]),
                cobrado=round(sum(float(p["cob"] or 0) for p in rs), 2),
                div=sum(1 for p in rs if p["div"] == "SIM" and not p["dup"]),
                verif=sum(1 for p in rs if p["achou"] == "VERIFICAR" and not p["dup"])
                      + sum(1 for p in ss if p["achou"] == "VERIFICAR"))
    for m, e in modelo.items():
        e["combo"] = COMBO[min(e["ndep"], 5)]
        e["vagr"] = e["nagr"] * VA
        e["total"] = round(VT + e["combo"] + e["vagr"], 2)

    ws = wv["TITULARES"]
    campos = [("nome", 2), ("contrato", 3), ("nosis", 5), ("ndep", 7), ("nagr", 8),
              ("ndep_sis", 9), ("nagr_sis", 10), ("depok", 11), ("depsohap", 12),
              ("depinc", 13), ("agrinc", 14), ("eleg", 15), ("combo", 17), ("vagr", 18),
              ("total", 19), ("cobrado", 20), ("div", 26), ("verif", 27)]
    erros, n, restante = [], 0, dict(modelo)
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v in (None, ""):
            continue
        n += 1
        m = str(v)
        if m not in modelo:
            erros.append("RE {} está em TITULARES mas não no modelo".format(m))
            continue
        restante.pop(m, None)
        e = modelo[m]
        for campo, col in campos:
            got = ws.cell(row=r, column=col).value
            exp = e[campo]
            if isinstance(exp, float):
                got = round(got or 0, 2); exp = round(exp, 2)
            if got != exp:
                erros.append("RE {} · {}: planilha={!r} modelo={!r}".format(m, campo, got, exp))
    for m in restante:
        erros.append("RE {} está no modelo e falta em TITULARES".format(m))

    print("Titulares conferidos: {} (modelo: {})".format(n, len(modelo)))
    print("Divergências: {}".format(len(erros)))
    for x in erros[:25]:
        print("  ✗", x)

    # ---- simulacao ----
    sim = wv["SIMULAÇÃO DE IMPACTO"]
    es = 0
    for r in range(4, sim.max_row + 1):
        v = sim.cell(row=r, column=1).value
        if v in (None, ""):
            continue
        e = modelo.get(str(v))
        if not e:
            continue
        novo = VT + COMBO[min(e["ndep"] + e["depinc"], 5)] + (e["nagr"] + e["agrinc"]) * VA
        for col, exp in ((8, novo), (9, novo - e["total"]), (10, (novo - e["total"]) * 12)):
            if round(sim.cell(row=r, column=col).value or 0, 2) != round(exp, 2):
                es += 1
                if es <= 5:
                    erros.append("RE {} simulação col{}: {!r} != {:.2f}".format(
                        v, col, sim.cell(row=r, column=col).value, exp))
    print("Erros na simulação: {}".format(es))

    rg = wv["RESUMO GERENCIAL"]
    tot = sum(e["total"] for e in modelo.values())
    print("\nCusto mensal total  planilha={!r}  modelo={:.2f}".format(rg["B9"].value, tot))
    print("Cartões: titulares={} dependentes={} agregados={}".format(
        rg["B6"].value, rg["E6"].value, rg["H6"].value))
    print("Impacto mensal planilha={!r}".format(rg["H9"].value))
    if rg["B9"].value is not None and abs(rg["B9"].value - tot) > 0.05:
        erros.append("custo total diverge")
    if rg["B6"].value != len(modelo):
        erros.append("cartão de titulares: {} != {}".format(rg["B6"].value, len(modelo)))
    print("\nTOTAL DE DIVERGÊNCIAS: {}".format(len(erros) + es))
    return len(erros) + es


if __name__ == "__main__":
    a = sys.argv[1:]
    lh = int(a[2]) if len(a) > 2 else None
    ls = int(a[3]) if len(a) > 3 else None
    sys.exit(1 if main(a[0], a[1], lh, ls) else 0)
