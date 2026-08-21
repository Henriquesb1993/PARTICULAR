# -*- coding: utf-8 -*-
"""Converte colunas de formulas repetidas em FORMULAS COMPARTILHADAS (sharedFormula).

O openpyxl grava a formula inteira em cada celula. Numa pasta com ~1 milhao de
formulas isso vira mais de 150 MB de XML, e o Excel leva minutos para abrir.
O proprio OOXML resolve isso: uma coluna de formulas identicas (que so mudam o
numero da linha) e gravada UMA vez, e as demais celulas apenas apontam para ela.

Este passo roda depois de salvar e nao altera nenhum resultado — apenas a forma
de armazenar. Excel, LibreOffice e Google Sheets leem sharedFormula nativamente.
"""
import re
import shutil
import sys
import zipfile

# referencia de celula: coluna (com ou sem $) + linha (com ou sem $)
REF = re.compile(r'(?<![A-Za-z0-9_$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])')
CELULA = re.compile(r'<c\b[^>]*/>|<c\b[^>]*>.*?</c>', re.S)
ATTR_R = re.compile(r'\br="([A-Z]+)(\d+)"')
FORMULA = re.compile(r'<f(?![a-zA-Z])[^>]*>(.*?)</f>', re.S)
FORMULA_VAZIA = re.compile(r'<f(?![a-zA-Z])[^>]*/>')


def normalizar(formula, linha):
    """Troca cada linha RELATIVA pelo seu deslocamento — duas fórmulas com a
    mesma forma normalizada são a mesma fórmula compartilhada."""
    def troca(m):
        cifrao_col, col, cifrao_lin, num = m.groups()
        if cifrao_lin:                      # linha absoluta: não desloca
            return m.group(0)
        return "{}{}[{}]".format(cifrao_col, col, int(num) - linha)
    return REF.sub(troca, formula)


def compactar_sheet(xml, si_inicial):
    partes, pos, celulas = [], 0, []
    for m in CELULA.finditer(xml):
        partes.append(xml[pos:m.start()])
        pos = m.end()
        celulas.append(len(partes))
        partes.append(m.group(0))
    partes.append(xml[pos:])

    info = []
    for idx in celulas:
        txt = partes[idx]
        mr = ATTR_R.search(txt)
        mf = FORMULA.search(txt)
        if not mr or not mf or FORMULA_VAZIA.search(txt):
            info.append(None)
            continue
        col, linha = mr.group(1), int(mr.group(2))
        corpo = mf.group(1)
        if "t=" in txt[mf.start():mf.start() + 40]:     # já compartilhada/array
            info.append(None)
            continue
        info.append((idx, col, linha, corpo, normalizar(corpo, linha), mf.span()))

    porcol = {}
    for k, dado in enumerate(info):
        if dado is not None:
            porcol.setdefault(dado[1], []).append(k)

    si = si_inicial
    trocas = 0
    for col, ks in porcol.items():
        ks.sort(key=lambda k: info[k][2])
        i = 0
        while i < len(ks):
            chave = info[ks[i]][4]
            j = i + 1
            while (j < len(ks) and info[ks[j]][4] == chave
                   and info[ks[j]][2] == info[ks[j - 1]][2] + 1):
                j += 1
            if j - i >= 4:                               # vale a pena compartilhar
                idx0, _, l0, corpo0, _, span0 = info[ks[i]]
                l1 = info[ks[j - 1]][2]
                txt = partes[idx0]
                novo_f = '<f t="shared" ref="{c}{a}:{c}{b}" si="{s}">{f}</f>'.format(
                    c=col, a=l0, b=l1, s=si, f=corpo0)
                partes[idx0] = txt[:span0[0]] + novo_f + txt[span0[1]:]
                for k in ks[i + 1:j]:
                    idxk, _, _, _, _, spank = info[k]
                    t = partes[idxk]
                    partes[idxk] = (t[:spank[0]]
                                    + '<f t="shared" si="{}"/>'.format(si)
                                    + t[spank[1]:])
                si += 1
                trocas += (j - i)
            i = j
    return "".join(partes), si, trocas


def compactar(caminho, destino=None):
    destino = destino or caminho
    tmp = destino + ".tmp"
    zin = zipfile.ZipFile(caminho)
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6)
    si = 0
    total = 0
    antes = depois = 0
    for item in zin.infolist():
        dados = zin.read(item.filename)
        if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
            xml = dados.decode("utf-8")
            antes += len(xml)
            xml, si, n = compactar_sheet(xml, si)
            depois += len(xml)
            total += n
            dados = xml.encode("utf-8")
        zout.writestr(item, dados)
    zout.close()
    zin.close()
    shutil.move(tmp, destino)
    print("compactadas {:,} fórmulas · XML das abas {:.1f} MB → {:.1f} MB".format(
        total, antes / 1e6, depois / 1e6).replace(",", "."))


if __name__ == "__main__":
    compactar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
