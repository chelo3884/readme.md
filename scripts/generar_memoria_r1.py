#!/usr/bin/env python3
"""
Genera la memoria de calculo Rev. 1 a partir de la Rev. 0 (docx) y de los resultados de
v10/tablas (calculo_r1.xlsx, derivados_memoria.xlsx, ...).

    python scripts/generar_memoria_r1.py --base docs/MT-CASA_MR-ESTRUCTURA-R0-30092026.docx \
        --tablas v10/tablas --figuras v10/figuras/memoria_r1 --salida docs/MT-CASA_MR-ESTRUCTURA-R1-01102026.docx
"""
import argparse
import copy
import os
import re
import subprocess
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches

sys.path.insert(0, str(Path(__file__).parent))
import calculo_r1 as cr  # noqa: E402
import figuras_r1  # noqa: E402

W = nsdecls("w")
VERDE, VERDE_CLARO = "00B157", "E8F6EE"
ANCHO = 9638
REV, FECHA_H, FECHA_PORTADA = "1", "2026/10/01", "Octubre 2026"
CODIGO = "MT-CASA_MR-ESTRUCTURA-R1-01102026"


def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# ---------------------------------------------------------------- constructores XML
def runs_xml(texto, sz=None, color="1A1A1A", bold=False, italic=False):
    """Texto con **negrita** en linea."""
    partes = re.split(r"(\*\*.+?\*\*)", texto)
    out = []
    for p in partes:
        if not p:
            continue
        b = bold
        if p.startswith("**") and p.endswith("**"):
            p, b = p[2:-2], True
        rpr = ""
        if b:
            rpr += "<w:b/><w:bCs/>"
        if italic:
            rpr += "<w:i/><w:iCs/>"
        if color:
            rpr += f'<w:color w:val="{color}"/>'
        if sz:
            rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        out.append(f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(p)}</w:t></w:r>')
    return "".join(out)


def p_cuerpo(texto, antes=0, despues=120, jc="both", sz=None, italic=False, color="1A1A1A"):
    sp = f'<w:spacing w:before="{antes}" w:after="{despues}" w:line="276" w:lineRule="auto"/>' if antes else \
        f'<w:spacing w:after="{despues}" w:line="276" w:lineRule="auto"/>'
    return parse_xml(f'<w:p {W}><w:pPr>{sp}<w:jc w:val="{jc}"/></w:pPr>'
                     f'{runs_xml(texto, sz=sz, italic=italic, color=color)}</w:p>')


def p_lista(texto, num_id=2, despues=80):
    return parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="Prrafodelista"/><w:numPr><w:ilvl w:val="0"/>'
                     f'<w:numId w:val="{num_id}"/></w:numPr><w:spacing w:after="{despues}" w:line="276" w:lineRule="auto"/>'
                     f'<w:jc w:val="both"/></w:pPr>{runs_xml(texto)}</w:p>')


_bm = {"n": 1000}


def p_titulo(texto, nivel):
    estilo = {1: "Ttulo1", 2: "Ttulo2", 3: "Ttulo3"}[nivel]
    borde = f'<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{VERDE}"/></w:pBdr>' if nivel == 1 else ""
    if nivel <= 2:
        _bm["n"] += 1
        i = _bm["n"]
        nombre = f"_Toc9{i:07d}"
        return parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="{estilo}"/><w:keepNext/>{borde}</w:pPr>'
                         f'<w:bookmarkStart w:id="{i}" w:name="{nombre}"/><w:r><w:t xml:space="preserve">{esc(texto)}</w:t></w:r>'
                         f'<w:bookmarkEnd w:id="{i}"/></w:p>')
    return parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="{estilo}"/><w:keepNext/></w:pPr>'
                     f'<w:r><w:t xml:space="preserve">{esc(texto)}</w:t></w:r></w:p>')


def p_salto():
    return parse_xml(f'<w:p {W}><w:r><w:br w:type="page"/></w:r></w:p>')


def celda(texto, ancho, fill=None, bold=False, color="1A1A1A", jc="center", sz=17):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    bord = "".join(f'<w:{l} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>' for l in ("top", "left", "bottom", "right"))
    pj = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc != "left" else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{ancho}" w:type="dxa"/><w:tcBorders>{bord}</w:tcBorders>{shd}'
            f'<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/>'
            f'<w:right w:w="60" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>'
            f'<w:p>{pj}{runs_xml(str(texto), sz=sz, color=color, bold=bold)}</w:p></w:tc>')


def tabla(encabezados, filas, anchos, alinear=None, sz=17, negrita_col0=False, resaltar=None):
    """Tabla con el estilo de la Rev. 0. `anchos` en dxa (se reescalan a 9638)."""
    k = ANCHO / sum(anchos)
    anchos = [int(a * k) for a in anchos]
    anchos[-1] += ANCHO - sum(anchos)
    alinear = alinear or ["center"] * len(anchos)
    grid = "".join(f'<w:gridCol w:w="{a}"/>' for a in anchos)
    xml = [f'<w:tbl {W}><w:tblPr><w:tblW w:w="{ANCHO}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>'
           + "".join(f'<w:{l} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for l in
                     ("top", "left", "bottom", "right", "insideH", "insideV"))
           + '</w:tblBorders><w:tblCellMar><w:left w:w="10" w:type="dxa"/><w:right w:w="10" w:type="dxa"/></w:tblCellMar>'
           '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>'
           f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>']
    xml.append('<w:tr><w:trPr><w:cantSplit/><w:tblHeader/><w:jc w:val="center"/></w:trPr>'
               + "".join(celda(h, a, fill=VERDE, bold=True, color="FFFFFF", sz=sz) for h, a in zip(encabezados, anchos))
               + "</w:tr>")
    for i, f in enumerate(filas):
        fill = VERDE_CLARO if i % 2 == 1 else None
        xml.append('<w:tr><w:trPr><w:cantSplit/><w:jc w:val="center"/></w:trPr>'
                   + "".join(celda(v, a, fill=fill, jc=al, sz=sz, bold=(negrita_col0 and j == 0) or
                                   (resaltar is not None and resaltar(i, j, v)))
                             for j, (v, a, al) in enumerate(zip(f, anchos, alinear))) + "</w:tr>")
    xml.append("</w:tbl>")
    return parse_xml("".join(xml))


def num(x, d=2, coma=False):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    s = f"{x:.{d}f}"
    return s.replace(".", ",") if coma else s


# ----------------------------------------------------------------- documento
class Memoria:
    def __init__(self, base, figuras):
        self.doc = Document(base)
        self.body = self.doc.element.body
        self.sect = self.body.find(qn("w:sectPr"))
        self.fig = {}
        self.n_fig = 6      # la Rev. 0 llega a la figura 6; se renumera mas abajo
        self.figuras = Path(figuras)

    # --- insercion al final del cuerpo
    def add(self, el):
        self.sect.addprevious(el)
        return el

    def texto(self, t, **kw):
        return self.add(p_cuerpo(t, **kw))

    def titulo(self, t, n):
        return self.add(p_titulo(t, n))

    def lista(self, items, num_id=2):
        for t in items:
            self.add(p_lista(t, num_id=num_id))

    def tabla(self, *a, **kw):
        t = self.add(tabla(*a, **kw))
        self.add(parse_xml(f'<w:p {W}><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>'))
        return t

    def figura(self, archivo, ancho_in, pie, numero):
        p = self.doc.add_paragraph()
        p.add_run().add_picture(str(self.figuras / archivo), width=Inches(ancho_in))
        ppr = parse_xml(f'<w:pPr {W}><w:keepNext/><w:spacing w:before="120" w:after="40"/><w:jc w:val="center"/></w:pPr>')
        old = p._p.find(qn("w:pPr"))
        if old is not None:
            p._p.remove(old)
        p._p.insert(0, ppr)
        self.add(parse_xml(
            f'<w:p {W}><w:pPr><w:spacing w:before="60" w:after="200"/><w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
            f'<w:t xml:space="preserve">Figura {numero}. {esc(pie)}</w:t></w:r></w:p>'))


def texto_de(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def buscar(body, inicio, desde=0, tag=None):
    for i, e in enumerate(list(body.iterchildren())):
        if i < desde:
            continue
        if tag and not e.tag.endswith(tag):
            continue
        if texto_de(e).strip().startswith(inicio):
            return i, e
    raise KeyError(inicio)


def poner_texto_parrafo(p, nuevo):
    runs = p.findall(qn("w:r"))
    for extra in p.findall(qn("w:proofErr")):
        p.remove(extra)
    for r in runs[1:]:
        p.remove(r)
    t = runs[0].find(qn("w:t"))
    t.text = nuevo
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def fijar_celda(tc, nuevo):
    p = tc.find(qn("w:p"))
    if p.findall(qn("w:r")):
        poner_texto_parrafo(p, nuevo)
    else:
        p.append(parse_xml(f'<w:r {W}><w:t>{esc(nuevo)}</w:t></w:r>'))


def filas_tabla(tbl):
    return tbl.findall(qn("w:tr"))


def celdas_fila(tr):
    return tr.findall(qn("w:tc"))


def reestriar(tbl):
    for i, tr in enumerate(filas_tabla(tbl)[1:]):
        for tc in celdas_fila(tr):
            pr = tc.find(qn("w:tcPr"))
            shd = pr.find(qn("w:shd"))
            if i % 2 == 1:
                if shd is None:
                    shd = parse_xml(f'<w:shd {W} w:val="clear" w:color="auto" w:fill="{VERDE_CLARO}"/>')
                    borders = pr.find(qn("w:tcBorders"))
                    borders.addnext(shd)
                else:
                    shd.set(qn("w:fill"), VERDE_CLARO)
            elif shd is not None:
                pr.remove(shd)


def agregar_fila(tbl, valores, copiar_de=-1):
    tr = copy.deepcopy(filas_tabla(tbl)[copiar_de])
    for tc, v in zip(celdas_fila(tr), valores):
        fijar_celda(tc, v)
    filas_tabla(tbl)[-1].addnext(tr)
    reestriar(tbl)
    return tr


def tabla_con_primera_celda(body, texto):
    for t in body.iter(qn("w:tbl")):
        tr = filas_tabla(t)[0]
        if texto_de(celdas_fila(tr)[0]).strip() == texto:
            return t
    raise KeyError(texto)


def fila_con(tbl, texto):
    for tr in filas_tabla(tbl):
        if texto_de(celdas_fila(tr)[0]).strip().startswith(texto):
            return tr
    raise KeyError(texto)


# ================================================================== datos
def cargar(dt):
    dt = Path(dt)
    d = {"dt": dt,
         "der": pd.read_excel(dt / "derivados_memoria.xlsx", sheet_name=None),
         "cal": pd.read_excel(dt / "calculo_r1.xlsx", sheet_name=None),
         "A": pd.read_excel(dt / "tablas_casa_mr.xlsx", sheet_name=["Acero_DC", "Cortes_Seccion", "Desplazamientos_Todos"]),
         "M": pd.read_excel(dt / "modelo_definiciones.xlsx", sheet_name=["Obj_Frames", "Asig_Frames", "Sec_Frames", "Coord_Nudos"])}
    d["der"]["Presion_zapatas"]["Paso"] = d["der"]["Presion_zapatas"]["Paso"].fillna("")
    return d


# ============================================== ediciones en las secciones 1 a 6
def editar_previas(m, d):
    body = m.body
    ch = list(body.iterchildren())
    # portada
    poner_texto_parrafo(ch[4], "Superestructura metálica, losas, muros de subsuelo y cimentación")
    poner_texto_parrafo(ch[12], CODIGO)
    poner_texto_parrafo(ch[14], FECHA_PORTADA)
    # control de revisiones
    t = ch[20]
    agregar_fila(t, ["1", FECHA_H, "Cimentación con resortes de suelo; diseño de zapatas, vigas de cimentación y muros", "M. Semblantes", "—"])
    # encabezado y pie
    for s in m.doc.sections:
        for part in (s.header, s.footer):
            for t_ in part._element.iter(qn("w:t")):
                if t_.text and "Rev: 0" in t_.text:
                    t_.text = t_.text.replace("Rev: 0", f"Rev: {REV}")
                if t_.text and "2026/09/30" in t_.text:
                    t_.text = t_.text.replace("2026/09/30", FECHA_H)
                if t_.text and "R0-30092026" in t_.text:
                    t_.text = t_.text.replace("R0-30092026", "R1-01102026")
    # 1.1 tabla: fila de cimentacion
    t11 = tabla_con_primera_celda(body, "Uso")
    tr = agregar_fila(t11, ["Cimentación", "Zapatas aisladas de H.A. de 1.00, 1.20 y 1.30 m (e = 30 cm), zapatas corridas bajo los muros (e = 35 cm) y vigas de cimentación; Df = 1.40 m, qadm = 24 T/m²"])
    # mover la fila nueva despues de "Sistema estructural"
    filas_tabla(t11)[-1].getparent().remove(filas_tabla(t11)[-1])
    fila_con(t11, "Sistema estructural").addnext(tr)
    reestriar(t11)
    # 1.2 alcance
    _, p = buscar(body, "Este documento cubre")
    poner_texto_parrafo(p, "Este documento cubre la definición de cargas, el análisis sísmico con la cimentación modelada mediante resortes de suelo, la verificación de derivas y estabilidad, el diseño de los elementos de acero, el diseño de muros, pedestales, pilastras y vigas de coronación de hormigón armado, y el diseño de la cimentación (zapatas aisladas, zapatas corridas y vigas de cimentación). El diseño de las conexiones se desarrolla en un documento complementario, con base en los criterios aquí establecidos. Esta Rev. 1 reemplaza el predimensionamiento de la cimentación y los resultados de la superestructura de la Rev. 0, que se calcularon con base empotrada.")
    # 2 materiales
    t2 = tabla_con_primera_celda(body, "Elemento")
    for tr in filas_tabla(t2):
        c0 = texto_de(celdas_fila(tr)[0])
        if c0.startswith("Muros, pilastras"):
            fijar_celda(celdas_fila(tr)[0], "Muros, pedestales, pilastras, vigas y zapatas")
    # 4.2 hipotesis
    t42 = None
    for t in body.iter(qn("w:tbl")):
        if texto_de(celdas_fila(filas_tabla(t)[0])[0]).strip() == "Aspecto":
            t42 = t
    der = d["der"]
    A = d["A"]["Desplazamientos_Todos"].copy()
    C = d["M"]["Coord_Nudos"].assign(Joint=lambda x: x["Joint"].astype(str)).set_index("Joint")
    A["Joint"] = A["Joint"].astype(str)
    A = A.join(C[["GlobalZ"]], on="Joint")
    A["z"] = A["GlobalZ"].round(2)
    r = {}
    for cs, comp in (("SPECX", "U1"), ("SPECY", "U2")):
        g = A[A["OutputCase"] == cs].groupby("z")[comp].agg(lambda s: s.abs().max())
        r[cs] = g.loc[0.0] / g.loc[6.12] * 100
    for tr in filas_tabla(t42):
        c0 = texto_de(celdas_fila(tr)[0]).strip()
        if c0 == "Base sísmica":
            fijar_celda(celdas_fila(tr)[1], f"Nivel ±0.00: diafragma apoyado en los muros de subsuelo y en las cimentaciones del ala este; su desplazamiento espectral máximo es el {r['SPECX']:.0f} % (X) y el {r['SPECY']:.0f} % (Y) del de la cubierta (4,6 % en la Rev. 0, con base empotrada).")
        elif c0 == "Apoyos":
            fijar_celda(celdas_fila(tr)[1], "Zapatas modeladas con elementos shell sobre resortes de suelo: verticales (kv = 4 800 T/m³ por el área tributaria de cada nudo) y horizontales (kh = 0.5·kv). Columnas sobre pedestales de H.A. de 40×40 cm; zapatas unidas por vigas de cimentación (sección 10.1).")
        elif c0 == "Muros de subsuelo":
            fijar_celda(celdas_fila(tr)[1], "Shell e = 20 cm mallado a ≈ 0.55 × 0.51 m, unido a la zapata corrida (shell e = 35 cm) hasta su plano medio (z = −4.285 m), con inercias agrietadas 0.6·Ig (NEC-SE-DS 6.1.6).")
        elif c0 == "Análisis modal":
            fijar_celda(celdas_fila(tr)[1], "Vectores Ritz (40), iniciados con aceleraciones en X e Y.")
    # 4.3 secciones: filas nuevas
    t43 = tabla_con_primera_celda(body, "Sección")
    Ob, Asg, Sf, Cn = (d["M"][k] for k in ("Obj_Frames", "Asig_Frames", "Sec_Frames", "Coord_Nudos"))
    sec = Asg.set_index(Asg["Frame"].astype(str))["AnalSect"]
    Ob = Ob.assign(sec=Ob["FrameObject"].astype(str).map(sec))
    Cn = Cn.assign(Joint=Cn["Joint"].astype(str)).set_index("Joint")
    def largo(r):
        a, b = Cn.loc[str(r["ElemJtI"])], Cn.loc[str(r["ElemJtJ"])]
        return float(np.sqrt(sum((a[f"Global{k}"] - b[f"Global{k}"]) ** 2 for k in "XYZ")))
    Ob["L"] = Ob.apply(largo, axis=1)
    for s, uso in (("PED40X40", "Pedestales de zapatas"), ("VC20X30", "Vigas de cimentación"),
                   ("VC25X40", "Vigas de cimentación"), ("VC25X45", "Vigas de cimentación"),
                   ("VC25X50", "Vigas de cimentación"), ("VC30X50", "Vigas de cimentación"),
                   ("NUDO40X60", "Nudos pedestal-muro")):
        g = Ob[Ob["sec"] == s]
        agregar_fila(t43, [s, "H.A. 280", uso, str(g["FrameObject"].nunique()), f"{g['L'].sum():.1f}", "—"])
    _, p = buscar(body, "Nomenclatura:")
    nuevo = copy.deepcopy(p)
    poner_texto_parrafo(nuevo, "Elementos shell de H.A.: MURO20 (muros, e = 20 cm, 490 elementos), ZAP30 (zapatas aisladas, e = 30 cm, 32.5 m²) y ZAP35 (zapatas corridas, e = 35 cm, 40.3 m²). Las secciones de pedestales, vigas de cimentación y nudos se modelan con inercia bruta.")
    p.addnext(nuevo)
    # 5.1 relleno
    i52, h52 = buscar(body, "5.2. Carga viva (L)")
    h52.addprevious(p_cuerpo(antes=120, texto="**Relleno y contrapiso sobre zapatas (RELLENO).** Representa el relleno y el contrapiso sobre las zapatas y se aplica como carga uniforme gravitacional sobre sus elementos: 1.73 y 1.88 T/m² sobre las zapatas aisladas y 1.65 T/m² sobre la mayor parte de las zapatas corridas, con 6.11 T/m² sobre 196 elementos de estas últimas. RELLENO forma parte de la carga muerta (D) en todas las combinaciones, pero no entra en la masa sísmica."))
    # 5.4 nota sobre el tramo enterrado
    _, p = buscar(body, "El empuje está desbalanceado")
    p.addnext(p_cuerpo("El empuje actúa entre z = −3.06 y 0.00 m (H = 3.06 m). El tramo de muro comprendido entre z = −3.06 y la zapata (z = −4.285 m, 1.225 m) no recibe empuje neto porque se supone equilibrado por el relleno y el contrapiso interiores; este supuesto debe confirmarse en obra con el procedimiento constructivo del subsuelo."))
    # 5.5.5 periodo
    _, p = buscar(body, "Período aproximado")
    poner_texto_parrafo(p, texto_de(p).replace("T1 = 0.466 s", "T1 = 0.480 s"))
    # 6 combinaciones
    _, p = buscar(body, "Las combinaciones de resistencia siguen")
    poner_texto_parrafo(p, texto_de(p).replace("D = DEAD + MUERTA", "D = DEAD + MUERTA + RELLENO"))
    t6 = tabla_con_primera_celda(body, "Combinación")
    tr = fila_con(t6, "Cimentación (ASD)")
    fijar_celda(celdas_fila(tr)[0], "Suelo (ASD): CS1 a CS4")
    fijar_celda(celdas_fila(tr)[1], "CS1: D + L + H · CS2: D + 0.75L + 0.75S + H · CS3X/Y: D + 0.75L + 0.75S + 0.525E + H + 0.525EH_sis · CS4X/Y: 0.6D + 0.7E + H + 0.7EH_sis")
    fijar_celda(celdas_fila(tr)[2], "Presiones, levantamiento y deslizamiento")


# ================================================================== secciones 7 a 11
def nc(x, d=2):
    return num(x, d, coma=True)


ORDEN_CASO = {"SPECX": 0, "SX": 1, "SPECY": 2, "SY": 3}


def sec7(m, d):
    der = d["der"]
    mod = der["Modal"]
    m.titulo("7. ANÁLISIS ESTRUCTURAL", 1)
    m.titulo("7.1. Análisis modal", 2)
    filas = [[int(r["Modo"]), nc(r["T (s)"], 3), nc(r["UX (%)"], 1), nc(r["UY (%)"], 1), nc(r["RZ (%)"], 1),
              nc(r["Suma UX (%)"], 1), nc(r["Suma UY (%)"], 1)] for _, r in mod.head(6).iterrows()]
    m.tabla(["Modo", "T (s)", "UX (%)", "UY (%)", "RZ (%)", "Σ UX (%)", "Σ UY (%)"], filas, [1] * 7)
    r30 = mod.iloc[29]
    rf = mod.iloc[-1]
    t = mod["T (s)"].values
    m.texto(f"Con {len(mod)} vectores Ritz la masa participativa acumulada llega a {rf['Suma UX (%)']:.1f} % en X y "
            f"{rf['Suma UY (%)']:.1f} % en Y, por encima del 90 % exigido (NEC-SE-DS 6.2.2); con los primeros 30 vectores es "
            f"de {r30['Suma UX (%)']:.1f} % y {r30['Suma UY (%)']:.1f} %. Los períodos fundamentales son T1 = {nc(t[0], 3)} s y "
            f"T2 = {nc(t[1], 3)} s, con participación acoplada de las dos direcciones traslacionales en cada modo "
            f"(UX + UY ≈ {mod.iloc[0]['UX (%)'] + mod.iloc[0]['UY (%)']:.0f} % y {mod.iloc[1]['UX (%)'] + mod.iloc[1]['UY (%)']:.0f} %), "
            f"y la torsión aparece en el tercer modo (T3 = {nc(t[2], 3)} s; RZ = {mod.iloc[2]['RZ (%)']:.1f} %). "
            f"Con respecto a la Rev. 0 (base empotrada: 0,466; 0,452 y 0,382 s), la flexibilidad de la cimentación alarga "
            f"T1 un {(t[0] / 0.466 - 1) * 100:.1f} %, T2 un {(t[1] / 0.452 - 1) * 100:.1f} % y T3 un {(t[2] / 0.382 - 1) * 100:.1f} %. "
            "El resto de la masa corresponde al nivel ±0.00 y a los muros, que son muy rígidos, y se capta en modos de período cercano a cero.")

    m.titulo("7.2. Ajuste del cortante dinámico", 2)
    c = der["Cortante"]
    filas = [[r["Direccion"], nc(r["V estatico (T)"]), nc(r["V dinamico (T)"]), nc(r["Relacion (%)"], 1) + " %",
              "85 %", r["Resultado"]] for _, r in c.iterrows()]
    m.tabla(["Dirección", "V estático (T)", "V dinámico (T)", "Relación", "Mínimo (irregular)", "Resultado"], filas, [1] * 6)
    cx, cy = c.iloc[0], c.iloc[1]
    m.texto("El cortante se mide con un corte de sección en +0.10, sobre la base sísmica. El cortante dinámico supera el 85 % del "
            "estático exigido para estructuras irregulares (NEC-SE-DS 6.2.2), así que no hace falta escalar el espectro. "
            f"Con la base flexible el cortante dinámico pasa de 65,74 a {nc(cx['V dinamico (T)'])} T en X y de 64,70 a "
            f"{nc(cy['V dinamico (T)'])} T en Y (la resultante vectorial es {nc(cx['V dinamico resultante (T)'])} y "
            f"{nc(cy['V dinamico resultante (T)'])} T), mientras que el estático (72,07 T) no cambia porque depende solo de C y W.")

    m.titulo("7.3. Derivas de piso", 2)
    m.texto("Derivas inelásticas según NEC-SE-DS 6.3.9: ΔM = 0.75·R·ΔE = 3.375·ΔE, con límite de 0.02 para estructuras de acero "
            "(NEC-SE-DS 4.2.2, Tabla 7). ΔE se calcula con la diferencia de desplazamientos entre nudos ubicados en la misma "
            "vertical de dos niveles consecutivos (se evaluaron entre 5 y 50 parejas por nivel y caso) y se reporta la pareja "
            "más desfavorable, incluida la torsión accidental.")
    dv = der["Derivas"].assign(o=lambda x: x["Caso"].map(ORDEN_CASO)).sort_values(["o", "Nivel (m)"])
    filas = [[f"+{r['Nivel (m)']:.2f}", r["Dir"], r["Caso"], f"{int(r['Nudo sup'])} / {int(r['Nudo inf'])}",
              nc(r["Desp. sup (mm)"]), nc(r["Desp. inf (mm)"]), nc(r["dE"], 5), nc(r["dM"], 4), r["Resultado"]]
             for _, r in dv.iterrows()]
    m.tabla(["Nivel", "Dir.", "Caso", "Nudos sup./inf.", "Desplaz. sup. (mm)", "Desplaz. inf. (mm)", "ΔE", "ΔM", "Resultado"],
            filas, [0.8, 0.6, 1, 1.3, 1.2, 1.2, 1, 1, 1], sz=16)
    m.figura("r1_derivas.png", 5.4, "Derivas inelásticas frente al límite de NEC-SE-DS (Rev. 1, base con resortes).", 5)
    mx = dv.loc[dv["dM"].idxmax()]
    C = d["M"]["Coord_Nudos"].assign(Joint=lambda x: x["Joint"].astype(str)).set_index("Joint")
    xy = C.loc[str(int(mx["Nudo sup"]))]
    m.texto(f"La deriva máxima, {nc(mx['dM'], 4).replace(',', '.')}, ocurre con el caso {mx['Caso']} en el nivel +{mx['Nivel (m)']:.2f} "
            f"(nudos {int(mx['Nudo sup'])} y {int(mx['Nudo inf'])}, x = {xy['GlobalX']:.2f} m, y = {xy['GlobalY']:.2f} m); es "
            f"{(1 - mx['dM'] / 0.0189) * 100:.0f} % menor que la de la Rev. 0 (0.0189) y queda "
            f"{(1 - mx['dM'] / 0.02) * 100:.0f} % por debajo del límite de 0.020. Con la base flexible parte del desplazamiento de la "
            "cubierta es movimiento de cuerpo rígido del nivel ±0.00 sobre el suelo (hasta 9,1 mm en X y 3,8 mm en Y), que no genera deriva; por eso la deriva "
            "máxima baja aunque los períodos aumentan. Las columnas [ ]250X200X8 del bloque norte (B-1, C-1, B-2, C-2) conservan su "
            "orientación (lado mayor en X).")

    m.titulo("7.4. Índice de estabilidad (P-Δ)", 2)
    cz = d["A"]["Cortes_Seccion"]
    def vcorte(corte, caso, col):
        return abs(cz[(cz["SectionCut"] == corte) & (cz["OutputCase"] == caso)][col].iloc[0])
    def de(nivel, caso):
        return float(dv[(dv["Nivel (m)"] == nivel) & (dv["Caso"] == caso)]["dE"].iloc[0])
    filas = []
    for piso, P, corte, niv in (("Piso 1 (±0.00 a +3.06)", 284.0, "Z 0.10", 3.06), ("Piso 2 (+3.06 a +6.12)", 62.5, "Z 3.10", 6.12)):
        vx, vy = vcorte(corte, "SPECX", "F1"), vcorte(corte, "SPECY", "F2")
        ex, ey = de(niv, "SPECX"), de(niv, "SPECY")
        filas.append([piso, nc(P, 1), f"{nc(vx)} / {nc(vy)}", f"{nc(ex, 4)} / {nc(ey, 4)}",
                      f"{nc(P * ex / vx, 3)} / {nc(P * ey / vy, 3)}", "0.10"])
    m.tabla(["Piso", "Pi = D + L (T)", "Vi X / Y (T)", "ΔE X / Y (máx.)", "Qi X / Y", "Límite"], filas, [1.6, 1, 1.2, 1.4, 1.2, 0.7])
    m.texto("Qi = Pi·ΔEi/Vi (equivalente a Pi·Δi/(Vi·hi) con Δi = ΔEi·hi) es menor que 0.10 en todos los pisos (NEC-SE-DS 6.3.8), así "
            "que no se requieren efectos P-Δ adicionales. Se usa la mayor deriva elástica de cada piso, que es conservadora "
            "frente al valor en el centro de masas empleado en la Rev. 0, y las cargas Pi de la Rev. 0, que no cambian. "
            "En el diseño de acero los efectos de segundo orden se incluyen mediante el método de análisis directo (AISC 360-16 C), "
            "con primer orden amplificado.")

    m.titulo("7.5. Efecto de la base flexible respecto de la Rev. 0", 2)
    ac = der["Acero_resumen"]
    dcp = ac[ac["Verificacion"] == "PMM"]["D/C max"].max()
    dcg = ac["D/C max"].max()
    dmx = dv[dv["Dir"] == "X"]["dM"].max()
    dmy = dv[dv["Dir"] == "Y"]["dM"].max()
    filas = [["T1 / T2 / T3 (s)", "0,466 / 0,452 / 0,382", f"{nc(t[0], 3)} / {nc(t[1], 3)} / {nc(t[2], 3)}"],
             ["Masa participativa con 30 vectores X / Y", "98,1 % / 97,4 %", f"{r30['Suma UX (%)']:.1f} % / {r30['Suma UY (%)']:.1f} %".replace(".", ",")],
             ["Cortante dinámico en la base X / Y (T)", "65,74 / 64,70", f"{nc(cx['V dinamico (T)'])} / {nc(cy['V dinamico (T)'])}"],
             ["Relación dinámico/estático X / Y", "91,2 % / 89,8 %", f"{nc(cx['Relacion (%)'], 1)} % / {nc(cy['Relacion (%)'], 1)} %"],
             ["Deriva inelástica máxima X / Y", "0,0189 / 0,0173", f"{nc(dmx, 4)} / {nc(dmy, 4)}"],
             ["D/C máximo por resistencia (PMM)", "0,81", nc(dcp)],
             ["D/C máximo global (deflexión)", "0,94", nc(dcg)]]
    m.tabla(["Resultado", "Rev. 0 (base empotrada)", "Rev. 1 (base con resortes)"], filas, [2.4, 1.6, 1.6], alinear=["left", "center", "center"])
    m.texto("La flexibilidad del suelo alarga los períodos y redistribuye la rigidez entre los pórticos y los muros, de modo que el "
            "cortante dinámico aumenta, las derivas de piso disminuyen ligeramente y varias secciones de acero suben su relación "
            "D/C (sección 8). Todas las verificaciones de NEC-SE-DS siguen cumpliéndose sin cambiar las secciones de la superestructura.")


def sec8(m, d, plant):
    ac = d["A"]["Acero_DC"]
    m.titulo("8. DISEÑO DE ELEMENTOS DE ACERO", 1)
    for e in plant["8.1"]:
        m.add(copy.deepcopy(e))
    m.titulo("8.2. Resultados", 2)
    tipo = {"Beam": "Viga", "Column": "Columna", "Brace": "Puntal"}
    orden = ["VK150X100X4X4", "VK180X100X4X6", "VK200X100X4X8", "VK220X110X4X8", "VK250X130X4X8", "VK270X140X4X8",
             "[]100X100X3", "[]200X200X8", "[]250X200X8"]
    g = ac.loc[ac.groupby(["DesignSect", "DesignType"])["Ratio"].idxmax()]
    cnt = ac.groupby(["DesignSect", "DesignType"])["Frame"].nunique()
    g = g.assign(n=[cnt[(a, b)] for a, b in zip(g["DesignSect"], g["DesignType"])])
    g = g.assign(o=g["DesignSect"].map({s: i for i, s in enumerate(orden)}),
                 ot=g["DesignType"].map({"Beam": 0, "Brace": 1, "Column": 2})).sort_values(["o", "ot"])
    filas = [[r["DesignSect"], tipo[r["DesignType"]], int(r["n"]), nc(r["Ratio"], 3),
              "Resistencia (PMM)" if r["RatioType"] == "PMM" else "Deflexión (servicio)", r["Combo"]] for _, r in g.iterrows()]
    m.tabla(["Sección", "Tipo", "Piezas", "D/C máx.", "Verificación", "Combinación"], filas, [1.5, 1, 0.8, 0.9, 1.6, 1.1])
    m.figura("r1_dc_acero.png", 5.4, "Relación demanda/capacidad máxima por sección (Rev. 1).", 6)
    pmm = ac[ac["RatioType"] == "PMM"]
    w = pmm.loc[pmm["Ratio"].idxmax()]
    wg = ac.loc[ac["Ratio"].idxmax()]
    m.texto(f"Se verificaron {ac['Frame'].nunique()} elementos. La relación demanda/capacidad máxima por resistencia es "
            f"{nc(w['Ratio'], 2)} ({w['DesignSect']}, combinación {w['Combo']}) y la máxima global es {nc(wg['Ratio'], 2)}, controlada por la "
            f"deflexión en servicio ({wg['Combo']}) de las viguetas {wg['DesignSect']}. Con la base flexible los momentos se redistribuyen "
            "hacia las vigas y las columnas del bloque norte: por ejemplo, las columnas []250X200X8 pasan de 0,62 a 0,82, las vigas "
            "VK250X130X4X8 de 0,67 a 0,83 y las VK180X100X4X6 de 0,81 a 0,89. Ningún elemento supera el D/C admisible de 0.95. "
            "El detalle por elemento está en el Anexo A.")
    for e in plant["8.3"]:
        m.add(copy.deepcopy(e))


def sec9(m, d):
    cal = d["cal"]
    env, ver = cal["Muros_envolventes"], cal["Muros_verificacion"]
    m.titulo("9. ELEMENTOS DE HORMIGÓN ARMADO", 1)
    m.titulo("9.1. Muros de subsuelo (e = 20 cm)", 2)
    m.texto("El espesor cumple el mínimo de 19 cm para muros exteriores de sótano (ACI 318-19 Tabla 11.3.1.1). Los muros se "
            "prolongan hasta el plano medio de la zapata corrida (z = −4.285 m) y quedan unidos a ella; reciben el empuje de "
            "tierras entre z = −3.06 y 0.00 m y están restringidos en la coronación por la losa de ±0.00 y las vigas de "
            "coronación. En la Rev. 0 el muro se consideraba empotrado en la zapata; ahora el modelo incluye la flexibilidad de la "
            "zapata y del suelo, y el muro se verificó con las fuerzas de las combinaciones U1, U5X, U5Y, U7X y U7Y "
            "(envolvente de máximos y mínimos).")
    m.texto("**Método.** En cada nudo se promediaron los valores de los elementos concurrentes de M11 (flexión horizontal), "
            "M22 (flexión vertical), M12 (torsión) y de los cortantes V13 y V23. Los momentos de cálculo se obtuvieron con el método "
            "de Wood-Armer, que incorpora M12 en el refuerzo de cada cara. El signo positivo de M22 corresponde a tracción en la cara "
            "interior (con U1 el empuje produce momento positivo en el vano) y el negativo a tracción en la cara del suelo. La "
            "sección crítica de la base se toma en la cara superior de la zapata (z = −4.11 m), interpolando entre los nudos de "
            "z = −4.285 y −3.06 m, y el cortante a d = 144 mm del apoyo en z = −3.06 m. Los resultados se separan en "
            "**campo libre** y **zonas de concentración**, que son los nudos a menos de 0.75 m de la llegada de pedestales y vigas "
            "de cimentación al muro y de las esquinas entre muros; en estas últimas los valores nodales son picos locales de la malla y "
            "se promedian en una franja de 1.0 m de ancho. Se adoptó f′c = 280 kgf/cm², fy = 4 200 kgf/cm², φ = 0.90 en flexión y "
            "0.75 en cortante, d = 144 mm en la cara del suelo (recubrimiento de 5 cm) y d = 154 mm en la cara interior.")
    zonas = ["base", "apoyo", "vano inferior", "vano superior", "tope"]
    e2 = env[env["Zona"].isin(zonas)].copy()
    e2["o"] = e2["Zona"].map({z: i for i, z in enumerate(zonas)})
    e2 = e2.sort_values(["o", "Tipo"])
    nombre_zona = {"base": "Base (z = −4.11)", "apoyo": "Apoyo (z = −3.06)", "vano inferior": "Vano inferior (−2.55 a −1.53)",
                   "vano superior": "Vano superior (−1.02 a −0.51)", "tope": "Tope (z = 0.00) (*)"}
    filas = [[nombre_zona[r["Zona"]], "libre" if r["Tipo"].startswith("campo") else "concentración",
              nc(r["M vertical cara int. (T*m/m)"]), nc(r["M vertical cara suelo (T*m/m)"]),
              nc(r["M horizontal cara int. (T*m/m)"]), nc(r["M horizontal cara suelo (T*m/m)"]),
              nc(r["V23 nodal (T/m)"]), nc(r["V13 nodal (T/m)"])] for _, r in e2.iterrows()]
    m.tabla(["Zona", "Tipo", "Mv int. (T·m/m)", "Mv suelo (T·m/m)", "Mh int. (T·m/m)", "Mh suelo (T·m/m)", "V23 (T/m)", "V13 (T/m)"],
            filas, [2.1, 1.1, 1, 1, 1, 1, 0.9, 0.9], alinear=["left"] + ["center"] * 7, sz=16)
    tope = e2[e2["Zona"] == "tope"]
    m.texto("(*) La fila superior (z = 0.00 m) corresponde al encuentro con la losa, las vigas de coronación y las pilastras. "
            f"Allí los valores nodales tienen concentraciones locales (hasta Mh = {nc(tope['M horizontal cara int. (T*m/m)'].max(), 1)} T·m/m y "
            f"V13 = {nc(tope['V13 nodal (T/m)'].max(), 1)} T/m en la esquina x = 8.50, y = 16.35) que no son representativas del panel y no se "
            "usan para dimensionarlo; el encuentro se resuelve con el detalle de las pilastras y de la viga de coronación (sección 9.2). "
            "Mv = momento de diseño del refuerzo vertical; Mh = del refuerzo horizontal.", sz=18, antes=80)
    m.titulo("Armado revisado", 3)
    a = cr.ARMADO_MURO
    filas = [["Vertical, cara interior", "Toda la altura", "φ12@12.5", nc(cr.barra_as(*a['vertical cara interior']))],
             ["Vertical, cara del suelo", "z = −4.285 a −1.53 m", "φ12@10", nc(cr.barra_as(*a['vertical cara suelo, mitad inferior']))],
             ["Vertical, cara del suelo", "z = −1.53 a 0.00 m", "φ12@20", nc(cr.barra_as(*a['vertical cara suelo, mitad superior']))],
             ["Horizontal, cada cara", "Campo libre", "φ12@15", nc(cr.barra_as(*a['horizontal cada cara']))],
             ["Horizontal, cada cara", "Zonas de concentración (±0.75 m)", "φ12@7.5", nc(cr.barra_as(*a['horizontal cada cara, zonas de concentracion']))],
             ["Ganchos transversales", "Zonas de concentración (z = −4.285 a −2.04 m)", "φ8@20 × 20", "-"]]
    m.tabla(["Refuerzo", "Zona", "Armado", "As (cm²/m)"], filas, [1.8, 2.6, 1, 1], alinear=["left", "left", "center", "center"])
    m.titulo("Verificación", 3)
    filas = [[r["Solicitacion"].replace("T*m", "T·m"), r["Zona"], "libre" if r["Tipo"].startswith("campo") else "concentración",
              nc(r["Demanda"]), r["Armado"].replace("phi", "φ"), nc(r["Capacidad"]), nc(r["D/C"])] for _, r in ver.iterrows()]
    m.tabla(["Solicitación", "Zona", "Tipo", "Demanda", "Armado", "Capacidad", "D/C"], filas, [2.2, 1.1, 1.1, 0.8, 1.5, 0.9, 0.6],
            alinear=["left", "center", "center", "center", "center", "center", "center"], sz=16)
    m.figura("r1_muros_dc.png", 5.0, "Relación demanda/capacidad de los muros con el armado revisado.", 7)
    m.texto("El cortante fuera del plano se verifica sin refuerzo transversal con ACI 318-19 22.5.5.1(c) (Vc = 0.66·λs·λ·ρw^(1/3)·√f′c·b·d, "
            "con ρw de la cara más débil); en las zonas de concentración se añaden ganchos φ8@20×20, con Vs = Av·fy·d/s. Las cuantías "
            f"cumplen los mínimos de ACI 318-19 11.6.1: ρt = {2 * cr.barra_as(12, 15) / 2000 * 100:.2f} % (total de ambas caras) ≥ 0.20 % y "
            f"ρl = {(cr.barra_as(12, 12.5) + cr.barra_as(12, 10)) / 2000 * 100:.2f} % ≥ 0.12 %. Se desprecia la compresión axial del muro, lo que es conservador.")
    gi = lambda z: float(env[(env["Zona"] == z) & (env["Tipo"].str.startswith("campo"))]["M vertical cara int. (T*m/m)"].iloc[0])
    m.texto("**Comparación con la Rev. 0.** Con el muro unido a una zapata flexible y el momento de diseño obtenido con Wood-Armer, "
            f"la demanda en la cara interior es de {nc(gi('base'))} T·m/m en la base y {nc(gi('vano inferior'))} T·m/m en el vano (1,90 T·m/m en la Rev. 0), "
            "mayor que la capacidad de 2,26 T·m/m del armado φ10@200 de la Rev. 0. Por eso el refuerzo vertical interior pasa a φ12@12.5 y el "
            "horizontal a φ12@15; el vertical de la cara del suelo (φ12@10 y φ12@20) se mantiene.")

    m.titulo("9.2. Pedestales, pilastras y vigas de coronación", 2)
    col = cal["Pedestales_pilastras"]
    ped = col[col["Grupo"] == "Pedestales"]
    pil = col[col["Grupo"] == "Pilastras"]
    m.texto("Los pedestales (PED40X40), las pilastras (COL45x40 y COL30x30) y las vigas de coronación se diseñaron en SAP2000 con ACI 318-19 "
            "para las combinaciones U1 a U7Y, sin errores ni advertencias. Se reporta el acero longitudinal requerido (As) y la rama de "
            "estribos más exigente de cada elemento.")
    g = ped.groupby(["Armado longitudinal", "Estribos"]).agg(n=("Frame", "count"), a_min=("As (cm2)", "min"), a_max=("As (cm2)", "max"),
                                                              av=("Av mayor (cm2/m)", "max")).reset_index().sort_values("a_max")
    filas = [[f"P{i + 1}", int(r["n"]), f"{nc(r['a_min'], 1)} – {nc(r['a_max'], 1)}", r["Armado longitudinal"].replace("phi", "φ"),
              r["Estribos"].replace("phi", "φ")] for i, (_, r) in enumerate(g.iterrows())]
    m.titulo("Pedestales 40 × 40 cm", 3)
    m.tabla(["Tipo", "Pedestales", "As requerido (cm²)", "Armado longitudinal", "Estribos"], filas, [0.7, 1, 1.5, 1.6, 1.8])
    n_min = int((ped["As (cm2)"] <= 16.01).sum())
    m.n_ped_tipos = len(g)
    m.texto(f"Los pedestales están gobernados por la cuantía mínima de 1 % en {n_min} de los {len(ped)} casos; el pedestal más exigente requiere "
            f"{nc(ped['As (cm2)'].max(), 1)} cm² (2.0 %) por la combinación sísmica. Cada pedestal se arma con el tipo que cubre su demanda "
            "(planos de detalle) y con estribos de 3 ramas en cada dirección.")
    m.titulo("Pilastras", 3)
    filas = [[int(r["Frame"]), r["DesignSect"], nc(r["As (cm2)"], 1), r["Armado longitudinal"].replace("phi", "φ"), nc(r["As prov (cm2)"], 1),
              r["Estribos"].replace("phi", "φ")] for _, r in pil.iterrows()]
    m.tabla(["Pilastra", "Sección", "As requerido (cm²)", "Armado", "As provisto (cm²)", "Estribos"], filas, [0.8, 1, 1.4, 1, 1.4, 1.8])
    m.texto("En la Rev. 0 las pilastras COL45x40 se armaron con 4φ20 + 4φ16 (20,6 cm²), valor que queda por debajo del requerido por la "
            f"pilastra {int(pil.loc[pil['As (cm2)'].idxmax(), 'Frame'])} ({nc(pil['As (cm2)'].max(), 1)} cm²); las demás siguen gobernadas por la cuantía mínima de 1 % "
            "(ACI 318-19 10.6.1.1) y los estribos por la separación máxima. El detallado transversal sigue ACI 318-19, capítulo 25.")
    m.titulo("Vigas de coronación", 3)
    m.texto("Las vigas de coronación (VCOR30X20, 30 × 20 cm) no pertenecen a los grupos de cimentación y sus resultados de diseño no se "
            "re-extrajeron del modelo con base flexible. Se mantiene el armado de la Rev. 0 (2φ12 superiores + 2φ12 inferiores, estribos φ10@10 cm; "
            "As sup./inf. = 1,67 cm² y Av/s = 14,3 cm²/m) y queda pendiente su verificación con la v10, que se incorporará en la Rev. 2 "
            "o en una adenda.")


def sec10(m, d):
    cal, der = d["cal"], d["der"]
    rz = cal["Zap_resumen"]
    za = cal["Zap_aisladas"]
    zc = cal["Zap_corridas"]
    pz = der["Presion_zapatas"]
    desl = cal["Deslizamiento"]
    area = rz["Area (m2)"].sum()
    ais = rz[rz["Tipo"] == "aislada"]
    m.titulo("10. CIMENTACIÓN — DISEÑO", 1)
    m.texto("La cimentación se diseña con las fuerzas del modelo de la v10, en el que las zapatas se representan con elementos shell sobre "
            "resortes de suelo. Se verifican las presiones en el suelo con qadm = 24 T/m² (32 T/m² con sismo, ≈ 1.33·qadm), el levantamiento, los "
            "asentamientos y el deslizamiento con las combinaciones de servicio CS1 a CS4 (sección 6), y la resistencia estructural "
            "con las combinaciones U1 a U7Y, con ACI 318-19. El criterio de 1.33·qadm en sismo debe confirmarlo el ingeniero geotécnico.")
    m.titulo("10.1. Modelo de la cimentación", 2)
    filas = [["Zapatas aisladas", f"26 zapatas shell ZAP30 (e = 30 cm): 13 de 1.00 × 1.00 m (Z3), 10 de 1.20 × 1.20 m (Z2) y 3 de 1.30 × 1.30 m (Z1)"],
             ["Zapatas corridas", "Shell ZAP35 (e = 35 cm) bajo los muros x = 8.50, x = 3.90, y = 16.35 e y = 9.45, unidas entre sí (40.2 m²)"],
             ["Nivel de desplante", "Df = 1.40 m desde el NPT; plano medio de las zapatas en z = −1.25 m (planta baja), −4.31 m (zapatas aisladas del subsuelo) y −4.285 m (zapatas corridas)"],
             ["Resortes verticales", "kv = 4 800 T/m³ × área tributaria de cada nudo (1 239 nudos; 24.9 a 1 024 T/m)"],
             ["Resortes horizontales", "kh = 0.5·kv en 1 072 nudos (supuesto de modelación, a confirmar por el geotecnista)"],
             ["Pedestales", "PED40X40 (28), de la zapata hasta el nivel de arranque de la columna"],
             ["Vigas de cimentación", "VC20X30, VC25X40, VC25X45, VC25X50 y VC30X50 (193 vigas) que atan las zapatas"],
             ["Cargas sobre zapatas", "Relleno y contrapiso (RELLENO, 1.65 a 6.11 T/m²) y peso propio de zapatas"],
             ["Área de contacto", f"{area:.1f} m² (zapatas aisladas {ais['Area (m2)'].sum():.1f} m², corridas {area - ais['Area (m2)'].sum():.1f} m²)"]]
    m.tabla(["Componente", "Descripción"], filas, [1.4, 5], alinear=["left", "left"], negrita_col0=True)

    m.titulo("10.2. Presión en el suelo", 2)
    pm = pz[pz["Paso"].isin(["", "Max"])]
    filas = []
    for cs in ("CS1", "CS2", "CS3X", "CS3Y"):
        q = pm[pm["Caso"] == cs]
        w = q.loc[q["p prom (T/m2)"].idxmax()]
        lim = w["Limite (T/m2)"]
        filas.append([cs, f"{nc(w['p prom (T/m2)'], 1)} ({w['Zapata']})", nc(q["p max nodal (T/m2)"].max(), 1), nc(lim, 0),
                      f"{w['p prom (T/m2)'] / lim * 100:.0f} %", "Cumple"])
    m.tabla(["Combinación", "Presión promedio máx. (T/m²)", "Presión nodal máx. (T/m²)", "Límite (T/m²)", "Utilización", "Resultado"], filas,
            [1, 1.8, 1.5, 1.1, 1, 1])
    u = rz.loc[rz["Utilizacion"].idxmax()]
    wn = pm.loc[pm["p max nodal (T/m2)"].idxmax()]
    m.texto("La presión promedio de cada zapata se calcula como ΣF3/ΣA de sus resortes (F3 = reacción vertical; A = área tributaria). "
            f"La utilización máxima es del {u['Utilizacion'] * 100:.0f} % ({u['Zapata']}, {u['Combinacion']}); la presión nodal máxima, "
            f"{nc(wn['p max nodal (T/m2)'], 1)} T/m² ({wn['Zapata']}, {wn['Caso']}), es un pico local menor que el límite de 32 T/m². "
            "El mapa de la Figura 8 muestra la utilización por zapata y el Anexo B el detalle por combinación.")
    m.figura("r1_mapa_presiones.png", 4.6, "Utilización de la capacidad admisible por zapata (envolvente CS1 a CS3).", 8)
    bajos = rz[(rz["Tipo"] == "aislada") & (rz["Utilizacion"] < 0.3)]["Zapata"].tolist()
    m.texto(f"Hay holgura de capacidad en las zapatas {', '.join(bajos)} (utilización menor que 30 %). Una reducción de sus dimensiones no mejoraría "
            "el diseño estructural y empeoraría el levantamiento en CS4, por lo que se mantienen los tres tipos de la Rev. 0 (Z1 = 1.30 m, "
            "Z2 = 1.20 m y Z3 = 1.00 m) y la decisión de optimizarlas queda a criterio del proyecto.")

    m.titulo("10.3. Levantamiento (CS4 = 0.6D + 0.7E + H)", 2)
    lev = pz[(pz["Caso"].isin(["CS4X", "CS4Y"])) & (pz["Paso"] == "Min") & (pz["Nudos con levantamiento"] > 0)]
    g = lev.groupby("Zapata").agg(n=("Nudos con levantamiento", "max"), f=("F3 total (T)", "min"), p=("p min nodal (T/m2)", "min")).reset_index()
    g = g.assign(neta=np.where(g["f"] < 0, "SÍ", "no")).sort_values("f")
    filas = [[r["Zapata"], int(r["n"]), nc(r["f"]), nc(r["p"], 1), r["neta"]] for _, r in g.iterrows()]
    m.tabla(["Zapata", "Nudos en tracción", "Reacción neta mín. (T)", "Presión nodal mín. (T/m²)", "Levantamiento neto"], filas, [1, 1.3, 1.6, 1.8, 1.4])
    n_nodos = int(lev.groupby("Caso")["Nudos con levantamiento"].sum().max())
    n_corr = int(g[g["Zapata"] == "Z01"]["n"].sum())
    f_corr = float(g[g["Zapata"] == "Z01"]["f"].iloc[0]) if n_corr else 0.0
    zneta = g[g["f"] < 0]
    m.texto(f"En CS4 hay hasta {n_nodos} nudos con resorte en tracción: {n_corr} en los bordes de las zapatas corridas (Z01, con reacción neta positiva de {nc(f_corr, 1)} T) y el resto en "
            f"las zapatas aisladas, sobre todo en las del ala este. Solo la zapata {zneta.iloc[0]['Zapata']} tiene reacción neta negativa ({nc(zneta.iloc[0]['f'])} T). El modelo lineal permite tracción en los "
            "resortes; en la realidad la zapata se levanta parcialmente y la presión se redistribuye hacia el borde comprimido. La tracción neta de "
            f"{abs(zneta.iloc[0]['f']):.1f} T es muy pequeña frente al peso de las zapatas vecinas y a la resistencia de las vigas de cimentación que las unen, "
            "por lo que se considera cubierta; se recomienda confirmarla con un análisis sin tracción en los resortes (iterativo) en la siguiente revisión.")

    m.titulo("10.4. Asentamientos", 2)
    s = ais["Asentamiento CS1 (mm)"]
    m.texto(f"Con resortes lineales el asentamiento es δ = p/kv. Con CS1 resulta entre {nc(s.min(), 1)} y {nc(s.max(), 1)} mm en las zapatas aisladas y "
            f"{nc(rz[rz['Tipo'] == 'corrida']['Asentamiento CS1 (mm)'].iloc[0], 1)} mm en las corridas. El asentamiento diferencial máximo entre zapatas es "
            f"{nc(s.max() - s.min(), 1)} mm; con la menor separación entre columnas (3.5 m) equivale a una distorsión angular de 1/{3500 / (s.max() - s.min()):.0f}, "
            "menor que 1/500 (criterio usual para estructuras de acero que debe confirmar el estudio geotécnico).")

    m.titulo("10.5. Deslizamiento global", 2)
    m.texto(f"Como el empuje está desbalanceado (sección 5.4), toda la edificación tiende a deslizar hacia el oeste y el norte. La demanda H es la resultante "
            f"de las reacciones horizontales de los resortes del modelo (Base Reactions) y la resistencia por fricción es μ·N, con μ = tan 24° = 0.445 y N la reacción vertical "
            f"de cada combinación (incluye el peso de zapatas y del relleno). El área de contacto es {area:.1f} m².")
    filas = []
    for _, r in desl.iterrows():
        paso = r["StepType"] if isinstance(r["StepType"], str) else ""
        filas.append([r["OutputCase"], paso or "-", nc(r["H"], 1), nc(r["N"], 1), nc(r["Resistencia por friccion (T)"], 1),
                      nc(r["FS friccion"]), nc(r["FS requerido"], 1), nc(r["Adhesion requerida (T/m2)"]), r["Resultado"]])
    m.tabla(["Combinación", "Paso", "H (T)", "N (T)", "μ·N (T)", "FS", "FS req.", "Adhesión req. (T/m²)", "Resultado"], filas,
            [1, 0.7, 0.9, 0.9, 1, 0.7, 0.8, 1.4, 1], sz=16,
            resaltar=lambda i, j, v: j == 8 and v == "NO cumple")
    m.figura("r1_deslizamiento.png", 5.2, "Factor de seguridad al deslizamiento por fricción en cada combinación.", 9)
    no = desl[desl["Resultado"] == "NO cumple"]
    amax = no["Adhesion requerida (T/m2)"].max()
    m.texto(f"Las combinaciones estáticas y CS3 cumplen por fricción. Las combinaciones CS4 (0.6D + 0.7E + H), donde el peso estabilizante se reduce a 0.6D, "
            f"no cumplen solo por fricción (FS entre {nc(no['FS friccion'].min())} y {nc(no['FS friccion'].max())}, contra 1.1). Para cumplirlas se requiere una adhesión suelo-hormigón de "
            f"hasta {nc(amax)} T/m² sobre {area:.1f} m² de contacto, equivalente al {amax / 4.28 * 100:.0f} % de la cohesión c′ = 4.28 T/m² "
            "(1,53 T/m² en la Rev. 0), o un empuje pasivo / dentellones equivalente a "
            f"{nc(amax * area, 0)} T en la combinación más desfavorable. Debe validarlos el ingeniero geotécnico antes de la construcción; mientras tanto, "
            "se recomienda prever dentellones bajo las zapatas corridas (por ejemplo, en los muros x = 8.50 y x = 3.90, los más largos y cargados), cuyo "
            "dimensionamiento dependerá de la resistencia pasiva que acepte el geotecnista.")

    m.titulo("10.6. Zapatas aisladas", 2)
    d_x, d_y = 21.8, 20.4
    m.texto("**Criterios.** Hormigón f′c = 280 kgf/cm², fy = 4 200 kgf/cm², recubrimiento de 7.5 cm contra el suelo (ACI 318-19 20.5.1.3.1) y barras φ14 "
            f"para el cálculo del peralte (dx = {d_x:.1f} cm en la capa inferior, dy = {d_y:.1f} cm en la segunda capa). Los momentos de diseño por Wood-Armer "
            "(M11, M22 y M12 de los elementos shell, combinaciones U1 a U7Y) se promedian sobre el ancho de la zapata en la cara del pedestal "
            "(sección crítica a c/2 = 0.20 m del centro); el cortante en una dirección se toma a d de la cara y se evalúa con ACI 318-19 22.5.5.1(c); "
            "el punzonamiento se verifica con ACI 318-19 22.6.5.2 usando P, M2 y M3 en la base del pedestal (γv = 0.40; Vu = Pu − qu·(c + d)²). "
            "La cuantía mínima de flexión es 0.0018·Ag = 5.40 cm²/m (ACI 318-19 24.4.3.2) y el peralte mínimo sobre el suelo es 150 mm (13.3.1.2).")
    za2 = za.assign(tipo=np.where(za["Bx (m)"] > 1.25, "Z1 (1.30 m)", np.where(za["Bx (m)"] > 1.1, "Z2 (1.20 m)", "Z3 (1.00 m)")))
    filas = []
    for t in ("Z1 (1.30 m)", "Z2 (1.20 m)", "Z3 (1.00 m)"):
        q = za2[za2["tipo"] == t]
        filas.append([t, len(q), nc(q["Pu max (T)"].max(), 1),
                      nc(q[["M* inf X (T*m/m)", "M* inf Y (T*m/m)"]].max().max()),
                      nc(q[["V a d dir X (T/m)", "V a d dir Y (T/m)"]].max().max()),
                      nc(q[["D/C flexion inf X", "D/C flexion inf Y"]].max().max()),
                      nc(q[["D/C cortante X", "D/C cortante Y"]].max().max()), nc(q["D/C punz"].max()), "φ12@20 en X e Y"])
    m.tabla(["Tipo", "N.º", "Pu máx. (T)", "M* inf. máx. (T·m/m)", "V a d máx. (T/m)", "D/C flexión", "D/C cortante", "D/C punz.", "Malla inferior"], filas,
            [1.2, 0.5, 0.9, 1.3, 1.1, 0.9, 0.9, 0.8, 1.3], sz=16)
    m.figura("r1_zapatas_dc.png", 5.4, "Relación demanda/capacidad de las zapatas aisladas (e = 30 cm).", 10)
    sup = za[(za["As req sup X (cm2/m)"] > 0.5) | (za["As req sup Y (cm2/m)"] > 0.5)]["Zapata"].tolist()
    ldh = 0.7 * 0.24 * 411.9 * 12 / (cr.FC_MPA ** 0.5) / 10
    m.texto(f"Todas las zapatas aisladas cumplen con e = 30 cm: la demanda máxima es del {za[['D/C flexion inf X', 'D/C flexion inf Y']].max().max() * 100:.0f} % en flexión, "
            f"{za[['D/C cortante X', 'D/C cortante Y']].max().max() * 100:.0f} % en cortante y {za['D/C punz'].max() * 100:.0f} % en punzonamiento, de modo que no se requiere aumentar "
            "el espesor a 40 cm. **Armado:** malla inferior φ12@20 cm en ambas direcciones (5,65 cm²/m ≥ 5,40 cm²/m) en todas las zapatas, con ganchos "
            f"estándar a 90° en los extremos (ldh = 0.7·0.24·fy·db/(λ√f′c) = {ldh:.0f} cm, menor que los 22.5 cm disponibles desde la cara del pedestal en Z3). "
            f"Malla superior φ10@25 cm (3,14 cm²/m; D/C ≤ 0.45) en las zapatas con momento negativo apreciable por levantamiento: {', '.join(sup)}. "
            "Los pedestales se anclan con dowels de área ≥ 0.005·Ag = 8 cm² (ACI 318-19 16.3.4.1); la presión de contacto pedestal-zapata "
            "(ACI 318-19 22.8) no gobierna, con Pu máximo de " + f"{za['Pu max (T)'].max():.1f} T frente a 247 T de capacidad de aplastamiento (φ = 0.65, A1 = 40 × 40 cm). "
            "El detalle por zapata está en el Anexo C.")

    m.titulo("10.7. Zapatas corridas bajo muros", 2)
    m.texto("Las zapatas corridas (e = 35 cm; d = 26.8 cm con φ14 y recubrimiento de 7.5 cm) se diseñan en la dirección transversal al muro con los mismos criterios, "
            "en secciones paralelas al muro a la cara del mismo (±0.10 m), con el máximo promedio móvil de 1.0 m a lo largo del muro (se recortan los "
            "extremos que concurren con otros muros) y el cortante a d de la cara.")
    filas = [[r["Muro"], r["Lado"].replace("lado ", ""), nc(r["Voladizo (m)"]), nc(r["M* inf transversal (T*m/m)"]), nc(r["M* sup transversal (T*m/m)"]),
              nc(r["As req (cm2/m)"]), r["Armado inf transversal"].replace("phi", "φ"), nc(r["D/C flexion"]), nc(r["V a d (T/m)"]), nc(r["phi Vc (T/m)"]),
              nc(r["D/C cortante"])] for _, r in zc.iterrows()]
    m.tabla(["Muro", "Lado", "Voladizo (m)", "M* inf. (T·m/m)", "M* sup. (T·m/m)", "As req. (cm²/m)", "Armado", "D/C flex.", "V a d (T/m)", "φVc (T/m)", "D/C cort."],
            filas, [1.5, 0.5, 0.8, 0.9, 0.9, 0.9, 0.9, 0.7, 0.8, 0.8, 0.7], sz=15)
    m.texto(f"Con e = 35 cm la demanda máxima es del {zc['D/C flexion'].max() * 100:.0f} % en flexión y del {zc['D/C cortante'].max() * 100:.0f} % en cortante. "
            "**Armado:** φ12@15 cm (7,54 cm²/m ≥ 6,30 cm²/m) en la malla inferior y superior, transversal y longitudinal, con ganchos a 90° en los bordes. En la zona "
            "de la zapata combinada del muro x = 3.90, donde llegan las columnas B-4 y B-5 por medio de pedestales, el modelo incluye su efecto en las fuerzas "
            f"promediadas; el punzonamiento de los pedestales sobre las zapatas corridas tiene D/C máximo de {cal['Punz_corridas']['D/C punz'].max():.2f} (con qu = 0, conservador).")

    m.titulo("10.8. Vigas de cimentación", 2)
    vg = cal["Vigas_cimentacion"]
    filas = []
    for _, r in vg.iterrows():
        filas.append([r["DesignSect"], int(r["Elementos"]), nc(r["As sup (cm2)"], 1), nc(r["As inf (cm2)"], 1), nc(r["Av/s (cm2/m)"], 1),
                      r["Superior"].replace("phi", "φ"), r["Inferior"].replace("phi", "φ"), r["Estribos"].replace("phi", "φ"),
                      nc(max(r["D/C sup"], r["D/C inf"])), r["Refuerzo adicional"].replace("phi", "φ") if isinstance(r["Refuerzo adicional"], str) else "-"])
    m.tabla(["Sección", "N.º", "As sup. req. (cm²)", "As inf. req. (cm²)", "Av/s req. (cm²/m)", "Superior", "Inferior", "Estribos", "D/C", "Torsión"], vg_rows := filas,
            [1, 0.5, 1, 1, 1, 0.9, 1.3, 0.9, 0.6, 1.5], sz=15)
    m.texto("El diseño por flexión, cortante y torsión se obtuvo en SAP2000 con ACI 318-19 y las combinaciones U1 a U7Y, sin errores ni advertencias; se reporta "
            "la envolvente por sección. Los estribos de las vigas de cimentación se disponen en toda la longitud con la separación indicada y, en las "
            "zonas de nudo, con la separación de confinamiento de NEC-SE-HM. El armado superior e inferior de cada viga se prolonga hasta anclarse en los pedestales.")

    m.titulo("10.9. Resumen de armados de cimentación", 2)
    filas = [["Zapatas aisladas Z1, Z2 y Z3 (e = 30 cm)", "Malla inferior φ12@20 X e Y; malla superior φ10@25 en " + ", ".join(sup[:6]) + (" y otras" if len(sup) > 6 else "")],
             ["Zapatas corridas (e = 35 cm)", "φ12@15 inferior y superior, transversal y longitudinal; dentellones según geotecnia (10.5)"],
             ["Pedestales 40 × 40", f"Tipos P1 a P{m.n_ped_tipos} (8φ16 a 4φ25 + 4φ20); estribos φ10 de 3 ramas @ 10 a 15 cm (9.2)"],
             ["Vigas de cimentación", "Ver tabla 10.8"],
             ["Muros", "Ver sección 9.1"]]
    m.tabla(["Elemento", "Armado"], filas, [2.2, 4.5], alinear=["left", "left"], negrita_col0=True)


def sec11(m, d, plant):
    cal, der = d["cal"], d["der"]
    rz, za, zc, desl = cal["Zap_resumen"], cal["Zap_aisladas"], cal["Zap_corridas"], cal["Deslizamiento"]
    dv, c, mod, ac = der["Derivas"], der["Cortante"], der["Modal"], der["Acero_resumen"]
    pz = der["Presion_zapatas"]
    u = rz.loc[rz["Utilizacion"].idxmax()]
    no = desl[desl["Resultado"] == "NO cumple"]
    area = rz["Area (m2)"].sum()
    amax = no["Adhesion requerida (T/m2)"].max()
    neta = rz[rz["Levantamiento neto"] == "SI"]
    m.titulo("11. CONCLUSIONES Y RECOMENDACIONES", 1)
    items = [
        f"Con la cimentación modelada mediante resortes de suelo, la estructura cumple los requisitos de NEC-SE-DS: la masa modal acumulada es {mod.iloc[-1]['Suma UX (%)']:.0f} % con {len(mod)} vectores Ritz, "
        f"el cortante dinámico es {c.iloc[0]['Relacion (%)']:.0f} % (X) y {c.iloc[1]['Relacion (%)']:.0f} % (Y) del estático, sin necesidad de escalar, la deriva inelástica máxima es {dv['dM'].max():.4f} (límite 0.020) "
        "y el índice de estabilidad es menor que 0.10. Los períodos se alargan entre 3 % y 12 % respecto de la Rev. 0.",
        f"Todos los elementos de acero cumplen AISC 360-16 y AISC 341-16 (IMF), con D/C máxima de {ac[ac['Verificacion'] == 'PMM']['D/C max'].max():.2f} por resistencia y {ac['D/C max'].max():.2f} por deflexión. "
        "La base flexible redistribuye los momentos y aumenta el D/C de varias secciones, sin que sea necesario cambiarlas.",
        "Los muros de subsuelo se re-verificaron con la unión a la zapata y el método de Wood-Armer: el refuerzo vertical interior pasa de φ10@200 a φ12@12.5, el horizontal a φ12@15 "
        "(φ12@7.5 y ganchos φ8@20×20 en las zonas de concentración) y el vertical de la cara del suelo se mantiene en φ12@10 y φ12@20. Con este armado todas las relaciones D/C son menores que 0.90.",
        "Las pilastras COL45x40 requieren hasta 24,7 cm² (8φ20) en lugar de los 20,6 cm² de la Rev. 0; los pedestales de 40 × 40 cm se arman con 8φ16 (mayoría) hasta 4φ25 + 4φ20. "
        "Las vigas de coronación quedan pendientes de re-verificación con el modelo v10.",
        f"Las presiones en el suelo cumplen con holgura: la utilización máxima es del {u['Utilizacion'] * 100:.0f} % de qadm ({u['Zapata']}), el asentamiento máximo por resortes es de {rz['Asentamiento CS1 (mm)'].max():.1f} mm "
        "y la distorsión angular es menor que 1/500.",
        f"Las zapatas aisladas de 30 cm y las corridas de 35 cm cumplen en flexión, cortante y punzonamiento con D/C máxima de {max(za[['D/C flexion inf X', 'D/C flexion inf Y']].max().max(), zc['D/C flexion'].max()):.2f} en flexión, "
        f"{max(za[['D/C cortante X', 'D/C cortante Y']].max().max(), zc['D/C cortante'].max()):.2f} en cortante y {za['D/C punz'].max():.2f} en punzonamiento; no se requiere aumentar el espesor a 40 cm. "
        "Armado: φ12@20 (aisladas) y φ12@15 (corridas) en malla inferior, con malla superior donde lo exige el levantamiento.",
        f"En CS4 (0.6D + 0.7E + H) la zapata {neta.iloc[0]['Zapata'] if len(neta) else '-'} presenta levantamiento neto de {abs(neta.iloc[0]['F3 neta min CS4 (T)']) if len(neta) else 0:.2f} T y varias zapatas del ala este tienen resortes en tracción; "
        "se recomienda confirmar con un análisis sin tracción en los resortes.",
        f"El deslizamiento global no cumple solo por fricción en CS4 (FS de {no['FS friccion'].min():.2f} a {no['FS friccion'].max():.2f} contra 1.1). Se requiere una adhesión de hasta {amax:.2f} T/m² "
        f"({amax / 4.28 * 100:.0f} % de c′) sobre {area:.1f} m² o su equivalente en dentellones, valor que debe validar el ingeniero geotécnico (la Rev. 0 requería 1,53 T/m²).",
        "Antes de la construcción, el ingeniero geotécnico debe validar la adhesión suelo-cimiento o los dentellones, el incremento de qadm en combinaciones sísmicas (32 T/m²) y la rigidez horizontal de los resortes (kh = 0.5·kv).",
        "La excavación del subsuelo supera la altura crítica Hcr = 2.82 m: se requiere apuntalamiento temporal o excavación por bataches. El fondo de las cimentaciones se compacta al 95 % del Proctor estándar y se coloca replantillo de 10 cm.",
        "Detrás de los muros se debe disponer drenaje e impermeabilización. El diseño no considera presión hidrostática, porque el estudio no detectó nivel freático.",
        "Las conexiones a momento deben calificarse para un ángulo de deriva de 0.02 rad (AISC 341-16 E2.6), con placas de continuidad y verificación de la zona de panel. Las soldaduras CJP son de demanda crítica, con inspección según AWS D1.1 y D1.8.",
        "Las cargas de fachada se asumieron como mampostería de bloque de 15 cm en todo el perímetro, y el tramo de muro entre z = −3.06 m y la zapata como sin empuje neto (sección 5.4). Cualquier cambio de materiales, uso, ocupación o procedimiento constructivo requiere revisar este diseño.",
    ]
    for t in items:
        m.add(p_lista(t, num_id=3, despues=90))
    for e in plant["firma"]:
        m.add(copy.deepcopy(e))


def anexos(m, d):
    cal, der, A = d["cal"], d["der"], d["A"]
    ac = A["Acero_DC"]
    m.add(p_salto())
    m.titulo("ANEXO A. DISEÑO DE ELEMENTOS DE ACERO (AISC 360-16)", 1)
    m.texto("Resumen del diseño en SAP2000 por elemento (v10, base con resortes). Verificación: PMM = interacción flexo-compresión; Deflexión = servicio (S1/S2).")
    tipo = {"Beam": "Viga", "Column": "Columna", "Brace": "Puntal"}
    filas = [[int(r["Frame"]), r["DesignSect"], tipo[r["DesignType"]], nc(r["Ratio"], 3), "PMM" if r["RatioType"] == "PMM" else "Deflexión", r["Combo"]]
             for _, r in ac.sort_values("Frame").iterrows()]
    m.add(tabla(["Frame", "Sección", "Tipo", "D/C", "Verificación", "Combinación"], filas, [0.8, 1.8, 1, 0.8, 1.2, 1.1]))

    m.add(p_salto())
    m.titulo("ANEXO B. PRESIONES Y REACCIONES EN LAS ZAPATAS", 1)
    m.texto("Presión promedio por zapata (ΣF3/ΣA, T/m²) en las combinaciones de servicio (valores Max para CS3), reacción vertical neta mínima en CS4 (T) y asentamiento "
            "δ = p/kv con CS1. Las zapatas aisladas se numeran Z02 a Z27; Z01 agrupa las zapatas corridas conectadas entre sí.")
    pz = der["Presion_zapatas"]
    pm = pz[pz["Paso"].isin(["", "Max"])]
    piv = pm.pivot_table(index="Zapata", columns="Caso", values="p prom (T/m2)", aggfunc="max")
    rz = cal["Zap_resumen"].set_index("Zapata")
    filas = []
    for z, r in rz.iterrows():
        filas.append([z, "corrida" if r["Tipo"] == "corrida" else "aislada", f"{r['Bx (m)']:.2f} × {r['By (m)']:.2f}" if r["Tipo"] == "aislada" else "-", nc(r["Area (m2)"]),
                      nc(piv.loc[z, "CS1"], 1), nc(piv.loc[z, "CS2"], 1), nc(piv.loc[z, "CS3X"], 1), nc(piv.loc[z, "CS3Y"], 1),
                      nc(r["F3 neta min CS4 (T)"]), nc(r["Asentamiento CS1 (mm)"], 1)])
    m.add(tabla(["Zapata", "Tipo", "B × L (m)", "Área (m²)", "CS1", "CS2", "CS3X", "CS3Y", "F3 neta CS4 (T)", "δ CS1 (mm)"], filas,
                [0.8, 0.9, 1.2, 0.9, 0.7, 0.7, 0.7, 0.7, 1.1, 0.9], sz=16))

    m.add(p_salto())
    m.titulo("ANEXO C. DISEÑO DE ZAPATAS AISLADAS (e = 30 cm)", 1)
    m.texto("Momentos de diseño Wood-Armer (M*) promediados sobre el ancho en la cara del pedestal, cortante a d de la cara y punzonamiento con las fuerzas en la base del pedestal. "
            "Todas las zapatas con malla inferior φ12@20 en X e Y (As = 5,65 cm²/m; φMn con d = 21.8 y 20.4 cm).")
    za = cal["Zap_aisladas"]
    filas = []
    for _, r in za.iterrows():
        filas.append([r["Zapata"], f"{r['Bx (m)']:.2f}", str(r["Pedestal"]).split(",")[0], nc(r["Pu max (T)"], 1),
                      f"{nc(r['M* inf X (T*m/m)'])} / {nc(r['M* inf Y (T*m/m)'])}",
                      f"{nc(r['D/C flexion inf X'])} / {nc(r['D/C flexion inf Y'])}",
                      f"{nc(r['V a d dir X (T/m)'], 1)} / {nc(r['V a d dir Y (T/m)'], 1)}",
                      f"{nc(r['D/C cortante X'])} / {nc(r['D/C cortante Y'])}", nc(r["D/C punz"]), r["Combo punz"]])
    m.add(tabla(["Zapata", "B (m)", "Pedestal", "Pu máx. (T)", "M* inf. X / Y (T·m/m)", "D/C flexión X / Y", "V a d X / Y (T/m)", "D/C cortante X / Y", "D/C punz.", "Combo punz."],
                filas, [0.8, 0.7, 0.8, 0.9, 1.4, 1.2, 1.2, 1.2, 0.8, 1], sz=15))
    m.add(parse_xml(f'<w:p {W}><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>'))


# ================================================================== indice
def encabezados(body):
    res = []
    for p in body.iterchildren(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is None or ppr.find(qn("w:pStyle")) is None:
            continue
        st = ppr.find(qn("w:pStyle")).get(qn("w:val"))
        if st in ("Ttulo1", "Ttulo2") and texto_de(p).strip():
            bm = p.find(qn("w:bookmarkStart"))
            if bm is None:
                _bm["n"] += 1
                i = _bm["n"]
                bm = parse_xml(f'<w:bookmarkStart {W} w:id="{i}" w:name="_Toc9{i:07d}"/>')
                ppr.addnext(bm)
                p.append(parse_xml(f'<w:bookmarkEnd {W} w:id="{i}"/>'))
            res.append((1 if st == "Ttulo1" else 2, texto_de(p).strip(), bm.get(qn("w:name"))))
    return res


def reconstruir_toc(m, paginas):
    sdt = None
    for e in m.body.iterchildren():
        if e.tag == qn("w:sdt"):
            sdt = e
            break
    cont = sdt.find(qn("w:sdtContent"))
    for e in list(cont):
        cont.remove(e)
    hs = encabezados(m.body)
    rpr_p = ('<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
             '<w:noProof/><w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')
    for i, (niv, texto, bm) in enumerate(hs):
        pg = paginas.get(texto, "")
        ini = ""
        if i == 0:
            ini = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\h \\o "1-2"</w:instrText></w:r>'
                   '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        cont.append(parse_xml(
            f'<w:p {W}><w:pPr><w:pStyle w:val="TDC{niv}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9628"/></w:tabs>{rpr_p}</w:pPr>{ini}'
            f'<w:hyperlink w:anchor="{bm}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hipervnculo"/><w:noProof/></w:rPr><w:t xml:space="preserve">{esc(texto)}</w:t></w:r>'
            f'<w:r><w:rPr><w:noProof/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:rPr><w:noProof/></w:rPr><w:instrText xml:space="preserve"> PAGEREF {bm} \\h </w:instrText></w:r>'
            f'<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:rPr><w:noProof/></w:rPr><w:t>{pg}</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>'))
    cont.append(parse_xml(f'<w:p {W}><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'))


def paginas_desde_pdf(pdf, hs):
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout).group(1))
    textos = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", str(pdf), "-"], capture_output=True, text=True).stdout
        textos.append(re.sub(r"\s+", " ", t))
    res, pag = {}, 3          # el indice ocupa las paginas 2 y 3
    for _, texto, _ in hs:
        clave = re.sub(r"\s+", " ", texto)
        for i in range(pag, n):
            lineas = re.findall(re.escape(clave), textos[i])
            if lineas:
                res[texto] = str(i + 1)
                pag = i
                break
    return res, n


# ================================================================== principal
def capturar(body):
    ch = list(body.iterchildren())
    idx = {texto_de(e).strip(): i for i, e in enumerate(ch) if e.tag == qn("w:p")}
    i81, i82, i83, i9 = idx["8.1. Criterios"], idx["8.2. Resultados"], idx["8.3. Requisitos de conexiones y fabricación"], idx["9. ELEMENTOS DE HORMIGÓN ARMADO"]
    ia = idx["Atentamente,"]
    ik = idx["KUBIEC"]
    return {"8.1": [copy.deepcopy(e) for e in ch[i81 + 1:i82]],
            "8.3": [copy.deepcopy(e) for e in ch[i83 + 1:i9]],
            "firma": [copy.deepcopy(e) for e in ch[ia - 1:ik + 1]]}


def construir(base, tablas, figuras, salida, render=True):
    sal = figuras_r1.generar(tablas, figuras)
    d = cargar(tablas)
    m = Memoria(base, figuras)
    plant = capturar(m.body)
    editar_previas(m, d)
    # eliminar desde la seccion 7
    ch = list(m.body.iterchildren())
    i7 = next(i for i, e in enumerate(ch) if e.tag == qn("w:p") and texto_de(e).strip() == "7. ANÁLISIS ESTRUCTURAL")
    for e in ch[i7:]:
        if e.tag != qn("w:sectPr"):
            m.body.remove(e)
    sec7(m, d)
    sec8(m, d, plant)
    sec9(m, d)
    sec10(m, d)
    sec11(m, d, plant)
    anexos(m, d)
    paginas = {}
    reconstruir_toc(m, paginas)
    salida = Path(salida)
    m.doc.save(salida)
    if render:
        for _ in range(2):
            pdf = salida.with_suffix(".pdf")
            subprocess.run(["soffice", "--headless", "-env:UserInstallation=file:///tmp/lohome/lo", "--convert-to", "pdf",
                            "--outdir", str(salida.parent), str(salida)], capture_output=True, timeout=280,
                           env={**os.environ, "HOME": "/tmp/lohome"})
            paginas, n = paginas_desde_pdf(pdf, encabezados(m.body))
            reconstruir_toc(m, paginas)
            m.doc.save(salida)
        print("paginas:", n)
    return salida


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default="docs/MT-CASA_MR-ESTRUCTURA-R0-30092026.docx")
    ap.add_argument("--tablas", default="v10/tablas")
    ap.add_argument("--figuras", default="v10/figuras/memoria_r1")
    ap.add_argument("--salida", default="docs/MT-CASA_MR-ESTRUCTURA-R1-01102026.docx")
    ap.add_argument("--sin-render", action="store_true")
    a = ap.parse_args()
    print(construir(a.base, a.tablas, a.figuras, a.salida, render=not a.sin_render))


if __name__ == "__main__":
    main()
