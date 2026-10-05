#!/usr/bin/env python3
"""
Genera las laminas de detalles (DXF) a partir de la lamina E-10 (misma hoja, cajetin y capas):
    E-14  placas base y anclajes, losa colaborante
    E-15  conexion viga-columna (VK270 / VK250 / VK220) y viga secundaria - viga principal
    E-16  vista 3D de la estructura metalica (solo perfiles)
    python scripts/planos_generar.py [E14|E15|E16|todo]
Las medidas de los detalles estan en mm reales; el modelo esta en metros y cada detalle se amplia (1:5 -> k = 20).
"""
import sys
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).parent))
import planos_lib as L                    # noqa: E402
import prediseno_placas_base as PB        # noqa: E402
import prediseno_conexiones as CX         # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]
PLANTILLA = RAIZ / "planos" / "E-10_CUADRO_COLUMNAS_VIGAS_corregida.dxf"
FECHA = "2026-10-05"
K5 = 20.0       # escala 1:5
K10 = 10.0      # escala 1:10

# --------------------------------------------------------------------------- placas base
PLACA = 380     # mm, lado de la placa (pedestal de 400 mm)
GAUGE = 250     # mm entre ejes de anclajes
GROUT = 40
EMBEB = 400     # mm de empotramiento de los anclajes


def zigzag(det, x1, x2, z):
    """Linea de rotura (zigzag)."""
    n = 5
    pts = [(x1 + (x2 - x1) * i / n, z + (12 if i % 2 else -12) * (0 if i in (0, n) else 1)) for i in range(n + 1)]
    det.poly(pts, "E_CORTE", cerrada=False)


def pb_planta(det, d, b, t, marca):
    """Planta de la placa base (centro en 0,0). d = dimension de la columna en X, b = en Y (mm)."""
    h = PLACA / 2
    det.rect(-h, -h, PLACA, PLACA, "E_COLUMNAS")
    det.rect(-d / 2, -b / 2, d, b, "E_COLUMNAS")
    det.rect(-d / 2 + 8, -b / 2 + 8, d - 16, b - 16, "E_ARMADO_FINO")
    det.rect(-200, -200, 400, 400, "E_PEDESTALES", "DASHED")
    det.rect(-50, -50, 100, 100, "E_ARMADO_FINO", "DASHED")       # taco de corte (oculto)
    for sx in (-1, 1):
        for sy in (-1, 1):
            x, y = sx * GAUGE / 2, sy * GAUGE / 2
            det.circle(x, y, 9.5, "E_ARMADO")
            det.circle(x, y, 12.5, "E_ARMADO_FINO")
            det.rect(x - 25, y - 25, 50, 50, "E_ARMADO_FINO")      # arandela cuadrada 50 x 50 x 8
    det.eje(-h - 40, 0, h + 40, 0)
    det.eje(0, -h - 40, 0, h + 40)
    # cortes A-A
    for sx in (-1, 1):
        det.line(sx * (h + 55), -18, sx * (h + 55), 18, "E_CORTE")
    det.texto(-h - 90, 0, "A", 0.22, "E_TEXTO", "cen")
    det.texto(h + 90, 0, "A", 0.22, "E_TEXTO", "cen")
    # cotas
    det.dim_h(-h, h, -h, -45, f"{PLACA}")
    det.dim_h(-GAUGE / 2, GAUGE / 2, -h, -20, f"{GAUGE}")
    det.dim_h(-h, -GAUGE / 2, -h, -20, f"{int((PLACA - GAUGE) / 2)}")
    det.dim_h(GAUGE / 2, h, -h, -20, f"{int((PLACA - GAUGE) / 2)}")
    det.dim_v(-h, h, h, 55, f"{PLACA}")
    det.dim_v(-GAUGE / 2, GAUGE / 2, h, 25, f"{GAUGE}")
    det.dim_h(-d / 2, d / 2, h, 25, f"{d}")
    det.dim_v(-b / 2, b / 2, -h, -25, f"{b}")
    # llamadas
    det.nota(-h + 8, h - 8, -h - 60, h + 80, f"PL {PLACA}x{PLACA} (A36), t segun cuadro", 0.16)
    det.nota(GAUGE / 2, GAUGE / 2, h + 40, h + 85, "4 anclajes φ19 F1554 Gr.36", 0.16)
    det.nota(35, -35, 120, -h - 130, "taco de corte []100x100x6 (oculto)", 0.16)
    det.nota(h + 5, h - 60, h + 70, 20 + h + 10, "pedestal 40x40 (P1-P3)", 0.16) if False else None


def pb_corte(det, d, t, marca, largo_anclaje):
    """Corte A-A. z = 0 en la cara superior del pedestal; x = 0 en el eje de la columna. d = ancho visto (mm)."""
    zp = GROUT                        # cara inferior de la placa
    zt = GROUT + t                    # cara superior de la placa
    hc = 260                          # altura de columna dibujada
    # pedestal y su armado
    det.rect(-200, -450, 400, 450, "E_PEDESTALES")
    det.hatch([(-200, -450), (200, -450), (200, 0), (-200, 0)], "AR-CONC", 0.045, "E_HATCH")
    for sx in (-1, 1):
        det.line(sx * 165, -445, sx * 165, -8, "E_ARMADO")                 # barras verticales del pedestal
        det.line(sx * 140, -445, sx * 140, -8, "E_ARMADO_FINO", "DASHED")
    for z in (-30, -80, -130, -180, -300, -420):
        det.line(-170, z, 170, z, "E_ARMADO_FINO")                          # estribos phi10
    # grout y placa
    det.rect(-190, 0, 380, GROUT, "E_ARMADO_FINO")
    det.rect(-190, zp, 380, t, "E_COLUMNAS")
    det.hatch([(-190, zp), (190, zp), (190, zt), (-190, zt)], "ANSI31", 0.06, "E_HATCH")
    # columna cajon con relleno
    det.line(-d / 2, zt, -d / 2, zt + hc, "E_COLUMNAS")
    det.line(d / 2, zt, d / 2, zt + hc, "E_COLUMNAS")
    det.line(-d / 2 + 8, zt, -d / 2 + 8, zt + hc, "E_ARMADO_FINO")
    det.line(d / 2 - 8, zt, d / 2 - 8, zt + hc, "E_ARMADO_FINO")
    det.hatch([(-d / 2, zt), (-d / 2 + 8, zt), (-d / 2 + 8, zt + hc), (-d / 2, zt + hc)], "ANSI31", 0.05, "E_HATCH")
    det.hatch([(d / 2 - 8, zt), (d / 2, zt), (d / 2, zt + hc), (d / 2 - 8, zt + hc)], "ANSI31", 0.05, "E_HATCH")
    det.hatch([(-d / 2 + 8, zt), (d / 2 - 8, zt), (d / 2 - 8, zt + hc), (-d / 2 + 8, zt + hc)], "AR-SAND", 0.045, "E_HATCH")
    zigzag(det, -d / 2 - 15, d / 2 + 15, zt + hc)
    # soldadura de filete perimetral 6 mm
    for sx in (-1, 1):
        x0 = sx * d / 2
        det.solido([(x0, zt), (x0 + sx * 6, zt), (x0, zt + 6)], "E_HATCH", 0)
    # taco de corte y bolsillo
    det.rect(-50, zp - 75, 100, 75, "E_COLUMNAS")
    det.rect(-75, -45, 150, 45, "E_ARMADO_FINO", "DASHED")
    # anclajes
    for sx in (-1, 1):
        x = sx * GAUGE / 2
        top = zt + 8 + 15 + 7
        det.rect(x - 9.5, -EMBEB, 19, EMBEB + top, "E_ARMADO")
        det.rect(x - 25, zt, 50, 8, "E_ARMADO_FINO")                       # arandela superior
        det.rect(x - 15, zt + 8, 30, 15, "E_ARMADO")                       # tuerca
        det.rect(x - 15, zp - 23, 30, 15, "E_ARMADO_FINO")                 # tuerca de nivelacion
        det.rect(x - 25, zp - 8, 50, 8, "E_ARMADO_FINO")                   # arandela de nivelacion
        det.rect(x - 25, -EMBEB - 8, 50, 8, "E_ARMADO")                    # arandela de anclaje inferior
        det.rect(x - 15, -EMBEB, 30, 15, "E_ARMADO_FINO")                  # tuerca inferior
    # cotas verticales (a la derecha)
    xr = 200
    det.dim_v(0, zp, xr, 60, f"{GROUT}")
    det.dim_v(zp, zt, xr, 60, f"{t}")
    det.dim_v(-EMBEB, 0, xr, 60, f"hef {EMBEB}")
    det.dim_v(zp - 75, zp, -d / 2 - 15 if False else -200, -60, "75")
    det.dim_h(-200, 200, -450, -40, "400")
    det.dim_h(-d / 2, d / 2, zt + hc, 40, f"{d}")
    det.dim_h(-GAUGE / 2, GAUGE / 2, -EMBEB - 8, -40 if False else -85, f"{GAUGE}")
    # llamadas
    det.nota(-190, zp + t / 2, -420, zp + t + 80, f"PL {PLACA}x{PLACA}x{t}", 0.16, "der")
    det.nota(0, zp - 40, -330, zp - 120, "taco []100x100x6 L=75", 0.16, "der")
    det.nota(-GAUGE / 2, -250, -640, -250, "anclaje φ19 L=" + str(largo_anclaje), 0.16, "der")
    det.nota(-110, -135, -640, -160, "estribos φ10@50 (3 en 150 mm)", 0.16, "der")
    det.nota(165, -200, 330, -330, "8 barras verticales del pedestal", 0.16)
    det.nota(0, 20, 120, 120, "grout sin retraccion (e=40)", 0.16)
    det.nota(d / 2 - 4, zt + 120, d / 2 + 90, zt + 190, f"columna cajon (rellena f'c>=210)", 0.16)
    det.nota(d / 2, zt + 3, d / 2 + 120, zt + 70, "filete 6 mm perimetral", 0.16)


def hoja_E14(salida):
    doc = L.hoja_nueva(PLANTILLA, "E-14", ["- PLACAS BASE Y ANCLAJES,", "  LOSA COLABORANTE"], FECHA)
    ms = doc.modelspace()
    L.titulo(ms, 3.0, 56.4, "PLACAS BASE Y ANCLAJES - PRE-DISEÑO (A VERIFICAR EN CYPE)")
    df = PB.calcular()
    # --- PB-1 (200x200) arriba, PB-2 (250x200) abajo
    L.titulo(ms, 3.0, 53.9, "PB-1: COLUMNA []200X200X8", 0.26)
    d1 = L.Detalle(ms, 9.0, 46.0, K5)
    pb_planta(d1, 200, 200, "t", "PB-1")
    d1.texto(0, -280, "PLANTA  ESC 1:5", 0.2, "E_TEXTO", "cen")
    d1c = L.Detalle(ms, 25.0, 41.0, K5)
    pb_corte(d1c, 200, 32, "PB-1", 530)
    d1c.texto(0, -520, "CORTE A-A  ESC 1:5", 0.2, "E_TEXTO", "cen")
    L.titulo(ms, 3.0, 28.4, "PB-2: COLUMNA []250X200X8", 0.26)
    d2 = L.Detalle(ms, 9.0, 20.0, K5)
    pb_planta(d2, 250, 200, "25", "PB-2")
    d2.texto(0, -280, "PLANTA  ESC 1:5", 0.2, "E_TEXTO", "cen")
    d2c = L.Detalle(ms, 25.0, 15.0, K5)
    pb_corte(d2c, 250, 25, "PB-2", 515)
    d2c.texto(0, -520, "CORTE A-A  ESC 1:5", 0.2, "E_TEXTO", "cen")

    # --- cuadro de placas base
    x0 = 38.0
    L.titulo(ms, x0, 56.4, "CUADRO DE PLACAS BASE", 0.26)
    resumen = {}
    for tipo, g in df.groupby("tipo"):
        resumen[tipo] = g
    def cols(t):
        return ", ".join(resumen[t]["col"])
    filas = [["MARCA", "COLUMNA", "PLACA (mm)", "ANCLAJES (4)", "TACO", "SOLDADURA", "DEMANDA MAX."],
             ["PB-1L", "[]200X200X8", "PL 380x380x20", "φ19 F1554 Gr.36 L=510", "[]100x100x6 L=75", "filete 6 mm", f"Pu {resumen['PB-1L']['Pu'].max():.1f} T  Tu {resumen['PB-1L']['Tu'].max():.1f} T"],
             ["PB-1", "[]200X200X8", "PL 380x380x32", "φ19 F1554 Gr.36 L=530", "[]100x100x6 L=75", "filete 6 mm", f"V {resumen['PB-1']['V'].max():.1f} T  Tu {resumen['PB-1']['Tu'].max():.1f} T"],
             ["PB-2", "[]250X200X8", "PL 380x380x25", "φ19 F1554 Gr.36 L=515", "[]100x100x6 L=75", "filete 6 mm", f"Pu {resumen['PB-2']['Pu'].max():.1f} T  V {resumen['PB-2']['V'].max():.1f} T"]]
    y = L.tabla_simple(ms, x0, 55.4, [1.6, 2.6, 3.0, 4.4, 3.4, 2.4, 4.4], filas, 0.7, 0.15)
    y = L.parrafo(ms, x0, y - 0.5, [
        f"PB-1L: {cols('PB-1L')}.",
        f"PB-1: {cols('PB-1')}.",
        f"PB-2: {cols('PB-2')}.",
        "Columnas C-1 a C-4 y de la torre: sobre pilastras de H.A. (embebidas); no llevan placa base y su detalle de embebido no esta en esta lamina."], 0.15, 0.3)

    # --- verificacion
    L.titulo(ms, x0, y - 0.7, "VERIFICACION (PREDISEÑO)", 0.26)
    filas = [["MARCA", "COLUMNA CRITICA", "t req. (mm)", "T anclaje / φNsa", "T anc. conserv.", "V / φVtaco", "ESPESOR"]]
    for t in ("PB-1L", "PB-1", "PB-2"):
        g = resumen[t]
        i = g["t_req_mm"].idxmax()
        filas.append([t, g.loc[i, "col"], f"{g['t_req_mm'].max():.1f}", f"{g['DC_T'].max():.2f}", f"{g['DC_T_conserv'].max():.2f}",
                      f"{g['DC_V_lug'].max():.2f}", f"{int(g['t_adopt_mm'].max())}"])
    y2 = L.tabla_simple(ms, x0, y - 1.5, [1.6, 3.0, 2.2, 3.6, 3.4, 2.6, 2.0], filas, 0.7, 0.15)
    y2 = L.parrafo(ms, x0, y2 - 0.5, [
        "BASES DE CALCULO: AISC Design Guide 1 (placa con excentricidad), AISC 360-16 J8, ACI 318-19 cap. 17 y 22.8. A36 en placas, F1554 Gr.36 en anclajes, E70XX.",
        "Cargas: tabla de cargas REV11 en la base del pedestal. Momento en la placa = |Mu - Vu x 1.25 m| (altura del pedestal); a confirmar.",
        "T anclaje: Tu repartida en 4 anclajes y traccion por momento (DG1). 'conserv.' suma ademas el momento completo de la fila (no simultaneo).",
        "El cortante lo toma el taco de corte (aplastamiento phi=0.65 x 0.85 f'c x 10x7.5 cm = 11.6 T por direccion); los anclajes no se cuentan para cortante.",
        "Traccion en anclajes: se transmite a las barras verticales del pedestal (8 phi16 min.) con estribos phi10@50 en los 150 mm superiores (ACI 17.5.2).",
        "Distancia de anclaje al borde del pedestal: 75 mm (4.7 da). Si se exige 6 da para anclajes pretensados, aumentar el pedestal a 45x45 cm.",
        "Soldadura de la columna a la placa: filete perimetral 6 mm (phi Rn = 0.94 T/cm >= 0.58 T/cm); si el EOR la califica de demanda critica, usar CJP con respaldo.",
        "Nivelacion con tuercas y arandelas; relleno con grout sin retraccion f'c >= 350 kg/cm2; apretar las tuercas a ajuste ceñido (snug-tight).",
        "Longitud de anclaje L = hef 400 + grout 40 + t + 50 mm; extremo inferior con arandela 50x50x8 y tuerca (resistencia a arrancamiento no gobierna)."], 0.15, 0.3)

    # --- losa colaborante (parte inferior derecha)
    losa_colaborante(ms, x0 + 4.0, y2)
    doc.saveas(salida)
    return doc


def deck(det, x0, x1, ytop_beam, esp=100, h_deck=50, paso=150, valle=75):
    """Perfil de la placa colaborante (trapezoidal) desde x0 hasta x1, apoyada con su cara inferior en y=ytop_beam."""
    pts = []
    x = x0
    y0 = ytop_beam
    while x < x1:
        pts += [(x, y0 + h_deck), (x + (paso - valle) / 2 - 8, y0 + h_deck), (x + (paso - valle) / 2 + 4, y0),
                (x + (paso - valle) / 2 + 4 + valle - 12, y0), (x + (paso - valle) / 2 + valle - 4, y0 + h_deck)]
        x += paso
    pts.append((x, y0 + h_deck))
    det.poly(pts, "E_ARMADO", cerrada=False)


def losa_colaborante(ms, xorg, ytop):
    L.titulo(ms, xorg, ytop - 0.5, "LOSA COLABORANTE (ESPESOR TOTAL 10 cm)", 0.26)
    # DETALLE 1: corte tipico sobre viga principal VK220 (ala 110)
    d = L.Detalle(ms, xorg + 6.0, ytop - 6.2, K5)
    bf, tf, dd = 110, 8, 220
    ytb = 0                              # cara superior del ala superior de la viga
    d.rect(-bf / 2, ytb - tf, bf, tf, "E_VIGAS_ACERO")
    d.rect(-2, ytb - dd + tf, 4, dd - 2 * tf, "E_VIGAS_ACERO")
    d.rect(-bf / 2, ytb - dd, bf, tf, "E_VIGAS_ACERO")
    d.hatch([(-bf / 2, ytb - tf), (bf / 2, ytb - tf), (bf / 2, ytb), (-bf / 2, ytb)], "ANSI31", 0.05, "E_HATCH")
    deck(d, -230, 230, ytb)
    d.poly([(-230, ytb + 100), (230, ytb + 100)], "E_VIGAS_ACERO", False)
    d.hatch([(-230, ytb + 50), (230, ytb + 50), (230, ytb + 100), (-230, ytb + 100)], "AR-CONC", 0.045, "E_HATCH")
    for xm in np.arange(-210, 230, 150):
        d.circle(xm, ytb + 76, 2.7, "E_ARMADO")            # malla phi5
    d.line(-230, ytb + 76, 230, ytb + 76, "E_ARMADO_FINO", "DASHED")
    d.rect(-6.5, ytb, 13, 70, "E_ARMADO")                  # conector phi13 L=75
    d.rect(-10, ytb + 70, 20, 5, "E_ARMADO")
    d.dim_v(ytb, ytb + 100, 230, 40, "100")
    d.dim_v(ytb, ytb + 50, -230, -40, "50")
    d.dim_v(ytb - dd, ytb, -bf / 2 - 10, -50, "220") if False else None
    d.nota(120, ytb + 100, 260, ytb + 190, "concreto f'c >= 210 kg/cm2", 0.15)
    d.nota(60, ytb + 76, 260, ytb + 150, "malla electrosoldada φ5 @ 15x15", 0.15)
    d.nota(-80, ytb + 25, -420, ytb + 110, "placa colaborante e=0.76 mm, h=50", 0.15, "der")
    d.nota(0, ytb + 40, -420, ytb + 190, "conector φ13 L=75 @ 300 (solo vigas principales)", 0.15, "der")
    d.nota(-bf / 2 + 10, ytb - tf / 2, -420, ytb - 100, "viga principal VK (ala superior)", 0.15, "der")
    d.texto(0, ytb - dd - 60, "DETALLE 1: CORTE SOBRE VIGA PRINCIPAL  ESC 1:5", 0.18, "E_TEXTO", "cen")
    # DETALLE 2: borde de losa con cierre de borde
    d2 = L.Detalle(ms, xorg + 22.0, ytop - 6.2, K5)
    bf2 = 100
    d2.rect(-bf2 / 2 - 100, ytb - 8, bf2 + 100, 8, "E_VIGAS_ACERO")
    d2.hatch([(-bf2 / 2 - 100, ytb - 8), (bf2 / 2, ytb - 8), (bf2 / 2, ytb), (-bf2 / 2 - 100, ytb)], "ANSI31", 0.05, "E_HATCH")
    d2.rect(-100 - bf2 / 2, ytb - 150, 4, 142, "E_VIGAS_ACERO")
    deck(d2, -300, bf2 / 2 + 10, ytb, paso=150)
    d2.poly([(-300, ytb + 100), (bf2 / 2 + 10, ytb + 100), (bf2 / 2 + 10, ytb)], "E_VIGAS_ACERO", False)
    d2.hatch([(-300, ytb + 50), (bf2 / 2 + 10, ytb + 50), (bf2 / 2 + 10, ytb + 100), (-300, ytb + 100)], "AR-CONC", 0.045, "E_HATCH")
    d2.poly([(bf2 / 2 + 10, ytb + 100), (bf2 / 2 + 10, ytb - 4), (bf2 / 2 - 40, ytb - 4)], "E_ARMADO", False)   # cierre de borde L
    d2.nota(bf2 / 2 + 10, ytb + 50, bf2 / 2 + 90, ytb + 150, "cierre de borde L 100x50x3", 0.15)
    d2.nota(-300, ytb + 70, -300, ytb + 190, "borde de placa colaborante", 0.15, "izq")
    d2.nota(-bf2 / 2 - 100, ytb - 70, -bf2 / 2 - 120, ytb - 200, "viga de borde VK150X100X4X4", 0.15, "der")
    d2.texto(-100, ytb - 260, "DETALLE 2: BORDE DE LOSA  ESC 1:5", 0.18, "E_TEXTO", "cen")
    # DETALLE 3: apoyo y traslape de la placa sobre la viga
    d3 = L.Detalle(ms, xorg + 6.0, ytop - 16.0, K5)
    bf3 = 100
    d3.rect(-bf3 / 2, ytb - 8, bf3, 8, "E_VIGAS_ACERO")
    d3.hatch([(-bf3 / 2, ytb - 8), (bf3 / 2, ytb - 8), (bf3 / 2, ytb), (-bf3 / 2, ytb)], "ANSI31", 0.05, "E_HATCH")
    d3.rect(-2, ytb - 120, 4, 112, "E_VIGAS_ACERO")
    d3.poly([(-250, ytb + 50), (-20, ytb + 50), (-5, ytb), (60, ytb)], "E_ARMADO", False)
    d3.poly([(-60, ytb + 4), (-10, ytb + 4), (5, ytb + 54), (250, ytb + 54)], "E_ARMADO", False)
    d3.circle(0, ytb + 2, 10, "E_ARMADO_FINO")
    d3.dim_h(-60, 60, ytb, 70, "traslape >= 100")
    d3.dim_h(-bf3 / 2, bf3 / 2, ytb - 8, -35, "apoyo >= 40")
    d3.nota(0, ytb + 2, 120, ytb + 120, "soldadura de tapon φ19 @ 300", 0.15)
    d3.nota(40, ytb + 54, 140, ytb + 190, "tornillo autoperforante lateral @ 600", 0.15)
    d3.texto(0, ytb - 190, "DETALLE 3: APOYO Y TRASLAPE DE LA PLACA  ESC 1:5", 0.18, "E_TEXTO", "cen")
    # notas
    L.parrafo(ms, xorg + 17.0, ytop - 13.0, [
        "NOTAS: la losa se calcula como NO COMPUESTA (la memoria no considera accion compuesta); el peso propio es 190 kg/m2.",
        "Los conectores φ13 x 75 solo transmiten el cortante del diafragma (la cabeza queda a 25 mm de la cara superior).",
        "Placa colaborante galvanizada e = 0.76 mm, h = 50 mm; apoyo minimo 40 mm; traslape de extremo 100 mm.",
        "Malla electrosoldada φ5 @ 15x15 cm: 1.31 cm2/m (>= 0.0018 x 5 cm x 100); recubrimiento superior 25 mm.",
        "Apuntalar la placa a luces mayores a 2.0 m durante el vaciado; el hormigon se vibra y cura >= 7 dias.",
        "Verificar con el fabricante la capacidad en voladizo de los aleros (0.50 m) y la perforacion para pasantes."], 0.15, 0.3)


# --------------------------------------------------------------------------- conexiones (E-15)
def cx_elevacion(det, nombre, d, bf, tf, tw, tdp, h_tab):
    """Elevacion de la conexion viga-columna. Origen: eje de la columna (x) y cara inferior del ala inferior (y)."""
    cb = 200
    ext = 60                                      # el diafragma sobresale 60 mm de las caras de la columna
    xe = cb / 2 + ext                             # borde del diafragma (160)
    xr = 610                                      # corte de la viga
    y0, y1 = -300, d + 300
    # columna cajon
    for x in (-cb / 2, cb / 2):
        det.line(x, y0, x, y1, "E_COLUMNAS")
    for x in (-cb / 2 + 8, cb / 2 - 8):
        det.line(x, y0, x, y1, "E_ARMADO_FINO")
    zigzag(det, -cb / 2 - 15, cb / 2 + 15, y1)
    zigzag(det, -cb / 2 - 15, cb / 2 + 15, y0)
    # diafragmas exteriores (placas de continuidad)
    for yc in (tf / 2, d - tf / 2):
        det.rect(-xe, yc - tdp / 2, 2 * xe, tdp, "E_ARMADO")
        det.hatch([(-xe, yc - tdp / 2), (xe, yc - tdp / 2), (xe, yc + tdp / 2), (-xe, yc + tdp / 2)], "ANSI31", 0.05, "E_HATCH")
    # viga: alas y alma desde el borde del diafragma
    det.rect(xe, 0, xr - xe, tf, "E_VIGAS_ACERO")
    det.rect(xe, d - tf, xr - xe, tf, "E_VIGAS_ACERO")
    det.line(xe, tf, xr, tf, "E_VIGAS_ACERO")
    det.line(xe, d - tf, xr, d - tf, "E_VIGAS_ACERO")
    zigzag(det, xr, xr, 0)
    det.line(xr, 0, xr, d, "E_CORTE")
    # placa de corte (A36 e=8) soldada a la columna y al alma
    y_a, y_b = (d - h_tab) / 2, (d + h_tab) / 2
    det.rect(cb / 2, y_a, 160, h_tab, "E_ARMADO", "DASHED")
    # soldaduras de la viga al diafragma (CJP): triangulo en la junta
    for yf in (tf, d - tf):
        pass
    for xx, yy, sgn in ((xe, 0.0, 1), (xe, d, -1)):
        det.solido([(xe, yy), (xe + 8, yy), (xe, yy + sgn * 8)], "E_HATCH", 0)
    # cotas
    det.dim_v(0, d, xr, 55, f"{int(d)}")
    det.dim_v(y_a, y_b, xr, 22, f"{int(h_tab)}") if False else None
    det.dim_h(-cb / 2, cb / 2, y1, 45, f"{cb}")
    det.dim_h(cb / 2, xe, d + tdp, 25, f"{ext}")
    det.dim_h(xe, xr - 200, -tf, -40, "zona protegida: d") if False else None
    det.dim_v(d - tdp - 12, d + 12, -xe, -40, f"PL {int(tdp)}") if False else None
    # llamadas (todas hacia la derecha, sobre y bajo la viga)
    det.nota(xe - 25, d - tf / 2, 250, d + 190, f"diafragma exterior PL {int(tdp)} A572 Gr.50", 0.15)
    det.nota(xe + 8, d, 250, d + 110, "CJP ala-diafragma (demanda critica)", 0.15)
    det.nota(xe - 25, tf / 2, 250, -100, "CJP diafragma-columna (perimetral)", 0.15)
    det.nota(cb / 2 + 100, y_a + 20, 250, -190, f"placa de corte PL 8 x {int(h_tab)} A36 (filete 6 mm columna / 3 mm alma)", 0.15)
    det.nota(xr - 40, d / 2, 250, -280, f"{nombre}: d={int(d)} bf={int(bf)} tf={int(tf)} tw={int(tw)}", 0.15)
    det.texto(150, y0 - 150, f"CONEXION {nombre} - COLUMNA  ESC 1:5", 0.19, "E_TEXTO", "cen")


def cx_planta(det, bf, tdp):
    """Planta al nivel del diafragma: columna 200x200, anillo de 320x320 y ala de la viga."""
    cb, xe = 200, 160
    det.rect(-xe, -xe, 2 * xe, 2 * xe, "E_ARMADO")
    det.rect(-cb / 2, -cb / 2, cb, cb, "E_COLUMNAS")
    det.rect(-cb / 2 + 8, -cb / 2 + 8, cb - 16, cb - 16, "E_ARMADO_FINO")
    det.rect(xe, -bf / 2, 450, bf, "E_VIGAS_ACERO")
    det.rect(xe, -2, 450, 4, "E_ARMADO_FINO", "DASHED")
    det.eje(-xe - 40, 0, xe + 500, 0)
    det.eje(0, -xe - 40, 0, xe + 40)
    for sy in (-1, 1):
        det.line(xe, sy * bf / 2, xe, sy * xe, "E_ARMADO_FINO")
    det.dim_h(-xe, xe, -xe, -45, "320")
    det.dim_v(-bf / 2, bf / 2, xe + 450, 40, "bf")
    det.dim_v(-xe, xe, -xe, -50, "320")
    det.nota(-xe + 10, xe - 10, -xe + 10, xe + 110, f"anillo PL {int(tdp)} (A572 Gr.50)", 0.15)
    det.nota(xe + 200, bf / 2, xe + 250, xe + 70, "ala de la viga (CJP al anillo)", 0.15)
    det.texto(150, -xe - 130, "PLANTA AL NIVEL DEL DIAFRAGMA  ESC 1:5", 0.19, "E_TEXTO", "cen")


def sec_elevacion(det):
    """Viga secundaria (VK200) sobre la viga principal VK220: vista lateral; la principal se ve en seccion."""
    d, bf, tf, tw = 220, 110, 8, 4
    ds = 200
    # viga principal en seccion (centro en x=0); alas a ras superior con la secundaria
    det.rect(-bf / 2, d - tf, bf, tf, "E_VIGAS_ACERO")
    det.rect(-bf / 2, 0, bf, tf, "E_VIGAS_ACERO")
    det.rect(-tw / 2, tf, tw, d - 2 * tf, "E_VIGAS_ACERO")
    # cartela PL 10 pasante (transversal al alma), cara superior a ras interior de las alas
    det.rect(tw / 2, tf, 175, d - 2 * tf, "E_ARMADO")
    det.hatch([(tw / 2, tf), (tw / 2 + 175, tf), (tw / 2 + 175, d - tf), (tw / 2, d - tf)], "ANSI31", 0.05, "E_HATCH")
    # secundaria: ala superior a ras de la principal (y = d)
    xs = 60
    det.rect(xs, d - tf, 480, tf, "E_VIGAS_ACERO")
    det.rect(xs, d - ds, 480, tf, "E_VIGAS_ACERO")
    det.line(xs, d - tf, xs + 480, d - tf, "E_VIGAS_ACERO")
    det.line(xs, d - ds + tf, xs + 480, d - ds + tf, "E_VIGAS_ACERO")
    det.line(xs, d - ds, xs, d, "E_VIGAS_ACERO")
    det.line(xs + 480, d - ds, xs + 480, d, "E_CORTE")
    # pernos phi16 (linea vertical a 95 mm de la cara de la cartela)
    xb = 115
    for yb in (d - ds / 2 + 45, d - ds / 2 - 45):
        det.circle(xb, yb, 8, "E_ARMADO")
        det.circle(xb, yb, 14, "E_ARMADO_FINO")
    # cotas
    det.dim_v(0, d, -bf / 2, -45, f"{d}")
    det.dim_v(d - ds, d, xs + 480, 40, f"{ds}")
    det.dim_h(0, xs, d, 50, "60") if False else None
    det.dim_h(tw / 2, tw / 2 + 175, 0, -40, "175")
    det.dim_v(d - ds / 2 - 45, d - ds / 2 + 45, xb, 60, "90") if False else None
    det.nota(tw / 2 + 140, d / 2, 330, d / 2 - 140, "cartela PL 10 (A36) d-2tf, filete 5 mm a alma y alas", 0.15)
    det.nota(xb, d - ds / 2 + 45, 250, d + 100, "2 pernos φ16 A325-N (hilos incluidos)", 0.15)
    det.nota(-tw / 2, tf + 20, -20, -120, "viga principal VK220..270 en seccion", 0.15)
    det.nota(xs + 300, d - tf / 2, xs + 380, d + 90, "viga secundaria VK150..VK200 (ala superior a ras)", 0.15)
    det.texto(250, -110, "ELEVACION: SECUNDARIA - PRINCIPAL  ESC 1:5", 0.19, "E_TEXTO", "cen")


def sec_planta(det):
    bf, ds_b = 110, 100
    det.rect(-bf / 2, -250, bf, 500, "E_VIGAS_ACERO")
    det.rect(-2, -250, 4, 500, "E_ARMADO_FINO")
    det.rect(2, -87.5, 175, 175, "E_ARMADO", "DASHED")
    det.rect(60, -ds_b / 2, 500, ds_b, "E_VIGAS_ACERO")
    det.rect(60, -2, 500, 4, "E_ARMADO_FINO", "DASHED")
    det.circle(115, 0, 8, "E_ARMADO")
    det.eje(-bf / 2 - 40, 0, 600, 0)
    det.dim_h(-bf / 2, bf / 2, 250, 40, "110")
    det.dim_h(bf / 2, 60, -250, -40, "5") if False else None
    det.dim_v(-ds_b / 2, ds_b / 2, 560, 40, "100")
    det.nota(100, -87, 200, -200, "cartela PL 10 (oculta)", 0.15)
    det.nota(300, 40, 360, 160, "ala superior de la secundaria", 0.15)
    det.texto(250, -330, "PLANTA  ESC 1:5", 0.19, "E_TEXTO", "cen")


def hoja_E15(salida):
    doc = L.hoja_nueva(PLANTILLA, "E-15", ["- CONEXION VIGA-COLUMNA,", "  VIGA SECUNDARIA-PRINCIPAL"], FECHA)
    ms = doc.modelspace()
    L.titulo(ms, 3.0, 56.4, "CONEXIONES - PRE-DISE\u00d1O (A VERIFICAR EN CYPE)")
    df = CX.conexion_viga_columna()
    L.titulo(ms, 3.0, 54.0, "CONEXION VIGA-COLUMNA (PORTICOS IMF): DIAFRAGMAS EXTERIORES", 0.26)
    posiciones = [(6.0, "VK270", 270, 140), (26.0, "VK250", 250, 130), (46.0, "VK220", 220, 110)]
    for (ox, nombre, d, bf), (_, r) in zip(posiciones, df.iterrows()):
        det = L.Detalle(ms, ox + 2.0, 40.0, K5)
        cx_elevacion(det, nombre, d, bf, 8, 4, r["t diafragma (mm)"], r["altura placa corte (mm)"])
    # planta del anillo
    dp = L.Detalle(ms, 10.0, 24.0, K5)
    cx_planta(dp, 140, df.iloc[0]["t diafragma (mm)"])
    # secundaria-principal
    L.titulo(ms, 28.0, 31.0, "VIGA SECUNDARIA CON VIGA PRINCIPAL (ARTICULADA)", 0.26)
    de = L.Detalle(ms, 36.0, 23.0, K5)
    sec_elevacion(de)
    dpl = L.Detalle(ms, 54.0, 27.0, K5)
    sec_planta(dpl)
    # cuadros
    x0 = 3.0
    L.titulo(ms, x0, 15.2, "CUADRO DE CONEXIONES VIGA-COLUMNA", 0.26)
    filas = [["VIGA", "Mp (T*m)", "Mpr (T*m)", "Fpr ala (T)", "Vu (T)", "DIAFRAGMA", "ZONA PANEL D/C", "PLACA CORTE", "D/C ALMA-PLACA"]]
    for _, r in df.iterrows():
        filas.append([r["Viga"], f"{r['Mp (T*m)']:.1f}", f"{r['Mpr (T*m)']:.1f}", f"{r['Fpr ala (T)']:.1f}", f"{r['Vu (T)']:.1f}",
                      f"PL {int(r['t diafragma (mm)'])} A572 Gr.50", f"{r['zona de panel D/C']:.2f}", f"PL 8 x {int(r['altura placa corte (mm)'])}", f"{r['D/C alma-placa']:.2f}"])
    y = L.tabla_simple(ms, x0, 14.4, [1.6, 1.9, 2.0, 2.0, 1.6, 3.4, 2.8, 2.8, 2.8], filas, 0.7, 0.15)
    sec = CX.conexion_secundaria()
    L.titulo(ms, x0, y - 0.8, "CONEXION SECUNDARIA-PRINCIPAL (R = 6 T)", 0.26)
    filas = [["PERNOS", "\u03c6Rn corte (T)", "\u03c6Rn aplast. alma 4 mm (T)", "CAPACIDAD GRUPO (T)", "D/C PERNOS", "SOLDADURA CARTELA", "D/C SOLD."],
             ["2 \u03c616 A325-N", f"{sec['phiRn corte perno (T)']:.1f}", f"{sec['phiRn aplastamiento alma 4 mm (T)']:.1f}", f"{sec['capacidad grupo (T)']:.1f}",
              f"{sec['D/C pernos']:.2f}", "filete 5 mm, 2 lados", f"{CX.cartela_dc():.2f}"]]
    y = L.tabla_simple(ms, x0, y - 1.5, [2.4, 2.6, 3.6, 3.4, 2.4, 3.4, 2.4], filas, 0.7, 0.15)
    L.parrafo(ms, 28.0, 15.2, [
        "NOTAS GENERALES",
        "1. Pre-diseno para revision en CYPE; las conexiones del portico IMF deben calificarse (AISC 341-16 E2.6, ensayo o calificacion equivalente).",
        "2. Mpr = Cpr Ry Fy Zx (Cpr 1.15, Ry 1.1, Fy 3515 kg/cm2); Fpr = Mpr/(d - tf); Vu = 2 Mpr/Lh + Vg con L = 4.5 m y Vg = 2.0 T.",
        "3. Diafragmas exteriores PL 14 A572 Gr.50 en ambos niveles de ala; soldaduras de ala y de diafragma CJP (demanda critica), WPS y END segun AWS D1.8.",
        "4. Alma de 4 mm soldada a la placa de corte con filete 3 mm a ambos lados; 2 pernos \u03c616 A325-N de montaje en la placa de corte.",
        "5. Zona de panel: dos paredes de 8 mm de la columna; se ignora el relleno de hormigon (conservador).",
        "6. Zona protegida: d de la viga desde la cara del diafragma; no soldar conectores ni accesorios en ella.",
        "7. Secundaria articulada: cartela PL 10 pasante soldada al alma y alas de la principal; la secundaria no se corta (termina 5 mm despues del ala).",
        "8. Para primarias VK250 y VK270 la altura de la cartela se ajusta a d - 2 tf; el detalle es el mismo."], 0.15, 0.3)
    doc.saveas(salida)
    return doc


# --------------------------------------------------------------------------- vista 3D (E-16)
def _estructura_acero(s2k_path):
    import s2k_lector as S
    T = S.load(str(s2k_path))
    jc = T["JOINT COORDINATES"].assign(Joint=lambda d: d["Joint"].astype(str)).set_index("Joint")
    xyz = {j: (float(r["GlobalX"]), float(r["GlobalY"]), float(r["GlobalZ"])) for j, r in jc.iterrows()}
    cf = T["CONNECTIVITY - FRAME"]
    sec = T["FRAME SECTION ASSIGNMENTS"].set_index("Frame")["AnalSect"]
    miembros = []
    for _, r in cf.iterrows():
        s_ = sec.get(r["Frame"])
        if s_ is None or not (s_.startswith("VK") or s_.startswith("[]")):
            continue
        miembros.append((s_, xyz[str(r["JointI"])], xyz[str(r["JointJ"])]))
    return miembros


def _capa(seccion):
    if seccion.startswith("[]100"):
        return "E_ARMADO", None
    if seccion.startswith("[]"):
        return "E_COLUMNAS", None
    if seccion.startswith(("VK220", "VK250", "VK270")):
        return "E_VIGAS_ACERO", None
    return "E_ARMADO_FINO", None


def _proyectar(p, az, el):
    a, e = np.radians(az), np.radians(el)
    x, y, z = p
    xs = x * np.cos(a) + y * np.sin(a)
    prof = -x * np.sin(a) + y * np.cos(a)
    return xs, z * np.cos(e) + prof * np.sin(e)


def hoja_E16(salida, s2k_path):
    doc = L.hoja_nueva(PLANTILLA, "E-16", ["- VISTA 3D DE LA ESTRUCTURA", "  METALICA (SOLO PERFILES)"], FECHA, "1:100")
    ms = doc.modelspace()
    L.titulo(ms, 3.0, 56.4, "VISTA 3D DE LA ESTRUCTURA METALICA (SOLO PERFILES, SIN LOSAS NI MUROS)")
    mi = _estructura_acero(s2k_path)
    vistas = [("VISTA SUR-OESTE", 35, 28, 3.0, 29.0), ("VISTA NOR-ESTE", 215, 28, 3.0, 5.0)]
    # centrar cada vista en un rectangulo de 64 x 24 m
    for titulo_v, az, el, ox0, oy0 in vistas:
        pts = [(_proyectar(a, az, el), _proyectar(b, az, el), sec) for sec, a, b in mi]
        xs = [q[0] for p in pts for q in p[:2]]
        ys = [q[1] for p in pts for q in p[:2]]
        sc = min(64.0 / (max(xs) - min(xs)), 22.0 / (max(ys) - min(ys)), 2.2)
        tx = ox0 + (64.0 - sc * (max(xs) - min(xs))) / 2 - sc * min(xs)
        ty = oy0 + 1.5 - sc * min(ys)
        for (p1, p2, sec) in pts:
            capa, _ = _capa(sec)
            ms.add_line((tx + sc * p1[0], ty + sc * p1[1]), (tx + sc * p2[0], ty + sc * p2[1]), dxfattribs={"layer": capa})
        L.titulo(ms, ox0, oy0 + 24.0, f"{titulo_v}   (ESC 1:{100 / sc:.0f})", 0.26)
    # leyenda
    x0, y0 = 44.0, 54.0
    L.parrafo(ms, x0, y0, ["LEYENDA", "grueso: columnas []200X200X8 y []250X200X8", "medio: vigas principales VK220, VK250 y VK270",
                           "fino: vigas y viguetas VK150, VK180 y VK200", "puntales y columnetas []100X100X3"], 0.17, 0.35)
    doc.saveas(salida)
    # modelo 3D real (lineas 3D por capa) para abrirlo y girarlo en AutoCAD
    d3 = ezdxf_nuevo_3d(mi)
    d3.saveas(str(salida).replace(".dxf", "_modelo3D.dxf"))
    return doc


def ezdxf_nuevo_3d(miembros):
    import ezdxf
    doc = ezdxf.new("R2018")
    for nombre, color in (("E_COLUMNAS", 1), ("E_VIGAS_ACERO", 5), ("E_ARMADO_FINO", 3), ("E_ARMADO", 6)):
        doc.layers.add(nombre, color=color)
    ms = doc.modelspace()
    for sec, a, b in miembros:
        ms.add_line(a, b, dxfattribs={"layer": _capa(sec)[0]})
    return doc


if __name__ == "__main__":
    que = sys.argv[1] if len(sys.argv) > 1 else "todo"
    out = RAIZ / "planos"
    if que in ("PDF", "todo"):
        for nombre in ("E-10_CUADRO_COLUMNAS_VIGAS_corregida", "E-14_PLACAS_BASE_Y_LOSA", "E-15_CONEXIONES", "E-16_VISTA_3D_ACERO"):
            dd = L.abrir(out / f"{nombre}.dxf")
            L.vista_previa(dd, out / f"{nombre}.pdf", 33.1, 100)
        print("PDF ok")
    if que in ("XLSX", "todo"):
        import pandas as pd
        with pd.ExcelWriter(out / "prediseno_placas_base_y_conexiones.xlsx") as xw:
            PB.calcular().to_excel(xw, sheet_name="Placas_base", index=False)
            CX.conexion_viga_columna().to_excel(xw, sheet_name="Conexion_viga_columna", index=False)
            pd.Series(CX.conexion_secundaria()).to_frame("valor").to_excel(xw, sheet_name="Conexion_secundaria")
        print("XLSX ok")
    if que in ("E16", "todo"):
        d = hoja_E16(out / "E-16_VISTA_3D_ACERO.dxf", RAIZ / "v10" / "2026-09-30_CASA_MR_REV10cambios.s2k")
        L.vista_previa(d, out / "previews" / "E-16.png", 20, 110)
        print("E-16 ok")
    if que in ("E15", "todo"):
        d = hoja_E15(out / "E-15_CONEXIONES.dxf")
        L.vista_previa(d, out / "previews" / "E-15.png", 20, 110)
        print("E-15 ok")
    if que in ("E14", "todo"):
        d = hoja_E14(out / "E-14_PLACAS_BASE_Y_LOSA.dxf")
        L.vista_previa(d, out / "previews" / "E-14.png", 20, 110)
        print("E-14 ok")
