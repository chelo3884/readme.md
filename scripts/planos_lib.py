#!/usr/bin/env python3
"""Utilidades para generar laminas DXF con el formato de la lamina E-10 (A1, 84.0 x 59.4 m de modelo =
escala 1:100 de impresion; los detalles se dibujan ampliados: 1:5 -> k = 20, 1:10 -> k = 10)."""
import copy
from pathlib import Path

import ezdxf
from ezdxf.enums import TextEntityAlignment as A

REGISTRO_CORRECTO = "R. SENESCYT: 7402-R-16-28686"
REGISTRO_ERRONEO = "R. SENESCYT: 1006-2024-2847508"

CAPAS_CONTENIDO = ("E_TEXTO", "E_TITULO", "E_TABLA", "E_TEXTO_VIGA", "E_COTAS", "E_ARMADO", "E_ARMADO_FINO",
                   "E_HATCH", "E_CORTE", "E_COLUMNAS", "E_VIGAS_ACERO", "E_EJES", "E_LOSA", "E_PEDESTALES",
                   "E_ZAPATAS", "E_MUROS", "E_PILASTRAS", "E_VIGAS_CIM", "E_REPLANTILLO", "E_EJES_BURBUJA")


def abrir(ruta):
    """Lee un DXF (tambien los generados por LibreDWG desde DWG) sanando entradas defectuosas."""
    from ezdxf import recover
    doc, _ = recover.readfile(str(ruta))
    return doc


def corregir_registro(doc):
    """Reemplaza el numero de registro SENESCYT del proyectista en el cajetin. Devuelve cuantos textos cambio."""
    n = 0
    for e in doc.modelspace().query("TEXT"):
        if e.dxf.text.strip() == REGISTRO_ERRONEO:
            e.dxf.text = REGISTRO_CORRECTO
            n += 1
    return n


def hoja_nueva(plantilla, numero, lineas_contiene, fecha, escala="INDICADAS"):
    """Carga la lamina de plantilla, borra su contenido y actualiza el cajetin."""
    doc = abrir(plantilla)
    ms = doc.modelspace()
    for e in list(ms):
        if e.dxf.layer in CAPAS_CONTENIDO:
            ms.delete_entity(e)
    corregir_registro(doc)
    contiene = [e for e in ms.query("TEXT") if e.dxf.layer == "_DATEC"
                and e.dxf.text.strip().startswith(("- CUADRO DE COLUMNAS", "SECCIONES Y VIGAS"))]
    contiene.sort(key=lambda e: -e.dxf.insert.y)
    for i, e in enumerate(contiene):
        e.dxf.text = lineas_contiene[i] if i < len(lineas_contiene) else ""
        if i < len(lineas_contiene):
            e.dxf.insert = (e.dxf.insert.x, e.dxf.insert.y)
    if len(lineas_contiene) > 2:       # tercera linea: copia de la segunda un renglon mas abajo
        base = contiene[1]
        extra = ms.add_text(lineas_contiene[2], dxfattribs=dict(base.dxf.all_existing_dxf_attribs()))
        extra.dxf.insert = (base.dxf.insert.x, base.dxf.insert.y - 0.3)
    for e in ms.query("TEXT"):
        if e.dxf.layer != "_DATEC":
            continue
        t = e.dxf.text.strip()
        if t == "E-10":
            e.dxf.text = numero
        elif t == "S/E":
            e.dxf.text = escala
        elif t.startswith("2026-"):
            e.dxf.text = fecha
    return doc


class Detalle:
    """Dibuja en milimetros reales y los ubica en el modelo con la escala k (unidades de modelo por metro)."""

    def __init__(self, ms, ox, oy, k):
        self.ms, self.ox, self.oy, self.k = ms, ox, oy, k

    def P(self, x, y):
        return (self.ox + x / 1000.0 * self.k, self.oy + y / 1000.0 * self.k)

    # ---- geometria
    def line(self, x1, y1, x2, y2, layer="E_VIGAS_ACERO", lt=None):
        e = self.ms.add_line(self.P(x1, y1), self.P(x2, y2), dxfattribs={"layer": layer})
        if lt:
            e.dxf.linetype = lt
        return e

    def poly(self, pts, layer="E_VIGAS_ACERO", cerrada=True, lt=None):
        e = self.ms.add_lwpolyline([self.P(*p) for p in pts], close=cerrada, dxfattribs={"layer": layer})
        if lt:
            e.dxf.linetype = lt
        return e

    def rect(self, x, y, w, h, layer="E_VIGAS_ACERO", lt=None):
        return self.poly([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], layer, True, lt)

    def circle(self, x, y, r, layer="E_ARMADO", lt=None):
        e = self.ms.add_circle(self.P(x, y), r / 1000.0 * self.k, dxfattribs={"layer": layer})
        if lt:
            e.dxf.linetype = lt
        return e

    def hatch(self, pts, patron="ANSI31", escala=0.08, layer="E_HATCH", color=None):
        h = self.ms.add_hatch(dxfattribs={"layer": layer})
        if color is not None:
            h.dxf.color = color
        h.set_pattern_fill(patron, scale=escala)
        h.paths.add_polyline_path([self.P(*p) for p in pts], is_closed=True)
        return h

    def solido(self, pts, layer="E_HATCH", color=8):
        h = self.ms.add_hatch(color=color, dxfattribs={"layer": layer})
        h.set_solid_fill()
        h.paths.add_polyline_path([self.P(*p) for p in pts], is_closed=True)
        return h

    # ---- anotacion
    def texto(self, x, y, s, h=0.18, layer="E_TEXTO", alin="izq", rot=0.0):
        al = {"izq": A.LEFT, "cen": A.MIDDLE_CENTER, "der": A.RIGHT, "izqm": A.MIDDLE_LEFT}[alin]
        t = self.ms.add_text(s, height=h, rotation=rot, dxfattribs={"layer": layer, "style": "KUBIEC"})
        t.set_placement(self.P(x, y), align=al)
        return t

    def tick(self, x, y, layer="E_COTAS"):
        d = 0.09
        p = self.P(x, y)
        self.ms.add_line((p[0] - d, p[1] - d), (p[0] + d, p[1] + d), dxfattribs={"layer": layer})

    def dim_h(self, x1, x2, y, off, label=None, h=0.17):
        """Cota horizontal: lineas de extension desde y hasta y+off (mm sobre el modelo ampliado)."""
        yy = y + off
        self.line(x1, y, x1, yy + (6 if off > 0 else -6), "E_COTAS")
        self.line(x2, y, x2, yy + (6 if off > 0 else -6), "E_COTAS")
        self.line(x1, yy, x2, yy, "E_COTAS")
        self.tick(*self._mm(x1, yy)); self.tick(*self._mm(x2, yy))
        lab = label if label is not None else f"{abs(x2 - x1):g}"
        px, py = self.P((x1 + x2) / 2, yy)
        t = self.ms.add_text(lab, height=h, dxfattribs={"layer": "E_COTAS", "style": "KUBIEC"})
        t.set_placement((px, py + 0.05), align=A.BOTTOM_CENTER)

    def dim_v(self, y1, y2, x, off, label=None, h=0.17):
        xx = x + off
        self.line(x, y1, xx + (6 if off > 0 else -6), y1, "E_COTAS")
        self.line(x, y2, xx + (6 if off > 0 else -6), y2, "E_COTAS")
        self.line(xx, y1, xx, y2, "E_COTAS")
        self.tick(*self._mm(xx, y1)); self.tick(*self._mm(xx, y2))
        lab = label if label is not None else f"{abs(y2 - y1):g}"
        px, py = self.P(xx, (y1 + y2) / 2)
        t = self.ms.add_text(lab, height=h, rotation=90, dxfattribs={"layer": "E_COTAS", "style": "KUBIEC"})
        t.set_placement((px - 0.05, py), align=A.BOTTOM_CENTER)

    def _mm(self, x, y):       # inversa de P para usar tick() con coordenadas en mm
        return x, y

    def nota(self, x, y, xt, yt, s, h=0.16, alin="izq"):
        """Llamada: linea con punta hacia (x, y) y texto en (xt, yt)."""
        self.line(xt, yt, x, y, "E_COTAS")
        p = self.P(x, y)
        self.ms.add_circle(p, 0.04, dxfattribs={"layer": "E_COTAS"})
        self.texto(xt, yt, s, h, "E_TEXTO", alin)

    def eje(self, x1, y1, x2, y2):
        return self.line(x1, y1, x2, y2, "E_EJES", "CENTER")

    def soldadura(self, x, y, xt, yt, simbolo):
        """Simbolo simplificado: linea de referencia horizontal con el texto y flecha hacia la junta."""
        self.line(x, y, xt, yt, "E_COTAS")
        self.ms.add_line(self.P(xt, yt), (self.P(xt, yt)[0] + 2.6, self.P(xt, yt)[1]), dxfattribs={"layer": "E_COTAS"})
        p = self.P(xt, yt)
        t = self.ms.add_text(simbolo, height=0.15, dxfattribs={"layer": "E_TEXTO", "style": "KUBIEC"})
        t.set_placement((p[0] + 0.1, p[1] + 0.06), align=A.LEFT)
        tri = [(p[0] + 0.0, p[1]), (p[0] + 0.0, p[1] - 0.0)]
        a = self.P(x, y)
        self.ms.add_circle(a, 0.05, dxfattribs={"layer": "E_COTAS"})


def tabla_simple(ms, x0, y0, anchos, filas, h_fila=0.62, h=0.15, encabezado=True):
    """Tabla con lineas E_TABLA y texto E_TEXTO; coordenadas de modelo; (x0, y0) = esquina superior izquierda."""
    n = len(filas)
    W = sum(anchos)
    for i in range(n + 1):
        y = y0 - i * h_fila
        ms.add_line((x0, y), (x0 + W, y), dxfattribs={"layer": "E_TABLA"})
    x = x0
    for a in [0] + list(anchos):
        x += a
        ms.add_line((x, y0), (x, y0 - n * h_fila), dxfattribs={"layer": "E_TABLA"})
    for i, f in enumerate(filas):
        x = x0
        for j, (txt, a) in enumerate(zip(f, anchos)):
            t = ms.add_text(str(txt), height=h, dxfattribs={"layer": "E_TEXTO", "style": "KUBIEC"})
            t.set_placement((x + a / 2, y0 - i * h_fila - h_fila / 2), align=A.MIDDLE_CENTER)
            x += a
    return y0 - n * h_fila


def titulo(ms, x, y, s, h=0.30):
    t = ms.add_text(s, height=h, dxfattribs={"layer": "E_TITULO", "style": "KUBIEC"})
    t.set_placement((x, y), align=A.LEFT)


def parrafo(ms, x, y, lineas, h=0.15, paso=0.3, layer="E_TEXTO"):
    for i, s in enumerate(lineas):
        t = ms.add_text(s, height=h, dxfattribs={"layer": layer, "style": "KUBIEC"})
        t.set_placement((x, y - i * paso), align=A.LEFT)
    return y - len(lineas) * paso


def vista_previa(doc, ruta, ancho_in=17, dpi=110, extent=None):
    """PNG o PDF en blanco y negro (tal como se imprime) para revisar la lamina."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from ezdxf.addons.drawing import Frontend, RenderContext
    from ezdxf.addons.drawing.config import Configuration, ColorPolicy, BackgroundPolicy
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    fig = plt.figure(figsize=(ancho_in, ancho_in * 59.4 / 84.0))
    ax = fig.add_axes([0, 0, 1, 1])
    cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.BLACK)
    ctx = RenderContext(doc)
    Frontend(ctx, MatplotlibBackend(ax), config=cfg).draw_layout(doc.modelspace(), finalize=True)
    if extent:
        ax.set_xlim(extent[0], extent[2]); ax.set_ylim(extent[1], extent[3])
    fig.savefig(str(ruta), dpi=dpi)
    plt.close(fig)


def recorte(doc, x0, y0, x1, y1, ruta, px=1800):
    """PNG de una region (x0, y0)-(x1, y1) del modelo, en blanco y negro."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from ezdxf.addons.drawing import Frontend, RenderContext
    from ezdxf.addons.drawing.config import Configuration, ColorPolicy, BackgroundPolicy
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    w, h = x1 - x0, y1 - y0
    fig = plt.figure(figsize=(px / 100, px / 100 * h / w))
    ax = fig.add_axes([0, 0, 1, 1])
    cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.BLACK)
    Frontend(RenderContext(doc), MatplotlibBackend(ax), config=cfg).draw_layout(doc.modelspace(), finalize=True)
    ax.set_xlim(x0, x1); ax.set_ylim(y0, y1); ax.set_aspect("equal", adjustable="box")
    fig.savefig(str(ruta), dpi=100)
    plt.close(fig)
