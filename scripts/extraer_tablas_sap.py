#!/usr/bin/env python3
"""
Extrae de SAP2000 todas las tablas necesarias para armar la memoria de calculo,
calcula las tablas derivadas (derivas, cortante 85 %, D/C, presiones en el suelo)
y genera los graficos. Opcionalmente captura pantallas de SAP2000.

REQUISITOS (PC con Windows donde corre SAP2000):
    pip install comtypes pandas openpyxl matplotlib pillow

USO TIPICO (SAP2000 abierto, modelo con analisis y disenos corridos):
    python extraer_tablas_sap.py --listar                 # revisar nombres del modelo
    python extraer_tablas_sap.py --salida v10/tablas      # extrae + derivados + graficos
    python extraer_tablas_sap.py --solo A                 # solo un grupo de tablas (A, B, C, F, M)
    python extraer_tablas_sap.py --solo F --salida v11/tablas   # solo las tablas faltantes
    python extraer_tablas_sap.py --solo-derivados         # recalcula derivados y graficos
                                                          # desde los Excel (sin SAP2000)
    python extraer_tablas_sap.py --capturas               # capturas guiadas de SAP2000

ARCHIVOS GENERADOS en --salida (cada uno con una hoja INDICE):
    tablas_casa_mr.xlsx           A  superestructura (modal, cortes, desplazamientos, D/C)
    CIMENTACION_reacciones.xlsx   B  suelo (reacciones en resortes, reaccion en la base)
    CIMENTACION_fuerzas.xlsx      C  refuerzo de cimentacion y muros
    CIMENTACION_faltantes.xlsx    F  reacciones bajo vigas, pilastras del subsuelo, vigas de coronacion,
                                     pilastrones y combos omega (pedestales y colectoras)
    modelo_definiciones.xlsx      M  definiciones del modelo (secciones, cargas, combos...)
    derivados_memoria.xlsx           tablas listas para la memoria (se calculan de las anteriores)
y en --figuras (por defecto <salida>/../figuras): PNG de los graficos.

El script NO modifica el modelo: solo cambia las unidades de presentacion a Tonf, m, C.
Si un nombre de caso, combo, grupo o tabla no existe en el modelo, lo avisa y sigue.
"""
import argparse
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

# ============================================================================
# CONFIGURACION (editar aqui)
# ============================================================================
TON_M_C = 12  # eUnits.Ton_m_C de la API de SAP2000

PARAMETROS = {
    "I": 1.0,                 # importancia
    "R": 4.5,                 # factor de reduccion
    "phiP": 0.9,              # irregularidad en planta
    "phiE": 0.9,              # irregularidad en elevacion
    "factor_deriva": 0.75,    # DeltaM = factor_deriva * R * DeltaE  (NEC-SE-DS 6.3.9)
    "deriva_lim": 0.02,       # limite de deriva inelastica (acero)
    "min_cortante": 0.85,     # cortante dinamico minimo / estatico (estructura irregular)
    "min_masa": 0.90,         # masa participativa minima
    "dc_adm": 0.95,           # relacion D/C admisible
    "qadm": 24.0,             # T/m2, capacidad admisible (combos estaticos)
    "qadm_sismo": 32.0,       # T/m2, con sismo (a confirmar por el geotecnista)
    "espectro": "NEC_C_Z040", # funcion de espectro de respuesta (en g)
}
NIVELES = [0.0, 3.06, 6.12, 7.77]   # cotas de los diafragmas para derivas (m)

CASOS_SISMO = ["SX", "SY", "SPECX", "SPECY"]
CASOS_SISMO_EXTRA = ["~TorsionSPECX", "~TorsionSPECY"]
CASOS_SUELO = ["CS1", "CS2", "CS3X", "CS3Y", "CS4X", "CS4Y"]
CASOS_SUELO_EST = ["CS1", "CS2"]
CASOS_REFUERZO = ["U1", "U2a", "U2b", "U3a", "U3b", "U5X", "U5Y", "U7X", "U7Y"]
CASOS_MUROS = ["U1", "U5X", "U5Y", "U7X", "U7Y"]
CASOS_PESO = ["DEAD", "MUERTA", "VIVA", "VIVA_CUB", "GRANIZO"]
CASOS_OMEGA = ["@OMEGA"]   # marcador: combos cuyo nombre contiene "omega" (U5Xomega, U7Yomega...)
FRAMES_PILASTRAS_SUB = [148, 149, 179, 402, 403, 404, 425]   # pilastras del tramo del subsuelo
FRAMES_VCOR = list(range(405, 425))                            # vigas de coronacion
FRAMES_PILASTRONES = [623, 632]                                # pilastrones B-4 y B-5
FRAMES_COLECTORAS = list(range(633, 656))                      # colectoras


@dataclass
class Job:
    hoja: str                   # nombre de la hoja de Excel
    patron: str                 # regex sobre el nombre de la tabla de SAP2000
    grupo: str = None           # filtro de grupo (None = sin filtro)
    casos: list = None          # casos/combos (None = todos)
    memoria: str = ""           # seccion de la memoria donde se usa
    nudos_de_areas: bool = False  # el grupo no tiene nudos propios (solo areas o frames): se lee la
                                  # tabla completa y se filtra por los nudos de sus objetos
    excluir_nudos_de: list = None  # con nudos_de_areas: grupos cuyos nudos se descartan (evita doble conteo)
    frames: list = None           # filtro por etiqueta de frame (columna Frame); si el grupo no existe en
                                  # el modelo, se lee sin grupo y se filtra solo por estos frames


ARCHIVOS = {
    "A": ("tablas_casa_mr.xlsx", [
        Job("Modal_Masas", r"^Modal Participating Mass Ratios$", None, ["MODAL"], "7.1"),
        Job("Modal_Periodos", r"^Modal Periods And Frequencies$", None, ["MODAL"], "7.1"),
        Job("Cortes_Seccion", r"^Section Cut Forces - Analysis$", None,
            CASOS_SISMO + CASOS_SISMO_EXTRA, "7.2"),
        Job("Desplazamientos", r"^Joint Displacements$", "DESPLAZAMIENTOS",
            CASOS_SISMO + CASOS_SISMO_EXTRA, "7.3"),
        Job("Desplazamientos_Todos", r"^Joint Displacements$", None, CASOS_SISMO, "7.3 (derivas)"),
        Job("Acero_DC", r"^Steel Design \d+ - Summary Data - AISC 360-16", None, None, "8.2, Anexo A"),
        Job("Reac_Base_Todas", r"^Base Reactions$", None, None, "5.5.5, 10.3"),
        Job("Masas_Grupos", r"^Groups 3 - Masses and Weights$", None, None, "5.5.5"),
    ]),
    "B": ("CIMENTACION_reacciones.xlsx", [
        Job("Reac_ZapAisl", r"^Joint Reactions$", "ZAP_AISLADAS", CASOS_SUELO, "10.1", True),
        Job("Reac_ZapCorr", r"^Joint Reactions$", "ZAP_CORRIDAS", CASOS_SUELO, "10.2", True),
        Job("Reac_Base", r"^Base Reactions$", None, CASOS_SUELO, "10.3"),
    ]),
    "C": ("CIMENTACION_fuerzas.xlsx", [
        Job("Shell_ZapAisl", r"^Element Forces - Area Shells$", "ZAP_AISLADAS", CASOS_REFUERZO, "10.1"),
        Job("Shell_ZapCorr", r"^Element Forces - Area Shells$", "ZAP_CORRIDAS", CASOS_REFUERZO, "10.2"),
        Job("Frame_Pedest", r"^Element Forces - Frames$", "PEDESTALES", CASOS_REFUERZO, "9.2"),
        Job("Frame_Pilastras", r"^Element Forces - Frames$", "PILASTRAS", CASOS_REFUERZO, "9.2"),
        Job("Frame_VigasCim", r"^Element Forces - Frames$", "VIGAS_CIMENTACION", CASOS_REFUERZO, "10"),
        Job("Conc_Pedest", r"^Concrete Design \d - Column Summary.*ACI 318-19",
            "PEDESTALES", None, "9.2"),
        Job("Conc_Pilastras", r"^Concrete Design \d - Column Summary.*ACI 318-19",
            "PILASTRAS", None, "9.2"),
        Job("Conc_VigasCim", r"^Concrete Design \d - Beam Summary.*ACI 318-19",
            "VIGAS_CIMENTACION", None, "10"),
        Job("Muros", r"^Element Forces - Area Shells$", "MUROS", CASOS_MUROS, "9.1"),
    ]),
    # Tablas que faltaban en la primera tanda: reacciones bajo vigas de cimentacion, pilastras del
    # subsuelo, vigas de coronacion, pilastrones, anclajes y colectoras (combos omega).
    # Las que dependen de un grupo (VCOR, COLECTORAS) funcionan tambien sin crearlo en SAP2000:
    # si el grupo no existe se lee la tabla completa y se filtra por la lista de frames.
    "F": ("CIMENTACION_faltantes.xlsx", [
        # Nudos con resorte fuera de las zapatas: vigas de cimentacion y pilastrones B-4/B-5 (en el modelo
        # REV10 son 167: 165 bajo VIGAS_CIMENTACION y 2 bajo los frames 623 y 632). Se descartan los nudos de
        # ZAP_* para que la suma con Reac_ZapAisl/Corr cierre con la reaccion total sin doble conteo.
        Job("Reac_VigasCim", r"^Joint Reactions$", "@RESORTES", CASOS_SUELO, "10 (control de reacciones)",
            True, ["ZAP_AISLADAS", "ZAP_CORRIDAS"]),
        Job("Conc_Pilastras_Sub", r"^Concrete Design \d - Column Summary.*ACI 318-19", None, None, "9.2",
            frames=FRAMES_PILASTRAS_SUB),
        Job("Conc_VCOR", r"^Concrete Design \d - Beam Summary.*ACI 318-19", "VCOR", None, "9.2",
            frames=FRAMES_VCOR),
        Job("Frame_Pilastrones", r"^Element Forces - Frames$", None, CASOS_REFUERZO, "10.7",
            frames=FRAMES_PILASTRONES),
        Job("Frame_Pedest_omega", r"^Element Forces - Frames$", "PEDESTALES", CASOS_OMEGA, "anclajes"),
        Job("Frame_Colect_omega", r"^Element Forces - Frames$", "COLECTORAS", CASOS_OMEGA, "colectoras",
            frames=FRAMES_COLECTORAS),
    ]),
    "M": ("modelo_definiciones.xlsx", [
        Job("Sec_Frames", r"^Frame Section Properties 01 - General$", None, None, "4.3"),
        Job("Asig_Frames", r"^Frame Section Assignments$", None, None, "4.3"),
        Job("Obj_Frames", r"^Objects And Elements - Frames$", None, None, "4.3"),
        Job("Sec_Areas", r"^Area Section Properties$", None, None, "4.2"),
        Job("Mat_General", r"^Material Properties 01 - General$", None, None, "2"),
        Job("Mat_Mecanicas", r"^Material Properties 02 - Basic Mechanical Properties$", None, None, "2"),
        Job("Mat_Acero", r"^Material Properties 03a - Steel Data$", None, None, "2"),
        Job("Mat_Hormigon", r"^Material Properties 03b - Concrete Data$", None, None, "2"),
        Job("Mat_Refuerzo", r"^Material Properties 03e - Rebar Data$", None, None, "2"),
        Job("Patrones_Carga", r"^Load Pattern Definitions$", None, None, "5"),
        Job("Casos_Carga", r"^Load Case Definitions$", None, None, "6"),
        Job("Combinaciones", r"^Combination Definitions$", None, None, "6"),
        Job("Masa_Sismica", r"^Mass Source$", None, None, "4.2, 5.5"),
        Job("Sismo_Estatico", r"^Auto Seismic - User Coefficient$", None, None, "5.5.5"),
        Job("Espectro", r"^Function - Response Spectrum - User$", None, None, "5.5.2"),
        Job("Caso_Espectral", r"^Case - Response Spectrum 1 - General$", None, None, "5.5.2"),
        Job("Caso_Modal", r"^Case - Modal 1 - General$", None, None, "7.1"),
        Job("Grupos_Def", r"^Groups 1 - Definitions$", None, None, "-"),
        Job("Grupos_Asignaciones", r"^Groups 2 - Assignments$", None, None, "-"),
        Job("Conectividad_Areas", r"^Connectivity - Area$", None, None, "10 (areas tributarias)"),
        Job("Coord_Nudos", r"^Joint Coordinates$", None, None, "7.3, 10"),
    ]),
}

# Capturas guiadas de SAP2000: (nombre de archivo, que dejar en pantalla)
CAPTURAS = [
    ("fig01_modelo_3d", "Vista 3D del modelo (extruido, sin etiquetas)"),
    ("fig02_planta_subsuelo", "Planta del subsuelo (Z = -3.06) con secciones/etiquetas"),
    ("fig02_planta_baja", "Planta baja (Z = 0.00)"),
    ("fig02_planta_alta", "Planta alta (Z = +3.06)"),
    ("fig02_planta_cubierta", "Cubierta (Z = +6.12)"),
    ("fig03_cargas_muerta", "Cargas de area MUERTA (Display > Show Load Assigns > Area)"),
    ("fig03_cargas_viva", "Cargas de area VIVA"),
    ("fig03_empuje_muros", "Empuje de tierras sobre los muros (EMPUJE)"),
    ("fig04_modo1", "Deformada modal: modo 1 (Display > Show Deformed Shape, caso MODAL)"),
    ("fig04_modo2", "Deformada modal: modo 2"),
    ("fig04_modo3", "Deformada modal: modo 3 (torsional)"),
    ("fig05_deformada_specx", "Deformada del caso SPECX"),
    ("fig05_deformada_specy", "Deformada del caso SPECY"),
    ("fig06_dc_acero", "Relaciones D/C del diseno de acero (Design > Steel Frame Design > Display)"),
    ("fig07_muros_M22", "Momentos M22 en muros (Display > Show Forces/Stresses > Shells), U5X"),
    ("fig07_muros_V23", "Cortante V23 en muros, U5X"),
    ("fig08_reacciones_zapatas", "Reacciones (F3) en resortes de zapatas, combo CS1"),
    ("fig08_zapatas_M11", "Momentos M11 en zapatas, combo U1"),
]


# ============================================================================
# Utilidades
# ============================================================================
def _etiqueta(x):
    """Normaliza una etiqueta (1, 1.0, '1') a texto comparable."""
    t = str(x).strip()
    return t[:-2] if t.endswith(".0") else t


def hoja_valida(nombre, usados):
    base = re.sub(r"[\[\]\*\?/\\:]", "_", nombre)[:31]
    nom, k = base, 2
    while nom in usados:
        nom = f"{base[:28]}_{k}"
        k += 1
    usados.add(nom)
    return nom


# ============================================================================
# Conexion con SAP2000 y lectura de tablas
# ============================================================================
def conectar():
    try:
        import comtypes.client
    except ImportError:
        sys.exit("Falta comtypes. Instala con: pip install comtypes")
    try:
        sap = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
    except OSError:
        sys.exit("No encuentro SAP2000 abierto. Abre el modelo y vuelve a correr.")
    modelo = sap.SapModel
    modelo.SetPresentUnits(TON_M_C)
    return modelo


def _lista(res):
    """Primera lista de un GetNameList: (n, (nombres...), ret)."""
    for item in res:
        if isinstance(item, (list, tuple)):
            return list(item)
    return []


def nombres_casos(m):
    return _lista(m.LoadCases.GetNameList())


def nombres_combos(m):
    return _lista(m.RespCombo.GetNameList())


def nombres_grupos(m):
    return _lista(m.GroupDef.GetNameList())


def tablas_disponibles(m):
    res = m.DatabaseTables.GetAvailableTables()
    listas = [x for x in res if isinstance(x, (list, tuple))]
    return list(listas[0])


def _a_dataframe(res):
    """Convierte la respuesta de GetTableFor{Display,Editing}Array en DataFrame.
    comtypes devuelve los parametros de salida y al final el codigo de retorno."""
    ret = res[-1]
    if ret != 0:
        raise RuntimeError(f"la API devolvio {ret}")
    cuerpo = res[:-1]
    secuencias = [list(x) if x is not None else [] for x in cuerpo
                  if x is None or isinstance(x, (list, tuple))]
    enteros = [x for x in cuerpo if isinstance(x, int)]
    if len(secuencias) < 2 or not enteros:
        raise RuntimeError(f"respuesta inesperada: {[type(x).__name__ for x in res]}")
    campos, datos, n = secuencias[-2], secuencias[-1], int(enteros[-1])
    if n == 0 or not campos:
        return pd.DataFrame(columns=campos)
    filas = [datos[i * len(campos):(i + 1) * len(campos)] for i in range(n)]
    df = pd.DataFrame(filas, columns=campos)
    for c in df.columns:  # numericos donde se pueda
        conv = pd.to_numeric(df[c], errors="coerce")
        if conv.notna().sum() == df[c].replace("", pd.NA).notna().sum():
            df[c] = conv
    return df


def leer_tabla(m, clave, grupo, casos):
    """DataFrame de la tabla `clave` filtrada por grupo y casos/combos."""
    db = m.DatabaseTables
    db.SetLoadCasesSelectedForDisplay(casos)
    db.SetLoadCombinationsSelectedForDisplay(casos)
    res = db.GetTableForDisplayArray(clave, [], grupo or "", 0, [], 0, [])
    try:
        return _a_dataframe(res)
    except RuntimeError as e1:
        try:  # tablas de definicion del modelo: lectura "de edicion"
            return _a_dataframe(db.GetTableForEditingArray(clave, grupo or "", 0, [], 0, []))
        except Exception:  # noqa: BLE001
            raise e1


def nudos_de_areas_del_grupo(m, grupo, todos):
    """Etiquetas de los nudos de los objetos (areas y frames) asignados a `grupo`.
    Los grupos de zapatas solo contienen areas y los de vigas solo frames, por eso las tablas
    de nudos (Joint Reactions...) salen vacias si se filtran por grupo: se filtran por estos nudos."""
    if grupo == "@RESORTES":   # seudo-grupo: todos los nudos con resorte de suelo
        sp = leer_tabla(m, "Joint Spring Assignments 1 - Uncoupled", None, todos)
        return {_etiqueta(j) for j in sp["Joint"].dropna()}
    asig = leer_tabla(m, "Groups 2 - Assignments", None, todos)
    g = asig[asig["GroupName"] == grupo]
    nudos = set()
    areas = {_etiqueta(a) for a in g.loc[g["ObjectType"] == "Area", "ObjectLabel"]}
    if areas:
        conn = leer_tabla(m, "Connectivity - Area", None, todos)
        conn = conn[conn["Area"].map(_etiqueta).isin(areas)]
        cols = [c for c in conn.columns if re.fullmatch(r"Joint\d+", c)]
        nudos |= {_etiqueta(j) for c in cols for j in conn[c].dropna() if str(j).strip() != ""}
    frames = {_etiqueta(f) for f in g.loc[g["ObjectType"] == "Frame", "ObjectLabel"]}
    if frames:   # elementos de analisis (incluye los nudos intermedios de la malla automatica)
        el = leer_tabla(m, "Objects And Elements - Frames", None, todos)
        el = el[el["FrameObject"].map(_etiqueta).isin(frames)]
        nudos |= {_etiqueta(j) for c in ("ElemJtI", "ElemJtJ") for j in el[c].dropna()}
    return nudos


def seleccion_casos(pedidos, existentes):
    if pedidos is None:
        return existentes
    expandidos = []
    for c in pedidos:
        if c == "@OMEGA":
            omega = [e for e in existentes if re.search(r"omega|\u03a9", e, re.I)]
            if not omega:
                print("    AVISO: el modelo no tiene combinaciones omega")
            expandidos += omega
        else:
            expandidos.append(c)
    faltan = [c for c in expandidos if c not in existentes]
    if faltan:
        print(f"    AVISO: no existen en el modelo: {', '.join(faltan)}")
    return [c for c in expandidos if c in existentes] or existentes


def listar(m):
    print("\n== TABLAS DISPONIBLES ==")
    for k in tablas_disponibles(m):
        print(f"  {k}")
    print("\n== CASOS ==\n ", ", ".join(nombres_casos(m)))
    print("\n== COMBOS ==\n ", ", ".join(nombres_combos(m)))
    print("\n== GRUPOS ==\n ", ", ".join(nombres_grupos(m)))


# ============================================================================
# Extraccion
# ============================================================================
def extraer(m, salida, letras, csv):
    salida.mkdir(parents=True, exist_ok=True)
    claves = tablas_disponibles(m)
    existentes = nombres_casos(m) + nombres_combos(m)
    grupos = nombres_grupos(m)
    resumen = []

    for letra in letras:
        archivo, trabajos = ARCHIVOS[letra]
        print(f"\n[{letra}] {archivo}")
        hojas, usados, indice = {}, set(), []
        for job in trabajos:
            def anotar(estado, tabla="", filas=0, hoja=job.hoja):
                fila = {"Hoja": hoja, "Tabla SAP2000": tabla, "Filas": filas,
                        "Grupo": job.grupo or "", "Frames": (f"{len(job.frames)} frames" if job.frames else ""), "Casos": "todos" if job.casos is None
                        else ", ".join(job.casos), "Memoria": job.memoria, "Estado": estado}
                indice.append(fila)
                resumen.append({"Archivo": archivo, **fila})

            grupo = job.grupo
            if grupo and not grupo.startswith("@") and grupo not in grupos:
                if job.frames:
                    print(f"  - {job.hoja}: el grupo {grupo} no existe; se filtra por frames "
                          f"{min(job.frames)}-{max(job.frames)} ({len(job.frames)})")
                    grupo = None
                else:
                    print(f"  - {job.hoja}: el grupo {grupo} no existe, se omite")
                    anotar("grupo inexistente")
                    continue
            tablas = [k for k in claves if re.search(job.patron, k, re.I)]
            if not tablas:
                print(f"  - {job.hoja}: ninguna tabla coincide con /{job.patron}/ "
                      f"(corre --listar y ajusta el patron)")
                anotar("tabla no encontrada")
                continue
            sel = seleccion_casos(job.casos, existentes)
            for clave in tablas:
                sufijo = re.search(r"Column|Beam|Joint", clave) if len(tablas) > 1 else None
                nombre = job.hoja if len(tablas) == 1 else \
                    f"{job.hoja}_{sufijo.group(0) if sufijo else clave[:12]}"
                try:
                    if job.nudos_de_areas:
                        nudos = nudos_de_areas_del_grupo(m, job.grupo, existentes)
                        for g_ex in job.excluir_nudos_de or []:
                            nudos -= nudos_de_areas_del_grupo(m, g_ex, existentes)
                        df = leer_tabla(m, clave, None, sel)
                        df = df[df["Joint"].map(_etiqueta).isin(nudos)]
                    else:
                        df = leer_tabla(m, clave, grupo, sel)
                    if job.frames and "Frame" in df.columns:
                        permitidos = {_etiqueta(f) for f in job.frames}
                        df = df[df["Frame"].map(_etiqueta).isin(permitidos)]
                        faltan = permitidos - {_etiqueta(f) for f in df["Frame"]}
                        if faltan:
                            lista = sorted(faltan, key=lambda x: (len(x), x))
                            resto = f" (+{len(lista) - 8} mas)" if len(lista) > 8 else ""
                            print(f"    AVISO: {len(lista)} frames sin resultados: {', '.join(lista[:8])}{resto}")
                except Exception as e:  # noqa: BLE001
                    print(f"  - {job.hoja}: error leyendo '{clave}': {e}")
                    anotar(f"error: {e}", clave)
                    continue
                if df.empty:
                    print(f"  - {job.hoja}: '{clave}' vacia (analisis/diseno sin correr?)")
                    anotar("vacia", clave)
                    continue
                hojas[hoja_valida(nombre, usados)] = df
                print(f"  - {job.hoja}: {len(df)} filas  <- {clave}")
                anotar("ok", clave, len(df))
                if csv:
                    df.to_csv(salida / f"{archivo[:-5]}__{nombre}.csv", index=False)
        if hojas:
            with pd.ExcelWriter(salida / archivo, engine="openpyxl") as xw:
                pd.DataFrame(indice).to_excel(xw, sheet_name="INDICE", index=False)
                for nom, df in hojas.items():
                    df.to_excel(xw, sheet_name=nom, index=False)
            print(f"  => guardado {salida / archivo}")
        else:
            print("  => sin datos, no se genero el archivo")

    log = pd.DataFrame(resumen)
    if not log.empty:
        log.to_csv(salida / "extraccion_log.csv", index=False, encoding="utf-8-sig")
        malos = log[log["Estado"] != "ok"]
        print(f"\nResumen: {len(log) - len(malos)} tablas ok, {len(malos)} con problemas "
              f"(detalle en {salida / 'extraccion_log.csv'})")


# ============================================================================
# Tablas derivadas (se calculan desde los Excel; no necesitan SAP2000)
# ============================================================================
def _leer(dt, archivo, hoja):
    p = Path(dt) / archivo
    if not p.exists():
        return None
    try:
        return pd.read_excel(p, sheet_name=hoja)
    except ValueError:
        return None


def _num(texto):
    """Ultimo numero dentro de un texto ('Z 0.10' -> 0.10)."""
    h = re.findall(r"-?\d+(?:\.\d+)?", str(texto))
    return float(h[-1]) if h else np.nan


def d_modal(dt):
    m = _leer(dt, "tablas_casa_mr.xlsx", "Modal_Masas")
    if m is None:
        return None
    cols = {"Modo": m.get("StepNum"), "T (s)": m["Period"],
            "UX (%)": m["UX"] * 100, "UY (%)": m["UY"] * 100}
    if "RZ" in m:
        cols["RZ (%)"] = m["RZ"] * 100
    cols["Suma UX (%)"] = m["SumUX"] * 100
    cols["Suma UY (%)"] = m["SumUY"] * 100
    out = pd.DataFrame(cols)
    ok = (m["SumUX"].iloc[-1] >= PARAMETROS["min_masa"]) and (m["SumUY"].iloc[-1] >= PARAMETROS["min_masa"])
    out.attrs["conclusion"] = (f"Masa acumulada final: X {m['SumUX'].iloc[-1]*100:.1f} %, "
                               f"Y {m['SumUY'].iloc[-1]*100:.1f} % -> "
                               f"{'cumple' if ok else 'NO cumple'} {PARAMETROS['min_masa']*100:.0f} %")
    return out


def d_cortante(dt):
    c = _leer(dt, "tablas_casa_mr.xlsx", "Cortes_Seccion")
    if c is None:
        return None
    cortes = c["SectionCut"].dropna().unique()
    base = min(cortes, key=lambda s: _num(s) if not np.isnan(_num(s)) else 1e9)
    b = c[c["SectionCut"] == base]
    filas = []
    for dire, est, din, comp in (("X", "SX", "SPECX", "F1"), ("Y", "SY", "SPECY", "F2")):
        e = b[b["OutputCase"] == est]
        d = b[b["OutputCase"] == din]
        if e.empty or d.empty:
            continue
        ve, vd = abs(e[comp]).max(), abs(d[comp]).max()
        vd_res = float(np.hypot(d["F1"], d["F2"]).max())
        filas.append({"Corte": base, "Direccion": dire, "V estatico (T)": ve,
                      "V dinamico (T)": vd, "Relacion (%)": vd / ve * 100,
                      "Minimo (%)": PARAMETROS["min_cortante"] * 100,
                      "Resultado": "Cumple" if vd / ve >= PARAMETROS["min_cortante"] else "NO cumple",
                      "V dinamico resultante (T)": vd_res,
                      "Factor de ajuste": max(1.0, PARAMETROS["min_cortante"] * ve / vd)})
    return pd.DataFrame(filas) if filas else None


def d_derivas(dt):
    d = _leer(dt, "tablas_casa_mr.xlsx", "Desplazamientos_Todos")
    fuente = "todos los nudos"
    if d is None:
        d = _leer(dt, "tablas_casa_mr.xlsx", "Desplazamientos")
        fuente = "solo grupo DESPLAZAMIENTOS (resultado parcial)"
    c = _leer(dt, "modelo_definiciones.xlsx", "Coord_Nudos")
    if d is None or c is None:
        return None
    c = c.assign(Joint=c["Joint"].map(_etiqueta))[["Joint", "GlobalX", "GlobalY", "GlobalZ"]]
    d = d.assign(Joint=d["Joint"].map(_etiqueta)).merge(c, on="Joint")
    d["xy"] = list(zip(d["GlobalX"].round(2), d["GlobalY"].round(2)))
    d["z"] = d["GlobalZ"].round(2)
    filas = []
    for caso in CASOS_SISMO:
        sub = d[d["OutputCase"] == caso]
        if sub.empty:
            continue
        comp = "U1" if caso.endswith("X") else "U2"
        dire = "X" if comp == "U1" else "Y"
        sub = sub.groupby(["Joint", "xy", "z"], as_index=False)[comp].agg(lambda s: s.abs().max())
        for k in range(1, len(NIVELES)):
            top = sub[sub["z"] == round(NIVELES[k], 2)]
            bot = sub[sub["z"] == round(NIVELES[k - 1], 2)]
            par = top.merge(bot, on="xy", suffixes=("_sup", "_inf"))
            if par.empty:
                continue
            h = NIVELES[k] - NIVELES[k - 1]
            par["dE"] = (par[f"{comp}_sup"] - par[f"{comp}_inf"]).abs() / h
            par["dM"] = PARAMETROS["factor_deriva"] * PARAMETROS["R"] * par["dE"]
            w = par.loc[par["dM"].idxmax()]
            filas.append({"Nivel (m)": NIVELES[k], "Dir": dire, "Caso": caso,
                          "Nudo sup": w["Joint_sup"], "Nudo inf": w["Joint_inf"],
                          "Desp. sup (mm)": w[f"{comp}_sup"] * 1000,
                          "Desp. inf (mm)": w[f"{comp}_inf"] * 1000,
                          "dE": w["dE"], "dM": w["dM"], "Limite": PARAMETROS["deriva_lim"],
                          "Resultado": "Cumple" if w["dM"] <= PARAMETROS["deriva_lim"] else "NO cumple",
                          "Pares evaluados": len(par)})
    if not filas:
        return None
    out = pd.DataFrame(filas)
    out.attrs["fuente"] = fuente
    return out


def d_acero(dt):
    a = _leer(dt, "tablas_casa_mr.xlsx", "Acero_DC")
    if a is None:
        return None
    g = a.loc[a.groupby(["DesignSect", "RatioType"])["Ratio"].idxmax()]
    cnt = a.groupby(["DesignSect", "RatioType"])["Frame"].nunique().rename("Elementos")
    g = g.merge(cnt.reset_index(), on=["DesignSect", "RatioType"])
    out = g[["DesignSect", "RatioType", "Elementos", "Ratio", "Combo"]].rename(
        columns={"DesignSect": "Seccion", "RatioType": "Verificacion", "Ratio": "D/C max",
                 "Combo": "Combinacion"})
    out["Resultado"] = np.where(out["D/C max"] <= PARAMETROS["dc_adm"], "Cumple", "NO cumple")
    return out.sort_values(["Seccion", "Verificacion"]).reset_index(drop=True)


def d_pesos(dt):
    b = _leer(dt, "tablas_casa_mr.xlsx", "Reac_Base_Todas")
    if b is None:
        return None
    b = b[b["OutputCase"].isin(CASOS_PESO)]
    if b.empty:
        return None
    return b[["OutputCase", "GlobalFX", "GlobalFY", "GlobalFZ"]].rename(
        columns={"OutputCase": "Caso", "GlobalFZ": "Fz total (T)"})


def d_presiones(dt):
    """Presion en el suelo por zapata: sum(F3) / sum(area) de los grupos ZAP_*.
    Las zapatas son los conjuntos de areas conectadas por nudos."""
    ra = _leer(dt, "CIMENTACION_reacciones.xlsx", "Reac_ZapAisl")
    rc = _leer(dt, "CIMENTACION_reacciones.xlsx", "Reac_ZapCorr")
    conn = _leer(dt, "modelo_definiciones.xlsx", "Conectividad_Areas")
    asig = _leer(dt, "modelo_definiciones.xlsx", "Grupos_Asignaciones")
    coord = _leer(dt, "modelo_definiciones.xlsx", "Coord_Nudos")
    reac = [r for r in (ra, rc) if r is not None and not r.empty and "Joint" in r.columns]
    if not reac or conn is None or asig is None or coord is None:
        return None
    r = pd.concat(reac, ignore_index=True).drop_duplicates(
        subset=[c for c in ("Joint", "OutputCase", "StepType") if c in reac[0].columns])
    r = r.assign(Joint=r["Joint"].map(_etiqueta))

    areas = {_etiqueta(a) for a in asig.loc[asig["GroupName"].isin(["ZAP_AISLADAS", "ZAP_CORRIDAS"])
                                            & (asig["ObjectType"] == "Area"), "ObjectLabel"]}
    conn = conn[conn["Area"].map(_etiqueta).isin(areas)]
    jcols = [c for c in conn.columns if re.fullmatch(r"Joint\d+", c)]
    padre = {}

    def raiz(x):
        padre.setdefault(x, x)
        while padre[x] != x:
            padre[x] = padre[padre[x]]
            x = padre[x]
        return x

    trib, area_de_joint, area_total = {}, {}, {}
    for _, f in conn.iterrows():
        js = [_etiqueta(f[c]) for c in jcols if pd.notna(f[c]) and str(f[c]).strip() != ""]
        a = float(f["AreaArea"])
        for j in js:
            trib[j] = trib.get(j, 0.0) + a / len(js)
            raiz(js[0])
            padre[raiz(j)] = raiz(js[0])
        area_de_joint[_etiqueta(f["Area"])] = js[0]
        area_total[_etiqueta(f["Area"])] = a
    zap_de_joint = {j: raiz(j) for j in trib}
    zap_ids = {rr: f"Z{n + 1:02d}" for n, rr in enumerate(sorted(set(zap_de_joint.values()), key=int))}
    coord = coord.assign(Joint=coord["Joint"].map(_etiqueta)).set_index("Joint")

    r = r[r["Joint"].isin(trib)].copy()
    r["Zapata"] = r["Joint"].map(lambda j: zap_ids[zap_de_joint[j]])
    r["A_trib"] = r["Joint"].map(trib)
    r["p_nodal"] = r["F3"] / r["A_trib"]

    # Las combinaciones con sismo traen dos filas por nudo (StepType Max y Min): no se
    # deben sumar. Max gobierna la presion; Min gobierna el levantamiento.
    r["Paso"] = r["StepType"].fillna("") if "StepType" in r.columns else ""
    filas = []
    for (zid, caso, paso), g in r.groupby(["Zapata", "OutputCase", "Paso"]):
        js = g["Joint"].unique()
        cx = coord.loc[[j for j in js if j in coord.index], "GlobalX"].astype(float).mean()
        cy = coord.loc[[j for j in js if j in coord.index], "GlobalY"].astype(float).mean()
        a_tot = sum(trib[j] for j in js)
        lim = PARAMETROS["qadm"] if caso in CASOS_SUELO_EST else PARAMETROS["qadm_sismo"]
        p_prom = g["F3"].sum() / a_tot
        evalua = paso in ("", "Max")
        filas.append({"Zapata": zid, "Caso": caso, "Paso": paso, "Nudos": len(js),
                      "Area (m2)": a_tot, "X (m)": cx, "Y (m)": cy,
                      "F3 total (T)": g["F3"].sum(), "p prom (T/m2)": p_prom,
                      "p max nodal (T/m2)": g["p_nodal"].max(),
                      "p min nodal (T/m2)": g["p_nodal"].min(),
                      "Limite (T/m2)": lim,
                      "Resultado": ("Cumple" if p_prom <= lim else "NO cumple") if evalua else "-",
                      "Nudos con levantamiento": int((g["F3"] < 0).sum())})
    return pd.DataFrame(filas)


def d_control_reacciones(dt):
    """Equilibrio: suma de reacciones de los resortes (zapatas + vigas de cimentacion) contra la
    reaccion total en la base (Base Reactions), por combinacion y paso."""
    ra = _leer(dt, "CIMENTACION_reacciones.xlsx", "Reac_ZapAisl")
    rc = _leer(dt, "CIMENTACION_reacciones.xlsx", "Reac_ZapCorr")
    rb = _leer(dt, "CIMENTACION_reacciones.xlsx", "Reac_Base")
    rv = _leer(dt, "CIMENTACION_faltantes.xlsx", "Reac_VigasCim")
    zap = [r for r in (ra, rc) if r is not None and not r.empty]
    if not zap or rb is None:
        return None
    z = pd.concat(zap, ignore_index=True)
    z["StepType"] = z["StepType"].fillna("") if "StepType" in z.columns else ""
    z = z.drop_duplicates(subset=["Joint", "OutputCase", "StepType"])
    suma = lambda df: df.groupby(["OutputCase", "StepType"])[["F1", "F2", "F3"]].sum()
    out = suma(z).add_suffix(" zapatas")
    if rv is not None and not rv.empty:
        rv = rv.assign(StepType=rv["StepType"].fillna("") if "StepType" in rv.columns else "")
        rv = rv.drop_duplicates(subset=["Joint", "OutputCase", "StepType"])
        out = out.join(suma(rv).add_suffix(" vigas"), how="left").fillna(0.0)
    else:
        for c in ("F1", "F2", "F3"):
            out[f"{c} vigas"] = 0.0
    for c in ("F1", "F2", "F3"):
        out[f"{c} total"] = out[f"{c} zapatas"] + out[f"{c} vigas"]
    rb = rb.assign(StepType=rb["StepType"].fillna(""))
    rb = rb.set_index(["OutputCase", "StepType"])[["GlobalFX", "GlobalFY", "GlobalFZ"]]
    out = out.join(rb, how="left").reset_index()
    out["Dif FZ (T)"] = out["F3 total"] - out["GlobalFZ"]
    out["Dif FZ (%)"] = out["Dif FZ (T)"] / out["GlobalFZ"] * 100
    out["Dif FX (T)"] = out["F1 total"] - out["GlobalFX"]
    out["Dif FY (T)"] = out["F2 total"] - out["GlobalFY"]
    # en las combinaciones con sismo (pasos Max/Min) la suma de maximos nodales no es el maximo de la
    # suma: solo las combinaciones sin paso son comparables con la reaccion total de la base
    out["Comparable"] = np.where(out["StepType"] == "", "si", "no (suma de envolventes nodales)")
    out.attrs["conclusion"] = ("incluye las vigas de cimentacion" if rv is not None and not rv.empty
                               else "SIN reacciones bajo vigas (falta CIMENTACION_faltantes.xlsx)")
    return out.rename(columns={"OutputCase": "Caso", "StepType": "Paso"})


def derivados(dt):
    """Calcula las tablas derivadas y las guarda en derivados_memoria.xlsx."""
    dt = Path(dt)
    calculos = {"Modal": d_modal, "Cortante": d_cortante, "Derivas": d_derivas,
                "Acero_resumen": d_acero, "Pesos": d_pesos, "Presion_zapatas": d_presiones,
                "Control_reacciones": d_control_reacciones}
    res = {}
    print("\nTablas derivadas:")
    for nombre, fn in calculos.items():
        try:
            df = fn(dt)
        except Exception as e:  # noqa: BLE001
            print(f"  - {nombre}: error: {e}")
            continue
        if df is None or df.empty:
            print(f"  - {nombre}: faltan tablas de origen, se omite")
            continue
        res[nombre] = df
        extra = df.attrs.get("conclusion") or (f"[derivas con {df.attrs['fuente']}]"
                                               if "fuente" in df.attrs else "")
        print(f"  - {nombre}: {len(df)} filas {extra}")
    if not res:
        return res
    if "Presion_zapatas" in res:  # resumen por caso
        p = res["Presion_zapatas"]
        mx = p[p["Paso"].isin(["", "Max"])].groupby("Caso").agg(
            **{"p prom max (T/m2)": ("p prom (T/m2)", "max"),
               "p max nodal (T/m2)": ("p max nodal (T/m2)", "max"),
               "Zapatas que no cumplen": ("Resultado", lambda s: int((s == "NO cumple").sum()))})
        mn = p[p["Paso"].isin(["", "Min"])].groupby("Caso").agg(
            **{"p min nodal (T/m2)": ("p min nodal (T/m2)", "min"),
               "Nudos con levantamiento": ("Nudos con levantamiento", "sum")})
        res["Presion_resumen"] = mx.join(mn).reset_index()
    par = pd.DataFrame({"Parametro": list(PARAMETROS), "Valor": list(PARAMETROS.values())})
    res["Parametros"] = par
    with pd.ExcelWriter(dt / "derivados_memoria.xlsx", engine="openpyxl") as xw:
        for nom, df in res.items():
            df.to_excel(xw, sheet_name=nom[:31], index=False)
    print(f"  => guardado {dt / 'derivados_memoria.xlsx'}")
    return res


# ============================================================================
# Graficos
# ============================================================================
def figuras(dt, carpeta, res):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        print("\nFalta matplotlib: pip install matplotlib (se omiten los graficos)")
        return
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    AZUL, ROJO, GRIS, VERDE = "#1f4e79", "#c0392b", "#7f8c8d", "#2e8b57"
    print("\nGraficos:")

    def guardar(fig, nombre):
        fig.tight_layout()
        fig.savefig(carpeta / nombre, dpi=200)
        plt.close(fig)
        print(f"  - {nombre}")

    # Espectro
    e = _leer(dt, "modelo_definiciones.xlsx", "Espectro")
    if e is not None:
        e = e[e["Name"] == PARAMETROS["espectro"]].sort_values("Period")
        if not e.empty:
            f = PARAMETROS
            red = f["I"] / (f["R"] * f["phiP"] * f["phiE"])
            fig, ax = plt.subplots(figsize=(8, 4.5))
            ax.plot(e["Period"], e["Accel"], color=AZUL, lw=2, label="Espectro elastico")
            ax.plot(e["Period"], e["Accel"] * red, color=ROJO, lw=2,
                    label=f"Espectro de diseno (x I/(R*phiP*phiE) = {red:.3f})")
            ax.set(xlabel="T (s)", ylabel="Sa (g)", title="Espectro NEC-SE-DS")
            ax.grid(alpha=.3); ax.legend()
            guardar(fig, "graf_espectro.png")

    # Masa modal
    if "Modal" in res:
        m = res["Modal"]
        fig, ax = plt.subplots(figsize=(8, 4.5))
        ax.plot(range(1, len(m) + 1), m["Suma UX (%)"], color=AZUL, lw=2, label="Suma UX")
        ax.plot(range(1, len(m) + 1), m["Suma UY (%)"], color=ROJO, lw=2, label="Suma UY")
        ax.axhline(PARAMETROS["min_masa"] * 100, color=GRIS, ls="--", label="90 %")
        ax.set(xlabel="Modo", ylabel="Masa participativa acumulada (%)", ylim=(0, 105),
               title="Participacion modal")
        ax.grid(alpha=.3); ax.legend()
        guardar(fig, "graf_masa_modal.png")

    # Cortante
    if "Cortante" in res:
        c = res["Cortante"]
        x = np.arange(len(c))
        fig, ax = plt.subplots(figsize=(7, 4.5))
        ax.bar(x - .2, c["V estatico (T)"], .4, color=GRIS, label="Estatico")
        ax.bar(x + .2, c["V dinamico (T)"], .4, color=AZUL, label="Dinamico")
        for i, f in c.iterrows():
            ax.plot([i - .4, i + .4], [f["V estatico (T)"] * PARAMETROS["min_cortante"]] * 2,
                    color=ROJO, ls="--", label="85 % del estatico" if i == 0 else None)
        ax.set_xticks(x); ax.set_xticklabels(c["Direccion"])
        ax.set(ylabel="Cortante basal (T)", title="Cortante estatico vs dinamico")
        ax.legend()
        guardar(fig, "graf_cortante.png")

    # Derivas
    if "Derivas" in res:
        d = res["Derivas"]
        et = [f"+{r['Nivel (m)']:.2f} {r['Caso']}" for _, r in d.iterrows()]
        fig, ax = plt.subplots(figsize=(8, max(3.5, .35 * len(d) + 1.5)))
        col = [VERDE if r == "Cumple" else ROJO for r in d["Resultado"]]
        ax.barh(et, d["dM"], color=col)
        ax.axvline(PARAMETROS["deriva_lim"], color=ROJO, ls="--", label=f"Limite {PARAMETROS['deriva_lim']}")
        ax.invert_yaxis()
        ax.set(xlabel="Deriva inelastica dM", title="Derivas de piso")
        ax.legend()
        guardar(fig, "graf_derivas.png")

    # D/C acero
    if "Acero_resumen" in res:
        a = res["Acero_resumen"]
        et = [f"{r['Seccion']} ({r['Verificacion']})" for _, r in a.iterrows()]
        fig, ax = plt.subplots(figsize=(8, max(3.5, .35 * len(a) + 1.5)))
        col = [AZUL if r <= PARAMETROS["dc_adm"] else ROJO for r in a["D/C max"]]
        ax.barh(et, a["D/C max"], color=col)
        ax.axvline(PARAMETROS["dc_adm"], color=ROJO, ls="--", label=f"Admisible {PARAMETROS['dc_adm']}")
        ax.invert_yaxis()
        ax.set(xlabel="D/C maximo", title="Relacion demanda/capacidad por seccion")
        ax.legend()
        guardar(fig, "graf_dc_acero.png")

    # Presiones en el suelo
    if "Presion_zapatas" in res:
        p = res["Presion_zapatas"]
        p = p[p["Paso"].isin(["", "Max"])]
        mx = p.groupby("Caso")["p prom (T/m2)"].max()
        fig, ax = plt.subplots(figsize=(7, 4.5))
        ax.bar(mx.index, mx.values, color=AZUL)
        ax.axhline(PARAMETROS["qadm"], color=GRIS, ls="--", label=f"qadm {PARAMETROS['qadm']} T/m2")
        ax.axhline(PARAMETROS["qadm_sismo"], color=ROJO, ls="--",
                   label=f"qadm sismo {PARAMETROS['qadm_sismo']} T/m2")
        ax.set(ylabel="Presion promedio maxima (T/m2)", title="Presion en el suelo por combinacion")
        ax.legend()
        guardar(fig, "graf_presion_combos.png")

        env = p[p["Caso"].isin(CASOS_SUELO)].groupby("Zapata").agg(
            x=("X (m)", "first"), y=("Y (m)", "first"), a=("Area (m2)", "first"),
            pm=("p prom (T/m2)", "max"))
        fig, ax = plt.subplots(figsize=(8, 7))
        sc = ax.scatter(env["x"], env["y"], s=40 + 25 * env["a"], c=env["pm"], cmap="viridis")
        for z, r in env.iterrows():
            ax.annotate(f"{z}\n{r['pm']:.1f}", (r["x"], r["y"]), fontsize=6, ha="center", va="center")
        fig.colorbar(sc, label="Presion promedio maxima (T/m2)")
        ax.set(xlabel="X (m)", ylabel="Y (m)", aspect="equal",
               title="Mapa de presiones en zapatas (envolvente CS)")
        ax.grid(alpha=.3)
        guardar(fig, "graf_mapa_presiones.png")


# ============================================================================
# Capturas de pantalla de SAP2000 (solo Windows)
# ============================================================================
def _ventana_sap():
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:  # noqa: BLE001
        user32.SetProcessDPIAware()
    halladas = []

    def cb(h, _):
        if user32.IsWindowVisible(h):
            n = user32.GetWindowTextLengthW(h)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(h, buf, n + 1)
                if "SAP2000" in buf.value:
                    r = wintypes.RECT()
                    user32.GetWindowRect(h, ctypes.byref(r))
                    halladas.append(((r.right - r.left) * (r.bottom - r.top), h, r))
        return True

    proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(proc(cb), 0)
    if not halladas:
        return None
    _, h, r = max(halladas, key=lambda t: t[0])
    user32.SetForegroundWindow(h)
    return (r.left, r.top, r.right, r.bottom)


def capturas(m, carpeta, solo):
    if sys.platform != "win32":
        sys.exit("Las capturas solo funcionan en Windows.")
    try:
        from PIL import ImageGrab
    except ImportError:
        sys.exit("Falta pillow: pip install pillow")
    carpeta = Path(carpeta) / "sap"
    carpeta.mkdir(parents=True, exist_ok=True)
    lista = [c for c in CAPTURAS if not solo or c[0] in solo]
    print("\nCAPTURAS GUIADAS. Para cada figura: deja en SAP2000 la vista indicada y\n"
          "vuelve a esta ventana. [Enter] captura, [s] salta, [q] termina.\n"
          "La captura toma la ventana completa de SAP2000 (se puede recortar luego).\n")
    for nombre, texto in lista:
        resp = input(f"-> {nombre}: {texto}\n   Prepara la vista y presiona Enter... ").strip().lower()
        if resp == "q":
            break
        if resp == "s":
            continue
        try:
            m.View.RefreshView(0, False)
        except Exception:  # noqa: BLE001
            pass
        print("   Tienes 3 s para dar clic en SAP2000...")
        time.sleep(3)
        caja = _ventana_sap()
        if not caja:
            print("   No encontre la ventana de SAP2000; se omite.")
            continue
        time.sleep(0.5)
        ImageGrab.grab(bbox=caja, all_screens=True).save(carpeta / f"{nombre}.png")
        print(f"   guardada {carpeta / (nombre + '.png')}")


# ============================================================================
# Programa principal
# ============================================================================
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--salida", default="v10/tablas", help="carpeta de los Excel")
    ap.add_argument("--figuras", help="carpeta de graficos (por defecto <salida>/../figuras)")
    ap.add_argument("--solo", choices=sorted(ARCHIVOS), help="solo un grupo de tablas (A, B, C, M)")
    ap.add_argument("--listar", action="store_true", help="listar tablas/casos/grupos del modelo")
    ap.add_argument("--csv", action="store_true", help="guardar tambien cada tabla en CSV")
    ap.add_argument("--sin-figuras", action="store_true", help="no generar derivados ni graficos")
    ap.add_argument("--solo-derivados", action="store_true",
                    help="recalcular derivados y graficos desde los Excel (sin SAP2000)")
    ap.add_argument("--capturas", action="store_true", help="capturas guiadas de SAP2000")
    ap.add_argument("--capturas-solo", help="nombres de capturas separados por coma")
    args = ap.parse_args()

    salida = Path(args.salida)
    carpeta_fig = Path(args.figuras) if args.figuras else salida.parent / "figuras"

    if args.solo_derivados:
        figuras(salida, carpeta_fig, derivados(salida))
        return

    m = conectar()
    if args.listar:
        listar(m)
        return
    if args.capturas:
        capturas(m, carpeta_fig, args.capturas_solo.split(",") if args.capturas_solo else None)
        return
    extraer(m, salida, [args.solo] if args.solo else list(ARCHIVOS), args.csv)
    if args.solo:
        print("\nCon --solo no se recalculan los derivados ni los graficos "
              "(para hacerlo: python extraer_tablas_sap.py --solo-derivados --salida <carpeta>).")
    elif not args.sin_figuras:
        figuras(salida, carpeta_fig, derivados(salida))


if __name__ == "__main__":
    main()
