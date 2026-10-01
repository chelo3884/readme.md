#!/usr/bin/env python3
"""
Extrae de un modelo SAP2000 abierto las tablas que se usan para revisar la
casa (superestructura con base flexible, suelo, refuerzo de cimentacion, muros)
y las guarda en tres archivos Excel.

REQUISITOS (en tu PC con Windows, donde esta SAP2000):
    pip install comtypes pandas openpyxl

USO
    1. Abre el modelo en SAP2000, con el analisis corrido (candado cerrado).
       Para las tablas de diseno: Design > Steel Frame Design / Concrete Frame
       Design > Start Design con los combos que quieras (U1 ... U7Y).
    2. Ejecuta (desde la carpeta del repo):
           python scripts/extraer_tablas_sap.py --salida v10/tablas
    3. Opciones utiles:
           --listar          imprime las tablas, casos, combos y grupos del modelo
                             (usalo primero si algun nombre no coincide)
           --solo A          solo el archivo A, B o C (ver ARCHIVOS abajo)
           --csv             ademas guarda cada tabla como CSV

El script NO modifica el modelo: solo cambia las unidades de presentacion a
Tonf, m, C para la lectura de tablas.

Si un nombre de caso, combo, grupo o tabla de la configuracion no existe en tu
modelo, el script lo avisa y sigue con lo demas; corrige la configuracion de
abajo (CASOS_*, GRUPOS_*, PATRON_*) con lo que muestre --listar.
"""
import argparse
import re
import sys
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------------
# CONFIGURACION (editar aqui)
# ----------------------------------------------------------------------------
TON_M_C = 12  # eUnits.Ton_m_C de la API de SAP2000

CASOS_SISMO = ["SX", "SY", "SPECX", "SPECY"]
CASOS_SUELO = ["CS1", "CS2", "CS3X", "CS3Y", "CS4X", "CS4Y"]
CASOS_REFUERZO = ["U1", "U2a", "U2b", "U3a", "U3b", "U5X", "U5Y", "U7X", "U7Y"]
CASOS_MUROS = ["U1", "U5X", "U5Y", "U7X", "U7Y"]

# Una tabla = (hoja, patron regex sobre el nombre de la tabla, grupo, casos)
#   grupo = None  -> sin filtro de grupo
#   casos = None  -> todos los casos y combos (para tablas de diseno)
# Un patron puede coincidir con varias tablas: se exporta cada una en su hoja.
ARCHIVOS = {
    # A. Superestructura con base flexible
    "A": ("tablas_casa_mr.xlsx", [
        ("Modal_Masas", r"^Modal Participating Mass Ratios$", None, ["MODAL"]),
        ("Cortes_Seccion", r"^Section Cut Forces - Analysis$", None, CASOS_SISMO),
        ("Desplazamientos", r"^Joint Displacements$", "DESPLAZAMIENTOS", CASOS_SISMO),
        ("Acero_DC", r"^Steel Design.*Summary.*AISC 360-16", None, None),
    ]),
    # B. Suelo
    "B": ("CIMENTACION_reacciones.xlsx", [
        ("Reac_ZapAisl", r"^Joint Reactions$", "ZAP_AISLADAS", CASOS_SUELO),
        ("Reac_ZapCorr", r"^Joint Reactions$", "ZAP_CORRIDAS", CASOS_SUELO),
        ("Reac_Base", r"^Base Reactions$", None, CASOS_SUELO),
    ]),
    # C + D. Refuerzo de cimentacion y muros
    "C": ("CIMENTACION_fuerzas.xlsx", [
        ("Shell_ZapAisl", r"^Element Forces - Area Shells$", "ZAP_AISLADAS", CASOS_REFUERZO),
        ("Shell_ZapCorr", r"^Element Forces - Area Shells$", "ZAP_CORRIDAS", CASOS_REFUERZO),
        ("Frame_Pedest", r"^Element Forces - Frames$", "PEDESTALES", CASOS_REFUERZO),
        ("Frame_Pilastras", r"^Element Forces - Frames$", "PILASTRAS", CASOS_REFUERZO),
        ("Frame_VigasCim", r"^Element Forces - Frames$", "VIGAS_CIMENTACION", CASOS_REFUERZO),
        ("Conc_Pedest", r"^Concrete Design \d - (Column|Beam) Summary.*ACI 318-19", "PEDESTALES", None),
        ("Conc_Pilastras", r"^Concrete Design \d - (Column|Beam) Summary.*ACI 318-19", "PILASTRAS", None),
        ("Conc_VigasCim", r"^Concrete Design \d - (Column|Beam) Summary.*ACI 318-19", "VIGAS_CIMENTACION", None),
        ("Muros", r"^Element Forces - Area Shells$", "MUROS", CASOS_MUROS),
    ]),
}


# ----------------------------------------------------------------------------
# Conexion con SAP2000
# ----------------------------------------------------------------------------
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
    """Extrae la lista de nombres de un GetNameList: (n, (nombres...), ret)."""
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
    # (NumberTables, TableKey[], TableName[], ImportType[], ret)
    listas = [x for x in res if isinstance(x, (list, tuple))]
    return list(listas[0]), list(listas[1]) if len(listas) > 1 else []


# ----------------------------------------------------------------------------
# Lectura de una tabla
# ----------------------------------------------------------------------------
def leer_tabla(m, clave, grupo, casos):
    """Devuelve un DataFrame con la tabla `clave` filtrada por grupo y casos."""
    db = m.DatabaseTables
    db.SetLoadCasesSelectedForDisplay(casos)
    db.SetLoadCombinationsSelectedForDisplay(casos)
    res = db.GetTableForDisplayArray(clave, [], grupo or "", 0, [], 0, [])
    # comtypes devuelve: (TableVersion, Campos[], NumRegistros, Datos[], ret)
    ret = res[-1]
    if ret != 0:
        raise RuntimeError(f"GetTableForDisplayArray devolvio {ret}")
    campos, n, datos = list(res[1]), int(res[2]), list(res[3])
    if n == 0 or not campos:
        return pd.DataFrame(columns=campos)
    filas = [datos[i * len(campos):(i + 1) * len(campos)] for i in range(n)]
    df = pd.DataFrame(filas, columns=campos)
    for c in df.columns:  # numericos donde se pueda
        conv = pd.to_numeric(df[c], errors="coerce")
        if conv.notna().sum() == df[c].replace("", pd.NA).notna().sum():
            df[c] = conv
    return df


def seleccion_casos(m, pedidos, existentes):
    """Casos/combos pedidos que existen; avisa de los que no."""
    if pedidos is None:
        return existentes
    faltan = [c for c in pedidos if c not in existentes]
    if faltan:
        print(f"    AVISO: no existen en el modelo: {', '.join(faltan)}")
    return [c for c in pedidos if c in existentes] or existentes


def hoja_valida(nombre, usados):
    base = re.sub(r"[\[\]\*\?/\\:]", "_", nombre)[:31]
    nom, k = base, 2
    while nom in usados:
        nom = f"{base[:28]}_{k}"
        k += 1
    usados.add(nom)
    return nom


# ----------------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------------
def listar(m):
    claves, nombres = tablas_disponibles(m)
    print("\n== TABLAS DISPONIBLES ==")
    for k, n in zip(claves, nombres or claves):
        print(f"  {k}")
    print("\n== CASOS ==\n ", ", ".join(nombres_casos(m)))
    print("\n== COMBOS ==\n ", ", ".join(nombres_combos(m)))
    print("\n== GRUPOS ==\n ", ", ".join(nombres_grupos(m)))


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--salida", default="v10/tablas", help="carpeta de salida")
    ap.add_argument("--solo", choices=sorted(ARCHIVOS), help="solo un archivo")
    ap.add_argument("--listar", action="store_true", help="listar tablas/casos/grupos")
    ap.add_argument("--csv", action="store_true", help="guardar tambien CSV")
    args = ap.parse_args()

    m = conectar()
    if args.listar:
        listar(m)
        return

    salida = Path(args.salida)
    salida.mkdir(parents=True, exist_ok=True)
    claves, _ = tablas_disponibles(m)
    existentes = nombres_casos(m) + nombres_combos(m)
    grupos = nombres_grupos(m)

    for letra, (archivo, trabajos) in ARCHIVOS.items():
        if args.solo and letra != args.solo:
            continue
        print(f"\n[{letra}] {archivo}")
        hojas, usados = {}, set()
        for hoja, patron, grupo, casos in trabajos:
            if grupo and grupo not in grupos:
                print(f"  - {hoja}: el grupo {grupo} no existe, se omite")
                continue
            tablas = [k for k in claves if re.search(patron, k, re.I)]
            if not tablas:
                print(f"  - {hoja}: ninguna tabla coincide con /{patron}/ "
                      f"(corre --listar y ajusta el patron)")
                continue
            sel = seleccion_casos(m, casos, existentes)
            for clave in tablas:
                sufijo = re.search(r"Column|Beam|Joint", clave)
                nombre = hoja if len(tablas) == 1 else f"{hoja}_{sufijo.group(0) if sufijo else clave[:12]}"
                try:
                    df = leer_tabla(m, clave, grupo, sel)
                except Exception as e:  # noqa: BLE001
                    print(f"  - {hoja}: error leyendo '{clave}': {e}")
                    continue
                if df.empty:
                    print(f"  - {hoja}: '{clave}' vacia (analisis/diseno sin correr?)")
                    continue
                hojas[hoja_valida(nombre, usados)] = df
                print(f"  - {hoja}: {len(df)} filas  <- {clave}")
                if args.csv:
                    df.to_csv(salida / f"{archivo[:-5]}__{nombre}.csv", index=False)
        if hojas:
            with pd.ExcelWriter(salida / archivo, engine="openpyxl") as xw:
                for nom, df in hojas.items():
                    df.to_excel(xw, sheet_name=nom, index=False)
            print(f"  => guardado {salida / archivo}")
        else:
            print("  => sin datos, no se genero el archivo")


if __name__ == "__main__":
    main()
