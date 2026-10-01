#!/usr/bin/env python3
"""
Calculos de la memoria Rev. 1: muros de subsuelo, cimentacion y elementos de
hormigon, a partir de las tablas extraidas de SAP2000 (v10/tablas/*.xlsx).

    python scripts/calculo_r1.py --tablas v10/tablas --salida v10/tablas

Genera calculo_r1.xlsx (todas las tablas de resultados) y calculo_r1.json.
Unidades: Tonf, m (momentos T*m/m, cortantes T/m) y cm2 / cm donde se indica.
Criterios: ACI 318-19, f'c = 280 kgf/cm2, fy = 4200 kgf/cm2.
"""
import argparse
import json
import re
from pathlib import Path

import numpy as np
import pandas as pd

FC, FY = 280.0, 4200.0            # kgf/cm2
EC_LAMBDA = 1.0
PHI_F, PHI_V = 0.90, 0.75
FC_MPA = FC * 0.0980665            # 27.46 MPa
MU_FRIC = np.tan(np.radians(24.0))  # coeficiente de friccion suelo-hormigon
KV = 4800.0                        # T/m3

# ---------------------------------------------------------------- utilidades
def _et(x):
    t = str(x).strip()
    return t[:-2] if t.endswith(".0") else t


def leer(dt, archivo, hoja):
    return pd.read_excel(Path(dt) / archivo, sheet_name=hoja)


def as_req(mu, d, b=100.0):
    """As requerido (cm2) para Mu (T*m) en ancho b (cm) y d (cm); phi = 0.9."""
    mu = np.asarray(mu, dtype=float) * 1e5               # kgf*cm
    k = 2 * mu / (PHI_F * 0.85 * FC * b * d * d)
    with np.errstate(invalid="ignore"):
        r = 1 - np.sqrt(1 - k)
    return np.where(k >= 1, np.inf, 0.85 * FC * b * d / FY * r)


def phi_mn(a_s, d, b=100.0):
    """Capacidad phi*Mn (T*m) para As (cm2) en ancho b (cm)."""
    a = a_s * FY / (0.85 * FC * b)
    return PHI_F * a_s * FY * (d - a / 2) / 1e5


def barra_as(diam_mm, sep_cm=None):
    a1 = np.pi * (diam_mm / 10) ** 2 / 4
    return a1 if sep_cm is None else a1 * 100 / sep_cm


def phi_vc_muro(rho, d_cm, b=100.0):
    """phi*Vc (T) ACI 318-19 Tabla 22.5.5.1(c), sin refuerzo transversal, Nu = 0."""
    d_mm = d_cm * 10
    lam_s = min(1.0, np.sqrt(2 / (1 + 0.004 * d_mm)))
    vc = 0.66 * lam_s * EC_LAMBDA * rho ** (1 / 3) * np.sqrt(FC_MPA) * (b * 10) * d_mm  # N
    return PHI_V * vc / 9806.65


def wood_armer(m1, m2, m12):
    """Momentos de diseno de Wood-Armer (T*m/m) para la cara traccionada por +M.
    Devuelve (m1*, m2*) para el refuerzo en las direcciones locales 1 y 2."""
    m1, m2, m12 = (np.asarray(v, dtype=float) for v in (m1, m2, m12))
    a = np.abs(m12)
    x, y = m1 + a, m2 + a
    with np.errstate(divide="ignore", invalid="ignore"):
        y_si_x_neg = m2 + m12 ** 2 / np.abs(m1)       # x* = 0
        x_si_y_neg = m1 + m12 ** 2 / np.abs(m2)       # y* = 0
    xr = np.where(x >= 0, np.where(y >= 0, x, x_si_y_neg), 0.0)
    yr = np.where(y >= 0, np.where(x >= 0, y, y_si_x_neg), 0.0)
    return np.maximum(np.nan_to_num(xr), 0.0), np.maximum(np.nan_to_num(yr), 0.0)


def coords(dt):
    c = leer(dt, "modelo_definiciones.xlsx", "Coord_Nudos")[["Joint", "GlobalX", "GlobalY", "GlobalZ"]]
    c = c.assign(Joint=c["Joint"].map(_et))
    return c.set_index("Joint")


# -------------------------------------------------------------------- MUROS
def muros(dt, armado):
    """Envolventes de diseno de los muros y verificacion del armado propuesto.
    `armado` = dict con As (cm2/m) y d (cm) por zona/cara."""
    F = leer(dt, "CIMENTACION_fuerzas.xlsx", "Muros")
    C = coords(dt)
    F["Joint"] = F["Joint"].map(_et)
    F = F.join(C, on="Joint")
    F["z"] = F["GlobalZ"].round(2)
    F["Area"] = F["Area"].map(_et)

    # cota media de cada elemento -> tramo enterrado (por debajo de -3.06)
    zc = F.groupby("Area")["GlobalZ"].mean()
    F["enterrado"] = F["Area"].map(zc) < -3.06 + 1e-6

    # nudos donde llegan elementos tipo frame (apoyos de pedestales, vigas, pilastras...)
    O = leer(dt, "modelo_definiciones.xlsx", "Obj_Frames")
    A = leer(dt, "modelo_definiciones.xlsx", "Asig_Frames")
    sec = A.set_index(A["Frame"].map(_et))["AnalSect"]
    O["sec"] = O["FrameObject"].map(_et).map(sec)
    O = O[~O["sec"].isin(["VCOR30X20"])]
    muros_j = set(F["Joint"])
    att = {_et(j) for j in pd.concat([O["ElemJtI"], O["ElemJtJ"]]) if _et(j) in muros_j}
    P = C.loc[sorted(att)] if att else C.iloc[:0]
    # encuentros entre muros ortogonales (esquinas): tambien son zonas de concentracion
    orient = F.groupby("Area").apply(
        lambda g: "x" if g["GlobalX"].max() - g["GlobalX"].min() < 0.01 else "y")
    F["orient"] = F["Area"].map(orient)
    tipos = F.groupby("Joint")["orient"].agg(lambda s: set(s))
    esquinas = [j for j, t in tipos.items() if len(t) > 1]
    P = pd.concat([P, C.loc[esquinas]]) if esquinas else P

    # factores de zona por nudo (campo libre / concentracion)
    nodos = F.drop_duplicates("Joint")[["Joint", "GlobalX", "GlobalY", "GlobalZ"]].set_index("Joint")
    conc = {}
    for j, r in nodos.iterrows():
        d = np.hypot(P["GlobalX"] - r["GlobalX"], P["GlobalY"] - r["GlobalY"])
        dz = (P["GlobalZ"] - r["GlobalZ"]).abs()
        conc[j] = bool(((d <= 0.75) & (dz <= 0.8)).any())
    F["conc"] = F["Joint"].map(conc)

    # momentos de diseno Wood-Armer por cara (+M = cara interior; -M = cara del suelo)
    hi, vi = wood_armer(F["M11"].values, F["M22"].values, F["M12"].values)
    hs, vs = wood_armer(-F["M11"].values, -F["M22"].values, F["M12"].values)
    F["h_int"], F["v_int"], F["h_suelo"], F["v_suelo"] = hi, vi, hs, vs

    def zona(r):
        if r["enterrado"]:
            return "enterrado"
        if r["z"] == 0.0:
            return "tope"
        if r["z"] == -3.06:
            return "apoyo"
        return "vano"
    F["zona"] = F.apply(zona, axis=1)

    # valores nodales por tramo (enterrado / superior), caso y paso
    F["tramo"] = np.where(F["enterrado"], "E", "S")
    F["paso"] = F["StepType"].fillna("")
    vals = ["h_int", "v_int", "h_suelo", "v_suelo", "V13", "V23"]
    nodal = F.groupby(["tramo", "Joint", "OutputCase", "paso"], as_index=False)[vals].mean()
    nodal = nodal.join(nodos, on="Joint")
    nodal["z"] = nodal["GlobalZ"].round(2)
    nodal["conc"] = nodal["Joint"].map(conc)

    # promedio en franja de 1.0 m (mismo nivel y caso) para las zonas de concentracion
    fr_cols = [c + "_fr" for c in vals]
    nodal[fr_cols] = np.nan
    for _, g in nodal.groupby(["tramo", "z", "OutputCase", "paso"]):
        xy = g[["GlobalX", "GlobalY"]].values
        d = np.hypot(xy[:, None, 0] - xy[None, :, 0], xy[:, None, 1] - xy[None, :, 1])
        w = (d <= 0.5).astype(float)
        nodal.loc[g.index, fr_cols] = (w @ g[vals].values) / w.sum(1, keepdims=True)

    def zona(r):
        if r["tramo"] == "E":
            return "enterrado"
        if r["z"] == 0.0:
            return "tope"
        if r["z"] == -3.06:
            return "apoyo"
        return "vano inferior" if r["z"] <= -1.53 else "vano superior"
    nodal["zona"] = nodal.apply(zona, axis=1)

    # seccion critica en la cara superior de la zapata (e = 35 cm): z = -4.285 + 0.175
    fb = armado["e_zapata_m"] / 2 / 1.225
    e = nodal[nodal["tramo"] == "E"]
    # el nudo de la cara superior del tramo enterrado esta en la misma vertical (mismo x,y), otro Joint
    tope_xy = e[e["z"] == -3.06].assign(xy=lambda d: list(zip(d["GlobalX"].round(2), d["GlobalY"].round(2))))
    tope_xy = tope_xy.set_index(["xy", "OutputCase", "paso"])
    b = e[e["z"] <= -4.28].assign(xy=lambda d: list(zip(d["GlobalX"].round(2), d["GlobalY"].round(2))))
    filas_b = []
    for _, r in b.iterrows():
        key = (r["xy"], r["OutputCase"], r["paso"])
        if key not in tope_xy.index:
            continue
        t = tope_xy.loc[key]
        t = t.iloc[0] if isinstance(t, pd.DataFrame) else t
        fila = {"Joint": r["Joint"], "OutputCase": r["OutputCase"], "paso": r["paso"],
                "conc": r["conc"], "GlobalX": r["GlobalX"], "GlobalY": r["GlobalY"], "zona": "base"}
        for c in vals + fr_cols:
            fila[c] = r[c] + (t[c] - r[c]) * fb
        filas_b.append(fila)
    nodal = pd.concat([nodal, pd.DataFrame(filas_b)], ignore_index=True)

    # cortante a la distancia d del apoyo (z = -3.06): interpolacion con el nudo superior
    f = (armado["d_cortante_cm"] / 100.0) / 0.51
    sup = nodal[nodal["tramo"] == "S"]
    k0 = sup[sup["z"] == -3.06].set_index(["Joint", "OutputCase", "paso"])
    k1 = sup[sup["z"] == -2.55]
    k1 = k1.assign(xy=list(zip(k1["GlobalX"].round(2), k1["GlobalY"].round(2)))).set_index(
        ["xy", "OutputCase", "paso"])
    filas_d = []
    for (j, cs, ps), r in k0.iterrows():
        key = ((round(r["GlobalX"], 2), round(r["GlobalY"], 2)), cs, ps)
        if key not in k1.index:
            continue
        t = k1.loc[key]
        t = t.iloc[0] if isinstance(t, pd.DataFrame) else t
        filas_d.append({"conc": r["conc"],
                        "V23": abs(r["V23"] + (t["V23"] - r["V23"]) * f),
                        "V13": abs(r["V13"] + (t["V13"] - r["V13"]) * f),
                        "V23_fr": abs(r["V23_fr"] + (t["V23_fr"] - r["V23_fr"]) * f),
                        "V13_fr": abs(r["V13_fr"] + (t["V13_fr"] - r["V13_fr"]) * f)})
    vd = pd.DataFrame(filas_d)

    # --- envolventes
    filas = []
    for (zona_, conc_), g in nodal.groupby(["zona", "conc"]):
        col = "_fr" if conc_ else ""
        filas.append({"Zona": zona_, "Tipo": "concentracion (franja 1 m)" if conc_ else "campo libre",
                      "Nudos": g["Joint"].nunique(),
                      "M vertical cara int. (T*m/m)": g["v_int" + col].max(),
                      "M vertical cara suelo (T*m/m)": g["v_suelo" + col].max(),
                      "M horizontal cara int. (T*m/m)": g["h_int" + col].max(),
                      "M horizontal cara suelo (T*m/m)": g["h_suelo" + col].max(),
                      "V23 nodal (T/m)": g["V23" + col].abs().max(),
                      "V13 nodal (T/m)": g["V13" + col].abs().max()})
    env = pd.DataFrame(filas)
    ved = []
    for conc_, g in vd.groupby("conc"):
        col = "_fr" if conc_ else ""
        ved.append({"Tipo": "concentracion (franja 1 m)" if conc_ else "campo libre",
                    "V23 a d del apoyo (T/m)": g["V23" + col].max(),
                    "V13 a d del apoyo (T/m)": g["V13" + col].max()})
    return env, pd.DataFrame(ved), nodal, att


# ---------------------------------------------------------------- ZAPATAS
def zapatas_geometria(dt):
    """Zapatas = conjuntos de areas ZAP_* conectadas por nudos (misma numeracion
    que 'Presion_zapatas' de derivados_memoria.xlsx)."""
    conn = leer(dt, "modelo_definiciones.xlsx", "Conectividad_Areas")
    asig = leer(dt, "modelo_definiciones.xlsx", "Grupos_Asignaciones")
    C = coords(dt)
    grp = {}
    for g in ("ZAP_AISLADAS", "ZAP_CORRIDAS"):
        for a in asig.loc[(asig["GroupName"] == g) & (asig["ObjectType"] == "Area"), "ObjectLabel"]:
            grp[_et(a)] = g
    conn = conn[conn["Area"].map(_et).isin(grp)]
    jcols = [c for c in conn.columns if re.fullmatch(r"Joint\d+", c)]
    padre = {}

    def raiz(x):
        padre.setdefault(x, x)
        while padre[x] != x:
            padre[x] = padre[padre[x]]
            x = padre[x]
        return x
    trib, areas = {}, []
    for _, f in conn.iterrows():
        js = [_et(f[c]) for c in jcols if pd.notna(f[c]) and str(f[c]).strip() != ""]
        a = float(f["AreaArea"])
        for j in js:
            trib[j] = trib.get(j, 0.0) + a / len(js)
            padre[raiz(j)] = raiz(js[0])
        areas.append((_et(f["Area"]), js, a, grp[_et(f["Area"])]))
    zj = {j: raiz(j) for j in trib}
    ids = {r: f"Z{n + 1:02d}" for n, r in enumerate(sorted(set(zj.values()), key=int))}
    zap_de_joint = {j: ids[r] for j, r in zj.items()}
    info = {}
    for zid in ids.values():
        js = [j for j, z in zap_de_joint.items() if z == zid]
        xy = C.loc[js]
        ars = [a for a in areas if zap_de_joint[a[1][0]] == zid]
        tipos = {a[3] for a in ars}
        info[zid] = {"joints": js, "xmin": xy["GlobalX"].min(), "xmax": xy["GlobalX"].max(),
                     "ymin": xy["GlobalY"].min(), "ymax": xy["GlobalY"].max(),
                     "area": sum(a[2] for a in ars), "tipo": "aislada" if tipos == {"ZAP_AISLADAS"} else "corrida",
                     "z": float(xy["GlobalZ"].mean())}
    return zap_de_joint, trib, info


def nodal_wa(df, C, claves=("OutputCase", "paso")):
    """Valores nodales (promedio de elementos) y momentos Wood-Armer por cara:
    cara inferior = +M (traccion abajo), superior = -M."""
    df = df.copy()
    df["Joint"] = df["Joint"].map(_et)
    df["paso"] = df["StepType"].fillna("")
    bx, by = wood_armer(df["M11"].values, df["M22"].values, df["M12"].values)
    tx, ty = wood_armer(-df["M11"].values, -df["M22"].values, df["M12"].values)
    df["bx"], df["by"], df["tx"], df["ty"] = bx, by, tx, ty
    cols = ["bx", "by", "tx", "ty", "V13", "V23", "M11", "M22", "M12"]
    n = df.groupby(["Joint", *claves], as_index=False)[cols].mean()
    return n.join(C, on="Joint")


def corte(nodos, eje, valor, cols, tol=1e-6):
    """Valores interpolados sobre la recta eje=valor ('x' o 'y').
    Devuelve DataFrame con la coordenada a lo largo de la recta y los valores."""
    o = "GlobalY" if eje == "x" else "GlobalX"
    c = "GlobalX" if eje == "x" else "GlobalY"
    filas = []
    for pos, g in nodos.groupby(nodos[o].round(3)):
        g = g.sort_values(c)
        xs = g[c].values
        if valor < xs.min() - tol or valor > xs.max() + tol:
            continue
        fila = {"pos": pos}
        for col in cols:
            fila[col] = float(np.interp(valor, xs, g[col].values))
        filas.append(fila)
    return pd.DataFrame(filas, columns=["pos", *cols])


_trapz = getattr(np, "trapezoid", None) or getattr(np, "trapz")


def _promedio(corte_df, col):
    if corte_df.empty:
        return np.nan
    d = corte_df.sort_values("pos")
    if len(d) == 1:
        return float(d[col].iloc[0])
    return float(_trapz(d[col].values, d["pos"].values) / (d["pos"].max() - d["pos"].min()))


def elegir_barras(as_cm2, smax=30.0):
    """Armado en una capa (diametro, separacion) con area >= as_cm2."""
    for fi in (12, 14, 16, 18, 20):
        for s in (30, 25, 20, 15, 12.5, 10):
            if s > smax:
                continue
            if barra_as(fi, s) >= as_cm2 - 1e-9:
                return fi, s, barra_as(fi, s)
    return 20, 10, barra_as(20, 10)


def capacidad_punzonamiento(c, d):
    """phi*vc (kgf/cm2) para columna cuadrada c (m), peralte d (m), ACI 318-19 22.6.5.2."""
    bo = 4 * (c + d)
    beta = 1.0
    raiz = np.sqrt(FC_MPA)
    v1 = 0.33 * raiz
    v2 = 0.17 * (1 + 2 / beta) * raiz
    v3 = 0.083 * (40 * d / bo + 2) * raiz
    return PHI_V * min(v1, v2, v3) / 0.0980665, bo   # kgf/cm2


def jc_cuadrada(c, d):
    b1 = c + d
    return d * b1 ** 3 / 6 + b1 * d ** 3 / 6 + d * b1 * b1 ** 2 / 2


def zapatas_aisladas(dt, e_cm=30.0, db_mm=14, c_ped=0.40, cc_cm=7.5):
    zj, trib, info = zapatas_geometria(dt)
    C = coords(dt)
    S = leer(dt, "CIMENTACION_fuerzas.xlsx", "Shell_ZapAisl")
    N = nodal_wa(S, C)
    P = leer(dt, "CIMENTACION_fuerzas.xlsx", "Frame_Pedest")
    P = P[P["Station"] == 0].copy()
    P["paso"] = P["StepType"].fillna("")
    O = leer(dt, "modelo_definiciones.xlsx", "Obj_Frames")
    O["Frame"] = O["FrameObject"].map(_et)
    pj = O.drop_duplicates("Frame").set_index("Frame")["ElemJtI"].map(_et)
    P["Frame"] = P["Frame"].map(_et)
    P["jt"] = P["Frame"].map(pj)
    P["zap"] = P["jt"].map(lambda j: zj.get(j))
    # pedestal por zapata (por posicion en planta)
    ped_xy = {f: (C.loc[j, "GlobalX"], C.loc[j, "GlobalY"]) for f, j in pj.items() if f in set(P["Frame"])}
    filas = []
    h = e_cm / 100
    d_x = (e_cm - cc_cm - db_mm / 20) / 100            # capa inferior, barras en X
    d_y = (e_cm - cc_cm - db_mm / 10 - db_mm / 20) / 100  # segunda capa, barras en Y
    d_m = (d_x + d_y) / 2
    for zid, g in info.items():
        if g["tipo"] != "aislada":
            continue
        peds = [f for f, (x, y) in ped_xy.items()
                if g["xmin"] - 0.01 <= x <= g["xmax"] + 0.01 and g["ymin"] - 0.01 <= y <= g["ymax"] + 0.01
                and abs(C.loc[pj[f], "GlobalZ"] - g["z"]) < 0.2]
        if not peds:
            continue
        xc, yc = ped_xy[peds[0]]
        B, L = g["xmax"] - g["xmin"], g["ymax"] - g["ymin"]
        nodos = N[N["Joint"].isin(g["joints"])]
        res = {"Zapata": zid, "Pedestal": ", ".join(peds), "Bx (m)": B, "By (m)": L,
               "xc (m)": xc, "yc (m)": yc}
        # momentos de calculo en la cara del pedestal (promedio sobre el ancho)
        mom = {"bx": 0.0, "by": 0.0, "tx": 0.0, "ty": 0.0}
        vsh = {"V13": 0.0, "V23": 0.0}
        for (cs, ps), gg in nodos.groupby(["OutputCase", "paso"]):
            for lado in (-1, 1):
                cx = corte(gg, "x", xc + lado * c_ped / 2, ["bx", "tx"])
                cy = corte(gg, "y", yc + lado * c_ped / 2, ["by", "ty"])
                mom["bx"] = max(mom["bx"], _promedio(cx, "bx"))
                mom["tx"] = max(mom["tx"], _promedio(cx, "tx"))
                mom["by"] = max(mom["by"], _promedio(cy, "by"))
                mom["ty"] = max(mom["ty"], _promedio(cy, "ty"))
                if abs(lado) == 1:
                    sx = corte(gg, "x", xc + lado * (c_ped / 2 + d_x), ["V13"])
                    sy = corte(gg, "y", yc + lado * (c_ped / 2 + d_y), ["V23"])
                    if not sx.empty:
                        vsh["V13"] = max(vsh["V13"], abs(_promedio(sx, "V13")))
                    if not sy.empty:
                        vsh["V23"] = max(vsh["V23"], abs(_promedio(sy, "V23")))
        res.update({"M* inf X (T*m/m)": mom["bx"], "M* inf Y (T*m/m)": mom["by"],
                    "M* sup X (T*m/m)": mom["tx"], "M* sup Y (T*m/m)": mom["ty"],
                    "V a d dir X (T/m)": vsh["V13"], "V a d dir Y (T/m)": vsh["V23"]})
        as_min = 0.0018 * 100 * e_cm
        for etiqueta, mu, d in (("inf X", mom["bx"], d_x), ("inf Y", mom["by"], d_y),
                                ("sup X", mom["tx"], d_x), ("sup Y", mom["ty"], d_y)):
            req = float(as_req(mu, d * 100))
            fi, s, a_prov = elegir_barras(max(req, as_min))
            res[f"As req {etiqueta} (cm2/m)"] = req
            res[f"Armado {etiqueta}"] = f"phi{fi}@{int(s) if s == int(s) else s}"
            res[f"As prov {etiqueta} (cm2/m)"] = a_prov
            res[f"D/C flexion {etiqueta}"] = mu / phi_mn(a_prov, d * 100)
        # cortante en una direccion (con el armado inferior elegido)
        for etiqueta, v, d, key in (("X", vsh["V13"], d_x, "inf X"), ("Y", vsh["V23"], d_y, "inf Y")):
            rho = res[f"As prov {key} (cm2/m)"] / (100 * d * 100)
            pv = phi_vc_muro(rho, d * 100)
            res[f"phi Vc dir {etiqueta} (T/m)"] = pv
            res[f"D/C cortante {etiqueta}"] = v / pv
        # punzonamiento con las fuerzas en la base del pedestal
        pp = P[P["Frame"].isin(peds)]
        phi_vc, bo = capacidad_punzonamiento(c_ped, d_m)
        jc = jc_cuadrada(c_ped, d_m)
        gamma_v = 0.4
        mejor = (0, None)
        for _, r in pp.iterrows():
            pu = -r["P"]
            if pu <= 0:
                continue
            qu = pu / (B * L)
            vu = pu - qu * (c_ped + d_m) ** 2
            vu_esf = vu / (bo * d_m) + gamma_v * (abs(r["M2"]) + abs(r["M3"])) * (c_ped + d_m) / 2 / jc
            vu_esf = vu_esf / 10.0   # T/m2 -> kgf/cm2
            dc = vu_esf / phi_vc
            if dc > mejor[0]:
                mejor = (dc, (r["OutputCase"], r["paso"], pu, r["M2"], r["M3"], vu_esf))
        res["phi vc punz (kgf/cm2)"] = phi_vc
        if mejor[1]:
            cs, ps, pu, m2, m3, ve = mejor[1]
            res.update({"Combo punz": cs + (f" ({ps})" if ps else ""), "Pu (T)": pu, "M2 (T*m)": m2,
                        "M3 (T*m)": m3, "vu punz (kgf/cm2)": ve, "D/C punz": mejor[0]})
        res["Pu max (T)"] = float((-pp["P"]).max())
        filas.append(res)
    return pd.DataFrame(filas)


def _ventana_max(cdf, col, ancho=1.0):
    """Maximo del promedio movil (ventana `ancho` m) de `col` a lo largo del corte."""
    if cdf.empty:
        return np.nan
    d = cdf.sort_values("pos")
    p, v = d["pos"].values, d[col].values
    if p.max() - p.min() < ancho:
        return float(v.mean())
    return float(max(v[np.abs(p - q) <= ancho / 2].mean() for q in p
                     if q - ancho / 2 >= p.min() - 1e-9 and q + ancho / 2 <= p.max() + 1e-9))


# eje constante, coordenada, extension a lo largo del muro, recorte en cada extremo (m):
# se recortan 0.60 m donde el muro concurre con otro muro (efecto de esquina)
MUROS_LINEAS = {
    "x = 8.50 (este)": ("x", 8.50, 15.75, 31.70, 1.20, 0.0),
    "x = 3.90": ("x", 3.90, 9.45, 16.35, 0.60, 0.60),
    "y = 16.35": ("y", 16.35, 3.90, 8.50, 0.60, 0.60),
    "y = 9.45 (sur)": ("y", 9.45, -1.00, 3.90, 0.0, 0.60),
}


def zapatas_corridas(dt, e_cm=35.0, db_mm=14, t_muro=0.20, cc_cm=7.5):
    C = coords(dt)
    S = leer(dt, "CIMENTACION_fuerzas.xlsx", "Shell_ZapCorr")
    N = nodal_wa(S, C)
    d = (e_cm - cc_cm - db_mm / 20) / 100
    as_min = 0.0018 * 100 * e_cm
    filas = []
    for nombre, (eje, c0, a0, a1, r0, r1) in MUROS_LINEAS.items():
        a0, a1 = a0 + r0, a1 - r1
        o = "GlobalY" if eje == "x" else "GlobalX"      # coordenada a lo largo del muro
        t = "GlobalX" if eje == "x" else "GlobalY"      # coordenada transversal
        banda = N[(N[o] >= a0 - 0.3) & (N[o] <= a1 + 0.3) & (N[t] >= c0 - 1.6) & (N[t] <= c0 + 1.6)]
        bx, bt = ("bx", "tx") if eje == "x" else ("by", "ty")
        vcol = "V13" if eje == "x" else "V23"
        tmin, tmax = banda[t].min(), banda[t].max()
        for lado, signo in (("lado -", -1), ("lado +", 1)):
            cara = c0 + signo * t_muro / 2
            voladizo = abs((tmin if signo < 0 else tmax) - cara)
            if voladizo < 0.05:
                continue
            m_inf = m_sup = v_max = 0.0
            for (cs, ps), g in banda.groupby(["OutputCase", "paso"]):
                cm = corte(g, eje, cara, [bx, bt])
                cm = cm[(cm["pos"] >= a0) & (cm["pos"] <= a1)]
                cv = corte(g, eje, c0 + signo * (t_muro / 2 + d), [vcol])
                cv = cv[(cv["pos"] >= a0) & (cv["pos"] <= a1)]
                m_inf = max(m_inf, _ventana_max(cm, bx))
                m_sup = max(m_sup, _ventana_max(cm, bt))
                v_max = max(v_max, _ventana_max(cv.assign(**{vcol: cv[vcol].abs()}), vcol) if not cv.empty else 0)
            req = float(as_req(m_inf, d * 100))
            fi, sep, a_prov = elegir_barras(max(req, as_min))
            rho = a_prov / (100 * d * 100)
            pv = phi_vc_muro(rho, d * 100)
            filas.append({"Muro": nombre, "Lado": lado, "Voladizo (m)": voladizo,
                          "M* inf transversal (T*m/m)": m_inf, "M* sup transversal (T*m/m)": m_sup,
                          "As req (cm2/m)": req, "As min (cm2/m)": as_min,
                          "Armado inf transversal": f"phi{fi}@{int(sep) if sep == int(sep) else sep}",
                          "As prov (cm2/m)": a_prov, "D/C flexion": m_inf / phi_mn(a_prov, d * 100),
                          "D/C flexion malla superior": m_sup / phi_mn(a_prov, d * 100),
                          "V a d (T/m)": v_max, "phi Vc (T/m)": pv, "D/C cortante": v_max / pv,
                          "d (cm)": d * 100})
    return pd.DataFrame(filas)


def punz_pedestales_corridas(dt, e_cm=35.0, db_mm=14, c_ped=0.40, cc_cm=7.5):
    """Punzonamiento de los pedestales que descansan sobre las zapatas corridas (qu = 0, conservador)."""
    zj, trib, info = zapatas_geometria(dt)
    P = leer(dt, "CIMENTACION_fuerzas.xlsx", "Frame_Pedest")
    P = P[P["Station"] == 0].copy()
    P["paso"] = P["StepType"].fillna("")
    O = leer(dt, "modelo_definiciones.xlsx", "Obj_Frames")
    O["Frame"] = O["FrameObject"].map(_et)
    pj = O.drop_duplicates("Frame").set_index("Frame")["ElemJtI"].map(_et)
    P["Frame"] = P["Frame"].map(_et)
    P["zap"] = P["Frame"].map(pj).map(lambda j: zj.get(j))
    P = P[P["zap"].map(lambda z: z is not None and info[z]["tipo"] == "corrida")]
    d = (e_cm - cc_cm - db_mm / 10) / 100
    phi_vc, bo = capacidad_punzonamiento(c_ped, d)
    jc = jc_cuadrada(c_ped, d)
    filas = []
    for fr, g in P.groupby("Frame"):
        mejor = None
        for _, r in g.iterrows():
            pu = -r["P"]
            if pu <= 0:
                continue
            vu = pu / (bo * d) + 0.4 * (abs(r["M2"]) + abs(r["M3"])) * (c_ped + d) / 2 / jc
            vu = vu / 10.0
            if mejor is None or vu > mejor[0]:
                mejor = (vu, r["OutputCase"], pu)
        if mejor:
            filas.append({"Pedestal": fr, "Zapata": g["zap"].iloc[0], "Combo": mejor[1], "Pu (T)": mejor[2],
                          "vu (kgf/cm2)": mejor[0], "phi vc (kgf/cm2)": phi_vc, "D/C punz": mejor[0] / phi_vc})
    return pd.DataFrame(filas)


# ----------------------------------------------------- deslizamiento, levantamiento
def deslizamiento(dt, area_contacto, c_adh=0.0):
    b = leer(dt, "CIMENTACION_reacciones.xlsx", "Reac_Base")
    b["H"] = np.hypot(b["GlobalFX"], b["GlobalFY"])
    b["N"] = b["GlobalFZ"]
    b["Resistencia por friccion (T)"] = MU_FRIC * b["N"]
    b["FS friccion"] = b["Resistencia por friccion (T)"] / b["H"]
    fs_req = np.where(b["OutputCase"].isin(["CS1", "CS2"]), 1.5, 1.1)
    b["FS requerido"] = fs_req
    b["Adhesion requerida (T/m2)"] = np.maximum(0, (fs_req * b["H"] - b["Resistencia por friccion (T)"]) / area_contacto)
    b["Resultado"] = np.where(b["FS friccion"] >= fs_req, "Cumple", "NO cumple")
    return b[["OutputCase", "StepType", "GlobalFX", "GlobalFY", "N", "H", "Resistencia por friccion (T)",
              "FS friccion", "FS requerido", "Adhesion requerida (T/m2)", "Resultado"]]


# -------------------------------------------------------- elementos de hormigon
def vigas_cimentacion(dt):
    b = leer(dt, "CIMENTACION_fuerzas.xlsx", "Conc_VigasCim_Beam")
    for c in ("FTopArea", "FBotArea", "VRebar", "TLngArea", "TTrnRebar"):
        b[c] = b[c] * 1e4
    g = b.groupby("DesignSect").agg(
        Elementos=("Frame", "nunique"), **{"As sup (cm2)": ("FTopArea", "max"),
        "As inf (cm2)": ("FBotArea", "max"), "Av/s (cm2/m)": ("VRebar", "max"),
        "At long. (cm2)": ("TLngArea", "max"), "At/s transv. (cm2/m)": ("TTrnRebar", "max")}).reset_index()
    return g


def pedestales_pilastras(dt):
    out = []
    for hoja, nombre in (("Conc_Pedest_Column", "Pedestales"), ("Conc_Pilastras_Column", "Pilastras")):
        d = leer(dt, "CIMENTACION_fuerzas.xlsx", hoja)
        d["As (cm2)"] = d["PMMArea"] * 1e4
        d["Av mayor (cm2/m)"] = d["VMajRebar"] * 1e4
        d["Av menor (cm2/m)"] = d["VMinRebar"] * 1e4
        g = d.groupby(["DesignSect", "Frame"], as_index=False).agg(
            **{"As (cm2)": ("As (cm2)", "max"), "Av mayor (cm2/m)": ("Av mayor (cm2/m)", "max"),
               "Av menor (cm2/m)": ("Av menor (cm2/m)", "max")})
        g.insert(0, "Grupo", nombre)
        out.append(g)
    return pd.concat(out, ignore_index=True)


# ----------------------------------------------- armados propuestos y verificacion
ARMADO_MURO = {                       # (diametro mm, separacion cm)
    "vertical cara interior": (12, 12.5),
    "vertical cara suelo, mitad inferior": (12, 10),
    "vertical cara suelo, mitad superior": (12, 20),
    "horizontal cada cara": (12, 15),
    "horizontal cada cara, zonas de concentracion": (12, 7.5),
}
D_MURO = {"int": 15.4, "suelo": 14.4, "h_suelo": 13.2, "h_int": 14.2}   # cm
GANCHOS = {"diam": 8, "sep_h": 20.0, "sep_v": 20.0}


def muros_verificacion(env, ved):
    def get(zona, tipo, col):
        r = env[(env["Zona"] == zona) & (env["Tipo"].str.startswith(tipo))]
        return float(r[col].iloc[0]) if len(r) else np.nan

    libre, conc = "campo libre", "concentracion"
    a = ARMADO_MURO
    as_vi = barra_as(*a["vertical cara interior"])
    as_vsi = barra_as(*a["vertical cara suelo, mitad inferior"])
    as_vss = barra_as(*a["vertical cara suelo, mitad superior"])
    as_h = barra_as(*a["horizontal cada cara"])
    as_hc = barra_as(*a["horizontal cada cara, zonas de concentracion"])
    nm = lambda k: "phi%d@%s" % (a[k][0], ("%g" % a[k][1]))
    filas = []

    def fila(sol, zona, tipo, dem, arm, cap):
        filas.append({"Solicitacion": sol, "Zona": zona, "Tipo": tipo, "Demanda": dem,
                      "Armado": arm, "Capacidad": cap, "D/C": dem / cap})

    for zona in ("base", "apoyo", "vano inferior", "vano superior"):
        fila("M vertical, cara interior (T*m/m)", zona, "campo libre",
             get(zona, libre, "M vertical cara int. (T*m/m)"), nm("vertical cara interior"),
             phi_mn(as_vi, D_MURO["int"]))
    for zona, k, asv in (("apoyo", "vertical cara suelo, mitad inferior", as_vsi),
                         ("vano inferior", "vertical cara suelo, mitad inferior", as_vsi),
                         ("vano superior", "vertical cara suelo, mitad superior", as_vss)):
        fila("M vertical, cara del suelo (T*m/m)", zona, "campo libre",
             get(zona, libre, "M vertical cara suelo (T*m/m)"), nm(k), phi_mn(asv, D_MURO["suelo"]))
    fila("M vertical, cara del suelo (T*m/m)", "apoyo", "concentracion (franja 1 m)",
         get("apoyo", conc, "M vertical cara suelo (T*m/m)"), nm("vertical cara suelo, mitad inferior"),
         phi_mn(as_vsi, D_MURO["suelo"]))
    zonas_h = ("base", "apoyo", "vano inferior", "vano superior")
    dem_hi = max(get(z, libre, "M horizontal cara int. (T*m/m)") for z in zonas_h)
    dem_hs = max(get(z, libre, "M horizontal cara suelo (T*m/m)") for z in zonas_h)
    fila("M horizontal, cara interior (T*m/m)", "todas", "campo libre", dem_hi,
         nm("horizontal cada cara"), phi_mn(as_h, D_MURO["h_int"]))
    fila("M horizontal, cara del suelo (T*m/m)", "todas", "campo libre", dem_hs,
         nm("horizontal cada cara"), phi_mn(as_h, D_MURO["h_suelo"]))
    dem_hic = max(get(z, conc, "M horizontal cara int. (T*m/m)") for z in zonas_h)
    dem_hsc = max(get(z, conc, "M horizontal cara suelo (T*m/m)") for z in zonas_h)
    fila("M horizontal, cara interior (T*m/m)", "todas", "concentracion (franja 1 m)", dem_hic,
         nm("horizontal cada cara, zonas de concentracion"), phi_mn(as_hc, D_MURO["h_int"]))
    fila("M horizontal, cara del suelo (T*m/m)", "todas", "concentracion (franja 1 m)", dem_hsc,
         nm("horizontal cada cara, zonas de concentracion"), phi_mn(as_hc, D_MURO["h_suelo"]))
    # cortante a d del apoyo
    d = D_MURO["suelo"]
    rho_v = min(as_vi, as_vsi) / (100 * d)
    rho_h = as_h / (100 * D_MURO["h_suelo"])
    pv23 = phi_vc_muro(rho_v, d)
    pv13 = phi_vc_muro(rho_h, D_MURO["h_suelo"])
    v = lambda tipo, col: float(ved[ved["Tipo"].str.startswith(tipo)][col].iloc[0])
    fila("V23 a d del apoyo (T/m)", "apoyo", "campo libre", v(libre, "V23 a d del apoyo (T/m)"),
         "sin refuerzo transversal", pv23)
    fila("V13 a d del apoyo (T/m)", "apoyo", "campo libre", v(libre, "V13 a d del apoyo (T/m)"),
         "sin refuerzo transversal", pv13)
    # zonas de concentracion: ganchos
    rho_c = as_hc / (100 * D_MURO["h_suelo"])
    av = 100 / GANCHOS["sep_h"] * barra_as(GANCHOS["diam"])
    vs = av * FY * d / GANCHOS["sep_v"] / 1000.0
    pvn23 = PHI_V * (phi_vc_muro(min(as_vi, as_vsi) / (100 * d), d) / PHI_V + vs)
    pvn13 = PHI_V * (phi_vc_muro(rho_c, D_MURO["h_suelo"]) / PHI_V + vs)
    fila("V23 a d del apoyo (T/m)", "apoyo", "concentracion (franja 1 m)", v(conc, "V23 a d del apoyo (T/m)"),
         "ganchos phi%d@%g x %g" % (GANCHOS["diam"], GANCHOS["sep_h"], GANCHOS["sep_v"]), pvn23)
    fila("V13 a d del apoyo (T/m)", "apoyo", "concentracion (franja 1 m)", v(conc, "V13 a d del apoyo (T/m)"),
         "ganchos phi%d@%g x %g" % (GANCHOS["diam"], GANCHOS["sep_h"], GANCHOS["sep_v"]), pvn13)
    return pd.DataFrame(filas)


OPCIONES_COLUMNA = [("8phi16", 16.08), ("8phi18", 20.36), ("8phi20", 25.13), ("4phi25+4phi20", 32.2),
                    ("8phi25", 39.27)]
OPCIONES_COL30 = [("4phi16", 8.04), ("6phi14", 9.24), ("6phi16", 12.06)]


def armar_columnas(pp):
    filas = []
    for _, r in pp.iterrows():
        opciones = OPCIONES_COL30 if r["DesignSect"] == "COL30x30" else OPCIONES_COLUMNA
        nombre, a_prov = next(((n, a) for n, a in opciones if a >= r["As (cm2)"] - 1e-9), opciones[-1])
        req_av = max(r["Av mayor (cm2/m)"], r["Av menor (cm2/m)"])
        patas = 3 if r["DesignSect"] != "COL30x30" else 2
        av1 = patas * barra_as(10)
        sep = next((s for s in (15, 12.5, 10, 7.5) if av1 * 100 / s >= req_av), 7.5)
        filas.append({**r.to_dict(), "Armado longitudinal": nombre, "As prov (cm2)": a_prov,
                      "Estribos": "phi10@%g (%d ramas)" % (sep, patas), "Av prov (cm2/m)": av1 * 100 / sep})
    return pd.DataFrame(filas)


OPC_VIGA = {   # seccion: (barras superiores, barras inferiores, area sup, area inf, estribos)
    "VC20X30": ("2phi16", "2phi16", 4.02, 4.02),
    "VC25X40": ("4phi18", "4phi18", 10.18, 10.18),
    "VC25X45": ("4phi16", "4phi16", 8.04, 8.04),
    "VC25X50": ("3phi14", "4phi20", 4.62, 12.57),
    "VC30X50": ("4phi20", "6phi20 (2 capas)", 12.57, 18.85),
}


def armar_vigas(g):
    filas = []
    for _, r in g.iterrows():
        sup, inf, a_sup, a_inf = OPC_VIGA[r["DesignSect"]]
        sep = next((s for s in (12.5, 10, 8, 7.5, 6) if 2 * barra_as(10) * 100 / s >= r["Av/s (cm2/m)"]), 6)
        extra = ""
        if r["At long. (cm2)"] > 0:
            n = int(np.ceil(r["At long. (cm2)"] / barra_as(12)))
            n += n % 2
            extra = "+ %dphi12 laterales (torsion)" % n
        filas.append({**r.to_dict(), "Superior": sup, "Inferior": inf, "As sup prov": a_sup, "As inf prov": a_inf,
                      "Estribos": "phi10@%g" % sep, "Av prov (cm2/m)": 2 * barra_as(10) * 100 / sep,
                      "Refuerzo adicional": extra,
                      "D/C sup": r["As sup (cm2)"] / a_sup, "D/C inf": r["As inf (cm2)"] / a_inf})
    return pd.DataFrame(filas)


def resumen_zapatas(dt, za, info):
    pres = pd.read_excel(Path(dt) / "derivados_memoria.xlsx", sheet_name="Presion_zapatas")
    pres["Paso"] = pres["Paso"].fillna("")
    est = pres[pres["Paso"].isin(["", "Max"])]
    filas = []
    for zid, g in info.items():
        pz = est[est["Zapata"] == zid]
        if pz.empty:
            continue
        w = pz.loc[(pz["p prom (T/m2)"] / pz["Limite (T/m2)"]).idxmax()]
        cs1 = pz[pz["Caso"] == "CS1"]["p prom (T/m2)"]
        lev = pres[(pres["Zapata"] == zid) & (pres["Paso"].isin(["", "Min"]))]
        neto = lev[lev["Caso"].isin(["CS4X", "CS4Y"])]["F3 total (T)"].min()
        filas.append({"Zapata": zid, "Tipo": g["tipo"], "Area (m2)": g["area"],
                      "Bx (m)": g["xmax"] - g["xmin"], "By (m)": g["ymax"] - g["ymin"],
                      "p prom gobernante (T/m2)": w["p prom (T/m2)"], "Combinacion": w["Caso"],
                      "Limite (T/m2)": w["Limite (T/m2)"],
                      "Utilizacion": w["p prom (T/m2)"] / w["Limite (T/m2)"],
                      "p nodal max (T/m2)": pz["p max nodal (T/m2)"].max(),
                      "Asentamiento CS1 (mm)": float(cs1.max()) / KV * 1000 if len(cs1) else np.nan,
                      "F3 neta min CS4 (T)": neto,
                      "Levantamiento neto": "SI" if neto < 0 else "no"})
    return pd.DataFrame(filas)


def calcular_todo(dt):
    env, ved, nodal, att = muros(dt, {"d_cortante_cm": D_MURO["suelo"], "e_zapata_m": 0.35})
    mver = muros_verificacion(env, ved)
    za = zapatas_aisladas(dt)
    zc = zapatas_corridas(dt)
    _, _, info = zapatas_geometria(dt)
    area_total = sum(g["area"] for g in info.values())
    desl = deslizamiento(dt, area_total)
    vig = armar_vigas(vigas_cimentacion(dt))
    col = armar_columnas(pedestales_pilastras(dt))
    rz = resumen_zapatas(dt, za, info)
    pc = punz_pedestales_corridas(dt)
    return {"punz_corridas": pc, "muros_env": env, "muros_ved": ved, "muros_ver": mver, "zap_aisladas": za,
            "zap_corridas": zc, "deslizamiento": desl, "vigas": vig, "columnas": col,
            "zap_resumen": rz, "area_contacto": area_total}


# ----------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tablas", default="v10/tablas")
    ap.add_argument("--salida", default="v10/tablas")
    args = ap.parse_args()
    r = calcular_todo(args.tablas)
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    hojas = {"Muros_envolventes": r["muros_env"], "Muros_cortante_d": r["muros_ved"],
             "Muros_verificacion": r["muros_ver"], "Zap_aisladas": r["zap_aisladas"],
             "Zap_corridas": r["zap_corridas"], "Punz_corridas": r["punz_corridas"], "Zap_resumen": r["zap_resumen"],
             "Deslizamiento": r["deslizamiento"], "Vigas_cimentacion": r["vigas"],
             "Pedestales_pilastras": r["columnas"]}
    with pd.ExcelWriter(out / "calculo_r1.xlsx", engine="openpyxl") as xw:
        for n, df in hojas.items():
            df.to_excel(xw, sheet_name=n[:31], index=False)
    print("guardado", out / "calculo_r1.xlsx")


if __name__ == "__main__":
    main()
