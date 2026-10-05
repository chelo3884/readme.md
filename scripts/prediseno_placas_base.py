#!/usr/bin/env python3
"""
Pre-diseno de placas base y anclajes (AISC DG1 / ACI 318-19 cap. 17) para las columnas cajon
[]200X200X8 y []250X200X8 sobre pedestales de 40 x 40 cm, a partir de la tabla de cargas
de la REV11.  Resultado: v11/prediseno/placas_base.xlsx.   Unidades: Tonf, m, kgf/cm2.

SUPUESTOS (a confirmar):
  * Pu, Tu, Vu y Mu son cargas ultimas (envolvente de las combinaciones, incluidas las de Omega0
    para Tu) en la BASE DEL PEDESTAL (cara superior de la zapata).
  * Mu X se asocia a Vu X y Mu Y a Vu Y. El momento en la placa (cara superior del pedestal) es
    M_placa = |Mu - Vu*h_ped|, con h_ped = 1.25 m (altura de los pedestales PED40X40 del modelo).
  * Pu, Tu, Vu y Mu de cada fila no son simultaneos; se verifican como casos separados:
      caso A (compresion): Pu con M_placa y V
      caso B (tension):    Tu con M_placa y V (conservador: se supone el mismo momento)
"""
import numpy as np
import pandas as pd
from pathlib import Path

FC = 280.0            # kgf/cm2 hormigon del pedestal
FY_PL = 2530.0        # A36 (placas)
FUTA = 4080.0         # F1554 Gr.36 / A36: fu = 58 ksi
FEXX = 4920.0         # E70XX
H_PED = 1.25          # m
MU_FRIC = 0.4         # acero sobre mortero/grout
PHI_C, PHI_B, PHI_W = 0.65, 0.90, 0.75

# ----------------------------------------------------------------- cargas (tabla del usuario)
# columna, apoyo, q_serv, q_sis, Pu, Tu, VuX, VuY, MuX, MuY
CARGAS = """A'-1;P3;Z3;8.0;1.3;4.8;0.9;6.9;1.6
A'-2;P3;Z3;12.1;0.0;3.8;1.3;4.9;1.9
A'-3;P3;Z3;12.6;5.3;2.3;1.3;3.3;2.0
A'-4;P3;Z3;12.4;3.9;0.9;1.3;1.1;2.0
A'-5;P3;Z3;12.3;7.6;0.8;1.4;0.9;2.0
A-3;P3;Z3;31.2;8.2;2.1;1.9;3.5;1.9
A-4;P3;Z3;36.8;5.4;1.2;1.6;1.7;1.6
A-5;P3;Z3;25.3;1.8;0.5;1.6;0.4;1.8
B-1;P2;Z2;44.1;0.0;7.3;3.3;12.0;2.9
B-2;P2;Z2;36.2;0.0;5.0;1.5;8.3;1.2
B-3;P3;Z2;40.5;0.0;2.0;1.1;3.0;0.9
B-4;P3;ZC-2;16.2;0.0;0.6;1.0;0.9;0.9
B-5;P3;ZC-2;16.3;4.8;0.7;0.8;0.9;0.7
B-6;P3;Z3;9.7;1.4;7.6;7.8;11.3;13.3
B-7;P3;Z3;11.9;0.0;6.9;8.2;9.8;13.1
B-8;P3;Z2;9.6;2.1;7.3;6.5;10.8;10.7
C-5;P2;Z2;25.0;0.0;9.0;8.8;14.1;14.1
C-6;P3;Z3;8.7;0.7;7.2;8.7;10.6;14.1
C-7;P3;Z3;9.5;0.7;6.7;8.5;9.5;13.7
C-8;P3;Z1;7.8;1.8;7.4;7.2;10.8;12.0
D-4;P1;Z2;25.7;0.2;8.9;6.7;14.7;10.6
D-5;P3;Z3;19.2;4.8;8.3;6.6;13.9;10.2
E-4;P1;Z2;30.2;0.6;8.1;7.1;13.3;11.2
E-5;P3;Z2;22.3;5.0;8.4;7.2;13.9;11.2
F-4;P2;Z2;29.4;3.0;8.0;8.1;13.3;12.8
F-5;P3;Z2;22.2;7.4;8.1;8.3;13.5;12.8
G-4;P2;Z1;24.7;4.5;6.4;8.7;11.0;14.1
G-5;P3;Z1;21.2;10.6;6.2;8.4;10.9;13.7"""
# columnas de 250x200 (marcadas con dagger en la tabla): B-1, B-2 (sobre pedestal); C-1, C-2 van sobre pilastra
COL250 = {"B-1", "B-2"}


def anchor_ase_cm2(da_mm, hilos_pulg):
    """Area efectiva a tension de la varilla roscada (ANSI B1.1)."""
    p = 25.4 / hilos_pulg
    return np.pi / 4 * (da_mm - 0.9743 * p) ** 2 / 100.0


ANCLAJES = {"phi16": (16, 11), "phi19": (19.05, 10), "phi22": (22.2, 9), "phi25": (25.4, 8)}


def carga_columna(fila):
    nombre, ped, zap, pu, tu, vx, vy, mx, my = fila.split(";")
    d = dict(col=nombre, ped=ped, zap=zap, Pu=float(pu), Tu=float(tu), Vx=float(vx), Vy=float(vy),
             Mx=float(mx), My=float(my))
    d["Mx_pl"] = abs(d["Mx"] - d["Vx"] * H_PED)
    d["My_pl"] = abs(d["My"] - d["Vy"] * H_PED)
    d["V"] = float(np.hypot(d["Vx"], d["Vy"]))
    d["M"] = float(np.hypot(d["Mx_pl"], d["My_pl"]))
    return d


def caso_compresion(P, M, N, B, f, fpmax):
    """DG1 (cap. 3.4): placa con excentricidad. Devuelve (Y, T_total, fp_real, tipo).
    P (T) compresion, M (T*m) en la placa, N (m) dimension en la direccion del momento, f (m) distancia
    del eje a los anclajes de traccion, fpmax (T/m2) presion de aplastamiento de diseno."""
    e = M / P if P > 1e-6 else 1e9
    q = fpmax * B          # T/m
    ecrit = N / 2 - P / (2 * q)
    if e <= ecrit:         # pequena excentricidad: toda la placa comprimida, sin traccion en anclajes
        Y = N - 2 * e if e > N / 6 else N
        fp = P / (B * Y) if e > N / 6 else (P / (B * N)) * (1 + 6 * e / N)
        return Y, 0.0, fp, "pequena e"
    arg = (f + N / 2) ** 2 - 2 * P * (e + f) / q
    if arg < 0:
        return np.nan, np.inf, np.nan, "no cumple"
    Y = (f + N / 2) - np.sqrt(arg)
    T = q * Y - P
    return Y, max(T, 0.0), fpmax, "gran e"


def espesor_placa(fp_t, Y, m, n_dim=None):
    """DG1: espesor requerido (cm) con fp en kgf/cm2; Y y m en cm."""
    if Y >= m:
        return 1.5 * m * np.sqrt(fp_t / FY_PL)
    return 2.11 * np.sqrt(fp_t * Y * (m - Y / 2) / FY_PL)


def verificar(d, N, B, gauge, anc, d_col, b_col, lug=(0.10, 0.075), n_anc=4):
    """Verifica una columna con una placa N x B (m) y anclajes anc. Devuelve diccionario."""
    da, hil = ANCLAJES[anc]
    ase = anchor_ase_cm2(da, hil)
    Nsa = ase * FUTA / 1000.0                  # T por anclaje
    phiNsa = 0.75 * Nsa
    phiVsa = 0.65 * 0.6 * Nsa
    A1 = N * B * 1e4
    A2 = (0.40 * 0.40) * 1e4                   # pedestal 40 x 40
    fpmax = PHI_C * min(0.85 * FC * np.sqrt(A2 / A1), 1.7 * FC) * 10   # T/m2
    out = {"phiNsa": phiNsa, "phiVsa": phiVsa, "Ase": ase}
    # --- caso A: compresion con momento (direccion mas desfavorable = resultante)
    f = gauge / 2
    Y, T, fp, tipo = caso_compresion(d["Pu"], d["M"], N, B, f, fpmax)
    out.update({"Y_A": Y, "T_A": T, "fp_A": fp, "tipoA": tipo})
    m = (N * 100 - 0.95 * d_col * 100) / 2         # cm
    n = (B * 100 - 0.95 * b_col * 100) / 2
    l = max(m, n)
    fp_kg = (fp / 10) if np.isfinite(fp) else np.nan                       # kgf/cm2
    Y_cm = Y * 100 if np.isfinite(Y) else np.nan
    t_bear = espesor_placa(fp_kg, Y_cm, l) if np.isfinite(Y_cm) else np.nan
    # --- caso B: traccion. B1: Tu repartida en los anclajes; B2: Tu + momento de la fila (conservador)
    brazo = gauge                                  # m entre filas de anclajes
    Tanc_B1 = d["Tu"] / n_anc
    Tanc_B2 = d["Tu"] / n_anc + d["M"] / (2 * brazo)
    # espesor por traccion de anclaje: voladizo desde la cara de la columna hasta el eje del anclaje
    c = max((gauge - d_col) / 2, 0.0)              # m
    ancho_ef = d_col / 2 + c * 2                   # m, ancho efectivo de linea de rotura (aprox.)
    Mpl = max(Tanc_B1, T / 2 if np.isfinite(T) else 0.0) * c / ancho_ef            # T*m/m
    t_tens = np.sqrt(4 * Mpl * 1e3 / (PHI_B * FY_PL)) * 10 if Mpl > 0 else 0.0     # mm (Z = t^2/4 por cm de ancho)
    T_anc_A = T / 2 if np.isfinite(T) else np.inf                                    # 2 anclajes en el lado traccionado
    Tu_anc = max(T_anc_A, Tanc_B1)
    Tu_anc2 = max(T_anc_A, Tanc_B2)
    # cortante: friccion + taco
    V = d["V"]
    fric = MU_FRIC * d["Pu"]
    Vlug = PHI_C * 0.85 * FC * (lug[0] * 100) * (lug[1] * 100) / 1000.0        # T por direccion (aplastamiento)
    out.update({"t_bear_mm": t_bear * 10, "t_tens_mm": t_tens, "Tanc_max": Tu_anc, "Tanc_max_conserv": Tu_anc2,
                "DC_T": Tu_anc / phiNsa, "DC_T_conserv": Tu_anc2 / phiNsa, "V": V, "fric": fric, "Vlug": Vlug,
                "DC_V_lug": max(d["Vx"], d["Vy"]) / Vlug,
                "DC_V_anc": V / (n_anc * phiVsa)})
    # interaccion de acero (ACI 17.8) con el cortante resistido por los anclajes sin taco: se reporta ambos
    out["DC_inter"] = (Tu_anc / phiNsa) + (V / (n_anc * phiVsa))
    return out


ESPESORES = (20, 25, 32)      # mm comerciales


def calcular():
    """Verifica todas las columnas con la placa de 380 x 380 mm y 4 anclajes phi19 y clasifica el espesor."""
    res = []
    for linea in CARGAS.split("\n"):
        d = carga_columna(linea)
        dcol, bcol = (0.25, 0.20) if d["col"] in COL250 else (0.20, 0.20)
        r = verificar(d, 0.38, 0.38, 0.25, "phi19", dcol, bcol)
        t_req = max(r["t_bear_mm"], r["t_tens_mm"])
        t = next((e for e in ESPESORES if e >= t_req * 1.05), np.nan)   # 5 % de margen sobre el requerido
        if d["col"] in COL250:
            tipo, t = "PB-2", 25
        elif t == 20:
            tipo = "PB-1L"
        else:
            tipo, t = "PB-1", 32       # una sola placa gruesa para todas las columnas con cortante alto
        res.append({**d, **r, "t_req_mm": t_req, "t_adopt_mm": t, "tipo": tipo})
    return pd.DataFrame(res)


if __name__ == "__main__":
    df = calcular()
    pd.set_option("display.width", 250); pd.set_option("display.max_columns", 40)
    print(df[["col", "tipo", "Pu", "Tu", "V", "M", "T_A", "t_bear_mm", "t_tens_mm", "t_adopt_mm", "DC_T", "DC_T_conserv", "DC_V_lug", "DC_V_anc"]].round(2).to_string(index=False))
    print(df.groupby("tipo")[["t_req_mm", "DC_T", "DC_T_conserv", "DC_V_lug"]].max().round(2))
    print(df.groupby("tipo")["col"].apply(list))
