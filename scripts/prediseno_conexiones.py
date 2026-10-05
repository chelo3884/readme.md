#!/usr/bin/env python3
"""
Pre-diseno de la conexion viga-columna (portico IMF) para VK270, VK250 y VK220, y de la conexion
viga secundaria - viga principal.  Unidades: Tonf, m, kgf/cm2.  Se verifica solo lo necesario para dimensionar
el detalle; la calificacion sismica (AISC 341-16 E2.6 / K2) sigue pendiente de ensayo o calificacion equivalente.
"""
import numpy as np
import pandas as pd

FY, FU, RY = 3515.0, 4570.0, 1.1          # A572 Gr50 (vigas, columnas, diafragmas)
FY_PL = 2530.0                            # A36 (placas de corte)
FEXX = 4920.0
CPR = min(1.2, (FY + FU) / (2 * FY))
VIGAS = {"VK270": (27.0, 14.0, 0.8, 0.4), "VK250": (25.0, 13.0, 0.8, 0.4), "VK220": (22.0, 11.0, 0.8, 0.4)}  # d, bf, tf, tw (cm)
COL_B, COL_T = 20.0, 0.8                 # cm: columna []200X200X8 (la de 250 x 200 tiene el mismo espesor)
L_TIPICA, VG = 4.5, 2.0                   # m luz tipica entre ejes; T cortante gravitacional en el extremo (supuesto)


def filete_cm(phi=0.75, a_mm=6.0):
    """Resistencia de diseno de un filete (kgf por cm de longitud) con garganta 0.707a."""
    return phi * 0.6 * FEXX * 0.7071 * (a_mm / 10)


def conexion_viga_columna():
    filas = []
    for nombre, (d, bf, tf, tw) in VIGAS.items():
        zx = bf * tf * (d - tf) + tw * (d - 2 * tf) ** 2 / 4
        mp = FY * zx / 1e5                                  # T*m
        mpr = CPR * RY * mp
        fpr = mpr / ((d - tf) / 100)                         # T fuerza en el ala
        lh = L_TIPICA - COL_B / 100                          # m
        vu = 2 * mpr / lh + VG
        # diafragma exterior (A572 Gr50): espesor por fluencia en tension sobre el ancho del ala
        t_req = fpr * 1000 / (0.9 * FY * bf)                 # cm
        tdp = next(t for t in (1.0, 1.2, 1.4, 1.6, 2.0) if t >= t_req - 1e-9)
        # zona de panel: dos paredes paralelas al eje de la viga
        h_pz = d - tf
        vn_pz = 0.9 * 0.6 * FY * 2 * COL_T * h_pz / 1000
        vcol = 6.0                                           # T supuesto de cortante de columna
        dc_pz = (fpr - vcol) / vn_pz
        # placa de corte soldada (PL 8 A36) y soldaduras
        h_tab = float(np.floor((d - 2 * tf - 5.0) / 2.5) * 2.5)  # cm, multiplo de 2.5, 25 mm libres c/ala
        w_alma = filete_cm(a_mm=3.0) * 2 * h_tab / 1000      # T, filete 3 mm a ambos lados (alma de 4 mm)
        w_tab = filete_cm(a_mm=6.0) * 2 * h_tab / 1000
        v_alma = 0.9 * 0.6 * FY * tw * d / 1000              # T cortante del alma de la viga
        filas.append({"Viga": nombre, "d (mm)": d * 10, "bf (mm)": bf * 10, "Zx (cm3)": zx, "Mp (T*m)": mp,
                      "Mpr (T*m)": mpr, "Fpr ala (T)": fpr, "Vu (T)": vu, "t diafragma req (mm)": t_req * 10,
                      "t diafragma (mm)": tdp * 10, "zona de panel D/C": dc_pz, "altura placa corte (mm)": h_tab * 10,
                      "filete alma 3 mm (T)": w_alma, "D/C alma-placa": vu / w_alma,
                      "filete placa-columna 6 mm (T)": w_tab, "D/C placa-columna": vu / w_tab,
                      "cortante alma viga (T)": v_alma, "D/C alma": vu / v_alma})
    return pd.DataFrame(filas)


def conexion_secundaria(R=6.0):
    """Reaccion de la vigueta R (T) sobre la viga principal por medio de placa de corte + rigidizadores."""
    ab = np.pi * 1.6 ** 2 / 4            # phi16
    fnv = 0.5 * 8250 / 1.0               # A325-N, Fnv = 0.5 Fu (rosca incluida) -> kgf/cm2 (Fu = 8250)
    rn_bolt = 0.75 * fnv * ab / 1000     # T por perno
    rn_apl = 0.75 * 2.4 * 1.6 * 0.4 * FU / 1000      # aplastamiento en el alma de 4 mm (por perno)
    n = 2
    cap = n * min(rn_bolt, rn_apl)
    w_tab = filete_cm(a_mm=4.0) * 2 * 10.0 / 1000    # filete 4 mm a ambos lados, L = 100 mm
    return {"R (T)": R, "perno": "2 phi16 A325-N", "phiRn corte perno (T)": rn_bolt, "phiRn aplastamiento alma 4 mm (T)": rn_apl,
            "capacidad grupo (T)": cap, "D/C pernos": R / cap, "filete placa-rigidizador 4 mm, L=100 (T)": w_tab,
            "D/C soldadura": R / w_tab}


def cartela_dc(R=6.0, e=11.5, h=20.4, a_mm=5.0):
    """Soldadura cartela-alma de la principal (filete a ambos lados, longitud h cm) con excentricidad e (cm)."""
    fv = R * 1000 / (2 * h)                    # kgf/cm por cortante
    s_w = 2 * h ** 2 / 6                       # modulo de la linea soldada (cm2 por cm de garganta unitaria)
    fm = R * 1000 * e / s_w                    # kgf/cm por flexion
    return float(np.hypot(fv, fm) / filete_cm(a_mm=a_mm))


if __name__ == "__main__":
    pd.set_option("display.width", 250); pd.set_option("display.max_columns", 40)
    print(conexion_viga_columna().round(2).T.to_string())
    print(pd.Series(conexion_secundaria()).round(2).to_string())
