#!/usr/bin/env python3
"""Figuras de la memoria Rev. 1 (mismo estilo que la Rev. 0: barras verdes, limite en linea
discontinua). Lee v10/tablas/derivados_memoria.xlsx y calculo_r1.xlsx."""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

VERDE, GRIS, ROJO, NEGRO = "#00B157", "#9AA5A0", "#C0392B", "#1A1A1A"
plt.rcParams.update({"font.size": 8, "axes.linewidth": 0.8, "font.family": "DejaVu Sans"})


def _guardar(fig, ruta):
    fig.tight_layout()
    fig.savefig(ruta, dpi=200)
    plt.close(fig)
    return ruta


def _barras_h(etiquetas, valores, limite, etq_limite, xlabel, ruta, fmt="{:.4f}", xmax=None,
              colores=None, alto=None):
    n = len(valores)
    fig, ax = plt.subplots(figsize=(8.5, alto or max(3.2, 0.42 * n + 1.2)))
    y = np.arange(n)
    cols = colores or [VERDE] * n
    ax.barh(y, valores, color=cols, height=0.78)
    xm = xmax or max(max(valores), limite) * 1.2
    for yi, v in zip(y, valores):
        dentro = v > xm * 0.18
        ax.text(v - xm * 0.01 if dentro else v + xm * 0.01, yi, fmt.format(v), va="center",
                ha="right" if dentro else "left", color="white" if dentro else NEGRO,
                fontsize=7.5, fontweight="bold")
    ax.axvline(limite, color=NEGRO, ls="--", lw=1.6)
    ax.text(limite + xm * 0.005, n - 0.45, etq_limite, fontsize=7.5, va="bottom")
    ax.set_yticks(y)
    ax.set_yticklabels(etiquetas, fontsize=7)
    ax.invert_yaxis()
    ax.set_xlim(0, xm)
    ax.set_xlabel(xlabel)
    return _guardar(fig, ruta)


def generar(dt, carpeta):
    dt, carpeta = Path(dt), Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    der = pd.read_excel(dt / "derivados_memoria.xlsx", sheet_name=None)
    cal = pd.read_excel(dt / "calculo_r1.xlsx", sheet_name=None)
    sal = {}

    # Fig. derivas
    d = der["Derivas"].sort_values(["Dir", "Caso", "Nivel (m)"], ascending=[True, False, True])
    d = d.assign(o=d["Caso"].map({"SPECX": 0, "SX": 1, "SPECY": 0, "SY": 1})).sort_values(["Dir", "o", "Nivel (m)"])
    et = [f"+{r['Nivel (m)']:.2f}\n{r['Dir']} – {r['Caso']}" for _, r in d.iterrows()]
    sal["derivas"] = _barras_h(et, d["dM"].values, 0.020, "Límite NEC 0.020",
                               "Deriva inelástica ΔM = 0.75·R·ΔE", carpeta / "r1_derivas.png",
                               xmax=0.0245, alto=4.6)

    # Fig. D/C acero
    a = der["Acero_resumen"].copy()
    a["tipo"] = a["Verificacion"].map({"PMM": "resistencia", "Other": "servicio"})
    a = a.sort_values("D/C max", ascending=False)
    et = [f"{r['Seccion']} ({r['tipo']}, n={r['Elementos']})" for _, r in a.iterrows()]
    sal["dc_acero"] = _barras_h(et, a["D/C max"].values, 0.95, "Admisible 0.95", "Relación demanda/capacidad máxima",
                                carpeta / "r1_dc_acero.png", fmt="{:.2f}", xmax=1.15, alto=4.6)

    # Fig. muros D/C
    m = cal["Muros_verificacion"].copy()
    et = [f"{r['Solicitacion'].replace(' (T*m/m)', '').replace(' (T/m)', '')}\n{r['Zona']} – {r['Tipo'].split(' (')[0]}"
          for _, r in m.iterrows()]
    sal["muros"] = _barras_h(et, m["D/C"].values, 1.0, "Capacidad = 1.0", "Demanda / capacidad",
                             carpeta / "r1_muros_dc.png", fmt="{:.2f}", xmax=1.2, alto=6.2)

    # Fig. mapa de presiones
    p = der["Presion_zapatas"].copy()
    p["Paso"] = p["Paso"].fillna("")
    p = p[p["Paso"].isin(["", "Max"])]
    env = p.assign(u=p["p prom (T/m2)"] / p["Limite (T/m2)"]).groupby("Zapata").agg(
        x=("X (m)", "first"), y=("Y (m)", "first"), a=("Area (m2)", "first"), u=("u", "max"))
    fig, ax = plt.subplots(figsize=(7.2, 6.6))
    ais = env[env["a"] < 5]
    sc = ax.scatter(ais["x"], ais["y"], s=60 + 140 * ais["a"], c=ais["u"] * 100, cmap="YlGn", vmin=0, vmax=100,
                    edgecolor=NEGRO, linewidth=0.6)
    for z, r in ais.iterrows():
        ax.annotate(f"{z}\n{r['u'] * 100:.0f} %", (r["x"], r["y"]), fontsize=5.8, ha="center", va="center")
    corr = env[env["a"] >= 5]
    for z, r in corr.iterrows():
        ax.annotate(f"zapatas corridas ({z})\nutilización máx. {r['u'] * 100:.0f} %", (8.5, 24),
                    fontsize=6.5, ha="center", color=NEGRO, bbox=dict(boxstyle="round", fc="#E8F6EE", ec=VERDE))
    cb = fig.colorbar(sc, ax=ax, shrink=0.8)
    cb.set_label("Utilización de la capacidad (%), envolvente CS1–CS3")
    ax.set(xlabel="X (m)", ylabel="Y (m)", aspect="equal", title="Presión promedio / capacidad admisible por zapata")
    ax.grid(alpha=.3)
    sal["mapa"] = _guardar(fig, carpeta / "r1_mapa_presiones.png")

    # Fig. deslizamiento
    s = cal["Deslizamiento"].copy()
    s["StepType"] = s["StepType"].fillna("")
    et = [f"{r['OutputCase']}{' ' + r['StepType'] if r['StepType'] else ''}" for _, r in s.iterrows()]
    col = [VERDE if r == "Cumple" else ROJO for r in s["Resultado"]]
    fig, ax = plt.subplots(figsize=(8.5, 3.6))
    ax.bar(range(len(s)), s["FS friccion"], color=col, width=0.7)
    for i, (fs, rq) in enumerate(zip(s["FS friccion"], s["FS requerido"])):
        ax.text(i, fs + 0.04, f"{fs:.2f}", ha="center", fontsize=7.5, fontweight="bold")
        ax.plot([i - 0.4, i + 0.4], [rq, rq], color=NEGRO, ls="--", lw=1.4)
    ax.set_xticks(range(len(s)))
    ax.set_xticklabels(et, fontsize=7.5)
    ax.set(ylabel="FS por fricción (μ = 0.445)", title="Factor de seguridad al deslizamiento (línea discontinua = FS requerido)")
    return_ = _guardar(fig, carpeta / "r1_deslizamiento.png")
    sal["desliz"] = return_

    # Fig. D/C zapatas aisladas
    z = cal["Zap_aisladas"]
    fig, ax = plt.subplots(figsize=(8.5, 3.8))
    x = np.arange(len(z))
    w = 0.27
    ax.bar(x - w, z[["D/C flexion inf X", "D/C flexion inf Y"]].max(axis=1), w, color=VERDE, label="Flexión (malla inferior)")
    ax.bar(x, z[["D/C cortante X", "D/C cortante Y"]].max(axis=1), w, color="#6FCF97", label="Cortante a d")
    ax.bar(x + w, z["D/C punz"], w, color=GRIS, label="Punzonamiento")
    ax.axhline(1.0, color=NEGRO, ls="--", lw=1.4)
    ax.text(len(z) - 0.5, 1.02, "Capacidad = 1.0", ha="right", fontsize=7.5)
    ax.set_xticks(x)
    ax.set_xticklabels(z["Zapata"], rotation=90, fontsize=7)
    ax.set(ylim=(0, 1.15), ylabel="Demanda / capacidad", title="Zapatas aisladas (e = 30 cm): relación D/C")
    ax.legend(loc="upper left", fontsize=7, ncol=3, frameon=False)
    sal["zap_dc"] = _guardar(fig, carpeta / "r1_zapatas_dc.png")
    return sal


if __name__ == "__main__":
    import sys
    print(generar(sys.argv[1] if len(sys.argv) > 1 else "v10/tablas",
                  sys.argv[2] if len(sys.argv) > 2 else "v10/figuras/memoria_r1"))
