"""Cobertura de stock país por SKU contra tu FCST.

Stock: archivo Stock_Pais (kg) con SKU, Stock disp, XLIB, Disp+XLIB, Tránsito, Bloqueado, Alcance (sem).
FCST: hoja "CONSOLIDADO 2" del FCST semanal (SAP, Fecha de inicio de semana, FCST en t, todas las cadenas).

Cobertura propia = semanas que alcanza el Disp+XLIB, descontando primero lo que falta despachar de la
semana en curso (FCST nacional de la semana - Venta Real de la base) y después el FCST de las semanas siguientes.
"""
import datetime as dt

import pandas as pd

CRITICO, BAJO = 1.0, 1.5  # semanas


def leer_stock(path):
    s = pd.read_excel(path)
    s["SKU"] = pd.to_numeric(s["SKU"], errors="coerce")
    s = s[s["SKU"].notna()].copy()
    s["SKU"] = s["SKU"].astype(int)
    kg = ["Fcst sem (kg)", "Stock disp (kg)", "XLIB (kg)", "Disp+XLIB (kg)", "Bloqueado (kg)", "Tránsito (kg)"]
    g = s.groupby("SKU").agg(**{c: (c, "sum") for c in kg}, Producto=("Producto", "first"),
                             Categoria=("Categoría", "first"), Planta=("Planta Genérica", "first"))
    for c in kg:
        g[c.replace("(kg)", "(t)")] = g.pop(c) / 1000
    return g


def leer_fcst_semanas(paths, semana, anio):
    """FCST nacional por SKU para la semana en curso y las siguientes (hoja CONSOLIDADO 2 del archivo más nuevo)."""
    for path in reversed(paths):
        if "CONSOLIDADO 2" in pd.ExcelFile(path).sheet_names:
            c = pd.read_excel(path, sheet_name="CONSOLIDADO 2")
            c["SAP"] = pd.to_numeric(c["SAP"], errors="coerce")
            c = c[c["SAP"].notna()]
            # la fecha es el domingo anterior al lunes de la semana ISO
            c["sem"] = (pd.to_datetime(c["Fecha"]) + pd.Timedelta(days=1)).dt.isocalendar().week.astype(int)
            w = c[c["sem"] >= semana].pivot_table(index="SAP", columns="sem", values="FCST", aggfunc="sum").fillna(0)
            w.index = w.index.astype(int)
            return w
    return None


def cobertura(stock, fsem, real_semana, semana):
    """Semanas de cobertura del Disp+XLIB contra lo que falta de la semana en curso y el FCST de las siguientes."""
    d = stock.copy()
    sems = [c for c in fsem.columns if c >= semana]
    d["en_stock"] = True
    d = d.join(fsem[sems].add_prefix("FCST S"), how="outer")
    d["en_stock"] = d["en_stock"].fillna(False).astype(bool)
    d["Alcance archivo (sem)"] = d["Disp+XLIB (t)"] / d["Fcst sem (t)"].where(d["Fcst sem (t)"] > 0)
    for c in sems:
        d[f"FCST S{c}"] = d[f"FCST S{c}"].fillna(0)
    d["Venta Real S"] = real_semana.reindex(d.index).fillna(0)
    d["Falta semana (t)"] = (d[f"FCST S{semana}"] - d["Venta Real S"]).clip(lower=0)
    disp = d["Disp+XLIB (t)"].fillna(0)
    resto = disp - d["Falta semana (t)"]
    cob = pd.Series(0.0, index=d.index)
    vivo = resto >= 0
    cob[~vivo] = (disp / d["Falta semana (t)"].where(d["Falta semana (t)"] > 0)).fillna(0)[~vivo] - 1
    for c in sems[1:]:
        f = d[f"FCST S{c}"]
        cubre = vivo & (resto >= f)
        parcial = vivo & ~cubre
        cob[parcial] += (resto / f.where(f > 0)).fillna(0)[parcial]
        cob[cubre & (f > 0)] += 1
        resto = resto - f
        vivo = cubre
    cob[vivo] = float("inf")  # cubre todo el horizonte del FCST
    d["Cobertura (sem)"] = cob
    prom = d[[f"FCST S{c}" for c in sems[1:]]].mean(axis=1)
    d["FCST prom. próximas (t)"] = prom
    d["Estado"] = "OK"
    d.loc[d["Cobertura (sem)"] < BAJO, "Estado"] = "Bajo"
    d.loc[d["Cobertura (sem)"] < CRITICO, "Estado"] = "Crítico"
    d.loc[(d[f"FCST S{semana}"] == 0) & (prom == 0), "Estado"] = "Sin FCST"
    d.loc[~d["en_stock"] & (d["Estado"] != "Sin FCST"), "Estado"] = "No está en stock país"
    return d
