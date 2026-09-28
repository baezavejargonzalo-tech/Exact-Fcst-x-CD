"""Avance de la semana vs FCST, por cadena.

Uso:
    python seguimiento/avance_semana.py --fcst FCST_Lacteos_y_Jugos.xlsx --base base.xlsx \
        --semana 40 --anio 2026 [--fecha 2026-09-28]

- FCST: hoja "FCST" (Cadena, SubCadena Cliente, SAP, S33..S44, en toneladas). Se ignoran las filas TOTAL.
- Base: export con Semana (202640), CADENA, SKU, Solicitado, Venta Real, Quebrados, Grupo Marketing.
  Se toma solo el grupo del forecast (GRUPO LACTEOS Y JUGOS).

Guarda la foto del día en data/S<sem>/foto_<fecha>.csv, arma reportes/Avance_S<sem>.xlsx
y regenera dashboard/Avance_FCST.html con todas las fotos de la semana.
"""
import argparse
import datetime as dt
import glob
import json
import os

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GRUPO = "GRUPO LACTEOS Y JUGOS"
MEDIDAS = ["Solicitado", "Venta Real", "Quebrados"]
EXACT_MED = ["Venta Sell IN", "Venta Real", "Solicitado", "Quebrados"]
ATRIBUTOS = {"Categoria Producto": "cat", "Marca": "marca", "Foco": "foco", "SubCat DMD": "subcat"}


def exactitud_datos(fcst_path, base_path, anio):
    """Filas semana x cadena x SKU con FCST (de la hoja FCST) y Sell In / Venta Real / Solicitado de la base,
    para todas las semanas de la base que tengan columna S<nn> en el FCST."""
    f = pd.read_excel(fcst_path, sheet_name="FCST")
    f = f[f["SAP"].notna()].copy()
    f["SKU"] = f["SAP"].astype(int)
    f["CADENA"] = f["Cadena"].str.strip().str.upper()
    b = pd.read_excel(base_path)
    b = b[b["Grupo Marketing"] == GRUPO].copy()
    b["CADENA"] = b["CADENA"].str.strip().str.upper()
    b["SKU"] = b["SKU"].astype(int)
    b["sem"] = b["Semana"].astype(str).str[-2:].astype(int)
    attrs = b.groupby("SKU")[list(ATRIBUTOS)].first().rename(columns=ATRIBUTOS)
    nombres = pd.concat([b.groupby("SKU")["Nombre Producto"].first(), f.groupby("SKU")["Nombre Producto"].first()])
    nombres = nombres[~nombres.index.duplicated()]
    filas, semanas = [], []
    for sem in sorted(b["sem"].unique()):
        col = f"S{sem}"
        if col not in f.columns:
            continue
        F = f.groupby(["CADENA", "SKU"])[col].sum().rename("F")
        B = b[b["sem"] == sem].groupby(["CADENA", "SKU"])[EXACT_MED].sum()
        d = pd.concat([F, B], axis=1).fillna(0).reset_index()
        d = d[d[["F"] + EXACT_MED].abs().sum(axis=1) > 0]
        semanas.append({"sem": int(sem), "sellin": round(float(d["Venta Sell IN"].sum()), 3)})
        for r in d.itertuples(index=False):
            filas.append([int(sem), r.CADENA, int(r.SKU), round(r.F, 4), round(r[3], 4), round(r[4], 4), round(r[5], 4), round(r[6], 4)])
    skus = {int(k): {"n": nombres.get(k, ""), **{c: (None if pd.isna(v) or v == "-" else v) for c, v in attrs.loc[k].items()}}
            if k in attrs.index else {"n": nombres.get(k, "")} for k in {r[2] for r in filas}}
    return {"anio": anio, "semanas": semanas,
            "cols": ["sem", "cadena", "sku", "fcst", "sellin", "real", "solic", "queb"], "filas": filas, "skus": skus}


def leer_fcst(path, semana):
    f = pd.read_excel(path, sheet_name="FCST")
    f = f[f["SAP"].notna()].copy()
    col = f"S{semana}"
    f["SKU"] = f["SAP"].astype(int)
    f["CADENA"] = f["Cadena"].str.strip().str.upper()
    f["FCST"] = pd.to_numeric(f[col], errors="coerce").fillna(0)
    return (f.groupby(["CADENA", "SKU"], as_index=False)
             .agg(FCST=("FCST", "sum"), Producto=("Nombre Producto", "first"),
                  Subcategoria=("Subcateria", "first")))


def leer_base(path, sem_id):
    b = pd.read_excel(path)
    b = b[(b["Semana"].astype(str) == str(sem_id)) & (b["Grupo Marketing"] == GRUPO)].copy()
    b["CADENA"] = b["CADENA"].str.strip().str.upper()
    b["SKU"] = b["SKU"].astype(int)
    for m in MEDIDAS:
        b[m] = pd.to_numeric(b[m], errors="coerce").fillna(0)
    return (b.groupby(["CADENA", "SKU"], as_index=False)
             .agg(**{m: (m, "sum") for m in MEDIDAS}, ProductoBase=("Nombre Producto", "first")))


def exactitud_excel(ex, por=None):
    """Exactitud = 1 - min(|Real - FCST|, FCST) / FCST por cadena x SKU; los totales ponderan por FCST.
    Real = Sell In si la semana ya lo tiene, si no Venta Real."""
    d = pd.DataFrame(ex["filas"], columns=ex["cols"])
    con_si = {s["sem"]: s["sellin"] > 0 for s in ex["semanas"]}
    d["Real usado"] = d["sem"].map(lambda s: "Sell In" if con_si[s] else "Venta Real (semana en curso)")
    d["Real"] = d["sellin"].where(d["sem"].map(con_si), d["real"])
    d["Error"] = (d["Real"] - d["fcst"]).abs().clip(upper=d["fcst"]).where(d["fcst"] > 0, 0)
    if por:
        g = d.groupby(["sem", "Real usado", "cadena"], as_index=False)[["fcst", "Real", "Error"]].sum()
    else:
        g = d.copy()
        g["Producto"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("n", ""))
        g["Categoría"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("cat"))
        g["Marca"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("marca"))
    g["Exactitud"] = (1 - g["Error"] / g["fcst"]).where(g["fcst"] > 0)
    g["Real / FCST"] = (g["Real"] / g["fcst"]).where(g["fcst"] > 0)
    g = g.rename(columns={"sem": "Semana", "cadena": "Cadena", "sku": "SKU", "fcst": "FCST (t)", "Real": "Real (t)", "Error": "Error (t)"})
    keep = [c for c in ["Semana", "Real usado", "Cadena", "SKU", "Producto", "Categoría", "Marca", "FCST (t)", "Real (t)",
                        "Real / FCST", "Exactitud", "Error (t)"] if c in g.columns]
    return g[keep].sort_values(["Semana", "FCST (t)"], ascending=[False, False])


def foto(fcst, base):
    d = fcst.merge(base, on=["CADENA", "SKU"], how="outer")
    d["Producto"] = d["Producto"].fillna(d["ProductoBase"])
    d = d.drop(columns="ProductoBase")
    for c in ["FCST"] + MEDIDAS:
        d[c] = d[c].fillna(0).round(4)
    return d[(d[["FCST"] + MEDIDAS].abs().sum(axis=1)) > 0]


def resumen_cadena(d):
    r = d.groupby("CADENA", as_index=False)[["FCST"] + MEDIDAS].sum()
    r["Avance %"] = r["Solicitado"] / r["FCST"].where(r["FCST"] > 0)
    r["Falta vs FCST"] = r["FCST"] - r["Solicitado"]
    r["Quiebre %"] = r["Quebrados"] / r["Solicitado"].where(r["Solicitado"] > 0)
    return r.sort_values("FCST", ascending=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fcst", required=True)
    ap.add_argument("--base", required=True)
    ap.add_argument("--semana", type=int, required=True)
    ap.add_argument("--anio", type=int, default=dt.date.today().year)
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    a = ap.parse_args()

    sem_id = a.anio * 100 + a.semana
    d = foto(leer_fcst(a.fcst, a.semana), leer_base(a.base, sem_id))

    carpeta = os.path.join(ROOT, "data", f"S{a.semana}")
    os.makedirs(carpeta, exist_ok=True)
    d.to_csv(os.path.join(carpeta, f"foto_{a.fecha}.csv"), index=False)

    fotos = sorted(glob.glob(os.path.join(carpeta, "foto_*.csv")))
    historia = []
    for p in fotos:
        h = pd.read_csv(p)
        fecha = os.path.basename(p)[5:15]
        g = h.groupby("CADENA")[["FCST", "Solicitado"]].sum()
        historia.append({"fecha": fecha, "cadenas": {k: round(v, 3) for k, v in g["Solicitado"].items()},
                         "total": round(h["Solicitado"].sum(), 3)})

    r = resumen_cadena(d)
    os.makedirs(os.path.join(ROOT, "reportes"), exist_ok=True)
    xlsx = os.path.join(ROOT, "reportes", f"Avance_S{a.semana}.xlsx")
    with pd.ExcelWriter(xlsx) as w:
        tot = r[["FCST"] + MEDIDAS].sum()
        fila_tot = pd.DataFrame([{"CADENA": "TOTAL", **tot.to_dict(),
                                  "Avance %": tot["Solicitado"] / tot["FCST"],
                                  "Falta vs FCST": tot["FCST"] - tot["Solicitado"],
                                  "Quiebre %": tot["Quebrados"] / tot["Solicitado"]}])
        pd.concat([r, fila_tot]).to_excel(w, sheet_name=f"Por cadena {a.fecha}", index=False)
        det = d.copy()
        det["Avance %"] = det["Solicitado"] / det["FCST"].where(det["FCST"] > 0)
        det["Dif (Solic - FCST)"] = det["Solicitado"] - det["FCST"]
        det.sort_values(["CADENA", "FCST"], ascending=[True, False]).to_excel(w, sheet_name="Detalle SKU", index=False)

    ex = exactitud_datos(a.fcst, a.base, a.anio)
    with pd.ExcelWriter(xlsx, mode="a", engine="openpyxl") as w:
        exactitud_excel(ex).to_excel(w, sheet_name="Exactitud SKU", index=False)
        res = exactitud_excel(ex, por="CADENA")
        res.to_excel(w, sheet_name="Exactitud cadena", index=False)

    datos = {
        "semana": a.semana, "anio": a.anio, "fecha": a.fecha,
        "cadenas": json.loads(r.to_json(orient="records")),
        "skus": json.loads(d.to_json(orient="records")),
        "historia": historia,
        "exact": ex,
    }
    tpl = open(os.path.join(ROOT, "seguimiento", "plantilla.html"), encoding="utf-8").read()
    os.makedirs(os.path.join(ROOT, "dashboard"), exist_ok=True)
    out = os.path.join(ROOT, "dashboard", "Avance_FCST.html")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(tpl.replace("/*__DATOS__*/null", json.dumps(datos, ensure_ascii=False)))
    print(r.to_string(index=False))
    print("Excel:", xlsx, "\nDashboard:", out)


if __name__ == "__main__":
    main()
