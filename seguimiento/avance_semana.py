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

    datos = {
        "semana": a.semana, "anio": a.anio, "fecha": a.fecha,
        "cadenas": json.loads(r.to_json(orient="records")),
        "skus": json.loads(d.to_json(orient="records")),
        "historia": historia,
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
