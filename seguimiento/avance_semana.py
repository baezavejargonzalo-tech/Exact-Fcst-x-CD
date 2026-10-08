"""Avance de la semana y exactitud vs FCST, por cadena y SKU.

Uso:
    python seguimiento/avance_semana.py --fcst FCST_Lacteos_y_Jugos.xlsx fcst_sem_40.xlsx --base base.xlsx \
        --semana 40 --anio 2026 [--fecha 2026-09-29]

FCST (se aceptan varios; si dos archivos traen la misma semana, manda el último):
- Formato ancho: hoja "FCST" con Cadena, SAP y columnas S33..S44 (toneladas). Se ignoran las filas TOTAL.
- Formato consolidado: SAP, KAM 1 (cadena), Grupo Marketing y una columna "SEM NN" (toneladas).
Las cadenas se comparan sin espacios ("CANALTRADICIONAL" = "CANAL TRADICIONAL").

Base: export con Semana (202640), CADENA, SKU, Venta Sell IN, Solicitado, Venta Real, Quebrados, Grupo Marketing.

Guarda la foto del día en data/S<sem>/foto_<fecha>.csv, arma reportes/Avance_S<sem>.xlsx
y regenera dashboard/Avance_FCST.html.
"""
import argparse
import datetime as dt
import glob
import json
import os
import re

import pandas as pd

import stock as stk

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEDIDAS = ["Solicitado", "Venta Real", "Quebrados"]
EXACT_MED = ["Venta Sell IN", "Venta Real", "Solicitado", "Quebrados"]
ATRIBUTOS = {"Categoria Producto": "cat", "Marca": "marca", "Foco": "foco", "Grupo Marketing": "grupo"}


def clave(s):
    """Cadena o grupo sin espacios ni mayúsculas/minúsculas, para cruzar archivos."""
    return re.sub(r"\s+", "", str(s)).upper()


def leer_fcsts(paths):
    """Devuelve filas largas: sem, cad_key, cadena, SKU, FCST, grupo_key, Producto."""
    por_sem = {}
    for path in paths:
        hojas = pd.ExcelFile(path).sheet_names
        if "FCST" in hojas:
            f = pd.read_excel(path, sheet_name="FCST")
            f = f[pd.to_numeric(f["SAP"], errors="coerce").notna()].copy()
            cols = [c for c in f.columns if re.fullmatch(r"S\d{1,2}", str(c))]
            for c in cols:
                por_sem[int(c[1:])] = pd.DataFrame({
                    "cadena": f["Cadena"].str.strip().str.upper(), "SKU": f["SAP"].astype(int),
                    "FCST": pd.to_numeric(f[c], errors="coerce").fillna(0), "grupo": None,
                    "Producto": f["Nombre Producto"], "Subcategoria": f.get("Subcateria", f.get("Subcategoria"))})
        else:
            f = pd.read_excel(path, sheet_name=0)
            f = f[pd.to_numeric(f["SAP"], errors="coerce").notna()].copy()
            for c in f.columns:
                m = re.fullmatch(r"SEM\.?\s*(\d{1,2}).*", str(c).strip(), re.I)
                if m:
                    por_sem[int(m.group(1))] = pd.DataFrame({
                        "cadena": f["KAM 1"].str.strip().str.upper(), "SKU": f["SAP"].astype(int),
                        "FCST": pd.to_numeric(f[c], errors="coerce").fillna(0),
                        "grupo": f["Grupo Marketing"], "Producto": f["Nombre Producto"],
                        "Subcategoria": f.get("Subcategoria", f.get("Subcateria"))})
    out = []
    for sem, d in por_sem.items():
        d = d.copy()
        d["sem"] = sem
        d["cad_key"] = d["cadena"].map(clave)
        d["grupo_key"] = d["grupo"].map(lambda g: clave(g) if isinstance(g, str) else None)
        out.append(d)
    return pd.concat(out, ignore_index=True)


def leer_base(path):
    b = pd.read_excel(path)
    b["cadena"] = b["CADENA"].str.strip().str.upper()
    b["cad_key"] = b["cadena"].map(clave)
    b["SKU"] = b["SKU"].astype(int)
    b["sem"] = b["Semana"].astype(str).str[-2:].astype(int)
    for m in EXACT_MED:
        b[m] = pd.to_numeric(b[m], errors="coerce").fillna(0)
    return b


def mapa_subcategoria(fc, b):
    """Subcategoría por SKU: la de tu FCST (prefiere la versión con espacios), si no la SubCat DMD de la base."""
    f = fc[fc["Subcategoria"].notna()].copy()
    f["sp"] = f["Subcategoria"].astype(str).str.contains(" ")
    m = f.sort_values(["sp", "sem"], ascending=False).groupby("SKU")["Subcategoria"].first()
    base = b[b["SubCat DMD"].notna() & (b["SubCat DMD"] != "-")].groupby("SKU")["SubCat DMD"].first()
    return {**base.to_dict(), **m.to_dict()}


SUBCAT = {}


def cruzar(fc, b, sem):
    """Filas cadena x SKU de una semana: FCST + medidas de la base.
    Solo entran las cadenas de la base y los grupos que tienen FCST esa semana."""
    grupo_sku = b.groupby("SKU")["Grupo Marketing"].first()
    nombres_grupo = {clave(g): g for g in b["Grupo Marketing"].dropna().unique()}
    nombres_cad = b.groupby("cad_key")["cadena"].first()
    f = fc[(fc["sem"] == sem) & fc["cad_key"].isin(nombres_cad.index)]
    F = f.groupby(["cad_key", "SKU"]).agg(FCST=("FCST", "sum"), Producto=("Producto", "first"),
                                          gk=("grupo_key", "first")).reset_index()
    B = b[b["sem"] == sem].groupby(["cad_key", "SKU"]).agg(
        **{m: (m, "sum") for m in EXACT_MED}, ProductoB=("Nombre Producto", "first")).reset_index()
    d = F.merge(B, on=["cad_key", "SKU"], how="outer")
    d["Producto"] = d["Producto"].fillna(d["ProductoB"])
    d["Grupo"] = d["SKU"].map(grupo_sku)
    d["Grupo"] = d["Grupo"].fillna(d["gk"].map(nombres_grupo)).fillna("SIN GRUPO")
    d["Cadena"] = d["cad_key"].map(nombres_cad)
    for c in ["FCST"] + EXACT_MED:
        d[c] = d[c].fillna(0).round(4)
    grupos_con_fcst = set(d.loc[d["FCST"] > 0, "Grupo"])
    d = d[d["Grupo"].isin(grupos_con_fcst) & (d[["FCST"] + EXACT_MED].abs().sum(axis=1) > 0)]
    d["Subcategoria"] = d["SKU"].map(SUBCAT).fillna("SIN SUBCATEGORÍA")
    return d[["Cadena", "Grupo", "Subcategoria", "SKU", "Producto", "FCST"] + EXACT_MED]


def exactitud_datos(fc, b, anio):
    """Filas semana x cadena x SKU. Las semanas cerradas (con Sell In) se guardan en data/exactitud/
    para no perderlas cuando la base ya no las traiga."""
    attrs = b.groupby("SKU")[list(ATRIBUTOS)].first().rename(columns=ATRIBUTOS)
    carpeta = os.path.join(ROOT, "data", "exactitud")
    os.makedirs(carpeta, exist_ok=True)
    semanas = {}
    for sem in sorted(set(b["sem"]) & set(fc["sem"])):
        d = cruzar(fc, b, sem)
        if d.empty:
            continue
        filas, skus = [], {}
        for r in d.itertuples(index=False):
            filas.append([int(sem), r.Cadena, int(r.SKU), r.FCST, r[6], r[7], r[8], r[9], r.Grupo])
            a = attrs.loc[r.SKU].to_dict() if r.SKU in attrs.index else {}
            skus[str(int(r.SKU))] = {"n": r.Producto if isinstance(r.Producto, str) else "",
                                     **{c: (None if pd.isna(v) or v == "-" else v) for c, v in a.items()}}
        info = {"sem": int(sem), "sellin": round(float(d["Venta Sell IN"].sum()), 3),
                "real": round(float(d["Venta Real"].sum()), 3)}
        semanas[int(sem)] = {"info": info, "filas": filas, "skus": skus}
        if info["sellin"] > 0:
            with open(os.path.join(carpeta, f"S{sem}.json"), "w", encoding="utf-8") as fh:
                json.dump(semanas[int(sem)], fh, ensure_ascii=False)
    for p in glob.glob(os.path.join(carpeta, "S*.json")):
        sem = int(os.path.basename(p)[1:-5])
        if sem not in semanas:
            semanas[sem] = json.load(open(p, encoding="utf-8"))
    filas, skus, info = [], {}, []
    for sem in sorted(semanas):
        info.append(semanas[sem]["info"])
        filas += semanas[sem]["filas"]
        for k, v in semanas[sem]["skus"].items():
            skus.setdefault(int(k), v)
    for k, v in skus.items():
        v["sub"] = SUBCAT.get(k) or v.get("sub")
    return {"anio": anio, "semanas": info,
            "cols": ["sem", "cadena", "sku", "fcst", "sellin", "real", "solic", "queb", "grupo"], "filas": filas, "skus": skus}


def exactitud_excel(ex, por=None):
    """Exactitud = 1 - min(|Real - FCST|, FCST) / FCST por cadena x SKU; los totales ponderan por FCST.
    Real = Sell In si la semana está cerrada (ya tiene Sell In); si no, Solicitado."""
    d = pd.DataFrame(ex["filas"], columns=ex["cols"])
    con_si = {s["sem"]: s["sellin"] > 0 for s in ex["semanas"]}
    d["Real usado"] = d["sem"].map(lambda s: "Sell In (cerrada)" if con_si[s] else "Solicitado (en curso)")
    d["Real"] = d["sellin"].where(d["sem"].map(con_si), d["solic"])
    d["Error"] = (d["Real"] - d["fcst"]).abs().clip(upper=d["fcst"]).where(d["fcst"] > 0, 0)
    if por == "sub":
        d["Subcategoría"] = d["sku"].map(lambda k: ex["skus"].get(k, {}).get("sub"))
        g = d.groupby(["sem", "Real usado", "grupo", "Subcategoría"], as_index=False)[["fcst", "Real", "Error"]].sum()
    elif por:
        g = d.groupby(["sem", "Real usado", "grupo", "cadena"], as_index=False)[["fcst", "Real", "Error"]].sum()
    else:
        g = d.copy()
        g["Producto"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("n", ""))
        g["Categoría"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("cat"))
        g["Marca"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("marca"))
        g["Subcategoría"] = g["sku"].map(lambda k: ex["skus"].get(k, {}).get("sub"))
    g["Exactitud"] = (1 - g["Error"] / g["fcst"]).where(g["fcst"] > 0)
    g["Real / FCST"] = (g["Real"] / g["fcst"]).where(g["fcst"] > 0)
    g = g.rename(columns={"sem": "Semana", "grupo": "Grupo", "cadena": "Cadena", "sku": "SKU", "fcst": "FCST (t)",
                          "Real": "Real (t)", "Error": "Error (t)"})
    keep = [c for c in ["Semana", "Real usado", "Grupo", "Cadena", "Subcategoría", "SKU", "Producto", "Categoría", "Marca", "FCST (t)",
                        "Real (t)", "Real / FCST", "Exactitud", "Error (t)"] if c in g.columns]
    return g[keep].sort_values(["Semana", "FCST (t)"], ascending=[False, False])


def resumen_cadena(d):
    r = d.groupby("Cadena", as_index=False)[["FCST"] + MEDIDAS].sum()
    r["Avance %"] = r["Solicitado"] / r["FCST"].where(r["FCST"] > 0)
    r["Falta vs FCST"] = r["FCST"] - r["Solicitado"]
    r["Quiebre %"] = r["Quebrados"] / r["Solicitado"].where(r["Solicitado"] > 0)
    return r.sort_values("FCST", ascending=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fcst", required=True, nargs="+")
    ap.add_argument("--base", required=True)
    ap.add_argument("--semana", type=int, required=True)
    ap.add_argument("--anio", type=int, default=dt.date.today().year)
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    ap.add_argument("--stock", help="Stock_Pais_stock_<fecha>.xlsx (opcional)")
    a = ap.parse_args()

    fc = leer_fcsts(a.fcst)
    b = leer_base(a.base)
    SUBCAT.update(mapa_subcategoria(fc, b))
    d = cruzar(fc, b, a.semana)[["Cadena", "Grupo", "Subcategoria", "SKU", "Producto", "FCST"] + MEDIDAS]

    fuera = fc[(fc["sem"] == a.semana) & ~fc["cad_key"].isin(set(b["cad_key"]))].groupby("cadena")["FCST"].sum()
    if len(fuera):
        print("FCST de cadenas que no están en la base (no se comparan):", fuera.round(1).to_dict())

    carpeta = os.path.join(ROOT, "data", f"S{a.semana}")
    os.makedirs(carpeta, exist_ok=True)
    d.to_csv(os.path.join(carpeta, f"foto_{a.fecha}.csv"), index=False)

    historia = []
    for p in sorted(glob.glob(os.path.join(carpeta, "foto_*.csv"))):
        h = pd.read_csv(p)
        if "Grupo" not in h.columns:
            continue
        g = h.groupby(["Grupo", "Cadena"])[["FCST", "Solicitado"]].sum().round(3)
        historia.append({"fecha": os.path.basename(p)[5:15],
                         "filas": [[gr, cad, r.FCST, r.Solicitado] for (gr, cad), r in g.iterrows()]})

    ex = exactitud_datos(fc, b, a.anio)
    # semanas cerradas + la en curso (según la fecha del corte) y la siguiente
    sem_hoy = dt.date.fromisoformat(a.fecha).isocalendar()[1]
    ex["sem_actual"] = sem_hoy
    ex["semanas"] = [s for s in ex["semanas"] if s["sellin"] > 0 or s["sem"] in (sem_hoy, sem_hoy + 1)]
    semanas_ok = {s["sem"] for s in ex["semanas"]}
    ex["filas"] = [f for f in ex["filas"] if f[0] in semanas_ok]
    r = resumen_cadena(d)
    os.makedirs(os.path.join(ROOT, "reportes"), exist_ok=True)
    xlsx = os.path.join(ROOT, "reportes", f"Avance_S{a.semana}.xlsx")
    with pd.ExcelWriter(xlsx) as w:
        tot = r[["FCST"] + MEDIDAS].sum()
        fila_tot = pd.DataFrame([{"Cadena": "TOTAL", **tot.to_dict(),
                                  "Avance %": tot["Solicitado"] / tot["FCST"],
                                  "Falta vs FCST": tot["FCST"] - tot["Solicitado"],
                                  "Quiebre %": tot["Quebrados"] / tot["Solicitado"]}])
        pd.concat([r, fila_tot]).to_excel(w, sheet_name=f"Avance cadena {a.fecha}", index=False)
        rg = d.groupby(["Grupo", "Cadena"], as_index=False)[["FCST"] + MEDIDAS].sum()
        rg["Avance %"] = rg["Solicitado"] / rg["FCST"].where(rg["FCST"] > 0)
        rg.to_excel(w, sheet_name="Avance grupo x cadena", index=False)
        det = d.copy()
        det["Avance %"] = det["Solicitado"] / det["FCST"].where(det["FCST"] > 0)
        det["Dif (Solic - FCST)"] = det["Solicitado"] - det["FCST"]
        det.sort_values(["Cadena", "FCST"], ascending=[True, False]).to_excel(w, sheet_name="Avance SKU", index=False)
        exactitud_excel(ex, por="cadena").to_excel(w, sheet_name="Exactitud cadena", index=False)
        exactitud_excel(ex, por="sub").to_excel(w, sheet_name="Exactitud subcategoría", index=False)
        exactitud_excel(ex).to_excel(w, sheet_name="Exactitud SKU", index=False)

    stock_rows = None
    if a.stock:
        fsem = stk.leer_fcst_semanas(a.fcst, a.semana, a.anio)
        bs = b[b["sem"] == a.semana]
        cob = stk.cobertura(stk.leer_stock(a.stock), fsem, bs.groupby("SKU")["Venta Real"].sum(), a.semana)
        cob["Solicitado S"] = bs.groupby("SKU")["Solicitado"].sum().reindex(cob.index).fillna(0)
        cob["Quebrado S"] = bs.groupby("SKU")["Quebrados"].sum().reindex(cob.index).fillna(0)
        cob["Grupo"] = cob.index.map(b.groupby("SKU")["Grupo Marketing"].first()).fillna("")
        cob["Subcategoria"] = cob.index.map(SUBCAT)
        cob["Producto"] = cob["Producto"].fillna(cob.index.to_series().map(fc.groupby("SKU")["Producto"].first()))
        cob = cob[cob["Estado"] != "Sin FCST"]
        orden = {"Crítico": 0, "Bajo": 1, "OK": 2, "No está en stock país": 3}
        cob = cob.assign(_o=cob["Estado"].map(orden)).sort_values(["_o", "Cobertura (sem)"]).drop(columns="_o")
        with pd.ExcelWriter(xlsx, mode="a", engine="openpyxl") as w:
            out = cob.reset_index().rename(columns={"index": "SKU", "SAP": "SKU", "Venta Real S": f"Venta Real S{a.semana} (t)",
                                                    "Solicitado S": f"Solicitado S{a.semana} (t)", "Quebrado S": f"Quebrado S{a.semana} (t)",
                                                    "Falta semana (t)": f"Falta S{a.semana} (t)"})
            out["Cobertura (sem)"] = out["Cobertura (sem)"].replace(float("inf"), 99)
            out = out.drop(columns=["en_stock"])
            out.to_excel(w, sheet_name="Stock cobertura", index=False)
        sems = [c for c in cob.columns if c.startswith("FCST S")]
        stock_rows = []
        # alertas: SKU de las cadenas de la base (tienen grupo) con cobertura baja
        for k, r in cob[cob["Estado"].isin(["Crítico", "Bajo"]) & (cob["Grupo"] != "")].iterrows():
            stock_rows.append({"sku": int(k), "n": r["Producto"], "grupo": r["Grupo"], "sub": r["Subcategoria"], "cat": r["Categoria"],
                               "planta": r["Planta"], "estado": r["Estado"], "cob": None if r["Cobertura (sem)"] == float("inf") else round(r["Cobertura (sem)"], 2),
                               "alc": None if pd.isna(r["Alcance archivo (sem)"]) else round(r["Alcance archivo (sem)"], 1),
                               "disp": round(r["Stock disp (t)"], 2), "xlib": round(r["XLIB (t)"], 2), "transito": round(r["Tránsito (t)"], 2),
                               "bloq": round(r["Bloqueado (t)"], 2), "falta": round(r["Falta semana (t)"], 2),
                               "prox": round(r["FCST prom. próximas (t)"], 2), "fsem": [round(r[c], 2) for c in sems],
                               "sol": round(r["Solicitado S"], 2), "queb": round(r["Quebrado S"], 2)})
        stock_rows = {"fecha": a.fecha, "semanas": [int(c[6:]) for c in sems], "filas": stock_rows,
                      "n_ok": int(((cob["Estado"] == "OK") & (cob["Grupo"] != "")).sum()),
                      "n_fuera": int(((cob["Estado"] != "OK") & (cob["Grupo"] == "")).sum())}

    datos = {
        "semana": a.semana, "anio": a.anio, "fecha": a.fecha,
        "skus": json.loads(d.to_json(orient="records")),
        "historia": historia,
        "exact": ex,
        "stock": stock_rows,
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
