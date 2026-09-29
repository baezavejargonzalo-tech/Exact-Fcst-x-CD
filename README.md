# Exact-Fcst-x-CD

## Seguimiento diario de la semana vs FCST (por cadena)

```
python seguimiento/avance_semana.py --fcst <FCST_Lacteos_y_Jugos.xlsx> --base <base.xlsx> --semana 40 --anio 2026 --fecha AAAA-MM-DD
```

- `data/S<sem>/foto_<fecha>.csv`: foto de cada corte (cadena × SKU: FCST, Solicitado, Venta Real, Quebrados).
- `reportes/Avance_S<sem>.xlsx`: resumen por cadena + detalle SKU.
- `dashboard/Avance_FCST.html`: dashboard con la evolución de todos los cortes de la semana.

Avance = Solicitado / FCST de la semana. Solo SKU del Grupo Lácteos y Jugos.

## Exactitud

Pestaña "Exactitud" del dashboard y hojas "Exactitud cadena" / "Exactitud SKU" del Excel.

- Por cadena × SKU: `1 − min(|Real − FCST|, FCST) / FCST` (entre 0 y 100 %).
- Totales (cadena, categoría, total): ponderados por el peso del SKU en el FCST, equivale a `1 − Σ error / Σ FCST`.
- Semana cerrada (con Sell In): se mide contra Sell In. Semana en curso y la siguiente: contra Solicitado (las entregas se hacen durante la semana).
