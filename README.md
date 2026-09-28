# Exact-Fcst-x-CD

## Seguimiento diario de la semana vs FCST (por cadena)

```
python seguimiento/avance_semana.py --fcst <FCST_Lacteos_y_Jugos.xlsx> --base <base.xlsx> --semana 40 --anio 2026 --fecha AAAA-MM-DD
```

- `data/S<sem>/foto_<fecha>.csv`: foto de cada corte (cadena × SKU: FCST, Solicitado, Venta Real, Quebrados).
- `reportes/Avance_S<sem>.xlsx`: resumen por cadena + detalle SKU.
- `dashboard/Avance_FCST.html`: dashboard con la evolución de todos los cortes de la semana.

Avance = Solicitado / FCST de la semana. Solo SKU del Grupo Lácteos y Jugos.
