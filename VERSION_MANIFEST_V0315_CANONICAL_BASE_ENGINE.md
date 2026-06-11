# v0.3.15 - Motor canonico matriz base

## Objetivo
Corregir la ruptura de trazabilidad entre `Comparativa`, `Detalle` y la webapp para el flujo de matriz base / catalogo base Nestle.

## Cambios
- Se introduce un calculo canonico para matriz base.
- `Comparativa` y `Detalle` se pintan desde el mismo `records[]` calculado.
- Se preserva el operador real de Construdata por renglon:
  - `*` multiplicacion
  - `/` rendimiento / division
  - `%` porcentual
- Los renglones `%MO` y unidad `%` se calculan como porcentaje sobre la base de Mano de Obra de la matriz ya afectada por el factor de servicio, sin doble-aplicar factor.
- Matriz base sigue generando solo:
  - `Comparativa`
  - `Detalle`
- Matriz base no genera columnas de mercado ni `Analisis IA`.
- El endpoint de matriz base ahora prefiere `analysis_summary` canonico devuelto por el motor en vez de recalcular desde Excel.

## Validacion
- `Comparativa` total = suma de servicios en `Detalle`.
- `Detalle` incluye operadores `*`, `/` y `%` cuando existen en la matriz fuente.
- `analysis_summary.total` = total canonico = total de `Comparativa`.

## Archivos modificados
- `backend/base_matrix_proposer.py`
- `backend/main.py`
- `main.py`
