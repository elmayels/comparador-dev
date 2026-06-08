# VERSION MANIFEST v0.2.8.5 - Multi proveedor con detalles individuales

Base: v0.2.8.4 multi-tabs-fix / v0.2.8.3-clean acumulado.

## Objetivo
Cambiar la salida multi-proveedor para que no dependa de una comparativa horizontal difícil de leer. El flujo multi ahora ejecuta el análisis individual por cada proveedor y consolida el resultado en un solo Excel con una hoja `Detalle - Pn` por proveedor.

## Archivos modificados
- `backend/professional_mvp.py`

## Cambios funcionales
- Para 2+ proveedores, `create_professional_mvp_workbook` ejecuta siempre el análisis individual por cada proveedor. No hay modo legacy.
- Por cada proveedor se ejecuta el motor individual `build_comparativo(...)`.
- Se copia la hoja `Detalle` generada por el motor individual a:
  - `Detalle - P1`
  - `Detalle - P2`
  - `Detalle - P3`, etc.
- Se agrega una hoja `Analisis experto IA` multi-proveedor con ranking, hallazgos y acciones sugeridas.
- Se mantienen ocultas hojas técnicas como `Alcance y Riesgos` y `Trazabilidad Técnica`.
- Si un análisis individual falla o no genera `Detalle`, se crea un fallback `Detalle - Pn` con el formato actual para no entregar un workbook incompleto.

## Salida esperada multi-proveedor
Visible:
- `Resumen Profesional`
- `Detalle - P1`
- `Detalle - P2`
- `Analisis experto IA`

Oculta:
- `Alcance y Riesgos`
- `Trazabilidad Técnica`

## Compatibilidad
El layout legacy multi-proveedor queda eliminado.

## No cambia
- Cálculos del motor individual.
- Matching Construdata.
- Reglas de mano de obra/cuadrillas.
- Matriz propuesta.
- Frontend.
- Endpoints.

## Validación realizada
- `python -m py_compile backend/professional_mvp.py main.py backend/main.py processor.py backend/base_matrix_proposer.py`
- Generación de Excel multi con paquete de prueba v3.
- Hojas generadas confirmadas: `Resumen Profesional`, `Detalle - P1`, `Detalle - P2`, `Analisis experto IA`, `Alcance y Riesgos`, `Trazabilidad Técnica`.
