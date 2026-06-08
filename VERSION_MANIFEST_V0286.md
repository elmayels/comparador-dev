# VERSION_MANIFEST_V0286

Version: v0.2.8.6-multi-sin-legacy

Base: v0.2.8.5-multi-individual-details

## Objetivo
Eliminar definitivamente el modo legacy de salida multi-proveedor.

## Cambios
- `backend/professional_mvp.py`: se elimina la variable `MULTI_PROVIDER_OUTPUT_MODE` y cualquier posibilidad de usar el layout anterior para multi-proveedor.
- Multi-proveedor ahora siempre genera el layout basado en análisis individual por proveedor:
  - `Resumen Profesional`
  - `Detalle - P1`
  - `Detalle - P2`
  - `Detalle - Pn`
  - `Analisis experto IA`
  - hojas técnicas ocultas según configuración.

## Eliminado definitivamente
- Salida multi legacy con `Matriz MultiProveedor`.
- Salida multi legacy con `Comparativa` horizontal.
- Salida multi legacy con `Detalle` resumido.
- Variable `MULTI_PROVIDER_OUTPUT_MODE`.

## Sin cambios
- Cálculos.
- Matching Construdata.
- Reglas de mano de obra/cuadrillas.
- Matriz propuesta.
- Frontend.
- Endpoints.
