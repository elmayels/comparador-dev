# VERSION_MANIFEST_V0300_UNIFIED_ENGINE

## Version
v0.3.0-unified-analysis-engine

## Objetivo
Unificar el analisis de cotizaciones para 1 o N proveedores. Ya no existe un flujo conceptual separado para single vs multi ni un writer alterno de multi horizontal.

## Cambios principales
- `main.py` y `backend/main.py` llaman siempre al generador unificado para analisis de cotizaciones.
- `backend/professional_mvp.py` ahora genera un unico workbook homologado:
  - `Resumen Profesional`
  - `Comparativa`
  - `Detalle` cuando hay 1 proveedor
  - `Detalle - P1`, `Detalle - P2`, ... cuando hay N proveedores
  - `Analisis experto IA`
  - hojas tecnicas ocultas segun configuracion
- Para cada proveedor se ejecuta el motor individual completo mediante `build_comparativo` y se copian las hojas valiosas preservando estilos, anchos, merges, grafias y estructura.
- La `Comparativa` multi ahora es compacta y ejecutiva; el detalle granular vive en las hojas `Detalle - Pn`.
- Se elimino `USE_FAST_MULTI_PROVIDER` de `.env.example` y del runtime del endpoint.
- FE homologado: labels y textos hablan de analisis de cotizaciones, no de dos flujos separados.

## Matriz propuesta
El flujo de matriz propuesta se mantiene separado por la naturaleza del input: catalogo base Nestle sin proveedor. No se mezclo con el analisis de cotizaciones.

## Archivos modificados
- `main.py`
- `backend/main.py`
- `backend/professional_mvp.py`
- `frontend/index.html`
- `.env.example`

## Validacion realizada
- Compilacion Python:
  - `main.py`
  - `backend/main.py`
  - `backend/professional_mvp.py`
  - `processor.py`
  - `backend/base_matrix_proposer.py`
- Revision estatica: no quedan referencias runtime a `USE_FAST_MULTI_PROVIDER`, `_use_fast_mvp_multi`, `Matriz MultiProveedor` ni `Comparativa Multi`.

## Validacion pendiente recomendada
Ejecutar localmente con archivos reales porque el motor individual completo puede tardar varios minutos al cargar catalogos granulares Construdata. Probar:
1. 1 proveedor con resumen PU + matriz.
2. 2 proveedores con resumen PU + matriz por proveedor.
3. Matriz propuesta con solo catalogo base.
