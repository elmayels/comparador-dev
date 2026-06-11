# v0.3.13 - Web dashboard + JSON final estructurado

## Objetivo
Rehacer la vista final de la webapp posterior al análisis para que muestre información ejecutiva útil y homogénea para:

- matriz base / catálogo base Nestlé;
- un contratista;
- múltiples contratistas.

## Backend

Se agregó un extractor estructurado `v0.3.13` que lee el Excel final y construye un JSON de dashboard con:

- `run_type`;
- `summary`;
- `contractors` / `providers` / `provider_scores`;
- totales por fuente;
- totales de mercado cuando aplican;
- diferencias contra mercado cuando aplican;
- desglose por familia;
- partidas críticas;
- datos para gráficas;
- texto de `Analisis IA` cuando existe;
- banderas de disponibilidad.

El flujo de matriz base ahora devuelve JSON con `analysis_summary` y `download_url` en lugar de devolver únicamente el Excel como blob. Se agregó endpoint de descarga:

- `/api/v2/base-budget/{job_id}/download`

## Frontend

Se rehizo la vista final como dashboard por tabs:

- Resumen general;
- Desglose por familia;
- Gráfica de totales;
- Analisis IA, solo si existe;
- Partidas críticas, solo si existen.

Se retiró temporalmente el chat IA de la vista final.

## Homogeneidad

- Matriz base: muestra resumen, familias, gráfica y descarga Excel. No muestra Analisis IA si no existe.
- Un contratista: muestra resumen, familias, gráfica, Analisis IA y críticos cuando existan.
- Multi-contratista: muestra la misma estructura, con comparación por fuente.

## Archivos modificados

- `main.py`
- `backend/main.py`
- `frontend/index.html`

## Validación

- Compilación Python correcta.
- Se probó el extractor JSON con Excel de matriz base v0.3.12.
- Se probó el extractor JSON con Excel de contratista único v0.3.12.
- La generación de Excel no fue modificada.
