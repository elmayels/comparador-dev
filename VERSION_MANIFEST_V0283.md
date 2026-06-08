# VERSION_MANIFEST_V0283

Version: v0.2.8.3-analysis-summary-chat
Base: v0.2.8.2-reglas-pmd-ia
Fecha: 2026-06-07

## Objetivo
Hacer que el dashboard y el chat de la webapp usen datos estructurados reales del Excel generado, en lugar de depender de un preview limitado enviado por el frontend.

## Cambios funcionales
- Se genera `analysis_summary.json` al terminar cada job de `/comparar_async`.
- El estado de `/jobs/{job_id}` incluye `analysis_id`, `analysis_summary` y `analysis_summary_public`.
- El chat `/api/v2/analysis/ask` ahora puede recibir `analysis_id`/`job_id` y cargar el resumen estructurado guardado del servidor.
- El frontend guarda el `analysis_summary` retornado por el job y lo usa para dashboard, conceptos críticos y chat.
- El chat deja de mostrar recomendaciones de falta de referencia cuando el contador es 0; ahora informa que no detectó huecos estructurados o pide un código específico.

## Datos extraídos para el resumen
- Totales contratista / referencia / diferencia.
- Conceptos de `Comparativa`.
- Conceptos críticos por impacto y riesgo.
- Ranking básico de proveedor.
- Composición por tipo desde `Detalle`.
- Hallazgos de Mano de Obra por rol cuando hay diferencia contra referencia.
- Hallazgos de equipo/herramienta, incluyendo señales como montacargas, grúa, andamio, plataforma.
- Conceptos/insumos sin referencia estructurada.
- Texto relevante de `Analisis experto IA`.

## Archivos modificados
- `main.py`
- `backend/main.py`
- `frontend/index.html`

## Archivos nuevos
- `VERSION_MANIFEST_V0283.md`
- `README_V0283_ANALYSIS_SUMMARY_CHAT.md`

## No modificado
- `processor.py`
- `backend/professional_mvp.py`
- `backend/base_matrix_proposer.py`
- Cálculos, PU, importes, matches Construdata, reglas de MO/cuadrillas y generación del Excel técnico.

## Validaciones ejecutadas
- `python -m py_compile main.py backend/main.py`
- `node --check` sobre el script Vue extraído de `frontend/index.html`
