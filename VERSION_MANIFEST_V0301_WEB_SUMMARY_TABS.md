# v0.3.1 - Unified analysis: tabs finales y resumen web desde Detalle

## Objetivo
Corregir la salida real del flujo unificado para que no reaparezcan hojas descartadas y para que la webapp/chat usen datos estructurados sin depender de abrir el Excel.

## Cambios

### Excel
- Se elimina físicamente cualquier hoja heredada:
  - `Comparativa`
  - `Matriz MultiProveedor`
  - `Comparativa Multi`
  - `Detalle Multi`
- La salida visible del análisis queda homologada:
  - 1 proveedor: `Resumen Profesional`, `Detalle`, `Analisis experto IA`
  - N proveedores: `Resumen Profesional`, `Detalle - P1`, `Detalle - P2`, `Detalle - Pn`, `Analisis experto IA`
- Se conservan hojas técnicas ocultas cuando aplica:
  - `Alcance y Riesgos`
  - `Trazabilidad Técnica`
  - `Diagnóstico Catálogos`

### Webapp / IA contextual
- `_build_analysis_summary_from_excel()` ahora lee como fuente principal:
  - `Detalle`
  - `Detalle - Pn`
- Ya no depende de `Comparativa` para alimentar:
  - dashboard web
  - ranking de proveedores
  - conceptos críticos
  - hallazgos de mano de obra
  - hallazgos de equipo/herramienta/EPP
  - faltantes de referencia Construdata
  - chat `/api/v2/analysis/ask`
- `/jobs/{job_id}` devuelve `analysis_summary` con estructura útil cuando existen hojas de detalle.

## Validación ejecutada
- Compilación OK:
  - `main.py`
  - `backend/main.py`
  - `processor.py`
  - `backend/base_matrix_proposer.py`
  - `backend/professional_mvp.py`
- Prueba controlada con workbook dummy:
  - confirma que se eliminan `Comparativa` y `Matriz MultiProveedor`
  - confirma existencia de `Detalle - P1` y `Detalle - P2`
  - confirma que el resumen web detecta filas y proveedores desde `Detalle - Pn`

## Nota
La generación end-to-end completa con Construdata puede tardar varios minutos en este entorno. El cambio crítico de estructura y resumen web fue validado con prueba controlada para evitar afirmar una validación completa no ejecutada.
