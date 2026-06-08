# v0.3.2 - Homologacion de tabs y resumen web

## Regla final de salida Excel
En los tres casos de uso, las cuatro hojas visibles principales deben existir y mantenerse en este orden:

1. Resumen Profesional
2. Comparativa
3. Detalle
4. Analisis experto IA

Hojas tecnicas como `Alcance y Riesgos`, `Trazabilidad Técnica` o `Diagnóstico Catálogos` pueden existir, pero quedan ocultas si `EXCEL_HIDE_INTERNAL_SHEETS=1`.

## Casos cubiertos

### 1 proveedor
- Mantiene el proceso individual.
- La salida final visible es: `Resumen Profesional`, `Comparativa`, `Detalle`, `Analisis experto IA`.

### N proveedores
- Ejecuta el mismo proceso individual por proveedor.
- Consolida todo en una unica hoja `Detalle`, agregando columna `Proveedor` y secciones por proveedor.
- No genera `Detalle - P1`, `Detalle - P2` ni `Matriz MultiProveedor`.

### Matriz propuesta IA
- Homologa nombres y orden:
  - `Resumen Profesional`: presupuesto propuesto.
  - `Comparativa`: resumen por partida/matriz propuesta.
  - `Detalle`: desglose APU de la matriz propuesta.
  - `Analisis experto IA`: lectura ejecutiva de la propuesta.

## Webapp / chat IA
- El resumen estructurado se alimenta desde `Comparativa`, `Detalle` y `Analisis experto IA`.
- `Detalle` soporta multiples proveedores mediante la columna `Proveedor`.
- `/jobs/{job_id}` devuelve `analysis_summary` y `analysis_summary_public` cuando el Excel termina.

## Archivos modificados
- `backend/professional_mvp.py`
- `backend/base_matrix_proposer.py`
- `processor.py`
- `main.py`
- `backend/main.py`

## Validacion ejecutada
- Compilacion Python de archivos principales.
- Prueba controlada de workbook multi: confirma orden visible de las 4 tabs principales y ausencia de tabs prohibidas.
- Prueba controlada de matriz propuesta: confirma `Resumen Profesional`, `Comparativa`, `Detalle`, `Analisis experto IA` en orden y visibles.
