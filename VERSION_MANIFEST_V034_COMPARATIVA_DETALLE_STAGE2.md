# VERSION_MANIFEST_V034_COMPARATIVA_DETALLE_STAGE2

## Objetivo
Reconstrucción etapa 2 del Excel ideal:

1. Ajustar `Comparativa` para que cada proveedor tenga su propio bloque completo:
   - P.U.
   - Importe
   - % Part.
   - % ajuste
   - Mercado P.U.
   - Mercado Importe

2. Agregar `Detalle` por contratista:
   - `Detalle` si hay un solo contratista.
   - `Detalle - P1`, `Detalle - P2`, etc. si hay múltiples contratistas.

## Alcance
No se generan todavía:

- Resumen Profesional
- Analisis experto IA
- Trazabilidad
- Matriz Propuesta
- Matriz MultiProveedor

## Archivo modificado

- `backend/professional_mvp.py`

## Nota técnica
El detalle reutiliza el parser del flujo actual de un solo contratista (`read_provider` / `extract_conceptos`) para evitar reescribir cálculos ya existentes.
