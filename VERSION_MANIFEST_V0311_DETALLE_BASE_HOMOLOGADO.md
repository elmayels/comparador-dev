# v0.3.11 - Detalle matriz base homologado

## Objetivo
Separar formalmente el flujo del tab `Detalle` para matriz base/catalogo Nestle del flujo de contratistas.

## Cambios
- Matriz base sigue generando solo `Comparativa` y `Detalle`.
- `Detalle` de matriz base no incluye columnas de mercado porque la matriz generada ya es la referencia de mercado.
- Se homologa el patron visual con contratistas: servicio -> insumos -> bloque financiero.
- Se agregan subtotales por `MATERIALES`, `MANO DE OBRA`, `EQUIPO Y HERRAMIENTA`, `BASICOS`, `COSTO DIRECTO`, `COSTO INDIRECTO`, `FINANCIAMIENTO`, `UTILIDAD / CARGOS ADICIONALES` y `TOTAL COSTO UNITARIO`.
- No se modifica el flujo de contratistas ni `Detalle - Pn`.

## Validacion esperada
- Matriz base: `Comparativa`, `Detalle`.
- Contratistas: `Comparativa`, `Detalle`/`Detalle - Pn`, `Analisis IA`.
