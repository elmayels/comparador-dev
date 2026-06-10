# v0.3.9 - Detalle matriz base Nestle sin mercado

## Objetivo
Corregir el caso matriz base / catalogo base Nestle donde el tab `Detalle` se generaba con renglones vacios.

## Cambios
- El writer de matriz base ahora lee correctamente `detail_rows` cuando vienen como diccionarios desde `calculate_matrix_rows`.
- El tab `Detalle` de matriz base usa solo columnas propias de la matriz base:
  - Codigo
  - Concepto
  - Unidad
  - P. Unitario
  - Op.
  - Cantidad
  - Importe
  - %
- No se generan columnas `Mercado P. Unitario`, `Mercado Op.`, `Mercado Cantidad`, `Mercado Importe` en matriz base.
- Se mantiene `Analisis IA` para el caso matriz base.

## Restricciones respetadas
- No se modifica el layout de `Detalle` para contratistas.
- No se agregan tabs distintos de `Comparativa`, `Detalle`, `Analisis IA`.
- No se busca mercado adicional para matriz base: la matriz base ya es la referencia de mercado.
