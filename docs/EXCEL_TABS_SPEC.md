# Especificación Excel - V0.2

## Cambios aplicados

- Se elimina la hoja `Parámetros` del reporte de comparación.
- La hoja `Comparativa` se presenta en formato horizontal por proveedor, siguiendo la referencia funcional recibida:
  - columnas base: Partida, Descripción, Unidad, Cantidad
  - bloque de columnas por proveedor
  - cada proveedor incluye P.U., Importe, % Part., % ajuste, Mercado P.U. y Mercado Importe
- La hoja `Detalle APU` no tiene columna `Contratista`.
- En `Detalle APU`, cada proveedor vive en su propio bloque de columnas con su data declarada y su propia comparación contra mercado.
- El detalle incluye subtotales por secciones:
  - SUBTOTAL MATERIALES
  - SUBTOTAL MANO DE OBRA
  - SUBTOTAL MAQUINARIA
  - COSTO DIRECTO
  - COSTO INDIRECTO
  - FINANCIAMIENTO
  - TOTAL COSTO UNITARIO

## Regla canónica de porcentajes

Los insumos declarados como porcentaje no se aplican de forma genérica al total. Deben aplicarse sobre la base declarada por el contratista en su matriz/APU:

- `% SOBRE MATERIALES`: se aplica al subtotal directo de materiales.
- `% SOBRE MO`: se aplica al subtotal directo de mano de obra.
- `% SOBRE MAQUINARIA`: se aplica al subtotal directo de maquinaria/equipo.
- `% SOBRE DIRECTO`: se aplica al costo directo.
- `% SOBRE DIRECTO+IND`: se aplica al costo directo más indirectos, si así está declarado.

La V1 debe leer esta base desde el archivo real del contratista. La V0 lo modela con datos mockeados y fórmulas auditables.

## Separación de fuentes

- `data/construdata_matrices.xlsx` aplica al módulo independiente de presupuesto base.
- `data/construdata-materiales-052026.xlsx`, `data/construdata-manodeobra-052026*.xlsx`, `data/construdata-maquinaria-052026.xlsx` y referencias equivalentes aplican al detalle APU del contratista.
- El detalle APU no se genera copiando matrices Construdata; se reconstruye desde la matriz/APU declarada por cada proveedor.

## Hojas del reporte de comparación

1. `Resumen Ejecutivo`
2. `Comparativa`
3. `Detalle APU`
4. `Partidas Críticas`
5. `Insumos Críticos`
6. `Validaciones`
7. `Análisis IA`

