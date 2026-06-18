# Especificación Excel - V0.7 canónica

## Principios

- El Excel es un entregable ejecutivo y técnico, no una copia literal del Excel histórico.
- La hoja `Comparativa` puede ser horizontal por proveedor porque compara partidas comerciales comunes.
- El detalle de matrices/APU no debe forzar estructura horizontal común entre proveedores.
- Cada proveedor genera una hoja independiente: `Detalle - <Proveedor>`.
- El presupuesto base genera una hoja `Detalle Base` usando el mismo formato canónico del detalle de proveedor, excluyendo únicamente columnas de mercado.

## Hojas del reporte de comparación

1. `Resumen Ejecutivo`
2. `Comparativa`
3. `Detalle - <Proveedor>` por cada proveedor
4. `Partidas Críticas`
5. `Insumos Críticos`
6. `Validaciones`
7. `Análisis IA`

La hoja `Parámetros` queda eliminada del reporte de comparación.

## Comparativa

La hoja `Comparativa` usa formato horizontal por proveedor:

- Columnas base: `Partida`, `Descripción`, `Unidad`, `Cantidad`.
- Bloque de columnas por proveedor.
- Cada proveedor incluye sus columnas de precio, importe, participación, ajuste y comparación contra mercado.

El nombre de cada proveedor viene del textbox de la webapp, limitado a 10 caracteres para mantener legibilidad en Excel.

## Detalle por proveedor

Cada hoja `Detalle - <Proveedor>` se genera desde la matriz/APU propia del proveedor y se enriquece con referencias granulares desde `data`:

- `construdata-materiales-052026.xlsx`
- `construdata-manodeobra-052026*.xlsx`
- `construdata-maquinaria-052026.xlsx`
- referencias auxiliares equivalentes

`construdata_matrices.xlsx` no participa en la generación del detalle de proveedores.

## Detalle Base

La hoja `Detalle Base` pertenece al módulo independiente de presupuesto base.

Debe reutilizar el mismo layout canónico de `Detalle - <Proveedor>`, con estas diferencias:

- No incluye columnas de mercado.
- Usa el resultado de ingeniería + `data/construdata_matrices.xlsx` como matriz base/de mercado.
- Mantiene las mismas secciones, subtotales, fórmulas y estructura visual del detalle APU.

Esta decisión evita dos motores visuales distintos y mantiene la matriz base asociada al mismo modelo canónico de matriz/APU.

## Secciones canónicas del detalle

Todo detalle, sea base o proveedor, debe organizarse por bloques:

1. Partida principal.
2. Materiales.
3. Subtotal materiales.
4. Mano de obra.
5. Subtotal mano de obra.
6. Maquinaria / equipo.
7. Subtotal maquinaria.
8. Sección financiera.
9. Costo directo.
10. Indirectos.
11. Financiamiento / utilidad / porcentajes declarados.
12. Total costo unitario.

## Regla canónica de porcentajes

Los insumos declarados como porcentaje no se aplican de forma genérica al total. Se aplican sobre la base declarada en la matriz/APU:

- `% SOBRE MATERIALES`: aplica al subtotal directo de materiales.
- `% SOBRE MO`: aplica al subtotal directo de mano de obra.
- `% SOBRE MAQUINARIA`: aplica al subtotal directo de maquinaria/equipo.
- `% SOBRE DIRECTO`: aplica al costo directo.
- `% SOBRE DIRECTO+IND`: aplica al costo directo más indirectos.

V1 debe leer esta base desde el archivo real del proveedor o desde la matriz base generada. V0 lo modela con datos mockeados y fórmulas auditables.

## Motores alineados al modelo canónico

Los procesos usan motores distintos por responsabilidad, pero comparten el mismo modelo de matriz/APU y el mismo escritor de detalle:

- `BaseBudgetEngine`: conceptos de ingeniería + `construdata_matrices.xlsx`.
- `ContractorMatrixDetailEngine`: matriz/APU del proveedor + referencias granulares de `data`.
- `ExcelReportEngine`: reutiliza `_write_canonical_apu_detail_sheet()` para `Detalle Base` y `Detalle - <Proveedor>`.

Regla final:

```text
Matriz base y matrices de proveedores pueden venir de fuentes distintas, pero deben expresarse en el mismo modelo canónico de detalle APU.
```
