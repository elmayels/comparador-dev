# VERSION_MANIFEST_V033_COMPARATIVA_STAGE1

## Objetivo

Reestructurar desde raíz la generación del Excel de resultados, iniciando únicamente con el tab `Comparativa`.

## Regla de producto aplicada

En esta etapa el Excel final debe contener un único tab visible y existente:

1. `Comparativa`

No se generan todavía:

- `Detalle`
- `Resumen Profesional`
- `Analisis experto IA`
- `Trazabilidad Técnica`
- `Matriz Propuesta`
- `Matriz MultiProveedor`
- tabs por proveedor

## Estructura de Comparativa

Columnas base:

- `Partida`
- `Descripción`
- `Unidad`
- `Cantidad`

Por cada contratista/proveedor:

- `P.U.`
- `Importe`
- `% part`
- `% ajuste`

Columnas de mercado:

- `Mercado - P.U.`
- `Mercado - Importe`

## Cálculos

- `Importe` por contratista: `Cantidad x P.U.` cuando el importe no viene directo del parser.
- `% part`: `Importe de la partida / total del mismo contratista`.
- `% ajuste`: `(P.U. contratista - Mercado P.U.) / Mercado P.U.`.
- Fila `TOTAL`: solo totaliza columnas monetarias de importe de contratistas y `Mercado - Importe`.

## Regla azul 80/20

La concentración 80/20 se calcula de forma independiente por contratista, ordenando sus importes de mayor a menor e incluyendo partidas hasta cruzar aproximadamente el 80% acumulado.

Solo se colorean las cuatro celdas del bloque del contratista:

- `P.U.`
- `Importe`
- `% part`
- `% ajuste`

No se colorean columnas base ni columnas de mercado por esta regla.

## Archivos modificados

- `backend/professional_mvp.py`
- `backend/base_matrix_proposer.py`

## Validación realizada

Se generaron pruebas controladas para:

- 1 proveedor
- 2 proveedores
- matriz/presupuesto propuesto

En los tres casos el workbook resultante contiene únicamente la hoja `Comparativa`.
