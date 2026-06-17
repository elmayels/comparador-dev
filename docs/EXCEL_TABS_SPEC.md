# Especificación V0 - Tabs Excel Comparativa y Detalle

Este documento incorpora el ejemplo `Comparativo_cotizacion-138 2.xlsx` como referencia histórica para la generación del Excel de salida.

## 1. Tab Comparativa

La hoja `Comparativa` debe ser el resumen económico por partida/servicio.

Columnas canónicas observadas:

1. Partida
2. Descripción
3. Unidad
4. Cantidad
5. P.U.
6. Importe
7. % Part.
8. % ajuste
9. Mercado P.U.
10. Mercado Importe

Uso funcional:

- `P.U.` e `Importe` vienen de la propuesta del contratista.
- `% Part.` mide el peso de la partida dentro del total del contratista.
- `% ajuste` mide la desviación contra mercado.
- `Mercado P.U.` y `Mercado Importe` se calculan usando referencias granulares de `data` para materiales, mano de obra y maquinaria.
- Para un solo contratista no debe mostrarse ranking global.
- Para múltiples contratistas se deberá extender esta lógica por contratista, manteniendo un resumen económico claro.

## 2. Tab Detalle

La hoja `Detalle` debe reconstruir la matriz/APU del contratista.

Columnas canónicas observadas:

1. Código
2. Concepto
3. Unidad
4. P. Unitario
5. Op.
6. Cantidad
7. Importe
8. %
9. Columna separadora visual
10. Mercado P. Unitario
11. Mercado Op.
12. Mercado Cantidad
13. Mercado Importe

## 3. Estructura jerárquica del Detalle

Cada partida debe abrir un bloque con:

- Renglón principal de partida.
- Sección `MATERIALES`.
- Insumos de materiales.
- `SUBTOTAL MATERIALES`.
- Sección `MANO DE OBRA`.
- Insumos de mano de obra.
- `SUBTOTAL MANO DE OBRA`.
- Sección `EQUIPO Y HERRAMIENTA`, cuando aplique.
- Insumos o porcentajes como `%HERR`, `%EPP`, andamios u otros.
- `SUBTOTAL EQUIPO Y HERRAMIENTA`.
- `SECCION FINANCIERA`.
- `COSTO DIRECTO`.
- `COSTO INDIRECTO`.
- `TOTAL COSTO UNITARIO`.
- `TOTAL POR SERVICIO`.

## 4. Reglas de origen de datos

La matriz de detalle se genera a partir de la matriz/APU propia presentada por cada contratista.

No se debe usar `data/construdata_matrices.xlsx` para generar el Detalle del contratista.

Para el Detalle del contratista se usan:

- `data/construdata-materiales-052026.xlsx`
- `data/construdata-manodeobra-052026.xlsx`
- `data/construdata-manodeobra-052026-2.xlsx`
- `data/construdata-manodeobra-052026-3.xlsx`
- `data/construdata-maquinaria-052026.xlsx`
- Archivos auxiliares de porcentajes/parámetros, cuando existan.

`data/construdata_matrices.xlsx` aplica al módulo independiente de presupuesto base.

## 5. Operadores

El motor debe conservar los operadores históricos:

- `*`: cantidad directa multiplicada por precio unitario.
- `/`: rendimiento inverso o consumo calculado por división.
- `%`: porcentaje aplicado sobre una base.

## 6. Porcentajes especiales

Deben soportarse porcentajes como:

- Herramienta menor.
- EPP / equipo de protección personal.
- Andamios.
- Indirectos.
- Financiamiento.
- Utilidad.
- Otros porcentajes definidos por familia o sección.

## 7. Regla canónica

El tab `Comparativa` resume el resultado económico por partida.

El tab `Detalle` explica cómo se construye cada precio unitario desde la matriz/APU del contratista, enriquecida con referencias de mercado granular desde `data`.
