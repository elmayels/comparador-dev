# Especificación canónica de Excel - V0 corregida

## Criterio de diseño

El archivo histórico `Comparativo_cotizacion-138 2.xlsx` se usa únicamente como referencia funcional, no como plantilla visual literal.

La V0 corregida debe generar un Excel limpio, claro y profesional, con fondo blanco, encabezados suaves, tipografía legible y sin replicar el estilo oscuro/negro del archivo histórico.

## Hojas principales

### 1. Comparativa

Objetivo: resumir el resultado económico por partida/servicio.

Columnas canónicas:

- Partida
- Descripción
- Unidad
- Cantidad
- P.U. ofertado
- Importe ofertado
- % participación
- % ajuste
- Mercado P.U.
- Mercado importe
- Diferencia vs mercado
- Diferencia %
- Estado

La hoja debe calcular importes y diferencias con fórmulas. No debe ser solo una copia de valores.

### 2. Detalle

Objetivo: mostrar la reconstrucción de la matriz/APU del contratista.

Regla clave:

> El Detalle se genera desde la matriz/APU propia del contratista. No se copia una matriz completa de referencia ni se usa `construdata_matrices.xlsx` para este flujo.

Fuentes reales esperadas para el motor V1:

- Archivo del contratista: contiene la matriz/APU propia.
- `data/construdata-materiales-052026.xlsx`: referencias de materiales.
- `data/construdata-manodeobra-052026*.xlsx`: referencias de mano de obra.
- `data/construdata-maquinaria-052026.xlsx`: referencias de maquinaria/equipo.
- Archivos auxiliares de porcentajes/parámetros en `data`, si existen.

Bloques esperados:

- Partida principal
- Materiales
- Subtotal materiales
- Mano de obra
- Subtotal mano de obra
- Equipo y herramienta
- Subtotal equipo y herramienta
- Sección financiera
- Costo directo
- Costo indirecto
- Total costo unitario
- Total por servicio

Columnas canónicas:

- Código
- Concepto / insumo
- Unidad
- P. Unitario
- Op.
- Cantidad
- Importe
- %
- Separador visual
- Mercado P.U.
- Op. referencia
- Cantidad referencia
- Mercado importe
- Alerta / regla

## Qué NO debe hacerse

- No copiar literalmente la matriz completa del archivo histórico.
- No copiar el tema negro ni colores de fuente del archivo histórico.
- No usar `construdata_matrices.xlsx` para el detalle de contratistas.
- No tratar el archivo histórico como fuente de verdad visual.

## Qué SÍ debe hacerse

- Usar el archivo histórico para entender la estructura conceptual de `Comparativa` y `Detalle`.
- Generar una salida limpia y ejecutiva.
- Mantener fórmulas trazables.
- Separar claramente datos del contratista vs referencias de mercado.
- Dejar visible la regla con la que se obtuvo cada comparación.

## Iteración profesional del Excel

La versión profesional del reporte queda diseñada como un entregable ejecutivo y técnico, no como una copia visual del Excel histórico.

### Hojas del reporte de comparación

1. **Resumen Ejecutivo**: KPIs, ranking, hallazgos y navegación interna.
2. **Comparativa**: tabla económica principal normalizada, filtrable y auditable.
3. **Detalle APU**: reconstrucción ordenada de la matriz/APU del contratista, generada desde su propio archivo y enriquecida con referencias granulares de `data`.
4. **Partidas Críticas**: priorización de revisión por impacto económico y desviación.
5. **Insumos Críticos**: materiales, mano de obra, maquinaria y porcentajes que explican variaciones.
6. **Validaciones**: bitácora de alertas e inconsistencias.
7. **Parámetros**: trazabilidad del análisis, fuentes, reglas y limitaciones.
8. **Análisis IA**: narrativa ejecutiva sobre datos calculados.

### Reglas visuales

- Fondo blanco, limpio y profesional.
- Encabezados sobrios en azul corporativo.
- Tablas estructuradas con filtros.
- Congelamiento de encabezados.
- Formatos de moneda y porcentaje.
- Barras de datos para impacto económico.
- Escalas de color suaves para desviación.
- Estados de revisión con validación de datos.
- Hipervínculos internos desde Resumen Ejecutivo.

### Regla canónica del Detalle APU

El tab `Detalle APU` no se genera copiando una matriz Construdata ni pegando una matriz histórica completa. En V1, el motor debe:

1. Leer la matriz/APU propia del contratista.
2. Detectar bloques: materiales, mano de obra, maquinaria/equipo, porcentajes y sección financiera.
3. Cruzar cada insumo contra referencias granulares de `data`.
4. Calcular diferencias, desviaciones, alertas y reconciliación contra la `Comparativa`.
