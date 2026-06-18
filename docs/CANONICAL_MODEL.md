# Modelo canónico V0

## Módulos principales

### 1. Presupuestos base

Módulo independiente. No pertenece obligatoriamente al proceso de licitación.

Entrada:

- Archivo base de conceptos creado por ingenieros (`.xlsx`).
- `data/construdata_matrices.xlsx`.

Salida:

- Presupuesto base / matriz base presupuestada.
- Excel con tabs `Comparativa` y `Detalle`.

### 2. Comparador de propuestas

Funciona con:

- Una propuesta individual.
- Múltiples contratistas.

Puede usar o no presupuesto base como referencia opcional.

### 3. Detalle APU por contratista

Se genera desde:

- Matriz/APU propia presentada por el contratista.
- Referencias granulares de `data`:
  - materiales,
  - mano de obra,
  - maquinaria,
  - porcentajes/parámetros.

No usa `data/construdata_matrices.xlsx`.

## Motores conceptuales

- `BaseBudgetEngine`
- `ContractorMatrixDetailEngine`
- `ProposalComparisonEngine`
- `AiAnalysisService`
- `ExcelReportService`

## Regla principal

El sistema calcula. La IA interpreta. El tablero comunica.

## Corrección V0.7 - Detalle Base y Detalle de Proveedor

La generación de presupuesto base y la generación de detalle por proveedor no deben tener formatos ni rutinas divergentes para expresar una matriz/APU.

A partir de V0.7, ambos procesos quedan asociados al mismo modelo canónico de detalle APU:

- `Detalle Base`: matriz generada desde conceptos de ingeniería + `data/construdata_matrices.xlsx`.
- `Detalle - <Proveedor>`: matriz/APU propia del proveedor + referencias granulares de `data`.

La diferencia no está en la estructura del detalle, sino en la fuente y en las columnas disponibles:

- `Detalle Base` excluye columnas de mercado porque ya representa la matriz de mercado/base generada.
- `Detalle - <Proveedor>` incluye columnas de mercado porque compara la matriz declarada por el proveedor contra referencias granulares.

Esto obliga a reutilizar el mismo escritor de detalle y la mayor cantidad posible de clases/rutinas comunes. La implementación V0.7 introduce un escritor compartido de detalle canónico para evitar que presupuesto base y comparativa evolucionen como procesos desconectados.
