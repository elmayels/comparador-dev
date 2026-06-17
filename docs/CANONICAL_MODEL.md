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
