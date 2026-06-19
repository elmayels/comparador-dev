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

## V1 Alpha - Inicio de data real

A partir de esta versión se agrega una primera capa de lectura real de archivos XLSX conectada al modelo canónico:

- `CanonicalRun`: corrida de presupuesto base o comparativa.
- `CanonicalProvider`: proveedor con nombre corto, archivo de conceptos y archivo matriz/APU.
- `CanonicalConcept`: concepto comercial usado para comparativa.
- `CanonicalApuItem`: insumo, sección, operador, cantidad, precio, importe y match de mercado.
- `ReferenceCatalog`: catálogo granular construido desde archivos `data` de materiales, mano de obra y maquinaria.

El flujo correcto queda:

```text
XLSX cargados
  -> parsers tolerantes
  -> modelo canónico
  -> motor de cálculo/render
  -> Excel profesional
```

La V1 alpha no pretende resolver todavía la homologación perfecta de conceptos ni el matching semántico avanzado. Su objetivo es romper la dependencia de mocks estáticos y empezar a poblar los reportes desde archivos reales.

## V1.2 - Render canónico con contrato visual mínimo aceptable

Se formaliza que el modelo canónico es la capa interna de datos, no el formato visual obligado del Excel. El Excel debe expresar el modelo de forma útil para negocio e ingeniería.

### Comparativa canónica

La comparativa se basa en un `CanonicalCatalogSpine` derivado del catálogo/conceptos declarado. Este spine conserva:

- código/partida;
- descripción;
- unidad;
- cantidad;
- orden original;
- jerarquía;
- marca de concepto ejecutable;
- flags Pareto 80/20 por proveedor.

El render final conserva `A:D` como base congelada y agrega bloques de seis columnas por proveedor.

### Detalle APU canónico

La matriz/APU se guarda internamente como `CanonicalApuItem`, pero se renderiza como matriz PU auditable:

- proveedor/base a la izquierda;
- mercado a la derecha solo cuando aplica;
- secciones, subtotales y totales preservados;
- porcentajes en la posición donde fueron declarados;
- sin forzar una estructura horizontal común entre contratistas.

### Reutilización obligatoria

`Detalle Base` y `Detalle - <Proveedor>` deben compartir el mismo writer. La diferencia debe ser solo `include_market=True/False`.

## Corrección V1.3 - separación semántica entre Comparativa y Detalle

La separación canónica entre catálogo y matriz/APU queda definida así:

- `CanonicalConcept` representa partidas/conceptos del catálogo económico. Alimenta `Comparativa`, ranking, Pareto 80/20 y KPIs de monto.
- `CanonicalApuItem` representa insumos, secciones, porcentajes, subtotales y totales internos de la matriz/APU. Alimenta únicamente `Detalle - <Proveedor>` y validaciones técnicas.

La hoja `Comparativa` nunca debe mostrar detalle de matriz. Debe mostrar el P.U. total ofertado por concepto y el importe total calculado/declarado para cada proveedor. La trazabilidad del P.U. se audita en el tab de detalle individual del proveedor.

## V1.5 - Mercado dentro del modelo canónico

La comparativa de mercado no se resuelve como formato de Excel. Se modela en dos niveles:

1. **Nivel concepto/proveedor**
   - `CanonicalConcept.market_unit_price`
   - `CanonicalConcept.market_amount`
   - `CanonicalConcept.market_source`
   - `CanonicalConcept.market_state`

   Estos campos alimentan exclusivamente el tab `Comparativa`, donde se comparan precios unitarios totales por concepto.

2. **Nivel matriz/APU**
   - `CanonicalApuItem.market_unit_price`
   - `CanonicalApuItem.market_operator`
   - `CanonicalApuItem.market_quantity`
   - `CanonicalApuItem.market_amount`
   - `CanonicalApuItem.market_deviation`
   - `CanonicalApuItem.concept_key`

   Estos campos alimentan los tabs `Detalle - <Proveedor>`.

La regla canónica es:

```text
Comparativa = conceptos del catálogo + totales del proveedor + totales de mercado por concepto.
Detalle = matriz/APU del proveedor + comparación granular contra mercado.
```

El mercado por concepto se deriva desde la matriz/APU cuando viene declarado o desde la agregación de los matches granulares de mercado. No se inventa en el writer de Excel.
