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

## V1.6 - Mercado y sección financiera calculados en modelo canónico

A partir de V1.6, el mercado ya no se considera una decisión del writer de Excel. El flujo canónico queda así:

1. `CanonicalApuItem` recibe valores de mercado por match granular contra `data` o desde columnas explícitas de mercado si el archivo las trae.
2. `_post_process_market_financials()` propaga esos valores a filas estructurales del APU:
   - `Importe:`
   - `Volumen:`
   - `SUBTOTAL <sección>`
   - `COSTO DIRECTO`
   - `INDIRECTOS / UTILIDAD / FINANCIAMIENTO`
   - `PRECIO UNITARIO`
3. `apply_provider_market_to_concepts()` toma el `PRECIO UNITARIO` de mercado calculado y lo asocia al `CanonicalConcept` correspondiente.
4. `Comparativa` renderiza mercado desde `CanonicalConcept`.
5. `Detalle - <Proveedor>` renderiza mercado desde `CanonicalApuItem`.

Regla de separación:

- `Comparativa` nunca muestra insumos o matriz. Solo conceptos de catálogo y totales por P.U.
- `Detalle - <Proveedor>` muestra insumos, subtotales, mercado granular y sección financiera.

También se corrige la detección de secciones: una palabra como “Equipo” dentro de la descripción de un insumo no debe cambiar la sección activa. La sección solo cambia con encabezados explícitos del APU.

## V1.7 - Baseline manual PMD y vínculo comparativa/matriz

Esta versión formaliza los archivos PMD manuales como gold standard funcional del resultado:

- `Comparativa` renderiza exclusivamente conceptos/partidas del catálogo o de la comparativa manual. No puede mostrar insumos ni renglones de matriz/APU.
- `Detalle - <Proveedor>` renderiza la matriz/APU técnica. Es el único lugar donde aparecen materiales, mano de obra, maquinaria, porcentajes, subtotales y sección financiera.
- El precio de mercado mostrado en `Comparativa` debe venir del resultado final de mercado calculado en la matriz/APU (`PRECIO UNITARIO` de mercado) o, si el archivo manual ya lo trae, del bloque Mercado de la comparativa.
- Cada concepto debe estar vinculado con su análisis APU mediante `concept_key`, normalmente derivado del código de partida/análisis.
- El Pareto 80/20 es propiedad de cada proveedor (`pareto_concept_keys`), no del catálogo compartido.

### Flujo canónico actualizado

```text
Conceptos / Comparativa manual
→ CanonicalConcept
→ Comparativa

Matriz / Hoja1 / PU
→ CanonicalApuItem
→ mercado granular por insumo
→ subtotales por sección
→ sección financiera
→ precio unitario mercado
→ CanonicalConcept.market_unit_price
→ Comparativa
```

### Layouts soportados de matriz

El parser de matriz ya no depende solo de columnas fijas. Detecta dos familias de layout:

1. Layout PMD estándar:
   - A:H matriz proveedor
   - I separador
   - J:M mercado

2. Layout PMD desplazado/ancho:
   - código, descripción, unidad, cantidad, P.U. e importe en columnas desplazadas
   - mercado en columnas finales equivalentes


## V1.8 - Regla canónica de conceptos, matriz y mercado fallback

- La hoja `Comparativa` se alimenta exclusivamente de `CanonicalConcept` proveniente del archivo de conceptos/catálogo del contratista.
- Un concepto participa en `Comparativa` solo si declara `unidad` y `cantidad > 0`.
- Las filas de matriz/APU, insumos, subtotales, importes, volúmenes y sección financiera nunca se promueven a `Comparativa`.
- El sistema clasifica el rol real de cada archivo por estructura, no por nombre ni por el campo de carga. Si un proveedor carga el PU/APU en el slot de conceptos y el resumen por conceptos en el slot de matriz, el pipeline usa el archivo correcto para cada rol.
- La matriz/APU se expresa en `CanonicalApuItem` y se renderiza únicamente en `Detalle - <Proveedor>`.
- Cuando no existe match granular contra Construdata, el mercado del insumo usa el valor declarado por el contratista y marca estado `Sin referencia - usa contratista`. Esto evita columnas de mercado vacías y conserva trazabilidad.
- Los cálculos financieros de mercado se derivan dentro del modelo canónico: subtotales de sección, costo directo, indirectos, financiamiento, utilidad, subtotal financiero y precio unitario.

## V1.9 - Proveniencia visual de columnas de mercado

Regla canónica para el tab `Detalle - <Proveedor>`:

- La matriz del contratista se lee y se muestra tal como viene declarada.
- El sistema no recalcula los valores del contratista.
- Las columnas de mercado (`Mercado P. Unitario`, `Mercado Op.`, `Mercado Cantidad`, `Mercado Importe`) se calculan desde referencia Construdata o, si no existe referencia, usan fallback al valor del contratista.
- Cada fila APU conserva flags de proveniencia:
  - `market_unit_price_is_fallback`
  - `market_operator_is_fallback`
  - `market_quantity_is_fallback`
  - `market_amount_is_fallback`
- El Excel renderiza en negrita únicamente los valores de mercado que no son fallback o que difieren del valor declarado por el contratista.
- Los valores de mercado copiados como fallback permanecen visualmente normales.

Esta regla evita que el estilo sea un parche del Excel writer: el resaltado visual depende de la semántica del modelo canónico de mercado.


## V2.1 - Porcentajes Construdata y servicios no porcentuales

- La matriz del contratista se conserva como dato declarado; no se recalcula la columna del contratista.
- Las columnas de mercado calculan únicamente el carril de referencia.
- La detección de porcentajes ya no usa cualquier símbolo `%` dentro de la descripción. Un servicio como BORO-01 puede mencionar `10%` en su texto y seguir siendo un elemento normal de Equipo y Herramienta.
- Los porcentajes canónicos de mercado recomendados desde `construdata_matrices.xlsx` son:
  - `%MO1` / `%HERR` Herramienta menor: 3%
  - `%MO2` Andamios: 5%
  - `%MO3` Materiales menores: 5%
  - `%MO5` / `%EPP` Equipo de seguridad/protección personal: 2%
- Estos porcentajes afectan solo `Mercado Cantidad`; el valor declarado por el contratista se muestra intacto en sus columnas.
- Si un porcentaje aparece dentro de una sección, `Mercado P.U.` usa el acumulado/subtotal de mercado correspondiente y `Mercado Cantidad` usa primero el porcentaje recomendado de Construdata cuando exista.

## V2.2 - Reglas de cierre para columnas de mercado

- La matriz del contratista se conserva como dato declarado; no se recalculan las columnas del contratista.
- El carril de mercado debe respetar siempre el operador declarado en la fila del contratista. Si el operador es `/`, el importe mercado se calcula como `Mercado P.U. / Mercado Cantidad`; si el operador es `*`, se calcula como `Mercado P.U. * Mercado Cantidad`.
- Cuando exista una fila de costo indirecto declarada por el contratista, el carril de mercado debe usar siempre 25% sobre el costo directo de mercado. Este 25% reemplaza cualquier porcentaje indirecto declarado por el contratista solo para el carril de mercado.
- Los cargos de utilidad o financiamiento se mantienen con la lógica declarada mientras no exista una regla de mercado específica equivalente.

## V2.3 - Regla de primera hoja y vínculo catálogo ↔ matriz

Para evitar mezclar información entre pestañas de un mismo workbook, todo archivo XLSX cargado se procesa usando únicamente la primera hoja visible del archivo, sin importar si la hoja se llama `Sheet1`, `PU`, `CATALOGO`, `Comparativa` u otro nombre. Las hojas secundarias se ignoran dentro del flujo de carga normal.

La `Comparativa` se alimenta exclusivamente desde conceptos válidos del catálogo: unidad declarada, cantidad mayor a cero y precio unitario o importe calculable. La matriz/APU se usa para explicar y calcular el carril de mercado, pero no para crear filas de comparativa.

Cuando los códigos del catálogo no coinciden con los códigos del análisis PU, el modelo crea un vínculo canónico entre concepto y análisis. El vínculo intenta primero código exacto y, si no existe, usa orden relativo con validación de unidad, cantidad e importe. Esto resuelve casos como catálogo `1.1.1` vinculado al análisis `BS.01`.

La entidad lógica de vínculo queda representada por:

- `concept_code`
- `analysis_code`
- `concept_order`
- `analysis_order`
- `match_strategy`
- `match_confidence`
- `unit_match`
- `quantity_match`
- `amount_match`

El precio unitario de mercado se propaga así:

```text
Primer sheet del catálogo -> CanonicalConcept
Primer sheet de matriz/APU -> CanonicalApuItem / análisis PU
ConceptApuLink -> asocia concepto con análisis
Precio unitario mercado del análisis -> Mercado P.U. en Comparativa
Mercado Importe -> Cantidad catálogo * Mercado P.U.
```

## V2.5 - Matching de materiales contra Construdata

El catálogo granular de referencias debe cargarse antes del cálculo de mercado. Los archivos de materiales, mano de obra y maquinaria pueden no reportar `max_row` cuando se abren en modo lectura; por eso la carga canónica debe ser secuencial, usando la primera hoja visible y detección de columnas por encabezado.

### Flujo canónico

1. Leer referencias Construdata granular:
   - materiales
   - mano de obra
   - maquinaria
2. Indexar por tokens de código y descripción.
3. Para cada insumo APU:
   - determinar tipo esperado por sección;
   - generar candidatos;
   - puntuar por tokens, unidad y cobertura;
   - aplicar regla de precio máximo permitido;
   - aceptar referencia o usar fallback contratista.

### Regla de precio máximo

Si `precio_construdata > precio_contratista * 1.25`, el candidato se rechaza por defecto y se conserva el valor del contratista como mercado fallback. Esta regla evita que un match semánticamente plausible pero económicamente desproporcionado contamine el análisis.

### IA / reranking

La IA debe entrar después del generador de candidatos, no como búsqueda libre sobre todo Construdata. El modelo canónico debe entregar top-N candidatos con score, precio y unidad. La IA puede elegir o rechazar el mejor candidato y devolver una justificación auditable.
