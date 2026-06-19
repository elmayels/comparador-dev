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

## V1 Alpha - Reportes con archivos reales

Se agregan endpoints reales:

- `POST /api/comparisons/real-run`
  - `provider_names[]`
  - `concept_files[]`
  - `matrix_files[]`

- `POST /api/base-budgets/real-run`
  - `concepts_file`
  - `matrix_file` opcional

La comparativa conserva formato horizontal por proveedor. El detalle conserva una hoja por proveedor. La hoja `Detalle Base` y las hojas `Detalle - <Proveedor>` usan el mismo layout canónico; para base se omiten columnas de mercado.

Los nombres visibles de proveedor vienen de la UI y se limitan a 10 caracteres para preservar legibilidad en Excel.

## V1.2 Canonical Output Contract - baseline aceptable

Esta corrección restaura como contrato visual mínimo aceptable el formato de `Comparativo_cotizacion_2_proveedores.xlsx`, manteniendo por debajo el modelo canónico y los parsers reales.

### Comparativa

La hoja `Comparativa` se renderiza como comparativo horizontal por proveedor:

- `A:D` son columnas base del catálogo/concepto y quedan congeladas.
- Fila 1 contiene agrupadores: `Servicios / Cotización` y un bloque por proveedor.
- Fila 2 contiene encabezados operativos.
- Cada proveedor aporta exactamente seis columnas:
  - `P.U.`
  - `Importe`
  - `% Part.`
  - `% ajuste`
  - `Mercado P.U.`
  - `Mercado Importe`
- El mercado vive dentro del bloque de cada proveedor, no como un bloque global separado.
- El orden visual respeta el catálogo declarado.
- El Pareto 80/20 se calcula internamente por proveedor y se pinta en azul claro sin reordenar el catálogo.

### Detalle por proveedor

Cada proveedor genera su propia hoja: `Detalle - <Proveedor>`.

No se genera un detalle horizontal multi-proveedor porque cada contratista puede declarar una matriz/APU distinta.

Contrato visual:

- `A:H`: matriz/APU del proveedor.
- `I`: separador visual.
- `J:M`: comparación contra mercado.

Columnas:

- `Código`
- `Concepto`
- `Unidad`
- `P. Unitario`
- `Op.`
- `Cantidad`
- `Importe`
- `%`
- Separador
- `Mercado P. Unitario`
- `Mercado Op.`
- `Mercado Cantidad`
- `Mercado Importe`

### Detalle Base

`Detalle Base` utiliza el mismo writer canónico que `Detalle - <Proveedor>`, pero sin el bloque de mercado (`J:M`) porque la matriz base ya representa la matriz presupuestada/mercado generada desde conceptos base + matrices Construdata.

### Regla canónica

El modelo canónico no debe obligar al Excel a verse como tabla plana. La lógica correcta es:

```text
Archivos XLSX reales
  -> parsers
  -> modelo canónico
  -> cálculo / mercado / validaciones
  -> writer Excel con contrato visual aceptable
```

## Corrección V1.3 - contrato de Comparativa sin detalle de matriz

La hoja `Comparativa` no debe renderizar insumos, materiales, mano de obra, maquinaria ni filas internas de la matriz/APU. Su objetivo es comparar los conceptos del catálogo/oferta a nivel de partida usando únicamente los totales por precio unitario.

Contrato final:

- `Comparativa` se alimenta desde `CanonicalConcept` proveniente del catálogo de conceptos.
- `Comparativa` muestra una fila por partida/concepto de catálogo, en el orden declarado.
- `Comparativa` no consume ni renderiza `CanonicalApuItem`.
- Los insumos de matriz/APU se muestran exclusivamente en `Detalle - <Proveedor>`.
- A:D permanecen congeladas como base del catálogo.
- Cada proveedor agrega su bloque hacia la derecha: `P.U.`, `Importe`, `% Part.`, `% ajuste`, `Mercado P.U.`, `Mercado Importe`.
- El sombreado 80/20 se calcula sobre el importe total por concepto, no sobre insumos.

Si el parser de catálogo recibe accidentalmente una hoja de matriz/APU, debe filtrar filas típicas de detalle como `MATERIALES`, `MANO DE OBRA`, `SUBTOTAL`, `COSTO DIRECTO`, `UTILIDAD`, porcentajes y otros insumos. Esas filas pertenecen al detalle técnico, no al comparativo.

## V1.5 - Columnas de mercado en Comparativa y Detalle

Las columnas `Mercado P.U.` y `Mercado Importe` del tab `Comparativa` deben poblarse desde `CanonicalConcept.market_unit_price` y `CanonicalConcept.market_amount`.

Las columnas de mercado de `Detalle - <Proveedor>` deben poblarse desde `CanonicalApuItem.market_*`.

El Excel writer no debe calcular ni inventar valores de mercado. Solo debe renderizar los valores que ya existan en el modelo canónico.

## V1.6 - Reglas de mercado en Comparativa y Detalle

### Comparativa

- Columnas A:D se mantienen como base de catálogo y no contienen detalle de matriz.
- Cada bloque de proveedor muestra:
  - P.U.
  - Importe
  - % Part.
  - % ajuste
  - Mercado P.U.
  - Mercado Importe
- Las columnas de mercado provienen de `CanonicalConcept.market_unit_price` y `CanonicalConcept.market_amount`, calculadas desde el detalle APU y no desde el Excel writer.
- Si los proveedores tienen catálogos no idénticos, la comparativa usa la unión canónica de conceptos, preservando el orden por proveedor y catálogo.

### Detalle - <Proveedor>

- El bloque A:H muestra la matriz/APU del proveedor.
- El bloque J:M muestra mercado calculado o leído:
  - Mercado P. Unitario
  - Mercado Op.
  - Mercado Cantidad
  - Mercado Importe
- Las filas estructurales también deben tener mercado cuando exista base suficiente:
  - `Importe:`
  - `Volumen:`
  - `SUBTOTAL`
  - `COSTO DIRECTO`
  - `INDIRECTOS`
  - `PRECIO UNITARIO`
- El estilo del detalle es sobrio: encabezados técnicos, secciones grises, subtotales claros y financiero azul suave.

## V1.7 - Contrato de salida PMD/manual

### Comparativa

La hoja `Comparativa` debe mantener el objetivo del análisis manual:

- A:D = `Servicios / Cotización`.
- E:J, K:P, etc. = un bloque por proveedor.
- Cada bloque de proveedor contiene:
  - `P.U.`
  - `Importe`
  - `% Part.`
  - `% ajuste`
  - `Mercado P.U.`
  - `Mercado Importe`
- `Comparativa` no debe mostrar insumos, subtotales de materiales, mano de obra, maquinaria ni sección financiera.
- El sombreado azul de Pareto 80/20 solo se aplica dentro del bloque del proveedor correspondiente. A:D no se sombrean por Pareto.
- Las filas de notas, capítulos y subcapítulos pueden mostrarse, pero no participan en KPIs ni Pareto si no tienen precio unitario o importe.

### Detalle - <Proveedor>

Cada proveedor tiene su propia hoja:

- A:H = matriz/APU del proveedor.
- I = separador visual.
- J:M = mercado / referencia.
- Debe conservar estructura tipo PU:
  - Encabezado de partida/análisis.
  - Secciones.
  - Insumos.
  - Subtotales.
  - Costo directo.
  - Indirectos/utilidad/financiamiento.
  - Precio unitario.

### Performance

El writer reutiliza objetos de estilo para soportar matrices grandes tipo PMD sin inflar el archivo ni ralentizar el guardado.
