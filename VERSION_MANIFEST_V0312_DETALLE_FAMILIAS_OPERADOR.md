# VERSION_MANIFEST_V0312_DETALLE_FAMILIAS_OPERADOR

## Objetivo
Corregir el tab `Detalle` para representar correctamente la estructura PU/APU:

- Agrupacion por familias dentro de cada matriz/servicio.
- Subtotales por familia.
- Seccion financiera por servicio.
- Correccion de calculo de `Mercado Importe` respetando operador real.

## Cambios principales

### Contratistas
Archivo modificado: `backend/professional_mvp.py`

- `Detalle` conserva las 13 columnas aprobadas.
- Se agrupa por servicio y familia:
  - MATERIALES
  - MANO DE OBRA
  - EQUIPO Y HERRAMIENTA
  - BASICOS
  - OTROS
- Se agrega subtotal por familia.
- Se conserva seccion financiera.
- `Mercado Importe` se recalcula con operador normalizado:
  - `*`: Mercado P.U. x Mercado Cantidad
  - `/`: Mercado P.U. / Mercado Cantidad
  - `%`: base aplicable x porcentaje
- No se copia valor contratista como mercado cuando no hay match.

### Matriz base / catalogo Nestle
Archivo modificado: `backend/base_matrix_proposer.py`

- `Detalle` conserva solo 8 columnas, sin mercado.
- No genera `Analisis IA`.
- La matriz generada se trata como referencia de mercado.
- Se agrupa por familia y subtotal.
- La seccion financiera considera Costo Directo, Costo Indirecto, Financiamiento, Utilidad/Cargos y Total Costo Unitario.

## Validacion realizada

- `python -m py_compile backend/professional_mvp.py backend/base_matrix_proposer.py main.py backend/main.py processor.py`
- Prueba matriz base:
  - archivo: `/mnt/data/test_v0312_detalle_familias_base.xlsx`
  - hojas: `Comparativa`, `Detalle`
  - Detalle: 87 filas, 8 columnas, sin columnas Mercado.
- Prueba contratista unico:
  - archivo: `/mnt/data/test_v0312_detalle_familias_single.xlsx`
  - hojas: `Comparativa`, `Detalle`, `Analisis IA`
  - Detalle: 110 filas, 13 columnas, con columnas Mercado.

## Nota
La prueba multi-contratista completa puede tardar demasiado en este entorno, pero la correccion de `Detalle` se implemento en el writer comun usado por cada `Detalle - Pn`.
