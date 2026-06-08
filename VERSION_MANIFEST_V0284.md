# VERSION_MANIFEST_V0284

Version: v0.2.8.4-multi-tabs-fix
Base: v0.2.8.3-clean

## Objetivo
Corregir el flujo multi-proveedor rapido para que el Excel de salida siempre incluya las hojas visibles fundamentales:

- Resumen Profesional
- Matriz MultiProveedor
- Comparativa
- Detalle

## Archivos modificados

- backend/professional_mvp.py

## Cambios

- Se agrega `Comparativa` al workbook generado por `create_professional_mvp_workbook`.
- Se agrega `Detalle` al workbook generado por `create_professional_mvp_workbook`.
- `Comparativa` deja de tratarse como hoja interna/oculta en la presentacion para analistas.
- Se mantiene oculto:
  - Alcance y Riesgos
  - Trazabilidad Tecnica
- Se agregan colores pastel por proveedor en columnas/bloques multi.
- Se conserva el encabezado actual de `Detalle`:
  - Codigo
  - Concepto / Insumo
  - Unidad
  - Costo Contratista
  - Op.
  - Cantidad
  - Importe Contratista
  - Costo Mercado
  - Dif % Costo
  - Importe Mercado
  - Base Mercado %
  - Tipo
  - Match Construdata
  - Conf.
  - Estado
  - Nota

## No cambia

- Calculos
- PU
- Importes
- Matching Construdata
- Reglas de mano de obra/cuadrillas
- Motor de matriz propuesta
- Frontend
- Endpoints

## Validacion

- python -m py_compile backend/professional_mvp.py main.py backend/main.py processor.py
- Se genero Excel de prueba multi con:
  - Resumen Profesional visible
  - Matriz MultiProveedor visible
  - Comparativa visible
  - Detalle visible
  - Alcance y Riesgos oculto
  - Trazabilidad Tecnica oculta
