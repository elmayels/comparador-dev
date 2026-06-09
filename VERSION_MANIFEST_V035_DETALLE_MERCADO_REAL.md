# VERSION_MANIFEST_V035_DETALLE_MERCADO_REAL

## Versión
v0.3.5-detalle-mercado-real

## Objetivo
Corregir exclusivamente el tab `Detalle` para que las columnas de mercado se alimenten con el match real contra Construdata/granular market del motor existente.

## Cambios funcionales
- `read_provider()` ahora aplica `_step08_apply_market_catalog_pricing()` antes de construir filas profesionales.
- `Detalle` usa `granular_market_items` como fuente prioritaria.
- Las columnas de mercado se llenan desde:
  - `precio_mercado`
  - `op`
  - `cantidad`
  - `importe_mercado`
  - `codigo_mercado` / `descripcion_mercado` / `confidence` / `match_status` para trazabilidad interna.
- Si el matcher solo tiene `sin_match_fallback_contratista`, ya no se copia el costo del contratista al lado mercado.
- En renglones sin match real, el lado mercado queda vacío y `Mercado Op.` muestra `Sin match`.
- Se agregan filas financieras por concepto cuando el motor actual expone `_step36_financial_rows()`:
  - COSTO DIRECTO
  - COSTO INDIRECTO
  - FINANCIAMIENTO
  - UTILIDAD / CARGOS ADICIONALES
  - TOTAL COSTO UNITARIO
- Se calculan diferencias internas por PU e importe cuando existe mercado real.
- El coloreo de diferencias se mantiene únicamente del lado mercado con tolerancias:
  - 0.5% para P. Unitario
  - 0.0001 para Cantidad
  - $1.00 para Importe

## No cambia
- No se agregan `Resumen Profesional`, `Analisis experto IA`, `Trazabilidad` ni tabs adicionales.
- No se modifica la estructura visible de `Detalle`.
- No se modifica la estructura de `Comparativa` de v0.3.4.

## Archivos modificados
- `backend/professional_mvp.py`

## Validación realizada
- Compilación OK:
  - `backend/professional_mvp.py`
  - `main.py`
  - `backend/main.py`
  - `processor.py`
  - `backend/base_matrix_proposer.py`
- Generado Excel de prueba single:
  - `/mnt/data/test_v035_detalle_mercado_single.xlsx`
- Validado que `Detalle` contiene mercado real en matches contra Construdata y `Sin match` en renglones sin match real, sin copiar costos del contratista como mercado.
