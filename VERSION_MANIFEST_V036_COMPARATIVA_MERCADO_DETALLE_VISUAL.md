# VERSION_MANIFEST_V036_COMPARATIVA_MERCADO_DETALLE_VISUAL

Fecha: 2026-06-08
Base: v0.3.5-detalle-mercado-real

## Objetivo

Corrección enfocada solo en:

1. Consistencia del mercado entre `Detalle` y `Comparativa`.
2. Homologación del azul 80/20 en `Comparativa` a azul claro.
3. Reemplazo de fondo amarillo por negritas en diferencias del lado mercado en `Detalle`.

No se agregan tabs nuevos.
No se implementan todavía `Resumen Profesional`, `Analisis experto IA` ni `Trazabilidad`.

## Cambios técnicos

Archivo modificado:

- `backend/professional_mvp.py`

### Comparativa

- `Mercado P.U.` y `Mercado Importe` ahora se alimentan desde el mercado consolidado de los tabs `Detalle` / `Detalle - Pn`.
- Se agregó cache de detalle por proveedor para evitar recalcular y para garantizar que `Comparativa` y `Detalle` usen la misma fuente.
- Si una partida no tiene match real de mercado en el detalle, `Comparativa` deja mercado vacío y no calcula `% ajuste`.
- No se copia P.U. del contratista como mercado.
- No se usa mercado global compartido.

### Azul 80/20

- `COMPARATIVA_BLUE_80` cambia de `9DC3E6` a `D9EAF7`.
- La regla sigue aplicando por proveedor y solo dentro de su bloque.

### Detalle

- Se eliminó el fondo amarillo para diferencias contra mercado.
- Las diferencias relevantes se marcan con negritas únicamente en columnas de mercado:
  - `Mercado P. Unitario`
  - `Mercado Cantidad`
  - `Mercado Importe`
- Se mantienen tolerancias:
  - P.U.: >0.5% o >$1.00
  - Cantidad: >0.0001
  - Importe: >0.5% o >$1.00

## Validación controlada

Archivo generado:

- `/mnt/data/test_v036_comparativa_detalle_dummy.xlsx`

Validaciones:

- Hojas: `Comparativa`, `Detalle - P1`, `Detalle - P2`.
- Mercado en `Comparativa` coincide con mercado consolidado desde el detalle inyectado por proveedor.
- Filas sin match dejan mercado vacío.
- Azul 80/20 usa `D9EAF7`.
- No hay celdas con fondo amarillo `FFF2CC` en los tabs de detalle.

## Compilación

Compilado correctamente:

- `backend/professional_mvp.py`
- `main.py`
- `backend/main.py`
- `processor.py`
- `backend/base_matrix_proposer.py`
