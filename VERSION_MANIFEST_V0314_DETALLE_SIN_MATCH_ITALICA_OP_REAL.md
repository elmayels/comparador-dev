# v0.3.14 - Detalle sin match en italica + operador real

## Alcance
- Correccion enfocada en el tab `Detalle` para flujos con contratistas.
- No modifica `Comparativa` ni `Analisis IA`.
- Matriz base se mantiene sin columnas de mercado y sin `Analisis IA`.

## Cambios
- Si un renglon de contratista no tiene match real de mercado, las columnas de mercado se llenan con los valores del contratista:
  - Mercado P. Unitario = P. Unitario contratista
  - Mercado Op. = Op. contratista
  - Mercado Cantidad = Cantidad contratista
  - Mercado Importe = Importe contratista
- Esas celdas de mercado se formatean en italica para indicar que son valores heredados por falta de match, no mercado real.
- Esos valores heredados no cambian `market_match_real`, por lo que no contaminan la consolidacion de mercado en `Comparativa`.
- La columna `Op.` del contratista se obtiene desde la evidencia del renglon:
  - `formula_kind` porcentual => `%`
  - `dividir` / rendimiento inverso => `/`
  - `op`, `operacion`, `operator`, etc. => operador normalizado
  - fallback `*` solo cuando no hay evidencia contraria.
- `Mercado Importe` sigue calculandose con el operador real cuando existe match.

## Validaciones realizadas
- Compilacion Python OK para:
  - `backend/professional_mvp.py`
  - `backend/base_matrix_proposer.py`
  - `main.py`
  - `backend/main.py`
  - `processor.py`
- Prueba single generada con filas sin match heredadas e italica en columnas J:M.
- Prueba matriz base generada sin columnas de mercado y sin `Analisis IA`.
- Prueba controlada de operadores valida `/`, `%` y fallback sin match en italica.
