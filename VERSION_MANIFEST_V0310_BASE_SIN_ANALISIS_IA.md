# v0.3.10 - Matriz base sin Analisis IA

## Objetivo

Ajuste incremental sobre v0.3.9 para el flujo de matriz base / catalogo base Nestle.

## Cambio aplicado

Cuando el Excel se genera unicamente desde matriz base / catalogo base Nestle, la salida aprobada queda solamente con:

- Comparativa
- Detalle

No se genera ni se conserva el tab `Analisis IA` en este flujo.

## Comportamiento que se mantiene

- `Detalle` de matriz base se sigue poblando con datos reales.
- `Detalle` de matriz base no incluye columnas de mercado.
- Los flujos de contratistas reales conservan el comportamiento aprobado, incluido `Analisis IA` cuando aplique.
- No se agregan nuevos tabs.
