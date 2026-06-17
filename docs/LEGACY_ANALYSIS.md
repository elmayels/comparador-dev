# Análisis histórico del proyecto viejo

## Elementos rescatados

- Lenguaje visual de dashboard oscuro, cards, KPIs y tablas.
- Idea de soportar tema profesional light/dark, ahora implementado con variables CSS.
- Generación Excel con dos hojas principales: `Comparativa` y `Detalle`.
- Reglas conceptuales de operadores `*`, `/` y `%`.
- Separación de costos directos, indirectos, utilidad, financiamiento y componentes APU.
- Carpeta `data` como fuente real de referencias.

## Elementos descartados como base final

- Frontend monolítico `frontend/index.html`.
- Login hardcodeado antiguo.
- `processor.py` como arquitectura central.
- Funciones redefinidas y parches acumulados.
- Mezcla entre presupuesto base, comparador y detalle de contratistas.
- Jobs en filesystem como sustituto de historial real por usuario.

## Decisión de arquitectura V0

La V0 nace limpia. El proyecto viejo se conserva solo como referencia histórica. La carpeta `data` se copia porque contiene fuentes reales de referencia.
