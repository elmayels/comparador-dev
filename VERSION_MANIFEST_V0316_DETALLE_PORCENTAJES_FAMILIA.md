# v0.3.16 - Detalle porcentajes por familia y reconciliacion

## Objetivo
Corregir el tab Detalle de contratistas para que los renglones porcentuales se calculen con la base correcta de familia, y asegurar que Detalle, Comparativa y webapp/JSON usen el mismo resultado consolidado.

## Cambios principales
- Se preserva el orden relativo de los renglones dentro de cada familia usando el orden entregado por el parser.
- Se detectan renglones porcentuales por operador, unidad %, código % o descripciones frecuentes: herramienta menor, EPP, andamios, consumibles, corte/soldadura, desperdicio, etc.
- Se evita clasificar roles reales como porcentuales, por ejemplo Supervisor de Seguridad.
- Para porcentajes de familia:
  - Contratista = subtotal contratista de la base detectada x porcentaje.
  - Mercado = subtotal mercado de la base detectada x porcentaje.
- Se distingue mercado derivado por porcentaje de familia contra valores heredados por falta de match.
- Los valores heredados por falta de match siguen en itálica.
- Los valores derivados por porcentaje de familia no se muestran en itálica porque son cálculo de mercado derivado, no copia del contratista.
- La consolidación de Comparativa ahora reconoce el TOTAL COSTO UNITARIO normalizado desde Detalle para reconciliar proveedor/mercado.

## Validaciones
- Single contractor:
  - Detalle vs Comparativa reconciliado en total contratista.
  - Detalle vs Comparativa reconciliado en total mercado con delta <= $0.01 por redondeo.
  - 12 renglones porcentuales, 0 con operador incorrecto.
- Multi contractor:
  - P1 y P2 reconciliados contra Comparativa.
  - 12 renglones porcentuales por proveedor, 0 con operador incorrecto.
- Matriz base:
  - Sin columnas mercado.
  - Sin Analisis IA.
  - Porcentajes base con operador %.

## Archivos modificados
- backend/professional_mvp.py
