# VERSION_MANIFEST_V037_RESUMEN_CONTRATISTA_COMPARATIVA

Versión: v0.3.7-resumen-contratista-comparativa
Base: v0.3.6-comparativa-mercado-detalle-visual

## Objetivo
Agregar en el tab `Comparativa`, debajo de la fila totalizadora, un bloque textual de `Resumen individual por contratista`, sin modificar la estructura principal de `Comparativa` ni de `Detalle` / `Detalle - Pn`.

## Cambios principales
- Se agrega sección `Resumen individual por contratista` al final de `Comparativa`.
- Se genera un bloque por cada proveedor/contratista.
- El texto se construye con reglas disponibles del análisis:
  - concentración 80/20 por proveedor;
  - partidas de mayor peso económico;
  - ajuste global contra mercado del propio proveedor;
  - partidas con mayor desviación contra mercado;
  - renglones sin match de mercado;
  - hallazgos potenciales en mano de obra, equipo crítico, herramienta menor, EPP e indirectos si existen datos.
- No se crean tabs adicionales.
- Se mantiene el filtro solo sobre la tabla de servicios, no sobre el bloque textual.

## Archivo modificado
- backend/professional_mvp.py

## Alcance no incluido
- No se agrega `Resumen Profesional`.
- No se agrega `Analisis experto IA`.
- No se agrega `Trazabilidad Técnica`.
- No se modifica la estructura de `Detalle`.
