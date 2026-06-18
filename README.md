# Plataforma Canónica APU / Presupuesto Base - Versión 0

Esta es una V0 limpia y navegable de la plataforma canónica. Usa el proyecto viejo solo como referencia histórica y conserva la carpeta `data` con los archivos reales de referencia.

## Qué incluye

- Backend FastAPI mínimo.
- Frontend estático modular sin dependencias externas.
- Login mock por roles.
- Tema light/dark.
- Módulo independiente de Presupuestos Base.
- Módulo Comparador para propuesta individual y múltiples contratistas.
- Historial de presupuestos/corridas.
- Administración de usuarios.
- Referencias detectadas desde `data`.
- Generación mock de Excel `.xlsx` con tabs `Comparativa` y `Detalle`.
- Separación conceptual de motores:
  - `BaseBudgetEngine`
  - `ContractorMatrixDetailEngine`
  - `ProposalComparisonEngine`

## Usuarios demo

| Rol | Usuario | Password |
|---|---|---|
| Super Admin | `superadmin@demo.com` | `demo123` |
| Admin | `admin@demo.com` | `demo123` |
| Analista | `analista@demo.com` | `demo123` |

## Ejecutar

```bash
pip install -r requirements.txt
python run.py
```

Abrir:

```text
http://localhost:8000
```

## Decisiones canónicas implementadas

1. Los usuarios cargan solo archivos `.xlsx`.
2. Presupuesto base es independiente del proceso de licitación.
3. Presupuesto base usa archivo base de conceptos de ingeniería + `data/construdata_matrices.xlsx`.
4. Comparador funciona con o sin presupuesto base.
5. Detalle APU por contratista usa matriz/APU del contratista + referencias granulares desde `data`.
6. El detalle de contratistas no usa `construdata_matrices.xlsx`.
7. IA en V0 está mockeada y separada como capa interpretativa.
8. Excel mock genera tabs `Comparativa` y `Detalle`.

## Próximo paso recomendado

Conectar los motores reales de parsing/cálculo por etapas, manteniendo esta separación:

- Presupuesto base: conceptos ingeniería + matrices Construdata.
- Detalle APU contratista: matriz contratista + materiales/MO/maquinaria/porcentajes desde `data`.
- Comparador: ranking, desviaciones, KPIs y hallazgos.

## Actualización Excel profesional

La V0 ahora incluye un generador de Excel profesional para `/api/reports/comparison` con hojas ejecutivas y técnicas:

- Resumen Ejecutivo
- Comparativa
- Detalle APU
- Partidas Críticas
- Insumos Críticos
- Validaciones
- Parámetros
- Análisis IA

El diseño es limpio y profesional. El archivo histórico se usa únicamente como referencia funcional, no visual. El tab `Detalle APU` queda modelado como reconstrucción desde la matriz/APU del contratista + referencias granulares de `data`.


## V0.3 - Ajuste de carga y Excel

- Para comparativas, cada contratista debe registrar nombre visible, archivo de conceptos `.xlsx` y archivo matriz/APU `.xlsx`.
- La hoja `Comparativa` mantiene formato horizontal por proveedor.
- El detalle ya no se genera en una única hoja horizontal; se genera un tab por contratista: `Detalle - Proveedor A`, `Detalle - Proveedor B`, etc.
- Cada tab de detalle respeta la matriz/APU propia del contratista y aplica porcentajes sobre el subtotal correspondiente declarado por la matriz.
- Se mantiene eliminado el tab `Parámetros`.

## V1 alpha real-data

Esta entrega agrega lectura real inicial de XLSX:

- Carga por proveedor: nombre corto, catálogo de conceptos y matriz/APU.
- Parser inicial de conceptos y matriz/APU usando `openpyxl`.
- Modelo canónico en `backend/real_data.py`.
- Generación de Excel desde objetos canónicos reales.
- Endpoint de descarga por corrida real: `/api/real-runs/{run_id}/report`.

Limitación intencional: el parser es heurístico y tolerante. La siguiente iteración debe robustecer homologación de conceptos, detección de secciones y matching semántico contra referencias.
