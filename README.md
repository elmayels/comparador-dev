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


## Railway note

This V0 intentionally does not require pandas. Railway/Railpack may default to Python 3.13, and older pandas builds can fail there. The app uses openpyxl for the mock Excel generation. A `.python-version` file is included with Python 3.12 as a safer runtime target.

Start command:

```bash
python run.py
```
