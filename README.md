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

## V1.2 - Canonical Excel baseline restored

This version restores the minimum accepted Excel visual contract while keeping the V1 real-data canonical model.

Key points:

- `Comparativa` follows the accepted horizontal provider format: A:D frozen base columns plus one six-column block per provider.
- Provider blocks contain `P.U.`, `Importe`, `% Part.`, `% ajuste`, `Mercado P.U.`, `Mercado Importe`.
- Pareto 80/20 is calculated from the canonical concept spine and highlighted in blue without changing the declared catalog order.
- `Detalle - <Proveedor>` uses one sheet per provider. Each sheet is A:H provider matrix, I separator, J:M market block.
- `Detalle Base` uses the same canonical writer without market columns.
- This preserves the canonical model internally without forcing the Excel to become a flat technical table.

### V1.3 - Corrección de objetivo del tab Comparativa

Se corrige el contrato del Excel para que `Comparativa` muestre solamente conceptos del catálogo y totales de P.U. por proveedor. El detalle de matriz/APU queda exclusivamente en `Detalle - <Proveedor>`.


## V1.5 - Corrección mercado canónico

Se corrigió la propagación de valores de mercado para que no queden como columnas visuales vacías:

- `Comparativa` toma mercado desde `CanonicalConcept.market_*`.
- `Detalle - <Proveedor>` toma mercado desde `CanonicalApuItem.market_*`.
- El parser de matriz/APU puede leer columnas de mercado declaradas en la matriz.
- El modelo conserva `concept_key` para asociar insumos y totales APU con el concepto del catálogo.

## V1.7 - PMD manual baseline

La V1.7 toma como referencia funcional los archivos PMD manuales usados actualmente:

- Comparativa = resumen por concepto/partida.
- Detalle = matriz/APU que explica el precio unitario.
- Mercado en Comparativa = resultado final de mercado derivado del detalle, no cálculo aislado.
- Parser de matriz soporta layout estándar A:H/J:M y layout ancho tipo TAPIAL.
- Se optimizó el writer de detalle para matrices grandes.


### V1.8 - Corrección de conceptos válidos y fallback de mercado

Esta versión corrige el caso `prov3`, donde los archivos venían con el PU/APU en el slot de conceptos y el resumen de conceptos en el slot de matriz. El motor ahora clasifica el rol del archivo por estructura y no por nombre.

Reglas aplicadas:

- `Comparativa` solo incluye conceptos con unidad y cantidad mayor a cero.
- `Comparativa` no incluye insumos ni líneas financieras de matriz.
- `Detalle - <Proveedor>` conserva la matriz/APU estructurada.
- Si no hay referencia Construdata, mercado usa el valor del contratista como fallback trazable.

### V1.9 - Detalle: negritas por referencia real de mercado

Las columnas de mercado en `Detalle - <Proveedor>` ahora distinguen visualmente los valores con referencia real frente a los valores fallback. Cuando el mercado usa el mismo valor del contratista por falta de referencia, la celda queda normal; cuando hay valor de referencia o diferencia real, se muestra en negrita.


## V2.1 - Porcentajes Construdata y servicios no porcentuales

- La matriz del contratista se conserva como dato declarado; no se recalcula la columna del contratista.
- Las columnas de mercado calculan únicamente el carril de referencia.
- La detección de porcentajes ya no usa cualquier símbolo `%` dentro de la descripción. Un servicio como BORO-01 puede mencionar `10%` en su texto y seguir siendo un elemento normal de Equipo y Herramienta.
- Los porcentajes canónicos de mercado recomendados desde `construdata_matrices.xlsx` son:
  - `%MO1` / `%HERR` Herramienta menor: 3%
  - `%MO2` Andamios: 5%
  - `%MO3` Materiales menores: 5%
  - `%MO5` / `%EPP` Equipo de seguridad/protección personal: 2%
- Estos porcentajes afectan solo `Mercado Cantidad`; el valor declarado por el contratista se muestra intacto en sus columnas.
- Si un porcentaje aparece dentro de una sección, `Mercado P.U.` usa el acumulado/subtotal de mercado correspondiente y `Mercado Cantidad` usa primero el porcentaje recomendado de Construdata cuando exista.

### V2.2 - Ajustes de cierre de comparativa

- El carril de mercado respeta el operador declarado por fila (`*`, `/`, etc.).
- El costo indirecto de mercado se calcula siempre con 25% cuando exista indirecto declarado.
- Se mantiene intacta la matriz del contratista; solo se calcula el carril de mercado.

### V2.3 - Corrección CAT/PU multi-sheet y vínculo por orden

- Los archivos con varios tabs se procesan usando únicamente la primera hoja visible.
- Se evita mezclar `PU` y `CATALOGO` secundarios dentro del mismo workbook.
- Se agregó vínculo canónico catálogo ↔ matriz cuando los códigos no coinciden.
- Caso validado: catálogo `1.1.1` vinculado con matriz `BS.01`; la `Comparativa` ahora llena `Mercado P.U.` y `Mercado Importe` desde el detalle de mercado.

## V2.5 - Materiales Construdata + motor de candidatos

- Corrige la carga de referencias granulares de Construdata cuando los workbooks reportan `max_row=None` en modo lectura. Esto afectaba especialmente materiales, maquinaria y un archivo de mano de obra.
- `ReferenceCatalog` ahora lee la primera hoja visible secuencialmente y detecta columnas por encabezado: código, descripción, unidad y costo.
- El matching de mercado usa un score híbrido por tokens, unidad y tipo de sección.
- Se agrega una regla de seguridad: si el precio Construdata candidato supera en más de 25% al precio del contratista, se conserva el precio del contratista como fallback y se documenta el candidato rechazado.
- Esta capa queda lista para incorporar IA/LLM como reranker sobre los candidatos, sin cambiar el modelo canónico ni el Excel writer.

### V2.6 - Match Construdata visible en Detalle

Se agrega trazabilidad del match Construdata al modelo canónico y al Excel. Las hojas `Detalle - <Proveedor>` muestran una columna `Match Construdata` inmediatamente después de `Mercado Importe`, con `Código - Descripción` de la referencia usada para el match.


## V2.8 Base budget real endpoint

- El frontend de Nuevo presupuesto base llama `/api/base-budgets/real-run`; ya no navega a una pantalla mock.
- El archivo cargado se envía como `concepts_file` y se parsea con `parse_base_concepts`.
- El reporte descargado usa `build_real_base_report` y contiene conceptos/Detalle Base generados desde el archivo real + `data/construdata_matrices.xlsx`.
- `/api/reports/base` se conserva solo como demo legado; no es el flujo canónico de presupuesto base real.


## V2.9 - Prueba correcta del presupuesto base real

La pantalla `Presupuestos base` ya no muestra historicos mock. Para probar el flujo real:

1. Ejecutar la app con `python run.py`.
2. Entrar a la webapp y usar `Dashboard` -> `Nuevo presupuesto base real`, o abrir la ruta/hash `#base-budgets-new`.
3. Cargar un unico archivo `.xlsx` de catalogo de conceptos, por ejemplo `secador-nestle.xlsx`.
4. Presionar `Generar presupuesto base real`.
5. El frontend debe llamar a `POST /api/base-budgets/real-run` con el campo `concepts_file`.
6. La pantalla de resultado debe descargar desde `/api/real-runs/{run_id}/report`, no desde `/api/reports/base`.

Senales de que el flujo es real:
- El Excel debe contener `Detalle Base`.
- El resumen debe mostrar conceptos reales del archivo, no proyectos demo.
- El endpoint `/api/reports/base` queda solo como legado/demo y no se usa desde la UI de presupuesto base.

## V3.3 - Capa IA para analista APU

Esta versión agrega una capa canónica de análisis profesional para el tab `Análisis IA` y para el summary web.

### Bandera única

```bash
ENABLE_AI_ANALYSIS=1
```

Si la bandera está apagada o no existe API key, el sistema no llama servicios externos y genera un resumen experto determinístico basado únicamente en los datos calculados.

### Proveedores soportados

```bash
AI_ANALYSIS_PROVIDER=openai
OPENAI_API_KEY=...
AI_ANALYSIS_MODEL=gpt-4o-mini
```

O bien:

```bash
AI_ANALYSIS_PROVIDER=anthropic
ANTHROPIC_API_KEY=...
AI_ANALYSIS_MODEL=claude-haiku-4-5
```

### Regla canónica

La IA interpreta datos ya calculados. No modifica importes, operadores, cantidades, precios unitarios, matches Construdata ni validaciones.

### Tipos de corrida soportados

- Presupuesto base.
- Comparativa de un proveedor.
- Comparativa de múltiples proveedores.

El contexto enviado a IA es un JSON compacto con KPIs, cobertura, validaciones, top conceptos y resumen por proveedor. La instrucción exige lenguaje natural, profesional, conciso y sin datos inventados.

## V3.4 - Análisis IA APU mejorado

Esta versión fortalece el resumen del tab `Análisis IA` y el summary de webapp para que el análisis sea útil para un analista de precios unitarios:

- Incluye montos y porcentajes calculados por el modelo canónico.
- Usa nombres cortos de contratistas cuando la corrida es comparativa.
- Reporta sobrecostos contra mercado cuando existen referencias.
- Agrega resumen por Materiales, Mano de obra, Maquinaria/equipo, Básicos e Indirectos.
- Señala fallback, insumos sin referencia y conceptos sin matriz directa.
- Mantiene la bandera única `ENABLE_AI_ANALYSIS`.
- Si IA externa no está disponible, el fallback local conserva el mismo nivel de estructura y datos.

La IA sigue sin modificar precios, importes, operadores, cantidades, porcentajes ni matches Construdata.

## V3.5 - AI APU professional analysis upgrade

- The `Análisis IA` tab was upgraded from a short narrative to a professional APU diagnostic sheet.
- The narrative now must include concrete values when available: totals, direct cost, indirect cost, contractor names, market totals, overcosts, coverage, Materials, Labor, Equipment and Indirects.
- The same tab now includes auditable canonical evidence tables:
  - KPI block.
  - Section cost mix.
  - Critical concepts by impact.
  - Market / Construdata overcost alerts.
  - Provider ranking and provider-level alerts for comparison runs.
- External AI remains controlled by `ENABLE_AI_ANALYSIS=1`; if unavailable, the local expert fallback now produces a richer professional analysis from the same canonical context.

## V3.6 - Prueba real de Análisis IA

La capa de Análisis IA ya puede probarse sin ejecutar un Excel completo mediante:

```bash
GET  /api/ai-analysis/status
POST /api/ai-analysis/test
```

Configuración Anthropic:

```bash
export ENABLE_AI_ANALYSIS=1
export AI_ANALYSIS_PROVIDER=anthropic
export ANTHROPIC_API_KEY="tu_api_key"
export AI_ANALYSIS_MODEL="claude-haiku-4-5"
python run.py
```

Configuración OpenAI:

```bash
export ENABLE_AI_ANALYSIS=1
export AI_ANALYSIS_PROVIDER=openai
export OPENAI_API_KEY="tu_api_key"
export AI_ANALYSIS_MODEL="gpt-4o-mini"
python run.py
```

Prueba rápida:

```bash
curl -X POST http://localhost:8000/api/ai-analysis/test \
  -H "Content-Type: application/json"
```

Resultado esperado con API key válida:

```json
{
  "mode": "EXTERNAL_AI",
  "provider": "anthropic",
  "model": "claude-3-5-sonnet-latest",
  "hasKey": true,
  "sections": [ ... ]
}
```

Si la API falla o no hay key, el sistema no rompe el Excel y devuelve fallback local con `LOCAL_EXPERT_FALLBACK_NO_KEY` o `LOCAL_EXPERT_FALLBACK_AI_ERROR`. Cuando ocurra un error, el tab `Análisis IA` y `/api/ai-analysis/status` muestran el motivo sin exponer secretos.

## V3.7 - AI analysis hard fix

Corrections for the APU AI analysis layer:

- Provider-specific default model resolution. If `AI_ANALYSIS_PROVIDER=anthropic` and no model is explicitly set, the backend now uses `claude-haiku-4-5` instead of sending the OpenAI default `gpt-4o-mini` to Anthropic.
- Robust JSON extraction for AI responses. The backend now accepts strict JSON, fenced JSON, or responses with accidental leading/trailing text, and extracts the JSON object safely.
- JSONDecodeError no longer breaks the run. Invalid/non-JSON provider output is captured in `AI_ANALYSIS_LAST_ERROR` and the workbook uses the local expert fallback.
- `GET /api/ai-analysis/status` exposes `lastError` and `lastProviderResponsePreview` for troubleshooting without leaking API keys.
- The AI prompt is stricter: it must behave as an APU / Neodata / Construdata analyst, use calculated values, contractor names, overcosts, section breakdowns and market alerts.

Recommended Anthropic configuration:

```bash
ENABLE_AI_ANALYSIS=1
AI_ANALYSIS_PROVIDER=anthropic
ANTHROPIC_API_KEY=...
# Optional. If omitted, provider-compatible default is used.
AI_ANALYSIS_MODEL=claude-haiku-4-5
```

Recommended smoke test:

```bash
curl http://localhost:8000/api/ai-analysis/status
curl -X POST http://localhost:8000/api/ai-analysis/test -H "Content-Type: application/json" -d '{}'
```

Expected external mode:

```json
{"mode":"EXTERNAL_AI","provider":"anthropic","hasKey":true,"sections":[...]}
```


### V3.7.1 - Anthropic Haiku and env resolution fix

- `claude-haiku-4-5` is now the default Anthropic model when no explicit model is provided.
- If `AI_ANALYSIS_MODEL` starts with `claude` and no provider is defined, the backend resolves `AI_ANALYSIS_PROVIDER=anthropic` automatically.
- Added local `.env` loading with support for spaces around `=`, for example `AI_ANALYSIS_MODEL = claude-haiku-4-5`.
- Added common provider alias handling (`anthopic`, `claude` -> `anthropic`).
- Provider HTTP responses are parsed defensively; non-JSON provider payloads no longer throw raw `JSONDecodeError` to the user.

Recommended cheap Anthropic setup:

```bash
ENABLE_AI_ANALYSIS=1
AI_ANALYSIS_PROVIDER=anthropic
ANTHROPIC_API_KEY=your_key
AI_ANALYSIS_MODEL=claude-haiku-4-5
```


## V3.8.1 Product-facing diagnostic cleanup
- Product UI and Excel no longer expose AI provider/model/fallback/error technical details.
- Diagnostic remains available in Excel and HTML, but commercial wording focuses on KPIs, market evidence, overcosts, references and review priorities.

## V3.9 - Diagnóstico profesional estructurado

Esta versión rehace la capa de diagnóstico profesional como producto comercial:

- El diagnóstico se genera como JSON estructurado y se renderiza en HTML y Excel.
- La hoja `Análisis IA` ya no depende de texto largo; usa KPIs, tablas, alertas y plan de revisión.
- La vista web `/api/real-runs/{run_id}/ai-report` se presenta como `Diagnóstico profesional`, sin mostrar proveedor IA, modelo, fallback, errores técnicos ni JSON.
- Para comparativa de un solo proveedor, el diagnóstico usa lectura individual contra mercado y no genera ranking artificial.
- Para múltiples proveedores, sí se habilita comparación/ranking económico.
- No se modifican cálculos de precios, operadores, indirectos, Construdata ni comparativa.


## V4.2 - Diagnóstico profesional como resultado principal
- En comparativas, la pantalla posterior a ejecutar ya no muestra "Resultados de comparación".
- La ruta interna `comparison-results` ahora renderiza la vista "Diagnóstico profesional" como resultado principal.
- El diagnóstico se muestra embebido desde `/api/real-runs/{run_id}/ai-report`, con botones para descargar Excel, abrir en nueva pestaña y consultar detalle APU.

## V4.5 - Diagnóstico estable, Excel formal y métricas de riesgo

- La pantalla de resultados de comparativa mantiene el diagnóstico nativo al cambiar tema claro/oscuro; ya no reconstruye la ruta ni pierde la corrida.
- Se persiste la última corrida de comparativa en localStorage para evitar pérdida accidental de botones/enlaces después de cambios visuales.
- El tab `Análisis IA` del Excel ahora funciona como documento formal para el analista de precios unitarios, distinto a la vista HTML/dashboard.
- Se agrega el tab `Métricas de Riesgo`, con el valor, criterio, nivel, razón y acción esperada para justificar por qué el dictamen queda bajo, medio/revisión o alto.
- La vista web sigue siendo el diagnóstico visual principal; el Excel conserva una lectura documental y una explicación auditable de riesgo.

## V4.6 - Indirecto declarado por contratista

- El diagnóstico profesional ahora detecta si el contratista declara un indirecto distinto al 25%.
- La diferencia se muestra en KPIs, alertas contra mercado, plan de revisión, tab `Análisis IA` y `Métricas de Riesgo`.
- El carril de mercado conserva la regla canónica: indirecto de mercado fijo 25%.
- No se modifican cálculos de importes, operadores, Construdata ni comparativa.
