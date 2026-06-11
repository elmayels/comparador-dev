"""
Quantia AI Comparador

Arquitectura de datos:
  FIJO en servidor (cambia poco):
    - Precios Neodata → /admin/neodata  (sube Excel exportado de Neodata)

  VARIABLE por licitación (cada corrida):
    - Catálogo del proyecto Nestlé  → campo "catalogo" en /comparar
    - Archivos de proveedores (2-5) → campos archivo1..archivo5 en /comparar
"""

import os, tempfile, traceback, json, shutil, datetime, threading, time, uuid
# STEP28_PERF_TRACKING: IA apagada por defecto para diagnosticar performance.
os.environ.setdefault("QUANTIA_ENABLE_AI", "0")
os.environ.setdefault("ENABLE_AI_MATERIAL_MATCHING", "0")
from pathlib import Path
import sys

# STEP23_IMPORT_BOOTSTRAP: Railway puede ejecutar backend/main.py como script.
# Aseguramos que la raiz del proyecto (/app) tenga prioridad sobre /app/backend
# para que `import processor` cargue /app/processor.py y no un shadow module.
PROJECT_ROOT = Path(__file__).resolve().parent
if PROJECT_ROOT.name == "backend":
    PROJECT_ROOT = PROJECT_ROOT.parent
# Railway suele ejecutar /app/backend/main.py como script, lo que deja /app/backend
# antes que /app en sys.path. Eso hace que `import processor` cargue
# /app/backend/processor.py en lugar del implementation real. Forzamos /app primero.
_project_root_str = str(PROJECT_ROOT)
while _project_root_str in sys.path:
    sys.path.remove(_project_root_str)
sys.path.insert(0, _project_root_str)

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from typing import Optional

try:
    from .material_ai_matcher import get_material_ai_status, run_material_ai_smoke_test, run_network_diagnostics
    from .material_ai_trace import get_ai_usage, reset_ai_usage
    from .material_vector_index import ensure_construdata_vector_index, get_vector_index_status, get_vector_index_dir
    from .material_match_cache import cache_status, clear_material_match_cache
    from .processor import build_comparativo, extract_conceptos, extract_orden_catalogo, extract_precios_nacional, detect_valid_supplier_sheet, match_materials_against_benchmark, summarize_material_market_matches, load_labor_tabulador, build_labor_alias_index, match_labor_against_tabulador, summarize_labor_matches, calculate_executive_findings, build_backend_market_summary, attach_construdata_matrix_analysis, load_equipment_benchmark, build_equipment_index, match_equipment_against_benchmark, summarize_equipment_matches, extract_resumen_pu, apply_resumen_pu_to_conceptos, _step08_apply_market_catalog_pricing
except ImportError:
    from material_ai_matcher import get_material_ai_status, run_material_ai_smoke_test, run_network_diagnostics
    from material_ai_trace import get_ai_usage, reset_ai_usage
    from material_vector_index import ensure_construdata_vector_index, get_vector_index_status, get_vector_index_dir
    from material_match_cache import cache_status, clear_material_match_cache
    from processor import build_comparativo, extract_conceptos, extract_orden_catalogo, extract_precios_nacional, detect_valid_supplier_sheet, match_materials_against_benchmark, summarize_material_market_matches, load_labor_tabulador, build_labor_alias_index, match_labor_against_tabulador, summarize_labor_matches, calculate_executive_findings, build_backend_market_summary, attach_construdata_matrix_analysis, load_equipment_benchmark, build_equipment_index, match_equipment_against_benchmark, summarize_equipment_matches, extract_resumen_pu, apply_resumen_pu_to_conceptos, _step08_apply_market_catalog_pricing

# Capa profesional v0.2: agrega hojas ejecutivas multi-proveedor
# sin cambiar el motor base ni requerir Redis/DB.
try:
    from .professional_mvp import append_professional_mvp_workbook, create_professional_mvp_workbook, build_professional_mvp_analysis, professional_mvp_status
except Exception:
    try:
        from backend.professional_mvp import append_professional_mvp_workbook, create_professional_mvp_workbook, build_professional_mvp_analysis, professional_mvp_status
    except Exception as _mvp_import_exc:
        append_professional_mvp_workbook = None
        create_professional_mvp_workbook = None
        build_professional_mvp_analysis = None
        def professional_mvp_status():
            return {"available": False, "error": f"professional_mvp unavailable: {_mvp_import_exc}"}

# STEP35: import module object for diagnostics/policy; from-import above does not always
# leave a convenient reference and Railway can execute this file as a script.
try:
    import processor as _quantia_processor_module
except Exception:
    _quantia_processor_module = None

def _safe_ai_policy_status():
    try:
        if _quantia_processor_module is not None and hasattr(_quantia_processor_module, 'quantia_ai_policy_status'):
            return _quantia_processor_module.quantia_ai_policy_status()
    except Exception as exc:
        return {'error': f'{type(exc).__name__}: {exc}'}
    return {
        'quantia_enable_ai': os.getenv('QUANTIA_ENABLE_AI', '0'),
        'runtime_mode': os.getenv('QUANTIA_AI_RUNTIME_MODE', 'off'),
        'material_runtime_allowed': False,
        'concept_runtime_allowed': False,
        'expert_text_allowed': False,
        'note': 'processor policy helper unavailable'
    }

app = FastAPI(title="Quantia Comparador APU", version="0.2.5-matriz-propuesta")

# Configuracion profesional. Se mantiene simple para entregar hoy, pero con
# contratos claros y variables de entorno listas para produccion ligera.
APP_ENV = os.getenv("APP_ENV", "development").strip().lower()
MAX_PROVIDERS = int(os.getenv("MAX_PROVIDERS", "5") or "5")
MARKET_INDIRECT_PCT = float(os.getenv("MARKET_INDIRECT_PCT", "0.25") or "0.25")
REFERENCE_SOURCE_LABEL = os.getenv("REFERENCE_SOURCE_LABEL", "Neodata / Construdata")

def _cors_origins():
    raw = os.getenv("CORS_ORIGINS", "*").strip()
    if not raw or raw == "*":
        return ["*"]
    return [x.strip() for x in raw.split(",") if x.strip()]

_CORS_ORIGINS = _cors_origins()

app.add_middleware(
    CORSMiddleware,
    allow_origins=_CORS_ORIGINS,
    allow_credentials=("*" not in _CORS_ORIGINS),
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Rutas de datos persistentes en el servidor ─────────────
def _default_data_dir() -> Path:
    configured = os.getenv("DATA_DIR") or os.getenv("QUANTIA_DATA_DIR")
    if configured:
        return Path(configured)
    prod_data = Path("/data")
    try:
        if prod_data.exists() and os.access(str(prod_data), os.W_OK):
            return prod_data
    except Exception:
        pass
    # Proyecto empacado: la carpeta data vive en la raiz del proyecto.
    here = Path(__file__).resolve()
    project_root = here.parents[1] if here.parent.name == "backend" else here.parent
    return project_root / "data"

DATA_DIR      = _default_data_dir()

def _project_data_dir() -> Path:
    here = Path(__file__).resolve()
    project_root = here.parents[1] if here.parent.name == "backend" else here.parent
    return project_root / "data"

def _resolve_data_file(filename: str) -> Path:
    """Resolve a data artifact with fallback to bundled ./data.

    In Railway/Docker DATA_DIR is usually /data for persistent uploads/jobs.
    Some deployments do not copy catalog files into that volume, while the
    repository still contains ./data/construdata_matrices.xlsx.  For read-only
    catalog artifacts, prefer DATA_DIR when the file exists, otherwise fallback
    to the bundled project data folder.
    """
    configured = DATA_DIR / filename
    if configured.exists():
        return configured
    bundled = _project_data_dir() / filename
    if bundled.exists():
        return bundled
    return configured

NEODATA_PATH  = _resolve_data_file("neodata_precios.xlsx")   # Snapshot benchmark de mercado
CONSTRUDATA_MATRICES_PATH = _resolve_data_file("construdata_matrices.xlsx")  # Base maestra concepto -> matriz
BENCHMARK_META_PATH = _resolve_data_file("benchmark_meta.json")
TABULADOR_MO_PATH = _resolve_data_file("Tabulador de Mano de Obra 2024-2025-2026-1.xlsx")
DATA_DIR.mkdir(parents=True, exist_ok=True)

# STEP29: jobs asincronos persistentes.
# En Railway un job no puede depender solo de memoria: el proceso puede reiniciar,
# el proxy puede balancear a otra instancia, o el browser puede seguir consultando
# despues de un cold start. Por eso el estado se guarda en JSON dentro de JOBS_DIR.
def _default_jobs_dir() -> Path:
    configured = os.getenv("JOBS_DIR") or os.getenv("QUANTIA_JOBS_DIR")
    if configured:
        return Path(configured)
    # Si existe DATA_DIR/QUANTIA_DATA_DIR, usarlo para poder montar Railway Volume.
    data_cfg = os.getenv("DATA_DIR") or os.getenv("QUANTIA_DATA_DIR")
    if data_cfg:
        return Path(data_cfg) / "jobs"
    return Path("/tmp/quantia_jobs")

JOBS_DIR = _default_jobs_dir()
JOBS_DIR.mkdir(parents=True, exist_ok=True)
_COMPARISON_JOBS = {}
_COMPARISON_JOBS_LOCK = threading.Lock()

def _job_dir(job_id: str) -> Path:
    safe = "".join(ch for ch in str(job_id) if ch.isalnum() or ch in "-_")
    return JOBS_DIR / safe

def _job_status_path(job_id: str) -> Path:
    return _job_dir(job_id) / "status.json"

def _write_job_status(job_id: str, job: dict):
    try:
        d = _job_dir(job_id)
        d.mkdir(parents=True, exist_ok=True)
        tmp = d / "status.tmp"
        tmp.write_text(json.dumps(job, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        tmp.replace(_job_status_path(job_id))
    except Exception as exc:
        print(f"[job_status_write_error] {job_id}: {exc}")

def _read_job_status(job_id: str):
    p = _job_status_path(job_id)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"status": "error", "phase": "status_read_error", "error": str(exc), "job_id": job_id}

def _job_set(job_id: str, **kwargs):
    with _COMPARISON_JOBS_LOCK:
        job = _COMPARISON_JOBS.get(job_id) or _read_job_status(job_id) or {"job_id": job_id}
        job.update(kwargs)
        job["job_id"] = job_id
        job["updated_at"] = datetime.datetime.utcnow().isoformat() + "Z"
        _COMPARISON_JOBS[job_id] = dict(job)
        _write_job_status(job_id, job)
        return dict(job)

def _job_get(job_id: str):
    with _COMPARISON_JOBS_LOCK:
        job = _COMPARISON_JOBS.get(job_id)
    if not job:
        job = _read_job_status(job_id)
        if job:
            with _COMPARISON_JOBS_LOCK:
                _COMPARISON_JOBS[job_id] = dict(job)
    return dict(job) if job else None

def _cleanup_old_jobs(max_age_hours: int = 24):
    # No borrar jobs activos. Solo limpia terminados viejos.
    cutoff = time.time() - (max_age_hours * 3600)
    try:
        for child in JOBS_DIR.iterdir():
            try:
                if not child.is_dir() or child.stat().st_mtime >= cutoff:
                    continue
                status_path = child / "status.json"
                status = None
                if status_path.exists():
                    try:
                        status = json.loads(status_path.read_text(encoding="utf-8")).get("status")
                    except Exception:
                        status = None
                if status in {"done", "error", "cancelled", None}:
                    shutil.rmtree(child, ignore_errors=True)
            except Exception:
                pass
    except Exception:
        pass



def _job_trace_path(job_id: str) -> Path:
    return _job_dir(job_id) / "trace.jsonl"

def _job_trace(job_id: str, event: str, **data):
    payload = {"ts": time.time(), "iso": datetime.datetime.utcnow().isoformat() + "Z", "event": event}
    payload.update(data)
    try:
        path = _job_trace_path(job_id)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass
    return payload

def _read_job_trace(job_id: str, limit: int = 400):
    path = _job_trace_path(job_id)
    if not path.exists():
        return []
    rows = []
    try:
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines()[-limit:]:
            try:
                rows.append(json.loads(line))
            except Exception:
                rows.append({"raw": line})
    except Exception as exc:
        rows.append({"event": "trace_read_error", "error": str(exc)})
    return rows

def _job_stage(job_id: str, phase: str, progress: int, **extra):
    _job_trace(job_id, "stage", phase=phase, progress=progress, **extra)
    return _job_set(job_id, phase=phase, progress=progress, **extra)

def _safe_download_filename(base_name: str, extension: str = ".xlsx") -> str:
    """Return an ASCII-safe filename for FileResponse/download headers.

    Avoids spaces, accents, quotes, slashes and other characters that can
    break browser downloads in production.
    """
    import re
    import unicodedata

    name = str(base_name or "Comparativo_cotizacion")
    name = unicodedata.normalize("NFKD", name)
    name = "".join(ch for ch in name if not unicodedata.combining(ch))
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    if not name:
        name = "Comparativo_cotizacion"

    ext = extension if str(extension).startswith(".") else f".{extension}"
    # Keep final filename short enough for headers/filesystems.
    max_base_len = 120 - len(ext)
    return f"{name[:max_base_len]}{ext}"


def _has_upload(upload: Optional[UploadFile]) -> bool:
    return bool(upload is not None and getattr(upload, "filename", None))

def _select_resumen_for_provider(resumen_path: Optional[str], resumen_oficial_path: Optional[str]) -> Optional[str]:
    # En multi-proveedor profesional, el Resumen PU oficial comun tiene prioridad
    # para que todos los proveedores se comparen contra la misma cantidad base.
    return resumen_oficial_path or resumen_path

def _mvp_methodology_payload(resumen_oficial_path: Optional[str] = None) -> dict:
    return {
        "version": "v0.2.4-analistas",
        "base": REFERENCE_SOURCE_LABEL,
        "criterio": "Neodata-style: presupuesto base, matriz APU, insumos, indirectos, comparativo y trazabilidad.",
        "resumen_pu_oficial_comun": bool(resumen_oficial_path),
        "indirecto_mercado_pct": MARKET_INDIRECT_PCT,
        "scoring": {
            "precio": 0.35,
            "cobertura": 0.20,
            "desviacion_mercado": 0.20,
            "calidad_apu": 0.15,
            "riesgo_tecnico": 0.10,
        },
        "notas": [
            "El ranking es recomendacion explicable, no adjudicacion automatica.",
            "Los conceptos faltantes se tratan como riesgo de cobertura.",
            "Los conceptos adicionales deben revisarse antes de sumarse al alcance base.",
        ],
    }


def _append_professional_mvp_safe(output_path, paths, nombres, meta, catalogo_path=None, nacional_path=None, resumen_paths=None, resumen_oficial_path=None):
    if append_professional_mvp_workbook is None:
        return output_path
    if os.getenv("ENABLE_PROFESSIONAL_SHEETS", "1").strip().lower() in {"0", "false", "no"}:
        return output_path
    try:
        return append_professional_mvp_workbook(
            str(output_path),
            paths,
            nombres,
            meta=meta,
            catalogo_path=catalogo_path,
            nacional_path=nacional_path,
            resumen_paths=resumen_paths or [],
            resumen_oficial_path=resumen_oficial_path,
        )
    except Exception as exc:
        print(f"[professional_mvp_sheet_error] {type(exc).__name__}: {exc}")
        return output_path

def _validate_provider_count(n: int):
    if n > MAX_PROVIDERS:
        raise HTTPException(400, f"Máximo {MAX_PROVIDERS} proveedores por corrida en esta versión.")


def _use_unified_analysis_engine(n_proveedores: int) -> bool:
    """Motor único de análisis de cotizaciones.

    v0.3.0: no se separan flujos para 1 proveedor y multi-proveedor.
    Todo análisis con uno o más proveedores pasa por el generador unificado,
    que internamente ejecuta el análisis individual por proveedor y genera un
    solo Excel homologado.
    """
    return n_proveedores >= 1

def _create_unified_analysis_workbook_safe(output_path, paths, nombres, meta, catalogo_path=None, nacional_path=None, resumen_paths=None, resumen_oficial_path=None):
    if create_professional_mvp_workbook is None:
        raise RuntimeError("create_professional_mvp_workbook no esta disponible")
    return create_professional_mvp_workbook(
        str(output_path),
        paths,
        nombres,
        meta=meta,
        catalogo_path=catalogo_path,
        nacional_path=nacional_path,
        resumen_paths=resumen_paths or [],
        resumen_oficial_path=resumen_oficial_path,
    )



# Estado en memoria para reconstruccion asincrona del indice vectorial.
# Evita dejar bloqueado el Admin sin progreso mientras Voyage procesa miles de registros.
_VECTOR_INDEX_JOB = {
    "running": False,
    "job_id": None,
    "phase": "idle",
    "started_at": None,
    "updated_at": None,
    "total_records": 0,
    "records_done": 0,
    "total_batches": 0,
    "batch_no": 0,
    "built_records": 0,
    "skipped_records": 0,
    "progress_pct": 0,
    "error": None,
    "result": None,
}
_VECTOR_INDEX_JOB_LOCK = threading.Lock()

def _set_vector_index_job(**updates):
    with _VECTOR_INDEX_JOB_LOCK:
        _VECTOR_INDEX_JOB.update(updates)
        _VECTOR_INDEX_JOB["updated_at"] = datetime.datetime.utcnow().isoformat() + "Z"
        total = int(_VECTOR_INDEX_JOB.get("total_records") or 0)
        done = int(_VECTOR_INDEX_JOB.get("records_done") or 0)
        if total > 0:
            _VECTOR_INDEX_JOB["progress_pct"] = round(min(100, max(0, done * 100 / total)), 2)

def _get_vector_index_job():
    with _VECTOR_INDEX_JOB_LOCK:
        return dict(_VECTOR_INDEX_JOB)



def _read_benchmark_meta() -> dict:
    if BENCHMARK_META_PATH.exists():
        try:
            return json.loads(BENCHMARK_META_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def _write_benchmark_meta(meta: dict):
    BENCHMARK_META_PATH.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

def get_benchmark_status_payload() -> dict:
    meta = _read_benchmark_meta()
    loaded = NEODATA_PATH.exists()
    source_name = meta.get("source_name", "Neodata")
    source_budget_code = meta.get("source_budget_code", "CONCURSOS / APU")
    last_update = meta.get("loaded_at")
    concepts = meta.get("conceptos_cargados")
    materials = meta.get("materiales_cargados")
    if loaded and (concepts is None or materials is None):
        try:
            refs = extract_precios_nacional(str(NEODATA_PATH))
            concepts = len([k for k in refs.keys() if k != "__materials_index__"])
            materials = len(refs.get("__materials_index__", [])) if isinstance(refs, dict) else 0
        except Exception:
            concepts = concepts or 0
            materials = materials or 0
    concepts = concepts or 0
    materials = materials or 0
    if not last_update and loaded:
        try:
            last_update = datetime.datetime.fromtimestamp(NEODATA_PATH.stat().st_mtime).isoformat()
        except Exception:
            last_update = None
    label = f"{source_name} / {source_budget_code} cargado" if loaded else "Sin referencia cargada"
    return {
        "loaded": loaded,
        "source_name": source_name,
        "source_budget_code": source_budget_code,
        "conceptos_cargados": concepts,
        "materiales_cargados": materials,
        "last_update": last_update,
        "label": label,
        "vector_index": get_vector_index_status(load_materials_index()) if loaded else get_vector_index_status(),
        "material_cache": cache_status(),
        "message": (
            f"{label}. {concepts} conceptos y {materials} materiales disponibles."
            if loaded else "Sin referencia cargada. Sube una exportación del benchmark."
        ),
    }

# ── Utilidades Neodata / preview ───────────────────────────
def load_neodata_refs() -> dict:
    # Paso 07: preferir base maestra Construdata embebida si existe.
    path = CONSTRUDATA_MATRICES_PATH if CONSTRUDATA_MATRICES_PATH.exists() else NEODATA_PATH
    if not path.exists():
        return {}
    try:
        return extract_precios_nacional(str(path))
    except Exception:
        return {}


def load_materials_index():
    refs = load_neodata_refs()
    return refs.get("__materials_index__", []) if isinstance(refs, dict) else []

def load_labor_alias_index():
    if not TABULADOR_MO_PATH.exists():
        return {}
    try:
        return build_labor_alias_index(load_labor_tabulador(str(TABULADOR_MO_PATH), "2026"))
    except Exception:
        return {}

def load_equipment_index():
    try:
        return build_equipment_index(load_equipment_benchmark())
    except Exception:
        return []

def _base_clave(clave: str) -> str:
    return str(clave).split("::")[-1] if clave else ""

def _area_from_clave(clave: str) -> str:
    s = str(clave or "")
    return s.split("::")[0] if "::" in s else "General"


def build_preview_payload_fast(paths: list, nombres: list, catalogo_path: str = None, resumen_paths: list = None):
    """Preview liviano para Railway.

    El preview anterior ejecutaba matching de mercado, matriz Construdata e IA antes
    de generar el Excel. En produccion eso puede tardar varios minutos y Railway
    corta la conexion del browser (499). Este preview solo valida/parsa archivos,
    aplica Resumen PU para cantidades reales y devuelve KPIs ejecutivos rápidos.
    El calculo técnico completo queda exclusivamente en /comparar.
    """
    started = time.time()
    datos = []
    resumen_paths = resumen_paths or []
    total_global = 0.0
    rows = []

    for idx, (path, nombre) in enumerate(zip(paths or [], nombres or [])):
        conceptos = extract_conceptos(path)
        resumen_info = {"rows": 0, "applied": 0, "error": None, "path_received": bool(idx < len(resumen_paths) and resumen_paths[idx])}
        if idx < len(resumen_paths) and resumen_paths[idx]:
            try:
                resumen_rows = extract_resumen_pu(resumen_paths[idx])
                resumen_info["rows"] = len(resumen_rows)
                apply_resumen_pu_to_conceptos(conceptos, resumen_rows)
                resumen_info["applied"] = sum(
                    1 for c in conceptos.values()
                    if isinstance(c, dict) and (c.get("resumen_pu_match") or {}).get("applied")
                )
            except Exception as exc:
                resumen_info["error"] = str(exc)

        proveedor_total = 0.0
        proveedor_rows = []
        for orden, (clave, c) in enumerate(conceptos.items(), start=1):
            if str(clave).startswith('__') or not isinstance(c, dict):
                continue
            cantidad = c.get('cantidad_real_cotizacion', c.get('cantidad', 1))
            pu = c.get('pu')
            total = c.get('total')
            if not isinstance(total, (int, float)) and isinstance(cantidad, (int, float)) and isinstance(pu, (int, float)):
                total = cantidad * pu
            if isinstance(total, (int, float)):
                proveedor_total += float(total)
            proveedor_rows.append({
                'area': _area_from_clave(clave),
                'clave': _base_clave(clave),
                'clave_compuesta': clave,
                'descripcion': c.get('desc', ''),
                'unidad': c.get('unidad', ''),
                'cantidad': cantidad if isinstance(cantidad, (int, float)) else 1,
                'proveedores': {nombre: round(float(pu), 4)} if isinstance(pu, (int, float)) else {},
                'proveedores_total': {nombre: round(float(total), 4)} if isinstance(total, (int, float)) else {},
                'ganador': nombre,
                'ganadores': [nombre],
                'providers_count': 1,
                'comparable': False,
                'ahorro': None,
                'pu_neodata': None,
                'delta_pct_neodata': None,
                'status': 'preview_liviano',
                'resumen_pu_aplicado': bool((c.get('resumen_pu_match') or {}).get('applied')),
                'resumen_pu_fila': (c.get('resumen_pu_match') or {}).get('source_row'),
            })

        total_global += proveedor_total
        datos.append({
            'nombre': nombre,
            'conceptos': conceptos,
            'total_estimado': round(proveedor_total, 2),
            'status': 'ok',
            'resumen_pu': resumen_info,
        })
        rows.extend(proveedor_rows)

    # Pareto 80 sin reordenar visualmente: se calcula por importe, pero se conserva
    # el orden declarado en rows.
    impact = []
    for r in rows:
        val = 0.0
        if r.get('proveedores_total'):
            val = max([v for v in r['proveedores_total'].values() if isinstance(v, (int, float))] or [0.0])
        impact.append((r.get('clave_compuesta'), val))
    sorted_impact = sorted(impact, key=lambda x: x[1], reverse=True)
    cutoff = set()
    acc = 0.0
    for key, val in sorted_impact:
        if total_global <= 0:
            break
        if acc / total_global < 0.80 or not cutoff:
            cutoff.add(key)
            acc += val
    for r in rows:
        r['pareto_80'] = r.get('clave_compuesta') in cutoff

    provider_scores = []
    for d in datos:
        provider_scores.append({
            'nombre': d['nombre'],
            'score': 100.0,
            'total_estimado': d['total_estimado'],
            'sobrecostes': 0,
            'comparables': 0,
            'gana': 0,
            'conceptos': len([k for k in d['conceptos'].keys() if not str(k).startswith('__')]),
            'cobertura': 100.0,
            'resumen_pu': d.get('resumen_pu'),
        })

    executive_analysis = [{
        'proveedor': d['nombre'],
        'titulo': 'Preview rápido',
        'mensaje': 'Vista previa liviana: cantidades del Resumen PU aplicadas. El análisis técnico completo se genera en el Excel.',
        'total': d['total_estimado'],
        'resumen_pu': d.get('resumen_pu'),
    } for d in datos]

    # Para 2+ proveedores, devolver una verdadera matriz comparativa liviana.
    # El preview original listaba cada proveedor por separado; esto ocultaba la
    # comparación lateral. Reutilizamos la capa profesional, que es rápida y
    # trabaja contra una base común tipo Neodata/Resumen PU.
    if len(paths or []) > 1 and build_professional_mvp_analysis is not None:
        try:
            analysis = build_professional_mvp_analysis(paths, nombres, meta={"modo": "preview"}, catalogo_path=catalogo_path, resumen_paths=resumen_paths)
            rows = []
            total_global = 0.0
            for mr in analysis.get("matrix_rows", []):
                proveedores = {}
                proveedores_total = {}
                quoted_totals = []
                for pname, pr in (mr.get("providers") or {}).items():
                    if isinstance(pr.get("pu"), (int, float)):
                        proveedores[pname] = round(float(pr.get("pu")), 4)
                    if isinstance(pr.get("total"), (int, float)):
                        proveedores_total[pname] = round(float(pr.get("total")), 4)
                        quoted_totals.append(float(pr.get("total")))
                if quoted_totals:
                    total_global += max(quoted_totals)
                rows.append({
                    "area": "General",
                    "clave": mr.get("code") or mr.get("base_id"),
                    "clave_compuesta": mr.get("base_id"),
                    "descripcion": mr.get("desc") or "",
                    "unidad": mr.get("unit") or "",
                    "cantidad": mr.get("qty"),
                    "proveedores": proveedores,
                    "proveedores_total": proveedores_total,
                    "ganador": mr.get("best_provider") or "",
                    "ganadores": [mr.get("best_provider")] if mr.get("best_provider") else [],
                    "providers_count": len(analysis.get("providers", [])),
                    "comparable": len(quoted_totals) >= 2,
                    "ahorro": (max(quoted_totals) - min(quoted_totals)) if len(quoted_totals) >= 2 else None,
                    "pu_neodata": mr.get("ref_pu"),
                    "delta_pct_neodata": None,
                    "status": "preview_multiproveedor",
                    "riesgo": mr.get("risk") or "",
                    "pareto_80": False,
                })
            impact = []
            for r in rows:
                val = max([v for v in (r.get("proveedores_total") or {}).values() if isinstance(v, (int, float))] or [0.0])
                impact.append((r.get("clave_compuesta"), val))
            sorted_impact = sorted(impact, key=lambda x: x[1], reverse=True)
            cutoff = set(); acc = 0.0
            for key, val in sorted_impact:
                if total_global <= 0:
                    break
                if acc / total_global < 0.80 or not cutoff:
                    cutoff.add(key); acc += val
            for r in rows:
                r["pareto_80"] = r.get("clave_compuesta") in cutoff
            provider_scores = [{
                "nombre": s.get("provider"),
                "score": s.get("total_score"),
                "total_estimado": s.get("total_bid"),
                "total_comparable": s.get("total_comparable"),
                "cobertura": round((s.get("coverage_pct") or 0) * 100, 2),
                "faltantes": s.get("missing"),
                "unidad": s.get("unit_mismatches"),
                "adicionales": s.get("extras"),
                "dictamen": s.get("recommendation"),
            } for s in analysis.get("provider_scores", [])]
            executive_analysis = [{"proveedor": s.get("nombre"), "titulo": "Preview multi-proveedor", "mensaje": s.get("dictamen"), "total": s.get("total_estimado")} for s in provider_scores]
        except Exception as exc:
            # Si falla el agregado comparativo, mantenemos el preview liviano original.
            executive_analysis.append({"proveedor": "Sistema", "titulo": "Aviso preview multi", "mensaje": f"No se pudo agregar matriz multi-proveedor en preview: {type(exc).__name__}: {exc}", "total": None})

    return {
        'modo': 'preview_liviano',
        'preview_fast': True,
        'elapsed_seconds': round(time.time() - started, 3),
        'areas': list(dict.fromkeys([r.get('area') or 'General' for r in rows])),
        'rows': rows,
        'provider_scores': provider_scores,
        'executive_analysis': executive_analysis,
        'market_summary': [],
        'market_totals': {'proveedor': round(total_global, 2), 'mercado': None, 'diferencia': None},
        'message': 'Preview liviano para evitar timeouts 499. El calculo completo corre en /comparar.'
    }


def build_preview_payload(paths: list, nombres: list, catalogo_path: str = None, resumen_paths: list = None):
    refs = load_neodata_refs()
    materials_index = refs.get("__materials_index__", []) if isinstance(refs, dict) else []
    labor_alias_index = load_labor_alias_index()
    equipment_index = load_equipment_index()
    datos = []
    provider_analyses = []
    resumen_paths = resumen_paths or []
    resumen_by_path = {}
    for fp, rp in zip(paths or [], resumen_paths):
        if rp:
            try:
                resumen_by_path[str(fp)] = extract_resumen_pu(rp)
            except Exception as exc:
                resumen_by_path[str(fp)] = {"__error__": str(exc)}
    for path, nombre in zip(paths, nombres):
        conceptos = extract_conceptos(path)
        resumen_rows = resumen_by_path.get(str(path))
        if isinstance(resumen_rows, list) and resumen_rows:
            try:
                apply_resumen_pu_to_conceptos(conceptos, resumen_rows)
            except Exception:
                pass
        # Paso 17: la web debe calcular mercado y totales con cantidades reales
        # antes de construir ranking, hallazgos y resumen ejecutivo.
        try:
            _step08_apply_market_catalog_pricing(conceptos)
        except Exception:
            pass
        try:
            attach_construdata_matrix_analysis(conceptos, refs)
        except Exception:
            pass
        benchmark_is_matrix = isinstance(refs, dict) and str(refs.get("__benchmark_kind__", "")).startswith("construdata_matrices")
        material_match_cache = {}
        important_scope = None
        if not benchmark_is_matrix:
            try:
                from material_market_search import optimize_material_search_scope
                important_scope = optimize_material_search_scope(conceptos)
            except Exception:
                important_scope = None
        for clave, c in conceptos.items():
            if benchmark_is_matrix:
                c["materiales_benchmark_matches"] = []
                c["materiales_benchmark_resumen"] = summarize_material_market_matches([])
                c["mano_obra_benchmark_matches"] = []
                c["mano_obra_benchmark_resumen"] = summarize_labor_matches([])
                c["equipo_benchmark_matches"] = []
                c["equipo_benchmark_resumen"] = summarize_equipment_matches([])
            else:
                c["materiales_benchmark_matches"] = match_materials_against_benchmark(c.get("materiales_items"), materials_index, important_scope, clave, material_match_cache)
                c["materiales_benchmark_resumen"] = summarize_material_market_matches(c["materiales_benchmark_matches"])
                c["mano_obra_benchmark_matches"] = match_labor_against_tabulador(c.get("mano_obra_items"), labor_alias_index)
                c["mano_obra_benchmark_resumen"] = summarize_labor_matches(c["mano_obra_benchmark_matches"])
                c["equipo_benchmark_matches"] = match_equipment_against_benchmark(c.get("equipo_items"), equipment_index)
                c["equipo_benchmark_resumen"] = summarize_equipment_matches(c["equipo_benchmark_matches"])
        total = sum(c.get("total", 0) for c in conceptos.values() if isinstance(c.get("total"), (int, float)))
        total = round(total, 2)
        datos.append({
            "nombre": nombre,
            "conceptos": conceptos,
            "total_estimado": total,
            "status": "ok",
        })
        provider_analyses.append(calculate_executive_findings(nombre, conceptos, total, []))

    orden = []
    if catalogo_path:
        try:
            orden = extract_orden_catalogo(catalogo_path)
        except Exception:
            orden = []

    vistas = set()
    rows = []
    all_keys = []

    if orden:
        for item in orden:
            if item.get("tipo") == "concepto":
                k = item.get("clave")
                if k and k not in vistas:
                    vistas.add(k)
                    all_keys.append(k)
        for d in datos:
            for k in d["conceptos"].keys():
                if k not in vistas:
                    vistas.add(k)
                    all_keys.append(k)
    else:
        for d in datos:
            for k in d["conceptos"].keys():
                if k not in vistas:
                    vistas.add(k)
                    all_keys.append(k)

    areas = []
    area_seen = set()
    for k in all_keys:
        desc = ""
        unidad = ""
        cantidad = 0
        proveedores = {}
        provider_totals = {}
        min_pu = None
        min_names = []
        max_pu = None
        conceptos_por_proveedor = {}
        provider_count = 0

        for d in datos:
            c = d["conceptos"].get(k)
            if not c:
                continue
            conceptos_por_proveedor[d["nombre"]] = c
            provider_count += 1
            if not desc:
                desc = c.get("desc", "")
            if not unidad:
                unidad = c.get("unidad", "")
            if not cantidad:
                cantidad = c.get("cantidad", 0)
            pu = c.get("pu")
            total = c.get("total")
            if isinstance(pu, (int, float)):
                proveedores[d["nombre"]] = round(float(pu), 4)
                provider_totals[d["nombre"]] = round(float(total), 4) if isinstance(total, (int, float)) else None
                if min_pu is None or pu < min_pu:
                    min_pu = pu
                    min_names = [d["nombre"]]
                elif min_pu is not None and abs(pu - min_pu) < 1e-9:
                    min_names.append(d["nombre"])
                if max_pu is None or pu > max_pu:
                    max_pu = pu

        area = _area_from_clave(k)
        if area not in area_seen:
            area_seen.add(area)
            areas.append(area)

        ref = refs.get(k) or refs.get(_base_clave(k))
        pu_neodata = ref.get("pu_mercado") if ref else None
        delta_pct_neodata = ((min_pu - pu_neodata) / pu_neodata) if (pu_neodata and min_pu is not None) else None
        ahorro = (max_pu - min_pu) if (max_pu is not None and min_pu is not None and max_pu != min_pu) else None

        status = "ok"
        if provider_count == 1 and len(datos) > 1:
            status = "exclusivo"
        if delta_pct_neodata is not None and delta_pct_neodata > 0.15:
            status = "sobrecoste"
        elif delta_pct_neodata is not None and delta_pct_neodata < -0.10:
            status = "debajo_referencia"

        benchmark_owner = min_names[0] if min_names else None
        best_concept = conceptos_por_proveedor.get(benchmark_owner, {})
        rows.append({
            "area": area,
            "clave": _base_clave(k),
            "clave_compuesta": k,
            "descripcion": desc,
            "unidad": unidad,
            "cantidad": cantidad,
            "proveedores": proveedores,
            "proveedores_total": provider_totals,
            "ganador": min_names[0] if len(min_names) == 1 else ("Empate" if min_names else None),
            "ganadores": min_names,
            "providers_count": provider_count,
            "comparable": provider_count >= 2,
            "ahorro": round(ahorro, 2) if ahorro is not None else None,
            "pu_neodata": round(pu_neodata, 4) if isinstance(pu_neodata, (int, float)) else None,
            "delta_pct_neodata": round(delta_pct_neodata, 4) if isinstance(delta_pct_neodata, (int, float)) else None,
            "status": status,
            "materiales_benchmark_resumen": best_concept.get("materiales_benchmark_resumen"),
            "materiales_benchmark_matches": best_concept.get("materiales_benchmark_matches"),
            "mano_obra_benchmark_resumen": best_concept.get("mano_obra_benchmark_resumen"),
            "mano_obra_benchmark_matches": best_concept.get("mano_obra_benchmark_matches"),
            "equipo_benchmark_resumen": best_concept.get("equipo_benchmark_resumen"),
            "equipo_benchmark_matches": best_concept.get("equipo_benchmark_matches"),
            "indirecto_pct": round((best_concept.get("subtotal2") / best_concept.get("costo_directo")), 4) if best_concept.get("subtotal2") and best_concept.get("costo_directo") else None,
        })

    provider_scores = []
    if datos:
        min_total = min([d["total_estimado"] for d in datos if d["total_estimado"] is not None] or [0])
        for d in datos:
            nombre = d["nombre"]
            total = d["total_estimado"] or 0
            precio_score = 40 if (min_total and total == min_total) else (max(0, 40 * (min_total / total)) if total else 0)
            comparables = sum(1 for r in rows if r["comparable"] and nombre in r["proveedores"])
            sobrecostes = sum(1 for r in rows if nombre in r["proveedores"] and r.get("status") == "sobrecoste")
            gana = sum((1 / len(r.get("ganadores") or [])) for r in rows if nombre in (r.get("ganadores") or []))
            cobertura = (len(d["conceptos"]) / len(all_keys) * 100) if all_keys else 0
            score = round(precio_score + min(20, gana / max(len(rows),1) * 20) + min(10, cobertura / 10) + max(0, 30 - sobrecostes * 2), 1)
            provider_scores.append({
                "nombre": nombre,
                "score": score,
                "total_estimado": d["total_estimado"],
                "sobrecostes": sobrecostes,
                "comparables": comparables,
                "gana": round(gana, 2),
                "conceptos": len(d["conceptos"]),
                "cobertura": round(cobertura, 1),
            })
    provider_scores.sort(key=lambda x: (-x["score"], x["total_estimado"] if x["total_estimado"] is not None else float("inf")))

    market_summary = build_backend_market_summary(datos)
    market_by_provider = {m.get('proveedor'): m for m in market_summary if isinstance(m, dict)}
    for ps in provider_scores:
        ms = market_by_provider.get(ps.get('nombre')) or {}
        if ms:
            ps['total_proveedor'] = ms.get('total_proveedor')
            ps['total_oferta'] = ms.get('total_proveedor')
            ps['total_mercado'] = ms.get('total_mercado')
            ps['diferencia_total'] = ms.get('diferencia_total')
            ps['desviacion_total_pct'] = ms.get('diferencia_total_pct')

    return {
        "archivos": [{
            "archivo": d["nombre"],
            "conceptos": len(d["conceptos"]),
            "total_estimado": d["total_estimado"],
            "status": d["status"],
        } for d in datos],
        "areas": areas,
        "rows": rows[:500],
        "provider_scores": provider_scores,
        "executive_analysis": provider_analyses,
        "market_summary": market_summary,
        "neodata": "disponible" if refs else "no cargado",
        "neodata_last_update": (
            datetime.datetime.fromtimestamp(NEODATA_PATH.stat().st_mtime).isoformat()
            if NEODATA_PATH.exists() else None
        ),
        "neodata_version_label": f"{REFERENCE_SOURCE_LABEL} cargado" if refs else "Sin referencia cargada",
        "modo": "1_vs_nacional" if len(datos) == 1 else "multi_proveedor",
        "advertencia": ("Con 1 proveedor se necesita benchmark de mercado. Súbelo en /benchmark/upload"
                        if len(datos) == 1 and not refs else None),
        "metodologia": _mvp_methodology_payload(),
        "reference_source": REFERENCE_SOURCE_LABEL,
    }



# ── Health check ────────────────────────────────────────────
@app.get("/debug/runtime")
def debug_runtime():
    import sys, os, platform
    proc_file = None
    try:
        proc_file = getattr(_quantia_processor_module, "__file__", None) if _quantia_processor_module is not None else None
    except Exception:
        proc_file = None
    return {
        "status": "ok",
        "cwd": os.getcwd(),
        "python": sys.version,
        "platform": platform.platform(),
        "data_dir": os.getenv("DATA_DIR"),
        "jobs_dir": str(JOBS_DIR),
        "jobs_dir_exists": JOBS_DIR.exists(),
        "preview_deep": os.getenv("PREVIEW_DEEP", "0"),
        "processor_file": proc_file,
        "ai_policy": _safe_ai_policy_status(),
        "ai_env": {
            "QUANTIA_ENABLE_AI": os.getenv("QUANTIA_ENABLE_AI"),
            "QUANTIA_AI_RUNTIME_MODE": os.getenv("QUANTIA_AI_RUNTIME_MODE"),
            "ENABLE_AI_CONCEPT_MATCHING": os.getenv("ENABLE_AI_CONCEPT_MATCHING"),
            "ENABLE_AI_EXPERT_TEXT": os.getenv("ENABLE_AI_EXPERT_TEXT"),
            "ENABLE_AI_MATERIAL_MATCHING": os.getenv("ENABLE_AI_MATERIAL_MATCHING"),
            "ANTHROPIC_MODEL": os.getenv("ANTHROPIC_MODEL"),
            "VOYAGE_MODEL": os.getenv("VOYAGE_MODEL"),
            "ANTHROPIC_API_KEY_SET": bool(os.getenv("ANTHROPIC_API_KEY")),
            "VOYAGE_API_KEY_SET": bool(os.getenv("VOYAGE_API_KEY")),
        },
        "routes": sorted([getattr(r, "path", "") for r in app.routes]),
    }


@app.get("/debug/jobs")
def debug_jobs():
    rows = []
    try:
        for d in sorted(JOBS_DIR.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True)[:50]:
            if not d.is_dir():
                continue
            st = d.stat()
            status = None
            p = d / "status.json"
            if p.exists():
                try:
                    status = json.loads(p.read_text(encoding="utf-8"))
                except Exception as exc:
                    status = {"status": "status_read_error", "error": str(exc)}
            rows.append({
                "job_id": d.name,
                "mtime": datetime.datetime.fromtimestamp(st.st_mtime).isoformat(),
                "status": (status or {}).get("status"),
                "phase": (status or {}).get("phase"),
                "progress": (status or {}).get("progress"),
                "updated_at": (status or {}).get("updated_at"),
                "has_trace": (d / "trace.jsonl").exists(),
                "has_output": (d / "comparativo.xlsx").exists(),
            })
    except Exception as exc:
        return {"ok": False, "jobs_dir": str(JOBS_DIR), "error": str(exc)}
    return {"ok": True, "jobs_dir": str(JOBS_DIR), "jobs": rows}

@app.get("/debug/ping")
def debug_ping():
    return {"ok": True, "message": "backend reached"}

@app.get("/health")
@app.get("/api/health")
def health():
    bench = get_benchmark_status_payload()
    return {
        "status": "ok",
        "service": "Quantia AI Comparador",
        "neodata": "cargado" if bench["loaded"] else "no cargado",
        "neodata_last_update": bench["last_update"],
        "neodata_version_label": bench["label"],
        "benchmark_source": bench["source_name"],
        "benchmark_budget_code": bench["source_budget_code"],
        "benchmark_records": bench["conceptos_cargados"],
        "material_ai": get_material_ai_status(),
    }

def _professional_status_payload():
    payload = professional_mvp_status()
    payload.update({
        "app_version": app.version,
        "data_dir": str(DATA_DIR),
        "jobs_dir": str(JOBS_DIR),
        "max_providers": MAX_PROVIDERS,
        "market_indirect_pct": MARKET_INDIRECT_PCT,
        "reference_source": REFERENCE_SOURCE_LABEL,
        "excel_hide_internal_sheets": os.getenv("EXCEL_HIDE_INTERNAL_SHEETS", "1"),
    })
    return payload

@app.get("/api/professional/status")
def professional_status_endpoint():
    return _professional_status_payload()

# Alias de compatibilidad para clientes antiguos.
@app.get("/api/professional-mvp/status")
def professional_mvp_status_endpoint():
    return _professional_status_payload()

@app.get("/")
def root():
    index_file = FRONTEND_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    return JSONResponse(health())

@app.get("/app", response_class=HTMLResponse)
def app_entry():
    index_file = FRONTEND_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    # Fallback defensivo: nunca devolver 404 en /app por ausencia accidental del archivo en un patch.
    return HTMLResponse("""<!doctype html><html><head><meta charset=\"utf-8\"><title>Comparador</title></head><body><h1>Comparador</h1><p>frontend/index.html no fue encontrado en el despliegue actual.</p><p>Verifica que el patch incluya <code>frontend/index.html</code>.</p></body></html>""", status_code=200)

@app.get("/admin", response_class=HTMLResponse)
def admin_entry():
    admin_file = FRONTEND_DIR / "admin_benchmark.html"
    if admin_file.exists():
        return FileResponse(str(admin_file))
    return HTMLResponse("""<!doctype html><html><head><meta charset=\"utf-8\"><title>Admin</title></head><body><h1>Admin</h1><p>frontend/admin_benchmark.html no fue encontrado en el despliegue actual.</p></body></html>""", status_code=200)


# ── Conversión .xls → .xlsx ─────────────────────────────────
def xls_to_xlsx(xls_path: str, xlsx_path: str):
    import xlrd
    from openpyxl import Workbook
    xls_wb = xlrd.open_workbook(xls_path)
    xlsx_wb = Workbook()
    xlsx_wb.remove(xlsx_wb.active)
    for sheet in xls_wb.sheets():
        ws = xlsx_wb.create_sheet(sheet.name)
        for rx in range(sheet.nrows):
            for cx in range(sheet.ncols):
                val = sheet.cell_value(rx, cx)
                if sheet.cell_type(rx, cx) == 3:
                    val = xlrd.xldate_as_datetime(val, xls_wb.datemode)
                ws.cell(row=rx+1, column=cx+1, value=val)
    xlsx_wb.save(xlsx_path)


def save_and_convert(archivo: UploadFile, dest: str) -> str:
    ext = Path(archivo.filename).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        raise HTTPException(400, f"'{archivo.filename}' no es Excel (.xlsx o .xls)")
    ruta = dest + ext
    with open(ruta, "wb") as f:
        f.write(archivo.file.read())
    if ext == ".xls":
        xlsx_path = dest + ".xlsx"
        xls_to_xlsx(ruta, xlsx_path)
        return xlsx_path
    return ruta

@app.post("/preview")
async def preview(
    archivo1: UploadFile = File(...),
    archivo2: Optional[UploadFile] = File(default=None),
    archivo3: Optional[UploadFile] = File(default=None),
    archivo4: Optional[UploadFile] = File(default=None),
    archivo5: Optional[UploadFile] = File(default=None),
    resumen1: Optional[UploadFile] = File(default=None),
    resumen2: Optional[UploadFile] = File(default=None),
    resumen3: Optional[UploadFile] = File(default=None),
    resumen4: Optional[UploadFile] = File(default=None),
    resumen5: Optional[UploadFile] = File(default=None),
    nombre1: str = Form(default="Proveedor 1"),
    nombre2: str = Form(default=""),
    nombre3: str = Form(default=""),
    nombre4: str = Form(default=""),
    nombre5: str = Form(default=""),
    catalogo: Optional[UploadFile] = File(default=None),
    resumen_oficial: Optional[UploadFile] = File(default=None),
):
    archivos_raw = [archivo1, archivo2, archivo3, archivo4, archivo5]
    resumenes_raw = [resumen1, resumen2, resumen3, resumen4, resumen5]
    nombres_raw = [nombre1, nombre2, nombre3, nombre4, nombre5]
    pares = [(a, n, r) for a, n, r in zip(archivos_raw, nombres_raw, resumenes_raw) if a is not None and a.filename]
    if not pares:
        raise HTTPException(400, "Se necesita al menos 1 archivo de cotización.")
    _validate_provider_count(len(pares))
    if len(pares) == 1 and not (_has_upload(resumenes_raw[0]) or _has_upload(resumen_oficial)):
        raise HTTPException(400, "Falta el Resumen PU. Carga el resumen oficial común o el resumen del proveedor para usar cantidades reales.")

    with tempfile.TemporaryDirectory() as tmpdir:
        resumen_oficial_path = None
        if _has_upload(resumen_oficial):
            resumen_oficial_path = save_and_convert(resumen_oficial, str(Path(tmpdir) / "resumen_oficial"))
        paths = []
        nombres = []
        resumen_paths = []
        for i, (archivo, nombre, resumen) in enumerate(pares):
            path = save_and_convert(archivo, str(Path(tmpdir) / f"preview_{i}"))
            try:
                detect_valid_supplier_sheet(path)
            except Exception as e:
                raise HTTPException(400, str(e))
            paths.append(path)
            nombres.append(nombre or f"Proveedor {i+1}")
            resumen_path = None
            if _has_upload(resumen):
                resumen_path = save_and_convert(resumen, str(Path(tmpdir) / f"preview_resumen_{i}"))
            resumen_paths.append(_select_resumen_for_provider(resumen_path, resumen_oficial_path))

        catalogo_path = None
        if catalogo and catalogo.filename:
            ext_cat = Path(catalogo.filename).suffix.lower()
            dest_cat = str(Path(tmpdir) / f"catalogo{ext_cat}")
            with open(dest_cat, "wb") as f:
                f.write(catalogo.file.read())
            if ext_cat == ".xls":
                xlsx_cat = dest_cat.replace(".xls", ".xlsx")
                xls_to_xlsx(dest_cat, xlsx_cat)
                catalogo_path = xlsx_cat
            else:
                catalogo_path = dest_cat

        try:
            
            if os.getenv("PREVIEW_DEEP", "0").strip().lower() in {"1", "true", "yes", "si"}:
                return JSONResponse(build_preview_payload(paths, nombres, catalogo_path=catalogo_path, resumen_paths=resumen_paths))
            return JSONResponse(build_preview_payload_fast(paths, nombres, catalogo_path=catalogo_path, resumen_paths=resumen_paths))
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            raise HTTPException(500, f"No se pudo generar el preview: {e}")



# ── Endpoint principal: COMPARAR ────────────────────────────
@app.post("/comparar")
async def comparar(
    archivo1:  UploadFile        = File(...),
    archivo2:  Optional[UploadFile] = File(default=None),
    archivo3:  Optional[UploadFile] = File(default=None),
    archivo4:  Optional[UploadFile] = File(default=None),
    archivo5:  Optional[UploadFile] = File(default=None),
    resumen1:  Optional[UploadFile] = File(default=None),
    resumen2:  Optional[UploadFile] = File(default=None),
    resumen3:  Optional[UploadFile] = File(default=None),
    resumen4:  Optional[UploadFile] = File(default=None),
    resumen5:  Optional[UploadFile] = File(default=None),
    nombre1:   str = Form(default="Proveedor 1"),
    nombre2:   str = Form(default=""),
    nombre3:   str = Form(default=""),
    nombre4:   str = Form(default=""),
    nombre5:   str = Form(default=""),
    catalogo:  Optional[UploadFile] = File(default=None),
    resumen_oficial: Optional[UploadFile] = File(default=None),
    proyecto:  str = Form(default=""),
    cliente:   str = Form(default="Nestlé Servicios Industriales"),
    ubicacion: str = Form(default=""),
):
    # Recolectar archivos y nombres válidos
    archivos_raw = [archivo1, archivo2, archivo3, archivo4, archivo5]
    resumenes_raw = [resumen1, resumen2, resumen3, resumen4, resumen5]
    nombres_raw  = [nombre1, nombre2, nombre3, nombre4, nombre5]
    pares = [(a, n, r) for a, n, r in zip(archivos_raw, nombres_raw, resumenes_raw)
             if a is not None and a.filename]

    n_proveedores = len(pares)
    _validate_provider_count(n_proveedores)

    # Regla: mínimo 1 proveedor
    if n_proveedores < 1:
        raise HTTPException(400, "Se necesita al menos 1 archivo de cotización.")

    # Regla Paso 19: en modo 1 proveedor el Resumen PU es obligatorio.
    # Sin este archivo el Excel vuelve a cantidad=1 y el analisis queda invalido.
    if n_proveedores == 1 and not (_has_upload(resumenes_raw[0]) or _has_upload(resumen_oficial)):
        raise HTTPException(400, "Falta el Resumen PU. Carga el resumen oficial común o el resumen del proveedor para usar cantidades reales.")

    # Paso 21: la validacion antigua de Neodata ya no aplica.
    # Con la arquitectura nueva, un solo proveedor se compara contra los catalogos
    # granulares Construdata embebidos (materiales, mano de obra y maquinaria) y
    # contra el Resumen PU cargado para cantidades reales. Por lo tanto, NO se
    # debe exigir /admin/neodata ni un segundo proveedor.

    with tempfile.TemporaryDirectory() as tmpdir:
        resumen_oficial_path = None
        if _has_upload(resumen_oficial):
            resumen_oficial_path = save_and_convert(resumen_oficial, str(Path(tmpdir) / "resumen_oficial"))
        paths   = []
        nombres = []
        resumen_paths = []
        for i, (archivo, nombre, resumen) in enumerate(pares):
            dest_base = str(Path(tmpdir) / f"prov_{i}")
            try:
                path = save_and_convert(archivo, dest_base)
                try:
                    detect_valid_supplier_sheet(path)
                except Exception as e:
                    raise HTTPException(400, str(e))
                paths.append(path)
                nombres.append(nombre or f"Proveedor {i+1}")
                resumen_path = None
                if _has_upload(resumen):
                    resumen_path = save_and_convert(resumen, str(Path(tmpdir) / f"resumen_{i}"))
                resumen_paths.append(_select_resumen_for_provider(resumen_path, resumen_oficial_path))
            except HTTPException:
                raise
            except Exception as e:
                raise HTTPException(400, f"Error leyendo '{archivo.filename}': {e}")

        # Catálogo de licitación (orden oficial, opcional)
        catalogo_path = None
        if catalogo and catalogo.filename:
            ext_cat = Path(catalogo.filename).suffix.lower()
            dest_cat = str(Path(tmpdir) / f"catalogo{ext_cat}")
            with open(dest_cat, "wb") as f:
                f.write(catalogo.file.read())
            if ext_cat == ".xls":
                xlsx_cat = dest_cat.replace(".xls", ".xlsx")
                xls_to_xlsx(dest_cat, xlsx_cat)
                catalogo_path = xlsx_cat
            else:
                catalogo_path = dest_cat

        # Precios Neodata (siempre si existe)
        nacional_path = str(CONSTRUDATA_MATRICES_PATH) if CONSTRUDATA_MATRICES_PATH.exists() else (str(NEODATA_PATH) if NEODATA_PATH.exists() else None)

        try:
            output_path = Path(tmpdir) / "comparativo.xlsx"
            meta = {
                "proyecto":  proyecto or "Sin nombre",
                "cliente":   cliente,
                "ubicacion": ubicacion,
                "modo":      "analisis_cotizacion_unificado" if n_proveedores == 1 else "analisis_cotizacion_unificado",
                "resumen_oficial_comun": bool(resumen_oficial_path),
                "metodologia": _mvp_methodology_payload(resumen_oficial_path),
            }
            _create_unified_analysis_workbook_safe(
                output_path,
                paths,
                nombres,
                meta,
                catalogo_path=catalogo_path,
                nacional_path=nacional_path,
                resumen_paths=resumen_paths,
                resumen_oficial_path=resumen_oficial_path,
            )
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            linea = tb.splitlines()[-3] if tb and len(tb.splitlines()) >= 3 else ''
            raise HTTPException(500, f"Error: {e} | Linea: {linea}")

        nombre_salida = _safe_download_filename(f"Comparativo_{(proyecto or 'cotizacion')}", ".xlsx")
        final_path = f"/tmp/{nombre_salida}"
        shutil.copy(str(output_path), final_path)

        return FileResponse(
            path=final_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_salida,
        )



# ── v0.2.8.3: resumen estructurado del Excel para dashboard/chat ───────────
def _to_float_safe(v):
    try:
        if v is None or v == "":
            return None
        if isinstance(v, str):
            cleaned = v.replace("$", "").replace(",", "").replace("%", "").strip()
            if cleaned in {"", "—", "-"}:
                return None
            n = float(cleaned)
            if "%" in v:
                return n / 100.0
            return n
        if isinstance(v, bool):
            return None
        return float(v)
    except Exception:
        return None


def _norm_header_key(v):
    import unicodedata, re
    txt = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", txt.lower()).strip("_")


def _find_header_row(ws, required_any=None, max_rows=40):
    required_any = required_any or []
    for row in range(1, min(ws.max_row, max_rows) + 1):
        vals = [ws.cell(row, c).value for c in range(1, ws.max_column + 1)]
        keys = [_norm_header_key(v) for v in vals]
        joined = "|".join(keys)
        if any(req in joined for req in required_any):
            mapping = {k: i + 1 for i, k in enumerate(keys) if k}
            return row, mapping, vals
    return None, {}, []


def _get_by_header(ws, row, mapping, aliases):
    for alias in aliases:
        key = _norm_header_key(alias)
        if key in mapping:
            return ws.cell(row, mapping[key]).value
    # fallback contains match
    for k, col in mapping.items():
        if any(_norm_header_key(alias) in k for alias in aliases):
            return ws.cell(row, col).value
    return None


def _severity_from_diff(diff_pct, status=None, impact=None):
    st = str(status or "").lower()
    if "sin" in st or "incompat" in st:
        return "Alta"
    if diff_pct is not None and abs(diff_pct) >= 0.30:
        return "Alta"
    if diff_pct is not None and abs(diff_pct) >= 0.15:
        return "Media"
    if impact is not None and abs(impact) >= 10000:
        return "Media"
    return "Baja"


def _recommendation_from_row(status=None, diff_pct=None):
    st = str(status or "").lower()
    if "sin" in st:
        return "Solicitar desglose, soporte técnico y validación de alcance antes de usar la diferencia económica como criterio."
    if diff_pct is not None and diff_pct > 0.30:
        return "Solicitar apertura de matriz, jornales/rendimientos, costos de insumos y alcance incluido; priorizar negociación por impacto."
    if diff_pct is not None and diff_pct < -0.20:
        return "Validar que el alcance cotizado no esté incompleto frente a la especificación antes de considerarlo competitivo."
    return "Revisar en el Excel técnico si forma parte de las partidas principales o del paquete de negociación."


def _build_analysis_summary_from_excel(output_path: str, mode: str = None, provider_names=None):
    """Construye un resumen web real desde el Excel final.

    No recalcula el workbook: lee los valores ya escritos en Comparativa, Detalle y Analisis experto IA
    para alimentar dashboard y chat con evidencia estructurada.
    """
    import openpyxl, datetime, re
    provider_names = provider_names or []
    summary = {
        "version": "v0.2.8.3",
        "mode": mode or ("multi_provider" if len(provider_names) > 1 else "single_provider"),
        "generated_at": datetime.datetime.utcnow().isoformat() + "Z",
        "providers": [],
        "provider_scores": [],
        "rows": [],
        "critical_concepts": [],
        "top_concepts": [],
        "cost_breakdown": {},
        "labor_findings": [],
        "equipment_findings": [],
        "reference_gaps": [],
        "ai_findings": [],
        "totals": {},
        "data_quality": {"has_structured_summary": False, "warnings": []},
    }
    try:
        wb = openpyxl.load_workbook(output_path, data_only=True, read_only=True)
    except Exception as exc:
        summary["data_quality"]["warnings"].append(f"No se pudo leer el Excel para resumen web: {type(exc).__name__}: {exc}")
        return summary

    def sheet(name):
        return wb[name] if name in wb.sheetnames else None


    def _detail_provider_name(title, index):
        txt = str(title or "").strip()
        m = re.search(r"Detalle\s*-\s*P(\d+)", txt, flags=re.I)
        if m:
            pos = int(m.group(1)) - 1
            if provider_names and 0 <= pos < len(provider_names) and provider_names[pos]:
                return str(provider_names[pos])
            return f"Proveedor {pos + 1}"
        if txt.lower() == "detalle":
            return str(provider_names[0]) if provider_names else "Proveedor 1"
        return str(provider_names[index]) if provider_names and index < len(provider_names) else txt

    def _augment_summary_from_detail_sheets():
        """Alimenta dashboard/chat desde Detalle.

        Esta es la fuente de verdad del resultado técnico. No depende de Comparativa.
        """
        detail_names = [n for n in wb.sheetnames if str(n).strip().lower() == "detalle" or str(n).strip().lower().startswith("detalle - p")]
        detail_scores = []
        all_breakdown = summary.setdefault("cost_breakdown", {})
        labor = list(summary.get("labor_findings") or [])
        equipment = list(summary.get("equipment_findings") or [])
        gaps = list(summary.get("reference_gaps") or [])
        for idx, name in enumerate(detail_names):
            ws_det = wb[name]
            provider = _detail_provider_name(name, idx)
            hdr_row, mapping, _headers = _find_header_row(ws_det, required_any=["concepto_insumo", "costo_contratista", "importe_contratista", "match_construdata"], max_rows=60)
            if not hdr_row:
                continue
            totals_by_provider = {}
            total_c = 0.0
            total_m = 0.0
            rows_count = 0
            provider_critical = 0
            for r in range(hdr_row + 1, ws_det.max_row + 1):
                row_provider = _get_by_header(ws_det, r, mapping, ["Proveedor", "Contratista"])
                provider_row = str(row_provider).strip() if row_provider not in (None, "") and str(row_provider).strip().lower() != "proveedor" else provider
                code = _get_by_header(ws_det, r, mapping, ["Codigo", "Código"])
                desc = _get_by_header(ws_det, r, mapping, ["Concepto / Insumo", "Concepto", "Insumo"])
                unit = _get_by_header(ws_det, r, mapping, ["Unidad"])
                typ = str(_get_by_header(ws_det, r, mapping, ["Tipo"]) or "").strip()
                cost_c = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Costo Contratista"]))
                qty = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Cantidad"]))
                imp_c = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Importe Contratista"]))
                cost_m = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Costo Mercado", "Costo Referencia"]))
                imp_m = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Importe Mercado", "Importe Referencia"]))
                diff_pct = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Dif % Costo", "Dif %", "Diferencia %"]))
                base_pct = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Base Mercado %"]))
                match = _get_by_header(ws_det, r, mapping, ["Match Construdata", "Referencia Construdata"])
                conf = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Conf.", "Conf"]))
                status = str(_get_by_header(ws_det, r, mapping, ["Estado"]) or "").strip()
                note = _get_by_header(ws_det, r, mapping, ["Nota", "Comentario", "Recomendacion", "Recomendación"])
                if imp_c is None and cost_c is not None and qty is not None:
                    imp_c = cost_c * qty
                if imp_m is None and cost_m is not None and qty is not None:
                    imp_m = cost_m * qty
                if diff_pct is None and cost_c is not None and cost_m not in (None, 0):
                    diff_pct = (cost_c - cost_m) / cost_m
                # Saltar bandas/títulos/encabezados repetidos que no sean renglones técnicos.
                if str(code or "").strip().lower() in {"codigo", "código"} or str(desc or "").strip().lower() in {"concepto", "concepto / insumo", "insumo"}:
                    continue
                if not (code or desc) or not (typ or cost_c is not None or cost_m is not None or imp_c is not None or imp_m is not None or status):
                    continue
                rows_count += 1
                curp = totals_by_provider.setdefault(provider_row, {"total_c": 0.0, "total_m": 0.0, "rows": 0, "critical": 0})
                curp["rows"] += 1
                if imp_c is not None:
                    total_c += imp_c
                    curp["total_c"] += imp_c
                if imp_m is not None:
                    total_m += imp_m
                    curp["total_m"] += imp_m
                diff_amt = (imp_c - imp_m) if imp_c is not None and imp_m is not None else None
                risk = _severity_from_diff(diff_pct, status, diff_amt)
                if risk in {"Alta", "Media"}:
                    provider_critical += 1
                    try:
                        totals_by_provider.setdefault(provider_row, {"total_c": 0.0, "total_m": 0.0, "rows": 0, "critical": 0})["critical"] += 1
                    except Exception:
                        pass
                row = {
                    "clave": str(code or "").strip(),
                    "codigo": str(code or "").strip(),
                    "descripcion": str(desc or "").strip(),
                    "unidad": unit,
                    "cantidad": qty,
                    "tipo": typ,
                    "proveedor": provider_row,
                    "ganador": provider_row,
                    "pu_contratista": cost_c,
                    "importe_contratista": imp_c,
                    "pu_mercado": cost_m,
                    "importe_mercado": imp_m,
                    "diferencia_total": diff_amt,
                    "ahorro": abs(diff_amt) if diff_amt is not None else None,
                    "delta_pct_neodata": diff_pct,
                    "desviacion_total_pct": diff_pct,
                    "base_mercado_pct": base_pct,
                    "match_construdata": match,
                    "confianza": conf,
                    "status": status or ("sobre_referencia" if diff_pct is not None and diff_pct > 0.15 else "ok"),
                    "riesgo": risk,
                    "nota": note,
                    "recommendation": _recommendation_from_row(status, diff_pct),
                }
                summary["rows"].append(row)
                if typ:
                    key = typ.upper()
                    cur = all_breakdown.setdefault(key, {"contractor": 0.0, "market": 0.0, "difference": 0.0, "rows": 0})
                    if imp_c is not None:
                        cur["contractor"] += imp_c
                    if imp_m is not None:
                        cur["market"] += imp_m
                    cur["difference"] = cur["contractor"] - cur["market"]
                    cur["rows"] += 1
                text = f"{code or ''} {desc or ''} {typ}".lower()
                if "mano" in typ.lower() or any(w in text for w in ["supervisor", "soldador", "tubero", "ayudante", "oficial", "operador"]):
                    if diff_pct is not None and diff_pct > 0.10:
                        labor.append({"provider": provider_row, "role": str(desc or code or "Mano de Obra"), "code": str(code or ""), "contractor_cost": cost_c, "market_cost": cost_m, "difference_pct": diff_pct, "severity": risk, "match": match})
                if "equipo" in typ.lower() or any(w in text for w in ["montacargas", "grua", "grúa", "andamio", "manlift", "plataforma", "herramienta", "epp"]):
                    if (diff_pct is not None and abs(diff_pct) > 0.15) or any(w in text for w in ["montacargas", "grua", "grúa", "herramienta", "epp"]):
                        equipment.append({"provider": provider_row, "item": str(desc or code or "Equipo"), "code": str(code or ""), "contractor_cost": cost_c, "market_cost": cost_m, "difference_pct": diff_pct, "severity": risk, "match": match})
                if "sin" in status.lower() or "incompat" in status.lower() or (not match and conf in (None, 0)):
                    gaps.append({"provider": provider_row, "code": str(code or ""), "description": str(desc or ""), "reason": status or "Sin referencia Construdata estructurada"})
            if rows_count:
                for pname, vals in totals_by_provider.items():
                    tc = vals.get("total_c") or 0.0
                    tm = vals.get("total_m") or 0.0
                    diff = tc - tm if tm else None
                    detail_scores.append({
                        "nombre": pname,
                        "provider": pname,
                        "total_estimado": tc,
                        "total_oferta": tc,
                        "total_mercado": tm if tm else None,
                        "diferencia_total": diff,
                        "desviacion_total_pct": (diff / tm) if tm else None,
                        "conceptos": vals.get("rows") or 0,
                        "riesgos": vals.get("critical") or 0,
                        "dictamen": "Análisis calculado desde Detalle técnico del proveedor.",
                    })
        if detail_scores:
            summary["_detail_provider_scores"] = detail_scores
            summary["labor_findings"] = sorted(labor, key=lambda x: abs(x.get("difference_pct") or 0), reverse=True)[:20]
            summary["equipment_findings"] = sorted(equipment, key=lambda x: abs(x.get("difference_pct") or 0), reverse=True)[:20]
            summary["reference_gaps"] = gaps[:40]

    # 1) Comparativa individual o multi para conceptos base.
    ws = sheet("Comparativa")
    if ws:
        hdr_row, mapping, _ = _find_header_row(ws, required_any=["servicio", "codigo", "part", "concepto"], max_rows=15)
        if hdr_row:
            for r in range(hdr_row + 1, ws.max_row + 1):
                code = _get_by_header(ws, r, mapping, ["Servicio", "Codigo", "Código", "Part.", "Part"])
                desc = _get_by_header(ws, r, mapping, ["Descripcion", "Descripción", "Concepto", "Concepto base"])
                if not code and not desc:
                    continue
                unit = _get_by_header(ws, r, mapping, ["Unidad", "Uni.", "Uni"])
                qty = _to_float_safe(_get_by_header(ws, r, mapping, ["Cantidad", "Cant.", "Cant"] ))
                pu_cont = _to_float_safe(_get_by_header(ws, r, mapping, ["PU Contratista", "P.U. Contratista", "PU Proveedor", "P1 PU Contratista"] ))
                imp_cont = _to_float_safe(_get_by_header(ws, r, mapping, ["Importe Contratista", "Total Contratista", "P1 Importe", "Importe Total"] ))
                pu_ref = _to_float_safe(_get_by_header(ws, r, mapping, ["PU Mercado", "PU Referencia", "PU Base", "PU Base CD"] ))
                imp_ref = _to_float_safe(_get_by_header(ws, r, mapping, ["Importe Mercado", "Importe Referencia", "Importe Base", "Importe Base CD"] ))
                diff_pct = _to_float_safe(_get_by_header(ws, r, mapping, ["Dif %", "Diferencia %", "P1 Dif %", "Delta %"] ))
                status = _get_by_header(ws, r, mapping, ["Estado", "P1 Estado", "Estado comparativo"])
                note = _get_by_header(ws, r, mapping, ["Nota", "Comentario", "Recomendacion", "Recomendación"])
                if imp_cont is None and pu_cont is not None and qty is not None:
                    imp_cont = pu_cont * qty
                if imp_ref is None and pu_ref is not None and qty is not None:
                    imp_ref = pu_ref * qty
                diff_amt = (imp_cont - imp_ref) if imp_cont is not None and imp_ref is not None else None
                if diff_pct is None and pu_cont not in (None, 0) and pu_ref not in (None, 0):
                    diff_pct = (pu_cont - pu_ref) / pu_ref
                riesgo = _severity_from_diff(diff_pct, status, diff_amt)
                row = {
                    "clave": str(code or "").strip(),
                    "codigo": str(code or "").strip(),
                    "descripcion": str(desc or "").strip(),
                    "unidad": unit,
                    "cantidad": qty,
                    "pu_contratista": pu_cont,
                    "importe_contratista": imp_cont,
                    "pu_mercado": pu_ref,
                    "importe_mercado": imp_ref,
                    "diferencia_total": diff_amt,
                    "ahorro": abs(diff_amt) if diff_amt is not None else None,
                    "delta_pct_neodata": diff_pct,
                    "desviacion_total_pct": diff_pct,
                    "status": status or ("sobre_referencia" if diff_pct is not None and diff_pct > 0.15 else "ok"),
                    "riesgo": riesgo,
                    "nota": note,
                    "recommendation": _recommendation_from_row(status, diff_pct),
                    "ganador": provider_names[0] if provider_names else "Proveedor 1",
                }
                # Evita que una fila TOTAL se mezcle como concepto crítico, pero conserva para totales.
                if str(code or "").strip().upper() == "TOTAL":
                    summary["totals"]["contractor_total"] = imp_cont
                    summary["totals"]["market_total"] = imp_ref
                    summary["totals"]["difference_amount"] = diff_amt
                    summary["totals"]["difference_pct"] = diff_pct
                else:
                    summary["rows"].append(row)

    # 2) Fuente principal para web/chat: Detalle.
    _augment_summary_from_detail_sheets()

    # 2) Totales fallback desde filas.
    if not summary["totals"].get("contractor_total"):
        ctotal = sum((r.get("importe_contratista") or 0) for r in summary["rows"] if r.get("importe_contratista") is not None)
        mtotal = sum((r.get("importe_mercado") or 0) for r in summary["rows"] if r.get("importe_mercado") is not None)
        if ctotal or mtotal:
            summary["totals"].update({
                "contractor_total": ctotal,
                "market_total": mtotal,
                "difference_amount": ctotal - mtotal,
                "difference_pct": ((ctotal - mtotal) / mtotal) if mtotal else None,
            })

    # 3) Provider scores minimo para dashboard.
    detail_scores = summary.pop("_detail_provider_scores", []) or []
    if detail_scores:
        summary["provider_scores"] = sorted(detail_scores, key=lambda x: x.get("total_estimado") or 0)
    elif summary["mode"] == "single_provider":
        pname = provider_names[0] if provider_names else "Proveedor 1"
        summary["provider_scores"] = [{
            "nombre": pname,
            "provider": pname,
            "total_estimado": summary["totals"].get("contractor_total"),
            "total_oferta": summary["totals"].get("contractor_total"),
            "total_mercado": summary["totals"].get("market_total"),
            "diferencia_total": summary["totals"].get("difference_amount"),
            "desviacion_total_pct": summary["totals"].get("difference_pct"),
            "conceptos": len(summary["rows"]),
            "dictamen": "Análisis individual contra referencia de mercado/Construdata.",
        }]
    else:
        # Lee totales por proveedor desde la Comparativa compacta del motor unificado.
        ws_multi = sheet("Comparativa")
        providers = []
        if ws_multi:
            hdr_row, mapping, headers = _find_header_row(ws_multi, required_any=["proveedor", "p1", "p2"], max_rows=15)
            if hdr_row:
                col_info = []
                for idx, h in enumerate(headers, start=1):
                    txt = str(h or "")
                    if ("Importe" in txt or "Total" in txt) and ("P" in txt or "Proveedor" in txt):
                        col_info.append((idx, txt))
                for idx, label in col_info[:8]:
                    total = 0.0
                    count = 0
                    for r in range(hdr_row + 1, ws_multi.max_row + 1):
                        v = _to_float_safe(ws_multi.cell(r, idx).value)
                        if v is not None:
                            total += v; count += 1
                    if count:
                        name = label.replace("Importe", "").replace("Total", "").strip() or f"Proveedor {len(providers)+1}"
                        providers.append({"nombre": name, "provider": name, "total_estimado": total, "total_oferta": total, "conceptos": count, "dictamen": "Proveedor incluido en comparativa multi."})
        summary["provider_scores"] = sorted(providers, key=lambda x: x.get("total_estimado") or 0) or [
            {"nombre": n, "provider": n, "conceptos": len(summary["rows"]), "dictamen": "Proveedor incluido en análisis multi."}
            for n in provider_names
        ]

    summary["providers"] = summary["provider_scores"]

    # 4) Conceptos críticos.
    candidates = []
    for r in summary["rows"]:
        impact = r.get("diferencia_total")
        pct = r.get("desviacion_total_pct")
        risk = r.get("riesgo")
        if (impact is not None and abs(impact) > 0) or risk in {"Alta", "Media"} or "sin" in str(r.get("status", "")).lower():
            candidates.append({
                "code": r.get("codigo"),
                "description": r.get("descripcion"),
                "desc": r.get("descripcion"),
                "provider": r.get("ganador") or (provider_names[0] if provider_names else "Proveedor 1"),
                "impact": impact,
                "risk": risk or r.get("status"),
                "action": r.get("recommendation"),
                "difference_pct": pct,
            })
    candidates.sort(key=lambda x: abs(x.get("impact") or 0), reverse=True)
    summary["critical_concepts"] = candidates[:12]
    summary["top_concepts"] = candidates[:8]

    # 5) Detalle: composición de costo, MO, equipo, huecos de referencia.
    ws_det = sheet("Detalle")
    if ws_det:
        hdr_row, mapping, _ = _find_header_row(ws_det, required_any=["concepto_insumo", "costo_contratista", "match_construdata"], max_rows=10)
        breakdown = {}
        labor = []
        equipment = []
        gaps = []
        if hdr_row:
            for r in range(hdr_row + 1, ws_det.max_row + 1):
                code = _get_by_header(ws_det, r, mapping, ["Codigo", "Código"])
                desc = _get_by_header(ws_det, r, mapping, ["Concepto / Insumo", "Concepto", "Insumo"])
                typ = str(_get_by_header(ws_det, r, mapping, ["Tipo"]) or "").strip()
                cost_c = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Costo Contratista"]))
                imp_c = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Importe Contratista"]))
                cost_m = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Costo Mercado", "Costo Referencia"]))
                imp_m = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Importe Mercado", "Importe Referencia"]))
                status = str(_get_by_header(ws_det, r, mapping, ["Estado"]) or "")
                match = _get_by_header(ws_det, r, mapping, ["Match Construdata", "Referencia Construdata"])
                conf = _to_float_safe(_get_by_header(ws_det, r, mapping, ["Conf.", "Conf"] ))
                if typ:
                    key = typ.upper()
                    cur = breakdown.setdefault(key, {"contractor": 0.0, "market": 0.0, "difference": 0.0, "rows": 0})
                    if imp_c is not None: cur["contractor"] += imp_c
                    if imp_m is not None: cur["market"] += imp_m
                    cur["difference"] = cur["contractor"] - cur["market"]
                    cur["rows"] += 1
                text = f"{code or ''} {desc or ''} {typ}".lower()
                diff_pct = ((cost_c - cost_m) / cost_m) if cost_c is not None and cost_m not in (None, 0) else None
                if "mano" in typ.lower() or any(w in text for w in ["supervisor", "soldador", "tubero", "ayudante", "oficial", "operador"]):
                    if diff_pct is not None and diff_pct > 0.10:
                        labor.append({
                            "role": str(desc or code or "Mano de Obra"),
                            "code": str(code or ""),
                            "contractor_cost": cost_c,
                            "market_cost": cost_m,
                            "difference_pct": diff_pct,
                            "severity": _severity_from_diff(diff_pct),
                            "match": match,
                        })
                if "equipo" in typ.lower() or any(w in text for w in ["montacargas", "grua", "grúa", "andamio", "manlift", "plataforma"]):
                    if (diff_pct is not None and diff_pct > 0.20) or any(w in text for w in ["montacargas", "grua", "grúa"]):
                        equipment.append({
                            "item": str(desc or code or "Equipo"),
                            "code": str(code or ""),
                            "contractor_cost": cost_c,
                            "market_cost": cost_m,
                            "difference_pct": diff_pct,
                            "severity": _severity_from_diff(diff_pct),
                            "match": match,
                        })
                if "sin_match" in status.lower() or (match in (None, "") and conf in (None, 0)):
                    if code or desc:
                        gaps.append({"code": str(code or ""), "description": str(desc or ""), "reason": status or "Sin referencia Construdata estructurada"})
        summary["cost_breakdown"] = breakdown
        summary["labor_findings"] = sorted(labor, key=lambda x: abs(x.get("difference_pct") or 0), reverse=True)[:12]
        summary["equipment_findings"] = sorted(equipment, key=lambda x: abs(x.get("difference_pct") or 0), reverse=True)[:12]
        summary["reference_gaps"] = gaps[:20]

    # 6) Texto del análisis experto IA para el chat.
    ws_ai = sheet("Analisis experto IA")
    if ws_ai:
        texts = []
        for row in ws_ai.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and len(v.strip()) > 40:
                    texts.append(v.strip())
        summary["ai_findings"] = [{"text": t[:4000]} for t in texts[:4]]

    summary["data_quality"]["has_structured_summary"] = bool(summary["rows"] or summary["provider_scores"] or summary["cost_breakdown"])
    if not summary["data_quality"]["has_structured_summary"]:
        summary["data_quality"]["warnings"].append("No se encontraron hojas Comparativa/Detalle con estructura esperada para resumen web.")
    try:
        wb.close()
    except Exception:
        pass
    return summary


def _analysis_summary_path(job_id: str) -> Path:
    return _job_dir(job_id) / "analysis_summary.json"


def _write_analysis_summary(job_id: str, summary: dict):
    try:
        p = _analysis_summary_path(job_id)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(summary or {}, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    except Exception as exc:
        print(f"[analysis_summary_write_error] {job_id}: {exc}")


def _read_analysis_summary(job_id: str):
    try:
        p = _analysis_summary_path(job_id)
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"[analysis_summary_read_error] {job_id}: {exc}")
    return None

def _run_comparison_job(job_id: str, paths, nombres, output_path: str, meta: dict, catalogo_path=None, nacional_path=None, resumen_paths=None, resumen_oficial_path=None):
    started = time.perf_counter()
    old_trace = os.environ.get("QUANTIA_TRACE_FILE")
    os.environ["QUANTIA_TRACE_FILE"] = str(_job_trace_path(job_id))
    os.environ.setdefault("QUANTIA_ENABLE_AI", "0")
    os.environ.setdefault("ENABLE_AI_MATERIAL_MATCHING", "0")
    stop_heartbeat = threading.Event()

    def heartbeat():
        while not stop_heartbeat.wait(15):
            elapsed = round(time.perf_counter() - started, 1)
            _job_trace(job_id, "heartbeat", elapsed_seconds=elapsed, phase=(_job_get(job_id) or {}).get("phase"))
            # Mantiene vivo el estado sin fingir avance real.
            _job_set(job_id, elapsed_seconds=elapsed)

    hb = threading.Thread(target=heartbeat, daemon=True)
    hb.start()
    try:
        _job_set(job_id, status="running", started_at=datetime.datetime.utcnow().isoformat() + "Z")
        _job_stage(job_id, "building_excel", 20, ai_enabled=os.getenv("QUANTIA_ENABLE_AI", "0"), files=len(paths), resumen_files=len([p for p in (resumen_paths or []) if p]), reference_source=REFERENCE_SOURCE_LABEL)
        _job_trace(job_id, "unified_analysis_call", output_path=output_path, catalogo_path=bool(catalogo_path), nacional_path=bool(nacional_path))
        _job_stage(job_id, "building_unified_analysis_excel", 55)
        _job_trace(job_id, "unified_analysis_workbook_call", output_path=output_path, catalogo_path=bool(catalogo_path), nacional_path=bool(nacional_path), providers=len(paths or []))
        _create_unified_analysis_workbook_safe(
            output_path,
            paths,
            nombres,
            meta,
            catalogo_path=catalogo_path,
            nacional_path=nacional_path,
            resumen_paths=resumen_paths or [],
            resumen_oficial_path=resumen_oficial_path,
        )
        _job_stage(job_id, "validating_output", 92)
        if not Path(output_path).exists() or Path(output_path).stat().st_size <= 0:
            raise RuntimeError("El Excel no se genero o quedo vacio.")
        size = Path(output_path).stat().st_size
        elapsed = round(time.perf_counter() - started, 2)
        analysis_mode = "multi_provider" if len(paths or []) > 1 else "single_provider"
        try:
            _job_stage(job_id, "building_analysis_summary", 96)
            summary = _build_analysis_summary_from_excel(output_path, mode=analysis_mode, provider_names=nombres or [])
            _write_analysis_summary(job_id, summary)
            summary_public = {
                "mode": summary.get("mode"),
                "rows_count": len(summary.get("rows") or []),
                "critical_count": len(summary.get("critical_concepts") or []),
                "provider_count": len(summary.get("provider_scores") or []),
                "has_structured_summary": (summary.get("data_quality") or {}).get("has_structured_summary"),
            }
        except Exception as exc:
            summary = None
            summary_public = {"error": f"{type(exc).__name__}: {exc}", "has_structured_summary": False}
            _job_trace(job_id, "analysis_summary_error", error=summary_public["error"])
        _job_trace(job_id, "done", elapsed_seconds=elapsed, size_bytes=size, analysis_summary=summary_public)
        _job_set(
            job_id,
            status="done",
            phase="ready",
            progress=100,
            output_path=output_path,
            size_bytes=size,
            elapsed_seconds=elapsed,
            trace_url=f"/jobs/{job_id}/trace",
            analysis_id=job_id,
            analysis_summary=summary,
            analysis_summary_public=summary_public,
        )
    except Exception as e:
        tb = traceback.format_exc()
        print(tb)
        _job_trace(job_id, "failed", error=str(e), traceback=tb[-4000:], elapsed_seconds=round(time.perf_counter() - started, 2))
        _job_set(job_id, status="error", phase="failed", progress=100, error=str(e), traceback=tb[-4000:], trace_url=f"/jobs/{job_id}/trace")
    finally:
        stop_heartbeat.set()
        if old_trace is not None:
            os.environ["QUANTIA_TRACE_FILE"] = old_trace
        else:
            os.environ.pop("QUANTIA_TRACE_FILE", None)


@app.post("/comparar_async")
async def comparar_async(
    archivo1:  UploadFile        = File(...),
    archivo2:  Optional[UploadFile] = File(default=None),
    archivo3:  Optional[UploadFile] = File(default=None),
    archivo4:  Optional[UploadFile] = File(default=None),
    archivo5:  Optional[UploadFile] = File(default=None),
    resumen1:  Optional[UploadFile] = File(default=None),
    resumen2:  Optional[UploadFile] = File(default=None),
    resumen3:  Optional[UploadFile] = File(default=None),
    resumen4:  Optional[UploadFile] = File(default=None),
    resumen5:  Optional[UploadFile] = File(default=None),
    nombre1:   str = Form(default="Proveedor 1"),
    nombre2:   str = Form(default=""),
    nombre3:   str = Form(default=""),
    nombre4:   str = Form(default=""),
    nombre5:   str = Form(default=""),
    catalogo:  Optional[UploadFile] = File(default=None),
    resumen_oficial: Optional[UploadFile] = File(default=None),
    proyecto:  str = Form(default=""),
    cliente:   str = Form(default="Nestle Servicios Industriales"),
    ubicacion: str = Form(default=""),
):
    _cleanup_old_jobs()
    archivos_raw = [archivo1, archivo2, archivo3, archivo4, archivo5]
    resumenes_raw = [resumen1, resumen2, resumen3, resumen4, resumen5]
    nombres_raw  = [nombre1, nombre2, nombre3, nombre4, nombre5]
    pares = [(a, n, r) for a, n, r in zip(archivos_raw, nombres_raw, resumenes_raw) if a is not None and a.filename]
    if not pares:
        raise HTTPException(400, "Se necesita al menos 1 archivo de cotizacion.")
    _validate_provider_count(len(pares))
    if len(pares) == 1 and not (_has_upload(resumenes_raw[0]) or _has_upload(resumen_oficial)):
        raise HTTPException(400, "Falta el Resumen PU. Carga el resumen oficial común o el resumen del proveedor para usar cantidades reales.")

    job_id = uuid.uuid4().hex
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    paths, nombres, resumen_paths = [], [], []
    try:
        resumen_oficial_path = None
        if _has_upload(resumen_oficial):
            resumen_oficial_path = save_and_convert(resumen_oficial, str(job_dir / "resumen_oficial"))
        for i, (archivo, nombre, resumen) in enumerate(pares):
            path = save_and_convert(archivo, str(job_dir / f"prov_{i}"))
            try:
                detect_valid_supplier_sheet(path)
            except Exception as e:
                raise HTTPException(400, str(e))
            paths.append(path)
            nombres.append(nombre or f"Proveedor {i+1}")
            resumen_path = None
            if _has_upload(resumen):
                resumen_path = save_and_convert(resumen, str(job_dir / f"resumen_{i}"))
            resumen_paths.append(_select_resumen_for_provider(resumen_path, resumen_oficial_path))

        catalogo_path = None
        if catalogo and catalogo.filename:
            ext_cat = Path(catalogo.filename).suffix.lower()
            if ext_cat not in [".xlsx", ".xls"]:
                raise HTTPException(400, "El catalogo debe ser .xlsx o .xls")
            raw_cat = job_dir / f"catalogo{ext_cat}"
            with open(raw_cat, "wb") as f:
                f.write(await catalogo.read())
            if ext_cat == ".xls":
                xlsx_cat = str(job_dir / "catalogo.xlsx")
                xls_to_xlsx(str(raw_cat), xlsx_cat)
                catalogo_path = xlsx_cat
            else:
                catalogo_path = str(raw_cat)

        nacional_path = str(CONSTRUDATA_MATRICES_PATH) if CONSTRUDATA_MATRICES_PATH.exists() else (str(NEODATA_PATH) if NEODATA_PATH.exists() else None)
        output_path = str(job_dir / "comparativo.xlsx")
        meta = {
            "proyecto": proyecto or "Sin nombre",
            "cliente": cliente,
            "ubicacion": ubicacion,
            "modo": "1_proveedor_vs_nacional" if len(pares) == 1 else "multi_proveedor",
            "resumen_oficial_comun": bool(resumen_oficial_path),
            "metodologia": _mvp_methodology_payload(resumen_oficial_path),
        }
        _job_set(job_id, status="queued", phase="saved_uploads", progress=5, created_at=datetime.datetime.utcnow().isoformat(), filename=_safe_download_filename(f"Comparativo_{(proyecto or 'cotizacion')}", ".xlsx"), metodologia=_mvp_methodology_payload(resumen_oficial_path))
        thread = threading.Thread(target=_run_comparison_job, args=(job_id, paths, nombres, output_path, meta, catalogo_path, nacional_path, resumen_paths, resumen_oficial_path), daemon=True)
        thread.start()
        return JSONResponse({"job_id": job_id, "status": "queued", "poll_url": f"/jobs/{job_id}", "download_url": f"/jobs/{job_id}/download"}, status_code=202)
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(400, f"No se pudo iniciar el job: {e}")


@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    job = _job_get(job_id)
    if not job:
        raise HTTPException(404, f"Job no encontrado en {JOBS_DIR}. Puede haberse creado antes de activar persistencia, haberse borrado por restart sin Volume, o estar consultandose en otra instancia.")
    public = {k: v for k, v in job.items() if k not in {"traceback", "output_path"}}
    trace = _read_job_trace(job_id, limit=8)
    public["trace_tail"] = trace[-8:]
    public["trace_url"] = f"/jobs/{job_id}/trace"
    if job.get("status") == "error":
        public["error"] = job.get("error")
    return JSONResponse(public)


@app.get("/jobs/{job_id}/trace")
def get_job_trace(job_id: str):
    job = _job_get(job_id)
    if not job:
        raise HTTPException(404, f"Job no encontrado en {JOBS_DIR}. Revisa /debug/jobs y confirma que JOBS_DIR/DATA_DIR apunten a un Railway Volume.")
    return JSONResponse({"job_id": job_id, "job": {k: v for k, v in job.items() if k not in {"output_path"}}, "trace": _read_job_trace(job_id, limit=1000)})


@app.get("/jobs/{job_id}/download")
def download_job(job_id: str):
    job = _job_get(job_id)
    if not job:
        raise HTTPException(404, f"Job no encontrado en {JOBS_DIR}. Revisa /debug/jobs.")
    if job.get("status") != "done":
        raise HTTPException(409, f"El job aun no esta listo. Estado: {job.get('status')}")
    output_path = job.get("output_path")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(404, "El archivo del job no existe.")
    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=job.get("filename") or "Comparativo_cotizacion.xlsx",
    )


# ── Admin: subir/actualizar precios Neodata ─────────────────
@app.post("/benchmark/upload")
@app.post("/admin/benchmark/upload")
@app.post("/admin/neodata")
async def subir_benchmark(
    archivo: UploadFile = File(...),
    source_name: str = Form(default="ConstruBase"),
    source_budget_code: str = Form(default="CONCURSOS"),
    notes: str = Form(default=""),
):
    """
    Sube o reemplaza el snapshot benchmark de mercado.
    Formato esperado del Excel:
      - Codigo / Código
      - Descripcion / Descripción
      - Unidad
      - PrecioUnitarioReferencia (o cualquier encabezado con precio unitario / referencia / pu)
    """
    ext = Path(archivo.filename).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        raise HTTPException(400, "El archivo debe ser .xlsx o .xls")

    tmp = str(NEODATA_PATH) + ".tmp" + ext
    with open(tmp, "wb") as f:
        f.write(await archivo.read())

    if ext == ".xls":
        xls_to_xlsx(tmp, str(NEODATA_PATH))
        os.remove(tmp)
    else:
        shutil.move(tmp, str(NEODATA_PATH))

    conceptos = 0
    materiales = 0
    try:
        precios = extract_precios_nacional(str(NEODATA_PATH))
        conceptos = len([k for k in precios.keys() if k != "__materials_index__"])
        materiales_index = precios.get("__materials_index__", []) if isinstance(precios, dict) else []
        materiales = len(materiales_index)
        vector_index = None
        if os.getenv("GENERATE_VECTOR_INDEX_ON_UPLOAD", "1").strip() not in {"0", "false", "False", "no"}:
            vector_index = ensure_construdata_vector_index(materiales_index, force=True)
    except Exception as e:
        raise HTTPException(400, f"No se pudo leer el benchmark cargado: {e}")

    loaded_at = datetime.datetime.utcnow().isoformat()
    _write_benchmark_meta({
        "source_name": source_name,
        "source_budget_code": source_budget_code,
        "notes": notes,
        "conceptos_cargados": conceptos,
        "materiales_cargados": materiales,
        "loaded_at": loaded_at,
        "filename": archivo.filename,
        "vector_index": vector_index,
    })

    return JSONResponse({
        "status": "ok",
        "mensaje": f"Benchmark actualizado. {conceptos} conceptos y {materiales} materiales cargados.",
        "archivo": archivo.filename,
        "conceptos_cargados": conceptos,
        "materiales_cargados": materiales,
        "source_name": source_name,
        "source_budget_code": source_budget_code,
        "loaded_at": loaded_at,
        "vector_index": vector_index,
    })



# ── Admin: verificar estado de Neodata ──────────────────────
@app.get("/benchmark/status")
@app.get("/admin/benchmark")
@app.get("/admin/neodata")
def estado_benchmark():
    bench = get_benchmark_status_payload()
    return JSONResponse({
        "cargado": bench["loaded"],
        "mensaje": bench["message"],
        "source_name": bench["source_name"],
        "source_budget_code": bench["source_budget_code"],
        "conceptos_cargados": bench["conceptos_cargados"],
        "materiales_cargados": bench.get("materiales_cargados", 0),
        "loaded_at": bench["last_update"],
        "label": bench["label"],
        "vector_index": bench.get("vector_index"),
        "material_cache": bench.get("material_cache"),
        "material_ai": get_material_ai_status(),
    })



@app.get("/admin/ai/status")
@app.get("/ai/status")
def estado_material_ai():
    bench = get_benchmark_status_payload()
    ai = get_material_ai_status()
    payload = dict(ai)
    ready = bool(ai.get("active") and bench.get("loaded") and bench.get("materiales_cargados", 0) > 0)
    payload.update({
        "benchmark_loaded": bench.get("loaded", False),
        "benchmark_materials": bench.get("materiales_cargados", 0),
        "benchmark_label": bench.get("label"),
        "ready": ready,
        "ready_reason": "IA y catalogo de materiales cargados." if ready else "Falta configurar IA o cargar Construdata con materiales.",
    })
    return JSONResponse(payload)

@app.post("/admin/ai/test")
def probar_material_ai():
    return JSONResponse(run_material_ai_smoke_test())

@app.get("/admin/ai/network")
def diagnostico_red_material_ai():
    return JSONResponse(run_network_diagnostics())

@app.get("/admin/ai/usage")
def estado_material_ai_usage():
    return JSONResponse(get_ai_usage())

@app.post("/admin/ai/usage/reset")
def reiniciar_material_ai_usage():
    return JSONResponse(reset_ai_usage())



def _safe_vector_index_filename(filename: str) -> str:
    raw = Path(filename or "").name.strip()
    raw = "".join(ch for ch in raw if ch.isalnum() or ch in ("-", "_", "."))
    model = (os.getenv("VOYAGE_MODEL") or "voyage-4-lite").replace("/", "_")
    if not raw.lower().endswith(".json"):
        raw = f"construdata_uploaded_{model}.json"
    if not (raw.startswith("construdata_") and raw.endswith(f"_{model}.json")):
        raw = f"construdata_uploaded_{model}.json"
    return raw

@app.post("/admin/vector-index/upload")
async def subir_vector_index_voyage(archivo: UploadFile = File(...)):
    filename = _safe_vector_index_filename(archivo.filename or "")
    target_dir = Path(get_vector_index_dir())
    target_dir.mkdir(parents=True, exist_ok=True)
    final_path = target_dir / filename
    tmp_path = target_dir / f".{filename}.uploading"
    bytes_written = 0
    chunk_size = int(os.getenv("VECTOR_INDEX_UPLOAD_CHUNK_MB", "8")) * 1024 * 1024
    try:
        with open(tmp_path, "wb") as out:
            while True:
                chunk = await archivo.read(chunk_size)
                if not chunk:
                    break
                out.write(chunk)
                bytes_written += len(chunk)
        if bytes_written <= 0:
            tmp_path.unlink(missing_ok=True)
            raise HTTPException(400, "El archivo de índice está vacío.")
        tmp_path.replace(final_path)
    except HTTPException:
        raise
    except Exception as exc:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise HTTPException(500, f"No se pudo guardar el índice vectorial: {type(exc).__name__}: {exc}")
    return JSONResponse({
        "ok": True,
        "filename": filename,
        "path": str(final_path),
        "size_mb": round(bytes_written / (1024 * 1024), 2),
        "message": "Índice Voyage cargado. No fue necesario reconstruirlo en producción.",
    })

@app.get("/admin/vector-index/status")
def estado_vector_index():
    status = get_vector_index_status(load_materials_index())
    status["job"] = _get_vector_index_job()
    return JSONResponse(status)

@app.post("/admin/vector-index/rebuild")
def reconstruir_vector_index():
    materials_index = load_materials_index()
    if not materials_index:
        raise HTTPException(400, "No hay Construdata/materiales cargados para vectorizar.")
    current = _get_vector_index_job()
    if current.get("running"):
        return JSONResponse({"ok": True, "accepted": True, "already_running": True, "job": current})
    job_id = str(uuid.uuid4())[:12]
    _set_vector_index_job(running=True, job_id=job_id, phase="queued", started_at=datetime.datetime.utcnow().isoformat() + "Z", total_records=len(materials_index), records_done=0, total_batches=0, batch_no=0, built_records=0, skipped_records=0, progress_pct=0, error=None, result=None)

    def progress(evt):
        updates = {"phase": evt.get("phase") or "running"}
        for key in ("total_records", "records_done", "total_batches", "batch_no", "built_records", "skipped_records", "retry_stats", "skipped_preview"):
            if key in evt:
                updates[key] = evt.get(key)
        
        try:
            total = int(updates.get("total_records") or _get_vector_index_job().get("total_records") or 0)
            done = int(updates.get("records_done") or _get_vector_index_job().get("records_done") or 0)
            if total:
                updates["progress_pct"] = round((done * 100.0) / total, 2)
        except Exception:
            pass
        _set_vector_index_job(**updates)

    def worker():
        try:
            _set_vector_index_job(phase="running")
            result = ensure_construdata_vector_index(materials_index, force=True, progress_callback=progress)
            _set_vector_index_job(running=False, phase="finished" if result.get("ok") else "failed", result=result, records_done=len(materials_index))
        except Exception as exc:
            _set_vector_index_job(running=False, phase="failed", error=f"{type(exc).__name__}: {exc}")

    threading.Thread(target=worker, daemon=True).start()
    return JSONResponse({"ok": True, "accepted": True, "job": _get_vector_index_job()})
@app.get("/admin/material-cache/status")
def estado_material_cache():
    return JSONResponse(cache_status())

@app.post("/admin/material-cache/clear")
def limpiar_material_cache():
    return JSONResponse(clear_material_match_cache())

# -- Admin: jornada de mano de obra persistente -----------------------------
@app.get("/admin/labor/status")
def estado_labor_tabulador():
    loaded = TABULADOR_MO_PATH.exists()
    payload = {"cargado": loaded, "archivo": TABULADOR_MO_PATH.name}
    if loaded:
        stat = TABULADOR_MO_PATH.stat()
        payload.update({"tamano_kb": round(stat.st_size / 1024, 1), "updated_at": datetime.datetime.fromtimestamp(stat.st_mtime).isoformat()})
        try:
            rows = load_labor_tabulador(str(TABULADOR_MO_PATH), "2026")
            payload["registros_2026"] = len(rows) if hasattr(rows, "__len__") else 0
        except Exception as e:
            payload["warning"] = str(e)
    return JSONResponse(payload)

@app.post("/admin/labor/upload")
async def subir_labor_tabulador(archivo: UploadFile = File(...)):
    ext = Path(archivo.filename).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        raise HTTPException(400, "El archivo debe ser .xlsx o .xls")
    tmp = str(TABULADOR_MO_PATH) + ".tmp" + ext
    with open(tmp, "wb") as f:
        f.write(await archivo.read())
    if ext == ".xls":
        xls_to_xlsx(tmp, str(TABULADOR_MO_PATH))
        os.remove(tmp)
    else:
        shutil.move(tmp, str(TABULADOR_MO_PATH))
    try:
        rows = load_labor_tabulador(str(TABULADOR_MO_PATH), "2026")
        n = len(rows) if hasattr(rows, "__len__") else 0
    except Exception as e:
        raise HTTPException(400, f"Se guardo el archivo, pero no se pudo leer como tabulador 2026: {e}")
    return JSONResponse({"status": "ok", "mensaje": f"Tabulador de mano de obra 2026 actualizado. {n} registros detectados.", "archivo": archivo.filename, "registros_2026": n})

# ── Archivos estáticos ──────────────────────────────────────
_FRONTEND_CANDIDATES = [
    Path(__file__).resolve().parent / "frontend",
    Path(__file__).resolve().parent.parent / "frontend",
    Path.cwd() / "frontend",
]
FRONTEND_DIR = next((x for x in _FRONTEND_CANDIDATES if x.exists()), _FRONTEND_CANDIDATES[0])
if FRONTEND_DIR.exists():
    app.mount("/frontend", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)


@app.get("/debug/performance")
def debug_performance():
    # STEP35: this endpoint must never 500; it is used to diagnose production.
    try:
        active_jobs = len(_COMPARISON_JOBS) if '_COMPARISON_JOBS' in globals() else 0
        mem_jobs = {jid: {k: v for k, v in job.items() if k not in {"output_path", "traceback"}} for jid, job in list(globals().get('_COMPARISON_JOBS', {}).items())[-10:]}
    except Exception as exc:
        active_jobs = None
        mem_jobs = {"error": f"{type(exc).__name__}: {exc}"}
    disk_jobs = []
    try:
        for d in sorted(JOBS_DIR.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True)[:10]:
            if not d.is_dir():
                continue
            st_path = d / 'status.json'
            st = {}
            if st_path.exists():
                try:
                    st = json.loads(st_path.read_text(encoding='utf-8'))
                except Exception as exc:
                    st = {'status': 'status_read_error', 'error': str(exc)}
            disk_jobs.append({
                'job_id': d.name,
                'status': st.get('status'),
                'phase': st.get('phase'),
                'progress': st.get('progress'),
                'updated_at': st.get('updated_at'),
                'has_output': (d / 'comparativo.xlsx').exists(),
                'has_trace': (d / 'trace.jsonl').exists(),
            })
    except Exception as exc:
        disk_jobs = [{'error': f'{type(exc).__name__}: {exc}'}]
    return JSONResponse({
        'ok': True,
        'ai_policy': _safe_ai_policy_status(),
        'jobs_dir': str(JOBS_DIR),
        'data_dir': str(DATA_DIR),
        'active_memory_jobs': active_jobs,
        'memory_jobs': mem_jobs,
        'disk_jobs': disk_jobs,
    })

# STEP43: diagnóstico rápido de catálogos granulares usados por Detalle.
@app.get("/debug/catalogos")
def debug_catalogos():
    try:
        import processor as _p
        if hasattr(_p, 'quantia_market_catalog_diagnostics'):
            return _p.quantia_market_catalog_diagnostics(force=True)
        return {"status": "unavailable", "error": "processor.quantia_market_catalog_diagnostics no disponible"}
    except Exception as exc:
        return {"status": "error", "error": f"{type(exc).__name__}: {exc}"}



def _format_money_for_ai(v):
    try:
        if v is None:
            return "—"
        n = float(v)
        return f"${n:,.2f}"
    except Exception:
        return "—"


def _fallback_analysis_answer(question: str, context: dict) -> str:
    """Respuesta local para el chat del resultado.

    No recalcula precios ni inventa datos: resume el JSON de preview/dashboard enviado por el FE.
    """
    q = (question or "").lower()
    mode = context.get("mode") or "analysis"
    providers = context.get("provider_scores") or context.get("providers") or []
    critical = context.get("critical_concepts") or []
    kpis = context.get("kpis") or []
    preview = context.get("preview") or {}
    rows = context.get("rows") or preview.get("rows") or []
    totals = context.get("totals") or {}
    labor_findings = context.get("labor_findings") or []
    equipment_findings = context.get("equipment_findings") or []
    reference_gaps = context.get("reference_gaps") or []
    cost_breakdown = context.get("cost_breakdown") or {}

    if not rows and not providers and not critical and not totals and not cost_breakdown and not labor_findings and not equipment_findings and not reference_gaps:
        return (
            "El Excel ya fue generado, pero todavía no hay un resumen web suficiente para responder con evidencia. "
            "Descarga el Excel técnico o vuelve a procesar el análisis para generar el resumen estructurado. No voy a inventar conclusiones sin datos calculados."
        )

    if "proveedor" in q or "balance" in q or "mejor" in q:
        if providers:
            best = providers[0]
            name = best.get("nombre") or best.get("provider") or "Proveedor líder"
            total = best.get("total_oferta") or best.get("total_estimado") or best.get("total_comparable")
            dictamen = best.get("dictamen") or best.get("recommendation") or "mejor posicionado en el ranking disponible"
            return f"Con base en el resumen calculado, {name} aparece como mejor balance. Total reportado: {_format_money_for_ai(total)}. Lectura: {dictamen}. Recomiendo validar sus conceptos críticos en el Excel antes de tomar una decisión final."
        return "No hay ranking de proveedores disponible en el resumen web. Revisa las hojas Detalle del Excel."

    if "negoci" in q or "primero" in q or "crítico" in q or "critico" in q:
        if critical:
            items = []
            for r in critical[:5]:
                items.append(f"{r.get('code','—')}: impacto { _format_money_for_ai(r.get('impact')) }, riesgo {r.get('risk','Revisar')}. Acción: {r.get('action','Solicitar aclaración técnica y económica.')}")
            return "Priorizaría estos conceptos para revisión/negociación: " + " | ".join(items)
        return "No hay conceptos críticos estructurados en la vista web. Usa las hojas Detalle del Excel para ordenar por diferencia económica."

    if "mano" in q or "mo" in q or "material" in q or "equipo" in q or "herramienta" in q or "epp" in q or "montacargas" in q:
        parts = []
        if "mano" in q or "mo" in q:
            if labor_findings:
                top = []
                for r in labor_findings[:5]:
                    role = r.get("role") or r.get("code") or "MO"
                    cc = _format_money_for_ai(r.get("contractor_cost"))
                    mm = _format_money_for_ai(r.get("market_cost"))
                    pct = r.get("difference_pct")
                    pct_txt = f" ({pct:.1%})" if isinstance(pct, (int, float)) else ""
                    top.append(f"{role}: contratista {cc} vs mercado {mm}{pct_txt}")
                parts.append("Hallazgos de Mano de Obra: " + "; ".join(top) + ".")
            else:
                parts.append("No hay hallazgos estructurados de Mano de Obra en el resumen web; revisa el tab Detalle para roles y rendimientos.")
        if "equipo" in q or "montacargas" in q or "herramienta" in q or "epp" in q:
            if equipment_findings:
                top = []
                for r in equipment_findings[:5]:
                    item = r.get("item") or r.get("code") or "Equipo"
                    cc = _format_money_for_ai(r.get("contractor_cost"))
                    mm = _format_money_for_ai(r.get("market_cost"))
                    top.append(f"{item}: contratista {cc} vs mercado {mm}")
                parts.append("Hallazgos de equipo/herramienta: " + "; ".join(top) + ".")
            else:
                parts.append("No hay hallazgos estructurados de equipo crítico en el resumen web.")
        if "material" in q and cost_breakdown:
            material_keys = [k for k in cost_breakdown if "MATERIAL" in str(k).upper()]
            if material_keys:
                k = material_keys[0]
                b = cost_breakdown.get(k) or {}
                parts.append(f"Materiales acumulados: contratista {_format_money_for_ai(b.get('contractor'))} vs referencia {_format_money_for_ai(b.get('market'))}.")
        parts.append("Criterios PMD: indirecto mercado 25%, MO mercado 2026 con DOF 13%, herramienta menor 5%, EPP 10%, supervisión 1:5 y equipo crítico con soporte de renta/seguro/portes/operador/consumibles.")
        return " ".join(parts)

    if "falt" in q or "referencia" in q or "construdata" in q:
        if reference_gaps:
            items = []
            for r in reference_gaps[:6]:
                items.append(f"{r.get('code','—')}: {r.get('description','sin descripción')}")
            return "Conceptos/insumos con falta de referencia estructurada: " + " | ".join(items) + ". No uses la diferencia económica como criterio único; solicita desglose, soporte técnico y validación de alcance."
        count = len([r for r in rows if str(r.get("status", "")).lower().find("sin") >= 0 or str(r.get("riesgo", "")).lower().find("sin") >= 0])
        if count > 0:
            return f"Detecté {count} concepto(s) con señales de falta de referencia o riesgo similar en el resumen web. Para esos casos, no conviene usar la diferencia económica como criterio único; solicita desglose, soporte técnico y validación de alcance."
        return "No detecté conceptos con falta de referencia en el resumen estructurado disponible. Si sospechas un caso específico, dime el código de la partida para revisarlo contra el Detalle del Excel."

    if totals:
        c = _format_money_for_ai(totals.get("contractor_total"))
        m = _format_money_for_ai(totals.get("market_total"))
        d = _format_money_for_ai(totals.get("difference_amount"))
        pct = totals.get("difference_pct")
        pct_txt = f" ({pct:.1%})" if isinstance(pct, (int, float)) else ""
        return f"Resumen calculado: contratista {c}, referencia {m}, diferencia {d}{pct_txt}. Recomiendo revisar primero los conceptos de mayor impacto y cualquier partida sin referencia confiable."
    if kpis:
        plain = "; ".join([f"{k.get('label')}: {k.get('value')}" for k in kpis[:4]])
        return f"Resumen disponible: {plain}. Como recomendación general, revisa primero los conceptos de mayor impacto y cualquier partida sin referencia Construdata confiable."
    return "Puedo ayudarte a interpretar el análisis, pero necesito que la pregunta haga referencia a proveedor, negociación, mano de obra/materiales, conceptos críticos o referencias Construdata."


@app.post("/api/v2/analysis/ask")
async def ask_analysis(payload: dict = Body(default_factory=dict)):
    """Chat de apoyo sobre el análisis ya calculado.

    La IA, si está habilitada, redacta sobre el contexto calculado que envía el FE.
    Si no está habilitada, se responde con reglas locales. No recalcula precios.
    """
    question = str((payload or {}).get("question") or "").strip()
    if not question:
        raise HTTPException(400, "Escribe una pregunta sobre el análisis.")

    analysis_id = str((payload or {}).get("analysis_id") or (payload or {}).get("job_id") or "").strip()
    stored_summary = _read_analysis_summary(analysis_id) if analysis_id else None
    preview_payload = (payload or {}).get("preview") or {}
    # Prioriza el resumen estructurado del backend; usa payload del FE solo como fallback.
    context = stored_summary or {
        "mode": (payload or {}).get("mode"),
        "providers": (payload or {}).get("providers") or [],
        "provider_scores": (payload or {}).get("providers") or [],
        "critical_concepts": (payload or {}).get("critical_concepts") or [],
        "kpis": (payload or {}).get("kpis") or [],
        "preview": preview_payload,
        "rows": (preview_payload or {}).get("rows") or [],
    }
    if analysis_id:
        context["analysis_id"] = analysis_id
    # Normaliza llaves esperadas por el fallback.
    if "providers" not in context and "provider_scores" in context:
        context["providers"] = context.get("provider_scores") or []
    if "critical_concepts" not in context:
        context["critical_concepts"] = []

    # IA opcional y controlada. Nunca calcula ni modifica importes.
    try:
        if os.getenv("QUANTIA_ENABLE_AI", "0") == "1" and os.getenv("ENABLE_AI_EXPERT_TEXT", "0") == "1" and (os.getenv("ANTHROPIC_API_KEY") or "").strip():
            try:
                import quantia_ai_review as qair
            except Exception:
                from . import quantia_ai_review as qair
            ai_payload = {
                "system": (
                    "Eres un especialista senior PMD en precios unitarios, APU/Neodata y revisión de cotizaciones. "
                    "Responde solo con base en los datos calculados enviados. No inventes precios, no recalcules importes y distingue dato calculado de recomendación profesional. "
                    "Usa criterios PMD cuando aplique: indirecto mercado 25%, MO mercado 2026 con DOF 13%, herramienta menor 5%, EPP 10%, supervisión 1:5 y análisis de equipos críticos como montacargas."
                ),
                "messages": [{"role": "user", "content": json.dumps({"question": question, "context": context}, ensure_ascii=False)[:18000]}],
            }
            txt = qair._anthropic_message(ai_payload, max_tokens=int(os.getenv("QUANTIA_AI_CHAT_MAX_TOKENS", "700")), timeout=int(os.getenv("QUANTIA_AI_EXPERT_TIMEOUT", "30")))
            if txt:
                return {"answer": txt.strip(), "used_ai": True, "mode": context.get("mode")}
    except Exception as exc:
        # Fallback local si Claude falla.
        pass

    return {"answer": _fallback_analysis_answer(question, context), "used_ai": False, "mode": context.get("mode")}

# ── Generador de Matriz Propuesta IA desde catalogo base ──────────────────
@app.post("/api/v2/base-budget/propose-matrix")
async def propose_base_matrix(
    catalogo: UploadFile = File(...),
    proyecto: str = Form(default=""),
    cliente: str = Form(default="Nestlé"),
    ubicacion: str = Form(default=""),
    indirect_pct: Optional[float] = Form(default=None),
    sheet_name: Optional[str] = Form(default=None),
    start_row: Optional[int] = Form(default=None),
    end_row: Optional[int] = Form(default=None),
):
    """Genera una matriz propuesta estilo APU/Neodata usando solo catalogo base.

    No requiere cotizaciones, matriz de contratista ni Resumen PU. La fuente tecnica
    es data/construdata_matrices.xlsx y el analisis IA local selecciona/complementa
    matrices para proponer un presupuesto base auditable.
    """
    try:
        from .base_matrix_proposer import generate_matrix_proposal_workbook
    except Exception:
        from backend.base_matrix_proposer import generate_matrix_proposal_workbook

    if not catalogo or not catalogo.filename:
        raise HTTPException(400, "Carga un catalogo base del proyecto.")
    ext = Path(catalogo.filename).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        raise HTTPException(400, "El catalogo debe ser Excel (.xlsx o .xls).")
    if not CONSTRUDATA_MATRICES_PATH.exists():
        bundled = _project_data_dir() / "construdata_matrices.xlsx"
        raise HTTPException(400, f"No se encontro la base de matrices Construdata. Se reviso DATA_DIR={DATA_DIR} y bundled={bundled}.")

    job_id = uuid.uuid4().hex[:10]
    job_dir = _job_dir(f"matrix_{job_id}")
    job_dir.mkdir(parents=True, exist_ok=True)
    raw_path = job_dir / f"catalogo{ext}"
    with open(raw_path, "wb") as f:
        f.write(await catalogo.read())
    if ext == ".xls":
        xlsx_path = job_dir / "catalogo.xlsx"
        xls_to_xlsx(str(raw_path), str(xlsx_path))
        catalog_path = str(xlsx_path)
    else:
        catalog_path = str(raw_path)

    output_path = str(job_dir / "matriz_propuesta_ia.xlsx")
    pct = float(indirect_pct if indirect_pct is not None else os.getenv("BASE_BUDGET_INDIRECT_PCT", "0.25") or 0.25)
    try:
        result = generate_matrix_proposal_workbook(
            catalog_path=catalog_path,
            output_path=output_path,
            construdata_matrices_path=str(CONSTRUDATA_MATRICES_PATH),
            project_meta={"proyecto": proyecto, "cliente": cliente, "ubicacion": ubicacion},
            indirect_pct=pct,
            only_sheet=sheet_name or None,
            start_row=start_row,
            end_row=end_row,
        )
        print("matrix_proposal_done", json.dumps(result, ensure_ascii=False))
    except Exception as exc:
        tb = traceback.format_exc()
        print(tb)
        raise HTTPException(500, f"No se pudo generar la matriz propuesta: {exc}")

    filename = f"matriz_propuesta_ia_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    summary = _build_analysis_summary_from_excel(output_path, mode="base_matrix", provider_names=["Matriz base / Mercado"])
    summary.setdefault("files", {})["excel_url"] = f"/api/v2/base-budget/{job_id}/download"
    summary.setdefault("files", {})["filename"] = filename
    return JSONResponse({
        "status": "done",
        "run_type": "base_matrix",
        "job_id": job_id,
        "download_url": f"/api/v2/base-budget/{job_id}/download",
        "filename": filename,
        "analysis_summary": summary,
    })


# ── v0.3.13: dashboard web estructurado post-análisis ─────────────────────
def _v0313_norm_text(value):
    import unicodedata, re
    txt = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", txt.strip().lower())


def _v0313_sheet_text(ws, max_rows=120, max_cols=12):
    texts = []
    try:
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), min_col=1, max_col=min(ws.max_column, max_cols), values_only=True):
            parts = [str(v).strip() for v in row if isinstance(v, str) and str(v).strip()]
            if parts:
                texts.append(" ".join(parts))
    except Exception:
        pass
    return "\n".join(texts).strip()


def _v0313_detail_provider_name(sheet_name, idx, provider_names=None):
    import re
    provider_names = provider_names or []
    txt = str(sheet_name or "").strip()
    m = re.search(r"Detalle\s*-\s*P(\d+)", txt, flags=re.I)
    if m:
        pos = int(m.group(1)) - 1
        if 0 <= pos < len(provider_names) and provider_names[pos]:
            return str(provider_names[pos])
        return f"Proveedor {pos + 1}"
    if provider_names:
        return str(provider_names[0])
    return "Matriz base / Mercado" if txt.lower() == "detalle" else txt


def _v0313_empty_family_dict():
    return {
        "materiales": 0.0,
        "mano_obra": 0.0,
        "equipo_herramienta": 0.0,
        "basicos": 0.0,
        "otros": 0.0,
        "costo_indirecto": 0.0,
        "financiamiento": 0.0,
        "utilidad": 0.0,
        "total_costo_unitario": 0.0,
    }


def _v0313_family_key(label):
    txt = _v0313_norm_text(label)
    if "material" in txt:
        return "materiales"
    if "mano" in txt or txt in {"mo", "m_o"}:
        return "mano_obra"
    if "equipo" in txt or "herramient" in txt or "maquinaria" in txt:
        return "equipo_herramienta"
    if "basic" in txt or "basico" in txt:
        return "basicos"
    return "otros"


def _v0313_financial_key(label):
    txt = _v0313_norm_text(label)
    if "costo indirect" in txt or txt == "ci":
        return "costo_indirecto"
    if "financ" in txt:
        return "financiamiento"
    if "util" in txt or "cargo" in txt:
        return "utilidad"
    if "total costo unitario" in txt or "total por servicio" in txt:
        return "total_costo_unitario"
    return None


def _v0313_parse_dashboard_from_excel(output_path: str, mode: str = None, provider_names=None):
    """Lee el Excel final y produce JSON pensado para el dashboard web.

    Fuente de verdad: hojas Comparativa, Detalle/Detalle - Pn y Analisis IA.
    No depende de datos hardcodeados ni obliga al frontend a recalcular todo.
    """
    import openpyxl, datetime, json, math
    provider_names = provider_names or []
    result = {
        "version": "v0.3.13-dashboard-web",
        "run_type": mode or "unknown",
        "mode": mode or "unknown",
        "generated_at": datetime.datetime.utcnow().isoformat() + "Z",
        "summary": {
            "total_contractors": 0,
            "best_candidate": None,
            "has_ai_analysis": False,
            "has_market": False,
            "has_multiple_contractors": False,
            "has_base_matrix": False,
        },
        "contractors": [],
        "providers": [],
        "provider_scores": [],
        "rows": [],
        "critical_concepts": [],
        "top_concepts": [],
        "cost_breakdown": {},
        "charts": {"totals_by_contractor": [], "families_by_contractor": [], "market_vs_contractor": []},
        "ai_analysis": {"available": False, "text": None},
        "files": {},
        "dashboard": {"tabs": ["resumen", "familias", "grafica_totales", "analisis_ia"]},
        "data_quality": {"has_structured_summary": False, "warnings": []},
    }
    try:
        wb = openpyxl.load_workbook(output_path, data_only=True, read_only=True)
    except Exception as exc:
        result["data_quality"]["warnings"].append(f"No se pudo leer Excel para dashboard: {type(exc).__name__}: {exc}")
        return result

    try:
        sheet_names = list(wb.sheetnames)
        detail_names = [n for n in sheet_names if str(n).strip().lower() == "detalle" or str(n).strip().lower().startswith("detalle - p")]
        has_ai = "Analisis IA" in sheet_names or "Analisis experto IA" in sheet_names
        has_market_cols = False
        contractors = []
        critical = []
        family_chart = []

        # Detectar run_type si no viene claro.
        if not mode or mode in {"unknown", "single_provider", "multi_provider"}:
            if detail_names == ["Detalle"] and not has_ai:
                mode = "base_matrix"
            elif len(detail_names) > 1:
                mode = "multi_contractor"
            elif len(detail_names) == 1:
                mode = "single_contractor"
        result["run_type"] = mode
        result["mode"] = mode
        result["summary"]["has_base_matrix"] = mode == "base_matrix"

        for idx, name in enumerate(detail_names):
            ws = wb[name]
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            hmap = {_norm_header_key(v): c for c, v in enumerate(headers, start=1) if v not in (None, "")}
            # Cabeceras aprobadas del Detalle v0.3.x.
            c_code = hmap.get("codigo") or 1
            c_desc = hmap.get("concepto") or 2
            c_unit = hmap.get("unidad") or 3
            c_pu = hmap.get("p_unitario") or 4
            c_op = hmap.get("op") or 5
            c_qty = hmap.get("cantidad") or 6
            c_amount = hmap.get("importe") or 7
            c_pct = hmap.get("") or 8
            c_mpu = hmap.get("mercado_p_unitario")
            c_mop = hmap.get("mercado_op")
            c_mqty = hmap.get("mercado_cantidad")
            c_mamount = hmap.get("mercado_importe")
            sheet_has_market = bool(c_mpu and c_mamount)
            has_market_cols = has_market_cols or sheet_has_market
            provider = "Matriz base / Mercado" if mode == "base_matrix" else _v0313_detail_provider_name(name, idx, provider_names)
            families = _v0313_empty_family_dict()
            market_families = _v0313_empty_family_dict()
            total = 0.0
            market_total = 0.0
            items = []
            current_service = None
            current_family = None
            service_amount = {}

            for r in range(2, ws.max_row + 1):
                code = ws.cell(r, c_code).value if c_code else None
                desc = ws.cell(r, c_desc).value if c_desc else None
                unit = ws.cell(r, c_unit).value if c_unit else None
                pu = _to_float_safe(ws.cell(r, c_pu).value if c_pu else None)
                op = ws.cell(r, c_op).value if c_op else None
                qty = _to_float_safe(ws.cell(r, c_qty).value if c_qty else None)
                amount = _to_float_safe(ws.cell(r, c_amount).value if c_amount else None)
                pct = _to_float_safe(ws.cell(r, c_pct).value if c_pct else None)
                mpu = _to_float_safe(ws.cell(r, c_mpu).value if c_mpu else None) if c_mpu else None
                mamount = _to_float_safe(ws.cell(r, c_mamount).value if c_mamount else None) if c_mamount else None
                text = str(desc or "").strip()
                norm = _v0313_norm_text(text)
                if not any([code, desc, unit, pu, qty, amount, mpu, mamount]):
                    continue
                # Servicio/matriz principal: trae código y descripción, pero no pertenece a una familia activa.
                if code and text and (norm not in {"materiales", "mano de obra", "equipo y herramienta", "basicos", "otros"}) and (current_family is None or norm.startswith("instalacion") or str(code).upper().startswith("FLEX")) and not norm.startswith("subtotal"):
                    # Si parece renglón de insumo dentro de familia, no cambiar servicio.
                    if current_family is None or str(code).upper().startswith(("FLEX", "SERV", "CONC")):
                        current_service = {"code": str(code), "description": text, "amount": amount, "market_amount": mamount}
                        if amount:
                            service_amount[str(code)] = (service_amount.get(str(code)) or 0) + amount
                        continue
                if norm in {"materiales", "mano de obra", "equipo y herramienta", "basicos", "básicos", "otros", "seccion financiera", "seccion financiera"}:
                    current_family = None if "financiera" in norm else _v0313_family_key(text)
                    continue
                if norm.startswith("subtotal"):
                    fkey = _v0313_family_key(text)
                    if amount is not None:
                        families[fkey] = families.get(fkey, 0.0) + amount
                    if mamount is not None:
                        market_families[fkey] = market_families.get(fkey, 0.0) + mamount
                    continue
                finkey = _v0313_financial_key(text)
                if finkey:
                    if amount is not None:
                        families[finkey] = families.get(finkey, 0.0) + amount
                    if mamount is not None:
                        market_families[finkey] = market_families.get(finkey, 0.0) + mamount
                    if finkey == "total_costo_unitario" and amount is not None:
                        total += amount
                    if finkey == "total_costo_unitario" and mamount is not None:
                        market_total += mamount
                    continue
                # Renglón técnico dentro de familia.
                if current_family and amount is not None:
                    items.append({
                        "code": str(code or ""),
                        "description": text,
                        "family": current_family,
                        "amount": amount,
                        "market_amount": mamount,
                        "difference_amount": (amount - mamount) if mamount is not None else None,
                        "difference_percent": ((amount - mamount) / mamount) if mamount not in (None, 0) else None,
                        "source": provider,
                        "provider": provider,
                    })

            if not total:
                # fallback: suma familias directas si no hubo TOTAL COSTO UNITARIO.
                total = sum(families.get(k, 0.0) for k in ["materiales", "mano_obra", "equipo_herramienta", "basicos", "otros", "costo_indirecto", "financiamiento", "utilidad"])
            if sheet_has_market and not market_total:
                market_total = sum(market_families.get(k, 0.0) for k in ["materiales", "mano_obra", "equipo_herramienta", "basicos", "otros", "costo_indirecto", "financiamiento", "utilidad"])
            diff_amount = (total - market_total) if sheet_has_market and market_total else None
            diff_pct = (diff_amount / market_total) if diff_amount is not None and market_total else None
            row = {
                "id": "BASE" if mode == "base_matrix" else f"P{idx+1}",
                "name": provider,
                "nombre": provider,
                "provider": provider,
                "total": total or None,
                "total_estimado": total or None,
                "total_oferta": total or None,
                "market_total": market_total or None,
                "total_mercado": market_total or None,
                "difference_amount": diff_amount,
                "diferencia_total": diff_amount,
                "difference_percent": diff_pct,
                "desviacion_total_pct": diff_pct,
                "families": families,
                "market_families": market_families if sheet_has_market else None,
                "top_items": sorted(items, key=lambda x: abs(x.get("amount") or 0), reverse=True)[:8],
                "business_rules": [],
                "conceptos": len(items),
                "dictamen": "Matriz base usada como referencia de mercado." if mode == "base_matrix" else "Fuente incluida en dashboard comparativo.",
            }
            # Reglas de negocio simples, basadas en evidencia disponible.
            if row["top_items"]:
                row["business_rules"].append("Las partidas/familias principales se concentran en: " + ", ".join([x["code"] or x["description"][:24] for x in row["top_items"][:3]]) + ".")
            if diff_pct is not None:
                if diff_pct > 0.15:
                    row["business_rules"].append(f"La fuente se ubica {diff_pct:.1%} arriba del mercado calculado; revisar partidas de mayor impacto.")
                elif diff_pct < -0.15:
                    row["business_rules"].append(f"La fuente se ubica {abs(diff_pct):.1%} abajo del mercado; validar alcance y omisiones antes de considerarlo ahorro real.")
                else:
                    row["business_rules"].append("La fuente está razonablemente alineada con el mercado calculado.")
            contractors.append(row)
            for k, v in families.items():
                family_chart.append({"source": provider, "family": k, "amount": v})
            # Criticos desde top_items con diferencia.
            for it in row["top_items"]:
                if it.get("difference_amount") is not None and abs(it.get("difference_amount") or 0) > 0:
                    critical.append({
                        "code": it.get("code"), "description": it.get("description"), "provider": provider,
                        "impact": it.get("difference_amount"), "difference_pct": it.get("difference_percent"),
                        "risk": "Alta" if abs(it.get("difference_percent") or 0) >= 0.30 else "Media",
                        "action": "Revisar soporte de precio, alcance y match de mercado en el Excel técnico.",
                    })

        contractors = sorted(contractors, key=lambda x: x.get("total") if x.get("total") is not None else 9e99)
        for i, c in enumerate(contractors, 1):
            c["rank"] = i
        result["contractors"] = contractors
        result["providers"] = contractors
        result["provider_scores"] = contractors
        result["summary"]["total_contractors"] = len(contractors)
        result["summary"]["best_candidate"] = contractors[0]["name"] if contractors and mode != "base_matrix" else None
        result["summary"]["has_ai_analysis"] = has_ai
        result["summary"]["has_market"] = has_market_cols
        result["summary"]["has_multiple_contractors"] = len(contractors) > 1
        result["charts"]["totals_by_contractor"] = [{"name": c["name"], "total": c.get("total"), "market_total": c.get("market_total")} for c in contractors]
        result["charts"]["families_by_contractor"] = family_chart
        result["charts"]["market_vs_contractor"] = [{"name": c["name"], "contractor": c.get("total"), "market": c.get("market_total")} for c in contractors if c.get("market_total")]
        # cost_breakdown compatible legado: cuando hay una fuente, objeto por familia; cuando hay varias, conservar por fuente.
        if len(contractors) == 1:
            result["cost_breakdown"] = {k: {"contractor": v, "market": (contractors[0].get("market_families") or {}).get(k), "rows": 0} for k, v in (contractors[0].get("families") or {}).items()}
        else:
            result["cost_breakdown"] = {c["name"]: c.get("families") for c in contractors}
        result["critical_concepts"] = sorted(critical, key=lambda x: abs(x.get("impact") or 0), reverse=True)[:12]
        result["top_concepts"] = result["critical_concepts"][:8]
        result["rows"] = [it for c in contractors for it in (c.get("top_items") or [])]
        # Analisis IA: hoja nueva v0.3.8+ o nombre anterior.
        ai_sheet = wb["Analisis IA"] if "Analisis IA" in wb.sheetnames else (wb["Analisis experto IA"] if "Analisis experto IA" in wb.sheetnames else None)
        if ai_sheet:
            text = _v0313_sheet_text(ai_sheet)
            if text:
                result["ai_analysis"] = {"available": True, "text": text}
                result["ai_findings"] = [{"text": text[:4000]}]
        else:
            result["dashboard"]["tabs"] = [t for t in result["dashboard"]["tabs"] if t != "analisis_ia"]
        result["data_quality"]["has_structured_summary"] = bool(contractors)
        if not contractors:
            result["data_quality"]["warnings"].append("No se encontraron datos suficientes en Detalle para el dashboard.")
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return result


# Sobrescribe el extractor web anterior: mantiene compatibilidad de nombres legacy.
def _build_analysis_summary_from_excel(output_path: str, mode: str = None, provider_names=None):
    return _v0313_parse_dashboard_from_excel(output_path, mode=mode, provider_names=provider_names)


@app.get("/api/v2/base-budget/{job_id}/download")
def download_base_matrix_result(job_id: str):
    job_dir = _job_dir(f"matrix_{job_id}")
    output_path = job_dir / "matriz_propuesta_ia.xlsx"
    if not output_path.exists():
        raise HTTPException(404, "No se encontro el Excel de matriz base para descargar.")
    return FileResponse(str(output_path), filename=f"matriz_propuesta_ia_{job_id}.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
