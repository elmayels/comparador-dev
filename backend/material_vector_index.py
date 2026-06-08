"""
material_vector_index.py

Indice vectorial persistente para Construdata.

Objetivo:
- Vectorizar cada registro del catálogo Construdata una sola vez por versión de catálogo/modelo Voyage.
- En cada cotización, generar embedding solo del material único y buscar top candidatos localmente.
- Reducir dependencia de búsqueda textual y consumo de tokens/llamadas.
"""

import hashlib
import json
import os
import time
import urllib.request
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:
    from material_ai_trace import record_ai_event
except Exception:  # pragma: no cover
    def record_ai_event(*args, **kwargs):
        return None

try:
    from material_normalizer import normalize_material_description, normalize_unit
except Exception:  # pragma: no cover
    def normalize_unit(x):
        return (str(x or "").strip().lower())
    def normalize_material_description(desc, unidad=None):
        return {"normalized": str(desc or "").strip().upper(), "query": str(desc or "").strip()}

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
    return Path(__file__).resolve().parent / "data"

DATA_DIR = _default_data_dir()
VECTOR_DIR = DATA_DIR / "vector_index"
VECTOR_DIR.mkdir(parents=True, exist_ok=True)

def get_vector_index_dir() -> str:
    """Directorio persistente donde se guardan/suben indices Voyage."""
    VECTOR_DIR.mkdir(parents=True, exist_ok=True)
    return str(VECTOR_DIR)

DEFAULT_VOYAGE_MODEL = os.getenv("VOYAGE_MODEL", "voyage-4-lite")
DEFAULT_BATCH_SIZE = int(os.getenv("VOYAGE_INDEX_BATCH_SIZE", "96"))
DEFAULT_TOP_K = int(os.getenv("MATERIAL_VECTOR_TOP_K", "25"))


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return default


def _cosine(a, b) -> float:
    try:
        num = sum(float(x) * float(y) for x, y in zip(a, b))
        da = sum(float(x) * float(x) for x in a) ** 0.5
        db = sum(float(y) * float(y) for y in b) ** 0.5
        return num / (da * db) if da and db else 0.0
    except Exception:
        return 0.0


def catalog_fingerprint(materials_index: Iterable[dict]) -> str:
    """Hash estable de contenido relevante del catálogo."""
    h = hashlib.sha256()
    count = 0
    for item in materials_index or []:
        count += 1
        parts = [
            str(item.get("clave") or ""),
            str(item.get("descripcion") or item.get("descripcion_corta") or ""),
            str(item.get("unidad") or ""),
            str(item.get("precio") or ""),
            str(item.get("tipo_insumo") or ""),
        ]
        h.update(("|".join(parts) + "\n").encode("utf-8", errors="ignore"))
    h.update(f"count={count}".encode("utf-8"))
    return h.hexdigest()[:16]


def _index_path(catalog_hash: str, model: Optional[str] = None) -> Path:
    model = (model or os.getenv("VOYAGE_MODEL") or DEFAULT_VOYAGE_MODEL).replace("/", "_")
    return VECTOR_DIR / f"construdata_{catalog_hash}_{model}.json"


def _record_text(item: dict) -> str:
    desc = item.get("descripcion") or item.get("descripcion_corta") or ""
    unidad = item.get("unidad") or ""
    tipo = item.get("tipo_insumo") or item.get("tipo_kind") or ""
    attrs = []
    for k in ("tipo", "material", "grado", "diametro"):
        if item.get(k):
            attrs.append(f"{k}:{item.get(k)}")
    return " | ".join([str(desc), f"unidad:{unidad}", f"tipo_insumo:{tipo}", " ".join(attrs)]).strip()


def _record_payload(item: dict, idx: int) -> dict:
    return {
        "idx": idx,
        "clave": item.get("clave"),
        "descripcion": item.get("descripcion") or item.get("descripcion_corta"),
        "descripcion_corta": item.get("descripcion_corta"),
        "unidad": item.get("unidad"),
        "unidad_norm": item.get("unidad_norm") or normalize_unit(item.get("unidad")),
        "precio": item.get("precio"),
        "tipo_insumo": item.get("tipo_insumo"),
        "tipo_kind": item.get("tipo_kind"),
        "tokens": item.get("tokens") or [],
        "tipo": item.get("tipo"),
        "material": item.get("material"),
        "grado": item.get("grado"),
        "diametro": item.get("diametro"),
        "text": _record_text(item),
    }


def _normalize_voyage_text(text) -> str:
    """Texto seguro para Voyage: no vacio y limitado para evitar fallos por batch."""
    t = str(text or "").strip()
    if not t:
        t = "MATERIAL SIN DESCRIPCION"
    # Evita payloads demasiado grandes por descripcion larga accidental.
    return t[:1800]


def _voyage_embeddings(texts: List[str], *, model: Optional[str] = None, timeout: int = 45, max_retries: int = 4) -> Tuple[Optional[List[List[float]]], Optional[dict]]:
    """Llama Voyage con diagnostico util y reintentos.

    Devuelve (embeddings, error). Si hay error HTTP o respuesta parcial, no lo oculta.
    """
    key = os.getenv("VOYAGE_API_KEY")
    if not key:
        return None, {"type": "config", "message": "VOYAGE_API_KEY no configurada"}
    model = model or os.getenv("VOYAGE_MODEL") or DEFAULT_VOYAGE_MODEL
    safe_texts = [_normalize_voyage_text(t) for t in texts]
    payload = json.dumps({"model": model, "input": safe_texts}).encode("utf-8")
    last_error = None

    for attempt in range(max(1, max_retries)):
        req = urllib.request.Request(
            "https://api.voyageai.com/v1/embeddings",
            data=payload,
            headers={"content-type": "application/json", "Authorization": f"Bearer {key}"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = json.loads(resp.read().decode("utf-8"))
            data = raw.get("data") or []
            try:
                if data and all(isinstance(x, dict) and "index" in x for x in data):
                    data = sorted(data, key=lambda x: int(x.get("index", 0)))
            except Exception:
                pass
            out = []
            for x in data:
                if isinstance(x, dict) and isinstance(x.get("embedding"), list) and x.get("embedding"):
                    out.append(x.get("embedding"))
            if len(out) == len(safe_texts):
                record_ai_event("voyage", "embeddings", ok=True, detail=f"texts={len(safe_texts)} returned={len(out)} model={model}", candidates=len(safe_texts))
                return out, None
            last_error = {"type": "partial_response", "message": f"returned={len(out)} expected={len(safe_texts)}", "status": None}
            record_ai_event("voyage", "embeddings", ok=False, detail=json.dumps(last_error, ensure_ascii=False), candidates=len(safe_texts))
        except Exception as exc:
            body = ""
            status = getattr(exc, "code", None)
            reason = getattr(exc, "reason", str(exc))
            try:
                body = exc.read().decode("utf-8", errors="ignore")[:800]
            except Exception:
                body = ""
            if status:
                last_error = {"type": "HTTPError", "status": status, "reason": str(reason), "body": body}
            else:
                last_error = {"type": type(exc).__name__, "message": str(exc)}
            record_ai_event("voyage", "embeddings", ok=False, detail=json.dumps(last_error, ensure_ascii=False), candidates=len(safe_texts))

        status = (last_error or {}).get("status")
        if status in (401, 403):
            break
        delay = min(20.0, 1.5 * (2 ** attempt))
        if status == 429:
            delay = min(45.0, 8.0 * (attempt + 1))
        time.sleep(delay)
    return None, last_error or {"type": "unknown", "message": "Voyage no devolvio embeddings"}


def _voyage_embeddings_resilient(texts: List[str], *, model: Optional[str] = None, timeout: int = 90, min_batch: int = 1) -> Tuple[List[Optional[List[float]]], dict, List[dict]]:
    """Obtiene embeddings con reintentos y motivos de omision.

    Evita convertir un rate-limit en miles de omitidos. Solo omite registros cuando el item queda aislado
    o cuando el error es no recuperable.
    """
    texts = [_normalize_voyage_text(t) for t in texts]
    stats = {"calls": 0, "failed_items": 0, "partial_batches": 0, "split_batches": 0, "rate_limit_batches": 0, "http_errors": 0}
    failures: List[dict] = []

    def classify(err: Optional[dict]) -> str:
        if not err:
            return "unknown"
        if err.get("type") == "partial_response":
            return "partial_response"
        if err.get("type") == "HTTPError":
            stats["http_errors"] += 1
            status = err.get("status")
            if status == 429:
                return "rate_limit"
            if status in (500, 502, 503, 504):
                return "server_error"
            if status in (400, 413):
                return "bad_payload"
            return f"http_{status}"
        return err.get("type") or "unknown"

    def run(chunk: List[str], offset: int = 0) -> List[Optional[List[float]]]:
        if not chunk:
            return []
        stats["calls"] += 1
        embs, err = _voyage_embeddings(chunk, model=model, timeout=timeout)
        if embs is not None and len(embs) == len(chunk):
            return embs

        kind = classify(err)
        if kind == "partial_response":
            stats["partial_batches"] += 1
        if kind == "rate_limit":
            stats["rate_limit_batches"] += 1
            time.sleep(30)
            stats["calls"] += 1
            embs2, err2 = _voyage_embeddings(chunk, model=model, timeout=timeout, max_retries=2)
            if embs2 is not None and len(embs2) == len(chunk):
                return embs2
            err = err2 or err
            kind = classify(err)

        if len(chunk) > min_batch and kind not in ("http_401", "http_403"):
            stats["split_batches"] += 1
            mid = max(1, len(chunk) // 2)
            return run(chunk[:mid], offset) + run(chunk[mid:], offset + mid)

        stats["failed_items"] += len(chunk)
        detail = json.dumps(err or {"type": kind}, ensure_ascii=False)[:600]
        for i, t in enumerate(chunk):
            failures.append({"local_index": offset + i, "reason": kind, "detail": detail, "text_preview": t[:160]})
        return [None for _ in chunk]

    result = run(texts, 0)
    return result, stats, failures

def get_vector_index_status(materials_index: Optional[List[dict]] = None) -> dict:
    model = os.getenv("VOYAGE_MODEL") or DEFAULT_VOYAGE_MODEL
    catalog_hash = catalog_fingerprint(materials_index or []) if materials_index else None
    path = _index_path(catalog_hash, model) if catalog_hash else None
    existing = sorted(VECTOR_DIR.glob(f"construdata_*_{model}.json"), key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
    active_path = path if path and path.exists() else (existing[0] if existing else None)
    payload = {
        "model": model,
        "catalog_hash": catalog_hash,
        "exists": bool(active_path and active_path.exists()),
        "path": str(active_path) if active_path else None,
        "records": 0,
        "created_at": None,
        "updated_at": None,
    }
    if active_path and active_path.exists():
        try:
            data = json.loads(active_path.read_text(encoding="utf-8"))
            payload.update({
                "records": len(data.get("records") or []),
                "total_catalog_records": data.get("total_catalog_records") or len(data.get("records") or []),
                "skipped_records": data.get("skipped_records") or 0,
                "retry_stats": data.get("retry_stats") or {},
                "created_at": data.get("created_at"),
                "updated_at": data.get("updated_at"),
                "catalog_hash": data.get("catalog_hash") or catalog_hash,
                "model": data.get("model") or model,
            })
        except Exception as exc:
            payload["error"] = str(exc)
    return payload


def ensure_construdata_vector_index(materials_index: List[dict], *, force: bool = False, batch_size: Optional[int] = None, progress_callback=None) -> dict:
    """Crea/reutiliza el índice vectorial persistente.

    v16: construcción resiliente.
    - No aborta todo el índice si Voyage devuelve un batch incompleto.
    - Reintenta dividiendo el batch hasta aislar filas problemáticas.
    - Guarda índice con los registros embebidos correctamente y reporta skipped_records.
    """
    model = os.getenv("VOYAGE_MODEL") or DEFAULT_VOYAGE_MODEL
    batch_size = int(batch_size or os.getenv("VOYAGE_INDEX_BATCH_SIZE", DEFAULT_BATCH_SIZE))
    materials = list(materials_index or [])
    catalog_hash = catalog_fingerprint(materials)
    path = _index_path(catalog_hash, model)

    if path.exists() and not force:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return {
                "ok": True,
                "reused": True,
                "path": str(path),
                "records": len(data.get("records") or []),
                "total_catalog_records": data.get("total_catalog_records") or len(data.get("records") or []),
                "skipped_records": data.get("skipped_records") or 0,
                "catalog_hash": catalog_hash,
                "model": model,
            }
        except Exception:
            pass

    if not os.getenv("VOYAGE_API_KEY"):
        return {"ok": False, "reused": False, "path": str(path), "records": 0, "catalog_hash": catalog_hash, "model": model, "error": "VOYAGE_API_KEY no configurada; no se puede construir índice vectorial."}

    records = [_record_payload(item, i) for i, item in enumerate(materials)]
    total_batches = ((len(records) + batch_size - 1) // batch_size) if batch_size else 0
    if progress_callback:
        try:
            progress_callback({"phase": "started", "total_records": len(records), "batch_size": batch_size, "total_batches": total_batches, "records_done": 0, "batches_done": 0})
        except Exception:
            pass
    built = []
    skipped = []
    retry_stats = {"calls": 0, "failed_items": 0, "partial_batches": 0, "split_batches": 0}
    started = time.time()

    for start in range(0, len(records), batch_size):
        batch = records[start:start + batch_size]
        batch_no = (start // batch_size) + 1
        if progress_callback:
            try:
                progress_callback({"phase": "batch_started", "batch_no": batch_no, "total_batches": total_batches, "start": start, "end": min(start + len(batch), len(records)), "records_done": start, "built_records": len(built), "skipped_records": len(skipped)})
            except Exception:
                pass
        texts = [r["text"] for r in batch]
        embs, stats, failures = _voyage_embeddings_resilient(texts, model=model, timeout=90)
        failure_by_local = {int(f.get("local_index", -1)): f for f in (failures or [])}
        for k, v in stats.items():
            retry_stats[k] = retry_stats.get(k, 0) + int(v or 0)
        for local_i, (r, e) in enumerate(zip(batch, embs)):
            if e:
                r["embedding"] = e
                built.append(r)
            else:
                failure = failure_by_local.get(local_i, {})
                skipped.append({"idx": r.get("idx"), "clave": r.get("clave"), "descripcion": r.get("descripcion"), "reason": failure.get("reason") or "embedding_missing", "detail": failure.get("detail") or "Voyage no devolvio embedding para este registro"})
        if progress_callback:
            try:
                progress_callback({"phase": "batch_done", "batch_no": batch_no, "total_batches": total_batches, "records_done": min(start + len(batch), len(records)), "built_records": len(built), "skipped_records": len(skipped), "retry_stats": dict(retry_stats), "skipped_preview": skipped[-10:]})
            except Exception:
                pass

    ok = len(built) > 0
    payload = {
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "catalog_hash": catalog_hash,
        "model": model,
        "records": built,
        "total_catalog_records": len(records),
        "skipped_records": len(skipped),
        "skipped_preview": skipped[:25],
        "retry_stats": retry_stats,
        "elapsed_seconds": round(time.time() - started, 2),
    }
    if ok:
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)

    result = {
        "ok": ok,
        "reused": False,
        "path": str(path),
        "records": len(built),
        "total_catalog_records": len(records),
        "skipped_records": len(skipped),
        "catalog_hash": catalog_hash,
        "model": model,
        "retry_stats": retry_stats,
        "elapsed_seconds": payload["elapsed_seconds"],
    }
    if not ok:
        result["error"] = "Voyage no devolvio embeddings utiles para ningun registro del catalogo. Revisar VOYAGE_API_KEY/modelo/conectividad."
    elif skipped:
        result["warning"] = f"Indice construido parcialmente: {len(built)} de {len(records)} registros con embedding; {len(skipped)} omitidos por error de Voyage."
    if progress_callback:
        try:
            progress_callback({"phase": "finished", "result": dict(result)})
        except Exception:
            pass
    return result

def load_vector_index_for_catalog(materials_index: List[dict]) -> Optional[dict]:
    model = os.getenv("VOYAGE_MODEL") or DEFAULT_VOYAGE_MODEL
    catalog_hash = catalog_fingerprint(materials_index or [])
    expected_path = _index_path(catalog_hash, model)

    # Prefer exact catalog hash. If an index was uploaded manually to /data/vector_index,
    # fall back to the newest compatible Construdata index for the active Voyage model.
    candidates = []
    if expected_path.exists():
        candidates.append(expected_path)
    try:
        for p in sorted(VECTOR_DIR.glob(f"construdata_*_{model}.json"), key=lambda x: x.stat().st_mtime, reverse=True):
            if p not in candidates:
                candidates.append(p)
    except Exception:
        pass

    for path in candidates:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if data.get("records"):
                data["_loaded_from"] = str(path)
                data["_expected_catalog_hash"] = catalog_hash
                data["_exact_catalog_match"] = (data.get("catalog_hash") == catalog_hash)
                return data
        except Exception:
            continue
    return None

def vector_search_candidates(item: dict, materials_index: List[dict], *, top_k: Optional[int] = None) -> Tuple[List[dict], dict]:
    """Busca top candidatos Construdata usando índice vectorial persistente."""
    if not runtime_ai_allowed("material_matching"):
        return [], {"ok": False, "reason": "runtime_ai_disabled", "detail": external_ai_block_reason("material_matching")}

    top_k = int(top_k or os.getenv("MATERIAL_VECTOR_TOP_K", DEFAULT_TOP_K))
    index = load_vector_index_for_catalog(materials_index or [])
    if not index:
        return [], {"ok": False, "reason": "vector_index_missing"}

    model = index.get("model") or os.getenv("VOYAGE_MODEL") or DEFAULT_VOYAGE_MODEL
    normalized = normalize_material_description(item.get("descripcion"), item.get("unidad"))
    query = f"{normalized.get('normalized') or item.get('descripcion') or ''} | unidad:{item.get('unidad') or ''}".strip()
    emb, emb_error = _voyage_embeddings([query], model=model, timeout=30)
    if not emb or not emb[0]:
        return [], {"ok": False, "reason": "query_embedding_failed", "model": model, "error": emb_error}

    q = emb[0]
    scored = []
    input_unit = normalize_unit(item.get("unidad"))
    for rec in index.get("records") or []:
        sim = _cosine(q, rec.get("embedding") or [])
        unit_bonus = 0.025 if input_unit and input_unit == rec.get("unidad_norm") else 0.0
        score = max(0.0, min(0.999, sim + unit_bonus))
        cand = {k: rec.get(k) for k in ("clave", "descripcion", "descripcion_corta", "unidad", "unidad_norm", "precio", "tipo_insumo", "tipo_kind", "tokens", "tipo", "material", "grado", "diametro")}
        cand.update({
            "score": round(score, 4),
            "voyage_score": round(sim, 4),
            "fuente_precio": "construdata vector index",
            "candidate_reason": f"vector_index model={model} sim={round(sim,4)}",
        })
        scored.append(cand)
    scored.sort(key=lambda x: x.get("score") or 0, reverse=True)
    return scored[:top_k], {"ok": True, "model": model, "catalog_hash": index.get("catalog_hash"), "records_scanned": len(index.get("records") or []), "top_k": top_k}
