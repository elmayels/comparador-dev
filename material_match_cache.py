"""Cache/historial persistente de decisiones de match de materiales."""
import hashlib
import json
import os
import time
from pathlib import Path

try:
    from material_normalizer import normalize_material_description, normalize_unit
except Exception:  # pragma: no cover
    def normalize_unit(x): return str(x or "").strip().lower()
    def normalize_material_description(desc, unidad=None): return {"normalized": str(desc or "").strip().upper()}

try:
    from material_vector_index import catalog_fingerprint
except Exception:  # pragma: no cover
    def catalog_fingerprint(materials_index): return "unknown"

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
CACHE_DIR = DATA_DIR / "cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
CACHE_FILE = CACHE_DIR / "material_match_cache.json"
HISTORY_FILE = CACHE_DIR / "material_decision_history.json"


def _load(path: Path):
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except Exception:
            return {}
    return {}


def _save(path: Path, payload: dict):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def material_cache_key(item: dict, materials_index=None, threshold=None) -> str:
    n = normalize_material_description(item.get("descripcion"), item.get("unidad"))
    normalized = n.get("normalized") or str(item.get("descripcion") or "").strip().upper()
    unit = normalize_unit(item.get("unidad"))
    catalog_hash = catalog_fingerprint(materials_index or []) if materials_index is not None else "no_catalog"
    threshold = str(threshold if threshold is not None else os.getenv("MATERIAL_MATCH_THRESHOLD", "60"))
    voyage_model = os.getenv("VOYAGE_MODEL", "voyage-4-lite")
    anthropic_model = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-6")
    raw = "|".join([normalized, unit, catalog_hash, threshold, voyage_model, anthropic_model])
    return hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest()


def get_cached_material_match(item: dict, materials_index=None, threshold=None):
    if os.getenv("DISABLE_MATERIAL_MATCH_CACHE", "").strip():
        return None
    cache = _load(CACHE_FILE)
    key = material_cache_key(item, materials_index, threshold)
    val = cache.get(key)
    if isinstance(val, dict):
        val = dict(val)
        val["cache_hit"] = True
        return val
    return None


def save_material_match(item: dict, result: dict, materials_index=None, threshold=None):
    if not isinstance(result, dict):
        return
    key = material_cache_key(item, materials_index, threshold)
    cache = _load(CACHE_FILE)
    payload = dict(result)
    payload["cached_at"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    payload["cache_key"] = key
    cache[key] = payload
    _save(CACHE_FILE, cache)

    # Historial ligero por descripción normalizada; útil para auditoría/aprendizaje.
    hist = _load(HISTORY_FILE)
    n = normalize_material_description(item.get("descripcion"), item.get("unidad"))
    hist_key = (n.get("normalized") or str(item.get("descripcion") or "")).strip().upper() + "|" + normalize_unit(item.get("unidad"))
    hist[hist_key] = {
        "updated_at": payload["cached_at"],
        "source": ((result.get("market_resolution") or {}).get("source")),
        "confidence_pct": ((result.get("market_resolution") or {}).get("confidence_pct")),
        "descripcion_match": (((result.get("market_resolution") or {}).get("evidence") or {}).get("descripcion")),
        "precio": ((result.get("market_resolution") or {}).get("effective_market_price")),
        "cache_key": key,
    }
    _save(HISTORY_FILE, hist)


def cache_status():
    cache = _load(CACHE_FILE)
    hist = _load(HISTORY_FILE)
    return {
        "cache_file": str(CACHE_FILE),
        "history_file": str(HISTORY_FILE),
        "cache_entries": len(cache),
        "history_entries": len(hist),
    }


def clear_material_match_cache():
    _save(CACHE_FILE, {})
    _save(HISTORY_FILE, {})
    return cache_status()
