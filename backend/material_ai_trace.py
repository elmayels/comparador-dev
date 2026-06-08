import json
import os
from datetime import datetime
from pathlib import Path

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
TRACE_PATH = DATA_DIR / "material_ai_usage.json"
MAX_EVENTS = int(os.getenv("MATERIAL_AI_TRACE_MAX_EVENTS", "500"))

_DEFAULT = {
    "started_at": None,
    "updated_at": None,
    "counters": {
        "anthropic_calls": 0,
        "anthropic_errors": 0,
        "voyage_calls": 0,
        "voyage_errors": 0,
        "internet_calls": 0,
        "internet_errors": 0,
        "mercadolibre_calls": 0,
        "searxng_calls": 0,
    },
    "last_events": [],
}

def _now():
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

def _default_copy():
    return json.loads(json.dumps(_DEFAULT))

def _read():
    if TRACE_PATH.exists():
        try:
            data = json.loads(TRACE_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                base = _default_copy()
                base.update(data)
                counters = base.setdefault("counters", {})
                counters.update((data.get("counters") or {}))
                base["last_events"] = data.get("last_events") or []
                return base
        except Exception:
            pass
    return _default_copy()

def _write(data):
    try:
        DATA_DIR.mkdir(exist_ok=True)
        TRACE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

def record_ai_event(provider, action, ok=True, material=None, detail=None, candidates=None):
    data = _read()
    now = _now()
    if not data.get("started_at"):
        data["started_at"] = now
    data["updated_at"] = now
    counters = data.setdefault("counters", {})
    provider_key = str(provider or "unknown").lower()
    if provider_key == "anthropic":
        key = "anthropic_calls" if ok else "anthropic_errors"
    elif provider_key == "voyage":
        key = "voyage_calls" if ok else "voyage_errors"
    elif provider_key == "internet":
        key = "internet_calls" if ok else "internet_errors"
    elif provider_key == "mercadolibre":
        key = "mercadolibre_calls" if ok else "internet_errors"
    elif provider_key == "searxng":
        key = "searxng_calls" if ok else "internet_errors"
    else:
        key = provider_key + ("_calls" if ok else "_errors")
    counters[key] = int(counters.get(key) or 0) + 1
    event = {"ts": now, "provider": provider, "action": action, "ok": bool(ok), "material": str(material or "")[:220], "detail": str(detail or "")[:350], "candidates": candidates if candidates is not None else None}
    events = data.setdefault("last_events", [])
    events.append(event)
    data["last_events"] = events[-MAX_EVENTS:]
    _write(data)

def get_ai_usage():
    return _read()

def reset_ai_usage():
    data = _default_copy()
    data["started_at"] = _now()
    data["updated_at"] = data["started_at"]
    _write(data)
    return data
