"""
Quantia AI Comparador
"""

import os, json, urllib.request, urllib.error, ssl, socket
from material_normalizer import normalize_material_description
from material_ai_trace import record_ai_event, get_ai_usage
try:
    from ai_runtime_policy import runtime_ai_allowed
except Exception:
    def runtime_ai_allowed(scope="runtime"):
        return False

ANTHROPIC_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-latest")
VOYAGE_MODEL = os.getenv("VOYAGE_MODEL", "voyage-3-large")


def _quantia_ai_enabled():
    # STEP28: IA apagada por defecto para pruebas de performance en Railway.
    # Activar explicitamente con QUANTIA_ENABLE_AI=1.
    return (os.getenv("QUANTIA_ENABLE_AI", "0").strip().lower() in {"1", "true", "yes", "on"})


def describe_http_exception(exc):
    if isinstance(exc, urllib.error.HTTPError):
        body = ""
        try:
            body = exc.read().decode("utf-8", errors="replace")[:800]
        except Exception:
            body = ""
        return {"type": "HTTPError", "status": getattr(exc, "code", None), "reason": str(getattr(exc, "reason", "")), "body_preview": body, "suggestion": "La red/TLS funcionan; el proveedor rechazo la solicitud. Revisar modelo, permisos, credito o formato de API key."}
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        reason_text = repr(reason)
        suggestion = "No se pudo abrir conexion HTTPS desde el backend. Revisar internet, DNS, proxy/VPN/firewall o certificados TLS."
        if isinstance(reason, ssl.SSLCertVerificationError) or "CERTIFICATE_VERIFY_FAILED" in reason_text:
            suggestion = "Fallo de certificados TLS en Python local. En macOS suele resolverse ejecutando 'Install Certificates.command' del Python instalado, o configurando certificados corporativos/proxy."
        elif isinstance(reason, socket.gaierror) or "nodename nor servname" in reason_text or "Name or service not known" in reason_text:
            suggestion = "Fallo DNS: el backend no puede resolver el dominio del proveedor. Revisar red, VPN, proxy o DNS."
        elif "timed out" in reason_text.lower():
            suggestion = "Timeout de conexion: revisar salida a internet, firewall/VPN/proxy o aumentar timeout."
        return {"type": "URLError", "reason": reason_text, "suggestion": suggestion}
    return {"type": type(exc).__name__, "reason": str(exc), "suggestion": "Error no clasificado; revisar traceback local si persiste."}

def _safe_error_detail(exc):
    try:
        return json.dumps(describe_http_exception(exc), ensure_ascii=False)[:1200]
    except Exception:
        return type(exc).__name__

def ai_available():
    return _quantia_ai_enabled() and bool((os.getenv("ANTHROPIC_API_KEY") or "").strip())

def voyage_available():
    return _quantia_ai_enabled() and bool((os.getenv("VOYAGE_API_KEY") or "").strip())

def get_material_ai_status():
    anthropic_key = os.getenv("ANTHROPIC_API_KEY") or ""
    voyage_key = os.getenv("VOYAGE_API_KEY") or ""
    internet_url = os.getenv("MATERIAL_INTERNET_SEARCH_URL") or ""
    provider = os.getenv("MATERIAL_INTERNET_SEARCH_PROVIDER", "mercadolibre_mx")
    searxng_url = os.getenv("SEARXNG_URL") or ""
    anthropic = bool(anthropic_key.strip())
    voyage = bool(voyage_key.strip())
    internet = bool(internet_url.strip()) or provider.lower() in {"mercadolibre_mx", "mercadolibre", "searxng"} or bool(searxng_url.strip())
    active = bool(anthropic or voyage)
    if anthropic and voyage:
        mode = "claude+voyage"
    elif anthropic:
        mode = "claude"
    elif voyage:
        mode = "voyage_only"
    else:
        mode = "disabled"
    if anthropic and voyage:
        reason = "Claude y Voyage configurados; se pueden normalizar, rankear y validar candidatos."
    elif anthropic:
        reason = "Claude configurado; puede interpretar candidatos de Construdata e internet."
    elif voyage:
        reason = "Solo Voyage configurado; reranking semantico disponible, sin validacion Claude."
    else:
        reason = "IA granular por material desactivada. El modo profesional usa catálogos + índice local y una revisión Claude final controlada."
    return {
        "active": active,
        "enabled": active,
        "ready_for_ai_review": anthropic,
        "anthropic_enabled": anthropic,
        "voyage_enabled": voyage,
        "internet_search_enabled": internet,
        "internet_provider": provider if not internet_url else "custom_endpoint",
        "internet_endpoint": "MATERIAL_INTERNET_SEARCH_URL" if internet_url else ("SEARXNG_URL" if searxng_url else "mercadolibre_mx_default"),
        "mode": mode,
        "reason": reason,
        "env": {
            "QUANTIA_ENABLE_AI": "on" if _quantia_ai_enabled() else "off",
            "ANTHROPIC_API_KEY": "set" if (os.getenv("ANTHROPIC_API_KEY") or "").strip() else "missing",
            "VOYAGE_API_KEY": "set" if voyage else "missing",
            "MATERIAL_INTERNET_SEARCH_URL": "set" if internet_url else "missing",
            "MATERIAL_INTERNET_SEARCH_PROVIDER": provider,
            "SEARXNG_URL": "set" if searxng_url else "missing",
        },
        "anthropic_model": os.getenv("ANTHROPIC_MODEL", ANTHROPIC_MODEL),
        "voyage_model": os.getenv("VOYAGE_MODEL", VOYAGE_MODEL),
        "usage": get_ai_usage(),
    }

def _anthropic_message(payload, max_tokens=500, timeout=25):
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        return None
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({"model": os.getenv("ANTHROPIC_MODEL", ANTHROPIC_MODEL), "max_tokens": max_tokens, "messages": [{"role": "user", "content": json.dumps(payload, ensure_ascii=False)}]}).encode("utf-8"),
        headers={"content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        record_ai_event("anthropic", "messages", ok=True, detail=f"model={os.getenv('ANTHROPIC_MODEL', ANTHROPIC_MODEL)}")
        return "\n".join(part.get("text", "") for part in raw.get("content", []) if isinstance(part, dict))
    except Exception as exc:
        record_ai_event("anthropic", "messages", ok=False, detail=_safe_error_detail(exc))
        raise

def _extract_json_object(text):
    if not text:
        return None
    start, end = text.find("{"), text.rfind("}")
    if start >= 0 and end > start:
        return json.loads(text[start:end+1])
    return None

def normalize_with_ai(item):
    base = normalize_material_description(item.get("descripcion"), item.get("unidad"))
    if not ai_available():
        base["ai_status"] = "disabled:no_ANTHROPIC_API_KEY"
        return base
    prompt = {
        "descripcion": item.get("descripcion"),
        "unidad": item.get("unidad"),
        "instruccion": "Normaliza el material para busqueda de precio. No inventes precio. Devuelve JSON con nombre_normalizado, unidad_comercial, atributos_clave, terminos_busqueda."
    }
    try:
        parsed = _extract_json_object(_anthropic_message(prompt, max_tokens=350))
        record_ai_event("anthropic", "normalize_material", ok=True, material=item.get("descripcion"), detail="normalizacion IA ejecutada")
        if parsed:
            base.update({"ai_status": "ok", "ai": parsed})
            if parsed.get("nombre_normalizado"):
                base["normalized"] = parsed["nombre_normalizado"]
            if parsed.get("terminos_busqueda"):
                base["query"] = " ".join(parsed["terminos_busqueda"][:8]) if isinstance(parsed["terminos_busqueda"], list) else str(parsed["terminos_busqueda"])
        else:
            base["ai_status"] = "ok:no_json"
    except Exception as exc:
        record_ai_event("anthropic", "normalize_material", ok=False, material=item.get("descripcion"), detail=type(exc).__name__)
        base["ai_status"] = "error:" + _safe_error_detail(exc)
    return base

def validate_candidates_with_claude(item, candidates):
    if not ai_available() or not candidates:
        return None
    compact = []
    for idx, c in enumerate(candidates[:8]):
        compact.append({
            "idx": idx,
            "codigo": c.get("clave"),
            "descripcion": c.get("descripcion") or c.get("descripcion_corta"),
            "unidad": c.get("unidad"),
            "precio": c.get("precio"),
            "score_preliminar": c.get("score"),
            "tipo_insumo": c.get("tipo_insumo"),
            "fuente": c.get("fuente") or c.get("fuente_precio"),
            "url": c.get("url"),
        })
    payload = {
        "material_contratista": {"descripcion": item.get("descripcion"), "unidad": item.get("unidad"), "precio_unitario_contratista": item.get("precio_base")},
        "candidatos_catalogo_o_internet": compact,
        "reglas": [
            "No inventes precios.",
            "Escoge solo un candidato de la lista si representa el mismo material comparable.",
            "Si la unidad no es comparable y no hay forma clara de convertir, baja confianza.",
            ">=75 solo si puede usarse como precio mercado comparable.",
            "Si ningun candidato es valido, best_index debe ser null."
        ],
        "respuesta_json": {"best_index": "numero o null", "confidence_pct": "numero 0-100", "normalized_description": "descripcion normalizada", "unit_notes": "comentario de unidad", "reason": "criterio breve"}
    }
    try:
        parsed = _extract_json_object(_anthropic_message(payload, max_tokens=700))
        record_ai_event("anthropic", "validate_candidates", ok=True, material=item.get("descripcion"), detail=f"candidates={len(compact)}", candidates=len(compact))
        if not isinstance(parsed, dict):
            return None
        idx = parsed.get("best_index")
        if idx is not None:
            idx = int(idx)
            if idx < 0 or idx >= len(candidates[:8]):
                idx = None
        conf = parsed.get("confidence_pct")
        try:
            conf = max(0.0, min(100.0, float(conf)))
        except Exception:
            conf = 0.0
        return {"best_index": idx, "confidence_pct": conf, "normalized_description": parsed.get("normalized_description"), "unit_notes": parsed.get("unit_notes"), "reason": parsed.get("reason")}
    except Exception as exc:
        record_ai_event("anthropic", "validate_candidates", ok=False, material=item.get("descripcion"), detail=type(exc).__name__)
        return {"error": _safe_error_detail(exc)}

def explain_match(item, candidate, confidence, source):
    if not candidate:
        return "Sin evidencia suficiente; se conserva precio del contratista como fallback."
    if confidence >= 90:
        level = "match directo/casi exacto"
    elif confidence >= 75:
        level = "match aceptable con normalizacion menor"
    elif confidence >= 60:
        level = "match debil, solo referencia auxiliar"
    else:
        level = "match insuficiente para reemplazo automatico"
    desc = candidate.get("descripcion") or candidate.get("descripcion_encontrada") or ""
    return f"{level}. Fuente: {source}. Comparado contra '{desc}'."

def _cosine(a, b):
    try:
        num = sum(float(x) * float(y) for x, y in zip(a, b))
        da = sum(float(x) * float(x) for x in a) ** 0.5
        db = sum(float(y) * float(y) for y in b) ** 0.5
        return num / (da * db) if da and db else 0.0
    except Exception:
        return 0.0

def _voyage_embeddings(texts):
    key = os.getenv("VOYAGE_API_KEY")
    if not key or not texts:
        return None
    req = urllib.request.Request(
        "https://api.voyageai.com/v1/embeddings",
        data=json.dumps({"model": os.getenv("VOYAGE_MODEL", VOYAGE_MODEL), "input": texts}).encode("utf-8"),
        headers={"content-type": "application/json", "Authorization": f"Bearer {key}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        data = raw.get("data") or []
        record_ai_event("voyage", "embeddings", ok=True, detail=f"texts={len(texts)} model={os.getenv('VOYAGE_MODEL', VOYAGE_MODEL)}", candidates=len(texts))
        return [x.get("embedding") for x in data if x.get("embedding")]
    except Exception as exc:
        record_ai_event("voyage", "embeddings", ok=False, detail=_safe_error_detail(exc))
        return None

def refine_benchmark_with_voyage(item, benchmark_match):
    candidates = (benchmark_match or {}).get("top_candidates") or []
    if not candidates:
        return None
    query = f"{item.get('descripcion') or ''} unidad {item.get('unidad') or ''}".strip()
    texts = [query] + [f"{c.get('descripcion') or ''} unidad {c.get('unidad') or ''}".strip() for c in candidates[:8]]
    embeddings = _voyage_embeddings(texts)
    if not embeddings or len(embeddings) < 2:
        return None
    q = embeddings[0]
    scored = []
    for cand, emb in zip(candidates[:8], embeddings[1:]):
        sim = _cosine(q, emb)
        merged = dict(cand)
        merged["voyage_score"] = round(sim, 4)
        merged["score"] = max(float(merged.get("score") or 0), sim)
        scored.append(merged)
    scored.sort(key=lambda x: x.get("voyage_score") or 0, reverse=True)
    return scored[0] if scored else None



def run_network_diagnostics():
    targets = [
        ("anthropic", "https://api.anthropic.com/v1/messages"),
        ("voyage", "https://api.voyageai.com/v1/embeddings"),
        ("mercadolibre_mx", "https://api.mercadolibre.com/sites/MLM/search?q=cemento&limit=1"),
    ]
    out = []
    for name, url in targets:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "QuantiaDiagnostics/1.0"}, method="GET")
            with urllib.request.urlopen(req, timeout=12) as resp:
                out.append({"target": name, "ok": True, "status": getattr(resp, "status", None), "network_ok": True})
        except urllib.error.HTTPError as exc:
            out.append({"target": name, "ok": True, "http_status": getattr(exc, "code", None), "network_ok": True, "note": "HTTPError esperado en endpoint protegido; DNS/TLS/red OK."})
        except Exception as exc:
            out.append({"target": name, "ok": False, "network_ok": False, "error": describe_http_exception(exc)})
    return {"diagnostics": out}

def run_material_ai_smoke_test():
    """Ejecuta llamadas minimas reales para comprobar que las APIs responden y que el contador sube."""
    result = {"anthropic": None, "voyage": None}
    if ai_available():
        try:
            text = _anthropic_message({"test": "Responde solo JSON", "respuesta_json": {"ok": True}}, max_tokens=80, timeout=15)
            result["anthropic"] = {"ok": True, "response_preview": str(text or "")[:120]}
        except Exception as exc:
            result["anthropic"] = {"ok": False, "error": describe_http_exception(exc)}
    else:
        result["anthropic"] = {"ok": False, "error": "ANTHROPIC_API_KEY missing"}
    if voyage_available():
        emb = _voyage_embeddings(["cemento gris saco 50 kg", "cemento portland 50 kg"])
        result["voyage"] = {"ok": bool(emb and len(emb) == 2), "vectors": len(emb or [])}
        if not result["voyage"]["ok"]:
            result["voyage"]["error"] = "Ver usage.last_events para detalle exacto; la llamada fallo antes de devolver embeddings."
    else:
        result["voyage"] = {"ok": False, "error": "VOYAGE_API_KEY missing"}
    result["network_diagnostics"] = run_network_diagnostics()
    result["usage"] = get_ai_usage()
    return result
