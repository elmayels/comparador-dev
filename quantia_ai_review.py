"""
Quantia Step31 - IA profesional controlada.

La IA NO participa en el cálculo unitario normal. El cálculo sigue basado en:
- matriz del contratista
- catálogos Construdata granulares
- Resumen PU/cantidades reales

La IA se usa como capa de revisión controlada:
- una llamada Claude por reporte, no por insumo/concepto
- solo sobre casos relevantes/ambiguos
- con cache persistente en DATA_DIR/ai_cache
- sin llamadas Voyage en runtime normal; Voyage queda para índice/admin offline
"""
from __future__ import annotations

import hashlib
import json
import os
import time
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Tuple

_TRUE = {"1", "true", "yes", "on", "si", "sí"}


def _env_true(name: str, default: str = "0") -> bool:
    return str(os.getenv(name, default)).strip().lower() in _TRUE


def _data_dir() -> Path:
    return Path(os.getenv("DATA_DIR") or os.getenv("QUANTIA_DATA_DIR") or "data")


def _trace(event: str, **data: Any) -> None:
    path = os.getenv("QUANTIA_TRACE_FILE")
    payload = {"ts": time.time(), "event": event, **data}
    if path:
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
        except Exception:
            pass


def _cache_dir() -> Path:
    path = _data_dir() / "ai_cache"
    try:
        path.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return path


def _hash_payload(payload: Dict[str, Any]) -> str:
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def ai_runtime_policy() -> Dict[str, Any]:
    mode = (os.getenv("QUANTIA_AI_RUNTIME_MODE") or "off").strip().lower()
    enabled = _env_true("QUANTIA_ENABLE_AI", "0")
    return {
        "enabled": enabled,
        "mode": mode,
        "expert_review_enabled": enabled and mode in {"controlled_review", "runtime_review"} and _env_true("ENABLE_AI_EXPERT_TEXT", "1"),
        "material_runtime_ai": enabled and mode == "runtime_review" and _env_true("ENABLE_AI_MATERIAL_MATCHING", "0"),
        # STEP35: controlled_review SI permite IA conceptual para construir
        # Matriz Recomendada de Mercado. No habilita IA por insumo/material.
        "concept_runtime_ai": enabled and mode in {"controlled_review", "runtime_review"} and _env_true("ENABLE_AI_CONCEPT_MATCHING", "0"),
        "voyage_runtime_ai": enabled and mode in {"controlled_review", "runtime_review"} and _env_true("ENABLE_AI_CONCEPT_MATCHING", "0") and bool((os.getenv("VOYAGE_API_KEY") or "").strip()),
        "anthropic_key_set": bool((os.getenv("ANTHROPIC_API_KEY") or "").strip()),
        "voyage_key_set": bool((os.getenv("VOYAGE_API_KEY") or "").strip()),
        "max_claude_calls_per_report": int(os.getenv("QUANTIA_AI_MAX_CLAUDE_CALLS_PER_REPORT", "1")),
        "max_review_items": int(os.getenv("QUANTIA_AI_MAX_REVIEW_ITEMS", "12")),
    }


def _num(v: Any) -> float | None:
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


def _concept_total(c: Dict[str, Any]) -> float:
    for key in ("importe_resumen_pu", "importe", "total", "importe_contratista"):
        n = _num(c.get(key))
        if n is not None:
            return n
    q = _num(c.get("cantidad")) or 1.0
    pu = _num(c.get("pu")) or 0.0
    return q * pu


def _market_total(c: Dict[str, Any]) -> float | None:
    for key in ("importe_mercado", "market_total", "total_mercado"):
        n = _num(c.get(key))
        if n is not None:
            return n
    q = _num(c.get("cantidad")) or 1.0
    pu = _num(c.get("pu_mercado")) or _num(c.get("market_pu"))
    if pu is not None:
        return q * pu
    direct = _num(c.get("granular_market_direct"))
    if direct is not None:
        return q * direct * 1.25
    return None


def collect_review_context(dato: Dict[str, Any], provider_analysis: Dict[str, Any] | None = None) -> Dict[str, Any]:
    conceptos = dato.get("conceptos") or {}
    rows: List[Tuple[float, Dict[str, Any]]] = []
    material_issues: List[Dict[str, Any]] = []
    matrix_issues: List[Dict[str, Any]] = []

    for clave, c in conceptos.items():
        if not isinstance(c, dict) or str(clave).startswith("__"):
            continue
        prov = _concept_total(c)
        mkt = _market_total(c)
        impact = abs(prov - mkt) if isinstance(mkt, (int, float)) else prov
        rows.append((impact, {"clave": clave, "concepto": c}))

        for item in c.get("granular_market_items") or []:
            if not isinstance(item, dict):
                continue
            status = str(item.get("match_status") or "").lower()
            conf = _num(item.get("confidence")) or 0.0
            price_delta = _num(item.get("delta_precio_pct"))
            imp = _num(item.get("importe_proveedor")) or 0.0
            if "sin_match" in status or conf < 0.45 or (price_delta is not None and abs(price_delta) >= 0.20):
                material_issues.append({
                    "servicio": clave,
                    "insumo": item.get("descripcion"),
                    "tipo": item.get("section") or item.get("domain"),
                    "unidad": item.get("unidad"),
                    "precio_contratista": item.get("precio_proveedor"),
                    "match_status": item.get("match_status"),
                    "candidato_cd": item.get("descripcion_mercado"),
                    "precio_mercado": item.get("precio_mercado"),
                    "confianza": item.get("confidence"),
                    "delta_precio_pct": item.get("delta_precio_pct"),
                    "importe_contratista": imp,
                    "nota": item.get("match_reason"),
                })

        cm = c.get("construdata_concept_match") or {}
        comp = c.get("construdata_matrix_comparison") or {}
        if cm or comp:
            score = _num(cm.get("score"))
            status = str(cm.get("status") or comp.get("match_type") or "")
            if (score is None or score < 0.55 or "revision" in status.lower() or "sin" in status.lower()):
                matrix_issues.append({
                    "servicio": clave,
                    "concepto": c.get("desc"),
                    "status_match_cd": status,
                    "score": score,
                    "codigo_cd": comp.get("construdata_codigo"),
                    "descripcion_cd": comp.get("construdata_desc"),
                    "motivo": cm.get("reason") or comp.get("reason"),
                    "findings": c.get("construdata_matrix_findings") or [],
                })

    rows.sort(key=lambda x: x[0], reverse=True)
    max_items = ai_runtime_policy()["max_review_items"]
    top_concepts = []
    for impact, row in rows[:max_items]:
        c = row["concepto"]
        prov = _concept_total(c)
        mkt = _market_total(c)
        top_concepts.append({
            "servicio": row["clave"],
            "descripcion": c.get("desc"),
            "unidad": c.get("unidad"),
            "cantidad": c.get("cantidad"),
            "pu_contratista": c.get("pu"),
            "importe_contratista": prov,
            "importe_mercado": mkt,
            "impacto_abs": impact,
            "match_cd_status": (c.get("construdata_concept_match") or {}).get("status"),
            "match_cd_reason": (c.get("construdata_concept_match") or {}).get("reason"),
        })

    material_issues.sort(key=lambda r: abs(_num(r.get("importe_contratista")) or 0.0), reverse=True)
    return {
        "proveedor": dato.get("nombre"),
        "resumen": {
            "total_contratista": sum(_concept_total(c) for c in conceptos.values() if isinstance(c, dict)),
            "total_mercado_estimado": sum((_market_total(c) or 0.0) for c in conceptos.values() if isinstance(c, dict)),
            "total_servicios": len([1 for c in conceptos.values() if isinstance(c, dict)]),
        },
        "top_conceptos_por_impacto": top_concepts,
        "materiales_a_revisar": material_issues[:max_items],
        "matches_matriz_cd_a_revisar": matrix_issues[:max_items],
        "drivers_existentes": (provider_analysis or {}).get("drivers", [])[:max_items] if isinstance(provider_analysis, dict) else [],
    }


def _anthropic_message(payload: Dict[str, Any], max_tokens: int = 1400, timeout: int = 35) -> str | None:
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        return None
    model = os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-latest")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({
            "model": model,
            "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": json.dumps(payload, ensure_ascii=False)}],
        }).encode("utf-8"),
        headers={"content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = json.loads(resp.read().decode("utf-8"))
    return "\n".join(part.get("text", "") for part in raw.get("content", []) if isinstance(part, dict))


def generate_controlled_ai_review(dato: Dict[str, Any], provider_analysis: Dict[str, Any] | None = None) -> Dict[str, Any]:
    policy = ai_runtime_policy()
    context = collect_review_context(dato, provider_analysis)
    out: Dict[str, Any] = {"policy": policy, "context": context, "used_ai": False, "source": "deterministic"}

    if not policy.get("expert_review_enabled"):
        out["status"] = "skipped"
        out["reason"] = "IA experta desactivada por politica. Usa QUANTIA_ENABLE_AI=1 y QUANTIA_AI_RUNTIME_MODE=controlled_review para habilitar una revision Claude controlada."
        return out
    if not policy.get("anthropic_key_set"):
        out["status"] = "skipped"
        out["reason"] = "ANTHROPIC_API_KEY no configurada."
        return out

    payload = {
        "rol": "Eres un ingeniero senior de costos, presupuestos y precios unitarios. Evalua como revisor PMD, con criterio conservador y sin inventar datos.",
        "tarea": "Genera un dictamen técnico ejecutivo sobre la cotización. La IA NO decide precios: solo interpreta los resultados calculados por el sistema.",
        "reglas": [
            "No inventes códigos, precios ni matches.",
            "Distingue precio por insumo granular de matriz base Construdata por alcance.",
            "Señala si un match de material o matriz es débil y requiere revisión humana.",
            "Prioriza impacto económico real: cantidad x PU.",
            "Si hay materiales sin match, no afirmes sobreprecio: marca riesgo de referencia faltante.",
            "Devuelve JSON estricto con resumen_ejecutivo, riesgos, recomendaciones y items_criticos.",
        ],
        "datos": context,
        "respuesta_json": {
            "resumen_ejecutivo": ["bullet"],
            "riesgos": [{"severidad": "alta|media|baja", "servicio": "", "comentario": ""}],
            "recomendaciones": ["bullet"],
            "items_criticos": [{"tipo": "material|matriz|alcance|precio", "servicio": "", "descripcion": "", "accion": ""}],
        },
    }
    h = _hash_payload(payload)
    cache_path = _cache_dir() / f"expert_review_{h}.json"
    if cache_path.exists() and not _env_true("QUANTIA_AI_BYPASS_CACHE", "0"):
        try:
            cached = json.loads(cache_path.read_text(encoding="utf-8"))
            cached["from_cache"] = True
            _trace("ai_review_cache_hit", hash=h)
            return cached
        except Exception:
            pass

    _trace("ai_review_start", provider="anthropic", mode=policy.get("mode"), hash=h)
    try:
        text = _anthropic_message(payload, max_tokens=int(os.getenv("QUANTIA_AI_EXPERT_MAX_TOKENS", "1400")), timeout=int(os.getenv("QUANTIA_AI_EXPERT_TIMEOUT", "35")))
        parsed = None
        if text:
            start, end = text.find("{"), text.rfind("}")
            if start >= 0 and end > start:
                parsed = json.loads(text[start:end+1])
        result = {"policy": policy, "context": context, "used_ai": True, "source": "anthropic", "status": "ok", "raw_text": text, "parsed": parsed}
        cache_path.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        _trace("ai_review_end", provider="anthropic", hash=h, ok=True)
        return result
    except Exception as exc:
        _trace("ai_review_error", provider="anthropic", hash=h, error=f"{type(exc).__name__}: {exc}")
        out["status"] = "error"
        out["reason"] = f"{type(exc).__name__}: {str(exc)[:300]}"
        return out
