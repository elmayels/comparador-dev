"""
concept_ai_matcher.py - Homologacion semantica concepto proveedor vs Construdata.

Capa opcional Claude + Voyage:
- Si no hay claves, el sistema sigue con scoring deterministico.
- Si hay VOYAGE_API_KEY, reordena los candidatos top por embedding.
- Si hay ANTHROPIC_API_KEY, Claude decide match directo/compuesto/parcial/sin match con JSON estricto.
"""
import json
import os
import urllib.request
import urllib.error
from typing import List, Dict, Any, Optional
try:
    from material_ai_trace import record_ai_event
except Exception:
    def record_ai_event(*args, **kwargs):
        return None

try:
    from ai_runtime_policy import runtime_ai_allowed
except Exception:
    def runtime_ai_allowed(scope="runtime"):
        return False

ANTHROPIC_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-latest")
VOYAGE_MODEL = os.getenv("VOYAGE_MODEL", "voyage-4-lite")

_CONCEPT_AI_CALLS = {"claude": 0, "voyage": 0}

def _max_concept_ai_calls(kind: str) -> int:
    try:
        if kind == "voyage":
            return int(os.getenv("QUANTIA_AI_MAX_VOYAGE_CONCEPT_CALLS", "12"))
        return int(os.getenv("QUANTIA_AI_MAX_CLAUDE_CONCEPT_CALLS", "6"))
    except Exception:
        return 0

def _concept_ai_budget_available(kind: str) -> bool:
    return _CONCEPT_AI_CALLS.get(kind, 0) < _max_concept_ai_calls(kind)

def _consume_concept_ai_budget(kind: str) -> None:
    _CONCEPT_AI_CALLS[kind] = _CONCEPT_AI_CALLS.get(kind, 0) + 1


def _quantia_ai_enabled():
    # STEP28: IA apagada por defecto para pruebas de performance.
    return (os.getenv("QUANTIA_ENABLE_AI", "0").strip().lower() in {"1", "true", "yes", "on"})

def anthropic_available() -> bool:
    # PASO 33: Claude conceptual permitido solo si la politica IA lo permite
    # y queda presupuesto por reporte/proceso. Nunca se usa por material/insumo.
    return runtime_ai_allowed("concept_matching") and _concept_ai_budget_available("claude") and bool((os.getenv("ANTHROPIC_API_KEY") or "").strip())

def voyage_available() -> bool:
    # PASO 33: Voyage conceptual permitido solo para reranking acotado de conceptos CD.
    return runtime_ai_allowed("concept_matching") and _concept_ai_budget_available("voyage") and bool((os.getenv("VOYAGE_API_KEY") or "").strip())

def get_concept_ai_status() -> dict:
    return {
        "anthropic_enabled": anthropic_available(),
        "voyage_enabled": voyage_available(),
        "anthropic_model": os.getenv("ANTHROPIC_MODEL", ANTHROPIC_MODEL),
        "voyage_model": os.getenv("VOYAGE_MODEL", VOYAGE_MODEL),
        "disabled_by_env": not _quantia_ai_enabled(),
        "claude_calls_used": _CONCEPT_AI_CALLS.get("claude", 0),
        "voyage_calls_used": _CONCEPT_AI_CALLS.get("voyage", 0),
        "claude_calls_max": _max_concept_ai_calls("claude"),
        "voyage_calls_max": _max_concept_ai_calls("voyage"),
        "mode": "claude+voyage" if anthropic_available() and voyage_available() else ("claude" if anthropic_available() else ("voyage" if voyage_available() else "deterministic")),
    }

def _extract_json(text: str) -> Optional[dict]:
    if not text:
        return None
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        return json.loads(text[start:end+1])
    return None

def _anthropic_message(payload: dict, max_tokens: int = 1200, timeout: int = 35) -> Optional[str]:
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        return None
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({
            "model": os.getenv("ANTHROPIC_MODEL", ANTHROPIC_MODEL),
            "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": json.dumps(payload, ensure_ascii=False)}]
        }).encode("utf-8"),
        headers={"content-type":"application/json", "x-api-key":key, "anthropic-version":"2023-06-01"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        _consume_concept_ai_budget("claude")
        record_ai_event("anthropic", "concept_match", ok=True, detail=f"model={os.getenv('ANTHROPIC_MODEL', ANTHROPIC_MODEL)} calls={_CONCEPT_AI_CALLS.get('claude')}")
        return "\n".join(part.get("text", "") for part in raw.get("content", []) if isinstance(part, dict))
    except Exception as exc:
        record_ai_event("anthropic", "concept_match", ok=False, detail=type(exc).__name__)
        raise

def _cosine(a, b) -> float:
    try:
        num = sum(float(x)*float(y) for x,y in zip(a,b))
        da = sum(float(x)*float(x) for x in a) ** 0.5
        db = sum(float(y)*float(y) for y in b) ** 0.5
        return num/(da*db) if da and db else 0.0
    except Exception:
        return 0.0

def _voyage_embeddings(texts: List[str], timeout: int = 45) -> Optional[List[List[float]]]:
    key = os.getenv("VOYAGE_API_KEY")
    if not key or not texts:
        return None
    safe = [str(t or "")[:1800] or "SIN TEXTO" for t in texts]
    req = urllib.request.Request(
        "https://api.voyageai.com/v1/embeddings",
        data=json.dumps({"model": os.getenv("VOYAGE_MODEL", VOYAGE_MODEL), "input": safe}).encode("utf-8"),
        headers={"content-type":"application/json", "Authorization": f"Bearer {key}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        _consume_concept_ai_budget("voyage")
        record_ai_event("voyage", "concept_embeddings", ok=True, detail=f"texts={len(texts)} model={os.getenv('VOYAGE_MODEL', VOYAGE_MODEL)} calls={_CONCEPT_AI_CALLS.get('voyage')}", candidates=len(texts))
    except Exception as exc:
        record_ai_event("voyage", "concept_embeddings", ok=False, detail=type(exc).__name__)
        raise
    data = raw.get("data") or []
    data = sorted(data, key=lambda x: x.get("index", 0))
    return [d.get("embedding") for d in data]

def voyage_rerank_concept_candidates(query_text: str, candidates: List[dict], top_k: int = 12) -> List[dict]:
    """Reordena candidatos ya acotados. Falla silenciosamente hacia el orden deterministico."""
    if not voyage_available() or not candidates:
        return candidates[:top_k]
    try:
        texts = [query_text] + [c.get("embedding_text") or c.get("descripcion") or "" for c in candidates[:30]]
        embs = _voyage_embeddings(texts)
        if not embs or len(embs) < 2:
            return candidates[:top_k]
        q = embs[0]
        ranked = []
        for cand, emb in zip(candidates[:30], embs[1:]):
            c = dict(cand)
            c["voyage_score"] = round(_cosine(q, emb), 4)
            # mezcla conservadora: no deja que embedding supere totalmente a reglas tecnicas
            c["final_retrieval_score"] = round((c.get("score", 0) * 0.65) + (max(c["voyage_score"], 0) * 0.35), 4)
            ranked.append(c)
        return sorted(ranked, key=lambda x: x.get("final_retrieval_score", 0), reverse=True)[:top_k]
    except Exception as exc:
        out = [dict(c) for c in candidates[:top_k]]
        for c in out:
            c["voyage_error"] = type(exc).__name__
        return out

def claude_decide_concept_match(provider: dict, candidates: List[dict]) -> Optional[dict]:
    """Claude decide si el concepto es directo, compuesto, parcial o sin match. Devuelve JSON o None."""
    if not anthropic_available() or not candidates:
        return None
    compact = []
    for c in candidates[:12]:
        compact.append({
            "codigo": c.get("codigo"),
            "descripcion": c.get("descripcion"),
            "unidad": c.get("unidad"),
            "pu": c.get("pu"),
            "score_reglas": c.get("score"),
            "score_voyage": c.get("voyage_score"),
            "tokens_match": c.get("tokens_match"),
            "matriz_resumen": c.get("matriz_resumen"),
        })
    payload = {
        "tarea": "Homologar un concepto de contratista contra conceptos Construdata/Neodata. No inventes codigos; solo usa candidatos.",
        "reglas": [
            "Distingue entre match_directo, match_compuesto, match_parcial y sin_match.",
            "match_compuesto aplica si el concepto del contratista engloba varias actividades que Construdata tiene separadas.",
            "No descartes herramienta, seguridad, acarreos o porcentajes como informacion tecnica; solo no los uses como actividad principal si aparecen en frases 'incluye'.",
            "SUPERVISOR DE SEGURIDAD es mano de obra/recurso humano; no lo homologues como EQUIPO DE SEGURIDAD/EPP ni como porcentaje %MO5.",
            "Si una palabra comun como seguridad coincide, verifica tambien tipo de insumo, rol y composicion de matriz antes de aceptar el match.",
            "Si la unidad es incompatible, baja confianza o marca revision.",
            "Devuelve JSON estricto."
        ],
        "concepto_proveedor": provider,
        "candidatos_construdata": compact,
        "respuesta_json": {
            "tipo_match": "match_directo|match_compuesto|match_parcial|sin_match",
            "confianza": "0 a 1",
            "matches": [{"codigo":"codigo candidato", "peso":"0 a 1", "motivo":"breve"}],
            "requiere_revision": "true/false",
            "explicacion": "breve"
        }
    }
    try:
        parsed = _extract_json(_anthropic_message(payload))
        if not isinstance(parsed, dict):
            return None
        allowed = {c.get("codigo") for c in compact}
        clean_matches = []
        for m in parsed.get("matches") or []:
            if isinstance(m, dict) and m.get("codigo") in allowed:
                try:
                    peso = float(m.get("peso", 1.0))
                except Exception:
                    peso = 1.0
                clean_matches.append({"codigo": m.get("codigo"), "peso": max(0.0, min(1.0, peso)), "motivo": str(m.get("motivo") or "")[:300]})
        parsed["matches"] = clean_matches
        try:
            parsed["confianza"] = max(0.0, min(1.0, float(parsed.get("confianza", 0))))
        except Exception:
            parsed["confianza"] = 0.0
        return parsed
    except Exception as exc:
        return {"error": type(exc).__name__, "tipo_match": "sin_match", "confianza": 0, "matches": [], "requiere_revision": True, "explicacion": "Claude no pudo validar el match"}
