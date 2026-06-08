"""
Quantia AI Comparador
"""

import os, json, urllib.parse, urllib.request
from material_normalizer import conversion_factor, today_iso, tokenize_material, normalize_unit, normalize_material_description
from material_ai_matcher import normalize_with_ai, explain_match, refine_benchmark_with_voyage, validate_candidates_with_claude, describe_http_exception, ai_available
from material_ai_trace import record_ai_event

MIN_REPLACE_CONFIDENCE = float(os.getenv("MATERIAL_MATCH_THRESHOLD", "60"))

def _request_json(url, headers=None, timeout=20):
    req = urllib.request.Request(url, headers=headers or {"User-Agent": "QuantiaMaterialMarket/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))

def _overlap_score(query, text, unit=None, market_unit=None):
    q = set(tokenize_material(query or ""))
    t = set(tokenize_material(text or ""))
    if not q or not t:
        return 0.0
    overlap = len(q & t) / max(1, len(q))
    bonus = 0.10 if normalize_unit(unit) and normalize_unit(unit) == normalize_unit(market_unit) else 0.0
    return max(0.0, min(0.92, overlap * 0.82 + bonus))

def _parse_custom_endpoint_rows(data):
    rows = data.get("results") if isinstance(data, dict) else data
    out = []
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        price = r.get("precio") or r.get("precio_unitario") or r.get("price")
        if price is None:
            continue
        out.append({
            "fuente": "internet custom endpoint",
            "descripcion": r.get("descripcion") or r.get("title") or r.get("name"),
            "unidad": r.get("unidad") or r.get("unit"),
            "precio": price,
            "url": r.get("url") or r.get("permalink"),
            "fecha_consulta": today_iso(),
            "score": float(r.get("score") or 0),
        })
    return out

def _search_custom_endpoint(item, normalized):
    endpoint = os.getenv("MATERIAL_INTERNET_SEARCH_URL")
    if not endpoint:
        return []
    query = normalized.get("query") or item.get("descripcion", "")
    url = endpoint.rstrip("/") + "?" + urllib.parse.urlencode({"q": query})
    try:
        data = _request_json(url)
        rows = _parse_custom_endpoint_rows(data)
        record_ai_event("internet", "custom_endpoint", ok=True, material=item.get("descripcion"), detail=f"results={len(rows)}")
        return rows
    except Exception as exc:
        record_ai_event("internet", "custom_endpoint", ok=False, material=item.get("descripcion"), detail=json.dumps(describe_http_exception(exc), ensure_ascii=False))
        return []

def _search_mercadolibre_mx(item, normalized):
    query = normalized.get("query") or item.get("descripcion", "")
    if not query:
        return []
    url = "https://api.mercadolibre.com/sites/MLM/search?" + urllib.parse.urlencode({"q": query, "limit": 10})
    try:
        data = _request_json(url, timeout=16)
        rows = []
        for r in data.get("results", [])[:10]:
            price = r.get("price")
            title = r.get("title")
            if price is None or not title:
                continue
            score = _overlap_score(query, title, item.get("unidad"), item.get("unidad"))
            rows.append({
                "fuente": "internet mercadolibre_mx",
                "descripcion": title,
                "unidad": item.get("unidad"),
                "precio": price,
                "url": r.get("permalink"),
                "fecha_consulta": today_iso(),
                "score": score,
                "seller": ((r.get("seller") or {}).get("nickname")),
            })
        record_ai_event("mercadolibre", "search_MLM", ok=True, material=item.get("descripcion"), detail=f"query={query}; results={len(rows)}", candidates=len(rows))
        return rows
    except Exception as exc:
        record_ai_event("mercadolibre", "search_MLM", ok=False, material=item.get("descripcion"), detail=json.dumps(describe_http_exception(exc), ensure_ascii=False))
        return []

def _search_searxng(item, normalized):
    base = os.getenv("SEARXNG_URL")
    if not base:
        return []
    query = (normalized.get("query") or item.get("descripcion", "")) + " precio Mexico material construccion"
    url = base.rstrip("/") + "/search?" + urllib.parse.urlencode({"q": query, "format": "json", "language": "es-MX"})
    try:
        data = _request_json(url, timeout=18)
        rows = []
        for r in data.get("results", [])[:10]:
            title = r.get("title") or ""
            content = r.get("content") or ""
            score = _overlap_score(query, title + " " + content, item.get("unidad"), item.get("unidad"))
            rows.append({
                "fuente": "internet searxng",
                "descripcion": title,
                "unidad": item.get("unidad"),
                "precio": None,
                "url": r.get("url"),
                "fecha_consulta": today_iso(),
                "score": score,
                "snippet": content[:300],
            })
        # Se filtran sin precio porque SearXNG es evidencia, no precio unitario estructurado.
        record_ai_event("searxng", "search", ok=True, material=item.get("descripcion"), detail=f"results={len(rows)}", candidates=len(rows))
        return [r for r in rows if r.get("precio") is not None]
    except Exception as exc:
        record_ai_event("searxng", "search", ok=False, material=item.get("descripcion"), detail=json.dumps(describe_http_exception(exc), ensure_ascii=False))
        return []

def search_internet(item, normalized):
    if os.getenv("ENABLE_INTERNET_PRICE_SEARCH", "0").strip().lower() not in {"1", "true", "yes", "on"}:
        record_ai_event("internet", "skipped", ok=True, material=item.get("descripcion"), detail="ENABLE_INTERNET_PRICE_SEARCH=0")
        return []
    custom = _search_custom_endpoint(item, normalized)
    if custom:
        return custom
    provider = (os.getenv("MATERIAL_INTERNET_SEARCH_PROVIDER") or "mercadolibre_mx").lower().strip()
    if os.getenv("SEARXNG_URL") and provider in {"searxng", "auto"}:
        sx = _search_searxng(item, normalized)
        if sx:
            return sx
    if provider in {"mercadolibre_mx", "mercadolibre", "mlm", "auto", ""}:
        return _search_mercadolibre_mx(item, normalized)
    return []

def _fallback(item, normalized, evidence=None, confidence=0, reason=None):
    return {
        "normalized": normalized,
        "source": "fallback contratista",
        "evidence": evidence or {},
        "market_unit": item.get("unidad"),
        "market_price_raw": None,
        "conversion_factor": 1,
        "conversion_reason": "sin conversion: fallback",
        "effective_market_price": item.get("precio_base"),
        "confidence_pct": round(float(confidence or 0), 2),
        "decision_criterion": reason or "Sin evidencia suficiente; se conserva precio unitario del contratista.",
        "ai_status": normalized.get("ai_status"),
        "voyage_used": False,
        "claude_used": str(normalized.get("ai_status") or "").startswith("ok"),
        "internet_used": bool((evidence or {}).get("fuente", "").startswith("internet")),
        "url": (evidence or {}).get("url"),
        "fecha_consulta": today_iso(),
    }

def _apply_ai_review(item, candidates, normalized, selected=None):
    ai_review = validate_candidates_with_claude(item, candidates)
    if ai_review and ai_review.get("best_index") is not None:
        selected = dict(candidates[int(ai_review["best_index"])])
        selected["score"] = max(float(selected.get("score") or 0), float(ai_review.get("confidence_pct") or 0) / 100)
        selected["ai_reason"] = ai_review.get("reason")
        if ai_review.get("normalized_description"):
            normalized["normalized"] = ai_review.get("normalized_description")
        normalized["ai_candidate_review"] = ai_review
    elif ai_review:
        normalized["ai_candidate_review"] = ai_review
    return selected, ai_review

def select_material_market_price(item, benchmark_match):
    """
    Cascada corregida:
    1) Normalizacion local barata.
    2) Candidatos Construdata preliminares.
    3) Voyage ANTES de Claude para reranking semantico.
    4) Claude decide solo con top candidatos ya ordenados.
    5) Threshold configurable, default 60%.
    """
    normalized = normalize_material_description(item.get("descripcion"), item.get("unidad"))
    normalized["ai_status"] = "local_normalization_before_voyage"

    candidates = list((benchmark_match or {}).get("top_candidates") or [])
    selected = (benchmark_match or {}).get("selected") or (candidates[0] if candidates else None)
    voyage_used = False
    claude_used = False
    ai_review = None

    # Voyage debe correr antes que Claude cuando hay candidatos.
    # Si los candidatos ya vienen del índice vectorial persistente, NO se vuelve
    # a llamar Voyage: ya se pagó 1 embedding del material y se comparó localmente
    # contra todo Construdata.
    if candidates and (candidates[0].get("voyage_score") is not None or candidates[0].get("fuente_precio") == "construdata vector index"):
        voyage_used = True
        selected = selected or candidates[0]
        candidates = sorted(candidates, key=lambda x: float(x.get("score") or 0), reverse=True)[:8]
        normalized["ai_status"] = "voyage_vector_index_ok_before_claude"
    else:
        voyage_selected = refine_benchmark_with_voyage(item, {"top_candidates": candidates}) if candidates else None
        if voyage_selected:
            voyage_used = True
            selected = voyage_selected
            # Mantener al candidato Voyage primero y evitar duplicarlo.
            selected_key = (voyage_selected.get("clave"), voyage_selected.get("descripcion"), voyage_selected.get("unidad"))
            rest = []
            for c in candidates:
                key = (c.get("clave"), c.get("descripcion"), c.get("unidad"))
                if key != selected_key:
                    rest.append(c)
            candidates = [voyage_selected] + rest[:7]
            normalized["ai_status"] = "voyage_rerank_ok_before_claude"
        elif candidates:
            normalized["ai_status"] = "voyage_not_available_or_failed_local_candidates_used"

    # Claude decide despues del reranking de Voyage/local.
    if candidates and ai_available():
        selected, ai_review = _apply_ai_review(item, candidates, normalized, selected)
        claude_used = bool(ai_review and not ai_review.get("error"))
        if claude_used:
            normalized["ai_status"] = "claude_after_voyage_ok" if voyage_used else "claude_after_local_candidates_ok"
        elif ai_review and ai_review.get("error"):
            normalized["ai_status"] = "claude_error_after_candidates"

    if ai_review and ai_review.get("best_index") is None and candidates:
        evidence = candidates[0]
        internet_candidates = search_internet(item, normalized)
        if internet_candidates:
            internet_voyage = refine_benchmark_with_voyage(item, {"top_candidates": internet_candidates})
            if internet_voyage:
                voyage_used = True
                internet_candidates = [internet_voyage] + [c for c in internet_candidates if c is not internet_voyage][:7]
            selected, internet_review = _apply_ai_review(item, internet_candidates, normalized, internet_candidates[0])
            claude_used = claude_used or bool(internet_review and not internet_review.get("error"))
            candidates = internet_candidates
        else:
            return _fallback(item, normalized, evidence, ai_review.get("confidence_pct") or 0, ai_review.get("reason") or "Claude ejecuto validacion, pero no encontro candidato comparable; se conserva precio contratista.")

    confidence = round(float((selected or {}).get("score") or 0) * 100, 2) if selected else 0.0
    evidence = selected
    source = (selected or {}).get("fuente") or (selected or {}).get("fuente_precio") or "construdata excel admin"

    if not selected or confidence < MIN_REPLACE_CONFIDENCE:
        internet_candidates = search_internet(item, normalized)
        if internet_candidates:
            internet_voyage = refine_benchmark_with_voyage(item, {"top_candidates": internet_candidates})
            if internet_voyage:
                voyage_used = True
                internet_candidates = [internet_voyage] + [c for c in internet_candidates if c is not internet_voyage][:7]
            internet_selected = internet_candidates[0]
            internet_selected, internet_review = _apply_ai_review(item, internet_candidates, normalized, internet_selected)
            claude_used = claude_used or bool(internet_review and not internet_review.get("error"))
            internet_conf = round(float((internet_selected or {}).get("score") or 0) * 100, 2) if internet_selected else 0.0
            if internet_selected and internet_conf > confidence:
                selected = internet_selected
                evidence = internet_selected
                confidence = internet_conf
                source = internet_selected.get("fuente") or "internet"

    market_unit = selected.get("unidad") if selected else item.get("unidad")
    factor, factor_reason = conversion_factor(item.get("unidad"), market_unit, (selected or {}).get("descripcion", ""))
    is_internet = bool(str(source).startswith("internet"))

    if selected and confidence >= MIN_REPLACE_CONFIDENCE:
        try:
            effective_price = float(selected.get("precio")) * (factor if factor else 1)
        except Exception:
            return _fallback(item, normalized, selected, confidence, "Precio de fuente no numerico; se conserva fallback contratista.")
        criterion = selected.get("ai_reason") or explain_match(item, selected, confidence, source)
        prefix = []
        if voyage_used: prefix.append("Voyage ejecutado antes de Claude")
        if claude_used: prefix.append("Claude ejecutado para validar candidatos")
        elif candidates and confidence >= MIN_REPLACE_CONFIDENCE: prefix.append("Claude no ejecutado/no disponible; decision por score Voyage/local")
        if is_internet: prefix.append("busqueda internet ejecutada")
        if prefix:
            criterion = "; ".join(prefix) + ". " + criterion
        return {
            "normalized": normalized,
            "source": source,
            "evidence": evidence,
            "market_unit": market_unit,
            "market_price_raw": selected.get("precio"),
            "conversion_factor": factor,
            "conversion_reason": factor_reason,
            "effective_market_price": effective_price,
            "confidence_pct": confidence,
            "decision_criterion": criterion,
            "ai_status": normalized.get("ai_status"),
            "voyage_used": voyage_used,
            "claude_used": claude_used,
            "internet_used": is_internet,
            "url": selected.get("url") or selected.get("clave"),
            "fecha_consulta": selected.get("fecha_consulta") or today_iso(),
        }

    final_reason = explain_match(item, evidence, confidence, source)
    prefix = []
    if voyage_used: prefix.append("Voyage ejecutado antes de Claude")
    else: prefix.append("Voyage no ejecutado: sin candidatos o sin VOYAGE_API_KEY/error")
    if claude_used: prefix.append("Claude ejecutado")
    elif candidates and ai_available(): prefix.append("Claude intento/no valido candidato usable")
    elif not ai_available(): prefix.append("Claude no configurado")
    if is_internet: prefix.append("busqueda internet ejecutada")
    if prefix:
        final_reason = "; ".join(prefix) + ". " + final_reason
    return _fallback(item, normalized, evidence, confidence, final_reason)
def optimize_material_search_scope(conceptos, coverage=0.80):
    rows = []
    for clave, c in (conceptos or {}).items():
        for item in c.get("materiales_items") or []:
            importe = item.get("importe")
            if isinstance(importe, (int, float)) and importe > 0:
                rows.append((importe, clave, item))
    rows.sort(reverse=True, key=lambda x: x[0])
    total = sum(x[0] for x in rows)
    if not rows or total <= 0:
        return None
    acc = 0
    selected = set()
    for importe, clave, item in rows:
        acc += importe
        selected.add((clave, item.get("descripcion"), item.get("unidad"), item.get("precio_base")))
        if acc / total >= coverage:
            break
    return selected
