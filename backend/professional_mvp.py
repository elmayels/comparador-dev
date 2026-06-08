"""Capa profesional para comparativos APU tipo Neodata/Construdata.

Capa liviana: no usa Redis/DB. Lee los mismos Excel del motor actual y agrega
hojas ejecutivas al workbook ya generado: resumen profesional, comparativa, detalle por proveedor y riesgos de alcance.
"""
from __future__ import annotations

import datetime as dt
import math
import os
import re
import tempfile
import unicodedata
from copy import copy
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from .processor import extract_conceptos, extract_precios_nacional, extract_resumen_pu, apply_resumen_pu_to_conceptos, build_comparativo
    try:
        from .processor import quantia_market_catalog_diagnostics
    except Exception:
        quantia_market_catalog_diagnostics = None
except Exception:
    from processor import extract_conceptos, extract_precios_nacional, extract_resumen_pu, apply_resumen_pu_to_conceptos, build_comparativo
    try:
        from processor import quantia_market_catalog_diagnostics
    except Exception:
        quantia_market_catalog_diagnostics = None

NAVY = "17365D"; BLUE = "1F4E79"; SKY = "D9EAF7"; GREEN = "D9EAD3"; YELLOW = "FFF2CC"; RED = "F4CCCC"; GRAY = "E7E6E6"; WHITE = "FFFFFF"; TEXT = "1F2937"


def env_float(name: str, default: float) -> float:
    try:
        return float(str(os.getenv(name, default)).replace(",", "."))
    except Exception:
        return default


def env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return str(raw).strip().lower() in {"1", "true", "yes", "si", "sí", "on"}


def settings() -> Dict[str, Any]:
    return {
        "max_providers": int(env_float("MAX_PROVIDERS", 5)),
        "market_indirect_pct": env_float("MARKET_INDIRECT_PCT", 0.25),
        "strict_unit_matching": env_bool("STRICT_UNIT_MATCHING", True),
        "weights": {"precio": 0.35, "cobertura": 0.20, "mercado": 0.20, "calidad": 0.15, "riesgo": 0.10},
    }


def _data_dir_candidates() -> List[Path]:
    candidates: List[Path] = []
    for key in ("DATA_DIR", "QUANTIA_DATA_DIR"):
        raw = os.getenv(key)
        if raw:
            candidates.append(Path(raw))
    try:
        candidates.append(Path(__file__).resolve().parents[1] / "data")
    except Exception:
        pass
    candidates.append(Path("data"))
    # Deduplicar preservando orden.
    out: List[Path] = []
    seen = set()
    for c in candidates:
        key = str(c.resolve()) if c.exists() else str(c)
        if key not in seen:
            out.append(c); seen.add(key)
    return out


def granular_reference_available() -> bool:
    """Detecta catálogos granulares embebidos usados por el motor histórico.

    La capa profesional puede no recibir un archivo maestro Neodata/construdata_matrices.xlsx,
    pero el motor sí puede estar calculando mercado por catálogos de materiales,
    mano de obra y maquinaria. Esta señal evita mostrar "no disponible" cuando
    existe referencia granular en el paquete.
    """
    patterns = (
        "construdata-materiales*.xlsx",
        "construdata-manodeobra*.xlsx",
        "construdata-maquinaria*.xlsx",
    )
    for d in _data_dir_candidates():
        try:
            if d.exists() and any(any(d.glob(pat)) for pat in patterns):
                return True
        except Exception:
            continue
    return False


def reference_label(neodata_meta: Dict[str, Any]) -> str:
    if neodata_meta.get("available"):
        return "Maestro Neodata/Construdata disponible"
    if granular_reference_available():
        return "Catálogos granulares Construdata disponibles (ver Detalle)"
    return "No disponible / no leída"


def as_float(v: Any) -> Optional[float]:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return None if math.isnan(f) or math.isinf(f) else f
    s = str(v).strip().replace("$", "").replace("%", "").replace(" ", "")
    if not s:
        return None
    if s.count(",") == 1 and s.count(".") >= 1 and s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))


def norm_text(v: Any) -> str:
    s = strip_accents(str(v or "").lower()).replace("_x000d_", " ")
    s = re.sub(r"[^a-z0-9%/\.\-\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    stop = {"de","del","la","las","los","el","en","y","o","a","para","por","con","sin","incluye","incl","suministro","colocacion","instalacion","servicio","obra","trabajo","precio","unitario","segun","planos","especificaciones"}
    return " ".join(t for t in s.split() if t not in stop)


def norm_unit(v: Any) -> str:
    s = strip_accents(str(v or "").lower()).strip().replace("²", "2").replace("³", "3").replace(".", "").replace(" ", "")
    aliases = {"ml":"m","metro":"m","metros":"m","mt":"m","m":"m","m2":"m2","mt2":"m2","metro2":"m2","metros2":"m2","m3":"m3","mt3":"m3","metro3":"m3","metros3":"m3","pz":"pza","pza":"pza","pzas":"pza","pieza":"pza","piezas":"pza","kg":"kg","kgs":"kg","kilo":"kg","ton":"ton","tons":"ton","tonelada":"ton","lt":"lt","lts":"lt","l":"lt","hr":"hr","hrs":"hr","hora":"hr","horas":"hr","jor":"jor","jornada":"jor","jornal":"jor","dia":"jor","dias":"jor","viaje":"viaje","viajes":"viaje","lote":"lote","lotes":"lote","gl":"global","global":"global","serv":"serv","servicio":"serv","%":"%","porc":"%","porcentaje":"%"}
    return aliases.get(s, s)


def unit_ok(a: Any, b: Any, strict: bool = True) -> Tuple[bool, str]:
    ua, ub = norm_unit(a), norm_unit(b)
    if not ua or not ub:
        return True, "unidad_no_informada"
    if ua == ub:
        return True, "unidad_equivalente"
    if not strict and ({ua, ub} <= {"m", "ml"} or {ua, ub} <= {"hr", "hora"} or {ua, ub} <= {"jor", "dia"}):
        return True, "unidad_equivalente_no_estricta"
    return False, f"unidad_incompatible:{ua}!={ub}"


def sim(a: Any, b: Any) -> float:
    na, nb = norm_text(a), norm_text(b)
    if not na or not nb:
        return 0.0
    ta, tb = set(na.split()), set(nb.split())
    jac = len(ta & tb) / max(1, len(ta | tb))
    seq = SequenceMatcher(None, na[:500], nb[:500]).ratio() * 0.86
    return max(jac, seq)


def desc(raw: Dict[str, Any]) -> str:
    return str(raw.get("desc") or raw.get("descripcion") or raw.get("descripcion_resumen") or raw.get("concepto") or "").strip()


def code(key: Any, raw: Dict[str, Any]) -> str:
    for k in ("codigo_resumen", "codigo", "clave", "code"):
        if raw.get(k) not in (None, ""):
            return str(raw.get(k)).strip()
    return str(key or "").split("::")[-1].strip()


def unit(raw: Dict[str, Any]) -> str:
    return str(raw.get("unidad") or raw.get("unidad_resumen") or raw.get("unit") or "").strip()


def qty(raw: Dict[str, Any]) -> Optional[float]:
    for k in ("cantidad_real_cotizacion", "cantidad_resumen", "cantidad", "qty"):
        f = as_float(raw.get(k))
        if f is not None:
            return f
    return None


def pu(raw: Dict[str, Any]) -> Optional[float]:
    for k in ("pu", "pu_resumen", "precio_unitario", "precio", "pu_mercado", "precio_nacional"):
        f = as_float(raw.get(k))
        if f is not None:
            return f
    return None


def total(raw: Dict[str, Any], q: Optional[float], p: Optional[float]) -> Optional[float]:
    for k in ("total", "importe", "importe_resumen", "total_estimado"):
        f = as_float(raw.get(k))
        if f is not None:
            return f
    return q * p if q is not None and p is not None else None


def read_provider(filepath: str, resumen_path: Optional[str], provider: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    diag = {"provider": provider, "file": Path(filepath).name, "resumen_applied": False, "errors": []}
    conceptos = extract_conceptos(filepath)
    if resumen_path:
        try:
            rows = extract_resumen_pu(resumen_path)
            apply_resumen_pu_to_conceptos(conceptos, rows)
            diag["resumen_rows"] = len(rows); diag["resumen_applied"] = True
        except Exception as exc:
            diag["errors"].append(f"Resumen PU no aplicado: {type(exc).__name__}: {exc}")
    out = []
    for i, (k, raw) in enumerate((conceptos or {}).items(), 1):
        if str(k).startswith("__") or not isinstance(raw, dict):
            continue
        d = desc(raw)
        if not d:
            continue
        q, p = qty(raw), pu(raw)
        out.append({"order": i, "key": str(k), "code": code(k, raw), "desc": d, "unit": unit(raw), "unit_norm": norm_unit(unit(raw)), "qty": q, "pu": p, "total": total(raw, q, p), "raw": raw})
    diag["concepts"] = len(out); diag["total_bid"] = round(sum(x.get("total") or 0 for x in out), 2)
    return out, diag


def base_from_resumen(path: str) -> List[Dict[str, Any]]:
    rows = []
    for i, r in enumerate(extract_resumen_pu(path), 1):
        d = str(r.get("descripcion_resumen") or "").strip()
        if not d:
            continue
        q = as_float(r.get("cantidad_resumen")); p = as_float(r.get("pu_resumen")); t = as_float(r.get("importe_resumen"))
        rows.append({"base_id": f"BASE-{i:04d}", "order": i, "code": str(r.get("codigo_resumen") or "").strip(), "desc": d, "unit": str(r.get("unidad_resumen") or "").strip(), "unit_norm": norm_unit(r.get("unidad_resumen")), "qty": q, "pu_input": p, "total_input": t, "source": "resumen_oficial"})
    return rows


def base_from_provider(rows: List[Dict[str, Any]], source: str) -> List[Dict[str, Any]]:
    return [{"base_id": f"BASE-{i:04d}", "order": i, "code": r.get("code"), "desc": r.get("desc"), "unit": r.get("unit"), "unit_norm": r.get("unit_norm"), "qty": r.get("qty"), "pu_input": r.get("pu"), "total_input": r.get("total"), "source": source} for i, r in enumerate(rows, 1)]


def load_neodata(nacional_path: Optional[str]) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Any]], Dict[str, Any]]:
    meta = {"available": False, "path": nacional_path, "count": 0, "error": None}
    if not nacional_path or not Path(str(nacional_path)).exists():
        return [], {}, meta
    try:
        refs = extract_precios_nacional(str(nacional_path))
    except Exception as exc:
        meta["error"] = f"{type(exc).__name__}: {exc}"; return [], {}, meta
    rows, by_code = [], {}
    for i, (k, raw) in enumerate((refs or {}).items(), 1):
        if str(k).startswith("__") or not isinstance(raw, dict):
            continue
        d = desc(raw)
        if not d:
            continue
        q, p = qty(raw), pu(raw)
        row = {"order": i, "code": code(k, raw), "desc": d, "unit": unit(raw), "unit_norm": norm_unit(unit(raw)), "qty": q, "pu": p, "total": total(raw, q, p)}
        rows.append(row)
        if row["code"]:
            by_code[norm_text(row["code"])] = row
    meta.update({"available": True, "count": len(rows), "kind": (refs or {}).get("__benchmark_kind__")})
    return rows, by_code, meta


def attach_neodata(base_rows: List[Dict[str, Any]], nacional_path: Optional[str], strict: bool) -> Dict[str, Any]:
    refs, by_code, meta = load_neodata(nacional_path)
    if not refs:
        return meta
    indexed = [(set(norm_text(r["desc"]).split()), r) for r in refs]
    for b in base_rows:
        selected, score = None, 0.0
        c = norm_text(b.get("code"))
        if c and c in by_code:
            selected, score = by_code[c], 1.0
        else:
            bt = set(norm_text(b.get("desc")).split())
            candidates = []
            for toks, r in indexed:
                if bt & toks and unit_ok(b.get("unit"), r.get("unit"), strict)[0]:
                    candidates.append((len(bt & toks), r))
            candidates.sort(key=lambda x: x[0], reverse=True)
            for _, r in candidates[:100]:
                s = sim(b.get("desc"), r.get("desc"))
                if s > score:
                    selected, score = r, s
        if selected and score >= 0.52:
            b["ref_pu"] = selected.get("pu")
            b["ref_total"] = (b.get("qty") * selected.get("pu")) if b.get("qty") is not None and selected.get("pu") is not None else selected.get("total")
            b["neodata_code"] = selected.get("code")
            b["neodata_desc"] = selected.get("desc")
            b["neodata_match_score"] = score
        else:
            b["ref_pu"] = b.get("pu_input")
            b["ref_total"] = b.get("total_input")
    return meta


def match_base(base: Dict[str, Any], rows: List[Dict[str, Any]], used: set, idx: int, strict: bool) -> Dict[str, Any]:
    c = norm_text(base.get("code"))
    if c:
        for r in rows:
            if r["order"] in used:
                continue
            if norm_text(r.get("code")) == c:
                ok, reason = unit_ok(base.get("unit"), r.get("unit"), strict)
                return {"row": r, "status": "cotizado" if ok else "unidad_incompatible", "similarity": 1.0, "method": "codigo", "unit_ok": ok, "reason": reason}
    if idx < len(rows) and rows[idx]["order"] not in used:
        r = rows[idx]
        s = sim(base.get("desc"), r.get("desc")); ok, reason = unit_ok(base.get("unit"), r.get("unit"), strict)
        if ok and s >= 0.35:
            return {"row": r, "status": "cotizado", "similarity": s, "method": "orden", "unit_ok": True, "reason": reason}
        if not ok and s >= 0.55:
            return {"row": r, "status": "unidad_incompatible", "similarity": s, "method": "orden", "unit_ok": False, "reason": reason}
    best = None; best_score = 0.0; best_ok = True; best_reason = ""
    for r in rows:
        if r["order"] in used:
            continue
        ok, reason = unit_ok(base.get("unit"), r.get("unit"), strict)
        s = sim(base.get("desc"), r.get("desc")) + (0.10 if ok else -0.12)
        if s > best_score:
            best, best_score, best_ok, best_reason = r, s, ok, reason
    if best and best_score >= 0.50:
        return {"row": best, "status": "cotizado" if best_ok else "unidad_incompatible", "similarity": min(1.0, best_score), "method": "descripcion_unidad", "unit_ok": best_ok, "reason": best_reason}
    return {"row": None, "status": "no_cotizado", "similarity": 0.0, "method": "sin_match", "unit_ok": True, "reason": "sin_match"}


def build_professional_mvp_analysis(filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> Dict[str, Any]:
    cfg = settings(); strict = bool(cfg["strict_unit_matching"]); meta = meta or {}; resumen_paths = resumen_paths or []
    provider_names = [(nombres[i] if i < len(nombres) and nombres[i] else f"Proveedor {i+1}") for i in range(len(filepaths or []))]
    prov_rows, diagnostics = [], []
    for i, fp in enumerate(filepaths or []):
        rp = resumen_paths[i] if i < len(resumen_paths) else None
        rows, diag = read_provider(fp, rp, provider_names[i])
        prov_rows.append(rows); diagnostics.append(diag)
    base_rows, base_source = [], ""
    for candidate, source in ((resumen_oficial_path, "resumen_oficial"), (catalogo_path, "catalogo_resumen_pu"), (next((p for p in resumen_paths if p), None), "resumen_pu_primer_proveedor")):
        if candidate:
            try:
                base_rows = base_from_resumen(candidate)
                if base_rows:
                    base_source = source; break
            except Exception:
                pass
    if not base_rows and prov_rows:
        base_rows = base_from_provider(prov_rows[0], "primer_proveedor"); base_source = "primer_proveedor"
    if not base_rows:
        base_source = "sin_base"
    neodata = attach_neodata(base_rows, nacional_path, strict)
    used = [set() for _ in prov_rows]
    matrix = []
    for b_idx, b in enumerate(base_rows):
        row = {"base_id": b.get("base_id"), "code": b.get("code"), "desc": b.get("desc"), "unit": b.get("unit"), "qty": b.get("qty"), "source": b.get("source"), "ref_pu": b.get("ref_pu", b.get("pu_input")), "ref_total": b.get("ref_total", b.get("total_input")), "neodata_code": b.get("neodata_code"), "neodata_desc": b.get("neodata_desc"), "neodata_match_score": b.get("neodata_match_score"), "providers": {}, "best_provider": "", "best_total": None, "risk": ""}
        risks = []
        for p_idx, rows in enumerate(prov_rows):
            m = match_base(b, rows, used[p_idx], b_idx, strict)
            r = m["row"]
            if r:
                used[p_idx].add(r["order"])
            q = b.get("qty") if b.get("qty") is not None else (r.get("qty") if r else None)
            p = r.get("pu") if r else None
            t = r.get("total") if r and r.get("total") is not None else (q * p if q is not None and p is not None else None)
            ref = row.get("ref_pu")
            delta = (p - ref) / ref if p is not None and ref not in (None, 0) else None
            if m["status"] != "cotizado": risks.append(f"{provider_names[p_idx]}: {m['status']}")
            elif isinstance(delta, (int, float)) and abs(delta) >= 0.35: risks.append(f"{provider_names[p_idx]}: desviación >35%")
            row["providers"][provider_names[p_idx]] = {"status": m["status"], "code": r.get("code") if r else "", "desc": r.get("desc") if r else "", "unit": r.get("unit") if r else "", "qty": q, "pu": p, "total": t, "delta_ref_pct": delta, "similarity": m["similarity"], "method": m["method"], "unit_ok": m["unit_ok"]}
            if m["status"] == "cotizado" and t is not None and (row["best_total"] is None or t < row["best_total"]):
                row["best_total"] = t; row["best_provider"] = provider_names[p_idx]
        row["risk"] = "; ".join(risks[:4])
        matrix.append(row)
    extras = {provider_names[i]: [r for r in rows if r["order"] not in used[i]] for i, rows in enumerate(prov_rows)}
    scores = score_providers(provider_names, matrix, extras, prov_rows, cfg)
    findings = executive_findings(matrix, scores, extras)
    return {"version": "quantia-comparador-apu-v0.2.4-analistas", "generated_at": dt.datetime.utcnow().isoformat()+"Z", "meta": meta, "settings": cfg, "base_source": base_source, "base_count": len(base_rows), "providers": provider_names, "diagnostics": diagnostics, "neodata": neodata, "matrix_rows": matrix, "extra_rows": extras, "provider_scores": scores, "executive_findings": findings}


def score_providers(names: List[str], matrix: List[Dict[str, Any]], extras: Dict[str, List[Dict[str, Any]]], prov_rows: List[List[Dict[str, Any]]], cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    totals = {}
    for name in names:
        totals[name] = sum((row["providers"].get(name, {}).get("total") or 0) for row in matrix if row["providers"].get(name, {}).get("status") == "cotizado")
    min_total = min([v for v in totals.values() if v > 0], default=None)
    out = []
    n = len(matrix)
    for i, name in enumerate(names):
        quoted = missing = mismatches = anomalies = 0; deltas = []; ref_total = 0.0
        for row in matrix:
            pr = row["providers"].get(name, {})
            st = pr.get("status")
            if st == "cotizado":
                quoted += 1
                if isinstance(row.get("ref_total"), (int, float)): ref_total += row["ref_total"]
                d = pr.get("delta_ref_pct")
                if isinstance(d, (int, float)):
                    deltas.append(abs(d))
                    if abs(d) >= 0.35: anomalies += 1
            elif st == "unidad_incompatible":
                mismatches += 1; anomalies += 1
            else:
                missing += 1
        coverage = quoted / n if n else 0
        avg = sum(deltas)/len(deltas) if deltas else None
        price_score = max(0, min(100, 100*min_total/totals[name])) if min_total and totals[name] > 0 else 0
        coverage_score = coverage*100
        market_score = 75 if avg is None else max(0, 100-min(1, avg)*100)
        quality_score = max(0, 100-mismatches*18-len(extras.get(name, []))*2)
        risk_score = max(0, 100-missing*12-anomalies*10)
        w = cfg["weights"]
        total_score = price_score*w["precio"] + coverage_score*w["cobertura"] + market_score*w["mercado"] + quality_score*w["calidad"] + risk_score*w["riesgo"]
        rec = recommendation(total_score, coverage, avg, missing, mismatches, totals[name], min_total, provider_count=len(names))
        out.append({"provider": name, "rank": 0, "total_score": round(total_score, 2), "total_comparable": round(totals[name], 2), "total_bid": round(sum(r.get("total") or 0 for r in prov_rows[i]), 2), "total_reference": round(ref_total, 2), "concepts_base": n, "quoted": quoted, "missing": missing, "unit_mismatches": mismatches, "extras": len(extras.get(name, [])), "anomalies": anomalies, "coverage_pct": coverage, "avg_abs_market_delta_pct": avg, "price_score": round(price_score,2), "coverage_score": round(coverage_score,2), "market_score": round(market_score,2), "quality_score": round(quality_score,2), "risk_score": round(risk_score,2), "recommendation": rec})
    out.sort(key=lambda x: x["total_score"], reverse=True)
    for i, row in enumerate(out, 1): row["rank"] = i
    return out


def recommendation(score: float, coverage: float, avg: Optional[float], missing: int, mismatches: int, total: float, min_total: Optional[float], provider_count: int = 1) -> str:
    if provider_count <= 1:
        if missing and coverage < 0.85:
            return "Proveedor único incompleto: revisar alcance antes de negociar."
        if mismatches:
            return "Proveedor único: requiere homologación de unidades antes de dictaminar."
        if avg is not None and avg > 0.35:
            return "Proveedor único: desviación alta contra base; auditar matrices y rendimientos."
        return "Proveedor único: evaluación técnica informativa; no hay ranking competitivo."
    if missing and coverage < 0.85: return "Cotización incompleta: revisar alcance antes de negociar."
    if mismatches: return "Requiere homologación de unidades antes de dictaminar."
    if min_total and total == min_total and score >= 75: return "Mejor candidato económico-técnico para negociación."
    if avg is not None and avg > 0.35: return "Desviación alta contra base; auditar matrices y rendimientos."
    if score >= 80: return "Oferta sólida; revisar hallazgos de alto impacto."
    if score >= 65: return "Oferta viable con condiciones; requiere ajuste técnico/comercial."
    return "Alto riesgo relativo; no recomendar sin aclaraciones."


def executive_findings(matrix: List[Dict[str, Any]], scores: List[Dict[str, Any]], extras: Dict[str, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    out = []
    if scores:
        if len(scores) <= 1:
            out.append({"tipo":"evaluacion","severidad":"info","mensaje":f"Proveedor único evaluado: {scores[0]['provider']}. El score {scores[0]['total_score']}/100 es informativo, no ranking competitivo."})
        else:
            out.append({"tipo":"ranking","severidad":"info","mensaje":f"Proveedor mejor evaluado: {scores[0]['provider']} con score {scores[0]['total_score']}/100."})
    miss = sum(1 for row in matrix for pr in row["providers"].values() if pr.get("status") == "no_cotizado")
    unit = sum(1 for row in matrix for pr in row["providers"].values() if pr.get("status") == "unidad_incompatible")
    extra = sum(len(v) for v in extras.values())
    if miss: out.append({"tipo":"alcance","severidad":"alta","mensaje":f"Hay {miss} partidas no cotizadas contra la base común."})
    if unit: out.append({"tipo":"unidad","severidad":"alta","mensaje":f"Hay {unit} coincidencias con unidad incompatible; no comparar directo."})
    if extra: out.append({"tipo":"adicionales","severidad":"media","mensaje":f"Se detectaron {extra} conceptos adicionales fuera de base; no se suman al total comparable."})
    deltas = []
    for row in matrix:
        for p, pr in row["providers"].items():
            d = pr.get("delta_ref_pct")
            if isinstance(d, (int, float)) and abs(d) >= 0.35:
                deltas.append((abs(d), p, row.get("desc"), d))
    for _, p, dsc, d in sorted(deltas, reverse=True)[:8]:
        out.append({"tipo":"precio","severidad":"media","mensaje":f"{p}: {str(dsc)[:90]} presenta desviación de {d:.1%} contra base."})
    return out[:12]


def style_title(cell):
    cell.fill = PatternFill("solid", fgColor=NAVY); cell.font = Font(bold=True, color=WHITE, size=14); cell.alignment = Alignment(horizontal="left", vertical="center")


def style_header(cell):
    cell.fill = PatternFill("solid", fgColor=BLUE); cell.font = Font(bold=True, color=WHITE, size=9); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); cell.border = Border(bottom=Side(style="thin", color="B7B7B7"))


def style_cell(cell, fill: Optional[str] = None, bold: bool = False, wrap: bool = False):
    if fill: cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(color=TEXT, size=9, bold=bold); cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap); cell.border = Border(bottom=Side(style="hair", color="D9D9D9"))


def fmt_money(cell): cell.number_format = '$#,##0.00'
def fmt_pct(cell): cell.number_format = '0.00%'
def fmt_num(cell): cell.number_format = '#,##0.0000'


ANALYST_SUMMARY_SHEET = "Resumen Profesional"
ANALYST_TRACE_SHEET = "Trazabilidad Técnica"
DEFAULT_ANALYST_INTERNAL_SHEETS = [
    "Alcance y Riesgos",
    "Diagnóstico Catálogos",
    "Diagnostico Catalogos",
    "Trazabilidad Técnica",
]


def analyst_hide_internal_sheets() -> bool:
    """Controla si las hojas técnicas se ocultan para entrega a analistas.

    Default: ocultas. Para auditoría/desarrollo usar EXCEL_HIDE_INTERNAL_SHEETS=0.
    """
    return env_bool("EXCEL_HIDE_INTERNAL_SHEETS", True) and not env_bool("EXCEL_SHOW_INTERNAL_SHEETS", False)


def analyst_internal_sheet_names() -> List[str]:
    raw = os.getenv("EXCEL_INTERNAL_SHEETS", "").strip()
    if raw:
        names = [x.strip() for x in raw.split(",") if x.strip()]
    else:
        names = list(DEFAULT_ANALYST_INTERNAL_SHEETS)
    if ANALYST_TRACE_SHEET not in names:
        names.append(ANALYST_TRACE_SHEET)
    return names


def apply_analyst_workbook_presentation(wb) -> None:
    """Aplica solo cambios de presentación al Excel final.

    No modifica valores, fórmulas ni cálculos: únicamente nombres de pestañas
    visibles y estado hidden/visible según configuración.
    """
    # Asegurar hojas visibles base para analistas.
    for name in wb.sheetnames:
        if (
            name == ANALYST_SUMMARY_SHEET
            or name == "Comparativa"
            or name == "Analisis experto IA"
            or name == "Detalle"
        ):
            wb[name].sheet_state = "visible"

    if analyst_hide_internal_sheets():
        for name in analyst_internal_sheet_names():
            if name in wb.sheetnames:
                wb[name].sheet_state = "hidden"

    # Excel exige al menos una hoja visible. Preferir resumen profesional, luego Detalle.
    visible = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]
    if not visible:
        for candidate in (ANALYST_SUMMARY_SHEET, "Detalle"):
            if candidate in wb.sheetnames:
                wb[candidate].sheet_state = "visible"
                visible = [wb[candidate]]
                break
    if visible:
        try:
            wb.active = wb.worksheets.index(visible[0])
        except Exception:
            pass


def _safe_sheet_title(base: str, existing: List[str]) -> str:
    """Nombre de hoja Excel <=31 chars y único."""
    raw = re.sub(r"[\\/\?\*\[\]:]", "-", str(base or "Hoja")).strip() or "Hoja"
    raw = raw[:31]
    title = raw
    i = 2
    while title in existing:
        suffix = f" {i}"
        title = raw[:31-len(suffix)] + suffix
        i += 1
    return title


def _copy_sheet_to_workbook(src_ws, dst_wb, title: str):
    """Copia valores, estilos, dimensiones y merges entre workbooks."""
    title = _safe_sheet_title(title, dst_wb.sheetnames)
    dst_ws = dst_wb.create_sheet(title)
    dst_ws.sheet_view.showGridLines = getattr(src_ws.sheet_view, "showGridLines", False)
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws[cell.coordinate]
            new_cell.value = cell.value
            if cell.has_style:
                new_cell._style = copy(cell._style)
            if cell.number_format:
                new_cell.number_format = cell.number_format
            if cell.font:
                new_cell.font = copy(cell.font)
            if cell.fill:
                new_cell.fill = copy(cell.fill)
            if cell.border:
                new_cell.border = copy(cell.border)
            if cell.alignment:
                new_cell.alignment = copy(cell.alignment)
            if cell.protection:
                new_cell.protection = copy(cell.protection)
            if cell.comment:
                new_cell.comment = copy(cell.comment)
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))
    for key, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[key].width = dim.width
        dst_ws.column_dimensions[key].hidden = dim.hidden
    for key, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[key].height = dim.height
        dst_ws.row_dimensions[key].hidden = dim.hidden
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes
    try:
        dst_ws.auto_filter.ref = src_ws.auto_filter.ref
    except Exception:
        pass
    return dst_ws


def _provider_short_name(provider: str, idx: int) -> str:
    name = str(provider or f"P{idx+1}").strip()
    name = re.sub(r"\s+", " ", name)
    # Mantener legible y evitar nombres de hoja demasiado largos.
    return name[:18] or f"P{idx+1}"


def _build_single_provider_workbook(provider_path: str, provider_name: str, output_path: str, meta: Dict[str, Any], catalogo_path: Optional[str], nacional_path: Optional[str], resumen_path: Optional[str], resumen_oficial_path: Optional[str]) -> str:
    """Ejecuta el motor individual completo para un proveedor.

    Intenta usar el mismo resumen/base disponible para que el Detalle resultante
    sea homologable al Excel individual. No usa el camino fast multi.
    """
    local_meta = dict(meta or {})
    local_meta["modo"] = "single_provider_embedded_in_multi"
    # Para el motor individual, el resumen del proveedor es el más importante.
    # Si existe resumen_oficial o catalogo, se pasa como catalogo/base; el resumen
    # del proveedor se pasa en resumen_paths para que alimente cantidades/PU reales.
    return build_comparativo(
        [provider_path],
        [provider_name],
        output_path,
        local_meta,
        catalogo_path=(resumen_oficial_path or catalogo_path),
        nacional_path=nacional_path,
        resumen_paths=[resumen_path] if resumen_path else [],
    )


def _write_multi_provider_ai_sheet(wb, analysis: Dict[str, Any]):
    ws = wb.create_sheet("Analisis experto IA"); ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:F1"); ws["A1"] = "DICTAMEN EJECUTIVO DE PRECIOS UNITARIOS - MULTI CONTRATISTA"; style_title(ws["A1"])
    ws.merge_cells("A2:F2"); ws["A2"] = "Lectura multi-proveedor generada con base en los totales, cobertura, faltantes y desviaciones calculadas. Los detalles técnicos están en la hoja Detalle."; style_cell(ws["A2"], SKY, True, True)
    r = 4
    headers = ["Rank", "Proveedor", "Total comparable", "Cobertura", "Riesgos", "Recomendación"]
    for c,h in enumerate(headers,1):
        ws.cell(r,c,h); style_header(ws.cell(r,c))
    r += 1
    scores = analysis.get("provider_scores", []) or []
    for s in scores:
        risk_bits = []
        if s.get("missing"): risk_bits.append(f"faltantes: {s.get('missing')}")
        if s.get("unit_mismatches"): risk_bits.append(f"unidades: {s.get('unit_mismatches')}")
        if s.get("extras"): risk_bits.append(f"adicionales: {s.get('extras')}")
        vals = [s.get("rank"), s.get("provider"), s.get("total_comparable"), s.get("coverage_pct"), "; ".join(risk_bits) or "Sin riesgos principales", s.get("recommendation")]
        for c,v in enumerate(vals,1):
            fill = GREEN if c == 2 and s.get("rank") == 1 else YELLOW if c == 5 and risk_bits else None
            ws.cell(r,c,v); style_cell(ws.cell(r,c), fill, c in (1,2), c in (5,6))
            if c == 3: fmt_money(ws.cell(r,c))
            if c == 4: fmt_pct(ws.cell(r,c))
        r += 1
    r += 2
    ws.cell(r,1,"Hallazgos y acciones sugeridas"); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6); style_header(ws.cell(r,1)); r += 1
    for c,h in enumerate(["Prioridad", "Tipo", "Hallazgo", "Acción sugerida", "Proveedor", "Nota"],1):
        ws.cell(r,c,h); style_header(ws.cell(r,c))
    r += 1
    findings = analysis.get("executive_findings", []) or []
    if not findings:
        findings = [{"tipo":"lectura", "severidad":"media", "mensaje":"No se detectaron hallazgos críticos con el resumen disponible."}]
    for f in findings[:15]:
        msg = f.get("mensaje") or ""
        accion = "Revisar el detalle del proveedor involucrado y solicitar aclaración técnica/económica si el hallazgo impacta la decisión."
        proveedor = ""
        for p in analysis.get("providers", []) or []:
            if p and p in msg:
                proveedor = p; break
        vals = [f.get("severidad") or "media", f.get("tipo") or "hallazgo", msg, accion, proveedor, "Ver hoja Detalle - proveedor correspondiente."]
        fill = RED if vals[0] == "alta" else YELLOW if vals[0] == "media" else SKY
        for c,v in enumerate(vals,1):
            ws.cell(r,c,v); style_cell(ws.cell(r,c), fill if c <= 2 else None, c <= 2, c in (3,4,6))
        r += 1
    for col,width in {"A":14,"B":20,"C":72,"D":70,"E":28,"F":48}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"



def write_detalle_provider_summary(wb, analysis: Dict[str, Any], provider_name: str, idx: int, error: Optional[str] = None):
    """Fallback: detalle resumido por proveedor con el formato actual."""
    ws = wb.create_sheet(_safe_sheet_title(f"Detalle - P{idx+1}", wb.sheetnames)); ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:P1"); ws["A1"] = f"MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO — {provider_name}"; style_title(ws["A1"])
    headers = ["Codigo","Concepto / Insumo","Unidad","Costo Contratista","Op.","Cantidad","Importe Contratista","Costo Mercado","Dif % Costo","Importe Mercado","Base Mercado %","Tipo","Match Construdata","Conf.","Estado","Nota"]
    for c,h in enumerate(headers,1): ws.cell(3,c,h); style_header(ws.cell(3,c))
    r = 4
    pastel = _provider_pastel(idx)
    for row in analysis.get("matrix_rows", []) or []:
        pr = (row.get("providers") or {}).get(provider_name, {})
        ws.cell(r,1,f"Analisis: {row.get('code') or ''} | {row.get('desc') or ''} | Cantidad real: {row.get('qty') or ''} | PU Mercado: {row.get('ref_pu') if row.get('ref_pu') is not None else ''}")
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=16); style_cell(ws.cell(r,1), pastel, True, True); r += 1
        vals = [
            pr.get("code") or row.get("code"), pr.get("desc") or row.get("desc"), pr.get("unit") or row.get("unit"),
            pr.get("pu"), "*", pr.get("qty") or row.get("qty"), pr.get("total"),
            row.get("ref_pu"), pr.get("delta_ref_pct"), row.get("ref_total"), None,
            "SERVICIO / RESUMEN MULTI", row.get("neodata_desc") or row.get("neodata_code") or "Referencia base común",
            row.get("neodata_match_score"), pr.get("status") or "sin_dato",
            ("Detalle fallback. " + error) if error else "Detalle resumido por proveedor. El motor individual no generó hoja Detalle."
        ]
        for c,v in enumerate(vals,1):
            ws.cell(r,c,v); style_cell(ws.cell(r,c), None, c in (1,12,15), c in (2,13,16))
            if c in (4,7,8,10): fmt_money(ws.cell(r,c))
            if c in (9,11,14): fmt_pct(ws.cell(r,c))
            if c == 6: fmt_num(ws.cell(r,c))
        r += 2
    for col,width in {"A":18,"B":72,"C":10,"D":17,"E":8,"F":12,"G":18,"H":17,"I":13,"J":18,"K":14,"L":24,"M":48,"N":10,"O":26,"P":78}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A4"



def _copy_if_exists(src_wb, dst_wb, src_name: str, dst_name: str, tab_color: Optional[str] = None, title_suffix: Optional[str] = None):
    """Copia una hoja si existe, preservando grafías, estilos, anchos y merges."""
    if src_name not in src_wb.sheetnames:
        return None
    ws = _copy_sheet_to_workbook(src_wb[src_name], dst_wb, dst_name)
    if tab_color:
        try:
            ws.sheet_properties.tabColor = tab_color
        except Exception:
            pass
    if title_suffix:
        try:
            for row in range(1, min(ws.max_row, 3) + 1):
                for col in range(1, min(ws.max_column, 6) + 1):
                    cell = ws.cell(row, col)
                    if isinstance(cell.value, str) and cell.value.strip():
                        cell.value = f"{cell.value} — {title_suffix}"
                        return ws
        except Exception:
            pass
    return ws


def write_unified_comparativa(wb, analysis: Dict[str, Any]):
    """Hoja 2 homologada: Comparativa.

    Debe existir para 1 proveedor, multi-proveedor y matriz propuesta. En el
    análisis de cotizaciones resume por concepto/base; el desglose técnico se
    conserva exclusivamente en la hoja Detalle.
    """
    ws = wb.create_sheet("Comparativa")
    ws.sheet_view.showGridLines = False
    providers = analysis.get("providers", []) or []
    provider_count = max(1, len(providers))
    end_col = max(12, 6 + provider_count * 4 + 2)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws.cell(1, 1, "COMPARATIVA EJECUTIVA DE COTIZACIONES")
    style_title(ws.cell(1, 1))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws.cell(2, 1, "Vista resumida por concepto. El desglose técnico, insumos, MO, equipo y referencias Construdata están en la hoja Detalle.")
    style_cell(ws.cell(2, 1), SKY, True, True)

    base_headers = ["Codigo", "Concepto", "Unidad", "Cantidad", "PU Mercado", "Importe Mercado"]
    headers = list(base_headers)
    for idx, provider in enumerate(providers or ["Proveedor 1"]):
        label = f"P{idx+1}"
        if provider:
            label = f"{label} {str(provider)[:18]}"
        headers.extend([f"{label} PU", f"{label} Importe", f"{label} Dif %", f"{label} Estado"])
    headers.extend(["Mejor proveedor", "Nota"])

    header_row = 4
    for c, h in enumerate(headers, 1):
        ws.cell(header_row, c, h)
        style_header(ws.cell(header_row, c))

    r = header_row + 1
    rows = analysis.get("matrix_rows", []) or []
    for row in rows:
        vals = [
            row.get("code"),
            row.get("desc"),
            row.get("unit"),
            row.get("qty"),
            row.get("ref_pu"),
            row.get("ref_total"),
        ]
        provider_values = []
        best_name = row.get("best_provider") or ""
        best_total = None
        for provider in providers or ["Proveedor 1"]:
            pr = (row.get("providers") or {}).get(provider, {}) or {}
            pu_v = pr.get("pu")
            total_v = pr.get("total")
            delta_v = pr.get("delta_ref_pct")
            status_v = pr.get("status") or "sin_dato"
            provider_values.extend([pu_v, total_v, delta_v, status_v])
            total_f = as_float(total_v)
            if total_f is not None and (best_total is None or total_f < best_total):
                best_total = total_f
                best_name = provider
        note_bits = []
        if row.get("neodata_desc") or row.get("neodata_code"):
            note_bits.append("Referencia mercado detectada")
        if not row.get("ref_pu"):
            note_bits.append("Sin PU mercado")
        if row.get("risk"):
            note_bits.append(str(row.get("risk")))
        vals.extend(provider_values)
        vals.extend([best_name, "; ".join(note_bits) or "Revisar hoja Detalle para soporte técnico."])
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v)
            style_cell(ws.cell(r, c), None, c in (1, len(headers)-1), c in (2, len(headers)))
            header = str(headers[c-1]).lower()
            if "importe" in header or header.endswith(" pu") or "mercado" in header and "pu" in header:
                fmt_money(ws.cell(r, c))
            if "dif %" in header:
                fmt_pct(ws.cell(r, c))
            if header == "cantidad":
                fmt_num(ws.cell(r, c))
        r += 1

    # Fallback si no hay matrix_rows, dejar una hoja válida y explicativa.
    if r == header_row + 1:
        ws.cell(r, 1, "Sin conceptos estructurados en Comparativa")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
        style_cell(ws.cell(r, 1), YELLOW, True, True)
        r += 1

    for col_idx, h in enumerate(headers, 1):
        htxt = str(h)
        if col_idx == 2:
            width = 72
        elif col_idx == len(headers):
            width = 58
        elif "Estado" in htxt:
            width = 22
        else:
            width = 16
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A5"
    try:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{max(r-1,4)}"
    except Exception:
        pass
    return ws


def _append_detail_sheet_to_combined(src_ws, dst_ws, provider_name: str, idx: int):
    """Agrega el Detalle individual a una sola hoja Detalle homologada.

    Se conserva formato/grafía original del Detalle individual, pero se agrega
    una columna A = Proveedor para que web/dashboard/chat puedan entender N
    proveedores dentro del mismo tab.
    """
    if src_ws is None:
        return None
    start_row = dst_ws.max_row + 2 if dst_ws.max_row and dst_ws.max_row > 1 else 1
    max_col = max(1, src_ws.max_column)
    end_col = max_col + 1
    # Separador del proveedor.
    dst_ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=end_col)
    title = dst_ws.cell(start_row, 1, f"PROVEEDOR {idx+1}: {provider_name}")
    style_title(title)
    try:
        dst_ws.sheet_properties.tabColor = _provider_pastel(idx)
    except Exception:
        pass
    target_row = start_row + 1
    header_like_terms = {"codigo", "código", "concepto / insumo", "costo contratista", "importe contratista", "match construdata"}
    for src_r in range(1, src_ws.max_row + 1):
        values = [src_ws.cell(src_r, c).value for c in range(1, max_col + 1)]
        normalized = {norm_text(v) for v in values if isinstance(v, str)}
        is_header = bool(normalized.intersection({norm_text(x) for x in header_like_terms})) and any("concepto" in x for x in normalized)
        dst_a = dst_ws.cell(target_row, 1, "Proveedor" if is_header else provider_name)
        if is_header:
            style_header(dst_a)
        else:
            style_cell(dst_a, _provider_pastel(idx) if src_r == 1 else None, bold=src_r == 1, wrap=True)
        for src_c in range(1, max_col + 1):
            src_cell = src_ws.cell(src_r, src_c)
            dst_cell = dst_ws.cell(target_row, src_c + 1)
            dst_cell.value = src_cell.value
            if src_cell.has_style:
                dst_cell._style = copy(src_cell._style)
            if src_cell.number_format:
                dst_cell.number_format = src_cell.number_format
            if src_cell.font:
                dst_cell.font = copy(src_cell.font)
            if src_cell.fill:
                dst_cell.fill = copy(src_cell.fill)
            if src_cell.border:
                dst_cell.border = copy(src_cell.border)
            if src_cell.alignment:
                dst_cell.alignment = copy(src_cell.alignment)
            if src_cell.protection:
                dst_cell.protection = copy(src_cell.protection)
            if src_cell.comment:
                dst_cell.comment = copy(src_cell.comment)
        try:
            dst_ws.row_dimensions[target_row].height = src_ws.row_dimensions[src_r].height
        except Exception:
            pass
        target_row += 1
    # Copiar merges desplazados una columna y al bloque correspondiente.
    row_offset = start_row
    for merged in src_ws.merged_cells.ranges:
        try:
            min_col, min_row, max_col_m, max_row_m = merged.bounds
            dst_ws.merge_cells(
                start_row=row_offset + min_row,
                start_column=min_col + 1,
                end_row=row_offset + max_row_m,
                end_column=max_col_m + 1,
            )
        except Exception:
            pass
    # Anchos: columna proveedor + ancho original desplazado.
    dst_ws.column_dimensions["A"].width = 22
    for key, dim in src_ws.column_dimensions.items():
        try:
            new_col = get_column_letter(openpyxl.utils.column_index_from_string(key) + 1)
            dst_ws.column_dimensions[new_col].width = dim.width
        except Exception:
            pass
    dst_ws.freeze_panes = "A5"
    return dst_ws


def _copy_single_provider_sheets(src_wb, dst_wb, provider_name: str, idx: int, total_providers: int, combined_detail_ws=None):
    """Copia hojas valiosas del análisis individual al workbook unificado.

    Regla final de producto: siempre una sola hoja Detalle. Para multi, cada
    análisis individual se anexa en la misma hoja Detalle con columna Proveedor.
    """
    color = _provider_pastel(idx)
    src_detail = src_wb["Detalle"] if "Detalle" in src_wb.sheetnames else None
    if combined_detail_ws is not None and src_detail is not None:
        return _append_detail_sheet_to_combined(src_detail, combined_detail_ws, provider_name, idx)
    if src_detail is not None:
        return _copy_if_exists(src_wb, dst_wb, "Detalle", "Detalle", tab_color=color)
    return None

def create_professional_mvp_workbook(workbook_path: str, filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> str:
    """Crea el Excel homologado de análisis de cotización.

    Regla final: los 3 casos de uso deben entregar las mismas 4 tabs visibles,
    en el mismo orden:
      1. Resumen Profesional
      2. Comparativa
      3. Detalle
      4. Analisis experto IA

    El motor acepta N proveedores. Si N > 1, cada proveedor se analiza con el
    proceso individual y se anexa dentro del mismo tab Detalle.
    """
    filepaths = list(filepaths or [])
    nombres = list(nombres or [])
    resumen_paths = list(resumen_paths or [])
    analysis = build_professional_mvp_analysis(filepaths, nombres, meta, catalogo_path, nacional_path, resumen_paths, resumen_oficial_path)

    wb = openpyxl.Workbook()
    if wb.active:
        wb.remove(wb.active)

    write_exec(wb, analysis)
    write_unified_comparativa(wb, analysis)
    detalle_ws = wb.create_sheet("Detalle")
    detalle_ws.sheet_view.showGridLines = False

    total_providers = len(filepaths)
    copied_ai = False
    for idx, fp in enumerate(filepaths):
        provider_name = (nombres[idx] if idx < len(nombres) and nombres[idx] else f"Proveedor {idx+1}")
        rp = resumen_paths[idx] if idx < len(resumen_paths) else None
        tmp = tempfile.NamedTemporaryFile(prefix=f"quantia_unified_p{idx+1}_", suffix=".xlsx", delete=False)
        tmp.close()
        try:
            _build_single_provider_workbook(fp, provider_name, tmp.name, meta or {}, catalogo_path, nacional_path, rp, resumen_oficial_path)
            src = openpyxl.load_workbook(tmp.name)
            _copy_single_provider_sheets(src, wb, provider_name, idx, total_providers, combined_detail_ws=detalle_ws)
            if total_providers == 1 and "Analisis experto IA" in src.sheetnames and "Analisis experto IA" not in wb.sheetnames:
                _copy_if_exists(src, wb, "Analisis experto IA", "Analisis experto IA")
                copied_ai = True
            src.close()
        except Exception as exc:
            # Fallback controlado dentro del mismo tab Detalle, no crea Detalle - Pn.
            write_detalle_provider_summary(wb, analysis, provider_name, idx, error=str(exc))
            fallback_name = f"Detalle - P{idx+1}"
            if fallback_name in wb.sheetnames:
                try:
                    _append_detail_sheet_to_combined(wb[fallback_name], detalle_ws, provider_name, idx)
                    del wb[fallback_name]
                except Exception:
                    pass
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass

    # Si no se pudo poblar Detalle con contenido real, dejar diagnóstico visible.
    if detalle_ws.max_row <= 1:
        detalle_ws.cell(1, 1, "DETALLE")
        style_title(detalle_ws.cell(1, 1))
        detalle_ws.cell(3, 1, "No se generaron renglones de detalle técnico. Revisar archivos de entrada y trazabilidad.")
        style_cell(detalle_ws.cell(3, 1), YELLOW, True, True)
        detalle_ws.column_dimensions["A"].width = 90

    if not copied_ai or total_providers != 1:
        if "Analisis experto IA" in wb.sheetnames:
            del wb["Analisis experto IA"]
        _write_multi_provider_ai_sheet(wb, analysis)

    write_risks(wb, analysis)
    write_trace(wb, analysis)

    # Defensa final: no deben existir tabs por proveedor ni layouts alternos.
    for obsolete in list(wb.sheetnames):
        if obsolete.startswith("Detalle - ") or obsolete in {"Matriz MultiProveedor", "Comparativa Multi", "Detalle Multi"}:
            del wb[obsolete]

    desired = ["Resumen Profesional", "Comparativa", "Detalle", "Analisis experto IA", "Alcance y Riesgos", "Trazabilidad Técnica", "Diagnóstico Catálogos"]
    ordered = [wb[n] for n in desired if n in wb.sheetnames]
    ordered += [ws for ws in wb._sheets if ws.title not in {w.title for w in ordered}]
    wb._sheets = ordered

    apply_analyst_workbook_presentation(wb)
    wb.save(workbook_path)
    return workbook_path

def append_professional_mvp_workbook(workbook_path: str, filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> str:
    """Compatibilidad: en v0.3.0 también usa el generador unificado.

    Ya no se agregan hojas multi alternas sobre un workbook preexistente.
    """
    return create_professional_mvp_workbook(
        workbook_path,
        filepaths,
        nombres,
        meta=meta,
        catalogo_path=catalogo_path,
        nacional_path=nacional_path,
        resumen_paths=resumen_paths or [],
        resumen_oficial_path=resumen_oficial_path,
    )


def write_exec(wb, a: Dict[str, Any]):
    ws = wb.create_sheet("Resumen Profesional"); ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:L1"); ws["A1"] = "RESUMEN PROFESIONAL - COMPARATIVO DE PRECIOS UNITARIOS"; style_title(ws["A1"])
    meta = a.get("meta", {}); cfg = a.get("settings", {}); neo = a.get("neodata", {})
    rows = [("Proyecto", meta.get("proyecto") or "Sin nombre"), ("Cliente", meta.get("cliente") or ""), ("Ubicación", meta.get("ubicacion") or ""), ("Modo", meta.get("modo") or "multi_proveedor"), ("Base común", a.get("base_source")), ("Referencia de mercado", reference_label(neo)), ("Indirecto mercado", cfg.get("market_indirect_pct")), ("Unidad estricta", "Sí" if cfg.get("strict_unit_matching") else "No"), ("Generado", a.get("generated_at"))]
    r = 3
    for k, v in rows:
        ws.cell(r,1,k); ws.cell(r,2,v); style_cell(ws.cell(r,1), GRAY, True); style_cell(ws.cell(r,2), wrap=True)
        if k == "Indirecto mercado": fmt_pct(ws.cell(r,2))
        r += 1
    r += 1; ws.cell(r,1,"Criterio profesional"); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=12); style_header(ws.cell(r,1)); r += 1
    for t in ["La comparación usa un universo base común: Resumen PU oficial si existe; si no, primer Resumen PU/proveedor.", "Faltantes y unidades incompatibles se penalizan; no se inventan precios.", "Adicionales fuera de base se muestran separados hasta aprobación del usuario.", "Neodata/Construdata o catálogos granulares se usan como referencia cuando hay evidencia suficiente."]:
        ws.cell(r,1,"•"); ws.cell(r,2,t); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=12); style_cell(ws.cell(r,1)); style_cell(ws.cell(r,2), wrap=True); r += 1
    r += 1; ws.cell(r,1,"Evaluación proveedor" if len(a.get("providers", [])) <= 1 else "Ranking proveedor"); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=12); style_header(ws.cell(r,1)); r += 1
    headers = ["Rank","Proveedor","Score","Total comparable","Total oferta","Total referencia","Cobertura","Faltantes","Unidad","Adicionales","Desv. prom. base","Dictamen"]
    for c,h in enumerate(headers,1): ws.cell(r,c,h); style_header(ws.cell(r,c))
    r += 1
    for s in a.get("provider_scores", []):
        vals = [s.get("rank"),s.get("provider"),s.get("total_score"),s.get("total_comparable"),s.get("total_bid"),s.get("total_reference"),s.get("coverage_pct"),s.get("missing"),s.get("unit_mismatches"),s.get("extras"),s.get("avg_abs_market_delta_pct"),s.get("recommendation")]
        for c,v in enumerate(vals,1):
            fill = GREEN if c == 3 and isinstance(v,(int,float)) and v >= 80 else RED if c in (8,9) and isinstance(v,(int,float)) and v > 0 else None
            ws.cell(r,c,v); style_cell(ws.cell(r,c), fill, c in (1,2), c==12)
            if c in (4,5,6): fmt_money(ws.cell(r,c))
            if c in (7,11): fmt_pct(ws.cell(r,c))
        r += 1
    r += 1; ws.cell(r,1,"Hallazgos ejecutivos"); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=12); style_header(ws.cell(r,1)); r += 1
    for f in a.get("executive_findings", []):
        fill = RED if f.get("severidad") == "alta" else YELLOW if f.get("severidad") == "media" else SKY
        ws.cell(r,1,f.get("severidad")); ws.cell(r,2,f.get("tipo")); ws.cell(r,3,f.get("mensaje")); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=12)
        for c in (1,2,3): style_cell(ws.cell(r,c), fill, c<3, True)
        r += 1
    widths = {"A":20,"B":28,"C":12,"D":18,"E":18,"F":18,"G":12,"H":10,"I":10,"J":12,"K":14,"L":58}
    for c,w in widths.items(): ws.column_dimensions[c].width = w
    ws.freeze_panes = "A16"


def write_matrix(wb, a: Dict[str, Any]):
    raise RuntimeError("write_matrix legacy deshabilitado: la salida unificada usa Comparativa homologada; no genera Matriz MultiProveedor.")
    ws = wb.create_sheet("Comparativa"); ws.sheet_view.showGridLines = False
    providers = a.get("providers", [])
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=max(10,10+len(providers)*4)); ws.cell(1,1,"MATRIZ CONTRA BASE COMÚN / REFERENCIA"); style_title(ws.cell(1,1))
    headers = ["ID","Código","Concepto base","Unidad","Cantidad","PU base","Total base","Match Neodata","Ganador","Riesgo"]
    r = 3; c = 1
    for h in headers:
        ws.cell(r,c,h); style_header(ws.cell(r,c)); c += 1
    for p in providers:
        for h in (f"{p} PU", f"{p} Total", f"{p} Δ Base", f"{p} Estado"):
            ws.cell(r,c,h); style_header(ws.cell(r,c)); c += 1
    r += 1
    for row in a.get("matrix_rows", []):
        vals = [row.get("base_id"),row.get("code"),row.get("desc"),row.get("unit"),row.get("qty"),row.get("ref_pu"),row.get("ref_total"),row.get("neodata_match_score"),row.get("best_provider"),row.get("risk")]
        for c,v in enumerate(vals,1):
            ws.cell(r,c,v); style_cell(ws.cell(r,c), YELLOW if c==10 and v else None, c in (1,9), c in (3,10))
            if c in (6,7): fmt_money(ws.cell(r,c))
            if c == 8: fmt_pct(ws.cell(r,c))
            if c == 5: fmt_num(ws.cell(r,c))
        c = 11
        for p in providers:
            pr = row.get("providers", {}).get(p, {})
            vals = [pr.get("pu"), pr.get("total"), pr.get("delta_ref_pct"), pr.get("status")]
            for j,v in enumerate(vals):
                fill = None
                if j == 3: fill = RED if v in ("no_cotizado","unidad_incompatible") else GREEN if v == "cotizado" else None
                if j == 2 and isinstance(v,(int,float)) and abs(v) >= 0.35: fill = YELLOW
                ws.cell(r,c+j,v); style_cell(ws.cell(r,c+j), fill, wrap=j==3)
                if j in (0,1): fmt_money(ws.cell(r,c+j))
                if j == 2: fmt_pct(ws.cell(r,c+j))
            c += 4
        r += 1
    for col,width in {"A":12,"B":16,"C":60,"D":10,"E":12,"F":16,"G":18,"H":14,"I":24,"J":38}.items(): ws.column_dimensions[col].width = width
    for col in range(11, 11+len(providers)*4): ws.column_dimensions[get_column_letter(col)].width = 16
    ws.freeze_panes = "A4"; ws.auto_filter.ref = ws.dimensions


def _risk_priority_from_provider_status(status: str, delta: Any) -> str:
    if status in {"no_cotizado", "unidad_incompatible"}:
        return "Alta"
    if isinstance(delta, (int, float)) and abs(delta) >= 0.35:
        return "Alta"
    if isinstance(delta, (int, float)) and abs(delta) >= 0.15:
        return "Media"
    return "Baja"


def _provider_pastel(idx: int) -> str:
    # Paleta pastel suave para identificar proveedores sin fatigar lectura.
    palette = ["D9EAF7", "E2F0D9", "FCE4D6", "EADCF8", "FFF2CC"]
    return palette[idx % len(palette)]


def write_comparativa(wb, a: Dict[str, Any]):
    raise RuntimeError("write_comparativa legacy deshabilitado: la salida unificada usa Comparativa homologada.")
    ws = wb.create_sheet("Comparativa"); ws.sheet_view.showGridLines = False
    providers = a.get("providers", []) or []
    end_col = max(14, 14 + max(0, len(providers)-1)*5)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws.cell(1,1,"COMPARATIVO DE COTIZACIÓN — MULTI PROVEEDOR"); style_title(ws.cell(1,1))
    ws.cell(2,1,"Criterio azul"); ws.cell(2,2,"Azul = partidas principales por importe dentro de la base común. Las filas conservan el orden declarado.")
    style_cell(ws.cell(2,1), SKY, True); style_cell(ws.cell(2,2), SKY, wrap=True)
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=end_col)
    ws.cell(3,1,"Fuente cantidades"); ws.cell(3,2,f"Base común: {a.get('base_source') or 'sin_base'} | Proveedores: {len(providers)} | Referencia: {reference_label(a.get('neodata', {}))}")
    style_cell(ws.cell(3,1), GRAY, True); style_cell(ws.cell(3,2), GRAY, wrap=True)
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=end_col)

    base_headers = ["Servicio","Descripcion","Unidad","Cantidad","PU Mercado","Importe Mercado"]
    provider_headers = []
    for p in providers:
        provider_headers.extend([f"{p} PU", f"{p} Importe", f"{p} Dif %", f"{p} Estado", f"{p} Nota"])
    tail_headers = ["Mejor proveedor", "Estado", "Nota"]
    headers = base_headers + provider_headers + tail_headers
    r = 5
    for c, h in enumerate(headers, 1):
        ws.cell(r,c,h); style_header(ws.cell(r,c))
    r += 1

    rows = a.get("matrix_rows", []) or []
    totals = []
    for row in rows:
        total = row.get("best_total") or row.get("ref_total") or 0
        if isinstance(total, (int, float)):
            totals.append((total, row))
    total_sum = sum(t for t, _ in totals) or 0
    accum = 0.0
    blue_ids = set()
    for total, row in sorted(totals, key=lambda x: x[0], reverse=True):
        if total_sum and accum / total_sum < 0.80:
            blue_ids.add(row.get("base_id")); accum += total

    for row in rows:
        fill_base = SKY if row.get("base_id") in blue_ids else None
        vals = [row.get("code"), row.get("desc"), row.get("unit"), row.get("qty"), row.get("ref_pu"), row.get("ref_total")]
        for c, v in enumerate(vals, 1):
            ws.cell(r,c,v); style_cell(ws.cell(r,c), fill_base, c == 1, c == 2)
            if c in (5,6): fmt_money(ws.cell(r,c))
            if c == 4: fmt_num(ws.cell(r,c))
        c = 7
        risk_notes = []
        for idx, p in enumerate(providers):
            pr = (row.get("providers") or {}).get(p, {})
            status = pr.get("status")
            delta = pr.get("delta_ref_pct")
            note = pr.get("method") or ""
            if status and status != "cotizado":
                risk_notes.append(f"{p}: {status}")
            elif isinstance(delta, (int, float)) and abs(delta) >= 0.35:
                risk_notes.append(f"{p}: desviación {delta:.1%}")
            vals = [pr.get("pu"), pr.get("total"), delta, status, note]
            pastel = _provider_pastel(idx)
            for j, v in enumerate(vals):
                fill = pastel
                if j == 3 and v in {"no_cotizado", "unidad_incompatible"}:
                    fill = RED
                elif j == 2 and isinstance(v, (int, float)) and abs(v) >= 0.35:
                    fill = YELLOW
                ws.cell(r,c+j,v); style_cell(ws.cell(r,c+j), fill, j == 3, j in (3,4))
                if j in (0,1): fmt_money(ws.cell(r,c+j))
                if j == 2: fmt_pct(ws.cell(r,c+j))
            c += 5
        estado = "REVISAR" if risk_notes else "OK"
        vals = [row.get("best_provider"), estado, "; ".join(risk_notes[:4]) or row.get("risk") or ""]
        for j, v in enumerate(vals):
            fill = YELLOW if estado == "REVISAR" else GREEN if j == 1 else None
            ws.cell(r,c+j,v); style_cell(ws.cell(r,c+j), fill, j < 2, j == 2)
        r += 1
    widths = {"A":16,"B":72,"C":10,"D":12,"E":16,"F":18}
    for col, width in widths.items(): ws.column_dimensions[col].width = width
    for col in range(7, 7+len(providers)*5): ws.column_dimensions[get_column_letter(col)].width = 16
    for col in range(7+len(providers)*5, 10+len(providers)*5): ws.column_dimensions[get_column_letter(col)].width = 24
    ws.freeze_panes = "A6"; ws.auto_filter.ref = ws.dimensions


def write_detalle(wb, a: Dict[str, Any]):
    """Hoja visible Detalle para fast multi con el encabezado actual.

    No pretende sustituir el detalle granular del motor completo; expone el
    detalle auditable disponible por servicio/proveedor usando la misma forma
    de columnas del Detalle actual.
    """
    ws = wb.create_sheet("Detalle"); ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:P1"); ws["A1"] = "MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO"; style_title(ws["A1"])
    headers = ["Codigo","Concepto / Insumo","Unidad","Costo Contratista","Op.","Cantidad","Importe Contratista","Costo Mercado","Dif % Costo","Importe Mercado","Base Mercado %","Tipo","Match Construdata","Conf.","Estado","Nota"]
    for c, h in enumerate(headers, 1):
        ws.cell(3,c,h); style_header(ws.cell(3,c))
    r = 4
    providers = a.get("providers", []) or []
    for row in a.get("matrix_rows", []) or []:
        # Encabezado de análisis por servicio, igual al patrón actual.
        title = f"Analisis: {row.get('code') or ''} | {row.get('desc') or ''} | Cantidad real: {row.get('qty') or ''} | PU Mercado: {row.get('ref_pu') if row.get('ref_pu') is not None else ''}"
        ws.cell(r,1,title); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=16)
        style_cell(ws.cell(r,1), SKY, True, True); r += 1
        for idx, p in enumerate(providers):
            pr = (row.get("providers") or {}).get(p, {})
            pastel = _provider_pastel(idx)
            ws.cell(r,1,p); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=16)
            style_cell(ws.cell(r,1), pastel, True, True); r += 1
            estado = pr.get("status") or "sin_dato"
            note_bits = []
            if pr.get("method"): note_bits.append(f"match: {pr.get('method')}")
            if pr.get("unit_ok") is False: note_bits.append("unidad incompatible")
            if pr.get("similarity") is not None: note_bits.append(f"similitud {pr.get('similarity'):.2f}") if isinstance(pr.get("similarity"), (int,float)) else None
            # Línea resumen total del concepto/proveedor.
            vals = [
                pr.get("code") or row.get("code"),
                pr.get("desc") or row.get("desc"),
                pr.get("unit") or row.get("unit"),
                pr.get("pu"),
                "*",
                pr.get("qty") or row.get("qty"),
                pr.get("total"),
                row.get("ref_pu"),
                pr.get("delta_ref_pct"),
                row.get("ref_total"),
                None,
                "SERVICIO / RESUMEN MULTI",
                row.get("neodata_desc") or row.get("neodata_code") or "Referencia base común",
                row.get("neodata_match_score"),
                estado,
                "; ".join(note_bits) or row.get("risk") or "Detalle resumido del flujo multi rápido. Para insumo granular por proveedor, ejecutar motor completo."
            ]
            for c, v in enumerate(vals, 1):
                fill = pastel
                if c == 15 and v in {"no_cotizado", "unidad_incompatible"}: fill = RED
                if c == 9 and isinstance(v, (int, float)) and abs(v) >= 0.35: fill = YELLOW
                ws.cell(r,c,v); style_cell(ws.cell(r,c), fill, c in (1,12,15), c in (2,13,16))
                if c in (4,7,8,10): fmt_money(ws.cell(r,c))
                if c in (9,11,14): fmt_pct(ws.cell(r,c))
                if c == 6: fmt_num(ws.cell(r,c))
            r += 1
        # Bloque financiero de referencia por servicio para conservar lectura APU.
        ws.cell(r,2,"COSTO DIRECTO / BASE"); ws.cell(r,7,row.get("ref_total")); ws.cell(r,10,row.get("ref_total")); ws.cell(r,12,"RESUMEN FINANCIERO MULTI")
        for c in range(1,17): style_cell(ws.cell(r,c), GRAY, c in (2,12), c in (2,16))
        fmt_money(ws.cell(r,7)); fmt_money(ws.cell(r,10)); r += 1
        r += 1
    for col,width in {"A":18,"B":72,"C":10,"D":17,"E":8,"F":12,"G":18,"H":17,"I":13,"J":18,"K":14,"L":24,"M":48,"N":10,"O":26,"P":78}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A4"; ws.auto_filter.ref = ws.dimensions


def write_risks(wb, a: Dict[str, Any]):
    ws = wb.create_sheet("Alcance y Riesgos"); ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1"); ws["A1"] = "ALCANCE, FALTANTES, ADICIONALES Y RIESGOS DE HOMOLOGACIÓN"; style_title(ws["A1"])
    r = 3; headers = ["Tipo","Proveedor","Código","Concepto","Unidad","PU","Total","Comentario"]
    for c,h in enumerate(headers,1): ws.cell(r,c,h); style_header(ws.cell(r,c))
    r += 1
    for row in a.get("matrix_rows", []):
        for p in a.get("providers", []):
            pr = row.get("providers", {}).get(p, {}); st = pr.get("status")
            if st not in {"no_cotizado","unidad_incompatible"}: continue
            vals = ["Faltante" if st=="no_cotizado" else "Unidad incompatible", p, row.get("code"), row.get("desc"), row.get("unit"), None, None, row.get("risk")]
            for c,v in enumerate(vals,1): ws.cell(r,c,v); style_cell(ws.cell(r,c), RED if st=="unidad_incompatible" else YELLOW, c<=2, c in (4,8))
            r += 1
    for p, rows in a.get("extra_rows", {}).items():
        for ex in rows[:80]:
            vals = ["Adicional fuera de base", p, ex.get("code"), ex.get("desc"), ex.get("unit"), ex.get("pu"), ex.get("total"), "No suma al total comparable base hasta aprobación del usuario."]
            for c,v in enumerate(vals,1):
                ws.cell(r,c,v); style_cell(ws.cell(r,c), SKY, c<=2, c in (4,8))
                if c in (6,7): fmt_money(ws.cell(r,c))
            r += 1
    for col,width in {"A":22,"B":26,"C":16,"D":70,"E":10,"F":16,"G":18,"H":60}.items(): ws.column_dimensions[col].width = width
    ws.freeze_panes = "A4"; ws.auto_filter.ref = ws.dimensions


def write_trace(wb, a: Dict[str, Any]):
    ws = wb.create_sheet("Trazabilidad Técnica"); ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:E1"); ws["A1"] = "TRAZABILIDAD TÉCNICA"; style_title(ws["A1"])
    rows = [("Versión", a.get("version")), ("Generado", a.get("generated_at")), ("Base común", a.get("base_source")), ("Conceptos base", a.get("base_count")), ("Neodata disponible", a.get("neodata",{}).get("available")), ("Neodata path", a.get("neodata",{}).get("path")), ("Neodata registros", a.get("neodata",{}).get("count")), ("Configuración", str(a.get("settings")))]
    try:
        if quantia_market_catalog_diagnostics:
            cd = quantia_market_catalog_diagnostics(False)
            rows.extend([
                ("Catálogos DATA_DIR resuelto", cd.get("resolved_data_dir")),
                ("Catálogos registros", str(cd.get("record_counts"))),
                ("Catálogos estado", cd.get("status")),
                ("Catálogos advertencia", cd.get("warning")),
            ])
    except Exception:
        pass
    r = 3
    for k,v in rows:
        ws.cell(r,1,k); ws.cell(r,2,v); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5); style_cell(ws.cell(r,1), GRAY, True); style_cell(ws.cell(r,2), wrap=True); r += 1
    r += 1; ws.cell(r,1,"Diagnóstico por proveedor"); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5); style_header(ws.cell(r,1)); r += 1
    for c,h in enumerate(["Proveedor","Archivo","Conceptos","Total detectado","Notas"],1): ws.cell(r,c,h); style_header(ws.cell(r,c))
    r += 1
    for d in a.get("diagnostics", []):
        vals = [d.get("provider"), d.get("file"), d.get("concepts"), d.get("total_bid"), "; ".join(d.get("errors") or [])]
        for c,v in enumerate(vals,1):
            ws.cell(r,c,v); style_cell(ws.cell(r,c), RED if c==5 and v else None, c==1, c==5)
            if c == 4: fmt_money(ws.cell(r,c))
        r += 1
    for col,width in {"A":24,"B":44,"C":14,"D":20,"E":90}.items(): ws.column_dimensions[col].width = width


def professional_mvp_status() -> Dict[str, Any]:
    return {"version": "quantia-comparador-apu-v0.2.4-analistas", "mode": "local-sin-redis-sin-db", "settings": settings(), "principles": ["Base común tipo Neodata/Resumen PU antes de comparar proveedores.", "Faltantes y unidades incompatibles penalizan; no se imputan como precio real.", "Adicionales fuera de base se reportan separados.", "Excel sigue siendo entregable formal; la UI es apoyo operativo."]}
