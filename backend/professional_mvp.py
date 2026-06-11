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




def _v035_processor_attr(name: str):
    """Obtiene funciones internas del processor raíz sin romper deploys backend/main.py.

    Se usa para reutilizar el matching real Construdata/granular del flujo
    histórico en lugar de inventar mercado en el nuevo layout de Excel.
    """
    try:
        try:
            from . import processor as _processor
        except Exception:
            import processor as _processor
        return getattr(_processor, name, None)
    except Exception:
        return None


def _v035_apply_real_market_pricing(conceptos: Dict[str, Any], diag: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Enriquece conceptos con granular_market_items usando el matcher real.

    No crea placeholders. Si el matcher no está disponible, deja el concepto sin
    mercado y guarda diagnóstico, para que Detalle muestre Sin match/blanco.
    """
    fn = _v035_processor_attr("_step08_apply_market_catalog_pricing")
    if callable(fn):
        try:
            return fn(conceptos)
        except Exception as exc:
            if isinstance(diag, dict):
                diag.setdefault("errors", []).append(f"Mercado Construdata no aplicado: {type(exc).__name__}: {exc}")
    else:
        if isinstance(diag, dict):
            diag.setdefault("errors", []).append("Mercado Construdata no aplicado: función _step08_apply_market_catalog_pricing no disponible")
    return conceptos


def _v035_financial_rows_for_concept(raw: Dict[str, Any]) -> List[Tuple[Any, Any, Any, Any, Any]]:
    fn = _v035_processor_attr("_step36_financial_rows") or _v035_processor_attr("_step13_financial_rows")
    if callable(fn):
        try:
            rows = fn(raw)
            if isinstance(rows, list):
                return rows
        except Exception:
            return []
    return []

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
    # V0.3.5: el Detalle necesita mercado real por insumo. Reutilizar el
    # matcher Construdata del motor actual antes de construir las filas.
    _v035_apply_real_market_pricing(conceptos, diag)
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


# ============================================================================
# V0.3.3 - Reestructura Excel desde cero: SOLO TAB Comparativa
# ============================================================================
# Esta etapa reinicia la salida Excel y deja un único tab visible llamado
# "Comparativa" para los 3 casos: presupuesto/matriz propuesta, 1 contratista
# y N contratistas. No se generan Detalle, Resumen Profesional, Analisis experto
# IA, Trazabilidad ni tabs auxiliares.

COMPARATIVA_BLUE_80 = "D9EAF7"
COMPARATIVA_GROUP_FILL = "17365D"
COMPARATIVA_SUBHEADER_FILL = "1F4E79"
COMPARATIVA_TOTAL_FILL = "D9EAD3"
COMPARATIVA_BASE_FILL = "E7E6E6"
COMPARATIVA_MARKET_FILL = "D9EAF7"


def _v033_provider_display_name(name: Any, idx: int) -> str:
    raw = str(name or "").strip()
    return raw or f"Contratista {idx + 1}"


def _v033_money(v: Any) -> Optional[float]:
    f = as_float(v)
    return round(float(f), 2) if f is not None else None


def _v033_num(v: Any) -> Optional[float]:
    f = as_float(v)
    return float(f) if f is not None else None


def _v033_sheet_style_base(ws) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 82
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 13
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(bottom=Side(style="hair", color="D9E2F3"))
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(name="Calibri", size=9, color=TEXT)


def _v033_style_group_header(cell, fill: str = COMPARATIVA_GROUP_FILL) -> None:
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=True, color=WHITE, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(bottom=Side(style="thin", color="B7B7B7"))


def _v033_style_sub_header(cell, fill: str = COMPARATIVA_SUBHEADER_FILL) -> None:
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=True, color=WHITE, size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(bottom=Side(style="thin", color="B7B7B7"))


def _v033_calc_blue_80_rows(row_indexes: List[int], importes_by_row: Dict[int, float]) -> set:
    """Devuelve las filas que concentran aproximadamente el 80% del importe.

    Se evalúa por contratista, de mayor a menor importe, incluyendo la partida
    que cruza el umbral del 80%.
    """
    positives = [(r, float(importes_by_row.get(r) or 0.0)) for r in row_indexes if (importes_by_row.get(r) or 0) > 0]
    total_val = sum(v for _, v in positives)
    if total_val <= 0:
        return set()
    threshold = total_val * 0.80
    selected = set()
    running = 0.0
    for r, v in sorted(positives, key=lambda x: x[1], reverse=True):
        if running < threshold:
            selected.add(r)
            running += v
        else:
            break
    return selected


def write_comparativa_stage1_workbook(workbook_path: str, analysis: Dict[str, Any]) -> str:
    """Crea desde cero un Excel con un único tab: Comparativa.

    Estructura obligatoria:
      Partida | Descripción | Unidad | Cantidad | [Contratista: P.U., Importe, % part, % ajuste]... | Mercado - P.U. | Mercado - Importe
    """
    providers = list(analysis.get("providers") or [])
    if not providers:
        providers = ["Contratista 1"]
    providers = [_v033_provider_display_name(p, i) for i, p in enumerate(providers)]
    rows = list(analysis.get("matrix_rows") or [])
    detail_cache = _v036_prepare_detail_cache(analysis, providers)
    detail_market_by_provider = {
        provider: _v036_consolidate_market_from_detail(detail_cache.get(provider, []), rows)
        for provider in providers
    }
    analysis["_v036_detail_cache"] = detail_cache

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparativa"

    # Encabezado en dos niveles: fila 1 grupos, fila 2 columnas reales.
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws.cell(1, 1, "Servicios / Cotización")
    _v033_style_group_header(ws.cell(1, 1), COMPARATIVA_BASE_FILL)
    ws.cell(1, 1).font = Font(bold=True, color=TEXT, size=10)

    headers = ["Partida", "Descripción", "Unidad", "Cantidad"]
    col = 5
    provider_blocks: Dict[str, Tuple[int, int]] = {}
    for idx, provider in enumerate(providers):
        start = col
        end = col + 3
        provider_blocks[provider] = (start, end)
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        ws.cell(1, start, provider)
        _v033_style_group_header(ws.cell(1, start), _provider_pastel(idx))
        ws.cell(1, start).font = Font(bold=True, color=TEXT, size=10)
        headers.extend(["P.U.", "Importe", "% part", "% ajuste"])
        col += 4

    market_start = col
    market_end = col + 1
    ws.merge_cells(start_row=1, start_column=market_start, end_row=1, end_column=market_end)
    ws.cell(1, market_start, "Mercado")
    _v033_style_group_header(ws.cell(1, market_start), COMPARATIVA_MARKET_FILL)
    ws.cell(1, market_start).font = Font(bold=True, color=TEXT, size=10)
    headers.extend(["Mercado - P.U.", "Mercado - Importe"])

    for c, h in enumerate(headers, 1):
        ws.cell(2, c, h)
        if c <= 4:
            _v033_style_sub_header(ws.cell(2, c), "666666")
        elif market_start <= c <= market_end:
            _v033_style_sub_header(ws.cell(2, c), "5B9BD5")
        else:
            _v033_style_sub_header(ws.cell(2, c), COMPARATIVA_SUBHEADER_FILL)

    first_data_row = 3
    row_idx = first_data_row
    provider_importes_by_excel_row: Dict[str, Dict[int, float]] = {p: {} for p in providers}

    for item in rows:
        partida = item.get("code") or item.get("partida") or item.get("part") or ""
        descripcion = item.get("desc") or item.get("descripcion") or item.get("concepto") or ""
        unidad = item.get("unit") or item.get("unidad") or ""
        cantidad = _v033_num(item.get("qty") or item.get("cantidad"))
        mercado_pu = _v033_money(item.get("ref_pu") or item.get("market_pu") or item.get("pu_mercado"))
        mercado_importe = _v033_money(item.get("ref_total") or item.get("market_total") or ((cantidad or 0) * mercado_pu if mercado_pu is not None else None))

        ws.cell(row_idx, 1, partida)
        ws.cell(row_idx, 2, descripcion)
        ws.cell(row_idx, 3, unidad)
        ws.cell(row_idx, 4, cantidad)
        ws.cell(row_idx, 4).number_format = '#,##0.0000'

        for idx, provider in enumerate(providers):
            start, _ = provider_blocks[provider]
            pdata = (item.get("providers") or {}).get(provider, {}) or {}
            # Fallback por si el nombre original estaba en otra grafía.
            if not pdata and item.get("providers"):
                for k, v in (item.get("providers") or {}).items():
                    if norm_text(k) == norm_text(provider):
                        pdata = v or {}
                        break
            prov_pu = _v033_money(pdata.get("pu") or pdata.get("precio_unitario"))
            prov_importe = _v033_money(pdata.get("total") or pdata.get("importe") or ((cantidad or 0) * prov_pu if prov_pu is not None else None))
            ajuste = None
            if prov_pu is not None and mercado_pu not in (None, 0):
                ajuste = (prov_pu - mercado_pu) / mercado_pu
            ws.cell(row_idx, start, prov_pu)
            ws.cell(row_idx, start + 1, prov_importe)
            ws.cell(row_idx, start + 2, None)  # % part se llena después de conocer total proveedor.
            ws.cell(row_idx, start + 3, ajuste)
            provider_importes_by_excel_row[provider][row_idx] = float(prov_importe or 0.0)

        ws.cell(row_idx, market_start, mercado_pu)
        ws.cell(row_idx, market_start + 1, mercado_importe)
        row_idx += 1

    last_data_row = row_idx - 1

    # Si no hay matrix_rows, dejar estructura vacía pero válida.
    if last_data_row < first_data_row:
        ws.cell(first_data_row, 1, "Sin partidas detectadas")
        ws.cell(first_data_row, 2, "No se recibieron conceptos estructurados para construir Comparativa.")
        last_data_row = first_data_row
        row_idx = first_data_row + 1

    # % part y azul 80/20 por contratista.
    data_rows = list(range(first_data_row, last_data_row + 1))
    blue_fill = PatternFill("solid", fgColor=COMPARATIVA_BLUE_80)
    for provider in providers:
        start, end = provider_blocks[provider]
        total_provider = sum(provider_importes_by_excel_row.get(provider, {}).values())
        blue_rows = _v033_calc_blue_80_rows(data_rows, provider_importes_by_excel_row.get(provider, {}))
        for r in data_rows:
            importe = provider_importes_by_excel_row.get(provider, {}).get(r, 0.0)
            ws.cell(r, start + 2, (importe / total_provider) if total_provider > 0 else None)
            if r in blue_rows:
                for c in range(start, end + 1):
                    ws.cell(r, c).fill = blue_fill

    # Fila totalizadora.
    total_row = last_data_row + 1
    ws.cell(total_row, 1, "TOTAL")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=4)
    for provider in providers:
        start, _ = provider_blocks[provider]
        total_provider = sum(provider_importes_by_excel_row.get(provider, {}).values())
        ws.cell(total_row, start + 1, total_provider)
    market_total = sum((_v033_money(ws.cell(r, market_start + 1).value) or 0.0) for r in data_rows)
    ws.cell(total_row, market_start + 1, market_total)

    # Formatos y estilo final.
    _v033_sheet_style_base(ws)
    for c in range(1, len(headers) + 1):
        # Reaplicar encabezados tras base style.
        if c <= 4:
            _v033_style_sub_header(ws.cell(2, c), "666666")
        elif market_start <= c <= market_end:
            _v033_style_sub_header(ws.cell(2, c), "5B9BD5")
        else:
            _v033_style_sub_header(ws.cell(2, c), COMPARATIVA_SUBHEADER_FILL)
    for merged in list(ws.merged_cells.ranges):
        min_col, min_row, max_col, max_row = merged.bounds
        if min_row == 1:
            # La celda superior izquierda conserva estilo; no tocar rangos merged.
            pass

    # Reaplicar encabezados de grupo.
    _v033_style_group_header(ws.cell(1, 1), COMPARATIVA_BASE_FILL); ws.cell(1, 1).font = Font(bold=True, color=TEXT, size=10)
    for idx, provider in enumerate(providers):
        start, _ = provider_blocks[provider]
        _v033_style_group_header(ws.cell(1, start), _provider_pastel(idx)); ws.cell(1, start).font = Font(bold=True, color=TEXT, size=10)
    _v033_style_group_header(ws.cell(1, market_start), COMPARATIVA_MARKET_FILL); ws.cell(1, market_start).font = Font(bold=True, color=TEXT, size=10)

    for r in range(first_data_row, total_row + 1):
        ws.cell(r, 2).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(r, 4).number_format = '#,##0.0000'
        for provider in providers:
            start, _ = provider_blocks[provider]
            ws.cell(r, start).number_format = '$#,##0.00'
            ws.cell(r, start + 1).number_format = '$#,##0.00'
            ws.cell(r, start + 2).number_format = '0.00%'
            ws.cell(r, start + 3).number_format = '0.00%'
        ws.cell(r, market_start).number_format = '$#,##0.00'
        ws.cell(r, market_start + 1).number_format = '$#,##0.00'

    for c in range(1, len(headers) + 1):
        letter = get_column_letter(c)
        if c == 2:
            ws.column_dimensions[letter].width = 82
        elif c in (1, 3, 4):
            ws.column_dimensions[letter].width = {1: 16, 3: 12, 4: 13}.get(c, 14)
        else:
            ws.column_dimensions[letter].width = 15

    for c in range(1, len(headers) + 1):
        cell = ws.cell(total_row, c)
        cell.fill = PatternFill("solid", fgColor=COMPARATIVA_TOTAL_FILL)
        cell.font = Font(bold=True, color=TEXT, size=9)
        cell.border = Border(top=Side(style="thin", color="666666"), bottom=Side(style="thin", color="666666"))

    try:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{total_row}"
    except Exception:
        pass
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 30

    # Defensa estricta: solo debe existir Comparativa.
    for sh in list(wb.sheetnames):
        if sh != "Comparativa":
            del wb[sh]

    Path(workbook_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(workbook_path)
    return workbook_path


def create_professional_mvp_workbook(workbook_path: str, filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> str:
    """V0.3.3: salida desde cero, únicamente tab Comparativa."""
    analysis = build_professional_mvp_analysis(
        list(filepaths or []),
        list(nombres or []),
        meta,
        catalogo_path,
        nacional_path,
        list(resumen_paths or []),
        resumen_oficial_path,
    )
    return write_comparativa_stage1_workbook(workbook_path, analysis)


def append_professional_mvp_workbook(workbook_path: str, filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> str:
    """V0.3.3: compatibilidad; no agrega tabs, reconstruye solo Comparativa."""
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


# ============================================================================
# V0.3.4 - Comparativa con mercado por proveedor + Detalle por contratista
# ============================================================================
# Etapa 2 de reconstruccion del Excel ideal:
#   1) Comparativa: bloque por contratista con P.U., Importe, % Part., % ajuste,
#      Mercado P.U. y Mercado Importe. No hay mercado global.
#   2) Detalle: una hoja por contratista, reutilizando la matriz parseada del
#      proceso individual cuando esta disponible. No genera Resumen Profesional,
#      Analisis experto IA, Trazabilidad ni Matriz MultiProveedor.

V034_MARKET_DIFF_FILL = "FFFFFF"  # v0.3.6: no yellow fill; diferencias en Detalle van en negritas
V034_SEPARATOR_FILL = "FFFFFF"



# V0.3.7 - Resumen textual individual por contratista dentro de Comparativa
# V0.3.8 - Reglas PU/APU ampliadas + tab Analisis IA con veredicto consolidado.
V037_SUMMARY_TITLE_FILL = "1F4E78"
V037_SUMMARY_PROVIDER_FILL = "D9EAF7"
V038_AI_TITLE_FILL = "1F4E78"
V038_AI_SECTION_FILL = "D9EAF7"


def _v037_short_text(value: Any, max_len: int = 90) -> str:
    txt = str(value or "").strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt if len(txt) <= max_len else txt[: max_len - 1].rstrip() + "…"


def _v037_money_label(value: Optional[float]) -> str:
    if value is None:
        return "N/D"
    try:
        return f"${float(value):,.2f}"
    except Exception:
        return "N/D"


def _v037_pct_label(value: Optional[float]) -> str:
    if value is None:
        return "N/D"
    try:
        return f"{float(value) * 100:.2f}%"
    except Exception:
        return "N/D"


def _v037_detail_rows_for_provider(detail_cache: Dict[str, List[Dict[str, Any]]], provider: str, idx: int) -> List[Dict[str, Any]]:
    if provider in detail_cache:
        return detail_cache.get(provider) or []
    key = f"P{idx + 1}"
    if key in detail_cache:
        return detail_cache.get(key) or []
    # Tolerancia a nombres normalizados.
    target = norm_text(provider)
    for k, v in (detail_cache or {}).items():
        if norm_text(k) == target:
            return v or []
    return []


def _v037_top_concentration_notes(
    ws,
    provider: str,
    block_start: int,
    data_rows: List[int],
    provider_total: float,
    blue_rows: set,
    max_items: int = 3,
) -> List[str]:
    if not data_rows or provider_total <= 0:
        return []
    selected = []
    accum = 0.0
    for rr in data_rows:
        if rr not in blue_rows:
            continue
        importe = _v033_money(ws.cell(rr, block_start + 1).value) or 0.0
        accum += float(importe)
        selected.append((rr, importe))
    if not selected:
        return []
    selected_sorted = sorted(selected, key=lambda x: float(x[1] or 0), reverse=True)
    parts = []
    for rr, importe in selected_sorted[:max_items]:
        partida = ws.cell(rr, 1).value or ""
        desc = _v037_short_text(ws.cell(rr, 2).value, 70)
        pct = (float(importe or 0) / provider_total) if provider_total else 0
        parts.append(f"{partida} ({_v037_pct_label(pct)}): {desc}")
    coverage = accum / provider_total if provider_total else None
    return [
        f"Las partidas sombreadas en azul concentran {_v037_pct_label(coverage)} del importe cotizado de {provider}.",
        "Principales partidas dentro del 80/20: " + "; ".join(parts) + ".",
    ]


def _v037_adjustment_notes(
    ws,
    provider: str,
    block_start: int,
    data_rows: List[int],
    importe_total: float,
    mercado_total: float,
    max_items: int = 3,
) -> List[str]:
    notes: List[str] = []
    if mercado_total and mercado_total > 0:
        ajuste_global = (importe_total - mercado_total) / mercado_total
        if ajuste_global > 0.01:
            notes.append(f"El ajuste global contra mercado es {_v037_pct_label(ajuste_global)} arriba de la referencia calculada para su propia matriz.")
        elif ajuste_global < -0.01:
            notes.append(f"El total cotizado está {_v037_pct_label(abs(ajuste_global))} por debajo del mercado calculado; validar alcance, exclusiones y cantidades antes de considerarlo ahorro real.")
        else:
            notes.append("El total cotizado se mantiene prácticamente alineado contra el mercado calculado para su matriz.")
    diffs = []
    for rr in data_rows:
        prov_imp = _v033_money(ws.cell(rr, block_start + 1).value)
        market_imp = _v033_money(ws.cell(rr, block_start + 5).value)
        if prov_imp is None or market_imp in (None, 0):
            continue
        diff_abs = float(prov_imp) - float(market_imp)
        diff_pct = diff_abs / float(market_imp)
        if abs(diff_pct) >= 0.05 or abs(diff_abs) >= 1000:
            diffs.append((abs(diff_abs), diff_pct, rr, diff_abs))
    if diffs:
        diffs.sort(reverse=True, key=lambda x: x[0])
        chunks = []
        for _abs_val, diff_pct, rr, diff_abs in diffs[:max_items]:
            partida = ws.cell(rr, 1).value or ""
            direction = "sobre mercado" if diff_abs > 0 else "bajo mercado"
            chunks.append(f"{partida} {_v037_pct_label(abs(diff_pct))} {direction}")
        notes.append("Partidas con mayor desviación contra mercado: " + "; ".join(chunks) + ".")
    return notes



def _v037_detail_business_rule_notes(detail_rows: List[Dict[str, Any]], provider: str) -> List[str]:
    """Reglas PU/APU para resumen individual.

    V0.3.8 amplía el set de señales sin inventar hallazgos: solo comenta cuando
    hay evidencia textual o numérica en el detalle/matriz/match.
    """
    notes: List[str] = []
    if not detail_rows:
        return notes

    rows = [r for r in detail_rows if str(r.get("kind") or "").lower() in {"item", "financial"}]
    item_rows = [r for r in rows if str(r.get("kind") or "").lower() == "item"] or rows

    def row_txt(r: Dict[str, Any]) -> str:
        return norm_text(f"{r.get('codigo') or ''} {r.get('concepto') or ''} {r.get('unidad') or ''}")

    def relevant_rows(keywords: List[str]) -> List[Dict[str, Any]]:
        keys = [norm_text(k) for k in keywords]
        return [r for r in rows if any(k in row_txt(r) for k in keys)]

    def high_market_diff(r: Dict[str, Any], threshold: float = 0.05) -> bool:
        imp = _v033_money(r.get("importe")); mi = _v033_money(r.get("market_importe"))
        if imp is not None and mi not in (None, 0):
            return abs((imp - mi) / mi) >= threshold or abs(imp - mi) >= 1000
        pu = _v033_money(r.get("pu")); mp = _v033_money(r.get("market_pu"))
        if pu is not None and mp not in (None, 0):
            return abs((pu - mp) / mp) >= threshold or abs(pu - mp) >= 1
        return False

    # Confiabilidad / match de mercado.
    no_match = [r for r in item_rows if not r.get("market_match_real") and (str(r.get("market_op") or "").lower() == "sin match" or not r.get("market_pu"))]
    if no_match:
        examples = []
        for r in no_match[:3]:
            label = _v037_short_text(r.get("codigo") or r.get("concepto"), 55)
            if label:
                examples.append(label)
        suffix = f" Ejemplos: {', '.join(examples)}." if examples else ""
        notes.append(f"Existen {len(no_match)} renglones sin match de mercado Construdata; requieren soporte o validación manual.{suffix}")

    low_conf = []
    for r in item_rows:
        conf = _v033_num(r.get("confidence") or r.get("confianza") or r.get("conf"))
        if conf is not None and conf < 0.65:
            low_conf.append(r)
    if low_conf:
        notes.append(f"Se identifican {len(low_conf)} renglones con baja confianza de match; la conclusión debe tomarse con reserva en esos insumos.")

    # Matriz / apertura técnica.
    lote_rows = relevant_rows(["lote", "global", "paquete"])
    if len(lote_rows) >= 2:
        notes.append("La matriz contiene varios conceptos tipo LOTE/global; se recomienda solicitar apertura adicional antes de tomar el importe como comparable.")
    zero_rows = [r for r in item_rows if (_v033_money(r.get("importe")) or 0) == 0]
    if zero_rows:
        notes.append(f"Hay {len(zero_rows)} renglones con importe cero o no calculable; validar que no existan omisiones o errores de captura.")

    # Mano de obra.
    labor = relevant_rows(["mano de obra", "supervisor", "oficial", "ayudante", "cuadrilla", "soldador", "argonero", "tubero", "electrico", "eléctrico", "obra", "seguridad"])
    if labor:
        elevated = []
        below = []
        for r in labor:
            pu = _v033_money(r.get("pu")); mp = _v033_money(r.get("market_pu"))
            label = _v037_short_text(r.get("concepto") or r.get("codigo"), 45)
            if pu is not None and mp not in (None, 0):
                diff = (pu - mp) / mp
                if diff > 0.05:
                    elevated.append(label)
                elif diff < -0.10:
                    below.append(label)
        if elevated:
            notes.append("Se observan posibles sobrecostos de mano de obra contra mercado en: " + "; ".join(elevated[:4]) + ".")
        if below:
            notes.append("Hay mano de obra por debajo de mercado; validar alcance, cuadrilla y rendimientos en: " + "; ".join(below[:4]) + ".")

    # EHS / seguridad industrial.
    ehs = relevant_rows(["epp", "proteccion personal", "protección personal", "seguridad", "supervisor de seguridad", "trabajo en altura", "altura", "permiso", "maniobra", "izaje", "soldadura", "corte", "electrico", "eléctrico"])
    if ehs:
        notes.append("Se identifican señales EHS/seguridad en la matriz; validar EPP, supervisión, permisos y controles de trabajo en sitio según alcance.")

    # Equipo móvil / flotillas / logística.
    equipment = relevant_rows(["montacargas", "grua", "grúa", "manlift", "andamio", "plataforma", "tijera", "camion", "camión", "camioneta", "transporte", "acarreo", "flete", "traslado", "maniobra", "izaje", "combustible", "operador", "seguro", "renta", "maquinaria", "equipo"])
    if equipment:
        elevated_eq = []
        total_imp = sum(float(_v033_money(r.get("importe")) or 0.0) for r in item_rows)
        eq_imp = sum(float(_v033_money(r.get("importe")) or 0.0) for r in equipment)
        for r in equipment:
            if high_market_diff(r):
                elevated_eq.append(_v037_short_text(r.get("concepto") or r.get("codigo"), 45))
        if elevated_eq:
            notes.append("Equipo/logística con diferencia contra mercado: " + "; ".join(elevated_eq[:4]) + ".")
        if total_imp > 0 and eq_imp / total_imp >= 0.20:
            notes.append(f"El equipo móvil/logística representa {_v037_pct_label(eq_imp / total_imp)} del detalle disponible; validar operador, combustible, seguro, traslado y jornada mínima.")

    # Herramienta menor, EPP, indirectos y financieros.
    herramienta = relevant_rows(["herramienta menor"])
    if herramienta:
        notes.append("Se identifican renglones de herramienta menor; validar porcentaje aplicado contra la referencia PMD usual del 5% cuando el dato esté disponible.")
    epp = relevant_rows(["epp", "proteccion personal", "protección personal", "seguridad e higiene"])
    if epp:
        notes.append("Se identifican cargos de EPP/seguridad; validar que el porcentaje sea consistente con alcance y política del proyecto.")

    financieros = relevant_rows(["indirecto", "financiamiento", "utilidad", "costo indirecto", "cargo adicional"])
    if financieros:
        for r in financieros:
            txt = norm_text(r.get("concepto") or "")
            val = _v033_money(r.get("pu") or r.get("importe"))
            market = _v033_money(r.get("market_pu") or r.get("market_importe"))
            if "indirect" in txt and val is not None and market is not None:
                if val > market:
                    notes.append("El costo indirecto calculado por el contratista supera la referencia de mercado disponible; revisar porcentaje y base aplicada.")
                elif val < market:
                    notes.append("El costo indirecto está por debajo de la referencia de mercado disponible; validar que no existan exclusiones de alcance.")
                break
        if any("financ" in norm_text(r.get("concepto") or "") for r in financieros):
            notes.append("Se detecta financiamiento/cargo financiero; validar que la base y el porcentaje aplicado estén justificados contractual y comercialmente.")

    # Calidad industrial / alimentaria.
    quality = relevant_rows(["acero inoxidable", "inoxidable", "304", "316", "sanitario", "soldadura sanitaria", "pulido", "acabado", "limpieza", "alimentaria", "grado alimenticio", "prueba", "puesta en marcha", "cip", "tuberia sanitaria", "tubería sanitaria", "valvula", "válvula"])
    if quality:
        notes.append("Para componentes industriales/sanitarios, validar materiales, acabado, limpieza, pruebas y criterios de aceptación de planta.")

    # Alcance / adjudicación.
    if no_match or lote_rows or zero_rows:
        notes.append("Antes de adjudicar, revisar alcance, exclusiones y soporte de matriz en partidas sin match, globales o con información insuficiente.")

    return notes

def _v037_build_provider_summary_notes(
    ws,
    provider: str,
    idx: int,
    block_start: int,
    data_rows: List[int],
    provider_total: float,
    market_total: float,
    blue_rows: set,
    detail_rows: List[Dict[str, Any]],
) -> List[str]:
    notes: List[str] = []
    notes.extend(_v037_top_concentration_notes(ws, provider, block_start, data_rows, provider_total, blue_rows))
    notes.extend(_v037_adjustment_notes(ws, provider, block_start, data_rows, provider_total, market_total))
    notes.extend(_v037_detail_business_rule_notes(detail_rows, provider))
    if not notes:
        notes.append("No se identificaron hallazgos concluyentes con la información disponible; revisar detalle de matriz y soporte del contratista.")
    # Compactar y limitar para lectura ejecutiva.
    clean: List[str] = []
    seen = set()
    for note in notes:
        txt = str(note or "").strip()
        if not txt:
            continue
        key = norm_text(txt)
        if key in seen:
            continue
        seen.add(key)
        clean.append(txt)
        if len(clean) >= 6:
            break
    return clean


def _v037_append_provider_summaries_to_comparativa(
    ws,
    providers: List[str],
    provider_blocks: Dict[str, Tuple[int, int]],
    data_rows: List[int],
    provider_totals: Dict[str, Dict[str, float]],
    provider_importes_by_excel_row: Dict[str, Dict[int, float]],
    detail_cache: Dict[str, List[Dict[str, Any]]],
    start_row: int,
    last_col: int,
) -> int:
    """Agrega resumen textual por contratista debajo de la fila totalizadora.

    Mantiene Comparativa y Detalle como están; solo añade una sección simple,
    textual, basada en datos calculados disponibles. No crea tablas nuevas.
    """
    if not providers:
        return start_row - 1
    # Separación visual: dejar una fila en blanco antes del título.
    title_row = start_row + 1
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=last_col)
    title_cell = ws.cell(title_row, 1, "Resumen individual por contratista")
    title_cell.fill = PatternFill("solid", fgColor=V037_SUMMARY_TITLE_FILL)
    title_cell.font = Font(bold=True, color="FFFFFF", size=10)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, last_col + 1):
        ws.cell(title_row, c).border = Border(top=Side(style="thin", color="666666"), bottom=Side(style="thin", color="666666"))
    row_cursor = title_row + 1
    for idx, provider in enumerate(providers):
        block_start, _block_end = provider_blocks[provider]
        provider_total = float((provider_totals.get(provider) or {}).get("importe") or 0.0)
        market_total = float((provider_totals.get(provider) or {}).get("mercado") or 0.0)
        blue_rows = _v033_calc_blue_80_rows(data_rows, provider_importes_by_excel_row.get(provider, {}))
        detail_rows = _v037_detail_rows_for_provider(detail_cache, provider, idx)
        notes = _v037_build_provider_summary_notes(ws, provider, idx, block_start, data_rows, provider_total, market_total, blue_rows, detail_rows)

        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=last_col)
        cell = ws.cell(row_cursor, 1, provider)
        cell.fill = PatternFill("solid", fgColor=V037_SUMMARY_PROVIDER_FILL)
        cell.font = Font(bold=True, color=TEXT, size=10)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        for c in range(1, last_col + 1):
            ws.cell(row_cursor, c).border = Border(bottom=Side(style="thin", color="D9E2F3"))
        row_cursor += 1
        for note in notes:
            ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=last_col)
            txt_cell = ws.cell(row_cursor, 1, f"• {note}")
            txt_cell.font = Font(name="Calibri", size=9, color=TEXT)
            txt_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            for c in range(1, last_col + 1):
                ws.cell(row_cursor, c).border = Border(bottom=Side(style="hair", color="E5E7EB"))
            ws.row_dimensions[row_cursor].height = 36
            row_cursor += 1
        row_cursor += 1
    return row_cursor - 1

def _v034_provider_market_pu(row: Dict[str, Any], pdata: Dict[str, Any]) -> Optional[float]:
    """Precio de mercado por proveedor.

    El modelo nuevo permite que cada proveedor tenga su propio PU de mercado
    calculado desde su matriz. Si el analisis todavia no trae ese dato granular,
    se usa la referencia de la partida como fallback controlado.
    """
    for key in ("market_pu", "mercado_pu", "ref_pu", "pu_mercado"):
        val = _v033_money(pdata.get(key))
        if val is not None:
            return val
    return _v033_money(row.get("ref_pu") or row.get("market_pu") or row.get("pu_mercado"))


def _v034_provider_market_total(row: Dict[str, Any], pdata: Dict[str, Any], qty_val: Optional[float], market_pu: Optional[float]) -> Optional[float]:
    for key in ("market_total", "mercado_importe", "ref_total", "importe_mercado"):
        val = _v033_money(pdata.get(key))
        if val is not None:
            return val
    val = _v033_money(row.get("ref_total") or row.get("market_total") or row.get("importe_mercado"))
    if val is not None:
        return val
    return _v033_money((qty_val or 0) * market_pu) if market_pu is not None else None


def _v034_set_block_right_border(ws, col_idx: int, first_row: int, last_row: int) -> None:
    thick = Side(style="medium", color="7F7F7F")
    for rr in range(first_row, last_row + 1):
        cell = ws.cell(rr, col_idx)
        left = cell.border.left
        top = cell.border.top
        bottom = cell.border.bottom
        cell.border = Border(left=left, right=thick, top=top, bottom=bottom)


def _v036_prepare_detail_cache(analysis: Dict[str, Any], providers: List[str]) -> Dict[str, List[Dict[str, Any]]]:
    """Obtiene el detalle real por proveedor una sola vez.

    V0.3.6: Comparativa debe tomar Mercado P.U./Importe desde el mismo
    Detalle del proveedor. Esta cache evita recalcular y asegura que Detalle
    y Comparativa usen la misma fuente.
    """
    injected = analysis.get("_v036_detail_cache")
    if isinstance(injected, dict):
        return injected
    meta = analysis.get("meta") or {}
    filepaths = list(meta.get("source_filepaths") or [])
    nombres = list(meta.get("source_provider_names") or analysis.get("providers") or providers or [])
    resumen_paths = list(meta.get("source_resumen_paths") or [])
    cache: Dict[str, List[Dict[str, Any]]] = {}
    if not filepaths:
        return cache
    for idx, fp in enumerate(filepaths):
        provider_name = providers[idx] if idx < len(providers) else (nombres[idx] if idx < len(nombres) else f"P{idx+1}")
        resumen_path = resumen_paths[idx] if idx < len(resumen_paths) else None
        try:
            detail_rows, _diag = _v034_extract_provider_detail_rows(fp, resumen_path, provider_name)
        except Exception as exc:
            detail_rows = [{
                "kind": "concept",
                "codigo": "ERROR",
                "concepto": f"No se pudo generar detalle: {type(exc).__name__}: {exc}",
                "unidad": "",
                "pu": None,
                "op": "",
                "cantidad": None,
                "importe": None,
                "pct": None,
                "market_pu": None,
                "market_op": "Sin match",
                "market_qty": None,
                "market_importe": None,
                "market_match_real": False,
            }]
        cache[provider_name] = detail_rows
    return cache


def _v036_norm_partida(value: Any) -> str:
    return norm_text(str(value or "")).replace(" ", "")


def _v036_consolidate_market_from_detail(detail_rows: List[Dict[str, Any]], base_rows: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Consolida mercado por partida desde el tab Detalle.

    Regla obligatoria: Detalle proveedor -> consolidacion por partida ->
    Comparativa. Si no hay match real, no se inventa mercado.
    """
    qty_by_code: Dict[str, Optional[float]] = {}
    desc_by_code: Dict[str, str] = {}
    for row in base_rows or []:
        code = _v036_norm_partida(row.get("code") or row.get("partida") or row.get("part") or row.get("codigo"))
        if not code:
            continue
        qty_by_code[code] = _v033_num(row.get("qty") or row.get("cantidad"))
        desc_by_code[code] = norm_text(row.get("desc") or row.get("descripcion") or row.get("concepto") or "")

    out: Dict[str, Dict[str, Any]] = {}
    current_code = ""
    item_sums: Dict[str, float] = {}
    item_has_market: Dict[str, bool] = {}

    for row in detail_rows or []:
        kind = row.get("kind")
        code_raw = row.get("codigo")
        code = _v036_norm_partida(code_raw)
        if kind == "concept":
            current_code = code
            qty = _v033_num(row.get("cantidad")) or qty_by_code.get(code)
            market_importe = _v033_money(row.get("market_importe"))
            market_pu = _v033_money(row.get("market_pu"))
            has_market = bool(row.get("market_match_real") and (market_importe is not None or market_pu is not None))
            if has_market:
                if market_importe is None and qty not in (None, 0) and market_pu is not None:
                    market_importe = _v033_money(qty * market_pu)
                if market_pu is None and qty not in (None, 0) and market_importe is not None:
                    market_pu = _v033_money(market_importe / qty)
                out[code] = {"market_pu": market_pu, "market_importe": market_importe, "has_market": True}
            continue

        if kind in {"item", "financial"} and current_code:
            if not row.get("market_match_real"):
                continue
            imp = _v033_money(row.get("market_importe"))
            if imp is None:
                imp = _v034_calc_market_importe(row.get("market_pu"), row.get("market_op"), row.get("market_qty"))
            if imp is not None:
                item_sums[current_code] = item_sums.get(current_code, 0.0) + float(imp)
                item_has_market[current_code] = True

    for code, total in item_sums.items():
        # No sobrescribir un total de concepto si ya estaba disponible; ese es el dato financiero final.
        if code in out and out[code].get("has_market"):
            continue
        qty = qty_by_code.get(code)
        market_pu = _v033_money(total / qty) if qty not in (None, 0) else None
        out[code] = {"market_pu": market_pu, "market_importe": _v033_money(total), "has_market": bool(item_has_market.get(code))}
    return out


def _v036_market_for_comparativa(item: Dict[str, Any], provider: str, pdata: Dict[str, Any], detail_market: Dict[str, Dict[str, Any]], cantidad: Optional[float]) -> Tuple[Optional[float], Optional[float], bool]:
    """Devuelve mercado para Comparativa desde Detalle consolidado.

    Si Detalle no trae match real para la partida, devuelve vacio. Solo usa
    campos del row/pdata si ya vienen marcados como mercado real/consolidado.
    """
    code = _v036_norm_partida(item.get("code") or item.get("partida") or item.get("part") or item.get("codigo"))
    found = detail_market.get(code) or {}
    if found.get("has_market"):
        market_pu = _v033_money(found.get("market_pu"))
        market_importe = _v033_money(found.get("market_importe"))
        if market_importe is None and market_pu is not None and cantidad is not None:
            market_importe = _v033_money(cantidad * market_pu)
        if market_pu is None and market_importe is not None and cantidad not in (None, 0):
            market_pu = _v033_money(market_importe / cantidad)
        return market_pu, market_importe, True

    # Fallback solo si el dato ya viene de una etapa consolidada de mercado real.
    if pdata.get("market_match_real") or pdata.get("market_from_detail") or pdata.get("mercado_real"):
        market_pu = _v033_money(pdata.get("market_pu") or pdata.get("mercado_pu") or pdata.get("pu_mercado"))
        market_importe = _v033_money(pdata.get("market_total") or pdata.get("mercado_importe") or pdata.get("importe_mercado"))
        if market_importe is None and market_pu is not None and cantidad is not None:
            market_importe = _v033_money(cantidad * market_pu)
        return market_pu, market_importe, market_pu is not None or market_importe is not None
    return None, None, False


def _v036_apply_market_diff_bold(ws, row_idx: int) -> None:
    """Marca diferencias solo con negritas en columnas de mercado.

    No usa fondo amarillo ni altera columnas de contratista.
    """
    def bold_market_cell(col_idx: int) -> None:
        cell = ws.cell(row_idx, col_idx)
        cell.font = copy(cell.font)
        cell.font = Font(name=cell.font.name or "Calibri", size=cell.font.sz or 9, color=cell.font.color.rgb if getattr(cell.font.color, 'type', None) == 'rgb' else TEXT, bold=True)

    left = _v033_num(ws.cell(row_idx, 4).value); market = _v033_num(ws.cell(row_idx, 10).value)
    if left is not None and market not in (None, 0):
        if abs(left - market) > 1.0 or abs((left - market) / market) > 0.005:
            bold_market_cell(10)
    left = _v033_num(ws.cell(row_idx, 6).value); market = _v033_num(ws.cell(row_idx, 12).value)
    if left is not None and market is not None and abs(left - market) > 0.0001:
        bold_market_cell(12)
    left = _v033_num(ws.cell(row_idx, 7).value); market = _v033_num(ws.cell(row_idx, 13).value)
    if left is not None and market not in (None, 0):
        if abs(left - market) > 1.0 or abs((left - market) / market) > 0.005:
            bold_market_cell(13)


def write_comparativa_stage2_workbook(workbook_path: str, analysis: Dict[str, Any]) -> str:
    """Crea Excel de etapa 2: Comparativa + Detalle por contratista.

    Comparativa usa mercado por proveedor dentro de cada bloque:
      P.U. | Importe | % Part. | % ajuste | Mercado P.U. | Mercado Importe
    """
    providers = list(analysis.get("providers") or [])
    if not providers:
        providers = ["Contratista 1"]
    providers = [_v033_provider_display_name(p, i) for i, p in enumerate(providers)]
    rows = list(analysis.get("matrix_rows") or [])
    detail_cache = _v036_prepare_detail_cache(analysis, providers)
    detail_market_by_provider = {
        provider: _v036_consolidate_market_from_detail(detail_cache.get(provider, []), rows)
        for provider in providers
    }
    analysis["_v036_detail_cache"] = detail_cache

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparativa"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws.cell(1, 1, "Servicios / Cotización")
    _v033_style_group_header(ws.cell(1, 1), COMPARATIVA_BASE_FILL)
    ws.cell(1, 1).font = Font(bold=True, color=TEXT, size=10)

    headers = ["Partida", "Descripción", "Unidad", "Cantidad"]
    col = 5
    provider_blocks: Dict[str, Tuple[int, int]] = {}
    provider_totals: Dict[str, Dict[str, float]] = {p: {"importe": 0.0, "mercado": 0.0, "cantidad": 0.0} for p in providers}
    provider_importes_by_excel_row: Dict[str, Dict[int, float]] = {p: {} for p in providers}

    for idx, provider in enumerate(providers):
        start = col
        end = col + 5
        provider_blocks[provider] = (start, end)
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        ws.cell(1, start, provider)
        _v033_style_group_header(ws.cell(1, start), _provider_pastel(idx))
        ws.cell(1, start).font = Font(bold=True, color=TEXT, size=10)
        headers.extend(["P.U.", "Importe", "% Part.", "% ajuste", "Mercado P.U.", "Mercado Importe"])
        col += 6

    for c, h in enumerate(headers, 1):
        ws.cell(2, c, h)
        if c <= 4:
            _v033_style_sub_header(ws.cell(2, c), "666666")
        else:
            _v033_style_sub_header(ws.cell(2, c), COMPARATIVA_SUBHEADER_FILL)

    first_data_row = 3
    row_idx = first_data_row
    for item in rows:
        partida = item.get("code") or item.get("partida") or item.get("part") or ""
        descripcion = item.get("desc") or item.get("descripcion") or item.get("concepto") or ""
        unidad = item.get("unit") or item.get("unidad") or ""
        cantidad = _v033_num(item.get("qty") or item.get("cantidad"))

        ws.cell(row_idx, 1, partida)
        ws.cell(row_idx, 2, descripcion)
        ws.cell(row_idx, 3, unidad)
        ws.cell(row_idx, 4, cantidad)
        ws.cell(row_idx, 4).number_format = '#,##0.0000'

        for idx, provider in enumerate(providers):
            start, _ = provider_blocks[provider]
            pdata = (item.get("providers") or {}).get(provider, {}) or {}
            if not pdata and item.get("providers"):
                for k, v in (item.get("providers") or {}).items():
                    if norm_text(k) == norm_text(provider):
                        pdata = v or {}
                        break
            prov_pu = _v033_money(pdata.get("pu") or pdata.get("precio_unitario"))
            prov_importe = _v033_money(pdata.get("total") or pdata.get("importe") or ((cantidad or 0) * prov_pu if prov_pu is not None else None))
            mercado_pu, mercado_importe, mercado_ok = _v036_market_for_comparativa(
                item,
                provider,
                pdata,
                detail_market_by_provider.get(provider, {}),
                cantidad,
            )
            ajuste = None
            if mercado_ok and prov_pu is not None and mercado_pu not in (None, 0):
                ajuste = (prov_pu - mercado_pu) / mercado_pu

            ws.cell(row_idx, start, prov_pu)
            ws.cell(row_idx, start + 1, prov_importe)
            ws.cell(row_idx, start + 2, None)  # % Part. se calcula despues.
            ws.cell(row_idx, start + 3, ajuste)
            ws.cell(row_idx, start + 4, mercado_pu)
            ws.cell(row_idx, start + 5, mercado_importe)

            provider_importes_by_excel_row[provider][row_idx] = float(prov_importe or 0.0)
            provider_totals[provider]["importe"] += float(prov_importe or 0.0)
            provider_totals[provider]["mercado"] += float(mercado_importe or 0.0)
            provider_totals[provider]["cantidad"] += float(cantidad or 0.0)
        row_idx += 1

    last_data_row = row_idx - 1
    if last_data_row < first_data_row:
        ws.cell(first_data_row, 1, "Sin partidas detectadas")
        ws.cell(first_data_row, 2, "No se recibieron conceptos estructurados para construir Comparativa.")
        last_data_row = first_data_row
        row_idx = first_data_row + 1

    data_rows = list(range(first_data_row, last_data_row + 1))
    blue_fill = PatternFill("solid", fgColor=COMPARATIVA_BLUE_80)
    for provider in providers:
        start, end = provider_blocks[provider]
        total_provider = provider_totals[provider]["importe"]
        blue_rows = _v033_calc_blue_80_rows(data_rows, provider_importes_by_excel_row.get(provider, {}))
        for rr in data_rows:
            importe = provider_importes_by_excel_row.get(provider, {}).get(rr, 0.0)
            ws.cell(rr, start + 2, (importe / total_provider) if total_provider > 0 else None)
            if rr in blue_rows:
                for cc in range(start, end + 1):
                    ws.cell(rr, cc).fill = blue_fill

    total_row = last_data_row + 1
    ws.cell(total_row, 1, "TOTAL")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=4)
    for provider in providers:
        start, _ = provider_blocks[provider]
        qty_total = provider_totals[provider]["cantidad"] or 0.0
        importe_total = provider_totals[provider]["importe"]
        mercado_total = provider_totals[provider]["mercado"]
        # P.U. total = PU ponderado por cantidad. Evita sumar PU simples.
        ws.cell(total_row, start, (importe_total / qty_total) if qty_total > 0 else None)
        ws.cell(total_row, start + 1, importe_total)
        ws.cell(total_row, start + 2, None)
        ws.cell(total_row, start + 3, None)
        ws.cell(total_row, start + 4, (mercado_total / qty_total) if qty_total > 0 else None)
        ws.cell(total_row, start + 5, mercado_total)

    _v033_sheet_style_base(ws)
    _v033_style_group_header(ws.cell(1, 1), COMPARATIVA_BASE_FILL)
    ws.cell(1, 1).font = Font(bold=True, color=TEXT, size=10)
    for idx, provider in enumerate(providers):
        start, end = provider_blocks[provider]
        _v033_style_group_header(ws.cell(1, start), _provider_pastel(idx))
        ws.cell(1, start).font = Font(bold=True, color=TEXT, size=10)
        _v034_set_block_right_border(ws, end, 1, total_row)

    for c, h in enumerate(headers, 1):
        if c <= 4:
            _v033_style_sub_header(ws.cell(2, c), "666666")
        else:
            _v033_style_sub_header(ws.cell(2, c), COMPARATIVA_SUBHEADER_FILL)

    for rr in range(first_data_row, total_row + 1):
        ws.cell(rr, 2).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(rr, 4).number_format = '#,##0.0000'
        for provider in providers:
            start, _ = provider_blocks[provider]
            for cc in (start, start + 1, start + 4, start + 5):
                ws.cell(rr, cc).number_format = '$#,##0.00'
            for cc in (start + 2, start + 3):
                ws.cell(rr, cc).number_format = '0.00%'

    for c in range(1, len(headers) + 1):
        letter = get_column_letter(c)
        if c == 2:
            ws.column_dimensions[letter].width = 82
        elif c in (1, 3, 4):
            ws.column_dimensions[letter].width = {1: 16, 3: 12, 4: 13}.get(c, 14)
        else:
            ws.column_dimensions[letter].width = 15

    for c in range(1, len(headers) + 1):
        cell = ws.cell(total_row, c)
        cell.fill = PatternFill("solid", fgColor=COMPARATIVA_TOTAL_FILL)
        cell.font = Font(bold=True, color=TEXT, size=9)
        cell.border = Border(top=Side(style="thin", color="666666"), bottom=Side(style="thin", color="666666"), right=cell.border.right)

    # V0.3.7: debajo del total, agregar resumen textual individual por contratista.
    # No modifica la estructura de Comparativa ni Detalle; solo añade lectura ejecutiva simple.
    summary_last_row = _v037_append_provider_summaries_to_comparativa(
        ws,
        providers,
        provider_blocks,
        data_rows,
        provider_totals,
        provider_importes_by_excel_row,
        detail_cache,
        total_row + 1,
        len(headers),
    )

    try:
        # El filtro se mantiene sobre la tabla de servicios, no sobre el bloque textual.
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{total_row}"
    except Exception:
        pass
    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 32
    if summary_last_row:
        for rr in range(total_row + 2, summary_last_row + 1):
            if ws.row_dimensions[rr].height is None:
                ws.row_dimensions[rr].height = 24

    return _v034_append_detail_sheets(wb, workbook_path, analysis, detail_cache=detail_cache)


def _v034_item_market_value(item: Dict[str, Any], *keys: str) -> Optional[float]:
    for key in keys:
        val = _v033_money(item.get(key))
        if val is not None:
            return val
    return None


def _v034_calc_market_importe(pu_val: Optional[float], op_val: Any, qty_val: Optional[float], base_val: Optional[float] = None) -> Optional[float]:
    """Calcula importe mercado respetando operador, sin usar importe contratista."""
    pu_num = _v033_num(pu_val)
    qty_num = _v033_num(qty_val)
    op = str(op_val or "*").strip()
    if pu_num is None:
        return None
    if op == "/":
        if qty_num in (None, 0):
            return None
        return _v033_money(pu_num / qty_num)
    if op == "%":
        base_num = _v033_num(base_val)
        if base_num is None or qty_num is None:
            return None
        return _v033_money(base_num * qty_num)
    if qty_num is None:
        return None
    return _v033_money(pu_num * qty_num)


def _v034_is_real_market_match(item: Dict[str, Any]) -> bool:
    """Determina si el mercado proviene de Construdata y no del fallback contratista."""
    if item.get("fallback_contratista") is True:
        return False
    status = str(item.get("match_status") or "").strip().lower()
    if "sin_match" in status or "fallback" in status:
        return False
    if item.get("codigo_mercado") or item.get("descripcion_mercado") or item.get("source_file"):
        return True
    conf = _v033_num(item.get("confidence"))
    return bool(conf is not None and conf > 0)


def _v035_real_market_fields(item: Dict[str, Any]) -> Dict[str, Any]:
    """Mapea columnas de mercado visibles desde granular_market_items.

    Si el renglón no tiene match real contra Construdata, no copia valores del
    contratista como mercado. Deja PU/Cantidad/Importe vacíos y marca el operador
    como Sin match para que el analista lo vea sin contaminar cálculos.
    """
    if not _v034_is_real_market_match(item):
        return {"market_pu": None, "market_op": "Sin match", "market_qty": None, "market_importe": None, "market_match_real": False}
    market_pu = _v034_item_market_value(item, "precio_mercado", "market_pu", "mercado_pu", "precio_nacional")
    market_qty = _v034_item_market_value(item, "cantidad_mercado", "market_qty", "mercado_cantidad", "cantidad", "factor")
    market_op = item.get("market_op") or item.get("mercado_op") or item.get("op_mercado") or item.get("op") or item.get("operacion") or "*"
    market_importe = _v034_item_market_value(item, "importe_mercado", "market_importe", "mercado_importe")
    if market_importe is None:
        market_importe = _v034_calc_market_importe(market_pu, market_op, market_qty, item.get("base_mercado"))
    return {"market_pu": market_pu, "market_op": market_op, "market_qty": market_qty, "market_importe": market_importe, "market_match_real": True}


def _v034_extract_provider_detail_rows(provider_file: str, resumen_file: Optional[str], provider_name: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Reutiliza el parser y matcher individual para obtener Detalle real.

    La prioridad es granular_market_items, que ya trae match Construdata,
    precio_mercado, importe_mercado, confianza y estado. Si no existe match real,
    NO se copia el valor del contratista como mercado.
    """
    rows, diag = read_provider(provider_file, resumen_file, provider_name)
    detail_rows: List[Dict[str, Any]] = []
    provider_total = sum((r.get("total") or 0.0) for r in rows)
    for concept in rows:
        raw = concept.get("raw") or {}
        service_code = concept.get("code") or raw.get("clave") or ""
        service_desc = concept.get("desc") or raw.get("desc") or ""
        concept_importe = _v033_money(concept.get("total"))
        market_pu_concept = _v033_money(raw.get("granular_market_unit") or raw.get("pu_mercado") or raw.get("ref_pu"))
        if market_pu_concept is None:
            # Usar cálculo financiero completo si está disponible.
            fin_rows = _v035_financial_rows_for_concept(raw)
            if fin_rows:
                last_market = fin_rows[-1][2] if len(fin_rows[-1]) > 2 else None
                market_pu_concept = _v033_money(last_market)
        market_importe_concept = _v033_money(((concept.get("qty") or 0) * market_pu_concept) if market_pu_concept is not None else None)
        detail_rows.append({
            "kind": "concept",
            "codigo": service_code,
            "concepto": service_desc,
            "unidad": concept.get("unit"),
            "pu": concept.get("pu"),
            "op": "",
            "cantidad": concept.get("qty"),
            "importe": concept_importe,
            "pct": (concept_importe or 0) / provider_total if provider_total else None,
            "market_pu": market_pu_concept,
            "market_op": "*" if market_pu_concept is not None else "",
            "market_qty": concept.get("qty") if market_pu_concept is not None else None,
            "market_importe": market_importe_concept,
            "market_match_real": bool(market_pu_concept is not None),
        })

        granular_items = list(raw.get("granular_market_items") or [])
        if granular_items:
            for item in granular_items:
                importe = _v033_money(item.get("importe_proveedor") or item.get("importe"))
                market = _v035_real_market_fields(item)
                match_label = f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(" -")
                status = item.get("match_status") or ("match" if market.get("market_match_real") else "sin_match_mercado")
                diff_pu = None
                diff_pu_pct = None
                diff_imp = None
                diff_imp_pct = None
                prov_pu = _v033_money(item.get("precio_proveedor") or item.get("precio_base") or item.get("pu"))
                prov_imp = importe
                if prov_pu is not None and market.get("market_pu") not in (None, 0):
                    diff_pu = prov_pu - float(market["market_pu"])
                    diff_pu_pct = diff_pu / float(market["market_pu"])
                if prov_imp is not None and market.get("market_importe") not in (None, 0):
                    diff_imp = prov_imp - float(market["market_importe"])
                    diff_imp_pct = diff_imp / float(market["market_importe"])
                detail_rows.append({
                    "kind": "item",
                    "codigo": item.get("codigo") or item.get("code") or "",
                    "concepto": item.get("descripcion") or item.get("desc") or "",
                    "unidad": item.get("unidad") or item.get("unit") or "",
                    "pu": prov_pu,
                    "op": item.get("op") or item.get("operacion") or "*",
                    "cantidad": _v033_num(item.get("cantidad") or item.get("factor") or item.get("qty")),
                    "importe": importe,
                    "pct": (importe or 0) / provider_total if provider_total else None,
                    "market_pu": market.get("market_pu"),
                    "market_op": market.get("market_op"),
                    "market_qty": market.get("market_qty"),
                    "market_importe": market.get("market_importe"),
                    "market_match_real": market.get("market_match_real"),
                    "match_construdata": match_label,
                    "confidence": item.get("confidence"),
                    "estado": status,
                    "diff_pu": diff_pu,
                    "diff_pu_pct": diff_pu_pct,
                    "diff_importe": diff_imp,
                    "diff_importe_pct": diff_imp_pct,
                })
        else:
            # Fallback de presentación: pinta la matriz del contratista, pero deja
            # mercado en blanco/Sin match. No se copian valores contratista.
            for list_name in ("materiales_items", "mano_obra_items", "equipo_items", "basicos_items"):
                for item in raw.get(list_name) or []:
                    importe = _v033_money(item.get("importe"))
                    detail_rows.append({
                        "kind": "item",
                        "codigo": item.get("codigo") or item.get("code") or "",
                        "concepto": item.get("descripcion") or item.get("desc") or "",
                        "unidad": item.get("unidad") or item.get("unit") or "",
                        "pu": _v033_money(item.get("precio_base") or item.get("pu") or item.get("precio_unitario")),
                        "op": item.get("op") or item.get("operacion") or "*",
                        "cantidad": _v033_num(item.get("factor") or item.get("cantidad") or item.get("qty")),
                        "importe": importe,
                        "pct": (importe or 0) / provider_total if provider_total else None,
                        "market_pu": None,
                        "market_op": "Sin match",
                        "market_qty": None,
                        "market_importe": None,
                        "market_match_real": False,
                        "estado": "sin_match_mercado",
                    })

        # Filas financieras de mercado: se agregan si el motor las puede calcular.
        fin_rows = _v035_financial_rows_for_concept(raw)
        if fin_rows:
            for label, prov_amt, market_amt, pct_val, note in fin_rows:
                detail_rows.append({
                    "kind": "financial",
                    "codigo": "",
                    "concepto": str(label or "").upper(),
                    "unidad": "",
                    "pu": _v033_money(prov_amt),
                    "op": "",
                    "cantidad": None,
                    "importe": _v033_money(prov_amt),
                    "pct": pct_val if isinstance(pct_val, (int, float)) else ((_v033_money(prov_amt) or 0) / provider_total if provider_total else None),
                    "market_pu": _v033_money(market_amt),
                    "market_op": "",
                    "market_qty": None,
                    "market_importe": _v033_money(market_amt),
                    "market_match_real": True,
                    "estado": "calculo_financiero_mercado",
                    "nota": note,
                })
    return detail_rows, diag

def _v034_style_detail_sheet(ws, last_row: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    widths = {
        "A": 16, "B": 72, "C": 12, "D": 15, "E": 8, "F": 12, "G": 16, "H": 11,
        "I": 4, "J": 17, "K": 10, "L": 13, "M": 17,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for c in range(1, 14):
        cell = ws.cell(1, c)
        if c == 9:
            cell.fill = PatternFill("solid", fgColor=V034_SEPARATOR_FILL)
        elif c >= 10:
            _v033_style_sub_header(cell, "5B9BD5")
        else:
            _v033_style_sub_header(cell, COMPARATIVA_SUBHEADER_FILL)
    for r in range(2, last_row + 1):
        for c in range(1, 14):
            cell = ws.cell(r, c)
            cell.border = Border(bottom=Side(style="hair", color="D9E2F3"))
            cell.font = Font(name="Calibri", size=9, color=TEXT)
            cell.alignment = Alignment(vertical="center", wrap_text=(c == 2))
        ws.cell(r, 4).number_format = '$#,##0.00'
        ws.cell(r, 6).number_format = '#,##0.0000'
        ws.cell(r, 7).number_format = '$#,##0.00'
        ws.cell(r, 8).number_format = '0.00%'
        ws.cell(r, 10).number_format = '$#,##0.00'
        ws.cell(r, 12).number_format = '#,##0.0000'
        ws.cell(r, 13).number_format = '$#,##0.00'
    try:
        ws.auto_filter.ref = f"A1:M{last_row}"
    except Exception:
        pass


def _v034_apply_market_diff_fill(ws, row_idx: int) -> None:
    """Compatibilidad v0.3.6: ya no pinta amarillo; solo negritas en mercado."""
    _v036_apply_market_diff_bold(ws, row_idx)


def _v034_create_detail_sheet(wb, sheet_name: str, detail_rows: List[Dict[str, Any]]) -> None:
    ws = wb.create_sheet(_safe_sheet_title(sheet_name, wb.sheetnames))
    headers = ["Código", "Concepto", "Unidad", "P. Unitario", "Op.", "Cantidad", "Importe", "%", "", "Mercado P. Unitario", "Mercado Op.", "Mercado Cantidad", "Mercado Importe"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    row_kinds: Dict[int, str] = {}
    r = 2
    for row in detail_rows:
        vals = [
            row.get("codigo"), row.get("concepto"), row.get("unidad"), row.get("pu"), row.get("op"),
            row.get("cantidad"), row.get("importe"), row.get("pct"), "",
            row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("market_importe"),
        ]
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v)
        row_kinds[r] = str(row.get("kind") or "item")
        r += 1
    if r == 2:
        ws.cell(2, 1, "Sin detalle detectado")
        ws.cell(2, 2, "No fue posible leer renglones de matriz para este contratista.")
        row_kinds[2] = "concept"
        r = 3

    last_row = r - 1
    _v034_style_detail_sheet(ws, last_row)

    # Reaplicar estilos semánticos y diferencias despues del estilo base, porque
    # _v034_style_detail_sheet normaliza fuentes/fondos.
    for rr, kind in row_kinds.items():
        if kind == "concept":
            for c in range(1, 14):
                ws.cell(rr, c).fill = PatternFill("solid", fgColor="E7E6E6")
                ws.cell(rr, c).font = Font(bold=True, color=TEXT, size=9)
        elif kind == "financial":
            for c in range(1, 14):
                ws.cell(rr, c).fill = PatternFill("solid", fgColor="D9EAF7")
                ws.cell(rr, c).font = Font(bold=True, color=TEXT, size=9)
        # Diferencias: solo negritas en J/L/M, nunca amarillo.
        _v036_apply_market_diff_bold(ws, rr)



# ==========================================================================
# V0.3.8 - Analisis IA consolidado
# ==========================================================================

def _v038_detect_providers_from_comparativa(ws) -> List[Dict[str, Any]]:
    providers: List[Dict[str, Any]] = []
    max_col = ws.max_column or 0
    col = 5
    idx = 0
    while col <= max_col:
        name = str(ws.cell(1, col).value or f"Contratista {idx + 1}").strip()
        if not name:
            name = f"Contratista {idx + 1}"
        providers.append({"name": name, "start_col": col, "end_col": min(col + 5, max_col), "idx": idx})
        col += 6
        idx += 1
    return providers


def _v038_find_total_row(ws) -> int:
    for rr in range(1, (ws.max_row or 1) + 1):
        if norm_text(ws.cell(rr, 1).value) == "total":
            return rr
    return max(3, ws.max_row or 3)


def _v038_comparativa_metrics(ws, provider_info: Dict[str, Any], total_row: int) -> Dict[str, Any]:
    start = int(provider_info["start_col"])
    name = str(provider_info.get("name") or "Contratista")
    data_rows = list(range(3, max(2, total_row)))
    total = _v033_money(ws.cell(total_row, start + 1).value) or 0.0
    mercado = _v033_money(ws.cell(total_row, start + 5).value) or 0.0
    ajuste = ((total - mercado) / mercado) if mercado else None
    rows = []
    for rr in data_rows:
        partida = ws.cell(rr, 1).value
        if not partida:
            continue
        importe = _v033_money(ws.cell(rr, start + 1).value) or 0.0
        mercado_imp = _v033_money(ws.cell(rr, start + 5).value)
        diff_pct = None
        if mercado_imp not in (None, 0):
            diff_pct = (importe - mercado_imp) / mercado_imp
        rows.append({
            "row": rr,
            "partida": str(partida),
            "descripcion": str(ws.cell(rr, 2).value or ""),
            "importe": float(importe or 0.0),
            "mercado_importe": mercado_imp,
            "diff_pct": diff_pct,
        })
    rows_sorted = sorted(rows, key=lambda r: r.get("importe") or 0.0, reverse=True)
    top_rows = rows_sorted[:3]
    return {"name": name, "total": total, "mercado": mercado, "ajuste": ajuste, "rows": rows, "top_rows": top_rows}


def _v038_detail_sheet_name(wb, provider_info: Dict[str, Any], total_providers: int) -> Optional[str]:
    idx = int(provider_info.get("idx") or 0)
    if total_providers == 1 and "Detalle" in wb.sheetnames:
        return "Detalle"
    candidate = f"Detalle - P{idx + 1}"
    if candidate in wb.sheetnames:
        return candidate
    # Tolerancia a nombres con contratista completo.
    for sh in wb.sheetnames:
        if sh.startswith("Detalle") and str(idx + 1) in sh:
            return sh
    return None


def _v038_detail_signals(ws) -> Dict[str, Any]:
    signals = {
        "no_match": 0,
        "elevated_market": [],
        "labor": False,
        "ehs": False,
        "equipment": False,
        "quality": False,
        "financial": False,
        "generic": 0,
    }
    if ws is None:
        return signals
    for rr in range(2, (ws.max_row or 1) + 1):
        codigo = ws.cell(rr, 1).value
        concepto = ws.cell(rr, 2).value
        txt = norm_text(f"{codigo or ''} {concepto or ''} {ws.cell(rr, 3).value or ''}")
        if not txt:
            continue
        market_op = str(ws.cell(rr, 11).value or "").strip().lower()
        market_pu = _v033_money(ws.cell(rr, 10).value)
        market_imp = _v033_money(ws.cell(rr, 13).value)
        pu = _v033_money(ws.cell(rr, 4).value)
        imp = _v033_money(ws.cell(rr, 7).value)
        if market_op == "sin match" or (market_pu is None and market_imp is None):
            signals["no_match"] += 1
        if pu is not None and market_pu not in (None, 0):
            diff = (pu - market_pu) / market_pu
            if abs(diff) >= 0.05:
                signals["elevated_market"].append((abs(diff), diff, _v037_short_text(concepto or codigo, 55)))
        elif imp is not None and market_imp not in (None, 0):
            diff = (imp - market_imp) / market_imp
            if abs(diff) >= 0.05:
                signals["elevated_market"].append((abs(diff), diff, _v037_short_text(concepto or codigo, 55)))
        if any(k in txt for k in ["supervisor", "oficial", "ayudante", "cuadrilla", "soldador", "tubero", "electrico", "eléctrico", "mano de obra"]):
            signals["labor"] = True
        if any(k in txt for k in ["epp", "seguridad", "altura", "permiso", "maniobra", "izaje", "soldadura", "corte"]):
            signals["ehs"] = True
        if any(k in txt for k in ["montacargas", "grua", "grúa", "manlift", "plataforma", "camion", "camión", "camioneta", "transporte", "acarreo", "flete", "traslado", "combustible", "operador", "seguro", "renta"]):
            signals["equipment"] = True
        if any(k in txt for k in ["inoxidable", "304", "316", "sanitario", "limpieza", "alimentaria", "cip", "valvula", "válvula", "pulido", "prueba"]):
            signals["quality"] = True
        if any(k in txt for k in ["indirecto", "financiamiento", "utilidad", "cargo adicional", "herramienta menor"]):
            signals["financial"] = True
        if any(k in txt for k in ["lote", "global", "paquete"]):
            signals["generic"] += 1
    signals["elevated_market"] = sorted(signals["elevated_market"], reverse=True, key=lambda x: x[0])[:5]
    return signals


def _v038_build_ai_text(wb) -> List[Tuple[str, str]]:
    if "Comparativa" not in wb.sheetnames:
        return [("Veredicto general", "No se encontró la hoja Comparativa para construir el análisis.")]
    comp = wb["Comparativa"]
    total_row = _v038_find_total_row(comp)
    providers_info = _v038_detect_providers_from_comparativa(comp)
    metrics = []
    for pinfo in providers_info:
        metric = _v038_comparativa_metrics(comp, pinfo, total_row)
        dname = _v038_detail_sheet_name(wb, pinfo, len(providers_info))
        metric["detail_sheet"] = dname
        metric["signals"] = _v038_detail_signals(wb[dname]) if dname else {}
        metrics.append(metric)

    if not metrics:
        return [("Veredicto general", "No se detectaron contratistas/proveedores para análisis.")]

    best = min(metrics, key=lambda m: (m.get("ajuste") if m.get("ajuste") is not None else 999, m.get("total") or 0))
    lowest = min(metrics, key=lambda m: m.get("total") or 0)
    highest_risk = max(metrics, key=lambda m: ((m.get("ajuste") or 0) + 0.03 * (m.get("signals") or {}).get("no_match", 0)))

    if len(metrics) == 1:
        m = metrics[0]
        if m.get("ajuste") is None:
            verdict = f"{m['name']} fue analizado contra mercado, pero la referencia disponible no permite calcular un ajuste global confiable. Conviene revisar cobertura de matches y soporte de matriz."
        elif m["ajuste"] > 0.05:
            verdict = f"{m['name']} se encuentra {_v037_pct_label(m['ajuste'])} por arriba del mercado calculado. La recomendación es revisar partidas críticas, soporte de precios y alcance antes de negociar."
        elif m["ajuste"] < -0.05:
            verdict = f"{m['name']} se encuentra {_v037_pct_label(abs(m['ajuste']))} por debajo del mercado calculado. Esto puede representar oportunidad, pero requiere validar alcance, exclusiones y suficiencia técnica."
        else:
            verdict = f"{m['name']} se mantiene razonablemente alineado contra mercado con la información disponible."
    else:
        verdict = f"{best['name']} presenta el mejor balance económico frente al mercado disponible. {highest_risk['name']} concentra el mayor riesgo relativo por desviación, cobertura o señales técnicas detectadas. {lowest['name']} tiene el menor total cotizado, sujeto a validar alcance y exclusiones."

    comparison_lines = []
    for m in metrics:
        adj = _v037_pct_label(m.get("ajuste")) if m.get("ajuste") is not None else "N/D"
        top = "; ".join([f"{r['partida']} ({_v037_money_label(r['importe'])})" for r in m.get("top_rows", [])[:3]]) or "sin partidas relevantes detectadas"
        comparison_lines.append(f"• {m['name']}: total {_v037_money_label(m.get('total'))}, mercado {_v037_money_label(m.get('mercado'))}, ajuste global {adj}. Principales partidas: {top}.")

    risk_lines: List[str] = []
    rec_lines: List[str] = []
    limitation_lines: List[str] = []
    for m in metrics:
        sig = m.get("signals") or {}
        if sig.get("no_match"):
            limitation_lines.append(f"• {m['name']}: {sig.get('no_match')} renglones sin match de mercado; revisar manualmente antes de tomar decisión final.")
        if sig.get("elevated_market"):
            labels = "; ".join([x[2] for x in sig.get("elevated_market", [])[:3]])
            risk_lines.append(f"• {m['name']}: diferencias relevantes contra mercado en {labels}.")
        if sig.get("equipment"):
            risk_lines.append(f"• {m['name']}: validar equipo móvil/logística, incluyendo operador, combustible, seguro, traslado y jornada mínima cuando aplique.")
        if sig.get("ehs"):
            risk_lines.append(f"• {m['name']}: validar EPP, supervisión de seguridad, permisos y controles de trabajo en sitio.")
        if sig.get("quality"):
            risk_lines.append(f"• {m['name']}: validar calidad industrial/sanitaria, limpieza, pruebas y criterios de aceptación de planta.")
        if sig.get("financial"):
            risk_lines.append(f"• {m['name']}: revisar indirectos, herramienta menor, EPP, financiamiento o cargos adicionales para evitar duplicidades.")
        if sig.get("generic"):
            limitation_lines.append(f"• {m['name']}: existen conceptos globales/LOTE; solicitar apertura adicional si son partidas de alto impacto.")
    if not risk_lines:
        risk_lines.append("• No se detectaron riesgos técnicos concluyentes con la información disponible; revisar los detalles para confirmar cobertura y alcance.")
    rec_lines.append(f"• Usar {best['name']} como referencia competitiva inicial, siempre que el alcance técnico esté completo.")
    if highest_risk['name'] != best['name']:
        rec_lines.append(f"• Solicitar a {highest_risk['name']} soporte de precios y apertura de matriz en partidas críticas antes de negociar/adjudicar.")
    rec_lines.append("• Validar partidas bajo mercado para descartar omisiones de alcance y partidas sobre mercado para soportar negociación.")
    rec_lines.append("• Revisar seguridad, logística, calidad industrial y condiciones de entrega cuando aparezcan señales en la matriz.")
    if not limitation_lines:
        limitation_lines.append("• El análisis depende de la calidad de las matrices recibidas y de la cobertura Construdata disponible; cualquier partida sin match confiable debe revisarse manualmente.")

    return [
        ("1. Veredicto general", verdict),
        ("2. Comparación entre contratistas", "\n".join(comparison_lines)),
        ("3. Principales riesgos detectados", "\n".join(risk_lines)),
        ("4. Recomendaciones de negociación / validación", "\n".join(rec_lines)),
        ("5. Limitaciones del análisis", "\n".join(limitation_lines)),
    ]


def _v038_create_analisis_ia_sheet(wb) -> None:
    if "Analisis IA" in wb.sheetnames:
        del wb["Analisis IA"]
    ws = wb.create_sheet("Analisis IA")
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    title = ws.cell(1, 1, "Analisis IA")
    title.fill = PatternFill("solid", fgColor=V038_AI_TITLE_FILL)
    title.font = Font(bold=True, color="FFFFFF", size=13)
    title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    row = 3
    for section, text in _v038_build_ai_text(wb):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cell = ws.cell(row, 1, section)
        cell.fill = PatternFill("solid", fgColor=V038_AI_SECTION_FILL)
        cell.font = Font(bold=True, color=TEXT, size=10)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        body = ws.cell(row, 1, text)
        body.font = Font(color=TEXT, size=10)
        body.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        # Altura proporcional simple para lectura.
        lines = max(2, str(text or "").count("\n") + 2)
        ws.row_dimensions[row].height = min(150, 18 * lines)
        row += 2

    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 18 if col > 1 else 26
    ws.freeze_panes = "A3"


def _v034_append_detail_sheets(wb, workbook_path: str, analysis: Dict[str, Any], detail_cache: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> str:
    meta = analysis.get("meta") or {}
    filepaths = list(meta.get("source_filepaths") or [])
    nombres = list(meta.get("source_provider_names") or analysis.get("providers") or [])
    resumen_paths = list(meta.get("source_resumen_paths") or [])
    providers = list(analysis.get("providers") or nombres or [])
    if not filepaths:
        # Mantener etapa legible aunque no haya archivos fuente en meta.
        _v034_create_detail_sheet(wb, "Detalle", [])
    else:
        for idx, fp in enumerate(filepaths):
            pname = nombres[idx] if idx < len(nombres) and nombres[idx] else (providers[idx] if idx < len(providers) else f"P{idx+1}")
            rp = resumen_paths[idx] if idx < len(resumen_paths) else None
            try:
                cache_key = providers[idx] if idx < len(providers) else pname
                if detail_cache and cache_key in detail_cache:
                    detail_rows = detail_cache.get(cache_key) or []
                elif detail_cache and pname in detail_cache:
                    detail_rows = detail_cache.get(pname) or []
                else:
                    detail_rows, _diag = _v034_extract_provider_detail_rows(fp, rp, pname)
            except Exception as exc:
                detail_rows = [{"kind": "concept", "codigo": "ERROR", "concepto": f"No se pudo generar detalle: {type(exc).__name__}: {exc}", "unidad": "", "pu": None, "op": "", "cantidad": None, "importe": None, "pct": None, "market_pu": None, "market_op": "Sin match", "market_qty": None, "market_importe": None, "market_match_real": False}]
            sheet_name = "Detalle" if len(filepaths) == 1 else f"Detalle - P{idx+1}"
            _v034_create_detail_sheet(wb, sheet_name, detail_rows)

    # V0.3.8: nuevo tab Analisis IA con veredicto consolidado.
    try:
        _v038_create_analisis_ia_sheet(wb)
    except Exception as exc:
        # No romper el Excel si el texto IA falla; dejar diagnóstico controlado.
        if "Analisis IA" in wb.sheetnames:
            del wb["Analisis IA"]
        ws_ai = wb.create_sheet("Analisis IA")
        ws_ai.cell(1, 1, "Analisis IA")
        ws_ai.cell(3, 1, f"No se pudo generar el análisis IA: {type(exc).__name__}: {exc}")

    # Defensa: solo Comparativa + Detalle(s) + Analisis IA.
    allowed = {"Comparativa", "Analisis IA"}
    for sh in wb.sheetnames:
        if sh == "Detalle" or sh.startswith("Detalle - P"):
            allowed.add(sh)
    for sh in list(wb.sheetnames):
        if sh not in allowed:
            del wb[sh]
    Path(workbook_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(workbook_path)
    return workbook_path


def create_professional_mvp_workbook(workbook_path: str, filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> str:
    """V0.3.4: salida etapa 2, Comparativa + Detalle por contratista."""
    meta = dict(meta or {})
    meta["source_filepaths"] = list(filepaths or [])
    meta["source_provider_names"] = list(nombres or [])
    meta["source_resumen_paths"] = list(resumen_paths or [])
    analysis = build_professional_mvp_analysis(
        list(filepaths or []),
        list(nombres or []),
        meta,
        catalogo_path,
        nacional_path,
        list(resumen_paths or []),
        resumen_oficial_path,
    )
    return write_comparativa_stage2_workbook(workbook_path, analysis)


def append_professional_mvp_workbook(workbook_path: str, filepaths: List[str], nombres: List[str], meta: Optional[Dict[str, Any]] = None, catalogo_path: Optional[str] = None, nacional_path: Optional[str] = None, resumen_paths: Optional[List[Optional[str]]] = None, resumen_oficial_path: Optional[str] = None) -> str:
    """V0.3.4: compatibilidad; reconstruye Comparativa + Detalle(s)."""
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


# ===========================================================================
# V0.3.12 - Detalle agrupado por familias + operador real de mercado
# ===========================================================================
# Objetivo:
# - El Detalle de contratistas debe mostrar cada servicio como matriz APU:
#   Materiales, Mano de Obra, Equipo y Herramienta, Basicos, Otros, subtotales
#   por familia y seccion financiera.
# - Mercado Importe se recalcula respetando el operador real (*, /, %), en vez
#   de confiar en importes preexistentes que pueden venir mal calculados.

V0312_FAMILY_ORDER = ["MATERIALES", "MANO DE OBRA", "EQUIPO Y HERRAMIENTA", "BASICOS", "OTROS"]
V0312_FAMILY_FILL = "F2F2F2"
V0312_SERVICE_FILL = "E7E6E6"
V0312_FIN_FILL = "D9EAF7"
V0312_TOTAL_FILL = "B4C6E7"


def _v0312_norm_operator(value: Any, default: str = "*") -> str:
    raw = str(value or "").strip()
    txt = norm_text(raw)
    if raw in {"*", "x", "X"} or txt in {"x", "multiplicacion", "multiplicar", "producto"}:
        return "*"
    if raw == "/" or txt in {"division", "dividir", "rendimiento"}:
        return "/"
    if raw == "%" or "%" in raw or txt in {"porcentaje", "porcentual"}:
        return "%"
    return default


def _v0312_family_label(value: Any = None, text: Any = None) -> str:
    raw = f"{value or ''} {text or ''}"
    txt = norm_text(raw)
    text_only = norm_text(text)
    # Si la descripcion del renglon dice material/consumible, priorizar Materiales
    # aunque la matriz origen venga etiquetada de forma imprecisa.
    if any(k in text_only for k in ["material", "acero", "valvula", "válvula", "tuberia", "tubería", "tubo", "brida", "tornill", "soldadura", "consumible", "lamina", "lámina", "cople", "conduit"]):
        return "MATERIALES"
    # Luego roles laborales explícitos para evitar clasificar "tubero" como tuberia.
    if any(k in txt for k in ["mano de obra", "mo", "jornal", "jor", "oficial", "ayudante", "soldador", "tubero", "electric", "supervisor", "cuadrilla", "seguridad"]):
        return "MANO DE OBRA"
    if any(k in txt for k in ["material", "acero", "valvula", "válvula", "tuberia", "tubería", "tubo", "brida", "tornill", "soldadura", "consumible", "lamina", "lámina", "cople", "conduit"]):
        return "MATERIALES"
    if any(k in txt for k in ["equipo", "herramient", "maquinaria", "montacargas", "grua", "grúa", "manlift", "plataforma", "camion", "camión", "camioneta", "flete", "transporte", "acarreo", "renta"]):
        return "EQUIPO Y HERRAMIENTA"
    if any(k in txt for k in ["basico", "basica", "básico", "básica", "matriz auxiliar", "auxiliar"]):
        return "BASICOS"
    if norm_text(value) in {"materiales", "material"}:
        return "MATERIALES"
    return "OTROS"


def _v0312_calc_amount(pu_val: Any, op_val: Any, qty_val: Any, base_val: Any = None) -> Optional[float]:
    pu = _v033_num(pu_val)
    qty = _v033_num(qty_val)
    base = _v033_num(base_val)
    op = _v0312_norm_operator(op_val)
    if op == "%":
        if base is None or qty is None:
            return None
        pct = qty / 100.0 if abs(qty) > 1 else qty
        return _v033_money(base * pct)
    if pu is None:
        return None
    if op == "/":
        if qty in (None, 0):
            return None
        return _v033_money(pu / qty)
    if qty is None:
        return None
    return _v033_money(pu * qty)


# Reemplaza calculo base de mercado: ya no confia primero en importe_mercado
# recibido; recalcula si tiene PU, cantidad y operador.
def _v034_calc_market_importe(pu_val: Optional[float], op_val: Any, qty_val: Optional[float], base_val: Optional[float] = None) -> Optional[float]:
    return _v0312_calc_amount(pu_val, op_val, qty_val, base_val)


def _v035_real_market_fields(item: Dict[str, Any]) -> Dict[str, Any]:
    if not _v034_is_real_market_match(item):
        return {"market_pu": None, "market_op": "Sin match", "market_qty": None, "market_importe": None, "market_match_real": False}
    market_pu = _v034_item_market_value(item, "precio_mercado", "market_pu", "mercado_pu", "precio_nacional")
    market_qty = _v034_item_market_value(item, "cantidad_mercado", "market_qty", "mercado_cantidad", "cantidad", "factor", "qty")
    market_op = _v0312_norm_operator(item.get("market_op") or item.get("mercado_op") or item.get("op_mercado") or item.get("op") or item.get("operacion") or "*")
    base_val = item.get("base_mercado") or item.get("base") or item.get("costo_directo")
    market_importe = _v0312_calc_amount(market_pu, market_op, market_qty, base_val)
    # Fallback solo si no se puede calcular con operador; nunca usa importe del proveedor.
    if market_importe is None:
        market_importe = _v034_item_market_value(item, "importe_mercado", "market_importe", "mercado_importe")
    return {"market_pu": market_pu, "market_op": market_op, "market_qty": market_qty, "market_importe": market_importe, "market_match_real": True}


def _v0312_enrich_detail_rows(detail_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Agrega familia y normaliza operadores/importes de mercado sin alterar columnas."""
    current_concept = ""
    for row in detail_rows:
        kind = str(row.get("kind") or "item")
        if kind == "concept":
            current_concept = str(row.get("concepto") or "")
            row["family"] = "SERVICIO"
            continue
        if kind == "financial":
            row["family"] = "FINANCIERO"
            row["op"] = _v0312_norm_operator(row.get("op"), default="") if row.get("op") else ""
            row["market_op"] = _v0312_norm_operator(row.get("market_op"), default="") if row.get("market_op") else ""
            continue
        fam = row.get("family") or row.get("familia") or row.get("tipo") or row.get("categoria")
        row["family"] = _v0312_family_label(fam, f"{row.get('codigo') or ''} {row.get('concepto') or ''} {row.get('unidad') or ''}")
        row["op"] = _v0312_norm_operator(row.get("op") or row.get("operacion") or "*")
        if row.get("importe") is None:
            row["importe"] = _v0312_calc_amount(row.get("pu"), row.get("op"), row.get("cantidad"), row.get("base"))
        if row.get("market_match_real"):
            row["market_op"] = _v0312_norm_operator(row.get("market_op") or row.get("op") or "*")
            row["market_importe"] = _v0312_calc_amount(row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("base_mercado") or row.get("base"))
        elif row.get("market_op") not in (None, "", "Sin match"):
            row["market_op"] = _v0312_norm_operator(row.get("market_op"))
    return detail_rows


_v0312_previous_extract_provider_detail_rows = _v034_extract_provider_detail_rows


def _v034_extract_provider_detail_rows(provider_file: str, resumen_file: Optional[str], provider_name: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    rows, diag = _v0312_previous_extract_provider_detail_rows(provider_file, resumen_file, provider_name)
    return _v0312_enrich_detail_rows(rows), diag


def _v0312_write_values(ws, rr: int, vals: List[Any], cols: int = 13) -> None:
    for c in range(1, cols + 1):
        ws.cell(rr, c, vals[c - 1] if c <= len(vals) else None)


def _v0312_style_detail_row(ws, rr: int, kind: str = "item", cols: int = 13) -> None:
    fill = None
    bold = False
    if kind == "service":
        fill = V0312_SERVICE_FILL; bold = True
    elif kind == "family":
        fill = V0312_FAMILY_FILL; bold = True
    elif kind == "subtotal":
        fill = V0312_FAMILY_FILL; bold = True
    elif kind == "financial":
        fill = V0312_FIN_FILL; bold = True
    elif kind == "total":
        fill = V0312_TOTAL_FILL; bold = True
    for c in range(1, cols + 1):
        cell = ws.cell(rr, c)
        cell.border = Border(bottom=Side(style="hair", color="D9E2F3"))
        cell.font = Font(name="Calibri", size=9, color=TEXT, bold=bold)
        cell.alignment = Alignment(vertical="center", wrap_text=(c == 2))
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
    # number formats for contractor detail columns
    if cols >= 13:
        for col in (4, 7, 10, 13):
            ws.cell(rr, col).number_format = '$#,##0.00'
        for col in (6, 12):
            ws.cell(rr, col).number_format = '#,##0.0000'
        ws.cell(rr, 8).number_format = '0.00%'


def _v0312_subtotal_row(label: str, rows: List[Dict[str, Any]], provider_total: float) -> Dict[str, Any]:
    imp = sum(float(_v033_money(r.get("importe")) or 0.0) for r in rows)
    mimp = sum(float(_v033_money(r.get("market_importe")) or 0.0) for r in rows if r.get("market_match_real"))
    has_mkt = any(r.get("market_match_real") and _v033_money(r.get("market_importe")) is not None for r in rows)
    return {
        "kind": "subtotal", "codigo": "", "concepto": f"SUBTOTAL {label}", "unidad": "",
        "pu": imp, "op": "", "cantidad": None, "importe": imp,
        "pct": (imp / provider_total) if provider_total else None,
        "market_pu": mimp if has_mkt else None, "market_op": "", "market_qty": None,
        "market_importe": mimp if has_mkt else None, "market_match_real": has_mkt,
    }


def _v034_create_detail_sheet(wb, sheet_name: str, detail_rows: List[Dict[str, Any]]) -> None:
    """V0.3.12: Detalle de contratista agrupado por servicio/familia.

    Mantiene las 13 columnas aprobadas para contratistas, pero organiza cada
    matriz como PU/APU: servicio, familias, renglones, subtotal por familia y
    bloque financiero. La matriz base usa otro writer en base_matrix_proposer.
    """
    detail_rows = _v0312_enrich_detail_rows(detail_rows or [])
    ws = wb.create_sheet(_safe_sheet_title(sheet_name, wb.sheetnames))
    headers = ["Código", "Concepto", "Unidad", "P. Unitario", "Op.", "Cantidad", "Importe", "%", "", "Mercado P. Unitario", "Mercado Op.", "Mercado Cantidad", "Mercado Importe"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    r = 2
    row_kinds: Dict[int, str] = {}
    provider_total = sum(float(_v033_money(x.get("importe")) or 0.0) for x in detail_rows if str(x.get("kind")) == "concept")
    if not provider_total:
        provider_total = sum(float(_v033_money(x.get("importe")) or 0.0) for x in detail_rows if str(x.get("kind")) == "item")

    current_concept: Optional[Dict[str, Any]] = None
    current_items: List[Dict[str, Any]] = []
    current_financial: List[Dict[str, Any]] = []

    def flush_current() -> None:
        nonlocal r, current_concept, current_items, current_financial
        if current_concept is None:
            return
        # Servicio / matriz principal
        vals = [
            current_concept.get("codigo"), current_concept.get("concepto"), current_concept.get("unidad"), current_concept.get("pu"), current_concept.get("op"),
            current_concept.get("cantidad"), current_concept.get("importe"), current_concept.get("pct"), "",
            current_concept.get("market_pu"), current_concept.get("market_op"), current_concept.get("market_qty"), current_concept.get("market_importe"),
        ]
        _v0312_write_values(ws, r, vals); row_kinds[r] = "service"; r += 1
        # Familias + subtotales
        for fam in V0312_FAMILY_ORDER:
            fam_rows = [x for x in current_items if x.get("family") == fam]
            if not fam_rows:
                continue
            _v0312_write_values(ws, r, ["", fam, "", "", "", "", "", "", "", "", "", "", ""]); row_kinds[r] = "family"; r += 1
            for row in fam_rows:
                # Recalculo defensivo de Mercado Importe con operador real.
                if row.get("market_match_real"):
                    row["market_importe"] = _v0312_calc_amount(row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("base_mercado") or row.get("base"))
                vals = [row.get("codigo"), row.get("concepto"), row.get("unidad"), row.get("pu"), row.get("op"), row.get("cantidad"), row.get("importe"), row.get("pct"), "", row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("market_importe")]
                _v0312_write_values(ws, r, vals); row_kinds[r] = "item"; r += 1
            st = _v0312_subtotal_row(fam, fam_rows, provider_total)
            vals = [st.get("codigo"), st.get("concepto"), st.get("unidad"), st.get("pu"), st.get("op"), st.get("cantidad"), st.get("importe"), st.get("pct"), "", st.get("market_pu"), st.get("market_op"), st.get("market_qty"), st.get("market_importe")]
            _v0312_write_values(ws, r, vals); row_kinds[r] = "subtotal"; r += 1
        # Seccion financiera
        if current_financial:
            _v0312_write_values(ws, r, ["", "SECCION FINANCIERA", "", "", "", "", "", "", "", "", "", "", ""]); row_kinds[r] = "financial"; r += 1
            for row in current_financial:
                vals = [row.get("codigo"), row.get("concepto"), row.get("unidad"), row.get("pu"), row.get("op"), row.get("cantidad"), row.get("importe"), row.get("pct"), "", row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("market_importe")]
                _v0312_write_values(ws, r, vals); row_kinds[r] = "financial"; r += 1
        # Total costo unitario por servicio: visible aunque no haya financieros.
        _v0312_write_values(ws, r, ["", "TOTAL POR SERVICIO", "", current_concept.get("pu"), "", current_concept.get("cantidad"), current_concept.get("importe"), current_concept.get("pct"), "", current_concept.get("market_pu"), "", current_concept.get("market_qty"), current_concept.get("market_importe")])
        row_kinds[r] = "total"; r += 2
        current_concept = None; current_items = []; current_financial = []

    for row in detail_rows:
        kind = str(row.get("kind") or "item")
        if kind == "concept":
            flush_current()
            current_concept = row
        elif kind == "financial":
            current_financial.append(row)
        else:
            current_items.append(row)
    flush_current()

    if r == 2:
        _v0312_write_values(ws, 2, ["Sin detalle detectado", "No fue posible leer renglones de matriz para este contratista."])
        row_kinds[2] = "service"; r = 3
    last_row = r - 1
    _v034_style_detail_sheet(ws, last_row)
    for rr, kind in row_kinds.items():
        _v0312_style_detail_row(ws, rr, kind, 13)
        if kind == "item":
            _v036_apply_market_diff_bold(ws, rr)
    try:
        ws.auto_filter.ref = f"A1:M{last_row}"
    except Exception:
        pass


# ===========================================================================
# V0.3.14 - Detalle: sin match heredado en italica + operador real proveedor
# ===========================================================================
# Objetivo:
# - En Detalle de contratistas, cuando no exista match real contra Construdata,
#   las columnas de mercado se llenan con los valores del contratista y se
#   pintan en italica para indicar que son valores heredados, no mercado real.
# - La columna Op. debe respetar la operacion real declarada/deducida del
#   renglon: *, / o %, sin forzar siempre *.
# - Matriz base no cambia: no tiene columnas de mercado ni Analisis IA.


def _v0314_operator_from_row(row: Dict[str, Any], default: str = "*") -> str:
    """Devuelve operador normalizado desde la evidencia real del renglon.

    Prioridad:
    1) flags/metadata de formula: porcentaje_mo => %, rendimiento/dividir => /
    2) columnas explicitas de operacion: op/operacion/operator/etc.
    3) fallback controlado: * solo si no hay evidencia contraria.
    """
    formula_kind = norm_text(row.get("formula_kind") or row.get("tipo_formula") or row.get("calculation_type") or "")
    if any(k in formula_kind for k in ["porcentaje", "percent", "pct", "porcentaje_mo"]):
        return "%"
    if any(k in formula_kind for k in ["rendimiento", "inverso", "division", "dividir"]):
        return "/"
    dividir = row.get("dividir") or row.get("is_yield") or row.get("rendimiento_inverso")
    if isinstance(dividir, str):
        if norm_text(dividir) in {"si", "sí", "true", "1", "x", "dividir"}:
            return "/"
    elif bool(dividir):
        return "/"
    for key in ("op", "operacion", "operator", "operador", "signo", "simbolo", "calculo", "expresion"):
        val = row.get(key)
        if val not in (None, ""):
            return _v0312_norm_operator(val, default=default)
    return default


def _v0314_apply_no_match_fallback(row: Dict[str, Any]) -> Dict[str, Any]:
    """Si no hay match real, completa mercado visible con valores proveedor.

    No cambia market_match_real: sigue False para que Comparativa y calculos no
    lo traten como referencia Construdata real. Solo mejora lectura visual.
    """
    kind = str(row.get("kind") or "item")
    if kind not in {"item"}:
        return row
    row["op"] = _v0314_operator_from_row(row, default="*")
    if row.get("importe") is None:
        row["importe"] = _v0312_calc_amount(row.get("pu"), row.get("op"), row.get("cantidad"), row.get("base"))
    if not row.get("market_match_real"):
        row["market_pu"] = row.get("pu")
        row["market_op"] = row.get("op")
        row["market_qty"] = row.get("cantidad")
        row["market_importe"] = row.get("importe")
        row["market_fallback_from_contractor"] = True
        row["sin_match_mercado"] = True
        row["estado"] = row.get("estado") or "sin_match_mercado"
    else:
        row["market_fallback_from_contractor"] = False
        row["market_op"] = _v0314_operator_from_row({**row, "op": row.get("market_op") or row.get("op")}, default=row.get("op") or "*")
        row["market_importe"] = _v0312_calc_amount(row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("base_mercado") or row.get("base"))
        if row.get("market_importe") is None:
            row["market_importe"] = _v034_item_market_value(row, "importe_mercado", "market_importe", "mercado_importe")
    return row


def _v0312_enrich_detail_rows(detail_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """V0.3.14 override: familia + operador real + fallback visible sin match."""
    current_concept = ""
    for row in detail_rows:
        kind = str(row.get("kind") or "item")
        if kind == "concept":
            current_concept = str(row.get("concepto") or "")
            row["family"] = "SERVICIO"
            # Si el servicio no trae operador explicito, no forzar nada en el
            # encabezado de servicio; el detalle de insumos lleva el operador.
            row["op"] = _v0312_norm_operator(row.get("op"), default="") if row.get("op") else ""
            continue
        if kind == "financial":
            row["family"] = "FINANCIERO"
            row["op"] = _v0314_operator_from_row(row, default="") if row.get("op") or row.get("formula_kind") or row.get("dividir") else ""
            row["market_op"] = _v0312_norm_operator(row.get("market_op"), default="") if row.get("market_op") else ""
            continue
        fam = row.get("family") or row.get("familia") or row.get("tipo") or row.get("categoria")
        row["family"] = _v0312_family_label(fam, f"{row.get('codigo') or ''} {row.get('concepto') or ''} {row.get('unidad') or ''}")
        _v0314_apply_no_match_fallback(row)
    return detail_rows


# Wrap extractor to preserve formula metadata where present in raw items.
_v0314_previous_extract_provider_detail_rows = _v034_extract_provider_detail_rows


def _v034_extract_provider_detail_rows(provider_file: str, resumen_file: Optional[str], provider_name: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    rows, diag = _v0314_previous_extract_provider_detail_rows(provider_file, resumen_file, provider_name)
    return _v0312_enrich_detail_rows(rows), diag


def _v0314_apply_market_fallback_italic(ws, rr: int) -> None:
    """Aplica italica solo a columnas mercado J:M para valores heredados."""
    for col in (10, 11, 12, 13):
        cell = ws.cell(rr, col)
        base = cell.font or Font(name="Calibri", size=9, color=TEXT)
        cell.font = Font(
            name=base.name or "Calibri",
            size=base.sz or 9,
            bold=False,
            italic=True,
            color=TEXT,
        )


def _v034_create_detail_sheet(wb, sheet_name: str, detail_rows: List[Dict[str, Any]]) -> None:
    """V0.3.14: Detalle contratista con fallback visible e italica.

    Mantiene columnas aprobadas. Los valores heredados por sin match se muestran
    en mercado con italica y no se consideran mercado real.
    """
    detail_rows = _v0312_enrich_detail_rows(detail_rows or [])
    ws = wb.create_sheet(_safe_sheet_title(sheet_name, wb.sheetnames))
    headers = ["Código", "Concepto", "Unidad", "P. Unitario", "Op.", "Cantidad", "Importe", "%", "", "Mercado P. Unitario", "Mercado Op.", "Mercado Cantidad", "Mercado Importe"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    r = 2
    row_kinds: Dict[int, str] = {}
    row_fallbacks: Dict[int, bool] = {}
    provider_total = sum(float(_v033_money(x.get("importe")) or 0.0) for x in detail_rows if str(x.get("kind")) == "concept")
    if not provider_total:
        provider_total = sum(float(_v033_money(x.get("importe")) or 0.0) for x in detail_rows if str(x.get("kind")) == "item")

    current_concept: Optional[Dict[str, Any]] = None
    current_items: List[Dict[str, Any]] = []
    current_financial: List[Dict[str, Any]] = []

    def write_row(row: Dict[str, Any], kind: str) -> None:
        nonlocal r
        vals = [
            row.get("codigo"), row.get("concepto"), row.get("unidad"), row.get("pu"), row.get("op"), row.get("cantidad"), row.get("importe"), row.get("pct"), "",
            row.get("market_pu"), row.get("market_op"), row.get("market_qty"), row.get("market_importe"),
        ]
        _v0312_write_values(ws, r, vals)
        row_kinds[r] = kind
        row_fallbacks[r] = bool(row.get("market_fallback_from_contractor"))
        r += 1

    def flush_current() -> None:
        nonlocal r, current_concept, current_items, current_financial
        if current_concept is None:
            return
        write_row(current_concept, "service")
        for fam in V0312_FAMILY_ORDER:
            fam_rows = [x for x in current_items if x.get("family") == fam]
            if not fam_rows:
                continue
            _v0312_write_values(ws, r, ["", fam, "", "", "", "", "", "", "", "", "", "", ""]); row_kinds[r] = "family"; r += 1
            for row in fam_rows:
                _v0314_apply_no_match_fallback(row)
                write_row(row, "item")
            st = _v0312_subtotal_row(fam, fam_rows, provider_total)
            write_row(st, "subtotal")
        if current_financial:
            _v0312_write_values(ws, r, ["", "SECCION FINANCIERA", "", "", "", "", "", "", "", "", "", "", ""]); row_kinds[r] = "financial"; r += 1
            for row in current_financial:
                write_row(row, "financial")
        _v0312_write_values(ws, r, ["", "TOTAL POR SERVICIO", "", current_concept.get("pu"), "", current_concept.get("cantidad"), current_concept.get("importe"), current_concept.get("pct"), "", current_concept.get("market_pu"), "", current_concept.get("market_qty"), current_concept.get("market_importe")])
        row_kinds[r] = "total"; r += 2
        current_concept = None; current_items = []; current_financial = []

    for row in detail_rows:
        kind = str(row.get("kind") or "item")
        if kind == "concept":
            flush_current()
            current_concept = row
        elif kind == "financial":
            current_financial.append(row)
        else:
            current_items.append(row)
    flush_current()

    if r == 2:
        _v0312_write_values(ws, 2, ["Sin detalle detectado", "No fue posible leer renglones de matriz para este contratista."])
        row_kinds[2] = "service"; r = 3
    last_row = r - 1
    _v034_style_detail_sheet(ws, last_row)
    for rr, kind in row_kinds.items():
        _v0312_style_detail_row(ws, rr, kind, 13)
        if kind == "item":
            if row_fallbacks.get(rr):
                _v0314_apply_market_fallback_italic(ws, rr)
            else:
                _v036_apply_market_diff_bold(ws, rr)
    try:
        ws.auto_filter.ref = f"A1:M{last_row}"
    except Exception:
        pass

