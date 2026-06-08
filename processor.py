"""
Quantia AI Comparador
"""

"""
processor.py — Motor de extracción y comparación de cotizaciones

Estrategia de detección automática de formato:
  1. Busca filas con código COM-XXX-XX y precio unitario numérico (formato PPT/RAR)
  2. Busca patrón "Clave:" + "Precio unitario:" en filas separadas (formato CIOC APU)
  3. Fallback: busca columnas con headers conocidos (Descripción, P.U., Importe)
"""

import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
from pathlib import Path
import re
import json
from difflib import SequenceMatcher
from material_market_search import select_material_market_price, optimize_material_search_scope
from material_vector_index import vector_search_candidates, ensure_construdata_vector_index, get_vector_index_status
from material_match_cache import get_cached_material_match, save_material_match, cache_status
from concept_ai_matcher import get_concept_ai_status, voyage_rerank_concept_candidates, claude_decide_concept_match


_BENCHMARK_CACHE = {}


def _quantia_data_dir() -> Path:
    """Directorio unico para catalogos/base Neodata-Construdata.

    Orden de busqueda profesional:
    1) DATA_DIR / QUANTIA_DATA_DIR si se define en Docker/Railway/local.
    2) data/ empacado junto al proyecto para pruebas locales.
    3) ./data como fallback defensivo.
    """
    for key in ("DATA_DIR", "QUANTIA_DATA_DIR"):
        raw = os.getenv(key)
        if raw:
            path = Path(raw).expanduser()
            if path.exists():
                return path
    packaged = Path(__file__).resolve().parent / "data"
    if packaged.exists():
        return packaged
    return Path("data")


def _quantia_data_path(*parts) -> Path:
    return _quantia_data_dir().joinpath(*parts)


def _base_clave(clave):
    if clave is None:
        return ""
    s = str(clave).strip()
    return s.split("::")[-1]

# ──────────────────────────────────────────────────────────────
# TEMPLATE V1 (multi-hoja proveedores / catálogo)
# ──────────────────────────────────────────────────────────────

def _norm(x):
    if x is None: return ""
    return str(x).strip().lower()

def _clean_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s

def _row_text(row):
    return " ".join(_clean_text(x).lower() for x in row if _clean_text(x))

def _as_float(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            if value != value:
                return None
        except Exception:
            pass
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(",", "")
        if not s:
            return None
        try:
            return float(s)
        except Exception:
            return None
    return None

def _find_last_numeric(row):
    nums = [_as_float(x) for x in row]
    nums = [n for n in nums if n is not None]
    return nums[-1] if nums else None

def _find_row_value_by_label(rows, start, end, labels):
    labels = [l.lower() for l in labels]
    for i in range(start, min(end, len(rows))):
        txt = _row_text(rows[i])
        if any(label in txt for label in labels):
            val = _find_last_numeric(rows[i])
            if val is not None:
                return val
    return None


def _find_row_importe_by_label(rows, start, end, labels):
    labels = [l.lower() for l in labels]
    for i in range(start, min(end, len(rows))):
        txt = _row_text(rows[i])
        if any(label in txt for label in labels):
            row = rows[i]
            preferred = _as_float(row[6] if len(row) > 6 else None)
            if preferred is not None:
                return preferred
            nums = [_as_float(x) for x in row]
            nums = [n for n in nums if n is not None]
            if len(nums) >= 2:
                return nums[-2]
            if nums:
                return nums[-1]
    return None

def _normalize_text(value):
    if value is None:
        return ""
    s = str(value).strip().lower()
    repl = {
        "á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n",
        "°":" grados ","/":" ","-":" ","(":" ",")":" ",",":" ",":":" ", ";":" ",
        '"':" pulgadas "
    }
    for a,b in repl.items():
        s = s.replace(a,b)
    s = re.sub(r'[^a-z0-9\. ]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _header_has_token(header, *tokens):
    h = _normalize_text(header)
    return all(t in h for t in tokens)


def _find_header_index(headers, candidates):
    for idx, header in enumerate(headers):
        h = _normalize_text(header)
        if not h:
            continue
        for candidate in candidates:
            if callable(candidate):
                try:
                    if candidate(h):
                        return idx
                except Exception:
                    continue
            else:
                c = _normalize_text(candidate)
                if h == c or c in h:
                    return idx
    return None


def _material_kind_from_tipo(tipo_text):
    tipo = _normalize_text(tipo_text)
    if not tipo:
        return None
    if "mano" in tipo or "obra" in tipo:
        return "mano_obra"
    if "equipo" in tipo or "herramienta" in tipo:
        return "equipo"
    if "material" in tipo:
        return "material"
    if "basic" in tipo:
        return "basico"
    return tipo

STOPWORDS = {
    "de","la","el","los","las","para","con","por","del","en","y","a","o","un","una","que","tipo",
    "incluye","incluyen","aplicada","aplicado","sobre","mano","obra","equipo","herramienta",
    "preparacion","superficie","servicio","general","trabajo","no","num","numero","cuadrilla"
}

def _tokenize_text(text):
    raw = _normalize_text(text).split()
    return [t for t in raw if t and t not in STOPWORDS and len(t) > 1]

UNIT_MAP = {
    "lt": ["lt", "l", "litro", "litros"],
    "kg": ["kg", "kilo", "kilogramo", "kilogramos"],
    "pza": ["pza", "pieza", "pz", "pzas"],
    "m": ["m", "metro", "metros"],
    "m2": ["m2", "m²", "metro2", "metros2"],
    "m3": ["m3", "m³", "metro3", "metros3"],
    "jor": ["jor", "jornal", "jorn"],
    "saco": ["saco", "sacos"],
    "rollo": ["rollo", "rollos"],
    "tramo": ["tramo", "tramos"],
    "lote": ["lote", "lotes"],
    "serv": ["serv", "servicio", "servicios"],
}

def normalize_unit(unit):
    u = _normalize_text(unit)
    if not u:
        return ""
    for k, vals in UNIT_MAP.items():
        if u in vals:
            return k
    return u

def _extract_fractional_inches(text):
    tokens = []
    for m in re.finditer(r'(\d+(?:[\.-]\d+)?(?:\s*1/2|\s*1/4|\s*3/4)?)\s*pulgadas?', _normalize_text(text)):
        tokens.append(m.group(1).strip())
    for m in re.finditer(r'(\d+(?:[\.-]\d+)?(?:/\d+)?)', str(text)):
        part = m.group(1)
        if "/" in part or "." in part or len(part) <= 2:
            tokens.append(part)
    return list(dict.fromkeys(tokens))

MATERIAL_TYPE_KEYWORDS = {
    "tubo": ["tubo", "tuberia"],
    "codo": ["codo"],
    "tee": ["tee", "yee"],
    "reduccion": ["reduccion", "reductor"],
    "pintura": ["pintura", "esmalte", "primario"],
    "soldadura": ["soldadura", "electrodo", "argon", "alambre"],
    "silicon": ["silicon"],
    "perfil": ["perfil", "ptr", "lamina", "placa"],
    "panel": ["panel"],
    "fijacion": ["tornillo", "tuerca", "rondana", "ancla", "fijacion"],
    "thinner": ["thinner", "solvente"],
    "cinta": ["cinta"],
    "hule": ["hule"],
    "chapeton": ["chapeton"],
    "fester": ["fester", "cm200"],
}

def _extract_material_attributes(text, unit=""):
    norm = _normalize_text(text)
    tokens = _tokenize_text(text)
    tipo = None
    for k, vals in MATERIAL_TYPE_KEYWORDS.items():
        if any(v in norm for v in vals):
            tipo = k
            break
    material = None
    for m in ["acero inoxidable", "acero al carbon", "pvc", "cpvc", "comex", "inoxidable", "carbono", "aluminio"]:
        if m in norm:
            material = m
            break
    grade = None
    for g in ["316l","316","304l","304","ced 80","ced 40","sch 80","sch 40","7018","100"]:
        if g in norm:
            grade = g
            break
    diameter = None
    diams = _extract_fractional_inches(text)
    if diams:
        diameter = diams[0]
    return {
        "tokens": tokens,
        "tipo": tipo,
        "material": material,
        "grado": grade,
        "diametro": diameter,
        "unidad_norm": normalize_unit(unit),
        "norm": norm,
    }


LABOR_ALIAS_MAP = {
    "SUPERVISOR DE OBRA": ["sup-o", "supervisor obra", "supervisor de obra", "residente", "cabo de oficios"],
    "SUPERVISOR DE SEGURIDAD": ["sup-seg", "supervisor seguridad", "supervisor de seguridad"],
    "OFICIAL SOLDADOR": ["of-sol", "soldador", "argonero", "soldador/argonero", "pailero", "armador"],
    "OFICIAL TUBERO": ["of-tub", "tubero", "plomero"],
    "AYUDANTE GENERAL": ["ayu-gral", "ayudante general", "ayudante", "peon", "maniobrista"],
    "OFICIAL ELECTRICISTA": ["electricista", "electrico", "eléctrico"],
    "OFICIAL ALBANIL": ["albanil", "albañil", "oficial albanil", "oficial albañil"],
}

LABOR_CLASS_KEYWORDS = {
    "SUPERVISOR DE OBRA": ["supervisor obra", "supervisor de obra", "residente", "cabo de oficios", "sobrestante"],
    "SUPERVISOR DE SEGURIDAD": ["supervisor seguridad", "supervisor de seguridad", "seguridad industrial"],
    "OFICIAL SOLDADOR": ["soldador", "argonero", "pailero", "armador", "soldadura"],
    "OFICIAL TUBERO": ["tubero", "plomero", "tuberia", "tubería"],
    "AYUDANTE GENERAL": ["ayudante", "peon", "peón", "maniobrista", "personal"],
    "OFICIAL ELECTRICISTA": ["electricista", "eléctrico", "electrico", "cableado", "conduit"],
    "OFICIAL ALBANIL": ["albanil", "albañil", "mampost", "concreto", "cimbra", "yesero", "pintor", "herrero"],
}


def _classify_labor(desc):
    norm = _normalize_text(desc)
    if not norm:
        return None
    if "cuadrilla de trabajo" in norm:
        return None
    for target, aliases in LABOR_ALIAS_MAP.items():
        if any(a in norm for a in aliases):
            return target
    for target, kws in LABOR_CLASS_KEYWORDS.items():
        if any(k in norm for k in kws):
            return target
    return None

def _section_kind(text):
    t = _normalize_text(text)
    if not t:
        return None
    if "material" in t:
        return "materiales"
    if "mano de obra" in t or "mano obra" in t:
        return "mano_obra"
    if "equipo y herramienta" in t or "equipo herramienta" in t or t == "equipo" or t == "equipo y":
        return "equipo"
    if "basico" in t:
        return "basicos"
    return None

def _parse_component_row(row):
    c0 = _clean_text(row[0] if len(row) > 0 else "")
    c1 = _clean_text(row[1] if len(row) > 1 else "")
    if not c0 and not c1:
        return None
    if c0.upper().startswith("SUBTOTAL"):
        return None
    if _section_kind(c0) or _section_kind(c1):
        return None
    txt = _normalize_text(c0 + " " + c1)
    if "precio unitario" in txt or "costo directo" in txt or "indirect" in txt or "utilidad" in txt or "subtotal1" in txt or "subtotal2" in txt:
        return None
    unidad = _clean_text(row[2] if len(row) > 2 else "")
    precio_base = _as_float(row[3] if len(row) > 3 else None)
    factor = _as_float(row[5] if len(row) > 5 else None)
    importe = _as_float(row[6] if len(row) > 6 else None)
    if precio_base is None and factor is None and importe is None:
        return None
    return {
        "codigo": c0,
        "descripcion": c1 or c0,
        "unidad": unidad,
        "unidad_norm": normalize_unit(unidad),
        "precio_base": precio_base,
        "factor": factor,
        "importe": importe,
        "attributes": _extract_material_attributes(c1 or c0, unidad),
    }

def _collect_components(rows, start, end):
    components = {"materiales": [], "mano_obra": [], "equipo": [], "basicos": []}
    current = None
    for i in range(start, min(end, len(rows))):
        row = rows[i]
        t0 = _clean_text(row[0] if len(row) > 0 else "")
        t1 = _clean_text(row[1] if len(row) > 1 else "")
        sec = _section_kind(t0) or _section_kind(t1)
        if sec:
            current = sec
            continue
        if current is None:
            continue
        txt = _row_text(row)
        if "costo directo" in txt or "indirect" in txt or "utilidad" in txt or "precio unitario" in txt or "subtotal1" in txt or "subtotal2" in txt:
            continue
        item = _parse_component_row(row)
        if item:
            components[current].append(item)
    return components

def _token_overlap_score(a_tokens, b_tokens):
    sa = set(a_tokens or [])
    sb = set(b_tokens or [])
    if not sa or not sb:
        return 0.0
    inter = len(sa & sb)
    union = len(sa | sb)
    return inter / union if union else 0.0

def _material_match_score(input_attrs, cand):
    score = 0.0
    text_score = _token_overlap_score(input_attrs.get("tokens"), cand.get("tokens"))
    score += 0.40 * text_score
    if input_attrs.get("unidad_norm") and input_attrs.get("unidad_norm") == cand.get("unidad_norm"):
        score += 0.18
    if input_attrs.get("tipo") and input_attrs.get("tipo") == cand.get("tipo"):
        score += 0.16
    if input_attrs.get("material") and input_attrs.get("material") == cand.get("material"):
        score += 0.12
    if input_attrs.get("grado") and input_attrs.get("grado") == cand.get("grado"):
        score += 0.08
    if input_attrs.get("diametro") and input_attrs.get("diametro") == cand.get("diametro"):
        score += 0.06
    price_in = cand.get("_price_input")
    price_ref = cand.get("precio")
    if price_in and price_ref:
        delta = abs(price_in - price_ref) / max(abs(price_ref), 1)
        if delta <= 0.20:
            score += 0.05
        elif delta <= 0.5:
            score += 0.02
    return round(score, 4)


def _candidate_unit_bonus(input_unit, candidate_unit):
    iu = normalize_unit(input_unit)
    cu = normalize_unit(candidate_unit)
    if not iu or not cu:
        return 0.0
    if iu == cu:
        return 0.16
    compatible = {("kg", "ton"), ("ton", "kg"), ("lt", "gal"), ("gal", "lt"), ("pza", "pieza"), ("m", "ml")}
    return 0.06 if (iu, cu) in compatible else -0.05

def search_material_in_benchmark(item, materials_index, top_k=12):
    """
    Genera varios candidatos desde el Excel Construdata/benchmark cargado por Admin.
    No decide precio por sí solo: entrega candidatos para Voyage/Claude y fallback seguro.
    """
    attrs = item.get("attributes") or _extract_material_attributes(item.get("descripcion"), item.get("unidad"))
    input_tokens = set(attrs.get("tokens") or [])
    prelim = []
    for cand in materials_index or []:
        cand_tokens = set(cand.get("tokens") or [])
        shared = input_tokens & cand_tokens
        if not shared:
            continue
        c = dict(cand)
        c["_price_input"] = item.get("precio_base")
        base_score = _material_match_score(attrs, c)
        base_score += min(0.14, len(shared) * 0.025)
        base_score += _candidate_unit_bonus(item.get("unidad"), c.get("unidad"))
        if attrs.get("diametro") and c.get("diametro") and attrs.get("diametro") != c.get("diametro"):
            base_score -= 0.18
        if attrs.get("grado") and c.get("grado") and attrs.get("grado") != c.get("grado"):
            base_score -= 0.12
        c["score"] = round(max(0, min(0.98, base_score)), 4)
        c["fuente_precio"] = "construdata excel admin"
        c["candidate_reason"] = f"tokens compartidos: {', '.join(sorted(shared)[:8])}"
        prelim.append(c)

    prelim.sort(key=lambda x: x["score"], reverse=True)
    top = prelim[:top_k]
    if not top:
        return {"selected": None, "top_candidates": [], "confidence": 0.0, "status": "sin_match"}

    selected = top[0] if top[0]["score"] >= 0.50 else None
    selected_score = selected.get("score", 0.0) if isinstance(selected, dict) else 0.0
    return {
        "selected": selected,
        "top_candidates": top,
        "confidence": round(selected_score, 4),
        "status": "candidatos" if top else "sin_match",
    }

def _material_dedupe_key(item):
    return (_normalize_text(item.get("descripcion")), normalize_unit(item.get("unidad")))

def match_materials_against_benchmark(materials_items, materials_index, important_scope=None, clave_servicio=None, memo=None):
    """
    Matching de materiales con arquitectura de bajo costo:
    1. Cache de corrida.
    2. Cache persistente entre corridas.
    3. Índice vectorial Construdata completo si existe.
    4. Fallback textual local.
    5. Voyage/Claude solo sobre candidatos reducidos.
    """
    results = []
    memo = memo if memo is not None else {}
    for item in materials_items or []:
        unit = normalize_unit(item.get("unidad"))
        if unit in {"jor", "%"}:
            continue

        scope_key = (clave_servicio, item.get("descripcion"), item.get("unidad"), item.get("precio_base"))
        dedupe_key = _material_dedupe_key(item)

        if dedupe_key in memo:
            result = dict(memo[dedupe_key])
            result["dedup_reused"] = True
            resolution = dict(result.get("market_resolution") or {})
            if resolution:
                resolution["decision_source"] = "cache memoria de corrida"
                if resolution.get("source") == "fallback contratista":
                    resolution["effective_market_price"] = item.get("precio_base")
                result["market_resolution"] = resolution
            results.append({"input": item, "match": result})
            continue

        threshold_pct = float(os.getenv("MATERIAL_MATCH_THRESHOLD", "60"))
        threshold = threshold_pct / 100.0

        cached = get_cached_material_match(item, materials_index, threshold_pct)
        if cached:
            result = dict(cached)
            result["persistent_cache_reused"] = True
            resolution = dict(result.get("market_resolution") or {})
            if resolution:
                resolution["decision_source"] = "cache persistente"
                if resolution.get("source") == "fallback contratista":
                    resolution["effective_market_price"] = item.get("precio_base")
                result["market_resolution"] = resolution
            memo[dedupe_key] = dict(result)
            results.append({"input": item, "match": result})
            continue

        vector_candidates, vector_meta = vector_search_candidates(
            item,
            materials_index,
            top_k=int(os.getenv("MATERIAL_VECTOR_TOP_K", "25")),
        )

        result = search_material_in_benchmark(item, materials_index)
        local_candidates = list(result.get("top_candidates") or [])

        merged_candidates = []
        seen_cands = set()
        for cand in list(vector_candidates or []) + local_candidates:
            key = (cand.get("clave"), cand.get("unidad"), cand.get("descripcion"))
            if key in seen_cands:
                continue
            seen_cands.add(key)
            merged_candidates.append(cand)
        merged_candidates.sort(key=lambda x: float(x.get("score") or 0), reverse=True)

        if merged_candidates:
            result["top_candidates"] = merged_candidates[:25]
            result["selected"] = merged_candidates[0]
            result["confidence"] = float(merged_candidates[0].get("score") or 0)
            result["vector_index_meta"] = vector_meta

        best_score = max(
            [float((result.get("selected") or {}).get("score") or 0)]
            + [float(c.get("score") or 0) for c in (result.get("top_candidates") or [])]
        ) if (result.get("selected") or result.get("top_candidates")) else 0.0

        should_deep_search = (
            important_scope is None
            or scope_key in important_scope
            or best_score >= threshold
            or bool(vector_candidates)
            or bool(os.getenv("FORCE_MATERIAL_AI_ALL", "").strip())
        )

        if should_deep_search:
            market_resolution = select_material_market_price(item, result)
            market_resolution["vector_index_meta"] = vector_meta
            if vector_candidates:
                market_resolution["decision_source"] = "vector_index_construdata"
            result["market_resolution"] = market_resolution

            selected = result.get("selected") or {}
            effective = market_resolution.get("effective_market_price")
            if effective is not None:
                selected = dict(selected) if selected else {}
                selected.setdefault("descripcion", (market_resolution.get("evidence") or {}).get("descripcion") or item.get("descripcion"))
                selected.setdefault("unidad", market_resolution.get("market_unit") or item.get("unidad"))
                selected["precio_evidencia"] = market_resolution.get("market_price_raw")
                selected["precio"] = effective
                selected["score"] = (market_resolution.get("confidence_pct") or 0) / 100
                selected["fuente_precio"] = market_resolution.get("source")
                result["selected"] = selected
        else:
            result["market_resolution"] = {
                "normalized": {"normalized": _normalize_text(item.get("descripcion")), "ai_status": "skipped:no_candidate_above_threshold"},
                "source": "fallback contratista",
                "evidence": result.get("selected") or {},
                "market_unit": item.get("unidad"),
                "market_price_raw": None,
                "conversion_factor": 1,
                "effective_market_price": item.get("precio_base"),
                "confidence_pct": 0,
                "decision_criterion": "Sin candidato Construdata/vectorial suficiente; se conserva precio contratista como fallback.",
                "fecha_consulta": None,
                "vector_index_meta": vector_meta,
            }

        memo[dedupe_key] = dict(result)
        save_material_match(item, result, materials_index, threshold_pct)
        results.append({"input": item, "match": result})
    return results

def summarize_material_market_matches(matches):
    comparados = len(matches or [])
    con_match = sum(1 for x in matches or [] if x.get("match", {}).get("selected"))
    scores = [x["match"]["selected"]["score"] for x in matches or [] if x.get("match", {}).get("selected")]
    return {
        "comparados": comparados,
        "con_match": con_match,
        "sin_match": comparados - con_match,
        "score_promedio": round(sum(scores)/len(scores), 4) if scores else None,
    }

def load_labor_tabulador(filepath=None, year="2026"):
    filepath = filepath or str((_quantia_data_path("Tabulador de Mano de Obra 2024-2025-2026-1.xlsx")))
    wb = openpyxl.load_workbook(filepath, data_only=True)
    target = None
    for s in wb.sheetnames:
        if year in str(s):
            target = s
            break
    target = target or wb.sheetnames[0]
    ws = wb[target]
    rows = list(ws.iter_rows(values_only=True))
    idx_row = None
    for i,row in enumerate(rows[:40]):
        row_norm = [_normalize_text(c) for c in row]
        if "codigo" in row_norm and "concepto" in row_norm and any("salario real" in c for c in row_norm):
            idx_row = i
            headers = row_norm
            break
    if idx_row is None:
        return {}
    col_codigo = headers.index("codigo")
    col_concepto = headers.index("concepto")
    col_unidad = headers.index("unidad")
    col_salario_real = next((i for i,c in enumerate(headers) if c.strip() == "salario real"), None)
    if col_salario_real is None:
        col_salario_real = next(i for i,c in enumerate(headers) if "salario real" in c and "factor" not in c)
    data = {}
    for row in rows[idx_row+1:]:
        cod = row[col_codigo] if col_codigo < len(row) else None
        concepto = row[col_concepto] if col_concepto < len(row) else None
        if not cod or not concepto:
            continue
        salario = _as_float(row[col_salario_real] if col_salario_real < len(row) else None)
        unidad = row[col_unidad] if col_unidad < len(row) else None
        norm_conc = _normalize_text(concepto)
        data[str(cod).strip()] = {"codigo": str(cod).strip(), "concepto": str(concepto).strip(), "unidad": str(unidad or "").strip(), "salario": salario, "norm": norm_conc}
    return data

def build_labor_alias_index(labor_table):
    alias_index = {}
    for entry in labor_table.values():
        norm = entry.get("norm","")
        for target, aliases in LABOR_ALIAS_MAP.items():
            if any(a in norm for a in aliases):
                alias_index[target] = entry
    return alias_index



def _best_labor_reference_by_text(item, labor_table):
    item_norm = _normalize_text(item.get("descripcion") or item.get("codigo") or "")
    item_tokens = set(_tokenize_text(item_norm))
    best = None
    best_score = 0.0
    for entry in labor_table.values():
        cand_tokens = set(_tokenize_text(entry.get("concepto") or entry.get("norm") or ""))
        if not cand_tokens:
            continue
        overlap = _token_overlap_score(item_tokens, cand_tokens)
        if overlap > best_score:
            best_score = overlap
            best = entry
    return best if best_score >= 0.25 else None


def _labor_unit_factor_to_hr(unit):
    u = normalize_unit(unit)
    if u in {"hr", "hrs", "hora", "horas", "hh", "h"}:
        return 1.0
    if u in {"jor", "jorn", "jornal", "jornada", "dia", "dias"}:
        return 8.0
    return None


def _convert_labor_quantity_to_hr(quantity, unit):
    if not isinstance(quantity, (int, float)):
        return quantity, normalize_unit(unit)
    factor = _labor_unit_factor_to_hr(unit)
    if factor is None:
        return float(quantity), normalize_unit(unit)
    return float(quantity) * factor, "hr"


def _convert_labor_price_to_hr(price, unit):
    if not isinstance(price, (int, float)):
        return price, normalize_unit(unit)
    factor = _labor_unit_factor_to_hr(unit)
    if factor is None or factor == 0:
        return float(price), normalize_unit(unit)
    return float(price) / factor, "hr"


def _build_labor_comparison_basis(item, ref):
    provider_price = item.get("precio_base")
    provider_unit = item.get("unidad") or item.get("unidad_norm")
    ref_price = ref.get("salario")
    ref_unit = ref.get("unidad")

    provider_factor_hr = _labor_unit_factor_to_hr(provider_unit)
    ref_factor_hr = _labor_unit_factor_to_hr(ref_unit)

    provider_price_cmp = float(provider_price) if isinstance(provider_price, (int, float)) else None
    provider_unit_cmp = normalize_unit(provider_unit)
    ref_price_cmp = float(ref_price) if isinstance(ref_price, (int, float)) else None
    ref_unit_cmp = normalize_unit(ref_unit)

    if provider_factor_hr and ref_factor_hr:
        provider_price_cmp, provider_unit_cmp = _convert_labor_price_to_hr(provider_price_cmp, provider_unit)
        ref_price_cmp, ref_unit_cmp = _convert_labor_price_to_hr(ref_price_cmp, ref_unit)

    input_factor_cmp = item.get("factor")
    input_factor_unit_cmp = provider_unit_cmp
    if provider_factor_hr:
        input_factor_cmp, input_factor_unit_cmp = _convert_labor_quantity_to_hr(item.get("factor"), provider_unit)

    return {
        "provider_price_cmp": provider_price_cmp,
        "provider_unit_cmp": provider_unit_cmp,
        "ref_price_cmp": ref_price_cmp,
        "ref_unit_cmp": ref_unit_cmp,
        "input_factor_cmp": input_factor_cmp,
        "input_factor_unit_cmp": input_factor_unit_cmp,
        "provider_unit_original": normalize_unit(provider_unit),
        "ref_unit_original": normalize_unit(ref_unit),
    }


def match_labor_item(item, labor_alias_index, labor_table=None):
    target = _classify_labor(item.get("descripcion", "") or item.get("codigo", ""))
    ref = labor_alias_index.get(target) if target else None
    if (not ref or not ref.get("salario")) and labor_table:
        ref = _best_labor_reference_by_text(item, labor_table)
        if ref and not target:
            target = ref.get("concepto")
    if not ref or not ref.get("salario"):
        return None

    basis = _build_labor_comparison_basis(item, ref)
    proveedor_cmp = basis.get("provider_price_cmp")
    mercado_cmp = basis.get("ref_price_cmp")
    if not proveedor_cmp or not mercado_cmp:
        return None

    delta = (proveedor_cmp - mercado_cmp) / mercado_cmp
    if delta <= 0.10:
        estado = "en_rango"
    elif delta <= 0.30:
        estado = "arriba"
    else:
        estado = "muy_arriba"
    return {
        "tipo": target or ref.get("concepto"),
        "referencia_codigo": ref.get("codigo"),
        "referencia_concepto": ref.get("concepto"),
        "precio_proveedor": item.get("precio_base"),
        "precio_tabulador": ref.get("salario"),
        "precio_proveedor_comparable": proveedor_cmp,
        "precio_tabulador_comparable": mercado_cmp,
        "unidad_comparable": basis.get("ref_unit_cmp") or basis.get("provider_unit_cmp") or "hr",
        "factor_comparable": basis.get("input_factor_cmp"),
        "factor_unidad_comparable": basis.get("input_factor_unit_cmp"),
        "conversion_regla": "1 JOR = 8 HR" if basis.get("provider_unit_original") in {"jor", "jorn", "jornal", "jornada", "dia", "dias"} or basis.get("ref_unit_original") in {"jor", "jorn", "jornal", "jornada", "dia", "dias"} else None,
        "delta": round(delta, 4),
        "estado": estado,
    }


def match_labor_against_tabulador(mano_obra_items, labor_alias_index, labor_table=None):
    results = []
    for item in mano_obra_items or []:
        unit = normalize_unit(item.get("unidad"))
        desc_norm = _normalize_text(item.get("descripcion") or item.get("codigo") or "")
        code_norm = _normalize_text(item.get("codigo") or "")
        if "%herr" in code_norm or "herramienta" in desc_norm or "equipo" in desc_norm:
            continue
        if unit not in {"jor", "", "hh", "jorn", "dia", "dias"} and not _classify_labor(desc_norm):
            continue
        mt = match_labor_item(item, labor_alias_index, labor_table)
        if mt:
            results.append({"input": item, "match": mt})
    return results

def summarize_labor_matches(matches):
    comparados = len(matches or [])
    muy_arriba = sum(1 for x in matches or [] if x.get("match",{}).get("estado") == "muy_arriba")
    arriba = sum(1 for x in matches or [] if x.get("match",{}).get("estado") == "arriba")
    en_rango = sum(1 for x in matches or [] if x.get("match",{}).get("estado") == "en_rango")
    deltas = [x["match"]["delta"] for x in matches or [] if x.get("match")]
    return {"comparados": comparados, "muy_arriba": muy_arriba, "arriba": arriba, "en_rango": en_rango, "delta_promedio": round(sum(deltas)/len(deltas),4) if deltas else None}


EQUIPMENT_BENCHMARK_FILE = _quantia_data_path("equipment_special_benchmark.json")

def load_equipment_benchmark(filepath=None):
    filepath = filepath or str(EQUIPMENT_BENCHMARK_FILE)
    try:
        data = json.loads(Path(filepath).read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []

def _equipment_key_tokens(text):
    tokens = set(_normalize_text(text).split())
    return {t for t in tokens if t not in {"renta","equipo","especial","servicio","jornada","dia","dias","hora","horas"}}

def _extract_equipment_signature(text):
    t = _normalize_text(text)
    sig = {
        "tipo": None,
        "capacidad_ton": None,
        "tokens": list(_equipment_key_tokens(text)),
    }
    if "montacargas" in t:
        sig["tipo"] = "montacargas"
    elif "grua" in t:
        sig["tipo"] = "grua"
    elif "plataforma" in t:
        sig["tipo"] = "plataforma"
    elif "retroexcavadora" in t:
        sig["tipo"] = "retroexcavadora"
    elif "generador" in t:
        sig["tipo"] = "generador"
    m = re.search(r'(\d+(?:[\.,]\d+)?)\s*(?:ton|tons|tonelada|toneladas)', t)
    if m:
        try:
            sig["capacidad_ton"] = float(m.group(1).replace(",", "."))
        except Exception:
            pass
    return sig

def build_equipment_index(rows):
    index = []
    for row in rows or []:
        desc = row.get("descripcion") or ""
        unidad = row.get("unidad") or ""
        sig = _extract_equipment_signature(desc)
        tokens = list(_equipment_key_tokens(desc))
        entry = dict(row)
        entry["unidad_norm"] = normalize_unit(unidad)
        entry["tokens"] = tokens
        entry["signature"] = sig
        index.append(entry)
    return index

def search_equipment_in_benchmark(item, equipment_index):
    item_desc = item.get("descripcion") or ""
    item_unit = normalize_unit(item.get("unidad"))
    item_sig = _extract_equipment_signature(item_desc)
    item_tokens = set(_equipment_key_tokens(item_desc))
    candidates = []
    for cand in equipment_index or []:
        cand_tokens = set(cand.get("tokens") or [])
        text_score = _token_overlap_score(item_tokens, cand_tokens)
        unit_score = 1.0 if item_unit and item_unit == cand.get("unidad_norm") else 0.0
        sig_score = 0.0
        cand_sig = cand.get("signature") or {}
        if item_sig.get("tipo") and item_sig.get("tipo") == cand_sig.get("tipo"):
            sig_score += 0.45
        if item_sig.get("capacidad_ton") and cand_sig.get("capacidad_ton"):
            delta = abs(item_sig["capacidad_ton"] - cand_sig["capacidad_ton"])
            if delta <= 0.25:
                sig_score += 0.35
            elif delta <= 1:
                sig_score += 0.15
        total = 0.45 * text_score + 0.25 * unit_score + 0.30 * sig_score
        if total >= 0.20:
            candidates.append({
                "descripcion": cand.get("descripcion"),
                "unidad": cand.get("unidad"),
                "precio_ref": cand.get("precio_ref"),
                "rango_bajo": cand.get("rango_bajo"),
                "rango_alto": cand.get("rango_alto"),
                "fuente": cand.get("fuente"),
                "evidencia": cand.get("evidencia"),
                "score": round(total, 4),
            })
    candidates.sort(key=lambda x: x["score"], reverse=True)
    selected = candidates[0] if candidates else None
    if selected and selected["score"] < 0.38:
        selected = None
    return {"selected": selected, "top": candidates[:10]}

def match_equipment_against_benchmark(equipo_items, equipment_index):
    results = []
    for item in equipo_items or []:
        mt = search_equipment_in_benchmark(item, equipment_index)
        results.append({"input": item, "match": mt})
    return results

def summarize_equipment_matches(matches):
    comparados = len(matches or [])
    con_match = sum(1 for x in matches or [] if (x.get("match") or {}).get("selected"))
    deltas = []
    criticals = 0
    for x in matches or []:
        sel = (x.get("match") or {}).get("selected")
        inp = x.get("input") or {}
        if sel and isinstance(inp.get("precio_base"), (int,float)) and isinstance(sel.get("precio_ref"), (int,float)) and sel.get("precio_ref"):
            delta = (float(inp["precio_base"]) - float(sel["precio_ref"])) / float(sel["precio_ref"])
            deltas.append(delta)
            if delta > 0.5:
                criticals += 1
    return {
        "comparados": comparados,
        "con_match": con_match,
        "sin_match": comparados - con_match,
        "delta_promedio": round(sum(deltas)/len(deltas),4) if deltas else None,
        "criticos": criticals,
    }


def calculate_executive_findings(provider_name, conceptos, total_proveedor, all_provider_totals):
    findings = []
    concept_rows = []
    total_mercado = 0.0
    total_contratista = 0.0
    all_items = []
    for clave, c in conceptos.items():
        total = c.get("total")
        if not isinstance(total, (int, float)):
            continue
        total_contratista += float(total)
        mercado_materiales = 0.0
        for m in c.get("materiales_benchmark_matches") or []:
            sel = (m.get("match") or {}).get("selected")
            inp = m.get("input") or {}
            if sel and isinstance(inp.get("factor"), (int,float)) and isinstance(sel.get("precio"), (int,float)):
                mercado_materiales += float(inp["factor"]) * float(sel["precio"])
        mercado_mo = 0.0
        for m in c.get("mano_obra_benchmark_matches") or []:
            mt = m.get("match") or {}
            inp = m.get("input") or {}
            factor_cmp = mt.get("factor_comparable") if isinstance(mt.get("factor_comparable"), (int, float)) else inp.get("factor")
            precio_cmp = mt.get("precio_tabulador_comparable") if isinstance(mt.get("precio_tabulador_comparable"), (int, float)) else mt.get("precio_tabulador")
            if isinstance(precio_cmp, (int, float)) and isinstance(factor_cmp, (int, float)):
                mercado_mo += float(factor_cmp) * float(precio_cmp)
        ci = c.get("subtotal2")
        ci_pct = None
        if isinstance(c.get("costo_directo"), (int,float)) and c.get("costo_directo"):
            ci_pct = (float(ci)/float(c.get("costo_directo"))) if isinstance(ci, (int,float)) else None
        mercado_equipo = 0.0
        for e in c.get("equipo_benchmark_matches") or []:
            mt = (e.get("match") or {}).get("selected") or {}
            inp = e.get("input") or {}
            if mt.get("precio_ref") and isinstance(inp.get("factor"), (int,float)):
                mercado_equipo += float(inp["factor"]) * float(mt["precio_ref"])
        market_total = mercado_materiales + mercado_mo + mercado_equipo
        total_mercado += market_total
        participacion = float(total)/float(total_proveedor) if total_proveedor else 0
        delta = ((float(total) - market_total) / market_total) if market_total else None
        impact = participacion * max(delta or 0, 0)
        concept_rows.append({
            "clave": clave,
            "descripcion": c.get("desc",""),
            "total_contratista": round(float(total),2),
            "total_mercado": round(market_total,2) if market_total else None,
            "participacion": round(participacion,4),
            "delta": round(delta,4) if delta is not None else None,
            "impacto": round(impact,4) if delta is not None else None,
            "indirecto_pct": round(ci_pct,4) if ci_pct is not None else None,
        })
        # driver items from material/labor/equipment percentages
        for mm in c.get("materiales_benchmark_matches") or []:
            sel=(mm.get("match") or {}).get("selected")
            inp=mm.get("input") or {}
            if sel and sel.get("precio") and inp.get("factor"):
                prov_total = (inp.get("factor") or 0) * (inp.get("precio_base") or 0)
                mkt_total = (inp.get("factor") or 0) * (sel.get("precio") or 0)
                delta_item = (prov_total - mkt_total) / mkt_total if mkt_total else None
                part_item = prov_total / total_proveedor if total_proveedor else 0
                if delta_item is not None and delta_item > 0:
                    all_items.append({"tipo":"material","concepto":inp.get("descripcion"),"clave_servicio":clave,"proveedor":round(prov_total,2),"mercado":round(mkt_total,2),"delta":round(delta_item,4),"participacion":round(part_item,4),"impacto":round(part_item*delta_item,4),"estado":"critico" if delta_item>0.5 and part_item>0.05 else ("alto" if delta_item>0.2 else "medio")})
        for lm in c.get("mano_obra_benchmark_matches") or []:
            mt=lm.get("match") or {}
            inp=lm.get("input") or {}
            factor_cmp = mt.get("factor_comparable") if isinstance(mt.get("factor_comparable"), (int, float)) else inp.get("factor")
            prov_price_cmp = mt.get("precio_proveedor_comparable") if isinstance(mt.get("precio_proveedor_comparable"), (int, float)) else inp.get("precio_base")
            mkt_price_cmp = mt.get("precio_tabulador_comparable") if isinstance(mt.get("precio_tabulador_comparable"), (int, float)) else mt.get("precio_tabulador")
            if mt and isinstance(mkt_price_cmp, (int, float)) and isinstance(factor_cmp, (int, float)):
                prov_total=(factor_cmp or 0)*(prov_price_cmp or 0)
                mkt_total=(factor_cmp or 0)*(mkt_price_cmp or 0)
                delta_item=(prov_total-mkt_total)/mkt_total if mkt_total else None
                part_item=prov_total/total_proveedor if total_proveedor else 0
                if delta_item is not None and delta_item>0:
                    all_items.append({"tipo":"mano_obra","concepto":mt.get("tipo"),"clave_servicio":clave,"proveedor":round(prov_total,2),"mercado":round(mkt_total,2),"delta":round(delta_item,4),"participacion":round(part_item,4),"impacto":round(part_item*delta_item,4),"estado":"critico" if delta_item>0.3 and part_item>0.03 else ("alto" if delta_item>0.15 else "medio")})
        matched_equipment = False
        for em in c.get("equipo_benchmark_matches") or []:
            mt = (em.get("match") or {}).get("selected") or {}
            inp = em.get("input") or {}
            if mt and mt.get("precio_ref") and inp.get("factor"):
                matched_equipment = True
                prov_total = (inp.get("factor") or 0) * (inp.get("precio_base") or 0)
                mkt_total = (inp.get("factor") or 0) * (mt.get("precio_ref") or 0)
                delta_item = (prov_total - mkt_total) / mkt_total if mkt_total else None
                part_item = prov_total / total_proveedor if total_proveedor else 0
                if delta_item is not None and delta_item > 0:
                    all_items.append({"tipo":"equipo_especial","concepto":inp.get("descripcion"),"clave_servicio":clave,"proveedor":round(prov_total,2),"mercado":round(mkt_total,2),"delta":round(delta_item,4),"participacion":round(part_item,4),"impacto":round(part_item*delta_item,4),"estado":"critico" if delta_item>0.5 and part_item>0.05 else ("alto" if delta_item>0.2 else "medio"),"fuente":"benchmark_manual"})
        eq = c.get("equipo")
        if isinstance(eq, (int,float)) and eq > 0 and not matched_equipment:
            part_eq = eq / total_proveedor if total_proveedor else 0
            if part_eq >= 0.05:
                all_items.append({"tipo":"equipo","concepto":"Equipo / herramienta sin referencia","clave_servicio":clave,"proveedor":round(eq,2),"mercado":None,"delta":None,"participacion":round(part_eq,4),"impacto":round(part_eq,4),"estado":"revisar","fuente":"sin_referencia"})
    concept_rows.sort(key=lambda x: x["total_contratista"], reverse=True)
    acum=0
    top_partidas=[]
    for r in concept_rows:
        acum += r["participacion"]
        rr=dict(r); rr["acumulado"]=round(acum,4)
        top_partidas.append(rr)
    all_items.sort(key=lambda x: x.get("impacto") or 0, reverse=True)
    indirectos = []
    for clave, c in conceptos.items():
        cd = c.get("costo_directo")
        ci = c.get("subtotal2")
        if isinstance(cd,(int,float)) and cd and isinstance(ci,(int,float)):
            pct = ci/cd
            if pct > 0.25:
                indirectos.append({"clave":clave,"descripcion":c.get("desc",""),"proveedor_pct":round(pct,4),"mercado_pct":0.25,"estado":"critico" if pct>0.30 else "arriba"})
    return {
        "provider_name": provider_name,
        "total_contratista": round(total_contratista,2),
        "total_mercado": round(total_mercado,2),
        "ajuste_global": round((total_contratista-total_mercado)/total_contratista,4) if total_contratista and total_mercado else None,
        "top_partidas": top_partidas[:50],
        "drivers": all_items[:50],
        "indirectos_alertas": indirectos[:20],
    }


def _looks_like_analysis_row(row):
    txt = _row_text(row)
    return "análisis:" in txt or "analisis:" in txt




def _parse_pct_value(value):
    num = _as_float(value)
    if num is None:
        return None
    if num > 1.5:
        num = num / 100.0
    if 0 <= num <= 1.5:
        return num
    return None


def _find_declared_indirect(rows, start, end, costo_directo=None, pu=None, utilidad=None, financiamiento=None):
    best_pct = None
    best_amount = None
    for i in range(start, min(end, len(rows))):
        row = rows[i]
        txt = _row_text(row)
        if not txt or "indirect" not in txt:
            continue
        nums = [_as_float(x) for x in row]
        nums = [n for n in nums if n is not None]
        pct_candidates = []
        for x in row:
            pct = _parse_pct_value(x)
            if pct is not None and 0 < pct <= 1.0:
                pct_candidates.append(pct)
        amount = None
        if nums:
            # amount usually is the largest numeric in row, pct is the smallest
            amount = max(nums)
        pct = None
        if pct_candidates:
            pct = min(pct_candidates)
        if pct is None and isinstance(costo_directo, (int, float)) and costo_directo and amount is not None:
            cand_pct = float(amount) / float(costo_directo)
            if 0 < cand_pct <= 1.0:
                pct = cand_pct
        if pct is not None:
            if best_pct is None or abs(pct - 0.30) < abs(best_pct - 0.30):
                best_pct = pct
                best_amount = amount
    if best_pct is not None and best_amount is None and isinstance(costo_directo, (int, float)):
        best_amount = float(costo_directo) * float(best_pct)
    # sane fallback from PU components only if explicit indirect not found
    if best_pct is None and isinstance(costo_directo, (int, float)) and costo_directo:
        pu_val = float(pu or 0)
        rem = pu_val - float(costo_directo) - float(utilidad or 0) - float(financiamiento or 0)
        if rem > 0:
            pct = rem / float(costo_directo)
            if 0 < pct <= 1.0:
                best_pct = pct
                best_amount = rem
    return best_amount, best_pct


def _extract_matrix_apu_blocks(rows):
    """
    Parser principal para matrices/APU reales.
    Cada fila 'Análisis:' abre un servicio y conserva detalle de materiales, mano de obra y equipo.
    """
    starts = [i for i, row in enumerate(rows) if _looks_like_analysis_row(row)]
    if not starts:
        return []
    conceptos = []
    starts.append(len(rows))
    for pos in range(len(starts) - 1):
        start = starts[pos]
        end = starts[pos + 1]
        row = rows[start]
        clave = _clean_text(row[1] if len(row) > 1 else "") or _clean_text(row[4] if len(row) > 4 else "") or f"SERVICIO_{pos+1:03d}"
        unidad = _clean_text(row[3] if len(row) > 3 else "") or _clean_text(row[21] if len(row) > 21 else "")
        cantidad = _as_float(row[5] if len(row) > 5 else None)
        desc = ""
        for j in range(start + 1, min(start + 6, end)):
            txtj = _clean_text(rows[j][0] if len(rows[j]) > 0 else "")
            if txtj and txtj.lower() not in {"partida:", "análisis:", "analisis:", "materiales", "mano de obra", "equipo y herramienta", "basicos"}:
                desc = txtj[:400]
                break
        if not desc:
            candidates = []
            for j in range(start, min(start + 6, end)):
                for cell in rows[j]:
                    txtc = _clean_text(cell)
                    if len(txtc) > 10 and "precio unitario" not in txtc.lower() and "subtotal" not in txtc.lower():
                        candidates.append(txtc)
            if candidates:
                desc = max(candidates, key=len)[:400]
        pu = _find_row_value_by_label(rows, start, end, ["precio unitario"])
        if pu is None:
            pu = _find_row_value_by_label(rows, start, end, ["costo directo"])
        if pu is None:
            continue
        total = _extract_declared_total(rows, start, end, pu)
        components = _collect_components(rows, start, end)
        materiales = _find_row_importe_by_label(rows, start, end, ["subtotal: materiales", "subtotal materiales"])
        mano_obra = _find_row_importe_by_label(rows, start, end, ["subtotal: mano de obra", "subtotal mano de obra", "mano de obra"])
        equipo = _find_row_importe_by_label(rows, start, end, ["subtotal: equipo y herramienta", "subtotal equipo y herramienta", "equipo y herramienta"])
        basicos = _find_row_importe_by_label(rows, start, end, ["subtotal: basicos", "subtotal basicos"])
        costo_directo = _find_row_value_by_label(rows, start, end, ["costo directo", "(cd)"])
        subtotal1 = _find_row_value_by_label(rows, start, end, ["subtotal1"])
        subtotal2 = _find_row_value_by_label(rows, start, end, ["subtotal2"])
        utilidad = _find_row_value_by_label(rows, start, end, ["utilidad", "(cu)"])
        financiamiento = _find_row_value_by_label(rows, start, end, ["financiamiento", "(cf)"])
        indirecto, indirecto_pct = _find_declared_indirect(rows, start, end, costo_directo, pu, utilidad, financiamiento)
        conceptos.append({
            "clave": clave.strip(),
            "desc": desc.strip(),
            "unidad": unidad.strip(),
            "cantidad": cantidad if cantidad is not None else 1,
            "pu": pu,
            "total": total,
            "seccion": "",
            "materiales": materiales,
            "mano_obra": mano_obra,
            "equipo": equipo,
            "basicos": basicos,
            "costo_directo": costo_directo,
            "subtotal1": subtotal1,
            "subtotal2": subtotal2,
            "indirecto": indirecto,
            "indirecto_pct": indirecto_pct,
            "utilidad": utilidad,
            "financiamiento": financiamiento,
            "materiales_items": components["materiales"],
            "mano_obra_items": components["mano_obra"],
            "equipo_items": components["equipo"],
            "basicos_items": components["basicos"],
        })
    return conceptos


def _extract_declared_total(rows, start, end, pu):
    # 1) Total en la fila de análisis
    row = rows[start]
    total = _as_float(row[6] if len(row) > 6 else None)
    if total is not None and total >= (pu or 0):
        return total

    # 2) Buscar "importe" o "total" explícito excluyendo subtotales y auxiliares
    for i in range(start, min(end, len(rows))):
        txt = _row_text(rows[i])
        if not txt:
            continue
        if "subtotal" in txt or "total de auxiliares" in txt:
            continue
        if ("importe" in txt or txt.startswith("total") or " total " in f" {txt} ") and "precio unitario" not in txt:
            val = _find_last_numeric(rows[i])
            if val is not None and val >= (pu or 0):
                return val

    # 3) Si no existe, usar PU como total declarado
    return pu

def archivo_parece_matriz_valida(filepath: str) -> bool:
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception:
        return False
    for sheet_name in wb.sheetnames:
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if _extract_matrix_apu_blocks(rows):
            return True
        txt = " ".join(_row_text(r) for r in rows[:25])
        if "precio unitario" in txt and ("análisis" in txt or "analisis" in txt):
            return True
    return False


def _is_header_v1(row):
    if not row: return False
    a=_norm(row[0]) if len(row)>0 else ""
    b=_norm(row[1]) if len(row)>1 else ""
    c=_norm(row[2]) if len(row)>2 else ""
    d=_norm(row[3]) if len(row)>3 else ""
    return a=="clave" and "cant" in b and "uni" in c and ("des" in d)

def extract_template_v1(filepath:str):
    wb=openpyxl.load_workbook(filepath,data_only=True)
    conceptos={}
    for sheet in wb.sheetnames:
        ws=wb[sheet]
        rows=list(ws.iter_rows(values_only=True))
        header=-1
        for i,r in enumerate(rows[:40]):
            if _is_header_v1(r):
                header=i
                break
        if header==-1:
            continue

        i=header+1
        current=None
        direct=0.0
        pcts={
            "indirectos de oficina":0,
            "indirectos de campo":0,
            "financiamiento":0,
            "utilidad":0,
            "cargos adicionales":0,
            "otro porcentaje":0
        }

        def flush():
            nonlocal current,direct,pcts
            if not current: return
            cargos=sum(direct*(v/100) for v in pcts.values())
            pu=direct+cargos
            conceptos[f"{sheet}::{current['clave']}"]={
                "clave":current["clave"],
                "desc":current["desc"],
                "unidad":current["unidad"],
                "cantidad":current["cantidad"],
                "pu":pu,
                "total":pu*current["cantidad"]
            }
            current=None
            direct=0.0
            for k in pcts: pcts[k]=0

        while i<len(rows):
            r=rows[i]; i+=1
            if not r: continue
            c0=r[0] if len(r)>0 else None
            c1=r[1] if len(r)>1 else None
            c2=r[2] if len(r)>2 else None
            c3=r[3] if len(r)>3 else None
            c5=r[5] if len(r)>5 else None
            c6=r[6] if len(r)>6 else None

            if isinstance(c0,str) and re.match(r'^[A-Z0-9-]{2,}$',str(c0)):
                flush()
                current={
                    "clave":str(c0).strip(),
                    "desc":str(c3).strip() if c3 else "",
                    "unidad":str(c2).strip() if c2 else "",
                    "cantidad":float(c1 or 0)
                }
                continue

            if not current: 
                continue

            desc=_norm(c3)

            if desc in pcts:
                try:
                    pcts[desc]=float(c1 or 0)
                except:
                    pass
                continue

            try:
                cant=float(c1 or 0)
                pu=float(c5 or 0)
                imp=float(c6) if c6 else cant*pu
                direct+=imp
            except:
                pass

        flush()
    return conceptos

# ══════════════════════════════════════════════════════════════
#  EXTRACCIÓN
# ══════════════════════════════════════════════════════════════



def extract_conceptos(filepath: str) -> dict:
    """
    Detecta automáticamente el formato del archivo y extrae los conceptos.
    Prioriza matrices/APU reales y conserva 1 bloque APU = 1 servicio.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    for sheet_name in wb.sheetnames:
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        conceptos = _extract_matrix_apu_blocks(rows)
        if conceptos:
            return _agrupar_duplicados(conceptos)
    for sheet_name in wb.sheetnames:
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        conceptos = _extract_formato_apu(rows)
        if conceptos:
            return _agrupar_duplicados(conceptos)
    try:
        wb_probe = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        if _is_construdata_matrix_workbook(wb_probe):
            wb_probe.close()
            matrices = load_construdata_matrices(filepath)
            if matrices:
                if cache_key is not None:
                    _BENCHMARK_CACHE.clear()
                    _BENCHMARK_CACHE[cache_key] = dict(matrices)
                return matrices
        wb_probe.close()
    except Exception:
        pass

    try:
        conceptos_v1 = extract_template_v1(filepath)
        if conceptos_v1 and len(conceptos_v1) >= 5:
            return conceptos_v1
    except Exception:
        pass
    for sheet_name in wb.sheetnames:
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        conceptos = _extract_formato_presupuesto(rows)
        if len(conceptos) >= 5:
            return _agrupar_duplicados(conceptos)
        conceptos = _extract_formato_headers(rows)
        if len(conceptos) >= 3:
            return _agrupar_duplicados(conceptos)
    raise ValueError(
        f"No se pudo extraer conceptos del archivo: {Path(filepath).name}. "
        "Verifica que sea una matriz/APU válida."
    )

def _es_codigo_concepto(val) -> bool:
    """Detecta si un valor parece un código de concepto (COM-PRE-01, A01, PR01, etc.)"""
    if not val or not isinstance(val, str):
        return False
    val = val.strip()
    # Códigos COM- estándar
    if re.match(r'^COM-[A-Z]{2,5}-\d{1,3}', val):
        return True
    # Códigos cortos tipo PR01, DD01, AN01, EM01, LA01, ES01, BA01, MA01, AC01
    if re.match(r'^[A-Z]{2,4}\d{2,3}$', val):
        return True
    return False


def _extract_formato_presupuesto(rows) -> list:
    """
    Formato: columnas fijas → [Código, Descripción, Unidad, Cantidad, P.U., Importe, ...]
    Ejemplo: PPT_Construccion_Expansion_Comedor (PROIINCSA/RAR)
    """
    conceptos = []
    seccion = ""

    for i, row in enumerate(rows):
        cod = row[0]
        if not cod:
            continue

        # Detectar sección
        if isinstance(cod, str) and not _es_codigo_concepto(cod) and row[1] and isinstance(row[1], str):
            if 'TOTAL' not in str(row[1]).upper():
                seccion = str(row[1])
            continue

        if not _es_codigo_concepto(cod):
            continue

        # Buscar columna de P.U. (puede estar en col 4, 5, o 6)
        pu = None
        pu_col = None
        for col_idx in [4, 5, 6]:
            if col_idx < len(row) and isinstance(row[col_idx], (int, float)) and row[col_idx] > 0:
                pu = row[col_idx]
                pu_col = col_idx
                break

        if not pu:
            continue

        # Descripción puede estar en múltiples filas
        desc_parts = [str(row[1]) if row[1] else '']
        for j in range(i+1, min(i+8, len(rows))):
            r2 = rows[j]
            if r2[0] is None and r2[1] and isinstance(r2[1], str) and not _es_codigo_concepto(r2[1]):
                desc_parts.append(r2[1])
            else:
                break

        total = row[pu_col + 1] if pu_col + 1 < len(row) and isinstance(row[pu_col + 1], (int, float)) else None
        cant  = row[3] if len(row) > 3 and isinstance(row[3], (int, float)) else 0
        uni   = row[2] if len(row) > 2 else ''

        conceptos.append({
            'clave':    str(cod).strip(),
            'desc':     ' '.join(desc_parts).strip()[:400],
            'unidad':   str(uni) if uni else '',
            'cantidad': cant or 0,
            'pu':       pu,
            'total':    total or (pu * cant if cant else pu),
            'seccion':  seccion,
        })

    return conceptos



def _extract_formato_apu(rows) -> list:
    """
    Formato CIOC / APU por bloques con filas 'Clave:' y etiquetas laterales.
    """
    conceptos = []

    for i, row in enumerate(rows):
        first = _clean_text(row[0] if len(row) > 0 else "")
        if first != 'Clave:':
            continue

        clave = _clean_text(row[4] if len(row) > 4 else "")
        if not clave:
            continue

        unidad = _clean_text(row[21] if len(row) > 21 else "")
        desc = ""
        if i >= 2 and _clean_text(rows[i - 2][0] if len(rows[i - 2]) > 0 else ""):
            desc = _clean_text(rows[i - 2][0])[:400]

        cantidad = pu = total = None
        end = min(i + 18, len(rows))
        for j in range(i + 1, end):
            r = rows[j]
            lbl = _clean_text(r[16] if len(r) > 16 else "")
            val = _as_float(r[21] if len(r) > 21 else None)
            if lbl == 'Cantidad:' and val is not None:
                cantidad = val
            if lbl == 'Precio unitario:' and val is not None:
                pu = val
            if lbl == 'Total' and val is not None:
                total = val

            # fallback horizontal simple
            lbl0 = _clean_text(r[0] if len(r) > 0 else "")
            val4 = _as_float(r[4] if len(r) > 4 else None)
            if lbl0 == 'Cantidad:' and cantidad is None and val4 is not None:
                cantidad = val4
            if lbl0 == 'Precio unitario:' and pu is None and val4 is not None:
                pu = val4

        if pu is None:
            continue

        conceptos.append({
            'clave': clave,
            'desc': desc,
            'unidad': unidad,
            'cantidad': cantidad if cantidad is not None else 1,
            'pu': pu,
            'total': total if total is not None else pu,
            'seccion': '',
        })

    return conceptos


def _extract_formato_headers(rows) -> list:
    """
    Fallback: busca la fila de encabezados y mapea columnas por nombre.
    Detecta variaciones: "P.U.", "Precio Unit.", "P. Unitario", "Importe", etc.
    """
    header_row_idx = None
    col_map = {}

    header_keywords = {
        'codigo':      ['código', 'clave', 'partida', 'part.', 'cod', 'item'],
        'descripcion': ['descripción', 'descripcion', 'concepto', 'trabajo'],
        'unidad':      ['uni', 'unidad', 'u.m.'],
        'cantidad':    ['cant', 'cantidad', 'volumen'],
        'pu':          ['p.u.', 'precio unit', 'precio unitario', 'p. unit'],
        'total':       ['importe', 'total', 'monto', 'precio total'],
    }

    for i, row in enumerate(rows[:50]):  # Buscar header en primeras 50 filas
        row_lower = [str(c).lower().strip() if c else '' for c in row]
        matches = 0
        temp_map = {}
        for field, keywords in header_keywords.items():
            for j, cell in enumerate(row_lower):
                if any(kw in cell for kw in keywords):
                    temp_map[field] = j
                    matches += 1
                    break
        if matches >= 3:
            header_row_idx = i
            col_map = temp_map
            break

    if not col_map or 'pu' not in col_map:
        return []

    conceptos = []
    for row in rows[header_row_idx+1:]:
        cod_val = row[col_map.get('codigo', 0)] if col_map.get('codigo') is not None else None
        pu_val  = row[col_map['pu']] if col_map.get('pu') is not None else None

        if not _es_codigo_concepto(cod_val) and not cod_val:
            continue
        if not isinstance(pu_val, (int, float)) or pu_val <= 0:
            continue

        cant  = row[col_map['cantidad']] if col_map.get('cantidad') is not None else 0
        total = row[col_map['total']]    if col_map.get('total')    is not None else None
        desc  = row[col_map['descripcion']] if col_map.get('descripcion') is not None else ''
        uni   = row[col_map['unidad']]   if col_map.get('unidad')   is not None else ''

        conceptos.append({
            'clave':    str(cod_val).strip() if cod_val else f"SIN_CODIGO_{len(conceptos)+1:03d}",
            'desc':     str(desc)[:400] if desc else '',
            'unidad':   str(uni) if uni else '',
            'cantidad': cant if isinstance(cant, (int,float)) else 0,
            'pu':       pu_val,
            'total':    total if isinstance(total, (int,float)) else pu_val * (cant or 0),
            'seccion':  '',
        })

    return conceptos



def _agrupar_duplicados(conceptos: list) -> dict:
    """
    Conserva cada servicio. Si una clave realmente se repite, se preserva
    agregando un sufijo incremental en vez de consolidar por precio/cantidad.
    """
    resultado = {}
    counts = defaultdict(int)
    for c in conceptos:
        base = str(c.get('clave') or '').strip() or f"SERVICIO_{len(resultado)+1:03d}"
        key = base
        if key in resultado:
            counts[base] += 1
            key = f"{base}__{counts[base]+1}"
        else:
            counts[base] = 1
        item = dict(c)
        item['clave'] = key
        item['clave_base'] = base
        resultado[key] = item
    return resultado


def extract_orden_catalogo(filepath: str) -> list:
    """
    Lee el catálogo base de Nestlé y extrae el orden oficial de conceptos.
    Devuelve lista de dicts: [{tipo, clave, desc, uni, cant, seccion}]
    """
    # intentar Template v1 primero
    try:
        conceptos_v1=extract_template_v1(filepath)
        if conceptos_v1:
            return conceptos_v1
    except Exception:
        pass

    wb = openpyxl.load_workbook(filepath, data_only=True)
    orden = []
    seccion_actual = ""

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        conceptos_encontrados = 0

        for i, row in enumerate(rows):
            cod = row[0]
            if not cod:
                continue

            # Detectar sección
            if isinstance(cod, str) and not _es_codigo_concepto(cod) and row[1]:
                if 'TOTAL' not in str(row[1]).upper():
                    seccion_actual = str(row[1])
                    orden.append({'tipo': 'seccion', 'clave': str(cod), 'desc': seccion_actual})
                continue

            if not _es_codigo_concepto(cod):
                continue

            # Buscar cantidad y unidad
            uni  = row[2] if len(row) > 2 else ''
            cant = row[3] if len(row) > 3 and isinstance(row[3], (int,float)) else 0
            desc = str(row[1]) if row[1] else ''

            orden.append({
                'tipo':    'concepto',
                'clave':   str(cod).strip(),
                'desc':    desc[:400],
                'uni':     str(uni) if uni else '',
                'cant':    cant,
                'seccion': seccion_actual,
            })
            conceptos_encontrados += 1

        if conceptos_encontrados >= 5:
            break  # Hoja correcta encontrada

    return orden


# ══════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN DEL COMPARATIVO EXCEL
# ══════════════════════════════════════════════════════════════

# Paleta de colores por proveedor (hasta 5)
PROV_COLORS = ["2E75B6", "375623", "843C0C", "7B2C7A", "0E4D6E"]
PROV_LIGHT  = ["DEEAF1", "E2EFDA", "FCE4D6", "F3D9F3", "D6EAF8"]

C_AZUL_OSC  = "1F3864"
C_GRIS_SEC  = "D6DCE4"
C_VERDE_OK  = "C6EFCE"
C_ROJO_MAL  = "FFC7CE"
C_AMARILLO  = "FFEB9C"
C_BLANCO    = "FFFFFF"
C_GRIS_FIL  = "F5F5F5"
# Colores adicionales usados por el resumen ejecutivo IA.
# Se definen aquí para evitar NameError si el módulo se carga desde Railway/backend.
C_GRIS_TXT  = "666666"
C_NEGRO_TXT = "1F1F1F"
C_NARANJA   = "F4B183"
C_AZUL_CLARO = "D9EAF7"

def _fill(c):   return PatternFill("solid", start_color=c, end_color=c)
def _brd():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)
def _fnt(bold=False, color="000000", size=9, name="Calibri", italic=False, underline=None):
    # STEP40 hotfix: aceptar italic/underline para evitar TypeError en hojas ejecutivas.
    return Font(bold=bold, color=color, size=size, name=name, italic=italic, underline=underline)
def _aln(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)



C_AMARILLO_PMD = "F5F527"


def _copy_cell_style(src, dst):
    if src.has_style:
        dst._style = src._style.copy()
    if src.font:
        dst.font = src.font.copy()
    if src.fill:
        dst.fill = src.fill.copy()
    if src.border:
        dst.border = src.border.copy()
    if src.alignment:
        dst.alignment = src.alignment.copy()
    if src.number_format:
        dst.number_format = src.number_format


def _copy_row_style(ws, src_row, dst_row, start_col=1, end_col=14):
    for col in range(start_col, end_col + 1):
        _copy_cell_style(ws.cell(src_row, col), ws.cell(dst_row, col))
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def _clear_range_values(ws, min_row, max_row, min_col=1, max_col=14):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                cell.value = None


def _set_font_size_workbook(wb, size=8):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                f = cell.font.copy(size=size) if cell.font else Font(size=size, name="Calibri")
                cell.font = f


def _market_component_total(concepto):
    """Costo directo de mercado para el concepto.

    Prioridad Paso 7:
    1) Si existe match Construdata por matriz, usa el PU directo Construdata
       asociado al concepto homologado. Ese valor se tratará como costo directo
       base y el reporte aplicará SIEMPRE 25% de indirecto para mercado.
    2) Si no hay match Construdata, cae al cálculo heredado por materiales/MO/equipo.

    Nota: esta función devuelve costo DIRECTO. El mercado final es directo * 1.25.
    """
    comp = concepto.get("construdata_matrix_comparison") or {}
    cd_pu = comp.get("construdata_pu")
    if isinstance(cd_pu, (int, float)) and cd_pu > 0:
        return float(cd_pu)

    total = 0.0
    for mm in concepto.get("materiales_benchmark_matches") or []:
        sel = (mm.get("match") or {}).get("selected") or {}
        inp = mm.get("input") or {}
        market_price = None
        resolution = (mm.get("match") or {}).get("market_resolution") or mm.get("market_resolution") or {}
        if isinstance(resolution.get("effective_market_price"), (int, float)):
            market_price = resolution.get("effective_market_price")
        elif isinstance(sel.get("precio"), (int, float)):
            market_price = sel.get("precio")
        if isinstance(inp.get("factor"), (int, float)) and isinstance(market_price, (int, float)):
            total += float(inp["factor"]) * float(market_price)
    for lm in concepto.get("mano_obra_benchmark_matches") or []:
        mt = lm.get("match") or {}
        inp = lm.get("input") or {}
        factor_cmp = mt.get("factor_comparable") if isinstance(mt.get("factor_comparable"), (int, float)) else inp.get("factor")
        precio_cmp = mt.get("precio_tabulador_comparable") if isinstance(mt.get("precio_tabulador_comparable"), (int, float)) else mt.get("precio_tabulador")
        if isinstance(factor_cmp, (int, float)) and isinstance(precio_cmp, (int, float)):
            total += float(factor_cmp) * float(precio_cmp)
    for em in concepto.get("equipo_benchmark_matches") or []:
        sel = (em.get("match") or {}).get("selected") or {}
        inp = em.get("input") or {}
        if isinstance(inp.get("factor"), (int, float)) and isinstance(sel.get("precio_ref"), (int, float)):
            total += float(inp["factor"]) * float(sel["precio_ref"])
    return total


try:
    MARKET_INDIRECT_PCT = float(os.getenv("MARKET_INDIRECT_PCT", "0.25") or "0.25")
except Exception:
    MARKET_INDIRECT_PCT = 0.25


def _market_pu_with_indirect(concepto):
    direct = _market_component_total(concepto)
    if isinstance(direct, (int, float)) and direct > 0:
        return round(float(direct) * (1.0 + MARKET_INDIRECT_PCT), 6)
    return None


def _concept_total_amount(concepto):
    cantidad = concepto.get('cantidad') if isinstance(concepto.get('cantidad'), (int, float)) else 1
    total = concepto.get('total')
    if isinstance(total, (int, float)):
        return float(total)
    pu = concepto.get('pu')
    if isinstance(pu, (int, float)):
        return float(pu) * float(cantidad)
    return None


def _concept_market_total_amount(concepto):
    cantidad = concepto.get('cantidad') if isinstance(concepto.get('cantidad'), (int, float)) else 1
    pu_market = _market_pu_with_indirect(concepto)
    if isinstance(pu_market, (int, float)):
        return float(pu_market) * float(cantidad)
    return None


def _concept_market_status(provider_total, market_total):
    if not isinstance(provider_total, (int, float)) or not isinstance(market_total, (int, float)) or market_total == 0:
        return 'SIN_REFERENCIA'
    delta = (float(provider_total) - float(market_total)) / float(market_total)
    if abs(delta) <= 0.05:
        return 'OK'
    if delta > 0.15:
        return 'SOBRE_MERCADO_ALTO'
    if delta > 0:
        return 'SOBRE_MERCADO'
    if delta < -0.20:
        return 'BAJO_MERCADO_REVISAR'
    return 'BAJO_MERCADO'


def _concept_indirect_amount_and_pct(concepto):
    cd = concepto.get("costo_directo")
    pct = concepto.get("indirecto_pct")
    amt = concepto.get("indirecto")
    if isinstance(pct, (int, float)):
        pct = float(pct)
        if pct > 1.5:
            pct = pct / 100.0
    else:
        pct = None
    if isinstance(amt, (int, float)) and isinstance(cd, (int, float)) and cd:
        cand_pct = float(amt) / float(cd)
        if pct is None and 0 < cand_pct <= 1.0:
            pct = cand_pct
    if pct is None:
        for candidate_amt in [concepto.get("subtotal2"), concepto.get("subtotal1")]:
            if isinstance(candidate_amt, (int, float)) and isinstance(cd, (int, float)) and cd:
                cand_pct = float(candidate_amt) / float(cd)
                if 0 < cand_pct <= 0.6:
                    amt = candidate_amt
                    pct = cand_pct
                    break
    if pct is None and isinstance(cd, (int, float)) and cd:
        pu = float(concepto.get("pu") or 0)
        residual = pu - float(cd) - float(concepto.get("utilidad") or 0) - float(concepto.get("financiamiento") or 0)
        cand_pct = residual / float(cd) if residual > 0 else None
        if cand_pct is not None and 0 < cand_pct <= 0.6:
            amt = residual
            pct = cand_pct
    if amt is None and isinstance(pct, (int, float)) and isinstance(cd, (int, float)):
        amt = float(cd) * float(pct)
    return amt, pct


def _select_blue_rows(service_rows):
    ordered = sorted(service_rows, key=lambda x: x.get("pct_part") or 0, reverse=True)
    acc = 0.0
    selected = []
    for row in ordered:
        if acc >= 0.80:
            break
        selected.append(row["excel_row"])
        acc += row.get("pct_part") or 0.0
    if not selected and ordered:
        selected = [ordered[0]["excel_row"]]
        acc = ordered[0].get("pct_part") or 0.0
    return set(selected), acc



def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    """
    Construye el PMD de 1 proveedor desde cero.
    NO inserta datos sobre una plantilla preformateada para evitar heredar
    colores, merges o estilos de datos ficticios.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativa"
    ws_det = wb.create_sheet("Detalle")

    # Eliminar hoja default extra si llegara a existir con otro nombre
    for name in list(wb.sheetnames):
        if name not in {"Comparativa", "Detalle"}:
            pass

    conceptos_items = list(dato["conceptos"].items())
    total_contratista = round(sum(v.get("total") or 0 for _, v in conceptos_items if isinstance(v.get("total"), (int, float))), 2)

    # ----- ESTILO Y LAYOUT COMPARATIVA -----
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A9'
    ws.column_dimensions['A'].width = 9.33
    ws.column_dimensions['B'].width = 52.33
    ws.column_dimensions['C'].width = 7.55
    ws.column_dimensions['D'].width = 10.11
    ws.column_dimensions['E'].width = 12.33
    ws.column_dimensions['F'].width = 17.89
    ws.column_dimensions['G'].width = 11.00
    ws.column_dimensions['H'].width = 12.00
    ws.column_dimensions['I'].width = 10.66
    ws.column_dimensions['J'].width = 14.11

    for r in range(1, 500):
        ws.row_dimensions[r].height = 15.0

    # Encabezado superior sin plantilla
    ws.merge_cells('E2:F2'); ws['E2'] = 'Cliente:'; ws['E2'].font = _fnt(bold=True, size=8)
    ws.merge_cells('G2:J2'); ws['G2'] = meta.get('cliente', '') or ''
    ws.merge_cells('E3:F3'); ws['E3'] = 'Proyecto:'; ws['E3'].font = _fnt(bold=True, size=8)
    ws.merge_cells('G3:J3'); ws['G3'] = meta.get('proyecto', '') or ''
    ws.merge_cells('E4:F4'); ws['E4'] = 'Archivo:'; ws['E4'].font = _fnt(bold=True, size=8)
    ws.merge_cells('G4:J4'); ws['G4'] = Path(meta.get('archivo_fuente') or dato.get('nombre') or '').name
    for ref in ['G2', 'G3', 'G4']:
        ws[ref].alignment = _aln('left', 'center')
        ws[ref].font = _fnt(size=8)

    # Franja proveedor / mercado
    ws.merge_cells('E6:H6')
    ws['E6'] = dato.get('nombre') or provider_analysis.get('provider_name') or 'Proveedor'
    ws['E6'].fill = _fill(C_AZUL_OSC)
    ws['E6'].font = _fnt(bold=True, color=C_BLANCO, size=8)
    ws['E6'].alignment = _aln('center', 'center')
    ws['E6'].border = _brd()

    ws.merge_cells('I6:J6')
    ws['I6'] = 'MERCADO'
    ws['I6'].fill = _fill(C_AMARILLO_PMD)
    ws['I6'].font = _fnt(bold=True, color='000000', size=8)
    ws['I6'].alignment = _aln('center', 'center')
    ws['I6'].border = _brd()

    headers = ['Part.', 'D e s c r i p c i ó n', 'Unidad', 'Cant.', 'P.U.', 'Importe', '% Part.', '% ajuste', 'P.U.', 'Importe']
    for cidx, h in enumerate(headers, start=1):
        cell = ws.cell(7, cidx, h)
        cell.font = _fnt(bold=True, color=('000000' if cidx in (9, 10) else C_BLANCO), size=8)
        cell.fill = _fill(C_AMARILLO_PMD if cidx in (9, 10) else C_AZUL_OSC)
        cell.alignment = _aln('center', 'center', wrap=True)
        cell.border = _brd()
    ws.row_dimensions[7].height = 15.6

    # ----- DATOS -----
    row = 9
    service_rows = []
    current_section = None
    for idx, (clave, c) in enumerate(conceptos_items, start=1):
        seccion = _clean_text(c.get('seccion'))
        if seccion and seccion != current_section:
            ws.cell(row, 1, f'{seccion}')
            ws.cell(row, 1).font = _fnt(bold=True, size=8)
            ws.cell(row, 1).alignment = _aln('left', 'center')
            for col in range(1, 11):
                ws.cell(row, col).border = _brd()
            ws.row_dimensions[row].height = 15.0
            current_section = seccion
            row += 1

        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'), (int, float)) and c.get('cantidad') is not None else 1
        pu = c.get('pu') if isinstance(c.get('pu'), (int, float)) else None
        importe_decl = c.get('total') if isinstance(c.get('total'), (int, float)) else None
        importe = importe_decl if importe_decl is not None else (pu * cantidad if isinstance(pu, (int, float)) and isinstance(cantidad, (int, float)) else None)

        market_direct = _market_component_total(c)
        market_pu = round(_market_pu_with_indirect(c), 2) if _market_pu_with_indirect(c) else None
        market_importe = round(market_pu * cantidad, 2) if market_pu is not None else None
        pct_part = (importe / total_contratista) if total_contratista and isinstance(importe, (int, float)) else None
        pct_adj = ((market_pu / pu) - 1) if (isinstance(market_pu, (int, float)) and isinstance(pu, (int, float)) and pu) else None

        vals = [clave, c.get('desc'), c.get('unidad'), cantidad, pu, importe, pct_part, pct_adj, market_pu, market_importe]
        for col, val in enumerate(vals, start=1):
            cell = ws.cell(row, col, value=val)
            cell.border = _brd()
            cell.font = _fnt(size=8)
            cell.alignment = _aln('left', 'center', wrap=True) if col == 2 else _aln('center', 'center')
            if col in (5, 6, 9, 10) and isinstance(val, (int, float)):
                cell.number_format = '$#,##0.00'
            if col in (7, 8) and isinstance(val, (int, float)):
                cell.number_format = '0.00%'
        # Altura segun descripcion
        desc = str(c.get('desc') or '')
        lines = max(1, min(4, (len(desc) // 70) + 1))
        ws.row_dimensions[row].height = 13.8 * lines
        service_rows.append({'excel_row': row, 'pct_part': pct_part or 0.0, 'concepto': c})
        row += 1

    blue_rows, blue_pct = _select_blue_rows(service_rows)
    blue_fill = _fill(PROV_LIGHT[0])
    for r in blue_rows:
        for cidx in range(1, 11):
            ws.cell(r, cidx).fill = blue_fill

    total_row = row + 1
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=4)
    ws.cell(total_row, 1, 'Importe Total')
    ws.cell(total_row, 1).alignment = _aln('center', 'center')
    for cidx in range(1, 11):
        ws.cell(total_row, cidx).border = _brd()
        ws.cell(total_row, cidx).font = _fnt(bold=True, size=8)
    ws.cell(total_row, 6, total_contratista)
    ws.cell(total_row, 6).number_format = '$#,##0.00'
    market_total = round(sum((r.get('concepto') and (_concept_market_total_amount(r['concepto']) or 0) or 0) for r in service_rows), 2)
    ws.cell(total_row, 10, market_total)
    ws.cell(total_row, 10).number_format = '$#,##0.00'

    pct_global_row = total_row + 1
    ws.merge_cells(start_row=pct_global_row, start_column=1, end_row=pct_global_row, end_column=4)
    ws.cell(pct_global_row, 1, '% de ajuste Global')
    ws.cell(pct_global_row, 1).alignment = _aln('center', 'center')
    ajuste_global = ((market_total / total_contratista) - 1) if total_contratista and market_total else None
    ws.cell(pct_global_row, 6, ajuste_global)
    if isinstance(ajuste_global, (int, float)):
        ws.cell(pct_global_row, 6).number_format = '0.00%'
    for cidx in range(1, 11):
        ws.cell(pct_global_row, cidx).border = _brd()
        ws.cell(pct_global_row, cidx).font = _fnt(size=8)

    comment_label_row = total_row + 3
    ws.merge_cells(start_row=comment_label_row, start_column=1, end_row=comment_label_row, end_column=4)
    ws.cell(comment_label_row, 1, 'Comentarios:')
    ws.cell(comment_label_row, 1).font = _fnt(bold=True, size=8)

    note1_row = total_row + 4
    note2_row = total_row + 7
    note3_row = total_row + 10
    for rr in [note1_row, note2_row, note3_row]:
        ws.merge_cells(start_row=rr, start_column=5, end_row=rr + (2 if rr != note3_row else 0), end_column=8)
        ws.cell(rr, 5).alignment = _aln('left', 'center', wrap=True)
        ws.cell(rr, 5).font = _fnt(size=8)

    ws.cell(note1_row, 5, f'Las partidas sombreadas en color azul representan el {blue_pct*100:.2f}% de la propuesta, habiendo áreas de oportunidad de revisión en lo siguiente:')
    # Indirecto global: priorizar promedio de indirectos explícitos declarados
    pcts = [(_concept_indirect_amount_and_pct(c)[1]) for _, c in conceptos_items]
    pcts = [p for p in pcts if isinstance(p, (int, float))]
    indirecto_pct = (sum(pcts) / len(pcts)) if pcts else None
    if indirecto_pct is not None:
        competitividad = 'siendo competitivo' if indirecto_pct <= 0.25 else 'NO siendo competitivo'
        ws.cell(note2_row, 5, f'Considera un indirecto global del {indirecto_pct*100:.2f} %, {competitividad} vs. el Mercado de 25.00 %.')
    else:
        ws.cell(note2_row, 5, 'No se pudo determinar de forma confiable el indirecto global del contratista.')
    top_driver = (provider_analysis.get('drivers') or [{}])[0]
    driver_txt = top_driver.get('concepto') or top_driver.get('tipo') or 'Sin hallazgo estructurado'
    ws.cell(note3_row, 5, f'Driver principal detectado: {driver_txt}.')

    # Bordes base visibles en celdas clave de comentarios
    for rr in range(comment_label_row, note3_row + 2):
        for cc in range(1, 11):
            ws.cell(rr, cc).font = _fnt(size=8)
            if cc in range(5, 9) or (rr in [comment_label_row, total_row, pct_global_row] and cc <= 6):
                ws.cell(rr, cc).border = _brd()

    # ----- DETALLE -----
    ws_det.sheet_view.showGridLines = False
    for col, width in {'A':14, 'B':56, 'C':10, 'D':13, 'E':6, 'F':10, 'G':14, 'H':8}.items():
        ws_det.column_dimensions[col].width = width
    ws_det.merge_cells('A1:H1')
    ws_det['A1'] = 'DETALLE TÉCNICO PROVEEDOR VS MERCADO'
    ws_det['A1'].fill = _fill(C_AZUL_OSC)
    ws_det['A1'].font = _fnt(bold=True, color=C_BLANCO, size=8)
    ws_det['A1'].alignment = _aln('center', 'center')
    headers_det = ['Código', 'Concepto', 'Unidad', 'P. Unitario', 'Op.', 'Cantidad', 'Importe', '%']
    for i, h in enumerate(headers_det, start=1):
        cell = ws_det.cell(3, i, h)
        cell.fill = _fill(C_AZUL_OSC)
        cell.font = _fnt(bold=True, color=C_BLANCO, size=8)
        cell.border = _brd()
        cell.alignment = _aln('center', 'center')
    det_row = 4
    for clave, c in conceptos_items:
        ws_det.merge_cells(start_row=det_row, start_column=1, end_row=det_row, end_column=8)
        cell = ws_det.cell(det_row, 1, f'Análisis: {clave}  |  {c.get("desc", "")}')
        cell.fill = _fill(C_GRIS_SEC)
        cell.font = _fnt(bold=True, size=8)
        cell.border = _brd()
        cell.alignment = _aln('left', 'center', wrap=True)
        det_row += 1
        for title, items in [('MATERIALES', c.get('materiales_items') or []), ('MANO DE OBRA', c.get('mano_obra_items') or []), ('EQUIPO Y HERRAMIENTA', c.get('equipo_items') or [])]:
            ws_det.merge_cells(start_row=det_row, start_column=1, end_row=det_row, end_column=8)
            hdr = ws_det.cell(det_row, 1, title)
            hdr.fill = _fill(C_GRIS_FIL)
            hdr.font = _fnt(bold=True, size=8)
            hdr.border = _brd()
            hdr.alignment = _aln('left', 'center')
            det_row += 1
            for item in items:
                vals = [item.get('codigo'), item.get('descripcion'), item.get('unidad'), item.get('precio_base'), '*', item.get('factor'), item.get('importe'), None]
                for col, val in enumerate(vals, start=1):
                    cc = ws_det.cell(det_row, col, val)
                    cc.border = _brd()
                    cc.font = _fnt(size=8)
                    cc.alignment = _aln('left', 'center', wrap=True) if col == 2 else _aln('center', 'center')
                    if col in (4, 7) and isinstance(val, (int, float)):
                        cc.number_format = '$#,##0.00'
                det_row += 1
        det_row += 1

    # Aux sheets
    _build_market_materials_sheet(wb, [dato])
    if 'Mercado materiales1' in wb.sheetnames:
        wb['Mercado materiales1'].title = 'Mercado materiales'
    _build_market_labor_sheet(wb, [dato])
    if 'Mercado mano de obra1' in wb.sheetnames:
        wb['Mercado mano de obra1'].title = 'Mercado mano de obra'
    _build_drivers_sheet(wb, [provider_analysis])
    if 'Hallazgos PMD1' in wb.sheetnames:
        wb['Hallazgos PMD1'].title = 'Hallazgos PMD'
    _build_construdata_matching_sheet(wb, [dato])
    try:
        _build_homologacion_ia_sheet(wb, [dato])
    except NameError:
        pass
    _build_construdata_matrix_detail_sheet(wb, [dato])
    _build_totales_mercado_sheet(wb, [dato])
    _build_hallazgos_ejecutivos_sheet(wb, [dato])
    _build_faltantes_matriz_base_sheet(wb, [dato])
    _build_elementos_fuera_matriz_sheet(wb, [dato])

    # Orden final de tabs
    desired = ['Comparativa', 'Detalle', 'Totales mercado', 'Hallazgos Ejecutivos', 'Faltantes matriz base', 'Elementos fuera de matriz', 'Mercado materiales', 'Mercado mano de obra', 'Hallazgos PMD', 'Construdata matching', 'Homologacion IA', 'Construdata matrices']
    wb._sheets = [wb[n] for n in desired if n in wb.sheetnames]
    _set_font_size_workbook(wb, 8)
    wb.save(output)
    return output





# -----------------------------------------------------------------------------
# Construdata / Neodata como MATRICES POR CONCEPTO
# -----------------------------------------------------------------------------
# Este bloque corrige el problema central del benchmark anterior: antes el Excel
# se interpretaba como catalogo plano (codigo -> precio). Ahora, si el archivo
# contiene hojas exportadas desde MDF/SQL Server como "conceptos" y
# "matrices_detalle", se construye una referencia real:
#
#   CodigoConcepto -> {descripcion, unidad, pu_mercado, matriz:[insumos...]}
#
# La funcion extract_precios_nacional() mantiene compatibilidad hacia atras: si
# el Excel NO tiene matrices, cae al parser plano existente.

CONSTRUDATA_MATRIX_SHEETS = {"matrices_detalle", "base_comparacion"}
CONSTRUDATA_CONCEPT_SHEETS = {"conceptos"}


def _norm_header(value):
    """Normaliza encabezados de Excel para mapear columnas flexibles."""
    return _normalize_text(value).replace(" ", "")


def _build_header_map(row):
    return {_norm_header(value): idx for idx, value in enumerate(row) if _clean_text(value)}


def _first_existing_col(header_map, candidates):
    for name in candidates:
        key = _norm_header(name)
        if key in header_map:
            return header_map[key]
    return None


def _cell(row, idx, default=None):
    if idx is None or idx >= len(row):
        return default
    value = row[idx]
    return default if value is None else value


def _is_construdata_matrix_workbook(wb):
    names = {_normalize_text(s).replace(" ", "_") for s in wb.sheetnames}
    names_compact = {_norm_header(s) for s in wb.sheetnames}
    return (
        "matrices_detalle" in names
        or "base_comparacion" in names
        or "matricesdetalle" in names_compact
        or "basecomparacion" in names_compact
    )


def _find_sheet_name(wb, candidates):
    by_norm = {_normalize_text(s).replace(" ", "_"): s for s in wb.sheetnames}
    by_compact = {_norm_header(s): s for s in wb.sheetnames}
    for c in candidates:
        cn = _normalize_text(c).replace(" ", "_")
        cc = _norm_header(c)
        if cn in by_norm:
            return by_norm[cn]
        if cc in by_compact:
            return by_compact[cc]
    return None


def load_construdata_matrices(filepath: str) -> dict:
    """
    Lee el Excel maestro generado desde el MDF de Neodata/Construdata y retorna
    una estructura de benchmark por concepto.

    Formato esperado recomendado:
      - Hoja conceptos
      - Hoja matrices_detalle

    Columnas minimas en matrices_detalle/base_comparacion:
      CodigoConcepto, Concepto, ConceptoUnidad, ConceptoPrecioUnitario,
      CodigoInsumo, Insumo, UnidadInsumo, TipoInsumo,
      CantidadMatriz, CostoInsumo, ImporteCalculado

    Retorna dict compatible con el resto del sistema:
      {
        "10301-001": {
          "clave": "10301-001",
          "desc": "...",
          "unidad": "M2",
          "pu_mercado": 123.45,
          "matriz": [...],
          "matriz_stats": {...}
        },
        "__materials_index__": [...],
        "__construdata_matrices__": {...},
        "__benchmark_kind__": "construdata_matrices"
      }
    """
    src = Path(filepath)
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    if not _is_construdata_matrix_workbook(wb):
        return {}

    conceptos = {}
    materials_index = []
    seen_materials = set()

    # 1) Cargar metadata de conceptos si existe la hoja conceptos.
    conceptos_sheet_name = _find_sheet_name(wb, ["conceptos"])
    if conceptos_sheet_name:
        ws = wb[conceptos_sheet_name]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if header:
            hm = _build_header_map(header)
            c_codigo = _first_existing_col(hm, ["CodigoConcepto", "Codigo", "Clave", "IdCodigoConcepto"])
            c_desc = _first_existing_col(hm, ["Concepto", "Descripcion", "Descripción", "ConceptoDescripcionLarga"])
            c_unidad = _first_existing_col(hm, ["ConceptoUnidad", "Unidad", "U.M.", "UM"])
            c_pu = _first_existing_col(hm, ["ConceptoPrecioUnitario", "Precio", "PrecioUnitario", "PU", "PU Mercado"])
            c_idexp = _first_existing_col(hm, ["IdExpInsConcepto", "IdExpIns"])
            for row in rows:
                codigo = _clean_text(_cell(row, c_codigo))
                if not codigo:
                    continue
                desc = _clean_text(_cell(row, c_desc))
                unidad = _clean_text(_cell(row, c_unidad))
                pu = _as_float(_cell(row, c_pu))
                conceptos[codigo] = {
                    "clave": codigo,
                    "codigo": codigo,
                    "id_expins_concepto": _cell(row, c_idexp),
                    "desc": desc[:300],
                    "desc_larga": desc[:600],
                    "unidad": unidad,
                    "pu_mercado": pu,
                    "tipo": "CONCEPTO",
                    "matriz": [],
                }

    # 2) Cargar matrices. Preferimos matrices_detalle; si no existe, base_comparacion.
    matrix_sheet_name = _find_sheet_name(wb, ["matrices_detalle", "base_comparacion"])
    if not matrix_sheet_name:
        return {}

    ws = wb[matrix_sheet_name]
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        return {}
    hm = _build_header_map(header)

    c_codigo_concepto = _first_existing_col(hm, ["CodigoConcepto", "ClaveConcepto", "Codigo Matriz", "Codigo"])
    c_id_concepto = _first_existing_col(hm, ["IdExpInsConcepto", "IdCodigoMatriz"])
    c_concepto = _first_existing_col(hm, ["Concepto", "DescripcionConcepto", "ConceptoDescripcionLarga"])
    c_unidad_concepto = _first_existing_col(hm, ["ConceptoUnidad", "UnidadConcepto"])
    c_pu_concepto = _first_existing_col(hm, ["ConceptoPrecioUnitario", "PrecioConcepto", "PrecioUnitarioConcepto", "PU"])

    c_renglon = _first_existing_col(hm, ["MatrizRenglon", "Renglon", "Renglón"])
    c_codigo_insumo = _first_existing_col(hm, ["CodigoInsumo", "ClaveInsumo", "Codigo Insumo"])
    c_insumo = _first_existing_col(hm, ["Insumo", "DescripcionInsumo", "InsumoDescripcionLarga"])
    c_unidad_insumo = _first_existing_col(hm, ["UnidadInsumo", "Unidad", "U.M.", "UM"])
    c_tipo = _first_existing_col(hm, ["TipoInsumo", "Tipo", "IdTipo"])
    c_cantidad = _first_existing_col(hm, ["CantidadMatriz", "Volumen", "Cantidad", "Rendimiento"])
    c_dividir = _first_existing_col(hm, ["Dividir"])
    c_precio = _first_existing_col(hm, ["CostoInsumo", "Costo", "Precio", "PrecioInsumo"])
    c_importe = _first_existing_col(hm, ["ImporteCalculado", "Importe", "Total"])
    c_expresion = _first_existing_col(hm, ["Expresion", "Expresión"])

    for row in rows:
        codigo_concepto = _clean_text(_cell(row, c_codigo_concepto))
        if not codigo_concepto:
            continue

        concepto_desc = _clean_text(_cell(row, c_concepto))
        concepto_unidad = _clean_text(_cell(row, c_unidad_concepto))
        concepto_pu = _as_float(_cell(row, c_pu_concepto))

        if codigo_concepto not in conceptos:
            conceptos[codigo_concepto] = {
                "clave": codigo_concepto,
                "codigo": codigo_concepto,
                "id_expins_concepto": _cell(row, c_id_concepto),
                "desc": concepto_desc[:300],
                "desc_larga": concepto_desc[:600],
                "unidad": concepto_unidad,
                "pu_mercado": concepto_pu,
                "tipo": "CONCEPTO",
                "matriz": [],
            }
        else:
            if concepto_desc and not conceptos[codigo_concepto].get("desc"):
                conceptos[codigo_concepto]["desc"] = concepto_desc[:300]
                conceptos[codigo_concepto]["desc_larga"] = concepto_desc[:600]
            if concepto_unidad and not conceptos[codigo_concepto].get("unidad"):
                conceptos[codigo_concepto]["unidad"] = concepto_unidad
            if concepto_pu is not None and conceptos[codigo_concepto].get("pu_mercado") is None:
                conceptos[codigo_concepto]["pu_mercado"] = concepto_pu

        codigo_insumo = _clean_text(_cell(row, c_codigo_insumo))
        insumo_desc = _clean_text(_cell(row, c_insumo))
        # Puede haber conceptos sin matriz; no generamos renglon vacio.
        if not codigo_insumo and not insumo_desc:
            continue

        unidad_insumo = _clean_text(_cell(row, c_unidad_insumo))
        tipo_raw = _clean_text(_cell(row, c_tipo))
        cantidad = _as_float(_cell(row, c_cantidad))
        precio = _as_float(_cell(row, c_precio))
        importe = _as_float(_cell(row, c_importe))
        dividir = _cell(row, c_dividir)

        renglon = {
            "renglon": _cell(row, c_renglon),
            "tipo": tipo_raw,
            "tipo_kind": _material_kind_from_tipo(tipo_raw),
            "codigo": codigo_insumo,
            "clave": codigo_insumo,
            "descripcion": insumo_desc[:600],
            "unidad": unidad_insumo,
            "unidad_norm": normalize_unit(unidad_insumo),
            "cantidad": cantidad,
            "precio": precio,
            "precio_base": precio,
            "importe": importe,
            "dividir": dividir,
            "expresion": _clean_text(_cell(row, c_expresion)),
        }
        conceptos[codigo_concepto]["matriz"].append(renglon)

        # Mantener indice de materiales/insumos para compatibilidad con el motor actual.
        # No se limita a materiales porque algunos benchmarks antiguos usan este indice
        # para busqueda semantica de insumos; se excluye mano de obra para no mezclar con tabulador.
        if renglon["tipo_kind"] != "mano_obra" and (codigo_insumo or insumo_desc):
            mkey = (codigo_insumo, insumo_desc, unidad_insumo, precio)
            if mkey not in seen_materials:
                seen_materials.add(mkey)
                materials_index.append({
                    "clave": codigo_insumo,
                    "descripcion": insumo_desc[:600],
                    "descripcion_corta": insumo_desc[:300],
                    "unidad": unidad_insumo,
                    "unidad_norm": normalize_unit(unidad_insumo),
                    "precio": precio,
                    "tipo_insumo": tipo_raw,
                    "tipo_kind": renglon["tipo_kind"],
                    "fuente_precio": "Construdata matriz",
                    **_extract_material_attributes(insumo_desc, unidad_insumo),
                })

    # 3) Enriquecer estadisticas por concepto.
    for codigo, c in conceptos.items():
        matriz = c.get("matriz") or []
        total_matriz = sum(x.get("importe") or 0 for x in matriz if isinstance(x.get("importe"), (int, float)))
        por_tipo = defaultdict(float)
        for x in matriz:
            tipo = x.get("tipo_kind") or x.get("tipo") or "sin_tipo"
            if isinstance(x.get("importe"), (int, float)):
                por_tipo[tipo] += x.get("importe") or 0
        c["matriz_stats"] = {
            "renglones": len(matriz),
            "total_matriz": round(total_matriz, 6),
            "total_por_tipo": dict(por_tipo),
        }

    conceptos["__materials_index__"] = materials_index
    conceptos["__construdata_matrices__"] = {k: v for k, v in conceptos.items() if not str(k).startswith("__")}
    conceptos["__benchmark_kind__"] = "construdata_matrices"
    conceptos["__matrix_sheet__"] = matrix_sheet_name
    return conceptos


def extract_precios_nacional(filepath: str) -> dict:
    """
    Lee benchmark/catálogo nacional y además construye un índice auxiliar de materiales.
    Soporta encabezados flexibles como Codigo/Código, Descripcion/Descripción,
    PrecioUnitarioReferencia, PU mercado, referencia, etc.
    Cachea por ruta + mtime para no recargar el benchmark si no cambió.
    """
    src = Path(filepath)
    cache_key = None
    try:
        stat = src.stat()
        cache_key = (str(src.resolve()), stat.st_mtime_ns, stat.st_size)
        cached = _BENCHMARK_CACHE.get(cache_key)
        if cached is not None:
            return dict(cached)
    except Exception:
        cache_key = None

    try:
        wb_probe = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        if _is_construdata_matrix_workbook(wb_probe):
            wb_probe.close()
            matrices = load_construdata_matrices(filepath)
            if matrices:
                if cache_key is not None:
                    _BENCHMARK_CACHE.clear()
                    _BENCHMARK_CACHE[cache_key] = dict(matrices)
                return matrices
        wb_probe.close()
    except Exception:
        pass

    try:
        conceptos_v1 = extract_template_v1(filepath)
        if conceptos_v1:
            if cache_key is not None:
                _BENCHMARK_CACHE.clear()
                _BENCHMARK_CACHE[cache_key] = dict(conceptos_v1)
            return conceptos_v1
    except Exception:
        pass

    wb = openpyxl.load_workbook(filepath, data_only=True)
    precios = {}
    materials_index = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        header_row_idx = None
        header_map = None
        for i, row in enumerate(rows[:30]):
            headers = [_normalize_text(c) for c in row]
            col_codigo = _find_header_index(headers, ["clave", "codigo", "código"])
            col_desc = _find_header_index(headers, ["descripcion", "descripción", "concepto", "insumo"])
            col_pu = _find_header_index(headers, [
                "preciounitarioreferencia",
                "precio unitario referencia",
                "pu referencia",
                "precio referencia",
                "precio unitario",
                lambda h: _header_has_token(h, "precio", "unitario"),
                lambda h: _header_has_token(h, "precio", "referencia"),
                lambda h: h == "pu",
                lambda h: " mercado" in f" {h} ",
            ])
            if col_codigo is not None and col_desc is not None and col_pu is not None:
                header_row_idx = i
                header_map = {
                    "codigo": col_codigo,
                    "descripcion": col_desc,
                    "precio": col_pu,
                    "unidad": _find_header_index(headers, ["unidad", "u.m", "um"]),
                    "descripcion_larga": _find_header_index(headers, ["descripcionlarga", "descripcion larga", "detalle", "descripcion extendida"]),
                    "tipo": _find_header_index(headers, ["tipo", "tipoinsumo", "tipo insumo", "clasificacion", "clasificación"]),
                }
                break

        if header_row_idx is None or not header_map:
            continue

        for row in rows[header_row_idx + 1:]:
            cod = row[header_map["codigo"]] if header_map["codigo"] < len(row) else None
            desc = row[header_map["descripcion"]] if header_map["descripcion"] < len(row) else None
            pu = _as_float(row[header_map["precio"]] if header_map["precio"] < len(row) else None)
            if not cod or not desc or pu is None:
                continue

            unidad = row[header_map["unidad"]] if header_map.get("unidad") is not None and header_map["unidad"] < len(row) else ""
            desc_larga = row[header_map["descripcion_larga"]] if header_map.get("descripcion_larga") is not None and header_map["descripcion_larga"] < len(row) else ""
            tipo_raw = row[header_map["tipo"]] if header_map.get("tipo") is not None and header_map["tipo"] < len(row) else ""
            tipo_kind = _material_kind_from_tipo(tipo_raw)

            clave = str(cod).strip()
            desc_corta = str(desc).strip()
            desc_ref = str(desc_larga).strip() if desc_larga else desc_corta
            unidad_txt = str(unidad or "").strip()

            precios[clave] = {
                "clave": clave,
                "desc": desc_corta[:300],
                "desc_larga": desc_ref[:600],
                "unidad": unidad_txt,
                "pu_mercado": pu,
                "tipo": str(tipo_raw or "").strip(),
            }

            if tipo_kind == "mano_obra":
                continue

            materials_index.append({
                "clave": clave,
                "descripcion": desc_ref[:600],
                "descripcion_corta": desc_corta[:300],
                "unidad": unidad_txt,
                "unidad_norm": normalize_unit(unidad_txt),
                "precio": pu,
                "tipo_insumo": str(tipo_raw or "").strip(),
                "tipo_kind": tipo_kind,
                **_extract_material_attributes(desc_ref, unidad_txt),
            })

    dedup = []
    seen = set()
    for item in materials_index:
        key = (
            item.get("clave", ""),
            item.get("descripcion", ""),
            item.get("unidad_norm", ""),
            item.get("precio"),
        )
        if key in seen:
            continue
        seen.add(key)
        dedup.append(item)

    precios["__materials_index__"] = dedup
    if cache_key is not None:
        _BENCHMARK_CACHE.clear()
        _BENCHMARK_CACHE[cache_key] = dict(precios)
    return precios



def _sheet_headers(ws, headers):
    for col, head in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=head)
        c.font = _fnt(bold=True, color=C_BLANCO)
        c.fill = _fill(C_AZUL_OSC)
        c.alignment = _aln('center','center', wrap=True)
        c.border = _brd()

def _autosize_sheet(ws, max_width=38):
    for col in ws.columns:
        try:
            letter = col[0].column_letter
        except Exception:
            continue
        vals = [str(c.value) for c in col if c.value is not None]
        width = min(max((len(v) for v in vals), default=8) + 2, max_width)
        ws.column_dimensions[letter].width = width

def _build_market_materials_sheet(wb, datos):
    ws = wb.create_sheet("Mercado materiales")
    headers = [
        "Proveedor","Servicio","Material original","Material normalizado","Unidad original","Cantidad",
        "Precio proveedor","Importe proveedor","Fuente precio mercado","Descripcion mercado","Unidad mercado",
        "Precio mercado evidencia","Factor conversion","Precio mercado comparable","Delta","Confianza %",
        "Criterio decision","IA estado","Claude usado","Voyage usado","Internet usado","URL / referencia","Fecha consulta","Top candidatos"
    ]
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d["conceptos"].items():
            for mm in c.get("materiales_benchmark_matches") or []:
                inp = mm.get("input") or {}
                match = mm.get("match") or {}
                sel = match.get("selected") or {}
                resolution = match.get("market_resolution") or {}
                norm = resolution.get("normalized") or {}
                evidence = resolution.get("evidence") or sel or {}
                top = "; ".join(f"{x.get('descripcion','')[:45]} [{x.get('score','')}]" for x in match.get("top_candidates",[])[:3])
                market_price = resolution.get("effective_market_price") if resolution else sel.get("precio")
                delta = None
                if market_price and inp.get("precio_base") is not None:
                    delta = (float(inp["precio_base"]) - float(market_price)) / float(market_price) if market_price else None
                vals = [
                    d["nombre"], clave, inp.get("descripcion"), norm.get("normalized"), inp.get("unidad"), inp.get("factor"),
                    inp.get("precio_base"), inp.get("importe"), resolution.get("source") or sel.get("fuente_precio"), evidence.get("descripcion") or sel.get("descripcion"), resolution.get("market_unit") or sel.get("unidad"),
                    resolution.get("market_price_raw") if resolution else sel.get("precio_evidencia"), resolution.get("conversion_factor"), market_price, delta, resolution.get("confidence_pct") if resolution else (sel.get("score") * 100 if sel.get("score") is not None else None),
                    resolution.get("decision_criterion"), resolution.get("ai_status"), "SI" if resolution.get("claude_used") else "NO", "SI" if resolution.get("voyage_used") else "NO", "SI" if resolution.get("internet_used") else "NO", resolution.get("url"), resolution.get("fecha_consulta"), top
                ]
                for col,val in enumerate(vals, start=1):
                    cell = ws.cell(row=rown, column=col, value=val)
                    cell.border=_brd()
                    cell.font = _fnt(size=8)
                    if col in [7,8,12,14]:
                        cell.number_format='$#,##0.00'
                    if col==15 and isinstance(val,(int,float)):
                        cell.number_format='0.00%'
                    if col==16 and isinstance(val,(int,float)):
                        cell.number_format='0.0'
                    if col in [3,4,10,17,18,22,23,24]:
                        cell.alignment = _aln('left','center', wrap=True)
                rown += 1
    _autosize_sheet(ws, max_width=44)

def _build_market_labor_sheet(wb, datos):
    ws = wb.create_sheet("Mercado mano de obra")
    headers = ["Proveedor","Servicio","Jornada","Precio proveedor","Tabulador","Delta","Estado","Referencia"]
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d["conceptos"].items():
            for lm in c.get("mano_obra_benchmark_matches") or []:
                inp = lm.get("input") or {}
                mt = lm.get("match") or {}
                vals = [d["nombre"], clave, mt.get("tipo"), mt.get("precio_proveedor_comparable") if mt.get("precio_proveedor_comparable") is not None else mt.get("precio_proveedor"), mt.get("precio_tabulador_comparable") if mt.get("precio_tabulador_comparable") is not None else mt.get("precio_tabulador"), mt.get("delta"), mt.get("estado"), mt.get("referencia_concepto")]
                for col,val in enumerate(vals, start=1):
                    cell = ws.cell(row=rown, column=col, value=val)
                    cell.border = _brd()
                    if col in [4,5]:
                        cell.number_format='$#,##0.00'
                    if col == 6 and isinstance(val,(int,float)):
                        cell.number_format='0.00%'
                rown += 1
    _autosize_sheet(ws)


def _build_equipment_market_sheet(wb, datos):
    ws = wb.create_sheet("Mercado equipo especial")
    headers = ["Proveedor","Servicio","Equipo","Precio proveedor","Mercado ref.","Delta","Estado","Fuente","Evidencia"]
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d["conceptos"].items():
            for em in c.get("equipo_benchmark_matches") or []:
                inp = em.get("input") or {}
                sel = (em.get("match") or {}).get("selected") or {}
                if not sel:
                    continue
                precio_prov = inp.get("precio_base")
                precio_ref = sel.get("precio_ref")
                delta = ((precio_prov - precio_ref) / precio_ref) if isinstance(precio_prov,(int,float)) and isinstance(precio_ref,(int,float)) and precio_ref else None
                estado = "critico" if (delta is not None and delta > 0.5) else ("alto" if (delta is not None and delta > 0.2) else "medio")
                vals = [d["nombre"], clave, inp.get("descripcion"), precio_prov, precio_ref, delta, estado, sel.get("fuente"), sel.get("evidencia")]
                for col,val in enumerate(vals, start=1):
                    cell = ws.cell(row=rown, column=col, value=val)
                    cell.border = _brd()
                    if col in [4,5] and isinstance(val,(int,float)):
                        cell.number_format = '$#,##0.00'
                    if col == 6 and isinstance(val,(int,float)):
                        cell.number_format = '0.00%'
                rown += 1
    _autosize_sheet(ws)

def _build_drivers_sheet(wb, provider_analyses):
    ws = wb.create_sheet("Hallazgos PMD")
    headers = ["Proveedor","Tipo","Concepto","Servicio","Participación","Delta","Impacto","Estado","Proveedor $","Mercado $"]
    _sheet_headers(ws, headers)
    rown = 2
    for analysis in provider_analyses:
        for item in analysis.get("drivers", [])[:40]:
            vals = [analysis.get("provider_name"), item.get("tipo"), item.get("concepto"), item.get("clave_servicio"), item.get("participacion"), item.get("delta"), item.get("impacto"), item.get("estado"), item.get("proveedor"), item.get("mercado")]
            for col,val in enumerate(vals, start=1):
                cell = ws.cell(row=rown, column=col, value=val)
                cell.border = _brd()
                if col in [5,6,7] and isinstance(val,(int,float)):
                    cell.number_format='0.00%'
                if col in [9,10] and isinstance(val,(int,float)):
                    cell.number_format='$#,##0.00'
            rown += 1
    _autosize_sheet(ws)





# -----------------------------------------------------------------------------
# Paso 04 - Homologacion proveedor -> concepto Construdata y comparacion matriz
# -----------------------------------------------------------------------------

def _safe_ratio_delta(value, ref):
    if not isinstance(value, (int, float)) or not isinstance(ref, (int, float)) or ref == 0:
        return None
    return (float(value) - float(ref)) / float(ref)


def _concept_tokens(text):
    stop = {
        'de','del','la','las','los','el','en','con','sin','para','por','y','o','a','un','una',
        'incluye','incl','suministro','colocacion','colocacion','material','mano','obra',
        'equipo','herramienta','precio','unitario','m2','m3','ml','kg','pza','pieza'
    }
    return [t for t in _normalize_text(text).split() if len(t) >= 3 and t not in stop]


def _concept_similarity(provider_concept, construdata_concept):
    p_desc = provider_concept.get('desc') or provider_concept.get('descripcion') or ''
    c_desc = construdata_concept.get('desc_larga') or construdata_concept.get('desc') or ''
    p_tokens = _concept_tokens(p_desc)
    c_tokens = _concept_tokens(c_desc)
    score = _token_overlap_score(p_tokens, c_tokens)
    p_unit = normalize_unit(provider_concept.get('unidad'))
    c_unit = normalize_unit(construdata_concept.get('unidad'))
    if p_unit and c_unit and p_unit == c_unit:
        score += 0.12
    p_clave = _base_clave(provider_concept.get('clave') or provider_concept.get('clave_base'))
    c_clave = _base_clave(construdata_concept.get('clave') or construdata_concept.get('codigo'))
    if p_clave and c_clave and p_clave == c_clave:
        score = 1.0
    return round(min(1.0, score), 4)


def _get_construdata_concepts(precios_nacional):
    if not precios_nacional:
        return {}
    matrices = precios_nacional.get('__construdata_matrices__')
    if isinstance(matrices, dict) and matrices:
        return matrices
    # fallback defensivo: tomar registros que tengan matriz
    return {k: v for k, v in precios_nacional.items() if isinstance(v, dict) and v.get('matriz')}


def match_concepto_construdata(provider_key, provider_concept, construdata_concepts, min_score=0.32):
    """
    Busca el concepto Construdata asociado a un concepto de proveedor.
    Prioridad:
      1. match exacto por clave/base
      2. match por codigo embebido en descripcion
      3. match por similitud texto + unidad
    """
    if not construdata_concepts:
        return {'status': 'sin_base_construdata', 'selected': None, 'score': 0.0, 'reason': 'No se cargo una base Construdata con matrices'}

    key = _base_clave(provider_concept.get('clave_base') or provider_concept.get('clave') or provider_key)
    if key in construdata_concepts:
        return {'status': 'match_exacto', 'selected': construdata_concepts[key], 'score': 1.0, 'reason': 'Clave proveedor = CodigoConcepto Construdata'}

    key_norm = _normalize_text(key)
    for cd_key, cd in construdata_concepts.items():
        if _normalize_text(cd_key) == key_norm:
            return {'status': 'match_exacto_normalizado', 'selected': cd, 'score': 0.98, 'reason': 'Clave normalizada coincide'}

    desc = provider_concept.get('desc') or ''
    embedded = re.findall(r'\b\d{4,6}[- ]\d{2,4}\b', str(desc))
    for code in embedded:
        code = code.replace(' ', '-')
        if code in construdata_concepts:
            return {'status': 'match_codigo_en_descripcion', 'selected': construdata_concepts[code], 'score': 0.94, 'reason': 'Codigo Construdata encontrado en descripcion del proveedor'}

    best = None
    best_score = 0.0
    for _, cd in construdata_concepts.items():
        sc = _concept_similarity(provider_concept, cd)
        if sc > best_score:
            best_score = sc
            best = cd

    if best and best_score >= min_score:
        status = 'match_probable' if best_score >= 0.55 else 'match_baja_confianza'
        return {'status': status, 'selected': best, 'score': best_score, 'reason': 'Similitud descripcion/unidad'}

    return {'status': 'sin_match', 'selected': best, 'score': best_score, 'reason': 'No se encontro similitud suficiente'}


def _provider_matrix_items(concept):
    out = []
    mapping = [
        ('materiales', 'materiales_items'),
        ('mano_obra', 'mano_obra_items'),
        ('equipo', 'equipo_items'),
        ('basicos', 'basicos_items'),
    ]
    for kind, key in mapping:
        for item in concept.get(key) or []:
            r = dict(item)
            r['tipo_kind'] = kind
            r['cantidad'] = r.get('factor')
            r['precio'] = r.get('precio_base')
            out.append(r)
    return out


def _component_match_score(provider_item, cd_item):
    p_desc = provider_item.get('descripcion') or ''
    c_desc = cd_item.get('descripcion') or cd_item.get('desc') or ''
    p_tokens = _concept_tokens(p_desc)
    c_tokens = _concept_tokens(c_desc)
    score = 0.0
    score += 0.58 * _token_overlap_score(p_tokens, c_tokens)
    if normalize_unit(provider_item.get('unidad')) and normalize_unit(provider_item.get('unidad')) == normalize_unit(cd_item.get('unidad')):
        score += 0.16
    if provider_item.get('tipo_kind') and cd_item.get('tipo_kind') and provider_item.get('tipo_kind') == cd_item.get('tipo_kind'):
        score += 0.14
    pc = _normalize_text(provider_item.get('codigo'))
    cc = _normalize_text(cd_item.get('codigo') or cd_item.get('clave'))
    if pc and cc and pc == cc:
        score = max(score, 0.98)
    return round(min(score, 1.0), 4)


def _match_matrix_component(provider_item, cd_items, used_indexes=None, min_score=0.30):
    used_indexes = used_indexes or set()
    best_idx = None
    best = None
    best_score = 0.0
    for idx, cd_item in enumerate(cd_items or []):
        if idx in used_indexes:
            continue
        sc = _component_match_score(provider_item, cd_item)
        if sc > best_score:
            best_score = sc
            best_idx = idx
            best = cd_item
    if best and best_score >= min_score:
        return best_idx, best, best_score
    return None, None, best_score


def compare_provider_matrix_vs_construdata(provider_key, provider_concept, construdata_match):
    cd = (construdata_match or {}).get('selected')
    result = {
        'provider_key': provider_key,
        'match_status': (construdata_match or {}).get('status'),
        'match_score': (construdata_match or {}).get('score'),
        'match_reason': (construdata_match or {}).get('reason'),
        'construdata_codigo': None,
        'construdata_desc': None,
        'construdata_pu': None,
        'provider_pu': provider_concept.get('pu'),
        'delta_pu': None,
        'summary': {},
        'rows': [],
        'findings': [],
    }
    if not cd:
        result['findings'].append({'severity': 'alta', 'type': 'concepto_sin_match', 'message': 'No se encontro concepto equivalente en Construdata'})
        return result

    result['construdata_codigo'] = cd.get('codigo') or cd.get('clave')
    result['construdata_desc'] = cd.get('desc_larga') or cd.get('desc')
    result['construdata_pu'] = cd.get('pu_mercado')
    result['delta_pu'] = _safe_ratio_delta(provider_concept.get('pu'), cd.get('pu_mercado'))

    provider_items = _provider_matrix_items(provider_concept)
    cd_items = list(cd.get('matriz') or [])
    used_cd = set()
    matched = 0
    missing = 0
    extra = 0
    amount_delta = 0.0

    for p_item in provider_items:
        idx, cd_item, score = _match_matrix_component(p_item, cd_items, used_cd)
        row = {
            'tipo': p_item.get('tipo_kind'),
            'provider_codigo': p_item.get('codigo'),
            'provider_desc': p_item.get('descripcion'),
            'provider_unidad': p_item.get('unidad'),
            'provider_cantidad': p_item.get('factor') if p_item.get('factor') is not None else p_item.get('cantidad'),
            'provider_precio': p_item.get('precio_base') if p_item.get('precio_base') is not None else p_item.get('precio'),
            'provider_importe': p_item.get('importe'),
            'cd_codigo': None,
            'cd_desc': None,
            'cd_unidad': None,
            'cd_cantidad': None,
            'cd_precio': None,
            'cd_importe': None,
            'match_score': score,
            'delta_cantidad': None,
            'delta_precio': None,
            'delta_importe': None,
            'estado': 'sin_match_insumo',
        }
        if cd_item:
            used_cd.add(idx)
            matched += 1
            row.update({
                'cd_codigo': cd_item.get('codigo') or cd_item.get('clave'),
                'cd_desc': cd_item.get('descripcion'),
                'cd_unidad': cd_item.get('unidad'),
                'cd_cantidad': cd_item.get('cantidad'),
                'cd_precio': cd_item.get('precio'),
                'cd_importe': cd_item.get('importe'),
                'estado': 'match_insumo',
            })
            row['delta_cantidad'] = _safe_ratio_delta(row['provider_cantidad'], row['cd_cantidad'])
            row['delta_precio'] = _safe_ratio_delta(row['provider_precio'], row['cd_precio'])
            row['delta_importe'] = _safe_ratio_delta(row['provider_importe'], row['cd_importe'])
            if isinstance(row['provider_importe'], (int, float)) and isinstance(row['cd_importe'], (int, float)):
                amount_delta += float(row['provider_importe']) - float(row['cd_importe'])
        else:
            extra += 1
        result['rows'].append(row)

    for idx, cd_item in enumerate(cd_items):
        if idx in used_cd:
            continue
        missing += 1
        result['rows'].append({
            'tipo': cd_item.get('tipo_kind'),
            'provider_codigo': None,
            'provider_desc': None,
            'provider_unidad': None,
            'provider_cantidad': None,
            'provider_precio': None,
            'provider_importe': None,
            'cd_codigo': cd_item.get('codigo') or cd_item.get('clave'),
            'cd_desc': cd_item.get('descripcion'),
            'cd_unidad': cd_item.get('unidad'),
            'cd_cantidad': cd_item.get('cantidad'),
            'cd_precio': cd_item.get('precio'),
            'cd_importe': cd_item.get('importe'),
            'match_score': None,
            'delta_cantidad': None,
            'delta_precio': None,
            'delta_importe': None,
            'estado': 'faltante_en_proveedor',
        })

    result['summary'] = {
        'proveedor_renglones': len(provider_items),
        'construdata_renglones': len(cd_items),
        'insumos_matcheados': matched,
        'insumos_extra_proveedor': extra,
        'insumos_faltantes_proveedor': missing,
        'diferencia_importe_matriz': round(amount_delta, 4),
    }

    if result['delta_pu'] is not None and abs(result['delta_pu']) >= 0.15:
        result['findings'].append({'severity': 'alta' if abs(result['delta_pu']) >= 0.30 else 'media', 'type': 'delta_pu', 'message': f"PU proveedor vs Construdata: {result['delta_pu']*100:.1f}%"})
    if missing:
        result['findings'].append({'severity': 'alta', 'type': 'insumos_faltantes', 'message': f'Faltan {missing} insumos de la matriz Construdata en proveedor'})
    if extra:
        result['findings'].append({'severity': 'media', 'type': 'insumos_extra', 'message': f'Proveedor incluye {extra} insumos no homologados contra Construdata'})
    for row in result['rows']:
        if row.get('estado') != 'match_insumo':
            continue
        if isinstance(row.get('delta_cantidad'), (int, float)) and abs(row['delta_cantidad']) >= 0.25:
            result['findings'].append({'severity': 'alta', 'type': 'rendimiento', 'message': f"Rendimiento/cantidad diferente en {row.get('provider_desc')}: {row['delta_cantidad']*100:.1f}%"})
        if isinstance(row.get('delta_precio'), (int, float)) and abs(row['delta_precio']) >= 0.20:
            result['findings'].append({'severity': 'media', 'type': 'precio_insumo', 'message': f"Precio diferente en {row.get('provider_desc')}: {row['delta_precio']*100:.1f}%"})
    return result


def attach_construdata_matrix_analysis(conceptos, precios_nacional):
    construdata_concepts = _get_construdata_concepts(precios_nacional)
    if not construdata_concepts:
        return conceptos
    concept_index = _build_construdata_token_index(construdata_concepts)
    for key, concept in conceptos.items():
        match = match_concepto_construdata(key, concept, construdata_concepts, concept_index=concept_index)
        comparison = compare_provider_matrix_vs_construdata(key, concept, match)
        concept['construdata_concept_match'] = match
        concept['construdata_matrix_comparison'] = comparison
        concept['construdata_matrix_findings'] = comparison.get('findings') or []
    return conceptos


def _build_construdata_matching_sheet(wb, datos):
    ws = wb.create_sheet('Construdata matching')
    headers = ['Proveedor','Concepto proveedor','Descripcion proveedor','Unidad','PU proveedor','Estado match','Score','Codigo Construdata','Descripcion Construdata','PU Construdata','Delta PU','Renglones prov.','Renglones CD','Matcheados','Extra prov.','Faltantes prov.','Hallazgos']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            comp = c.get('construdata_matrix_comparison') or {}
            summary = comp.get('summary') or {}
            findings = '; '.join(f.get('message','') for f in (comp.get('findings') or [])[:4])
            vals = [
                d.get('nombre'), clave, c.get('desc'), c.get('unidad'), c.get('pu'), comp.get('match_status'), comp.get('match_score'), comp.get('construdata_codigo'), comp.get('construdata_desc'), comp.get('construdata_pu'), comp.get('delta_pu'),
                summary.get('proveedor_renglones'), summary.get('construdata_renglones'), summary.get('insumos_matcheados'), summary.get('insumos_extra_proveedor'), summary.get('insumos_faltantes_proveedor'), findings
            ]
            for col, val in enumerate(vals, start=1):
                cell = ws.cell(row=rown, column=col, value=val)
                cell.border = _brd()
                cell.font = _fnt(size=8)
                if col in [5,10] and isinstance(val, (int, float)):
                    cell.number_format = '$#,##0.00'
                if col == 11 and isinstance(val, (int, float)):
                    cell.number_format = '+0.0%;-0.0%;0.0%'
                if col in [3,9,17]:
                    cell.alignment = _aln('left','center', wrap=True)
            rown += 1
    _autosize_sheet(ws, max_width=48)


def _build_construdata_matrix_detail_sheet(wb, datos):
    ws = wb.create_sheet('Construdata matrices')
    headers = ['Proveedor','Concepto proveedor','Codigo CD','Estado','Tipo','Insumo proveedor','Unidad prov.','Cant. prov.','Precio prov.','Importe prov.','Insumo CD','Unidad CD','Cant. CD','Precio CD','Importe CD','Score','Delta cant.','Delta precio','Delta importe']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            comp = c.get('construdata_matrix_comparison') or {}
            for r in comp.get('rows') or []:
                vals = [
                    d.get('nombre'), clave, comp.get('construdata_codigo'), r.get('estado'), r.get('tipo'), r.get('provider_desc'), r.get('provider_unidad'), r.get('provider_cantidad'), r.get('provider_precio'), r.get('provider_importe'),
                    r.get('cd_desc'), r.get('cd_unidad'), r.get('cd_cantidad'), r.get('cd_precio'), r.get('cd_importe'), r.get('match_score'), r.get('delta_cantidad'), r.get('delta_precio'), r.get('delta_importe')
                ]
                for col, val in enumerate(vals, start=1):
                    cell = ws.cell(row=rown, column=col, value=val)
                    cell.border = _brd()
                    cell.font = _fnt(size=8)
                    if col in [9,10,14,15] and isinstance(val, (int, float)):
                        cell.number_format = '$#,##0.00'
                    if col in [17,18,19] and isinstance(val, (int, float)):
                        cell.number_format = '+0.0%;-0.0%;0.0%'
                    if col in [6,11]:
                        cell.alignment = _aln('left','center', wrap=True)
                rown += 1
    _autosize_sheet(ws, max_width=48)



# -----------------------------------------------------------------------------
# Paso 05 - Homologacion semantica fuerte + conceptos compuestos
# -----------------------------------------------------------------------------

RAW_CONSTRUDATA_POSITIONAL_SCHEMA = {
    "codigo_concepto": 4,
    "concepto_desc": 5,
    "concepto_desc_larga": 6,
    "concepto_unidad": 7,
    "concepto_pu": 10,
    "renglon": 3,
    "codigo_insumo": 13,
    "insumo_desc": 14,
    "insumo_desc_larga": 15,
    "insumo_unidad": 16,
    "tipo_id": 17,
    "tipo": 18,
    "precio": 19,
    "cantidad": 20,
    "dividir": 21,
    "expresion": 22,
    "importe": 23,
}

CONCEPT_ACTIVITY_TOKENS = {
    "demolicion", "demoler", "retiro", "retirar", "desmontaje", "desmontar", "limpieza", "limpia", "tala",
    "trazo", "nivelacion", "excavacion", "relleno", "compactacion", "acarreo", "carga", "colado",
    "concreto", "cimbra", "acero", "muro", "losa", "firme", "pintura", "impermeabilizante", "barandal",
    "moldura", "suministro", "instalacion", "colocacion", "acabado", "aplicacion", "sellado", "reparacion",
}

MATCH_ONLY_NOISE_PHRASES = [
    r"\bincluye\b.*$",
    r"\bincl\.\b.*$",
    r"\bconsidera\b.*$",
    r"\bver especificaciones\b.*$",
    r"\bsegun planos\b.*$",
]

MATCH_ONLY_STOPWORDS = {
    'incluye','incl','materiales','material','mano','obra','herramienta','herramientas','equipo','seguridad',
    'acarreos','acarreo','limpieza','limpio','retirando','retirado','necesario','necesarios','todo','todos',
    'precio','unitario','proyecto','obra','trabajo','trabajos','servicio','servicios','suministro','colocacion',
    'con','sin','para','por','de','del','la','las','los','el','en','y','o','a','un','una','al','se','su','sus',
    'hasta','desde','sobre','bajo','segun','planos','especificaciones','generales','particulares'
}



def _matrix_stats(items):
    stats = {'total_renglones': len(items or []), 'materiales': 0, 'mano_obra': 0, 'equipo': 0, 'basicos': 0, 'importe_total': 0.0}
    for item in items or []:
        kind = item.get('tipo_kind')
        if kind in stats:
            stats[kind] += 1
        imp = item.get('importe')
        if isinstance(imp, (int, float)):
            stats['importe_total'] += float(imp)
    stats['importe_total'] = round(stats['importe_total'], 6)
    return stats

def _is_raw_construdata_query_sheet(ws):
    try:
        row = next(ws.iter_rows(values_only=True))
    except StopIteration:
        return False
    if not row or len(row) < 24:
        return False
    return bool(re.match(r"^\d{4,6}-\d{2,4}$", _clean_text(row[4] if len(row) > 4 else "")))


def _load_construdata_raw_query_workbook(filepath: str) -> dict:
    """Carga el Excel real de la consulta SQL recibida: sheet1 sin encabezados.

    La consulta trae ya Concepto -> Renglon de matriz -> Insumo. Este loader preserva
    el importe calculado por Neodata y clasifica formulas especiales:
    - porcentaje_mo: %MO1/%MO5 usan subtotal MO como base, cantidad como porcentaje.
    - rendimiento_inverso: Dividir=True.
    - normal: precio*cantidad.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    if not _is_raw_construdata_query_sheet(ws):
        return {}
    idx = RAW_CONSTRUDATA_POSITIONAL_SCHEMA
    conceptos = {}
    materials_index = []
    seen_materials = set()
    # primera pasada: leer renglones
    for row in ws.iter_rows(values_only=True):
        codigo = _clean_text(_cell(row, idx['codigo_concepto']))
        if not codigo:
            continue
        desc = _clean_text(_cell(row, idx['concepto_desc']))
        desc_larga = _clean_text(_cell(row, idx['concepto_desc_larga'])) or desc
        unidad = _clean_text(_cell(row, idx['concepto_unidad']))
        pu = _as_float(_cell(row, idx['concepto_pu']))
        if codigo not in conceptos:
            conceptos[codigo] = {
                'clave': codigo, 'codigo': codigo, 'desc': desc[:300], 'desc_larga': desc_larga[:900],
                'unidad': unidad, 'pu_mercado': pu, 'tipo': 'CONCEPTO', 'matriz': [],
                'partida_codigo': codigo.split('-')[0] if '-' in codigo else codigo[:5],
            }
        ins_code = _clean_text(_cell(row, idx['codigo_insumo']))
        ins_desc = _clean_text(_cell(row, idx['insumo_desc_larga'])) or _clean_text(_cell(row, idx['insumo_desc']))
        if not ins_code and not ins_desc:
            continue
        tipo_raw = _clean_text(_cell(row, idx['tipo']))
        cantidad = _as_float(_cell(row, idx['cantidad']))
        precio = _as_float(_cell(row, idx['precio']))
        importe = _as_float(_cell(row, idx['importe']))
        dividir = bool(_cell(row, idx['dividir']))
        formula_kind = 'normal'
        if ins_code.upper().startswith('%MO') or ('%MO' in ins_code.upper()):
            formula_kind = 'porcentaje_mo'
        elif dividir:
            formula_kind = 'rendimiento_inverso'
        renglon = {
            'renglon': _cell(row, idx['renglon']),
            'tipo': tipo_raw,
            'tipo_id': _cell(row, idx['tipo_id']),
            'tipo_kind': _material_kind_from_tipo(tipo_raw),
            'codigo': ins_code,
            'clave': ins_code,
            'descripcion': ins_desc[:700],
            'unidad': _clean_text(_cell(row, idx['insumo_unidad'])),
            'unidad_norm': normalize_unit(_cell(row, idx['insumo_unidad'])),
            'cantidad': cantidad,
            'precio': precio,
            'precio_base': precio,
            'importe': importe,
            'dividir': dividir,
            'expresion': _clean_text(_cell(row, idx['expresion'])),
            'formula_kind': formula_kind,
            'porcentaje': cantidad if formula_kind == 'porcentaje_mo' else None,
            'base_calculo': precio if formula_kind == 'porcentaje_mo' else None,
            'fuente': 'construdata_raw_query',
        }
        conceptos[codigo]['matriz'].append(renglon)
        # No se construye un indice plano masivo de insumos para este formato.
        # La comparacion principal debe ser concepto -> matriz Construdata.
        # Mantenerlo vacio evita costos altos de tokenizacion y de matching heredado.
    # segunda pasada: corrige porcentajes si el importe viniera mal o vacio
    for concepto in conceptos.values():
        subtotal_mo = 0.0
        for r in concepto.get('matriz') or []:
            if r.get('tipo_kind') == 'mano_obra' and r.get('formula_kind') != 'porcentaje_mo' and isinstance(r.get('importe'), (int, float)):
                subtotal_mo += float(r['importe'])
        for r in concepto.get('matriz') or []:
            if r.get('formula_kind') == 'porcentaje_mo':
                pct = r.get('porcentaje')
                if isinstance(pct, (int, float)):
                    r['base_calculo'] = subtotal_mo or r.get('base_calculo')
                    correcto = (subtotal_mo or 0.0) * float(pct)
                    # Para %MO el importe real es subtotal_mo * porcentaje.
                    # La consulta puede traer un importe calculado contra otra base temporal; se normaliza aquí.
                    r['importe'] = round(correcto, 6)
                    r['precio'] = r['base_calculo']
                    r['precio_base'] = r['base_calculo']
        concepto['matriz_stats'] = _matrix_stats(concepto.get('matriz') or [])
        concepto['_match_tokens'] = _concept_match_tokens(concepto.get('desc_larga') or concepto.get('desc') or '')
        concepto['_matrix_tokens'] = _matrix_summary_tokens(concepto.get('matriz') or [], limit=30)
        concepto['_core_text'] = _concept_core_text(concepto.get('desc_larga') or concepto.get('desc') or '')
        concepto['embedding_text'] = _concept_embedding_text(concepto)
    conceptos['__materials_index__'] = materials_index
    conceptos['__construdata_matrices__'] = {k:v for k,v in conceptos.items() if not k.startswith('__')}
    conceptos['__benchmark_kind__'] = 'construdata_matrices_raw_query'
    conceptos['__raw_query_schema__'] = RAW_CONSTRUDATA_POSITIONAL_SCHEMA
    try:
        wb.close()
    except Exception:
        pass
    return conceptos

# override detector + loader para soportar sheet1 sin encabezados ademas de los formatos anteriores
_is_construdata_matrix_workbook_previous = _is_construdata_matrix_workbook

def _is_construdata_matrix_workbook(wb):
    try:
        if _is_construdata_matrix_workbook_previous(wb):
            return True
    except Exception:
        pass
    try:
        if wb.sheetnames and _is_raw_construdata_query_sheet(wb[wb.sheetnames[0]]):
            return True
    except Exception:
        pass
    return False

_load_construdata_matrices_previous = load_construdata_matrices

def load_construdata_matrices(filepath: str) -> dict:
    try:
        raw = _load_construdata_raw_query_workbook(filepath)
        if raw:
            return raw
    except Exception as exc:
        print(f"Advertencia: no se pudo leer Construdata raw query: {exc}")
    return _load_construdata_matrices_previous(filepath)


def _concept_core_text(desc: str) -> str:
    """Texto solo para homologacion del concepto principal.

    No elimina nada de la matriz. Solo evita que frases tipo 'incluye mano de obra,
    herramienta, seguridad...' dominen el match del nombre del concepto.
    """
    s = _normalize_text(desc)
    for pattern in MATCH_ONLY_NOISE_PHRASES:
        s = re.sub(pattern, ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(ver|segun)\b.*$', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def _concept_match_tokens(desc: str):
    core = _concept_core_text(desc)
    return [t for t in core.split() if len(t) >= 3 and t not in MATCH_ONLY_STOPWORDS]


def _matrix_summary_tokens(items, limit=40):
    tokens = []
    for item in items or []:
        desc = item.get('descripcion') or item.get('desc') or ''
        for t in _tokenize_text(desc):
            if t not in MATCH_ONLY_STOPWORDS:
                tokens.append(t)
    # frecuencia simple, preservando orden de relevancia aproximado
    counts = defaultdict(int)
    for t in tokens:
        counts[t] += 1
    ranked = sorted(counts, key=lambda x: counts[x], reverse=True)
    return ranked[:limit]


def _concept_embedding_text(concept: dict) -> str:
    matriz_tokens = _matrix_summary_tokens(concept.get('matriz') or [], limit=24)
    return ' | '.join([
        f"codigo:{concept.get('codigo') or concept.get('clave')}",
        f"partida:{concept.get('partida_codigo') or ''}",
        f"unidad:{concept.get('unidad') or ''}",
        f"concepto:{concept.get('desc_larga') or concept.get('desc') or ''}",
        f"matriz:{' '.join(matriz_tokens)}",
    ])[:1800]


def _provider_embedding_text(provider_concept: dict) -> str:
    matriz = _provider_matrix_items(provider_concept)
    matriz_tokens = _matrix_summary_tokens(matriz, limit=24)
    return ' | '.join([
        f"unidad:{provider_concept.get('unidad') or ''}",
        f"concepto:{provider_concept.get('desc') or provider_concept.get('descripcion') or ''}",
        f"matriz:{' '.join(matriz_tokens)}",
    ])[:1800]


def _sequence_score(a: str, b: str) -> float:
    a = _concept_core_text(a)
    b = _concept_core_text(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def _price_similarity_bonus(provider_pu, cd_pu):
    if not isinstance(provider_pu, (int, float)) or not isinstance(cd_pu, (int, float)) or cd_pu <= 0:
        return 0.0
    ratio = abs(float(provider_pu) - float(cd_pu)) / max(abs(float(cd_pu)), 1.0)
    if ratio <= 0.25:
        return 0.06
    if ratio <= 0.60:
        return 0.03
    return 0.0


def _concept_candidate_score(provider_concept, cd):
    p_desc = provider_concept.get('desc') or provider_concept.get('descripcion') or ''
    c_desc = cd.get('desc_larga') or cd.get('desc') or ''
    p_tokens = provider_concept.get('_match_tokens') or _concept_match_tokens(p_desc)
    c_tokens = cd.get('_match_tokens') or _concept_match_tokens(c_desc)
    overlap = _token_overlap_score(p_tokens, c_tokens)
    # SequenceMatcher solo sobre core precomputado para evitar costo excesivo.
    p_core = provider_concept.get('_core_text') or _concept_core_text(p_desc)
    c_core = cd.get('_core_text') or _concept_core_text(c_desc)
    seq = SequenceMatcher(None, p_core[:160], c_core[:160]).ratio() if p_core and c_core else 0.0
    p_matrix_tokens = provider_concept.get('_matrix_tokens') or _matrix_summary_tokens(_provider_matrix_items(provider_concept), limit=30)
    c_matrix_tokens = cd.get('_matrix_tokens') or _matrix_summary_tokens(cd.get('matriz') or [], limit=30)
    matrix_overlap = _token_overlap_score(p_matrix_tokens, c_matrix_tokens)
    score = 0.52 * overlap + 0.18 * seq + 0.14 * matrix_overlap
    p_unit = normalize_unit(provider_concept.get('unidad'))
    c_unit = normalize_unit(cd.get('unidad'))
    if p_unit and c_unit and p_unit == c_unit:
        score += 0.12
    elif p_unit and c_unit and p_unit != c_unit:
        score -= 0.08
    score += _price_similarity_bonus(provider_concept.get('pu'), cd.get('pu_mercado'))
    # actividad fuerte compartida
    shared_activity = set(p_tokens) & set(c_tokens) & CONCEPT_ACTIVITY_TOKENS
    if shared_activity:
        score += 0.04
    return max(0.0, min(1.0, round(score, 4))), sorted(list(set(p_tokens) & set(c_tokens)))[:12]


def _build_construdata_token_index(construdata_concepts):
    token_index = defaultdict(set)
    unit_index = defaultdict(set)
    for code, cd in (construdata_concepts or {}).items():
        if str(code).startswith('__'):
            continue
        tokens = cd.get('_match_tokens') or _concept_match_tokens(cd.get('desc_larga') or cd.get('desc') or '')
        cd['_match_tokens'] = tokens
        for tok in set(tokens):
            token_index[tok].add(code)
        u = normalize_unit(cd.get('unidad'))
        if u:
            unit_index[u].add(code)
    return {'token_index': token_index, 'unit_index': unit_index}


def _candidate_codes_for_provider(provider_concept, construdata_concepts, concept_index):
    if not concept_index:
        return list(construdata_concepts.keys())
    p_tokens = set(provider_concept.get('_match_tokens') or _concept_match_tokens(provider_concept.get('desc') or ''))
    counts = defaultdict(int)
    for tok in p_tokens:
        for code in concept_index.get('token_index', {}).get(tok, ()): counts[code] += 1
    p_unit = normalize_unit(provider_concept.get('unidad'))
    unit_codes = concept_index.get('unit_index', {}).get(p_unit, set()) if p_unit else set()
    ranked_codes = []
    for code, cnt in counts.items():
        bonus = 1 if code in unit_codes else 0
        ranked_codes.append((cnt + bonus, code))
    ranked_codes.sort(reverse=True)
    codes = [code for _, code in ranked_codes[:120]]
    # Si muy pocos, agrega por unidad para no perder candidatos con sinonimos.
    if len(codes) < 80 and unit_codes:
        for code in list(unit_codes)[:100]:
            if code not in codes:
                codes.append(code)
    return codes or list(construdata_concepts.keys())[:250]


def _rank_construdata_candidates(provider_concept, construdata_concepts, top_k=25, concept_index=None):
    ranked = []
    p_unit = normalize_unit(provider_concept.get('unidad'))
    p_desc = provider_concept.get('desc') or ''
    provider_concept['_match_tokens'] = provider_concept.get('_match_tokens') or _concept_match_tokens(p_desc)
    provider_concept['_matrix_tokens'] = provider_concept.get('_matrix_tokens') or _matrix_summary_tokens(_provider_matrix_items(provider_concept), limit=30)
    provider_concept['_core_text'] = provider_concept.get('_core_text') or _concept_core_text(p_desc)
    p_tokens = set(provider_concept['_match_tokens'])
    candidate_codes = _candidate_codes_for_provider(provider_concept, construdata_concepts, concept_index)
    for code in candidate_codes:
        cd = construdata_concepts.get(code)
        if not cd or str(code).startswith('__'):
            continue
        # filtro ligero por unidad o token para evitar matches aleatorios
        c_unit = normalize_unit(cd.get('unidad'))
        c_tokens = set(cd.get('_match_tokens') or _concept_match_tokens(cd.get('desc_larga') or cd.get('desc') or ''))
        if p_unit and c_unit and p_unit != c_unit and not (p_tokens & c_tokens):
            continue
        score, shared = _concept_candidate_score(provider_concept, cd)
        if score <= 0:
            continue
        ranked.append({
            'codigo': cd.get('codigo') or cd.get('clave') or code,
            'descripcion': cd.get('desc_larga') or cd.get('desc'),
            'unidad': cd.get('unidad'),
            'pu': cd.get('pu_mercado'),
            'score': score,
            'tokens_match': ', '.join(shared),
            'matriz_resumen': ', '.join(_matrix_summary_tokens(cd.get('matriz') or [], limit=12)),
            'embedding_text': cd.get('embedding_text') or _concept_embedding_text(cd),
            '_concept': cd,
        })
    ranked.sort(key=lambda x: x['score'], reverse=True)
    # reranking Voyage si esta disponible; si no, retorna igual
    query = _provider_embedding_text(provider_concept)
    reranked = voyage_rerank_concept_candidates(query, ranked[:40], top_k=top_k)
    # recuperar objeto interno tras reranking
    by_code = {c['codigo']: c for c in ranked}
    out = []
    for c in reranked:
        base = by_code.get(c.get('codigo'), c)
        merged = dict(base)
        merged.update({k:v for k,v in c.items() if k not in {'_concept'}})
        merged['_concept'] = base.get('_concept')
        out.append(merged)
    return out[:top_k]


def _detect_composite_match(provider_concept, candidates):
    """Detecta cuando 1 concepto proveedor engloba N conceptos Construdata.

    No fuerza un compuesto si el top1 ya es muy fuerte. Busca cobertura adicional de tokens
    tecnicos por candidatos diferentes.
    """
    if not candidates:
        return None
    p_tokens = set(_concept_match_tokens(provider_concept.get('desc') or ''))
    if len(p_tokens) < 4:
        return None
    top = candidates[0]
    if top.get('score', 0) >= 0.72:
        return None
    selected = []
    covered = set()
    for c in candidates[:10]:
        c_tokens = set(_concept_match_tokens(c.get('descripcion') or ''))
        useful = (p_tokens & c_tokens) - covered
        activity_useful = useful & CONCEPT_ACTIVITY_TOKENS
        if len(useful) >= 2 or activity_useful:
            selected.append((c, useful))
            covered |= useful
        if len(selected) >= 3 or (len(covered) / max(len(p_tokens),1)) >= 0.58:
            break
    if len(selected) >= 2:
        matches = []
        total = sum(max(c.get('score',0),0.01) for c,_ in selected)
        for c,useful in selected:
            matches.append({'codigo': c['codigo'], 'peso': round(max(c.get('score',0),0.01)/total, 4), 'motivo': 'cubre tokens: ' + ', '.join(sorted(useful)[:8])})
        return {
            'tipo_match': 'match_compuesto',
            'confianza': round(min(0.86, 0.45 + (len(covered)/max(len(p_tokens),1))*0.45), 4),
            'matches': matches,
            'requiere_revision': True,
            'explicacion': 'El concepto del proveedor parece agrupar varias actividades que Construdata maneja como conceptos separados.'
        }
    return None


def match_concepto_construdata(provider_key, provider_concept, construdata_concepts, min_score=0.38, concept_index=None):
    """Homologacion fuerte: codigo -> candidatos deterministicos/Voyage -> Claude opcional -> compuesto."""
    if not construdata_concepts:
        return {'status': 'sin_base_construdata', 'selected': None, 'selected_matches': [], 'score': 0.0, 'reason': 'No se cargo una base Construdata con matrices', 'ai_status': get_concept_ai_status()}
    key = _base_clave(provider_concept.get('clave_base') or provider_concept.get('clave') or provider_key)
    if key in construdata_concepts:
        cd = construdata_concepts[key]
        return {'status': 'match_exacto', 'tipo_match': 'match_directo', 'selected': cd, 'selected_matches': [{'concept': cd, 'peso': 1.0, 'motivo': 'clave exacta'}], 'score': 1.0, 'reason': 'Clave proveedor = CodigoConcepto Construdata', 'candidates': [], 'ai_status': get_concept_ai_status()}
    desc = provider_concept.get('desc') or ''
    embedded = re.findall(r'\b\d{4,6}[- ]\d{2,4}\b', str(desc))
    for code in embedded:
        code = code.replace(' ', '-')
        if code in construdata_concepts:
            cd = construdata_concepts[code]
            return {'status': 'match_codigo_en_descripcion', 'tipo_match': 'match_directo', 'selected': cd, 'selected_matches': [{'concept': cd, 'peso': 1.0, 'motivo': 'codigo embebido'}], 'score': 0.94, 'reason': 'Codigo Construdata encontrado en descripcion del proveedor', 'candidates': [], 'ai_status': get_concept_ai_status()}
    candidates = _rank_construdata_candidates(provider_concept, construdata_concepts, top_k=20, concept_index=concept_index)
    provider_payload = {
        'clave': provider_key,
        'descripcion_completa': provider_concept.get('desc'),
        'descripcion_core_matching': _concept_core_text(provider_concept.get('desc') or ''),
        'unidad': provider_concept.get('unidad'),
        'pu': provider_concept.get('pu'),
        'matriz_resumen': ', '.join(_matrix_summary_tokens(_provider_matrix_items(provider_concept), limit=16)),
    }
    claude_decision = claude_decide_concept_match(provider_payload, candidates)
    if isinstance(claude_decision, dict) and not claude_decision.get('error') and claude_decision.get('matches'):
        selected_matches = []
        for m in claude_decision.get('matches') or []:
            cd = construdata_concepts.get(m.get('codigo'))
            if cd:
                selected_matches.append({'concept': cd, 'peso': m.get('peso', 1.0), 'motivo': m.get('motivo')})
        if selected_matches:
            selected = selected_matches[0]['concept']
            tipo = claude_decision.get('tipo_match') or ('match_compuesto' if len(selected_matches)>1 else 'match_directo')
            return {
                'status': 'match_ia_' + tipo,
                'tipo_match': tipo,
                'selected': selected,
                'selected_matches': selected_matches,
                'score': claude_decision.get('confianza', 0.0),
                'reason': claude_decision.get('explicacion') or 'Validado por Claude',
                'requires_review': bool(claude_decision.get('requiere_revision')) or tipo != 'match_directo',
                'candidates': candidates[:10],
                'claude_decision': claude_decision,
                'ai_status': get_concept_ai_status(),
            }
    composite = _detect_composite_match(provider_concept, candidates)
    if composite:
        selected_matches = []
        for m in composite.get('matches') or []:
            cd = construdata_concepts.get(m.get('codigo'))
            if cd:
                selected_matches.append({'concept': cd, 'peso': m.get('peso', 1.0), 'motivo': m.get('motivo')})
        if selected_matches:
            return {'status': 'match_compuesto_reglas', 'tipo_match': 'match_compuesto', 'selected': selected_matches[0]['concept'], 'selected_matches': selected_matches, 'score': composite['confianza'], 'reason': composite['explicacion'], 'requires_review': True, 'candidates': candidates[:10], 'ai_status': get_concept_ai_status()}
    best = candidates[0] if candidates else None
    if best and best.get('score', 0) >= min_score:
        cd = best['_concept']
        status = 'match_fuerte' if best['score'] >= 0.62 else 'match_debil_revision'
        return {'status': status, 'tipo_match': 'match_directo' if best['score'] >= 0.62 else 'match_parcial', 'selected': cd, 'selected_matches': [{'concept': cd, 'peso': 1.0, 'motivo': 'mejor score reglas'}], 'score': best['score'], 'reason': 'Score descripcion/unidad/matriz', 'requires_review': best['score'] < 0.62, 'candidates': candidates[:10], 'ai_status': get_concept_ai_status()}
    return {'status': 'sin_match', 'tipo_match': 'sin_match', 'selected': best.get('_concept') if best else None, 'selected_matches': [], 'score': best.get('score',0) if best else 0.0, 'reason': 'No se encontro similitud suficiente', 'requires_review': True, 'candidates': candidates[:10], 'ai_status': get_concept_ai_status()}


def _selected_cd_items_from_match(construdata_match):
    matches = (construdata_match or {}).get('selected_matches') or []
    if not matches and (construdata_match or {}).get('selected'):
        matches = [{'concept': construdata_match.get('selected'), 'peso': 1.0, 'motivo': ''}]
    cd_items = []
    codes = []
    descs = []
    pus = []
    for m in matches:
        cd = m.get('concept') or {}
        code = cd.get('codigo') or cd.get('clave')
        if code:
            codes.append(code)
        if cd.get('desc'):
            descs.append(cd.get('desc'))
        if isinstance(cd.get('pu_mercado'), (int, float)):
            pus.append(float(cd.get('pu_mercado')))
        for item in cd.get('matriz') or []:
            r = dict(item)
            r['source_concept_code'] = code
            r['source_concept_desc'] = cd.get('desc')
            cd_items.append(r)
    return codes, descs, pus, cd_items


def _component_is_percentage(item):
    code = str(item.get('codigo') or item.get('clave') or '').upper()
    return item.get('formula_kind') == 'porcentaje_mo' or code.startswith('%MO')


def compare_provider_matrix_vs_construdata(provider_key, provider_concept, construdata_match):
    result = {
        'provider_key': provider_key,
        'match_status': (construdata_match or {}).get('status'),
        'match_type': (construdata_match or {}).get('tipo_match'),
        'match_score': (construdata_match or {}).get('score'),
        'match_reason': (construdata_match or {}).get('reason'),
        'requires_review': (construdata_match or {}).get('requires_review'),
        'construdata_codigo': None,
        'construdata_desc': None,
        'construdata_pu': None,
        'provider_pu': provider_concept.get('pu'),
        'delta_pu': None,
        'summary': {},
        'rows': [],
        'findings': [],
        'candidates': (construdata_match or {}).get('candidates') or [],
    }
    codes, descs, pus, cd_items = _selected_cd_items_from_match(construdata_match)
    if not cd_items and not codes:
        result['findings'].append({'severity': 'alta', 'type': 'concepto_sin_match', 'message': 'No se encontro concepto equivalente en Construdata'})
        return result
    result['construdata_codigo'] = ' + '.join(codes)
    result['construdata_desc'] = ' + '.join(descs[:4])
    result['construdata_pu'] = round(sum(pus), 6) if pus else None
    result['delta_pu'] = _safe_ratio_delta(provider_concept.get('pu'), result['construdata_pu'])
    provider_items = _provider_matrix_items(provider_concept)
    used_cd = set()
    matched = missing = extra = 0
    amount_delta = 0.0
    for p_item in provider_items:
        idx, cd_item, score = _match_matrix_component(p_item, cd_items, used_cd, min_score=0.26)
        p_cant = p_item.get('factor') if p_item.get('factor') is not None else p_item.get('cantidad')
        p_precio = p_item.get('precio_base') if p_item.get('precio_base') is not None else p_item.get('precio')
        row = {
            'tipo': p_item.get('tipo_kind'),
            'provider_codigo': p_item.get('codigo'),
            'provider_desc': p_item.get('descripcion'),
            'provider_unidad': p_item.get('unidad'),
            'provider_cantidad': p_cant,
            'provider_precio': p_precio,
            'provider_importe': p_item.get('importe'),
            'cd_codigo': None, 'cd_desc': None, 'cd_unidad': None, 'cd_cantidad': None, 'cd_precio': None, 'cd_importe': None,
            'cd_source_concept': None,
            'formula_kind': None,
            'match_score': score,
            'delta_cantidad': None, 'delta_precio': None, 'delta_importe': None,
            'estado': 'sin_match_insumo',
        }
        if cd_item:
            used_cd.add(idx); matched += 1
            row.update({
                'cd_codigo': cd_item.get('codigo') or cd_item.get('clave'),
                'cd_desc': cd_item.get('descripcion'),
                'cd_unidad': cd_item.get('unidad'),
                'cd_cantidad': cd_item.get('cantidad'),
                'cd_precio': cd_item.get('precio'),
                'cd_importe': cd_item.get('importe'),
                'cd_source_concept': cd_item.get('source_concept_code'),
                'formula_kind': cd_item.get('formula_kind'),
                'estado': 'match_insumo',
            })
            row['delta_cantidad'] = _safe_ratio_delta(row['provider_cantidad'], row['cd_cantidad'])
            # porcentajes: precio/base no es comparable; se compara porcentaje e importe.
            if _component_is_percentage(cd_item):
                row['delta_precio'] = None
            else:
                row['delta_precio'] = _safe_ratio_delta(row['provider_precio'], row['cd_precio'])
            row['delta_importe'] = _safe_ratio_delta(row['provider_importe'], row['cd_importe'])
            if isinstance(row['provider_importe'], (int, float)) and isinstance(row['cd_importe'], (int, float)):
                amount_delta += float(row['provider_importe']) - float(row['cd_importe'])
        else:
            extra += 1
        result['rows'].append(row)
    for idx, cd_item in enumerate(cd_items):
        if idx in used_cd:
            continue
        missing += 1
        result['rows'].append({
            'tipo': cd_item.get('tipo_kind'), 'provider_codigo': None, 'provider_desc': None, 'provider_unidad': None,
            'provider_cantidad': None, 'provider_precio': None, 'provider_importe': None,
            'cd_codigo': cd_item.get('codigo') or cd_item.get('clave'), 'cd_desc': cd_item.get('descripcion'), 'cd_unidad': cd_item.get('unidad'),
            'cd_cantidad': cd_item.get('cantidad'), 'cd_precio': cd_item.get('precio'), 'cd_importe': cd_item.get('importe'),
            'cd_source_concept': cd_item.get('source_concept_code'), 'formula_kind': cd_item.get('formula_kind'),
            'match_score': None, 'delta_cantidad': None, 'delta_precio': None, 'delta_importe': None, 'estado': 'faltante_en_proveedor',
        })
    result['summary'] = {
        'proveedor_renglones': len(provider_items), 'construdata_renglones': len(cd_items), 'insumos_matcheados': matched,
        'insumos_extra_proveedor': extra, 'insumos_faltantes_proveedor': missing, 'diferencia_importe_matriz': round(amount_delta, 4),
        'conceptos_cd_asociados': len(codes), 'requiere_revision': bool(result.get('requires_review')),
    }
    if result['match_type'] == 'match_compuesto':
        result['findings'].append({'severity': 'media', 'type': 'concepto_compuesto', 'message': f"Concepto proveedor homologado como compuesto contra {len(codes)} conceptos Construdata; requiere validacion tecnica."})
    if result['delta_pu'] is not None and abs(result['delta_pu']) >= 0.15:
        result['findings'].append({'severity': 'alta' if abs(result['delta_pu']) >= 0.30 else 'media', 'type': 'delta_pu', 'message': f"PU proveedor vs Construdata: {result['delta_pu']*100:.1f}%"})
    if missing:
        result['findings'].append({'severity': 'alta', 'type': 'insumos_faltantes', 'message': f'Faltan {missing} insumos de la matriz Construdata en proveedor'})
    if extra:
        result['findings'].append({'severity': 'media', 'type': 'insumos_extra', 'message': f'Proveedor incluye {extra} insumos no homologados contra Construdata'})
    for row in result['rows']:
        if row.get('estado') != 'match_insumo':
            continue
        if row.get('formula_kind') == 'porcentaje_mo' and isinstance(row.get('delta_cantidad'), (int, float)) and abs(row['delta_cantidad']) >= 0.20:
            result['findings'].append({'severity': 'media', 'type': 'porcentaje_mo', 'message': f"Porcentaje MO diferente en {row.get('provider_desc')}: {row['delta_cantidad']*100:.1f}%"})
        elif isinstance(row.get('delta_cantidad'), (int, float)) and abs(row['delta_cantidad']) >= 0.25:
            result['findings'].append({'severity': 'alta', 'type': 'rendimiento', 'message': f"Rendimiento/cantidad diferente en {row.get('provider_desc')}: {row['delta_cantidad']*100:.1f}%"})
        if isinstance(row.get('delta_precio'), (int, float)) and abs(row['delta_precio']) >= 0.20:
            result['findings'].append({'severity': 'media', 'type': 'precio_insumo', 'message': f"Precio diferente en {row.get('provider_desc')}: {row['delta_precio']*100:.1f}%"})
    return result


def attach_construdata_matrix_analysis(conceptos, precios_nacional):
    construdata_concepts = _get_construdata_concepts(precios_nacional)
    if not construdata_concepts:
        return conceptos
    concept_index = _build_construdata_token_index(construdata_concepts)
    for key, concept in conceptos.items():
        match = match_concepto_construdata(key, concept, construdata_concepts, concept_index=concept_index)
        comparison = compare_provider_matrix_vs_construdata(key, concept, match)
        concept['construdata_concept_match'] = match
        concept['construdata_matrix_comparison'] = comparison
        concept['construdata_matrix_findings'] = comparison.get('findings') or []
    return conceptos


def _build_homologacion_ia_sheet(wb, datos):
    ws = wb.create_sheet('Homologacion IA')
    headers = ['Proveedor','Concepto proveedor','Descripcion proveedor','Unidad','PU proveedor','Tipo match','Estado','Score','Requiere revision','Codigos CD seleccionados','Motivo','Top candidatos']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            m = c.get('construdata_concept_match') or {}
            selected = m.get('selected_matches') or []
            codes = ', '.join((sm.get('concept') or {}).get('codigo') or (sm.get('concept') or {}).get('clave') or '' for sm in selected)
            cand = m.get('candidates') or []
            cand_txt = '; '.join(f"{x.get('codigo')} ({x.get('score')}) {str(x.get('descripcion') or '')[:80]}" for x in cand[:5])
            vals = [d.get('nombre'), clave, c.get('desc'), c.get('unidad'), c.get('pu'), m.get('tipo_match'), m.get('status'), m.get('score'), bool(m.get('requires_review')), codes, m.get('reason'), cand_txt]
            for col,val in enumerate(vals, start=1):
                cell=ws.cell(row=rown, column=col, value=val)
                cell.border=_brd(); cell.font=_fnt(size=8)
                if col == 5 and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
                if col in [3,11,12]: cell.alignment=_aln('left','center', wrap=True)
            rown += 1
    _autosize_sheet(ws, max_width=55)

# override sheets con columnas adicionales paso 5

def _build_construdata_matching_sheet(wb, datos):
    ws = wb.create_sheet('Construdata matching')
    headers = ['Proveedor','Concepto proveedor','Descripcion proveedor','Unidad','PU proveedor','Tipo match','Estado match','Score','Revision','Codigo Construdata','Descripcion Construdata','PU Construdata','Delta PU','Renglones prov.','Renglones CD','Matcheados','Extra prov.','Faltantes prov.','Hallazgos']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            comp = c.get('construdata_matrix_comparison') or {}
            summary = comp.get('summary') or {}
            findings = '; '.join(f.get('message','') for f in (comp.get('findings') or [])[:5])
            vals = [
                d.get('nombre'), clave, c.get('desc'), c.get('unidad'), c.get('pu'), comp.get('match_type'), comp.get('match_status'), comp.get('match_score'), bool(comp.get('requires_review')), comp.get('construdata_codigo'), comp.get('construdata_desc'), comp.get('construdata_pu'), comp.get('delta_pu'),
                summary.get('proveedor_renglones'), summary.get('construdata_renglones'), summary.get('insumos_matcheados'), summary.get('insumos_extra_proveedor'), summary.get('insumos_faltantes_proveedor'), findings
            ]
            for col, val in enumerate(vals, start=1):
                cell = ws.cell(row=rown, column=col, value=val)
                cell.border = _brd(); cell.font = _fnt(size=8)
                if col in [5,12] and isinstance(val, (int, float)): cell.number_format = '$#,##0.00'
                if col == 13 and isinstance(val, (int, float)): cell.number_format = '+0.0%;-0.0%;0.0%'
                if col in [3,11,19]: cell.alignment = _aln('left','center', wrap=True)
            rown += 1
    _autosize_sheet(ws, max_width=55)


def _build_construdata_matrix_detail_sheet(wb, datos):
    ws = wb.create_sheet('Construdata matrices')
    headers = ['Proveedor','Concepto proveedor','Codigo CD','Estado','Concepto CD origen','Tipo','Formula','Insumo proveedor','Unidad prov.','Cant. prov.','Precio/Base prov.','Importe prov.','Insumo CD','Unidad CD','Cant./% CD','Precio/Base CD','Importe CD','Score','Delta cant./%','Delta precio','Delta importe']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            comp = c.get('construdata_matrix_comparison') or {}
            for r in comp.get('rows') or []:
                vals = [
                    d.get('nombre'), clave, comp.get('construdata_codigo'), r.get('estado'), r.get('cd_source_concept'), r.get('tipo'), r.get('formula_kind'), r.get('provider_desc'), r.get('provider_unidad'), r.get('provider_cantidad'), r.get('provider_precio'), r.get('provider_importe'),
                    r.get('cd_desc'), r.get('cd_unidad'), r.get('cd_cantidad'), r.get('cd_precio'), r.get('cd_importe'), r.get('match_score'), r.get('delta_cantidad'), r.get('delta_precio'), r.get('delta_importe')
                ]
                for col, val in enumerate(vals, start=1):
                    cell = ws.cell(row=rown, column=col, value=val)
                    cell.border = _brd(); cell.font = _fnt(size=8)
                    if col in [11,12,16,17] and isinstance(val, (int, float)): cell.number_format = '$#,##0.00'
                    if col in [19,20,21] and isinstance(val, (int, float)): cell.number_format = '+0.0%;-0.0%;0.0%'
                    if col in [8,13]: cell.alignment = _aln('left','center', wrap=True)
                rown += 1
    _autosize_sheet(ws, max_width=50)

def build_comparativo(
    filepaths:     list,
    nombres:       list,
    output:        str,
    meta:          dict,
    catalogo_path: str = None,
    nacional_path: str = None,
):
    """
    Punto de entrada principal.
    """
    n_prov = len(filepaths)
    precios_nacional = {}
    materials_index = []
    if not nacional_path:
        default_cd = _quantia_data_path("construdata_matrices.xlsx")
        default_flat = _quantia_data_path("neodata_precios.xlsx")
        if default_cd.exists():
            nacional_path = str(default_cd)
        elif default_flat.exists():
            nacional_path = str(default_flat)
    if nacional_path:
        try:
            precios_nacional = extract_precios_nacional(nacional_path)
            # Si el benchmark es matriz Construdata, no usamos el indice plano de insumos
            # para el matching de materiales heredado porque puede tener decenas de miles
            # de renglones y volver lenta la ejecucion. La comparacion tecnica se hace
            # en attach_construdata_matrix_analysis() contra la matriz del concepto.
            if str(precios_nacional.get("__benchmark_kind__", "")).startswith("construdata_matrices"):
                materials_index = []
            else:
                materials_index = precios_nacional.get("__materials_index__", [])
        except Exception as e:
            print(f"Advertencia: no se pudo leer catálogo nacional: {e}")
    labor_path = str((_quantia_data_path("Tabulador de Mano de Obra 2024-2025-2026-1.xlsx")))
    labor_table = {}
    labor_alias_index = {}
    if Path(labor_path).exists():
        labor_table = load_labor_tabulador(labor_path, "2026")
        labor_alias_index = build_labor_alias_index(labor_table)
    equipment_index = build_equipment_index(load_equipment_benchmark())
    datos = []
    provider_analyses = []
    for fp, nombre in zip(filepaths, nombres):
        try:
            conceptos = extract_conceptos(fp)
            attach_construdata_matrix_analysis(conceptos, precios_nacional)
            _step08_apply_market_catalog_pricing(conceptos)
            important_scope = optimize_material_search_scope(conceptos)
            material_match_cache = {}
            benchmark_is_matrix = str(precios_nacional.get("__benchmark_kind__", "")).startswith("construdata_matrices")
            for _, c in conceptos.items():
                if benchmark_is_matrix:
                    # Paso 07: cuando existe matriz Construdata real, el mercado se calcula
                    # por concepto homologado + 25% de indirecto. Se omite el matching legado
                    # material/MO/equipo para evitar ruido, falsos positivos y lentitud.
                    c["materiales_benchmark_matches"] = []
                    c["materiales_benchmark_resumen"] = summarize_material_market_matches([])
                    c["mano_obra_benchmark_matches"] = []
                    c["mano_obra_benchmark_resumen"] = summarize_labor_matches([])
                    c["equipo_benchmark_matches"] = []
                    c["equipo_benchmark_resumen"] = summarize_equipment_matches([])
                else:
                    c["materiales_benchmark_matches"] = match_materials_against_benchmark(c.get("materiales_items"), materials_index, important_scope, _, material_match_cache)
                    c["materiales_benchmark_resumen"] = summarize_material_market_matches(c["materiales_benchmark_matches"])
                    c["mano_obra_benchmark_matches"] = match_labor_against_tabulador(c.get("mano_obra_items"), labor_alias_index, labor_table)
                    c["mano_obra_benchmark_resumen"] = summarize_labor_matches(c["mano_obra_benchmark_matches"])
                    c["equipo_benchmark_matches"] = match_equipment_against_benchmark(c.get("equipo_items"), equipment_index)
                    c["equipo_benchmark_resumen"] = summarize_equipment_matches(c["equipo_benchmark_matches"])
            total_estimado = round(sum(v["total"] for v in conceptos.values() if isinstance(v.get("total"), (int,float))),2)
            datos.append({'nombre': nombre, 'conceptos': conceptos, 'total_estimado': total_estimado})
            provider_analyses.append(calculate_executive_findings(nombre, conceptos, total_estimado, []))
        except Exception as e:
            datos.append({'nombre': nombre, 'conceptos': {}, 'error': str(e), 'total_estimado': 0.0})
    if n_prov == 1 and datos:
        meta = dict(meta or {})
        meta.setdefault("archivo_fuente", filepaths[0] if filepaths else "")
        return _build_single_provider_pmd(output, meta, datos[0], provider_analyses[0] if provider_analyses else {"provider_name": nombres[0], "drivers": []})
    orden_catalogo = []
    if catalogo_path:
        try:
            orden_catalogo = extract_orden_catalogo(catalogo_path)
        except Exception as e:
            print(f"Advertencia: no se pudo leer catálogo: {e}")
    if orden_catalogo:
        claves_en_catalogo = {x['clave'] for x in orden_catalogo if x['tipo']=='concepto'}
        todas_claves_set = set()
        for d in datos:
            todas_claves_set.update(d['conceptos'].keys())
        solo_prov = [c for c in todas_claves_set if c not in claves_en_catalogo]
        orden_final = orden_catalogo.copy()
        if solo_prov:
            orden_final.append({'tipo':'seccion','clave':'EX','desc':'CONCEPTOS ADICIONALES DE PROVEEDORES'})
            for c in sorted(solo_prov):
                orden_final.append({'tipo':'concepto','clave':c,'desc':'','uni':'','cant':0,'seccion':'ADICIONAL'})
    else:
        orden_final = []
        vistas = set()
        for d in datos:
            for clave in d['conceptos']:
                if clave not in vistas:
                    orden_final.append({'tipo':'concepto','clave':clave,'desc':'','uni':'','cant':0,'seccion':''})
                    vistas.add(clave)
    todas_claves = [x['clave'] for x in orden_final if x['tipo']=='concepto']
    totales = []
    for d in datos:
        t = sum(v['total'] for v in d['conceptos'].values() if isinstance(v.get('total'), (int,float)))
        totales.append(t)
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativa"
    ws.freeze_panes = "D9"
    _build_header(ws, meta, datos, n_prov, precios_nacional)
    data_rows = _build_data(ws, orden_final, datos, totales, n_prov, precios_nacional)
    _build_totals_row(ws, data_rows, totales, n_prov, precios_nacional)
    _build_resumen(wb, datos, totales, todas_claves, meta, precios_nacional)
    _build_market_materials_sheet(wb, datos)
    _build_market_labor_sheet(wb, datos)
    _build_drivers_sheet(wb, provider_analyses)
    _build_construdata_matching_sheet(wb, datos)
    _build_homologacion_ia_sheet(wb, datos)
    _build_construdata_matrix_detail_sheet(wb, datos)
    _build_totales_mercado_sheet(wb, datos)
    _build_hallazgos_ejecutivos_sheet(wb, datos)
    _build_faltantes_matriz_base_sheet(wb, datos)
    _build_elementos_fuera_matriz_sheet(wb, datos)
    wb.save(output)
    return output

def _build_header(ws, meta, datos, n_prov, precios_nacional={}):
    """Construye las filas de encabezado del comparativo."""

    # Definir anchos de columna
    # A=Partida, B=Descripción, C=Uni, + bloques por proveedor (P.U., Importe, %Part, %Dif) + Ganador + Obs
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 7

    tiene_nacional = bool(precios_nacional)
    col = 4  # Empezar en D
    for i in range(n_prov):
        ws.column_dimensions[chr(64+col)].width = 13   # P.U.
        ws.column_dimensions[chr(64+col+1)].width = 15 # Importe
        ws.column_dimensions[chr(64+col+2)].width = 8  # % Part
        ws.column_dimensions[chr(64+col+3)].width = 9  # % Dif
        col += 4

    # Columna Catálogo Nacional (siempre al final de proveedores)
    if tiene_nacional:
        ws.column_dimensions[chr(64+col)].width = 14   # P.U. Nacional
        ws.column_dimensions[chr(64+col+1)].width = 10 # % vs Nacional
        col += 2

    ws.column_dimensions[chr(64+col)].width = 16    # Ganador
    ws.column_dimensions[chr(64+col+1)].width = 14  # Ahorro
    ws.column_dimensions[chr(64+col+2)].width = 22  # Observaciones

    last_col = chr(64 + col + 2)

    # Fila 1: Título
    ws.row_dimensions[1].height = 34
    ws.merge_cells(f'A1:{last_col}1')
    c = ws['A1']
    c.value     = f"COMPARATIVO DE COTIZACIONES — {meta.get('proyecto','').upper()}"
    c.font      = Font(bold=True, size=13, color=C_BLANCO, name="Calibri")
    c.fill      = _fill(C_AZUL_OSC)
    c.alignment = _aln('center', 'center')

    # Filas 2-4: Info proyecto
    info = [
        ('Cliente:',   meta.get('cliente', '')),
        ('Proyecto:',  meta.get('proyecto', '')),
        ('Ubicación:', meta.get('ubicacion', '')),
    ]
    for i, (lbl, val) in enumerate(info, 2):
        ws.row_dimensions[i].height = 15
        ws.merge_cells(f'A{i}:B{i}')
        c = ws[f'A{i}']
        c.value = lbl; c.font = _fnt(bold=True, color=C_BLANCO, size=9)
        c.fill  = _fill("2E75B6"); c.alignment = _aln('right')
        ws.merge_cells(f'C{i}:{last_col}{i}')
        c = ws[f'C{i}']
        c.value = val; c.font = _fnt(size=9); c.alignment = _aln('left')

    # Fila 5: Nombres de proveedores (agrupados)
    ws.row_dimensions[5].height = 18
    col = 4
    for i, d in enumerate(datos):
        start = chr(64+col); end = chr(64+col+3)
        ws.merge_cells(f'{start}5:{end}5')
        c = ws[f'{start}5']
        c.value = d['nombre']
        c.font  = Font(bold=True, size=10, color=C_BLANCO, name="Calibri")
        c.fill  = _fill(PROV_COLORS[i % len(PROV_COLORS)])
        c.alignment = _aln('center', 'center')
        col += 4

    # Celda Catálogo Nacional en fila 5
    if tiene_nacional:
        ws.merge_cells(f'{chr(64+col)}5:{chr(64+col+1)}5')
        c = ws[f'{chr(64+col)}5']
        c.value = '📊 Catálogo Nacional'
        c.font  = Font(bold=True, size=9, color=C_BLANCO, name="Calibri")
        c.fill  = _fill("4A235A"); c.alignment = _aln('center','center')
        col += 2

    # Celda Análisis
    ws.merge_cells(f'{chr(64+col)}5:{last_col}5')
    c = ws[f'{chr(64+col)}5']
    c.value = 'ANÁLISIS'
    c.font  = Font(bold=True, size=10, color=C_BLANCO, name="Calibri")
    c.fill  = _fill(C_AZUL_OSC); c.alignment = _aln('center','center')

    # Fila 6: divisor
    ws.row_dimensions[6].height = 4
    ws.merge_cells(f'A6:{last_col}6')
    ws['A6'].fill = _fill(C_AZUL_OSC)

    # Fila 7: sub-headers
    ws.row_dimensions[7].height = 14
    col = 4
    for i, d in enumerate(datos):
        for j, sub in enumerate(['P.U.', 'Importe', '% Part.', '% Dif.']):
            ref = f'{chr(64+col+j)}7'
            c = ws[ref]; c.value = sub
            c.font  = _fnt(bold=True, color=C_BLANCO, size=8)
            c.fill  = _fill(PROV_COLORS[i % len(PROV_COLORS)])
            c.alignment = _aln('center','center'); c.border = _brd()
        col += 4

    if tiene_nacional:
        for j, sub in enumerate(['P.U. Nac.', '% vs Nac.']):
            ref = f'{chr(64+col+j)}7'
            c = ws[ref]; c.value = sub
            c.font  = _fnt(bold=True, color=C_BLANCO, size=8)
            c.fill  = _fill("4A235A")
            c.alignment = _aln('center','center'); c.border = _brd()
        col += 2

    for ref, txt in [(f'{chr(64+col)}7','🏆 Mejor'),(f'{chr(64+col+1)}7','Ahorro $'),(f'{chr(64+col+2)}7','Observaciones')]:
        c = ws[ref]; c.value = txt
        c.font = _fnt(bold=True,color=C_BLANCO,size=8); c.fill=_fill(C_AZUL_OSC)
        c.alignment=_aln('center','center'); c.border=_brd()

    # Fila 8: headers principales
    ws.row_dimensions[8].height = 36
    for ref, txt in [('A8','Partida'),('B8','Descripción'),('C8','Uni.')]:
        c = ws[ref]; c.value = txt
        c.font = Font(bold=True,size=9,color=C_BLANCO,name="Calibri")
        c.fill = _fill(C_AZUL_OSC); c.alignment=_aln('center','center',wrap=True); c.border=_brd()

    col = 4
    for i, d in enumerate(datos):
        for j in range(4):
            c = ws[f'{chr(64+col+j)}8']
            c.font = Font(bold=True,size=9,color=C_BLANCO,name="Calibri")
            c.fill = _fill(PROV_COLORS[i % len(PROV_COLORS)])
            c.alignment=_aln('center','center',wrap=True); c.border=_brd()
        col += 4

    for ref in [f'{chr(64+col)}8',f'{chr(64+col+1)}8',f'{chr(64+col+2)}8']:
        c=ws[ref]; c.font=Font(bold=True,size=9,color=C_BLANCO,name="Calibri")
        c.fill=_fill(C_AZUL_OSC); c.alignment=_aln('center','center'); c.border=_brd()


def _build_data(ws, orden_final, datos, totales, n_prov, precios_nacional={}):
    """Escribe las filas de conceptos. Devuelve lista de row numbers con datos."""
    ROW = 9
    data_rows = []
    alt = False

    tiene_nacional  = bool(precios_nacional)
    extra_nac       = 2 if tiene_nacional else 0
    last_data_col   = 3 + n_prov * 4 + extra_nac
    last_col_letter = chr(64 + last_data_col + 3)

    # Normalizar orden_final a lista de dicts si llegó como lista de strings (compatibilidad)
    if orden_final and isinstance(orden_final[0], str):
        orden_final = [{'tipo':'concepto','clave':c,'desc':'','uni':'','cant':0,'seccion':''} for c in orden_final]

    for item in orden_final:
        # ── FILA DE SECCIÓN ──
        if item['tipo'] == 'seccion':
            ws.row_dimensions[ROW].height = 15
            ws.merge_cells(f'A{ROW}:{last_col_letter}{ROW}')
            c = ws[f'A{ROW}']
            c.value = f"  {item.get('clave','')}   {item.get('desc','')}"
            c.font  = Font(bold=True, size=9, color="000000", name="Calibri")
            c.fill  = _fill(C_GRIS_SEC); c.alignment=_aln('left','center'); c.border=_brd()
            ROW += 1
            alt = False
            continue

        clave = item['clave']

        ws.row_dimensions[ROW].height = 28
        bg = C_GRIS_FIL if alt else C_BLANCO

        def sc(col_letter, val, fmt=None, bold=False, bg_ov=None, h='center'):
            c = ws[f'{col_letter}{ROW}']
            c.value = val
            c.font  = _fnt(bold=bold, size=9)
            c.fill  = _fill(bg_ov or bg)
            c.alignment = _aln(h, 'center', wrap=(col_letter=='B'))
            c.border = _brd()
            if fmt: c.number_format = fmt

        # Tomar desc/uni/cant del catálogo primero, luego del proveedor
        desc = item.get('desc') or ''
        uni  = item.get('uni')  or ''
        cant = item.get('cant') or 0
        if not desc or not uni:
            for d in datos:
                if clave in d['conceptos']:
                    if not desc: desc = d['conceptos'][clave].get('desc','')
                    if not uni:  uni  = d['conceptos'][clave].get('unidad','')
                    if not cant: cant = d['conceptos'][clave].get('cantidad',0)
                    break

        sc('A', clave,  h='left')
        sc('B', desc,   h='left')
        sc('C', uni)

        # Bloque por proveedor
        col = 4
        importes = []
        for i, d in enumerate(datos):
            c_data = d['conceptos'].get(clave)
            pu     = c_data['pu']    if c_data else None
            imp    = c_data['total'] if c_data else None
            pct_p  = imp / totales[i] if (imp and totales[i]) else None

            sc(chr(64+col),   pu,    '"$"#,##0.00')
            sc(chr(64+col+1), imp,   '"$"#,##0.00')
            sc(chr(64+col+2), pct_p, '0.00%')
            sc(chr(64+col+3), None)  # % Dif — se calcula después

            importes.append((col, imp))
            col += 4

        # % Diferencia vs mínimo
        vals_validos = [(c2, v) for c2,v in importes if v]
        if len(vals_validos) >= 2:
            min_v = min(v for _,v in vals_validos)
            max_v = max(v for _,v in vals_validos)
            for c2, v in importes:
                if v is None: continue
                pct_dif = (v - min_v) / min_v if (min_v and min_v != 0) else 0
                cell = ws[f'{chr(64+c2+3)}{ROW}']
                cell.value        = pct_dif
                cell.number_format = '+0.0%;-0.0%;"-"'
                cell.font         = _fnt(size=9)
                cell.border       = _brd()
                cell.alignment    = _aln('center','center')
                if pct_dif == 0:
                    cell.fill = _fill(C_VERDE_OK)
                elif pct_dif > 0.1:
                    cell.fill = _fill(C_ROJO_MAL)
                else:
                    cell.fill = _fill(C_AMARILLO)

        # Columnas Catálogo Nacional
        pu_nac = None
        if precios_nacional:
            nac = precios_nacional.get(clave) or precios_nacional.get(_base_clave(clave))
            pu_nac = nac['pu_mercado'] if nac else None

            c_pu = ws[f'{chr(64+col)}{ROW}']
            c_pu.value = pu_nac
            c_pu.number_format = '"$"#,##0.00'
            c_pu.font = _fnt(size=9, color="FFFFFF")
            c_pu.fill = _fill("4A235A" if pu_nac else "6C3483")
            c_pu.alignment = _aln('center','center'); c_pu.border = _brd()

            # % vs Nacional para cada proveedor (retroactivo en col+3 de cada uno)
            if pu_nac:
                col_prov = 4
                for d in datos:
                    c_data = d['conceptos'].get(clave)
                    pu_p   = c_data['pu'] if c_data and c_data.get('pu') else None
                    if pu_p:
                        pct_vs_nac = (pu_p - pu_nac) / pu_nac if pu_nac and pu_nac != 0 else 0
                        # Escribir en la col de % Dif de ese proveedor
                        cell = ws[f'{chr(64+col_prov+3)}{ROW}']
                        cell.value = pct_vs_nac
                        cell.number_format = '+0.0%;-0.0%;"-"'
                        cell.font = _fnt(size=9)
                        cell.border = _brd()
                        cell.alignment = _aln('center','center')
                        if abs(pct_vs_nac) <= 0.05:
                            cell.fill = _fill(C_VERDE_OK)
                        elif pct_vs_nac > 0.1:
                            cell.fill = _fill(C_ROJO_MAL)
                        else:
                            cell.fill = _fill(C_AMARILLO)
                    col_prov += 4

            # % vs Nacional en col separada
            c_pct = ws[f'{chr(64+col+1)}{ROW}']
            if pu_nac and vals_validos:
                min_pu = min(d['conceptos'][clave]['pu'] for d in datos
                             if clave in d['conceptos'] and d['conceptos'][clave].get('pu'))
                pct_vs = (min_pu - pu_nac) / pu_nac if pu_nac and pu_nac != 0 else 0
                c_pct.value = pct_vs
                c_pct.number_format = '+0.0%;-0.0%;"-"'
                c_pct.fill = _fill(C_VERDE_OK if pct_vs <= 0 else C_ROJO_MAL)
            else:
                c_pct.value = None
            c_pct.font = _fnt(size=9, color="FFFFFF" if pu_nac else "000000")
            c_pct.fill = _fill("4A235A") if not pu_nac else c_pct.fill
            c_pct.alignment = _aln('center','center'); c_pct.border = _brd()
            col += 2

        # Ganador + Ahorro
        # En modo 1 proveedor vs Nacional, el "ganador" es la comparación vs mercado
        if n_prov == 1 and precios_nacional:
            c_data = datos[0]['conceptos'].get(clave)
            pu_p   = c_data['pu'] if c_data and c_data.get('pu') else None
            if pu_p and pu_nac:
                pct = (pu_p - pu_nac) / pu_nac if pu_nac and pu_nac != 0 else 0
                estado = '✓ Bajo mercado' if pct <= 0 else ('⚠ Sobre mercado' if pct <= 0.1 else '❌ Excede mercado')
                color_e = C_VERDE_OK if pct <= 0 else (C_AMARILLO if pct <= 0.1 else C_ROJO_MAL)
                sc(chr(64+col),   estado, bold=True, bg_ov=color_e)
                sc(chr(64+col+1), abs((pu_p - pu_nac) * (c_data.get('cantidad') or 1)),
                   '"$"#,##0.00', bg_ov=color_e)
            else:
                sc(chr(64+col), 'Sin ref.'); sc(chr(64+col+1), None)
        elif len(vals_validos) >= 2:
            min_col, min_val = min(vals_validos, key=lambda x: x[1])
            max_col, max_val = max(vals_validos, key=lambda x: x[1])
            prov_idx = (min_col - 4) // 4
            ganador  = datos[prov_idx]['nombre']
            ahorro   = max_val - min_val
            sc(chr(64+col),   ganador, bold=True, bg_ov=C_VERDE_OK)
            sc(chr(64+col+1), ahorro,  '"$"#,##0.00', bg_ov=C_VERDE_OK)
        else:
            sc(chr(64+col),   'Sin datos')
            sc(chr(64+col+1), None)

        # Observaciones / alertas
        observaciones = []
        if precios_nacional and pu_nac:
            for d in datos:
                c_data_obs = d['conceptos'].get(clave)
                pu_obs = c_data_obs['pu'] if c_data_obs and c_data_obs.get('pu') else None
                if pu_obs:
                    pct_obs = (pu_obs - pu_nac) / pu_nac if pu_nac else 0
                    if pct_obs > 0.15:
                        observaciones.append(f"{d['nombre']}: sobrecoste {pct_obs*100:.1f}%")
                    elif pct_obs < -0.1:
                        observaciones.append(f"{d['nombre']}: por debajo {-pct_obs*100:.1f}%")
        sc(chr(64+col+2), '; '.join(observaciones[:2]) if observaciones else '', h='left')  # Observaciones

        data_rows.append(ROW)
        ROW += 1
        alt = not alt

    return data_rows


def _build_totals_row(ws, data_rows, totales, n_prov, precios_nacional={}):
    """Fila de totales generales."""
    if not data_rows:
        return
    ROW = max(data_rows) + 1
    ws.row_dimensions[ROW].height = 20

    tiene_nacional  = bool(precios_nacional)
    extra_nac       = 2 if tiene_nacional else 0
    last_data_col = 3 + n_prov * 4 + extra_nac
    last_col_letter = chr(64 + last_data_col + 3)

    def sc(col_letter, val, fmt=None):
        c = ws[f'{col_letter}{ROW}']
        c.value = val; c.font=Font(bold=True,size=9,color=C_BLANCO,name="Calibri")
        c.fill=_fill(C_AZUL_OSC); c.border=_brd(); c.alignment=_aln('center','center')
        if fmt: c.number_format=fmt

    sc('A','TOTAL GENERAL')
    ws.merge_cells(f'A{ROW}:C{ROW}')

    col = 4
    for i in range(n_prov):
        # Sumar solo columnas de importe
        first = data_rows[0]; last = data_rows[-1]
        imp_col = chr(64+col+1)
        ws[f'{imp_col}{ROW}'].value = f'=SUM({imp_col}{first}:{imp_col}{last})'
        ws[f'{imp_col}{ROW}'].number_format = '"$"#,##0.00'
        ws[f'{imp_col}{ROW}'].font = Font(bold=True,size=9,color=C_BLANCO,name="Calibri")
        ws[f'{imp_col}{ROW}'].fill = _fill(C_AZUL_OSC)
        ws[f'{imp_col}{ROW}'].border = _brd()
        ws[f'{imp_col}{ROW}'].alignment = _aln('center','center')
        for j in [0,2,3]:
            sc(chr(64+col+j), '')
        col += 4

    for extra in range(3):
        sc(chr(64+col+extra), '')


def _build_resumen(wb, datos, totales, todas_claves, meta, precios_nacional={}):
    """Hoja de Resumen Ejecutivo."""
    ws = wb.create_sheet("Resumen Ejecutivo")
    ws.column_dimensions['A'].width = 38
    for i in range(len(datos)):
        ws.column_dimensions[chr(66+i)].width = 18

    # Título
    ws.row_dimensions[1].height = 30
    last = chr(65 + len(datos))
    ws.merge_cells(f'A1:{last}1')
    c=ws['A1']; c.value='RESUMEN EJECUTIVO'
    c.font=Font(bold=True,size=13,color=C_BLANCO,name="Calibri")
    c.fill=_fill(C_AZUL_OSC); c.alignment=_aln('center','center')

    # Proyecto
    ws.row_dimensions[2].height=15
    ws['A2'].value=f"Proyecto: {meta.get('proyecto','')}"
    ws['A2'].font=_fnt(size=9)

    # Headers
    ws.row_dimensions[4].height=18
    ws['A4'].value='Indicador'; ws['A4'].font=_fnt(bold=True,color=C_BLANCO,size=9)
    ws['A4'].fill=_fill(C_AZUL_OSC); ws['A4'].border=_brd(); ws['A4'].alignment=_aln('left','center')
    for i,d in enumerate(datos):
        c=ws[f'{chr(66+i)}4']; c.value=d['nombre']
        c.font=Font(bold=True,size=9,color=C_BLANCO,name="Calibri")
        c.fill=_fill(PROV_COLORS[i%len(PROV_COLORS)]); c.border=_brd(); c.alignment=_aln('center','center')

    # Datos de resumen
    conceptos_comunes = set(datos[0]['conceptos'].keys())
    for d in datos[1:]:
        conceptos_comunes &= set(d['conceptos'].keys())

    resumen_rows = [
        ('Total oferta',        [f"${t:,.2f}" for t in totales]),
        ('No. conceptos cotizados', [str(len(d['conceptos'])) for d in datos]),
        ('Conceptos en común',  [str(len(conceptos_comunes))] * len(datos)),
        ('Conceptos exclusivos', [str(len(d['conceptos']) - len(conceptos_comunes)) for d in datos]),
    ]

    # Ganadores por concepto
    gana = [0] * len(datos)
    for clave in todas_claves:
        vals = []
        for i, d in enumerate(datos):
            if clave in d['conceptos']:
                t = d['conceptos'][clave].get('total')
                if isinstance(t,(int,float)):
                    vals.append((i,t))
        if len(vals) >= 2:
            min_i = min(vals, key=lambda x:x[1])[0]
            gana[min_i] += 1
    resumen_rows.append(('Conceptos donde es más económico', [str(g) for g in gana]))

    # Diferencia vs más barato
    if len(totales) >= 2:
        min_total = min(totales)
        difs = [f"+${t-min_total:,.2f} ({(t-min_total)/min_total*100:.1f}%)" if (t!=min_total and min_total!=0) else ("✓ MÁS ECONÓMICO" if t==min_total else f"+${t-min_total:,.2f}") for t in totales]
        resumen_rows.append(('Diferencia vs oferta más económica', difs))


    # Score por proveedor (precio + referencia nacional + cobertura)
    if precios_nacional:
        score_vals = []
        min_total = min(totales) if totales else 0
        for idx_prov, d in enumerate(datos):
            total = totales[idx_prov] if idx_prov < len(totales) else 0
            precio_score = 40 if (min_total and total == min_total) else max(0, 40 * (min_total / total)) if total else 0

            comparables = 0
            sobrecostes = 0
            under = 0
            for clave in todas_claves:
                cdat = d['conceptos'].get(clave)
                nac = precios_nacional.get(clave) or precios_nacional.get(_base_clave(clave))
                pu_nac = nac['pu_mercado'] if nac else None
                pu_p = cdat.get('pu') if cdat else None
                if pu_nac and pu_p:
                    comparables += 1
                    pct = (pu_p - pu_nac) / pu_nac if pu_nac else 0
                    if pct > 0.15:
                        sobrecostes += 1
                    elif pct <= 0:
                        under += 1

            ref_score = 0
            if comparables:
                ref_score = max(0, 30 - (sobrecostes / comparables) * 30)

            cobertura = len(d['conceptos']) / len(todas_claves) if todas_claves else 0
            cobertura_score = min(10, cobertura * 10)

            gana_count = gana[idx_prov] if idx_prov < len(gana) else 0
            consistencia_score = min(20, (gana_count / len(todas_claves)) * 20) if todas_claves else 0

            total_score = round(precio_score + ref_score + cobertura_score + consistencia_score, 1)
            score_vals.append(f"{total_score}/100")

        resumen_rows.append(('Score proveedor', score_vals))

    alt2 = False
    for i,(lbl,vals) in enumerate(resumen_rows,5):
        ws.row_dimensions[i].height=18
        bg = C_GRIS_FIL if alt2 else C_BLANCO
        c=ws[f'A{i}']; c.value=lbl; c.font=_fnt(size=9,bold=True)
        c.fill=_fill(bg); c.border=_brd(); c.alignment=_aln('left','center')
        for j,v in enumerate(vals):
            cell=ws[f'{chr(66+j)}{i}']; cell.value=v; cell.font=_fnt(size=9)
            # Resaltar al más económico en fila de totales
            is_best = (i==5 and totales[j]==min(totales))
            cell.fill=_fill(C_VERDE_OK if is_best else bg)
            cell.border=_brd(); cell.alignment=_aln('center','center')
        alt2=not alt2



# -----------------------------------------------------------------------------
# Paso 07 - Totales reales vs mercado + hallazgos ejecutivos consolidados
# -----------------------------------------------------------------------------

def _iter_conceptos_datos(datos):
    for d in datos or []:
        for clave, concepto in (d.get('conceptos') or {}).items():
            yield d, clave, concepto


def _flatten_construdata_findings(concepto):
    comp = concepto.get('construdata_matrix_comparison') or {}
    for f in comp.get('findings') or []:
        yield {
            'tipo': f.get('type') or 'CONSTRUDATA',
            'severidad': f.get('severity') or 'media',
            'mensaje': f.get('message') or '',
            'origen': 'Construdata matrices',
        }


def _indirect_finding(concepto):
    _amt, pct = _concept_indirect_amount_and_pct(concepto)
    if isinstance(pct, (int, float)):
        if pct > MARKET_INDIRECT_PCT + 0.001:
            return {
                'tipo': 'INDIRECTO_ALTO',
                'severidad': 'alta' if pct >= 0.30 else 'media',
                'mensaje': f"Indirecto declarado {pct*100:.2f}% vs mercado {MARKET_INDIRECT_PCT*100:.2f}%.",
                'origen': 'Indirecto proveedor',
            }
        return {
            'tipo': 'INDIRECTO_OK',
            'severidad': 'baja',
            'mensaje': f"Indirecto declarado {pct*100:.2f}% vs mercado {MARKET_INDIRECT_PCT*100:.2f}%.",
            'origen': 'Indirecto proveedor',
        }
    return {
        'tipo': 'INDIRECTO_NO_DETECTADO',
        'severidad': 'media',
        'mensaje': f"No se detectó indirecto explícito del proveedor; mercado se calcula siempre con {MARKET_INDIRECT_PCT*100:.2f}%.",
        'origen': 'Indirecto proveedor',
    }


def _rendimiento_findings(concepto):
    comp = concepto.get('construdata_matrix_comparison') or {}
    for r in comp.get('rows') or []:
        if r.get('estado') != 'match_insumo':
            continue
        delta = r.get('delta_cantidad')
        if isinstance(delta, (int, float)) and abs(delta) >= 0.25:
            yield {
                'tipo': 'RENDIMIENTO',
                'severidad': 'alta' if abs(delta) >= 0.50 else 'media',
                'mensaje': f"Rendimiento/cantidad diferente en {r.get('provider_desc') or r.get('cd_desc')}: {delta*100:.1f}% vs matriz base.",
                'origen': 'Comparación matriz',
            }
        if r.get('formula_kind') == 'porcentaje_mo' and isinstance(delta, (int, float)) and abs(delta) >= 0.15:
            yield {
                'tipo': 'PORCENTAJE_MO',
                'severidad': 'media',
                'mensaje': f"Porcentaje sobre MO diferente en {r.get('provider_desc') or r.get('cd_desc')}: {delta*100:.1f}%.",
                'origen': 'Comparación matriz',
            }


def _concept_hallazgos_ejecutivos(concepto):
    hallazgos = []
    provider_total = _concept_total_amount(concepto)
    market_total = _concept_market_total_amount(concepto)
    if isinstance(provider_total, (int, float)) and isinstance(market_total, (int, float)) and market_total:
        delta = (provider_total - market_total) / market_total
        if abs(delta) >= 0.05:
            hallazgos.append({
                'tipo': 'DIFERENCIA_MERCADO',
                'severidad': 'alta' if abs(delta) >= 0.20 else 'media',
                'mensaje': f"Diferencia vs mercado {delta*100:.1f}% considerando indirecto de mercado del {MARKET_INDIRECT_PCT*100:.0f}%.",
                'origen': 'Totales mercado',
            })
    else:
        hallazgos.append({
            'tipo': 'SIN_TOTAL_MERCADO',
            'severidad': 'media',
            'mensaje': 'No se pudo calcular total de mercado porque el concepto no tiene match Construdata usable.',
            'origen': 'Totales mercado',
        })
    ind = _indirect_finding(concepto)
    if ind:
        hallazgos.append(ind)
    hallazgos.extend(_flatten_construdata_findings(concepto))
    hallazgos.extend(_rendimiento_findings(concepto))
    # Quitar duplicados por mensaje para no saturar la hoja.
    seen = set()
    out = []
    for h in hallazgos:
        key = (h.get('tipo'), h.get('mensaje'))
        if key in seen:
            continue
        seen.add(key)
        out.append(h)
    return out


def _build_totales_mercado_sheet(wb, datos):
    ws = wb.create_sheet('Totales mercado')
    headers = ['Proveedor','Concepto','Descripción','Unidad','Cantidad','PU proveedor','Total proveedor','PU CD directo','PU mercado +25%','Total mercado','Diferencia $','Diferencia %','Estado','Match CD','Tipo match']
    _sheet_headers(ws, headers)
    rown = 2
    totals = defaultdict(lambda: {'prov': 0.0, 'market': 0.0})
    for d, clave, c in _iter_conceptos_datos(datos):
        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'), (int, float)) else 1
        provider_total = _concept_total_amount(c)
        market_direct = _market_component_total(c)
        market_pu = _market_pu_with_indirect(c)
        market_total = _concept_market_total_amount(c)
        diff = (provider_total - market_total) if isinstance(provider_total, (int, float)) and isinstance(market_total, (int, float)) else None
        diff_pct = (diff / market_total) if isinstance(diff, (int, float)) and market_total else None
        comp = c.get('construdata_matrix_comparison') or {}
        vals = [d.get('nombre'), clave, c.get('desc'), c.get('unidad'), cantidad, c.get('pu'), provider_total, market_direct, market_pu, market_total, diff, diff_pct, _concept_market_status(provider_total, market_total), comp.get('construdata_codigo'), comp.get('match_type')]
        for col, val in enumerate(vals, start=1):
            cell = ws.cell(row=rown, column=col, value=val)
            cell.border = _brd(); cell.font = _fnt(size=8)
            if col in [6,7,8,9,10,11] and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
            if col == 12 and isinstance(val, (int,float)): cell.number_format = '+0.0%;-0.0%;0.0%'
            if col in [3,14]: cell.alignment = _aln('left','center', wrap=True)
            else: cell.alignment = _aln('center','center')
        if isinstance(provider_total, (int,float)): totals[d.get('nombre')]['prov'] += provider_total
        if isinstance(market_total, (int,float)): totals[d.get('nombre')]['market'] += market_total
        rown += 1
    rown += 1
    ws.cell(row=rown, column=1, value='RESUMEN GLOBAL').font = _fnt(bold=True, size=9)
    rown += 1
    for provider, t in totals.items():
        prov = t['prov']; market = t['market']; diff = prov - market if market else None; pct = diff/market if market else None
        vals = [provider, '', '', '', '', '', prov, '', '', market, diff, pct, _concept_market_status(prov, market), '', '']
        for col, val in enumerate(vals, start=1):
            cell = ws.cell(row=rown, column=col, value=val)
            cell.border = _brd(); cell.font = _fnt(bold=True, size=8)
            if col in [7,10,11] and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
            if col == 12 and isinstance(val, (int,float)): cell.number_format = '+0.0%;-0.0%;0.0%'
        rown += 1
    _autosize_sheet(ws, max_width=55)


def _build_hallazgos_ejecutivos_sheet(wb, datos):
    ws = wb.create_sheet('Hallazgos Ejecutivos')
    headers = ['Proveedor','Concepto','Descripción','Match CD','Severidad','Tipo hallazgo','Hallazgo','PU proveedor','PU mercado +25%','Total proveedor','Total mercado','Diferencia %','Origen']
    _sheet_headers(ws, headers)
    rown = 2
    for d, clave, c in _iter_conceptos_datos(datos):
        comp = c.get('construdata_matrix_comparison') or {}
        provider_total = _concept_total_amount(c)
        market_total = _concept_market_total_amount(c)
        market_pu = _market_pu_with_indirect(c)
        diff_pct = ((provider_total - market_total) / market_total) if isinstance(provider_total, (int,float)) and isinstance(market_total, (int,float)) and market_total else None
        for h in _concept_hallazgos_ejecutivos(c):
            vals = [d.get('nombre'), clave, c.get('desc'), comp.get('construdata_codigo'), h.get('severidad'), h.get('tipo'), h.get('mensaje'), c.get('pu'), market_pu, provider_total, market_total, diff_pct, h.get('origen')]
            for col, val in enumerate(vals, start=1):
                cell = ws.cell(row=rown, column=col, value=val)
                cell.border = _brd(); cell.font = _fnt(size=8)
                if col in [8,9,10,11] and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
                if col == 12 and isinstance(val, (int,float)): cell.number_format = '+0.0%;-0.0%;0.0%'
                if col in [3,7]: cell.alignment = _aln('left','center', wrap=True)
                else: cell.alignment = _aln('center','center')
            sev = str(h.get('severidad') or '').lower()
            if sev == 'alta':
                for cc in range(1, len(headers)+1): ws.cell(row=rown, column=cc).fill = _fill(C_ROJO_MAL)
            elif sev == 'media':
                for cc in range(1, len(headers)+1): ws.cell(row=rown, column=cc).fill = _fill(C_AMARILLO)
            rown += 1
    _autosize_sheet(ws, max_width=70)


def _build_faltantes_matriz_base_sheet(wb, datos):
    ws = wb.create_sheet('Faltantes matriz base')
    headers = ['Proveedor','Concepto proveedor','Match CD','Tipo','Código CD','Insumo faltante en proveedor','Unidad','Cantidad/% CD','Precio/Base CD','Importe CD','Formula','Concepto CD origen']
    _sheet_headers(ws, headers)
    rown = 2
    for d, clave, c in _iter_conceptos_datos(datos):
        comp = c.get('construdata_matrix_comparison') or {}
        for r in comp.get('rows') or []:
            if r.get('estado') != 'faltante_en_proveedor':
                continue
            vals = [d.get('nombre'), clave, comp.get('construdata_codigo'), r.get('tipo'), r.get('cd_codigo'), r.get('cd_desc'), r.get('cd_unidad'), r.get('cd_cantidad'), r.get('cd_precio'), r.get('cd_importe'), r.get('formula_kind'), r.get('cd_source_concept')]
            for col, val in enumerate(vals, start=1):
                cell = ws.cell(row=rown, column=col, value=val)
                cell.border = _brd(); cell.font = _fnt(size=8)
                if col in [9,10] and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
                if col == 6: cell.alignment = _aln('left','center', wrap=True)
            rown += 1
    _autosize_sheet(ws, max_width=60)


def _build_elementos_fuera_matriz_sheet(wb, datos):
    ws = wb.create_sheet('Elementos fuera de matriz')
    headers = ['Proveedor','Concepto proveedor','Match CD','Tipo','Código proveedor','Elemento proveedor no homologado','Unidad','Cantidad','Precio','Importe','Nota']
    _sheet_headers(ws, headers)
    rown = 2
    for d, clave, c in _iter_conceptos_datos(datos):
        comp = c.get('construdata_matrix_comparison') or {}
        for r in comp.get('rows') or []:
            if r.get('estado') != 'sin_match_insumo':
                continue
            nota = 'Elemento declarado por proveedor fuera de la matriz base homologada; revisar si corresponde a alcance adicional, concepto compuesto o clasificación diferente.'
            vals = [d.get('nombre'), clave, comp.get('construdata_codigo'), r.get('tipo'), r.get('provider_codigo'), r.get('provider_desc'), r.get('provider_unidad'), r.get('provider_cantidad'), r.get('provider_precio'), r.get('provider_importe'), nota]
            for col, val in enumerate(vals, start=1):
                cell = ws.cell(row=rown, column=col, value=val)
                cell.border = _brd(); cell.font = _fnt(size=8)
                if col in [9,10] and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
                if col in [6,11]: cell.alignment = _aln('left','center', wrap=True)
            rown += 1
    _autosize_sheet(ws, max_width=70)


def build_backend_market_summary(datos):
    """Resumen serializable para backend/web."""
    out = []
    for d in datos or []:
        total_prov = 0.0
        total_market = 0.0
        hallazgos = []
        for clave, c in (d.get('conceptos') or {}).items():
            pt = _concept_total_amount(c)
            mt = _concept_market_total_amount(c)
            if isinstance(pt, (int,float)): total_prov += pt
            if isinstance(mt, (int,float)): total_market += mt
            for h in _concept_hallazgos_ejecutivos(c)[:6]:
                hallazgos.append({'concepto': clave, 'tipo': h.get('tipo'), 'severidad': h.get('severidad'), 'mensaje': h.get('mensaje')})
        diff = total_prov - total_market if total_market else None
        out.append({
            'proveedor': d.get('nombre'),
            'total_proveedor': round(total_prov, 2),
            'total_mercado': round(total_market, 2) if total_market else None,
            'diferencia_total': round(diff, 2) if isinstance(diff, (int,float)) else None,
            'diferencia_total_pct': round(diff / total_market, 4) if isinstance(diff, (int,float)) and total_market else None,
            'indirecto_mercado_pct': MARKET_INDIRECT_PCT,
            'hallazgos_ejecutivos': hallazgos[:50],
        })
    return out


def detect_valid_supplier_sheet(filepath: str):
    """
    Detecta la hoja válida de proveedor sin romper compatibilidad con backend/main.py.
    Devuelve: (sheet_name, conceptos_agrupados, reason)
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    best_sheet = None
    best_concepts = []
    best_reason = ""

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        candidates = [
            (_extract_matrix_apu_blocks(rows), "APU por bloques detectado"),
            (_extract_formato_apu(rows), "Formato CIOC / Clave detectado"),
            (_extract_formato_presupuesto(rows), "Presupuesto tabular detectado"),
            (_extract_formato_headers(rows), "Headers tabulares detectados"),
        ]

        for conceptos, reason in candidates:
            if conceptos and len(conceptos) > len(best_concepts):
                best_sheet = sheet_name
                best_concepts = conceptos
                best_reason = reason

    if best_sheet and best_concepts:
        return best_sheet, _agrupar_duplicados(best_concepts), best_reason

    raise ValueError(
        f"{Path(filepath).name}: el archivo no contiene una matriz de precios unitarios válida. "
        "Debes subir la matriz completa (APU), no solo el resumen de conceptos."
    )


# -----------------------------------------------------------------------------
# Paso 06 - Fix MO Supervisor de Seguridad + trazabilidad de homologacion
# -----------------------------------------------------------------------------

LABOR_ROLE_TOKENS = {
    'supervisor', 'segurista', 'residente', 'cabo', 'oficial', 'albanil', 'ayudante',
    'peon', 'operador', 'soldador', 'pintor', 'fierrero', 'carpintero', 'electricista',
    'plomero', 'maestro', 'maniobrista', 'topografo', 'chofer'
}

EQUIPMENT_SECURITY_TOKENS = {
    'equipo', 'epp', 'proteccion', 'casco', 'lentes', 'guantes', 'arnes', 'chaleco',
    'botas', 'seguridad', 'herramienta', 'andamio'
}


def _item_text_for_domain(item):
    return _normalize_text(' '.join(str(item.get(k) or '') for k in ['codigo', 'clave', 'descripcion', 'desc']))


def _item_has_any_token(item, tokens):
    txt = _item_text_for_domain(item)
    return any(re.search(r'\b' + re.escape(tok) + r'\b', txt) for tok in tokens)


def _is_supervisor_de_seguridad(item):
    txt = _item_text_for_domain(item)
    return bool(re.search(r'\bsupervisor\b', txt) and re.search(r'\bseguridad\b', txt))


def _is_equipo_de_seguridad(item):
    txt = _item_text_for_domain(item)
    code = str(item.get('codigo') or item.get('clave') or '').upper()
    return (
        bool(re.search(r'\bequipo\b', txt) and re.search(r'\bseguridad\b', txt))
        or bool(re.search(r'\bequipo\b', txt) and re.search(r'\bproteccion\b', txt))
        or code.startswith('%MO5')
    )


def _component_domain(item):
    kind = item.get('tipo_kind')
    if kind:
        return kind
    tipo = _material_kind_from_tipo(item.get('tipo') or item.get('categoria') or '')
    if tipo:
        return tipo
    if _is_supervisor_de_seguridad(item) or _item_has_any_token(item, LABOR_ROLE_TOKENS):
        return 'mano_obra'
    if _is_equipo_de_seguridad(item) or _item_has_any_token(item, EQUIPMENT_SECURITY_TOKENS):
        return 'equipo'
    return None


def _component_domains_compatible(provider_item, cd_item):
    """Evita falsos positivos por tokens genericos como 'seguridad'.

    Caso critico corregido:
      proveedor MO-SUP-SEG / SUPERVISOR DE SEGURIDAD (mano de obra)
      NO debe empatar con %MO5 / EQUIPO DE SEGURIDAD (equipo/herramienta).
    """
    p_domain = _component_domain(provider_item)
    c_domain = _component_domain(cd_item)
    if _is_supervisor_de_seguridad(provider_item) and _is_equipo_de_seguridad(cd_item):
        return False, 'MO supervisor de seguridad no es equivalente a equipo/EPP de seguridad'
    if p_domain and c_domain and p_domain != c_domain:
        # Compatibilidad estricta para matriz APU; si los dominios son distintos,
        # se reporta como extra/faltante en lugar de forzar un match por texto.
        return False, f'Tipo incompatible: proveedor={p_domain}, construdata={c_domain}'
    return True, ''


def _component_match_score(provider_item, cd_item):
    compatible, _reason = _component_domains_compatible(provider_item, cd_item)
    if not compatible:
        return 0.0
    p_desc = provider_item.get('descripcion') or provider_item.get('desc') or ''
    c_desc = cd_item.get('descripcion') or cd_item.get('desc') or ''
    p_tokens = _concept_tokens(p_desc)
    c_tokens = _concept_tokens(c_desc)
    score = 0.0
    score += 0.54 * _token_overlap_score(p_tokens, c_tokens)
    if normalize_unit(provider_item.get('unidad')) and normalize_unit(provider_item.get('unidad')) == normalize_unit(cd_item.get('unidad')):
        score += 0.16
    if _component_domain(provider_item) and _component_domain(provider_item) == _component_domain(cd_item):
        score += 0.22
    pc = _normalize_text(provider_item.get('codigo'))
    cc = _normalize_text(cd_item.get('codigo') or cd_item.get('clave'))
    if pc and cc and pc == cc:
        score = max(score, 0.98)
    # Penaliza matches basados solamente en 'seguridad' sin rol/equipo equivalente.
    shared = set(p_tokens) & set(c_tokens)
    if shared == {'seguridad'}:
        score = min(score, 0.18)
    return round(min(score, 1.0), 4)


def _match_matrix_component(provider_item, cd_items, used_indexes=None, min_score=0.30):
    used_indexes = used_indexes or set()
    best_idx = None
    best = None
    best_score = 0.0
    best_reject_reason = ''
    for idx, cd_item in enumerate(cd_items or []):
        if idx in used_indexes:
            continue
        compatible, reject_reason = _component_domains_compatible(provider_item, cd_item)
        if not compatible:
            if not best_reject_reason:
                best_reject_reason = reject_reason
            continue
        sc = _component_match_score(provider_item, cd_item)
        if sc > best_score:
            best_score = sc
            best_idx = idx
            best = cd_item
    if best and best_score >= min_score:
        return best_idx, best, best_score
    # Se conserva el score pero no se asigna match. El detalle se vera como extra/faltante.
    return None, None, best_score


def _candidate_explanation(provider_concept, candidate):
    cd = candidate.get('_concept') or {}
    p_desc = provider_concept.get('desc') or provider_concept.get('descripcion') or ''
    p_tokens = set(_concept_match_tokens(p_desc))
    c_tokens = set(cd.get('_match_tokens') or _concept_match_tokens(cd.get('desc_larga') or cd.get('desc') or ''))
    shared_concept = sorted(p_tokens & c_tokens)
    p_matrix = set(_matrix_summary_tokens(_provider_matrix_items(provider_concept), limit=40))
    c_matrix = set(cd.get('_matrix_tokens') or _matrix_summary_tokens(cd.get('matriz') or [], limit=40))
    shared_matrix = sorted(p_matrix & c_matrix)
    reasons = []
    if shared_concept:
        reasons.append('concepto: ' + ', '.join(shared_concept[:10]))
    if normalize_unit(provider_concept.get('unidad')) and normalize_unit(provider_concept.get('unidad')) == normalize_unit(cd.get('unidad')):
        reasons.append('unidad compatible: ' + str(provider_concept.get('unidad')))
    if shared_matrix:
        reasons.append('matriz: ' + ', '.join(shared_matrix[:10]))
    if isinstance(provider_concept.get('pu'), (int,float)) and isinstance(cd.get('pu_mercado'), (int,float)):
        reasons.append('PU proveedor/CD: %.2f / %.2f' % (provider_concept.get('pu'), cd.get('pu_mercado')))
    return ' | '.join(reasons)[:700]


def _rank_construdata_candidates(provider_concept, construdata_concepts, top_k=25, concept_index=None):
    ranked = []
    p_unit = normalize_unit(provider_concept.get('unidad'))
    p_desc = provider_concept.get('desc') or ''
    provider_concept['_match_tokens'] = provider_concept.get('_match_tokens') or _concept_match_tokens(p_desc)
    provider_concept['_matrix_tokens'] = provider_concept.get('_matrix_tokens') or _matrix_summary_tokens(_provider_matrix_items(provider_concept), limit=30)
    provider_concept['_core_text'] = provider_concept.get('_core_text') or _concept_core_text(p_desc)
    p_tokens = set(provider_concept['_match_tokens'])
    candidate_codes = _candidate_codes_for_provider(provider_concept, construdata_concepts, concept_index)
    for code in candidate_codes:
        cd = construdata_concepts.get(code)
        if not cd or str(code).startswith('__'):
            continue
        c_unit = normalize_unit(cd.get('unidad'))
        c_tokens = set(cd.get('_match_tokens') or _concept_match_tokens(cd.get('desc_larga') or cd.get('desc') or ''))
        if p_unit and c_unit and p_unit != c_unit and not (p_tokens & c_tokens):
            continue
        score, shared = _concept_candidate_score(provider_concept, cd)
        if score <= 0:
            continue
        cand = {
            'codigo': cd.get('codigo') or cd.get('clave') or code,
            'descripcion': cd.get('desc_larga') or cd.get('desc'),
            'unidad': cd.get('unidad'),
            'pu': cd.get('pu_mercado'),
            'score': score,
            'tokens_match': ', '.join(shared),
            'matriz_resumen': ', '.join(_matrix_summary_tokens(cd.get('matriz') or [], limit=12)),
            'embedding_text': cd.get('embedding_text') or _concept_embedding_text(cd),
            '_concept': cd,
        }
        cand['match_explanation'] = _candidate_explanation(provider_concept, cand)
        ranked.append(cand)
    ranked.sort(key=lambda x: x['score'], reverse=True)
    query = _provider_embedding_text(provider_concept)
    reranked = voyage_rerank_concept_candidates(query, ranked[:40], top_k=top_k)
    by_code = {c['codigo']: c for c in ranked}
    out = []
    for c in reranked:
        base = by_code.get(c.get('codigo'), c)
        merged = dict(base)
        merged.update({k:v for k,v in c.items() if k not in {'_concept'}})
        merged['_concept'] = base.get('_concept')
        if not merged.get('match_explanation'):
            merged['match_explanation'] = _candidate_explanation(provider_concept, merged)
        out.append(merged)
    return out[:top_k]


def compare_provider_matrix_vs_construdata(provider_key, provider_concept, construdata_match):
    result = {
        'provider_key': provider_key,
        'match_status': (construdata_match or {}).get('status'),
        'match_type': (construdata_match or {}).get('tipo_match'),
        'match_score': (construdata_match or {}).get('score'),
        'match_reason': (construdata_match or {}).get('reason'),
        'requires_review': (construdata_match or {}).get('requires_review'),
        'construdata_codigo': None,
        'construdata_desc': None,
        'construdata_pu': None,
        'provider_pu': provider_concept.get('pu'),
        'delta_pu': None,
        'summary': {},
        'rows': [],
        'findings': [],
        'candidates': (construdata_match or {}).get('candidates') or [],
    }
    codes, descs, pus, cd_items = _selected_cd_items_from_match(construdata_match)
    if not cd_items and not codes:
        result['findings'].append({'severity': 'alta', 'type': 'concepto_sin_match', 'message': 'No se encontro concepto equivalente en Construdata'})
        return result
    result['construdata_codigo'] = ' + '.join(codes)
    result['construdata_desc'] = ' + '.join(descs[:4])
    result['construdata_pu'] = round(sum(pus), 6) if pus else None
    result['delta_pu'] = _safe_ratio_delta(provider_concept.get('pu'), result['construdata_pu'])
    provider_items = _provider_matrix_items(provider_concept)
    used_cd = set()
    matched = missing = extra = 0
    amount_delta = 0.0
    for p_item in provider_items:
        idx, cd_item, score = _match_matrix_component(p_item, cd_items, used_cd, min_score=0.30)
        p_cant = p_item.get('factor') if p_item.get('factor') is not None else p_item.get('cantidad')
        p_precio = p_item.get('precio_base') if p_item.get('precio_base') is not None else p_item.get('precio')
        row = {
            'tipo': p_item.get('tipo_kind'),
            'provider_codigo': p_item.get('codigo'),
            'provider_desc': p_item.get('descripcion'),
            'provider_unidad': p_item.get('unidad'),
            'provider_cantidad': p_cant,
            'provider_precio': p_precio,
            'provider_importe': p_item.get('importe'),
            'cd_codigo': None, 'cd_desc': None, 'cd_unidad': None, 'cd_cantidad': None, 'cd_precio': None, 'cd_importe': None,
            'cd_source_concept': None,
            'formula_kind': None,
            'match_score': score,
            'delta_cantidad': None, 'delta_precio': None, 'delta_importe': None,
            'estado': 'sin_match_insumo',
            'motivo_no_match': None,
        }
        if cd_item:
            used_cd.add(idx); matched += 1
            row.update({
                'cd_codigo': cd_item.get('codigo') or cd_item.get('clave'),
                'cd_desc': cd_item.get('descripcion'),
                'cd_unidad': cd_item.get('unidad'),
                'cd_cantidad': cd_item.get('cantidad'),
                'cd_precio': cd_item.get('precio'),
                'cd_importe': cd_item.get('importe'),
                'cd_source_concept': cd_item.get('source_concept_code'),
                'formula_kind': cd_item.get('formula_kind'),
                'estado': 'match_insumo',
            })
            row['delta_cantidad'] = _safe_ratio_delta(row['provider_cantidad'], row['cd_cantidad'])
            if _component_is_percentage(cd_item):
                row['delta_precio'] = None
            else:
                row['delta_precio'] = _safe_ratio_delta(row['provider_precio'], row['cd_precio'])
            row['delta_importe'] = _safe_ratio_delta(row['provider_importe'], row['cd_importe'])
            if isinstance(row['provider_importe'], (int, float)) and isinstance(row['cd_importe'], (int, float)):
                amount_delta += float(row['provider_importe']) - float(row['cd_importe'])
        else:
            extra += 1
            # explica el caso mas frecuente sin generar falsos matches.
            for possible in cd_items:
                compatible, reason = _component_domains_compatible(p_item, possible)
                if not compatible and reason:
                    row['motivo_no_match'] = reason
                    break
        result['rows'].append(row)
    for idx, cd_item in enumerate(cd_items):
        if idx in used_cd:
            continue
        missing += 1
        result['rows'].append({
            'tipo': cd_item.get('tipo_kind'), 'provider_codigo': None, 'provider_desc': None, 'provider_unidad': None,
            'provider_cantidad': None, 'provider_precio': None, 'provider_importe': None,
            'cd_codigo': cd_item.get('codigo') or cd_item.get('clave'), 'cd_desc': cd_item.get('descripcion'), 'cd_unidad': cd_item.get('unidad'),
            'cd_cantidad': cd_item.get('cantidad'), 'cd_precio': cd_item.get('precio'), 'cd_importe': cd_item.get('importe'),
            'cd_source_concept': cd_item.get('source_concept_code'), 'formula_kind': cd_item.get('formula_kind'),
            'match_score': None, 'delta_cantidad': None, 'delta_precio': None, 'delta_importe': None, 'estado': 'faltante_en_proveedor',
            'motivo_no_match': None,
        })
    result['summary'] = {
        'proveedor_renglones': len(provider_items), 'construdata_renglones': len(cd_items), 'insumos_matcheados': matched,
        'insumos_extra_proveedor': extra, 'insumos_faltantes_proveedor': missing, 'diferencia_importe_matriz': round(amount_delta, 4),
        'conceptos_cd_asociados': len(codes), 'requiere_revision': bool(result.get('requires_review')),
    }
    if result['match_type'] == 'match_compuesto':
        result['findings'].append({'severity': 'media', 'type': 'concepto_compuesto', 'message': f"Concepto proveedor homologado como compuesto contra {len(codes)} conceptos Construdata; requiere validacion tecnica."})
    if result['delta_pu'] is not None and abs(result['delta_pu']) >= 0.15:
        result['findings'].append({'severity': 'alta' if abs(result['delta_pu']) >= 0.30 else 'media', 'type': 'delta_pu', 'message': f"PU proveedor vs Construdata: {result['delta_pu']*100:.1f}%"})
    if missing:
        result['findings'].append({'severity': 'alta', 'type': 'insumos_faltantes', 'message': f'Faltan {missing} insumos de la matriz Construdata en proveedor'})
    if extra:
        result['findings'].append({'severity': 'media', 'type': 'insumos_extra', 'message': f'Proveedor incluye {extra} insumos no homologados contra Construdata'})
    for row in result['rows']:
        if row.get('estado') != 'match_insumo':
            continue
        if row.get('formula_kind') == 'porcentaje_mo' and isinstance(row.get('delta_cantidad'), (int, float)) and abs(row['delta_cantidad']) >= 0.20:
            result['findings'].append({'severity': 'media', 'type': 'porcentaje_mo', 'message': f"Porcentaje MO diferente en {row.get('provider_desc')}: {row['delta_cantidad']*100:.1f}%"})
        elif isinstance(row.get('delta_cantidad'), (int, float)) and abs(row['delta_cantidad']) >= 0.25:
            result['findings'].append({'severity': 'alta', 'type': 'rendimiento', 'message': f"Rendimiento/cantidad diferente en {row.get('provider_desc')}: {row['delta_cantidad']*100:.1f}%"})
        if isinstance(row.get('delta_precio'), (int, float)) and abs(row['delta_precio']) >= 0.20:
            result['findings'].append({'severity': 'media', 'type': 'precio_insumo', 'message': f"Precio diferente en {row.get('provider_desc')}: {row['delta_precio']*100:.1f}%"})
    return result


def _build_homologacion_ia_sheet(wb, datos):
    ws = wb.create_sheet('Homologacion IA')
    headers = ['Proveedor','Concepto proveedor','Descripcion proveedor','Unidad','PU proveedor','Tipo match','Estado','Score','Requiere revision','Codigos CD seleccionados','Motivo','Por que se considero','Top candidatos con razones']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            m = c.get('construdata_concept_match') or {}
            selected = m.get('selected_matches') or []
            codes = ', '.join((sm.get('concept') or {}).get('codigo') or (sm.get('concept') or {}).get('clave') or '' for sm in selected)
            cand = m.get('candidates') or []
            selected_reason = ''
            selected_code = codes.split(',')[0].strip() if codes else ''
            for x in cand:
                if str(x.get('codigo')) == selected_code:
                    selected_reason = x.get('match_explanation') or ''
                    break
            cand_txt = '; '.join(f"{x.get('codigo')} ({x.get('score')}) {str(x.get('descripcion') or '')[:65]} :: {str(x.get('match_explanation') or '')[:180]}" for x in cand[:5])
            vals = [d.get('nombre'), clave, c.get('desc'), c.get('unidad'), c.get('pu'), m.get('tipo_match'), m.get('status'), m.get('score'), bool(m.get('requires_review')), codes, m.get('reason'), selected_reason, cand_txt]
            for col,val in enumerate(vals, start=1):
                cell=ws.cell(row=rown, column=col, value=val)
                cell.border=_brd(); cell.font=_fnt(size=8)
                if col == 5 and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
                if col in [3,11,12,13]: cell.alignment=_aln('left','center', wrap=True)
            rown += 1
    _autosize_sheet(ws, max_width=65)


def _build_construdata_matrix_detail_sheet(wb, datos):
    ws = wb.create_sheet('Construdata matrices')
    headers = ['Proveedor','Concepto proveedor','Codigo CD','Estado','Tipo','Insumo proveedor','Unidad prov.','Cant. prov.','Precio prov.','Importe prov.','Insumo CD','Unidad CD','Cant. CD','Precio CD','Importe CD','Score','Delta cant.','Delta precio','Delta importe','Motivo no match']
    _sheet_headers(ws, headers)
    rown = 2
    for d in datos:
        for clave, c in d.get('conceptos', {}).items():
            comp = c.get('construdata_matrix_comparison') or {}
            for r in comp.get('rows') or []:
                vals = [
                    d.get('nombre'), clave, comp.get('construdata_codigo'), r.get('estado'), r.get('tipo'), r.get('provider_desc'), r.get('provider_unidad'), r.get('provider_cantidad'), r.get('provider_precio'), r.get('provider_importe'),
                    r.get('cd_desc'), r.get('cd_unidad'), r.get('cd_cantidad'), r.get('cd_precio'), r.get('cd_importe'), r.get('match_score'), r.get('delta_cantidad'), r.get('delta_precio'), r.get('delta_importe'), r.get('motivo_no_match')
                ]
                for col, val in enumerate(vals, start=1):
                    cell = ws.cell(row=rown, column=col, value=val)
                    cell.border = _brd(); cell.font = _fnt(size=8)
                    if col in [9,10,14,15] and isinstance(val, (int, float)): cell.number_format = '$#,##0.00'
                    if col in [17,18,19] and isinstance(val, (int, float)): cell.number_format = '+0.0%;-0.0%;0.0%'
                    if col in [6,11,20]: cell.alignment = _aln('left','center', wrap=True)
                rown += 1
    _autosize_sheet(ws, max_width=55)

# -----------------------------------------------------------------------------
# Paso 07b - Loader rapido para construdata_matrices.xlsx raw 24 columnas
# -----------------------------------------------------------------------------
# Evita que openpyxl tarde demasiado leyendo el XML grande de la consulta SQL.
# Este override solo aplica cuando el xlsx tiene xl/worksheets/sheet1.xml con
# las 24 columnas raw indicadas por el usuario.

import zipfile as _zipfile
import xml.etree.ElementTree as _ET


def _xlsx_col_to_index(cell_ref: str):
    letters = ''.join(ch for ch in str(cell_ref or '') if ch.isalpha()).upper()
    if not letters:
        return None
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - ord('A') + 1)
    return value - 1


def _xml_cell_value(cell):
    # Soporta inlineStr y v numericos/booleanos.
    t = cell.attrib.get('t')
    if t == 'inlineStr':
        texts = []
        for elem in cell.iter():
            if elem.tag.endswith('}t') or elem.tag == 't':
                if elem.text:
                    texts.append(elem.text)
        return ''.join(texts)
    v = None
    for child in cell:
        if child.tag.endswith('}v') or child.tag == 'v':
            v = child.text
            break
    if v is None:
        return None
    if t == 'b':
        return str(v).strip() in {'1', 'true', 'TRUE'}
    try:
        if isinstance(v, str) and ('.' in v or 'E' in v.upper()):
            return float(v)
        return int(v)
    except Exception:
        try:
            return float(v)
        except Exception:
            return v


def _iter_raw_construdata_rows_fast(filepath: str):
    with _zipfile.ZipFile(filepath) as zf:
        if 'xl/worksheets/sheet1.xml' not in zf.namelist():
            return
        with zf.open('xl/worksheets/sheet1.xml') as fh:
            context = _ET.iterparse(fh, events=('end',))
            for _event, elem in context:
                if not (elem.tag.endswith('}row') or elem.tag == 'row'):
                    continue
                row = [None] * 24
                for cell in list(elem):
                    if not (cell.tag.endswith('}c') or cell.tag == 'c'):
                        continue
                    idx = _xlsx_col_to_index(cell.attrib.get('r'))
                    if idx is None or idx >= 24:
                        continue
                    row[idx] = _xml_cell_value(cell)
                elem.clear()
                yield row


def _is_raw_construdata_query_xlsx_fast(filepath: str) -> bool:
    try:
        for row in _iter_raw_construdata_rows_fast(filepath):
            if not row:
                return False
            codigo = _clean_text(row[4] if len(row) > 4 else '')
            return bool(re.match(r'^\d{4,6}-\d{2,4}$', codigo))
    except Exception:
        return False
    return False


def _load_construdata_raw_query_workbook_fast(filepath: str) -> dict:
    idx = RAW_CONSTRUDATA_POSITIONAL_SCHEMA
    conceptos = {}
    materials_index = []
    row_count = 0
    for row in _iter_raw_construdata_rows_fast(filepath):
        row_count += 1
        codigo = _clean_text(_cell(row, idx['codigo_concepto']))
        if not codigo:
            continue
        desc = _clean_text(_cell(row, idx['concepto_desc']))
        desc_larga = _clean_text(_cell(row, idx['concepto_desc_larga'])) or desc
        unidad = _clean_text(_cell(row, idx['concepto_unidad']))
        pu = _as_float(_cell(row, idx['concepto_pu']))
        if codigo not in conceptos:
            conceptos[codigo] = {
                'clave': codigo, 'codigo': codigo, 'desc': desc[:300], 'desc_larga': desc_larga[:900],
                'unidad': unidad, 'pu_mercado': pu, 'tipo': 'CONCEPTO', 'matriz': [],
                'partida_codigo': codigo.split('-')[0] if '-' in codigo else codigo[:5],
            }
        ins_code = _clean_text(_cell(row, idx['codigo_insumo']))
        ins_desc = _clean_text(_cell(row, idx['insumo_desc_larga'])) or _clean_text(_cell(row, idx['insumo_desc']))
        if not ins_code and not ins_desc:
            continue
        tipo_raw = _clean_text(_cell(row, idx['tipo']))
        cantidad = _as_float(_cell(row, idx['cantidad']))
        precio = _as_float(_cell(row, idx['precio']))
        importe = _as_float(_cell(row, idx['importe']))
        dividir_val = _cell(row, idx['dividir'])
        dividir = bool(dividir_val)
        formula_kind = 'normal'
        code_up = ins_code.upper()
        expr = _clean_text(_cell(row, idx['expresion']))
        if code_up.startswith('%MO') or '%MO' in code_up or '%MO' in expr.upper():
            formula_kind = 'porcentaje_mo'
        elif dividir:
            formula_kind = 'rendimiento_inverso'
        renglon = {
            'renglon': _cell(row, idx['renglon']),
            'tipo': tipo_raw,
            'tipo_id': _cell(row, idx['tipo_id']),
            'tipo_kind': _material_kind_from_tipo(tipo_raw),
            'codigo': ins_code,
            'clave': ins_code,
            'descripcion': ins_desc[:700],
            'unidad': _clean_text(_cell(row, idx['insumo_unidad'])),
            'unidad_norm': normalize_unit(_cell(row, idx['insumo_unidad'])),
            'cantidad': cantidad,
            'precio': precio,
            'precio_base': precio,
            'importe': importe,
            'dividir': dividir,
            'expresion': expr,
            'formula_kind': formula_kind,
            'porcentaje': cantidad if formula_kind == 'porcentaje_mo' else None,
            'base_calculo': precio if formula_kind == 'porcentaje_mo' else None,
            'fuente': 'construdata_raw_query_fast',
        }
        conceptos[codigo]['matriz'].append(renglon)
    for concepto in conceptos.values():
        subtotal_mo = 0.0
        for r in concepto.get('matriz') or []:
            if r.get('tipo_kind') == 'mano_obra' and r.get('formula_kind') != 'porcentaje_mo' and isinstance(r.get('importe'), (int, float)):
                subtotal_mo += float(r['importe'])
        for r in concepto.get('matriz') or []:
            if r.get('formula_kind') == 'porcentaje_mo':
                pct = r.get('porcentaje')
                if isinstance(pct, (int, float)):
                    r['base_calculo'] = subtotal_mo or r.get('base_calculo')
                    r['importe'] = round((subtotal_mo or 0.0) * float(pct), 6)
                    r['precio'] = r['base_calculo']
                    r['precio_base'] = r['base_calculo']
        concepto['matriz_stats'] = _matrix_stats(concepto.get('matriz') or [])
        concepto['_match_tokens'] = _concept_match_tokens(concepto.get('desc_larga') or concepto.get('desc') or '')
        concepto['_matrix_tokens'] = _matrix_summary_tokens(concepto.get('matriz') or [], limit=30)
        concepto['_core_text'] = _concept_core_text(concepto.get('desc_larga') or concepto.get('desc') or '')
        concepto['embedding_text'] = _concept_embedding_text(concepto)
    conceptos['__materials_index__'] = materials_index
    conceptos['__construdata_matrices__'] = {k:v for k,v in conceptos.items() if not str(k).startswith('__')}
    conceptos['__benchmark_kind__'] = 'construdata_matrices_raw_query_fast'
    conceptos['__raw_query_schema__'] = RAW_CONSTRUDATA_POSITIONAL_SCHEMA
    conceptos['__raw_rows_loaded__'] = row_count
    return conceptos


_load_construdata_matrices_step07_previous = load_construdata_matrices

def load_construdata_matrices(filepath: str) -> dict:
    if _is_raw_construdata_query_xlsx_fast(filepath):
        return _load_construdata_raw_query_workbook_fast(filepath)
    return _load_construdata_matrices_step07_previous(filepath)


_extract_precios_nacional_step07_previous = extract_precios_nacional

def extract_precios_nacional(filepath: str) -> dict:
    src = Path(filepath)
    cache_key = None
    try:
        stat = src.stat()
        cache_key = (str(src.resolve()), stat.st_mtime_ns, stat.st_size)
        cached = _BENCHMARK_CACHE.get(cache_key)
        if cached is not None:
            return dict(cached)
    except Exception:
        cache_key = None
    if _is_raw_construdata_query_xlsx_fast(filepath):
        matrices = load_construdata_matrices(filepath)
        if cache_key is not None:
            _BENCHMARK_CACHE.clear()
            _BENCHMARK_CACHE[cache_key] = dict(matrices)
        return matrices
    return _extract_precios_nacional_step07_previous(filepath)


# -----------------------------------------------------------------------------
# Paso 08 - Cambio de arquitectura: mercado granular por insumo declarado
# -----------------------------------------------------------------------------
# El mercado ya no se calcula primero por concepto -> matriz Construdata.
# Ahora el flujo principal toma la matriz del contratista y homologa cada elemento
# declarado contra los catálogos granulares Construdata de materiales, mano de
# obra y maquinaria/equipo. La matriz Construdata por concepto queda como tab 3
# de validación técnica: faltantes, extras y concepto compuesto.

STEP08_CATALOG_FILES = {
    'material': 'construdata-materiales-052026.xlsx',
    'mano_obra': 'construdata-manodeobra-052026.xlsx',
    'equipo': 'construdata-maquinaria-052026.xlsx',
}

_STEP08_CATALOG_CACHE = None
_STEP08_CATALOG_MTIME = None


def _step08_data_dir():
    return _quantia_data_dir()


def _step08_norm_code(value):
    """Normaliza codigos de insumo para matching exacto.

    Construdata suele traer codigos con puntos/sufijos de captura
    (ej. MO-SUP-SEG.) mientras el proveedor puede traer MO-SUP-SEG.
    Para comparacion de catalogo, la puntuacion no debe impedir el match
    exacto. Conservamos solo letras y numeros.
    """
    return re.sub(r'[^A-Z0-9]+', '', _normalize_text(value or '').upper())


def _step08_safe_float(value):
    return _as_float(value)


def _step08_read_catalog_file(path, domain):
    """Lee un catálogo Construdata granular y devuelve registros normalizados."""
    records = []
    if not path.exists():
        return records
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        wb.close()
        return records
    headers = [_normalize_text(x) for x in rows[0]]

    def idx(*names):
        targets = [_normalize_text(n) for n in names]
        for i, h in enumerate(headers):
            if not h:
                continue
            for t in targets:
                if h == t or t in h:
                    return i
        return None

    if domain == 'material':
        code_i = idx('codigo', 'código') or 1
        desc_i = idx('descripcion completa', 'descripción completa', 'descripcion') or 2
        unit_i = idx('unidad') or 4
        type_i = idx('tipo') or 3
        cost_i = idx('costo') or 11
        cost_label = 'Costo'
    elif domain == 'equipo':
        code_i = idx('codigo', 'código') or 2
        desc_i = idx('descripcion completa', 'descripción completa', 'descripcion') or 3
        unit_i = idx('unidad') or 4
        type_i = idx('tipo') or 5
        cost_i = idx('costo') or 6
        cost_label = 'Costo'
    else:
        code_i = idx('codigo', 'código') or 1
        desc_i = idx('descripcion completa', 'descripción completa', 'descripcion') or 2
        unit_i = idx('unidad') or 3
        type_i = idx('tipo') or 4
        # Para mano de obra el costo comparable es Salario Real.
        cost_i = idx('salario real')
        if cost_i is None:
            cost_i = idx('costo', 'salario integrado') or 9
        cost_label = headers[cost_i] if cost_i is not None and cost_i < len(headers) else 'Salario Real'

    for rnum, row in enumerate(rows[1:], start=2):
        code = _clean_text(row[code_i] if code_i is not None and code_i < len(row) else '')
        desc = _clean_text(row[desc_i] if desc_i is not None and desc_i < len(row) else '')
        if not code and not desc:
            continue
        unit = _clean_text(row[unit_i] if unit_i is not None and unit_i < len(row) else '')
        idtipo = row[type_i] if type_i is not None and type_i < len(row) else None
        cost = _step08_safe_float(row[cost_i] if cost_i is not None and cost_i < len(row) else None)
        tokens = _tokenize_text(f'{code} {desc}')
        records.append({
            'domain': domain,
            'codigo': code,
            'codigo_norm': _step08_norm_code(code),
            'descripcion': desc,
            'descripcion_larga': desc,
            'unidad': unit,
            'unidad_norm': normalize_unit(unit),
            'id_tipo': idtipo,
            'tipo_insumo': {'material': 'MATERIALES', 'mano_obra': 'MANO DE OBRA', 'equipo': 'EQUIPO Y HERRAMIENTA'}.get(domain, domain),
            'costo': cost,
            'cost_label': cost_label,
            'tokens': tokens,
            'norm': _normalize_text(f'{code} {desc}'),
            'source_file': path.name,
            'source_row': rnum,
        })
    wb.close()
    return records


def load_step08_market_catalogs(force=False):
    """Carga los tres catálogos granulares Construdata desde /data."""
    global _STEP08_CATALOG_CACHE, _STEP08_CATALOG_MTIME
    data_dir = _step08_data_dir()
    files = {k: data_dir / v for k, v in STEP08_CATALOG_FILES.items()}
    mtime = tuple((str(p), p.stat().st_mtime if p.exists() else None) for p in files.values())
    if _STEP08_CATALOG_CACHE is not None and _STEP08_CATALOG_MTIME == mtime and not force:
        return _STEP08_CATALOG_CACHE

    by_domain = {'material': [], 'mano_obra': [], 'equipo': []}
    for domain, path in files.items():
        by_domain[domain] = _step08_read_catalog_file(path, domain)

    by_code = defaultdict(list)
    for domain, rows in by_domain.items():
        for rec in rows:
            if rec.get('codigo_norm'):
                by_code[(domain, rec['codigo_norm'])].append(rec)

    _STEP08_CATALOG_CACHE = {'by_domain': by_domain, 'by_code': by_code, 'files': {k: str(v) for k, v in files.items()}}
    _STEP08_CATALOG_MTIME = mtime
    return _STEP08_CATALOG_CACHE


def _step08_domain_from_component_key(kind):
    if kind in ('materiales', 'material'):
        return 'material'
    if kind in ('mano_obra', 'mano obra'):
        return 'mano_obra'
    if kind in ('equipo', 'equipo_herramienta', 'maquinaria'):
        return 'equipo'
    return None


def _step08_iter_provider_items(concepto):
    mapping = [
        ('MATERIALES', 'material', 'materiales_items'),
        ('MANO DE OBRA', 'mano_obra', 'mano_obra_items'),
        ('EQUIPO Y HERRAMIENTA', 'equipo', 'equipo_items'),
        ('BASICOS', 'material', 'basicos_items'),
    ]
    for section_label, domain, key in mapping:
        for idx, item in enumerate(concepto.get(key) or [], start=1):
            r = dict(item)
            r['_section'] = section_label
            r['_domain'] = domain
            r['_component_key'] = key
            r['_idx'] = idx
            yield r


def _step08_item_tokens(item):
    return _tokenize_text(f"{item.get('codigo','')} {item.get('descripcion','')}")


def _step08_score_catalog_candidate(item, cand):
    p_tokens = set(_step08_item_tokens(item))
    c_tokens = set(cand.get('tokens') or [])
    if not p_tokens or not c_tokens:
        text_score = 0.0
    else:
        text_score = len(p_tokens & c_tokens) / max(len(p_tokens | c_tokens), 1)
    score = 0.62 * text_score
    if normalize_unit(item.get('unidad')) and normalize_unit(item.get('unidad')) == cand.get('unidad_norm'):
        score += 0.18
    pc = _step08_norm_code(item.get('codigo'))
    if pc and pc == cand.get('codigo_norm'):
        score = max(score, 0.98)
    # Penaliza confusiones supervisor de seguridad (MO) vs equipo de seguridad.
    item_desc = _normalize_text(item.get('descripcion'))
    cand_desc = _normalize_text(cand.get('descripcion'))
    if item.get('_domain') == 'mano_obra' and 'supervisor' in item_desc and 'equipo' in cand_desc:
        score -= 0.50
    if item.get('_domain') == 'equipo' and item.get('codigo') and str(item.get('codigo')).startswith('%') and cand.get('codigo_norm') == _step08_norm_code(item.get('codigo')):
        score = max(score, 0.99)
    return round(max(0.0, min(1.0, score)), 4)


def _step08_match_market_item(item, catalogs):
    domain = item.get('_domain') or 'material'
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    candidates = []
    if code_norm and (domain, code_norm) in by_code:
        candidates = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)]]
    else:
        domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
        item_tokens = set(_step08_item_tokens(item))
        prelim = []
        for cand in domain_rows:
            ct = set(cand.get('tokens') or [])
            # filtro barato: al menos un token común o unidad/codigo parecido
            if not (item_tokens & ct):
                continue
            sc = _step08_score_catalog_candidate(item, cand)
            if sc >= 0.18:
                c = dict(cand)
                c['score'] = sc
                c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
                prelim.append(c)
        prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
        candidates = prelim[:8]
    selected = candidates[0] if candidates and (candidates[0].get('score') or 0) >= 0.34 else None
    return {
        'selected': selected,
        'candidates': candidates,
        'status': 'match' if selected else ('candidatos_baja_confianza' if candidates else 'sin_match'),
        'confidence': selected.get('score') if selected else (candidates[0].get('score') if candidates else 0.0),
    }


def _step08_is_percent_item(item):
    code = str(item.get('codigo') or '').strip().upper()
    unit = normalize_unit(item.get('unidad'))
    desc = _normalize_text(item.get('descripcion'))
    return code.startswith('%') or unit == '%' or 'porcentaje' in desc


def _step08_market_importe_for_item(item, match):
    selected = (match or {}).get('selected') or {}
    price = selected.get('costo')
    qty = item.get('factor') if isinstance(item.get('factor'), (int, float)) else item.get('cantidad')
    provider_price = item.get('precio_base')
    provider_importe = item.get('importe')
    if _step08_is_percent_item(item):
        # Los porcentuales dependen de una base de cálculo dentro de la matriz.
        # En el catálogo granular se conserva el match/código, pero para no romper
        # el total directo se usa el importe declarado como proxy y se etiqueta.
        return provider_importe if isinstance(provider_importe, (int, float)) else None
    if isinstance(price, (int, float)) and isinstance(qty, (int, float)):
        return float(price) * float(qty)
    return None


def _step08_apply_market_catalog_pricing(conceptos):
    """Enriquece cada concepto con precio de mercado por cada insumo declarado."""
    try:
        catalogs = load_step08_market_catalogs()
    except Exception as exc:
        for c in (conceptos or {}).values():
            c['granular_market_items'] = []
            c['granular_market_error'] = str(exc)
        return conceptos

    for _, c in (conceptos or {}).items():
        enriched = []
        direct_market = 0.0
        provider_direct = 0.0
        matched_count = 0
        total_count = 0
        for item in _step08_iter_provider_items(c):
            total_count += 1
            match = _step08_match_market_item(item, catalogs)
            selected = match.get('selected') or {}
            market_price = selected.get('costo')
            market_importe = _step08_market_importe_for_item(item, match)
            provider_price = item.get('precio_base')
            provider_importe = item.get('importe')
            delta_price = None
            if isinstance(provider_price, (int, float)) and isinstance(market_price, (int, float)) and market_price:
                if not _step08_is_percent_item(item):
                    delta_price = (float(provider_price) - float(market_price)) / float(market_price)
            if isinstance(market_importe, (int, float)):
                direct_market += float(market_importe)
            if isinstance(provider_importe, (int, float)):
                provider_direct += float(provider_importe)
            if selected:
                matched_count += 1
            enriched.append({
                'section': item.get('_section'),
                'domain': item.get('_domain'),
                'codigo': item.get('codigo'),
                'descripcion': item.get('descripcion'),
                'unidad': item.get('unidad'),
                'cantidad': item.get('factor'),
                'precio_proveedor': provider_price,
                'importe_proveedor': provider_importe,
                'codigo_mercado': selected.get('codigo'),
                'descripcion_mercado': selected.get('descripcion'),
                'unidad_mercado': selected.get('unidad'),
                'precio_mercado': market_price,
                'importe_mercado': market_importe,
                'delta_precio_pct': delta_price,
                'match_status': match.get('status'),
                'confidence': match.get('confidence'),
                'match_reason': selected.get('match_reason') if selected else '',
                'source_file': selected.get('source_file'),
                'source_row': selected.get('source_row'),
                'is_percent_item': _step08_is_percent_item(item),
                'top_candidates': match.get('candidates') or [],
            })
        c['granular_market_items'] = enriched
        c['granular_market_direct'] = round(direct_market, 6) if direct_market else None
        c['granular_provider_direct'] = round(provider_direct, 6) if provider_direct else None
        c['granular_market_match_coverage'] = (matched_count / total_count) if total_count else None
    return conceptos


# Override Paso 07: el costo directo de mercado principal ahora sale del catálogo granular.
def _market_component_total(concepto):
    direct = concepto.get('granular_market_direct')
    if isinstance(direct, (int, float)) and direct > 0:
        return float(direct)
    comp = concepto.get('construdata_matrix_comparison') or {}
    cd_pu = comp.get('construdata_pu')
    if isinstance(cd_pu, (int, float)) and cd_pu > 0:
        return float(cd_pu)
    return 0.0


def _step08_market_status(delta):
    if delta is None:
        return 'SIN_REFERENCIA'
    if abs(delta) <= 0.05:
        return 'OK'
    if delta > 0.20:
        return 'SOBRE_MERCADO_ALTO'
    if delta > 0:
        return 'SOBRE_MERCADO'
    if delta < -0.20:
        return 'BAJO_MERCADO_REVISAR'
    return 'BAJO_MERCADO'


def _step08_row_fill_by_delta(delta):
    if delta is None:
        return None
    if delta > 0.20:
        return _fill(C_ROJO_MAL)
    if delta > 0.05:
        return _fill(C_AMARILLO)
    if delta < -0.20:
        return _fill('D9EAD3')
    return None


def _step08_build_matrix_base_sheet(wb, dato):
    ws = wb.create_sheet('Matriz base CD')
    ws.sheet_view.showGridLines = False
    headers = [
        'Servicio proveedor','Descripción proveedor','Tipo match','Score','Código CD','Concepto CD','Tipo elemento CD',
        'Código insumo CD','Insumo CD','Unidad','Cantidad','Costo/Base','Importe','Estado vs proveedor','Nota'
    ]
    _sheet_headers(ws, headers)
    row = 2
    for clave, c in (dato.get('conceptos') or {}).items():
        comp = c.get('construdata_matrix_comparison') or {}
        rows = comp.get('rows') or []
        if not rows:
            vals = [clave, c.get('desc'), comp.get('match_type') or 'SIN_MATCH', comp.get('score'), comp.get('construdata_codigo'), comp.get('construdata_desc'), '', '', '', '', '', '', '', 'SIN_MATRIZ_BASE', 'No hay matriz Construdata con confianza suficiente; usar tab Detalle para evaluación granular.']
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row, col, val); cell.border = _brd(); cell.font = _fnt(size=8)
                if col in (2,15): cell.alignment = _aln('left','center', wrap=True)
            row += 1
            continue
        for r in rows:
            estado = r.get('estado')
            nota = ''
            if estado == 'faltante_en_proveedor':
                nota = 'Elemento recomendado por la matriz base Construdata y no localizado en la matriz del contratista.'
            elif estado == 'sin_match_insumo':
                nota = 'Elemento declarado por contratista fuera de la matriz base; puede ser alcance adicional o concepto compuesto.'
            elif estado == 'match':
                nota = 'Elemento homologado contra matriz base.'
            vals = [
                clave, c.get('desc'), comp.get('match_type'), comp.get('score'), comp.get('construdata_codigo'), comp.get('construdata_desc'),
                r.get('tipo'), r.get('cd_codigo') or r.get('provider_codigo'), r.get('cd_desc') or r.get('provider_desc'),
                r.get('cd_unidad') or r.get('provider_unidad'), r.get('cd_cantidad') or r.get('provider_cantidad'),
                r.get('cd_precio') or r.get('provider_precio'), r.get('cd_importe') or r.get('provider_importe'), estado, nota
            ]
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row, col, val); cell.border = _brd(); cell.font = _fnt(size=8)
                if col in (2,6,9,15): cell.alignment = _aln('left','center', wrap=True)
                if col in (12,13) and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
                if col == 4 and isinstance(val, (int,float)): cell.number_format = '0.0%'
            if estado == 'faltante_en_proveedor':
                for cc in range(1, len(headers)+1): ws.cell(row, cc).fill = _fill(C_AMARILLO)
            elif estado == 'sin_match_insumo':
                for cc in range(1, len(headers)+1): ws.cell(row, cc).fill = _fill('FCE4D6')
            row += 1
    _autosize_sheet(ws, max_width=60)


# Override single-provider report for Paso 08: tabs limpios = Comparativa, Detalle, Matriz base CD.
def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comparativa'
    ws_det = wb.create_sheet('Detalle')

    conceptos_items = list((dato.get('conceptos') or {}).items())
    total_contratista = round(sum(v.get('total') or 0 for _, v in conceptos_items if isinstance(v.get('total'), (int, float))), 2)

    # ---------------- TAB 1: resumen todavía simple ----------------
    ws.sheet_view.showGridLines = False
    for col, width in {'A':12,'B':58,'C':10,'D':12,'E':13,'F':14,'G':14,'H':14,'I':12,'J':14,'K':14,'L':18}.items():
        ws.column_dimensions[col].width = width
    ws.merge_cells('A1:L1')
    ws['A1'] = f"COMPARATIVO DE COTIZACIÓN — {meta.get('proyecto','')}"
    ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10); ws['A1'].alignment = _aln('center','center')
    info = [('Cliente', meta.get('cliente','')), ('Proyecto', meta.get('proyecto','')), ('Archivo', Path(meta.get('archivo_fuente') or '').name)]
    for i, (lbl, val) in enumerate(info, start=2):
        ws.cell(i,1,lbl).font = _fnt(bold=True, size=8)
        ws.cell(i,2,val).font = _fnt(size=8)
    headers = ['Servicio','Descripción','Unidad','Cantidad','PU Contratista','Importe Contratista','PU Mercado +25%','Importe Mercado','Dif %','Cobertura mercado','Estado','Nota']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(6, col, h); cell.fill = _fill(C_AZUL_OSC); cell.font = _fnt(bold=True, color=C_BLANCO, size=8); cell.border = _brd(); cell.alignment = _aln('center','center', wrap=True)
        if col in (7,8,9):
            cell.fill = _fill(C_AMARILLO_PMD); cell.font = _fnt(bold=True, color='000000', size=8)
    row = 7
    total_market = 0.0
    for clave, c in conceptos_items:
        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'), (int,float)) and c.get('cantidad') else 1
        pu = c.get('pu') if isinstance(c.get('pu'), (int,float)) else None
        importe = _concept_total_amount(c)
        market_pu = _market_pu_with_indirect(c)
        market_total = (market_pu * cantidad) if isinstance(market_pu, (int,float)) else None
        if isinstance(market_total, (int,float)): total_market += market_total
        diff_pct = ((importe - market_total) / market_total) if isinstance(importe, (int,float)) and isinstance(market_total, (int,float)) and market_total else None
        coverage = c.get('granular_market_match_coverage')
        nota = 'Mercado calculado por insumos declarados + 25% indirecto.' if market_pu else 'Sin referencia granular suficiente.'
        vals = [clave, c.get('desc'), c.get('unidad'), cantidad, pu, importe, market_pu, market_total, diff_pct, coverage, _step08_market_status(diff_pct), nota]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row, col, val); cell.border = _brd(); cell.font = _fnt(size=8)
            if col == 2 or col == 12: cell.alignment = _aln('left','center', wrap=True)
            else: cell.alignment = _aln('center','center')
            if col in (5,6,7,8) and isinstance(val, (int,float)): cell.number_format = '$#,##0.00'
            if col in (9,10) and isinstance(val, (int,float)): cell.number_format = '0.00%'
            if col in (7,9): cell.fill = _fill(C_AMARILLO_PMD)
        fill = _step08_row_fill_by_delta(diff_pct)
        if fill:
            for cc in range(1, len(headers)+1):
                if cc not in (7,9): ws.cell(row,cc).fill = fill
        row += 1
    row += 1
    ws.cell(row, 1, 'TOTAL').font = _fnt(bold=True, size=8)
    ws.cell(row, 6, total_contratista).number_format = '$#,##0.00'; ws.cell(row, 6).font = _fnt(bold=True, size=8)
    ws.cell(row, 8, total_market).number_format = '$#,##0.00'; ws.cell(row, 8).font = _fnt(bold=True, size=8)
    if total_market:
        ws.cell(row, 9, (total_contratista-total_market)/total_market).number_format = '0.00%'
    for cc in range(1, len(headers)+1): ws.cell(row, cc).border = _brd()

    # ---------------- TAB 2: matriz contratista enriquecida ----------------
    ws_det.sheet_view.showGridLines = False
    widths = {'A':15,'B':58,'C':10,'D':13,'E':7,'F':10,'G':14,'H':16,'I':14,'J':14,'K':14,'L':42,'M':12,'N':16,'O':18}
    for col, width in widths.items(): ws_det.column_dimensions[col].width = width
    ws_det.merge_cells('A1:O1')
    ws_det['A1'] = 'MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO'
    ws_det['A1'].fill = _fill(C_AZUL_OSC); ws_det['A1'].font = _fnt(bold=True, color=C_BLANCO, size=9); ws_det['A1'].alignment = _aln('center','center')
    headers_det = ['Código', 'Concepto / Insumo', 'Unidad', 'Costo Contratista', 'Op.', 'Cantidad', 'Importe Contratista', 'Costo Mercado', 'Dif % Costo', 'Importe Mercado', 'Tipo', 'Match Construdata', 'Conf.', 'Estado', 'Nota']
    for i, h in enumerate(headers_det, 1):
        cell = ws_det.cell(3, i, h); cell.fill = _fill(C_AZUL_OSC); cell.font = _fnt(bold=True, color=C_BLANCO, size=8); cell.border = _brd(); cell.alignment = _aln('center','center', wrap=True)
        if i in (8,9): cell.fill = _fill(C_AMARILLO_PMD); cell.font = _fnt(bold=True, color='000000', size=8)
    det_row = 4
    for clave, c in conceptos_items:
        market_pu = _market_pu_with_indirect(c)
        title = f'Análisis: {clave}  |  {c.get("desc", "")}  |  PU Mercado +25%: {market_pu:,.2f}' if isinstance(market_pu,(int,float)) else f'Análisis: {clave}  |  {c.get("desc", "")}'
        ws_det.merge_cells(start_row=det_row, start_column=1, end_row=det_row, end_column=15)
        cell = ws_det.cell(det_row, 1, title); cell.fill = _fill(C_GRIS_SEC); cell.font = _fnt(bold=True, size=8); cell.border = _brd(); cell.alignment = _aln('left','center', wrap=True)
        det_row += 1
        current_section = None
        for item in c.get('granular_market_items') or []:
            if item.get('section') != current_section:
                current_section = item.get('section')
                ws_det.merge_cells(start_row=det_row, start_column=1, end_row=det_row, end_column=15)
                hdr = ws_det.cell(det_row, 1, current_section); hdr.fill = _fill(C_GRIS_FIL); hdr.font = _fnt(bold=True, size=8); hdr.border = _brd(); hdr.alignment = _aln('left','center')
                det_row += 1
            diff = item.get('delta_precio_pct')
            nota = ''
            if item.get('is_percent_item'):
                nota = 'Elemento porcentual: el costo depende de una base de cálculo; revisar contra tab Matriz base CD.'
            elif not item.get('codigo_mercado'):
                nota = 'Sin match en catálogo granular; requiere revisión manual o alias.'
            elif diff is not None and abs(diff) > 0.20:
                nota = 'Diferencia relevante contra costo unitario Construdata.'
            vals = [
                item.get('codigo'), item.get('descripcion'), item.get('unidad'), item.get('precio_proveedor'), '*', item.get('cantidad'), item.get('importe_proveedor'),
                item.get('precio_mercado'), item.get('delta_precio_pct'), item.get('importe_mercado'), item.get('section'),
                f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(' -'), item.get('confidence'), item.get('match_status'), nota
            ]
            for col, val in enumerate(vals, 1):
                cc = ws_det.cell(det_row, col, val); cc.border = _brd(); cc.font = _fnt(size=8)
                if col in (2,12,15): cc.alignment = _aln('left','center', wrap=True)
                else: cc.alignment = _aln('center','center')
                if col in (4,7,8,10) and isinstance(val, (int,float)): cc.number_format = '$#,##0.00'
                if col in (9,13) and isinstance(val, (int,float)): cc.number_format = '0.00%'
                if col in (8,9): cc.fill = _fill(C_AMARILLO_PMD)
            fill = _step08_row_fill_by_delta(diff)
            if fill:
                for cc in range(1, 16):
                    if cc not in (8,9): ws_det.cell(det_row, cc).fill = fill
            det_row += 1
        det_row += 1
    _autosize_sheet(ws_det, max_width=70)

    # ---------------- TAB 3: matriz base/concepto como referencia técnica ----------------
    _step08_build_matrix_base_sheet(wb, dato)
    wb._sheets = [wb[n] for n in ['Comparativa', 'Detalle', 'Matriz base CD'] if n in wb.sheetnames]
    _set_font_size_workbook(wb, 8)
    wb.save(output)
    return output

# Paso 08 performance patch: inverted token index for granular catalogs.
def load_step08_market_catalogs(force=False):
    global _STEP08_CATALOG_CACHE, _STEP08_CATALOG_MTIME
    data_dir = _step08_data_dir()
    files = {k: data_dir / v for k, v in STEP08_CATALOG_FILES.items()}
    mtime = tuple((str(p), p.stat().st_mtime if p.exists() else None) for p in files.values())
    if _STEP08_CATALOG_CACHE is not None and _STEP08_CATALOG_MTIME == mtime and not force:
        return _STEP08_CATALOG_CACHE
    by_domain = {'material': [], 'mano_obra': [], 'equipo': []}
    token_index = {'material': defaultdict(set), 'mano_obra': defaultdict(set), 'equipo': defaultdict(set)}
    for domain, path in files.items():
        by_domain[domain] = _step08_read_catalog_file(path, domain)
        for idx, rec in enumerate(by_domain[domain]):
            for tok in set(rec.get('tokens') or []):
                token_index[domain][tok].add(idx)
    by_code = defaultdict(list)
    for domain, rows in by_domain.items():
        for rec in rows:
            if rec.get('codigo_norm'):
                by_code[(domain, rec['codigo_norm'])].append(rec)
    _STEP08_CATALOG_CACHE = {'by_domain': by_domain, 'by_code': by_code, 'token_index': token_index, 'files': {k: str(v) for k, v in files.items()}}
    _STEP08_CATALOG_MTIME = mtime
    return _STEP08_CATALOG_CACHE


def _step08_match_market_item(item, catalogs):
    domain = item.get('_domain') or 'material'
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    candidates = []
    if code_norm and (domain, code_norm) in by_code:
        candidates = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)]]
    else:
        domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
        token_index = (catalogs.get('token_index') or {}).get(domain, {})
        item_tokens = set(_step08_item_tokens(item))
        idxs = set()
        for tok in item_tokens:
            idxs.update(token_index.get(tok, set()))
        prelim = []
        for idx in idxs:
            if idx >= len(domain_rows):
                continue
            cand = domain_rows[idx]
            sc = _step08_score_catalog_candidate(item, cand)
            if sc >= 0.18:
                c = dict(cand)
                c['score'] = sc
                c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
                prelim.append(c)
        prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
        candidates = prelim[:8]
    selected = candidates[0] if candidates and (candidates[0].get('score') or 0) >= 0.34 else None
    return {
        'selected': selected,
        'candidates': candidates,
        'status': 'match' if selected else ('candidatos_baja_confianza' if candidates else 'sin_match'),
        'confidence': selected.get('score') if selected else (candidates[0].get('score') if candidates else 0.0),
    }

# Paso 08 semantic patch: cuadrillas declaradas por contratista se valorizan por roles individuales.
def _step08_labor_role_candidates(catalogs):
    rows = (catalogs.get('by_domain') or {}).get('mano_obra', [])
    aliases = []
    for rec in rows:
        desc = _normalize_text(rec.get('descripcion'))
        aliases.append((desc, rec))
    return aliases


def _step08_find_labor_role(role_text, catalogs):
    role = _normalize_text(role_text)
    # normalización de plural/sinónimos frecuentes
    role = role.replace('peones', 'peon').replace('ay esp', 'ayudante especializado').replace('ayudantes', 'ayudante')
    if 'albanil' in role:
        role = 'oficial albanil'
    if 'ayudante general' in role:
        role = 'ayudante general'
    if 'ayudante especializado' in role or 'ay especializado' in role:
        role = 'ayudante especializado'
    if role.strip() == 'peon':
        role = 'peon'
    best = None; best_score = 0.0
    role_tokens = set(_tokenize_text(role))
    for desc, rec in _step08_labor_role_candidates(catalogs):
        dt = set(_tokenize_text(desc))
        if not role_tokens or not dt:
            continue
        sc = len(role_tokens & dt) / max(len(role_tokens | dt), 1)
        if role in desc:
            sc = max(sc, 0.92)
        if sc > best_score:
            best_score = sc; best = rec
    return best if best_score >= 0.34 else None, best_score


def _step08_parse_cuadrilla_composition(text):
    norm = _normalize_text(text)
    if 'cuadrilla' not in norm:
        return []
    # Remueve texto previo a paréntesis si existe.
    if '(' in str(text) and ')' in str(text):
        inner = str(text)[str(text).find('(')+1:str(text).rfind(')')]
        norm = _normalize_text(inner)
    parts = re.split(r'\s*\+\s*|,| y ', norm)
    comps = []
    for part in parts:
        part = part.strip()
        m = re.match(r'(\d+(?:\.\d+)?)\s+(.+)', part)
        if not m:
            continue
        qty = _as_float(m.group(1))
        role = m.group(2).strip()
        if qty and role:
            comps.append((qty, role))
    return comps


def _step08_match_cuadrilla_item(item, catalogs):
    text = f"{item.get('codigo','')} {item.get('descripcion','')}"
    comps = _step08_parse_cuadrilla_composition(text)
    if not comps:
        return None
    total = 0.0
    parts = []
    min_conf = 1.0
    for qty, role in comps:
        rec, score = _step08_find_labor_role(role, catalogs)
        if not rec or not isinstance(rec.get('costo'), (int,float)):
            return None
        total += float(qty) * float(rec['costo'])
        min_conf = min(min_conf, score)
        parts.append(f"{qty:g} x {rec.get('descripcion')} (${rec.get('costo'):,.2f})")
    selected = {
        'domain': 'mano_obra',
        'codigo': 'CUADRILLA_CALCULADA',
        'codigo_norm': 'CUADRILLA_CALCULADA',
        'descripcion': ' + '.join(parts),
        'unidad': item.get('unidad') or 'JOR',
        'unidad_norm': normalize_unit(item.get('unidad') or 'JOR'),
        'id_tipo': 2,
        'tipo_insumo': 'MANO DE OBRA',
        'costo': round(total, 6),
        'cost_label': 'Suma de roles MO Construdata',
        'tokens': _tokenize_text(' '.join(parts)),
        'source_file': STEP08_CATALOG_FILES['mano_obra'],
        'source_row': None,
        'score': max(0.80, min_conf),
        'match_reason': 'cuadrilla calculada por roles individuales del catálogo MO',
    }
    return {'selected': selected, 'candidates': [selected], 'status': 'match_cuadrilla_calculada', 'confidence': selected['score']}


def _step08_match_market_item(item, catalogs):
    # Prioridad especial: cuadrillas de MO se valorizan por roles, no contra un solo ayudante.
    if item.get('_domain') == 'mano_obra' and 'cuadrilla' in _normalize_text(item.get('descripcion')):
        special = _step08_match_cuadrilla_item(item, catalogs)
        if special:
            return special
    domain = item.get('_domain') or 'material'
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    candidates = []
    if code_norm and (domain, code_norm) in by_code:
        candidates = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)]]
    else:
        domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
        token_index = (catalogs.get('token_index') or {}).get(domain, {})
        item_tokens = set(_step08_item_tokens(item))
        idxs = set()
        for tok in item_tokens:
            idxs.update(token_index.get(tok, set()))
        prelim = []
        for idx in idxs:
            if idx >= len(domain_rows):
                continue
            cand = domain_rows[idx]
            # Evita que una cuadrilla se convierta en un solo oficial/ayudante.
            if 'cuadrilla' in _normalize_text(item.get('descripcion')) and 'cuadrilla' not in _normalize_text(cand.get('descripcion')):
                continue
            sc = _step08_score_catalog_candidate(item, cand)
            if sc >= 0.18:
                c = dict(cand)
                c['score'] = sc
                c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
                prelim.append(c)
        prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
        candidates = prelim[:8]
    selected = candidates[0] if candidates and (candidates[0].get('score') or 0) >= 0.34 else None
    return {'selected': selected, 'candidates': candidates, 'status': 'match' if selected else ('candidatos_baja_confianza' if candidates else 'sin_match'), 'confidence': selected.get('score') if selected else (candidates[0].get('score') if candidates else 0.0)}

# Paso 08 patch: parseo robusto de composiciones de cuadrilla con '+'.
def _step08_parse_cuadrilla_composition(text):
    raw = str(text or '')
    if 'cuadrilla' not in _normalize_text(raw):
        return []
    if '(' in raw and ')' in raw:
        raw = raw[raw.find('(')+1:raw.rfind(')')]
    raw = raw.replace('+', ' + ').replace(',', ' + ')
    parts = [p.strip() for p in re.split(r'\s+\+\s+|\s+y\s+', raw, flags=re.IGNORECASE) if p.strip()]
    comps = []
    for part in parts:
        norm = _normalize_text(part)
        m = re.match(r'(\d+(?:\.\d+)?)\s+(.+)', norm)
        if not m:
            continue
        qty = _as_float(m.group(1))
        role = m.group(2).strip()
        if qty and role:
            comps.append((qty, role))
    return comps

# Paso 08 fast technical matrix attach: deterministic only, no Voyage/Claude in report generation.
def _step08_fast_match_concept(provider_key, provider_concept, construdata_concepts, concept_index=None):
    if not construdata_concepts:
        return {'status': 'sin_base_construdata', 'selected': None, 'score': 0.0, 'reason': 'No se cargó base de matrices'}
    key = _base_clave(provider_concept.get('clave_base') or provider_concept.get('clave') or provider_key)
    if key in construdata_concepts:
        return {'status': 'match_exacto', 'selected': construdata_concepts[key], 'score': 1.0, 'reason': 'clave exacta'}
    p_tokens = set(_concept_match_tokens(provider_concept.get('desc') or '')) if '_concept_match_tokens' in globals() else set(_concept_tokens(provider_concept.get('desc') or ''))
    p_unit = normalize_unit(provider_concept.get('unidad'))
    counts = defaultdict(int)
    if concept_index:
        for tok in p_tokens:
            for code in concept_index.get('token_index', {}).get(tok, ()): counts[code] += 1
        if p_unit:
            for code in concept_index.get('unit_index', {}).get(p_unit, ()): counts[code] += 1
    candidate_codes = [c for _, c in sorted(((v,k) for k,v in counts.items()), reverse=True)[:150]] or list(construdata_concepts.keys())[:250]
    best = None; best_score = 0.0; best_reason = ''
    for code in candidate_codes:
        cd = construdata_concepts.get(code)
        if not isinstance(cd, dict):
            continue
        sc = _concept_similarity(provider_concept, cd)
        # bonifica tokens de matriz compartidos para conceptos compuestos/malos textos
        try:
            pmat = set(_matrix_summary_tokens(_provider_matrix_items(provider_concept), limit=30))
            cmat = set(_matrix_summary_tokens(cd.get('matriz') or [], limit=30))
            if pmat and cmat:
                sc += min(0.18, len(pmat & cmat) * 0.025)
        except Exception:
            pass
        sc = min(sc, 1.0)
        if sc > best_score:
            best = cd; best_score = sc; best_reason = 'match determinístico por descripción/unidad/matriz'
    status = 'match_probable' if best_score >= 0.45 else ('match_baja_confianza' if best_score >= 0.30 else 'sin_match')
    return {'status': status, 'selected': best if best_score >= 0.30 else None, 'score': round(best_score,4), 'reason': best_reason or 'sin candidatos suficientes'}


def attach_construdata_matrix_analysis(conceptos, precios_nacional):
    construdata_concepts = _get_construdata_concepts(precios_nacional)
    if not construdata_concepts:
        return conceptos
    concept_index = _build_construdata_token_index(construdata_concepts) if '_build_construdata_token_index' in globals() else None
    for key, concept in (conceptos or {}).items():
        match = _step08_fast_match_concept(key, concept, construdata_concepts, concept_index=concept_index)
        comparison = compare_provider_matrix_vs_construdata(key, concept, match)
        concept['construdata_concept_match'] = match
        concept['construdata_matrix_comparison'] = comparison
        concept['construdata_matrix_findings'] = comparison.get('findings') or []
    return conceptos

# Paso 08 runtime guard: por defecto el flujo principal es mercado granular.
# La matriz base por concepto puede activarse con ENABLE_STEP08_MATRIX_BASE=1,
# porque el catálogo de matrices CD es grande y no debe bloquear la web.
_STEP08_DETERMINISTIC_ATTACH = attach_construdata_matrix_analysis

def attach_construdata_matrix_analysis(conceptos, precios_nacional):
    if os.getenv('ENABLE_STEP08_MATRIX_BASE', '0') != '1':
        for key, concept in (conceptos or {}).items():
            concept['construdata_concept_match'] = {'status': 'desactivado_por_runtime', 'score': None, 'reason': 'Matriz base CD disponible bajo ENABLE_STEP08_MATRIX_BASE=1; el mercado principal se calcula por catálogo granular.'}
            concept['construdata_matrix_comparison'] = {'provider_key': key, 'match_status': 'NO_EJECUTADO', 'match_score': None, 'construdata_codigo': None, 'construdata_desc': None, 'rows': [], 'findings': [{'severity':'media','type':'matriz_base_no_ejecutada','message':'Validación de matriz base no ejecutada en este runtime.'}]}
            concept['construdata_matrix_findings'] = concept['construdata_matrix_comparison']['findings']
        return conceptos
    return _STEP08_DETERMINISTIC_ATTACH(conceptos, precios_nacional)

# Paso 08 safe writer: versión sin autosize/font-scan para evitar tiempos largos en Railway/openpyxl.
def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comparativa'
    ws_det = wb.create_sheet('Detalle')
    conceptos_items = list((dato.get('conceptos') or {}).items())
    total_contratista = round(sum(v.get('total') or 0 for _, v in conceptos_items if isinstance(v.get('total'), (int, float))), 2)

    # Tab 1
    for col, width in {'A':12,'B':58,'C':10,'D':12,'E':13,'F':14,'G':14,'H':14,'I':12,'J':14,'K':14,'L':24}.items(): ws.column_dimensions[col].width = width
    ws.merge_cells('A1:L1'); ws['A1'] = f"COMPARATIVO DE COTIZACIÓN — {meta.get('proyecto','')}"; ws['A1'].fill=_fill(C_AZUL_OSC); ws['A1'].font=_fnt(bold=True,color=C_BLANCO,size=10); ws['A1'].alignment=_aln('center','center')
    headers = ['Servicio','Descripción','Unidad','Cantidad','PU Contratista','Importe Contratista','PU Mercado +25%','Importe Mercado','Dif %','Cobertura mercado','Estado','Nota']
    for col,h in enumerate(headers,1):
        cell=ws.cell(4,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (7,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (7,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    row=5; total_market=0.0
    for clave,c in conceptos_items:
        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'),(int,float)) and c.get('cantidad') else 1
        pu = c.get('pu') if isinstance(c.get('pu'),(int,float)) else None
        importe = _concept_total_amount(c)
        market_pu = _market_pu_with_indirect(c)
        market_total = market_pu*cantidad if isinstance(market_pu,(int,float)) else None
        if isinstance(market_total,(int,float)): total_market += market_total
        diff_pct = ((importe-market_total)/market_total) if isinstance(importe,(int,float)) and isinstance(market_total,(int,float)) and market_total else None
        vals=[clave,c.get('desc'),c.get('unidad'),cantidad,pu,importe,market_pu,market_total,diff_pct,c.get('granular_market_match_coverage'),_step08_market_status(diff_pct),'Mercado por insumos declarados + 25% indirecto']
        for col,val in enumerate(vals,1):
            cell=ws.cell(row,col,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True) if col in (2,12) else _aln('center','center')
            if col in (5,6,7,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
            if col in (9,10) and isinstance(val,(int,float)): cell.number_format='0.00%'
            if col in (7,9): cell.fill=_fill(C_AMARILLO_PMD)
        row+=1
    row+=1
    for col,val in [(1,'TOTAL'),(6,total_contratista),(8,total_market),(9,((total_contratista-total_market)/total_market if total_market else None))]:
        cell=ws.cell(row,col,val); cell.font=_fnt(bold=True,size=8); cell.border=_brd()
        if col in (6,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
        if col==9 and isinstance(val,(int,float)): cell.number_format='0.00%'

    # Tab 2
    for col,width in {'A':15,'B':58,'C':10,'D':13,'E':7,'F':10,'G':14,'H':16,'I':14,'J':14,'K':18,'L':42,'M':12,'N':18,'O':24}.items(): ws_det.column_dimensions[col].width=width
    ws_det.merge_cells('A1:O1'); ws_det['A1']='MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO'; ws_det['A1'].fill=_fill(C_AZUL_OSC); ws_det['A1'].font=_fnt(bold=True,color=C_BLANCO,size=9); ws_det['A1'].alignment=_aln('center','center')
    headers_det=['Código','Concepto / Insumo','Unidad','Costo Contratista','Op.','Cantidad','Importe Contratista','Costo Mercado','Dif % Costo','Importe Mercado','Tipo','Match Construdata','Conf.','Estado','Nota']
    for col,h in enumerate(headers_det,1):
        cell=ws_det.cell(3,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (8,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (8,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    det_row=4
    for clave,c in conceptos_items:
        ws_det.cell(det_row,1,f'Análisis: {clave} | {c.get("desc","")}').fill=_fill(C_GRIS_SEC)
        ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
        ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=15)
        det_row+=1
        current=None
        for item in c.get('granular_market_items') or []:
            if item.get('section') != current:
                current=item.get('section')
                ws_det.cell(det_row,1,current).fill=_fill(C_GRIS_FIL); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
                ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=15)
                det_row+=1
            diff=item.get('delta_precio_pct')
            nota = 'Porcentual: depende de base de cálculo.' if item.get('is_percent_item') else ('Sin match: revisar alias/catálogo.' if not item.get('codigo_mercado') else '')
            vals=[item.get('codigo'),item.get('descripcion'),item.get('unidad'),item.get('precio_proveedor'),'*',item.get('cantidad'),item.get('importe_proveedor'),item.get('precio_mercado'),diff,item.get('importe_mercado'),item.get('section'),f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(' -'),item.get('confidence'),item.get('match_status'),nota]
            for col,val in enumerate(vals,1):
                cc=ws_det.cell(det_row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,12,15) else _aln('center','center')
                if col in (4,7,8,10) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
                if col in (9,13) and isinstance(val,(int,float)): cc.number_format='0.00%'
                if col in (8,9): cc.fill=_fill(C_AMARILLO_PMD)
            det_row+=1
        det_row+=1

    _step08_build_matrix_base_sheet(wb, dato)
    wb._sheets=[wb[n] for n in ['Comparativa','Detalle','Matriz base CD'] if n in wb.sheetnames]
    wb.save(output)
    return output

# -----------------------------------------------------------------------------
# Paso 09 - Ajustes catalogo granular / cuadrillas / operaciones reales
# -----------------------------------------------------------------------------
# Cambios:
# 1) Usa el catalogo nuevo construdata-manodeobra-052026-2.xlsx.
# 2) Para cuadrillas, primero busca costo oficial en catalogo MO; solo si no hay
#    match suficiente calcula por composicion.
# 3) Agrega aliases tecnicos (CAL HIDRATADA -> CALHIDRA, etc.) para mejorar
#    matching granular de materiales.
# 4) Si no hay match, usa el costo del contratista como costo mercado fallback.
# 5) El importe mercado respeta la operacion declarada por el contratista (*,/,%).

STEP08_CATALOG_FILES = {
    'material': 'construdata-materiales-052026.xlsx',
    'mano_obra': 'construdata-manodeobra-052026-2.xlsx',
    'equipo': 'construdata-maquinaria-052026.xlsx',
}

_STEP09_ALIAS_TOKENS = {
    'calhidra': {'calhidra', 'cal', 'hidratada', 'hidra', 'hidrato'},
    'hidratada': {'calhidra', 'cal', 'hidratada', 'hidra', 'hidrato'},
    'cal': {'cal', 'calhidra', 'hidratada'},
    'rafia': {'costal', 'saco', 'rafia'},
    'costal': {'costal', 'saco', 'rafia'},
    'saco': {'saco', 'costal', 'bulto'},
    'rompedora': {'rompedora', 'demoledor', 'rotomartillo', 'martillo'},
    'electric': {'electrica', 'electrico', 'electric'},
    'electrica': {'electrica', 'electrico', 'electric'},
    'peon': {'peon', 'ayudante'},
    'ayudante': {'ayudante', 'peon'},
    'albanil': {'albanil', 'oficial'},
}


def _step09_expand_tokens(tokens):
    expanded = set(tokens or [])
    for tok in list(expanded):
        expanded.update(_STEP09_ALIAS_TOKENS.get(tok, set()))
    # Alias por co-ocurrencia: cal + hidratada debe encontrar CALHIDRA.
    if {'cal', 'hidratada'} & expanded:
        expanded.update({'calhidra', 'hidra', 'hidrato'})
    return [t for t in expanded if t]


def _step08_item_tokens(item):
    return _step09_expand_tokens(_tokenize_text(f"{item.get('codigo','')} {item.get('descripcion','')}"))


def _step09_augment_catalog_tokens(records):
    for rec in records or []:
        rec['tokens'] = _step09_expand_tokens(rec.get('tokens') or _tokenize_text(f"{rec.get('codigo','')} {rec.get('descripcion','')}"))
        rec['norm'] = _normalize_text(f"{rec.get('codigo','')} {rec.get('descripcion','')}")
    return records


def load_step08_market_catalogs(force=False):
    global _STEP08_CATALOG_CACHE, _STEP08_CATALOG_MTIME
    data_dir = _step08_data_dir()
    files = {k: data_dir / v for k, v in STEP08_CATALOG_FILES.items()}
    mtime = tuple((str(p), p.stat().st_mtime if p.exists() else None) for p in files.values())
    if _STEP08_CATALOG_CACHE is not None and _STEP08_CATALOG_MTIME == mtime and not force:
        return _STEP08_CATALOG_CACHE
    by_domain = {'material': [], 'mano_obra': [], 'equipo': []}
    token_index = {'material': defaultdict(set), 'mano_obra': defaultdict(set), 'equipo': defaultdict(set)}
    for domain, path in files.items():
        rows = _step08_read_catalog_file(path, domain)
        by_domain[domain] = _step09_augment_catalog_tokens(rows)
        for idx, rec in enumerate(by_domain[domain]):
            for tok in set(rec.get('tokens') or []):
                token_index[domain][tok].add(idx)
    by_code = defaultdict(list)
    for domain, rows in by_domain.items():
        for rec in rows:
            if rec.get('codigo_norm'):
                by_code[(domain, rec['codigo_norm'])].append(rec)
    _STEP08_CATALOG_CACHE = {'by_domain': by_domain, 'by_code': by_code, 'token_index': token_index, 'files': {k: str(v) for k, v in files.items()}}
    _STEP08_CATALOG_MTIME = mtime
    return _STEP08_CATALOG_CACHE


def _parse_component_row(row):
    c0 = _clean_text(row[0] if len(row) > 0 else "")
    c1 = _clean_text(row[1] if len(row) > 1 else "")
    if not c0 and not c1:
        return None
    if c0.upper().startswith("SUBTOTAL"):
        return None
    if _section_kind(c0) or _section_kind(c1):
        return None
    txt = _normalize_text(c0 + " " + c1)
    if "precio unitario" in txt or "costo directo" in txt or "indirect" in txt or "utilidad" in txt or "subtotal1" in txt or "subtotal2" in txt:
        return None
    unidad = _clean_text(row[2] if len(row) > 2 else "")
    precio_base = _as_float(row[3] if len(row) > 3 else None)
    op = _clean_text(row[4] if len(row) > 4 else "*") or "*"
    factor = _as_float(row[5] if len(row) > 5 else None)
    importe = _as_float(row[6] if len(row) > 6 else None)
    if precio_base is None and factor is None and importe is None:
        return None
    return {
        "codigo": c0,
        "descripcion": c1 or c0,
        "unidad": unidad,
        "unidad_norm": normalize_unit(unidad),
        "precio_base": precio_base,
        "factor": factor,
        "importe": importe,
        "op": op,
        "operacion": op,
        "attributes": _extract_material_attributes(c1 or c0, unidad),
    }


def _step09_is_cuadrilla(item):
    return 'cuadrilla' in _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')}")


def _step09_direct_catalog_match(item, catalogs, domain):
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    if code_norm and (domain, code_norm) in by_code:
        candidates = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)]]
        return candidates[0], candidates
    domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
    token_index = (catalogs.get('token_index') or {}).get(domain, {})
    item_tokens = set(_step08_item_tokens(item))
    idxs = set()
    for tok in item_tokens:
        idxs.update(token_index.get(tok, set()))
    prelim = []
    for idx in idxs:
        if idx >= len(domain_rows):
            continue
        cand = domain_rows[idx]
        # Si el insumo proveedor es cuadrilla, no permitir que empate contra un rol individual salvo fallback calculado.
        if _step09_is_cuadrilla(item) and 'cuadrilla' not in _normalize_text(cand.get('descripcion')):
            continue
        sc = _step08_score_catalog_candidate(item, cand)
        # Bonificacion por alias fuertes.
        pt = set(_step08_item_tokens(item)); ct = set(cand.get('tokens') or [])
        if {'calhidra', 'cal', 'hidratada'} & pt and {'calhidra', 'cal', 'hidratada'} & ct:
            sc = max(sc, 0.72)
        if _step09_is_cuadrilla(item) and 'cuadrilla' in _normalize_text(cand.get('descripcion')):
            sc = max(sc, 0.65)
        if sc >= 0.18:
            c = dict(cand); c['score'] = sc; c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
            prelim.append(c)
    prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
    selected = prelim[0] if prelim and (prelim[0].get('score') or 0) >= 0.34 else None
    return selected, prelim[:8]


def _step08_match_market_item(item, catalogs):
    domain = item.get('_domain') or 'material'
    selected, candidates = _step09_direct_catalog_match(item, catalogs, domain)
    # Prioridad nueva: para cuadrillas primero usar costo oficial catalogado.
    if selected:
        return {'selected': selected, 'candidates': candidates, 'status': 'match', 'confidence': selected.get('score')}
    # Solo si no existe cuadrilla catalogada, calcularla por composicion.
    if domain == 'mano_obra' and _step09_is_cuadrilla(item):
        special = _step08_match_cuadrilla_item(item, catalogs)
        if special:
            special['selected']['match_reason'] = 'sin cuadrilla oficial suficiente; fallback calculado por roles MO'
            return special
    return {'selected': None, 'candidates': candidates or [], 'status': 'candidatos_baja_confianza' if candidates else 'sin_match', 'confidence': (candidates[0].get('score') if candidates else 0.0)}


def _step09_apply_operation(price, qty, op):
    if not isinstance(price, (int, float)):
        return None
    if not isinstance(qty, (int, float)):
        return None
    opn = str(op or '*').strip()
    if opn == '/':
        if qty == 0:
            return None
        return float(price) / float(qty)
    return float(price) * float(qty)


def _step09_percent_base_kind(item):
    code = _normalize_text(item.get('codigo'))
    desc = _normalize_text(item.get('descripcion'))
    text = f'{code} {desc}'
    if 'mat' in code or 'material' in desc:
        return 'material'
    if 'eq' in code or 'equipo' in desc or 'seguridad' in desc or item.get('_section') == 'EQUIPO Y HERRAMIENTA':
        # En los archivos del contratista los porcentuales de equipo/herramienta suelen usar base MO.
        return 'mano_obra'
    if 'mo' in code or 'herramienta' in desc or 'mano obra' in desc:
        return 'mano_obra'
    return 'mano_obra'


def _step09_market_importe_for_item(item, market_price, subtotals):
    qty = item.get('factor') if isinstance(item.get('factor'), (int, float)) else item.get('cantidad')
    op = item.get('op') or item.get('operacion') or '*'
    if _step08_is_percent_item(item):
        base_kind = _step09_percent_base_kind(item)
        base = subtotals.get(base_kind)
        if isinstance(base, (int, float)) and isinstance(qty, (int, float)):
            return float(base) * float(qty), base
        # fallback seguro: replica base declarada por contratista si no hay subtotal mercado.
        provider_base = item.get('precio_base')
        if isinstance(provider_base, (int, float)) and isinstance(qty, (int, float)):
            return float(provider_base) * float(qty), provider_base
        return item.get('importe'), provider_base
    return _step09_apply_operation(market_price, qty, op), market_price


def _step08_apply_market_catalog_pricing(conceptos):
    """Enriquece cada concepto con precio mercado por insumo, respetando op. contratista."""
    try:
        catalogs = load_step08_market_catalogs()
    except Exception as exc:
        for c in (conceptos or {}).values():
            c['granular_market_items'] = []
            c['granular_market_error'] = str(exc)
        return conceptos

    for _, c in (conceptos or {}).items():
        raw_items = list(_step08_iter_provider_items(c))
        first_pass = []
        for item in raw_items:
            match = _step08_match_market_item(item, catalogs)
            selected = match.get('selected') or {}
            provider_price = item.get('precio_base')
            market_price = selected.get('costo') if isinstance(selected.get('costo'), (int, float)) else provider_price
            fallback_used = not bool(selected)
            first_pass.append((item, match, selected, market_price, fallback_used))

        # Subtotales de mercado no porcentuales por dominio para calcular % sobre base correcta.
        subtotals = {'material': 0.0, 'mano_obra': 0.0, 'equipo': 0.0}
        for item, match, selected, market_price, fallback_used in first_pass:
            if _step08_is_percent_item(item):
                continue
            imp = _step09_apply_operation(market_price, item.get('factor'), item.get('op') or item.get('operacion') or '*')
            if isinstance(imp, (int, float)):
                subtotals[item.get('_domain') or 'material'] = subtotals.get(item.get('_domain') or 'material', 0.0) + float(imp)

        enriched = []
        direct_market = 0.0
        provider_direct = 0.0
        matched_count = 0
        total_count = 0
        for item, match, selected, market_price, fallback_used in first_pass:
            total_count += 1
            provider_price = item.get('precio_base')
            provider_importe = item.get('importe')
            market_importe, market_base = _step09_market_importe_for_item(item, market_price, subtotals)
            delta_price = None
            if isinstance(provider_price, (int, float)) and isinstance(market_price, (int, float)) and market_price:
                # En SIN_MATCH el mercado es proveedor; delta 0 pero status conserva SIN_MATCH.
                delta_price = (float(provider_price) - float(market_price)) / float(market_price)
            if isinstance(market_importe, (int, float)):
                direct_market += float(market_importe)
            if isinstance(provider_importe, (int, float)):
                provider_direct += float(provider_importe)
            if selected:
                matched_count += 1
            status = match.get('status') if selected else 'sin_match_fallback_contratista'
            reason = selected.get('match_reason') if selected else 'sin match en catalogo; costo mercado igual al costo contratista como fallback'
            enriched.append({
                'section': item.get('_section'),
                'domain': item.get('_domain'),
                'codigo': item.get('codigo'),
                'descripcion': item.get('descripcion'),
                'unidad': item.get('unidad'),
                'cantidad': item.get('factor'),
                'op': item.get('op') or item.get('operacion') or '*',
                'precio_proveedor': provider_price,
                'importe_proveedor': provider_importe,
                'codigo_mercado': selected.get('codigo'),
                'descripcion_mercado': selected.get('descripcion'),
                'unidad_mercado': selected.get('unidad') if selected else item.get('unidad'),
                'precio_mercado': market_price,
                'importe_mercado': market_importe,
                'base_mercado': market_base if _step08_is_percent_item(item) else None,
                'delta_precio_pct': delta_price,
                'match_status': status,
                'confidence': match.get('confidence') if selected else 0.0,
                'match_reason': reason,
                'source_file': selected.get('source_file'),
                'source_row': selected.get('source_row'),
                'is_percent_item': _step08_is_percent_item(item),
                'fallback_contratista': fallback_used,
                'top_candidates': match.get('candidates') or [],
            })
        c['granular_market_items'] = enriched
        c['granular_market_direct'] = round(direct_market, 6) if direct_market else None
        c['granular_provider_direct'] = round(provider_direct, 6) if provider_direct else None
        c['granular_market_match_coverage'] = (matched_count / total_count) if total_count else None
    return conceptos


# Writer actualizado: respeta Op. real y muestra fallback, base porcentual y match.
def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comparativa'
    ws_det = wb.create_sheet('Detalle')
    conceptos_items = list((dato.get('conceptos') or {}).items())
    total_contratista = round(sum(v.get('total') or 0 for _, v in conceptos_items if isinstance(v.get('total'), (int, float))), 2)

    for col, width in {'A':12,'B':58,'C':10,'D':12,'E':13,'F':14,'G':14,'H':14,'I':12,'J':14,'K':14,'L':24}.items():
        ws.column_dimensions[col].width = width
    ws.merge_cells('A1:L1')
    ws['A1'] = f"COMPARATIVO DE COTIZACIÓN — {meta.get('proyecto','')}"
    ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10); ws['A1'].alignment = _aln('center','center')
    headers = ['Servicio','Descripción','Unidad','Cantidad','PU Contratista','Importe Contratista','PU Mercado +25%','Importe Mercado','Dif %','Cobertura mercado','Estado','Nota']
    for col,h in enumerate(headers,1):
        cell=ws.cell(4,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (7,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (7,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    row=5; total_market=0.0
    for clave,c in conceptos_items:
        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'),(int,float)) and c.get('cantidad') else 1
        pu = c.get('pu') if isinstance(c.get('pu'),(int,float)) else None
        importe = _concept_total_amount(c)
        market_pu = _market_pu_with_indirect(c)
        market_total = market_pu*cantidad if isinstance(market_pu,(int,float)) else None
        if isinstance(market_total,(int,float)): total_market += market_total
        diff_pct = ((importe-market_total)/market_total) if isinstance(importe,(int,float)) and isinstance(market_total,(int,float)) and market_total else None
        vals=[clave,c.get('desc'),c.get('unidad'),cantidad,pu,importe,market_pu,market_total,diff_pct,c.get('granular_market_match_coverage'),_step08_market_status(diff_pct),'Mercado por insumos declarados con operación del contratista + 25% indirecto']
        for col,val in enumerate(vals,1):
            cell=ws.cell(row,col,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True) if col in (2,12) else _aln('center','center')
            if col in (5,6,7,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
            if col in (9,10) and isinstance(val,(int,float)): cell.number_format='0.00%'
            if col in (7,9): cell.fill=_fill(C_AMARILLO_PMD)
        row+=1
    row+=1
    for col,val in [(1,'TOTAL'),(6,total_contratista),(8,total_market),(9,((total_contratista-total_market)/total_market if total_market else None))]:
        cell=ws.cell(row,col,val); cell.font=_fnt(bold=True,size=8); cell.border=_brd()
        if col in (6,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
        if col==9 and isinstance(val,(int,float)): cell.number_format='0.00%'

    # Tab 2 limpio: matriz contratista + columnas amarillas de mercado.
    widths={'A':15,'B':58,'C':10,'D':13,'E':7,'F':10,'G':14,'H':16,'I':14,'J':14,'K':14,'L':18,'M':42,'N':12,'O':20,'P':26}
    for col,width in widths.items(): ws_det.column_dimensions[col].width=width
    ws_det.merge_cells('A1:P1'); ws_det['A1']='MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO'; ws_det['A1'].fill=_fill(C_AZUL_OSC); ws_det['A1'].font=_fnt(bold=True,color=C_BLANCO,size=9); ws_det['A1'].alignment=_aln('center','center')
    headers_det=['Código','Concepto / Insumo','Unidad','Costo Contratista','Op.','Cantidad','Importe Contratista','Costo Mercado','Dif % Costo','Importe Mercado','Base Mercado %','Tipo','Match Construdata','Conf.','Estado','Nota']
    for col,h in enumerate(headers_det,1):
        cell=ws_det.cell(3,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (8,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (8,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    det_row=4
    for clave,c in conceptos_items:
        ws_det.cell(det_row,1,f'Análisis: {clave} | {c.get("desc","")}').fill=_fill(C_GRIS_SEC)
        ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
        ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
        det_row+=1
        current=None
        for item in c.get('granular_market_items') or []:
            if item.get('section') != current:
                current=item.get('section')
                ws_det.cell(det_row,1,current).fill=_fill(C_GRIS_FIL); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
                ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
                det_row+=1
            diff=item.get('delta_precio_pct')
            if item.get('fallback_contratista'):
                nota='Sin match: costo mercado = costo contratista como fallback.'
            elif item.get('is_percent_item'):
                nota='Porcentual: importe mercado calculado sobre base mercado correspondiente.'
            elif item.get('match_reason'):
                nota=item.get('match_reason')
            else:
                nota=''
            vals=[item.get('codigo'),item.get('descripcion'),item.get('unidad'),item.get('precio_proveedor'),item.get('op') or '*',item.get('cantidad'),item.get('importe_proveedor'),item.get('precio_mercado'),diff,item.get('importe_mercado'),item.get('base_mercado'),item.get('section'),f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(' -'),item.get('confidence'),item.get('match_status'),nota]
            for col,val in enumerate(vals,1):
                cc=ws_det.cell(det_row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,13,16) else _aln('center','center')
                if col in (4,7,8,10,11) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
                if col in (9,14) and isinstance(val,(int,float)): cc.number_format='0.00%'
                if col in (8,9): cc.fill=_fill(C_AMARILLO_PMD)
            det_row+=1
        det_row+=1

    _step08_build_matrix_base_sheet(wb, dato)
    wb._sheets=[wb[n] for n in ['Comparativa','Detalle','Matriz base CD'] if n in wb.sheetnames]
    wb.save(output)
    return output

# Paso 09b - refinamientos de lectura de catalogo MO, alias CALHIDRA y exclusiones financieras.
def _step08_read_catalog_file(path, domain):
    records = []
    if not path.exists():
        return records
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        wb.close(); return records
    headers = [_normalize_text(x) for x in rows[0]]
    def idx(*names):
        targets = [_normalize_text(n) for n in names]
        for i, h in enumerate(headers):
            if not h: continue
            for t in targets:
                if h == t or t in h:
                    return i
        return None
    def pick(value, default):
        return value if value is not None else default

    if domain == 'material':
        code_i = pick(idx('codigo','código'), 1); desc_i = pick(idx('descripcion completa','descripción completa','descripcion'), 2)
        unit_i = pick(idx('unidad'), 4); type_i = pick(idx('tipo'), 3); cost_i = pick(idx('costo'), 11); extra_cost_i = None
        cost_label = 'Costo'
    elif domain == 'equipo':
        code_i = pick(idx('codigo','código'), 2); desc_i = pick(idx('descripcion completa','descripción completa','descripcion'), 3)
        unit_i = pick(idx('unidad'), 4); type_i = pick(idx('tipo'), 5); cost_i = pick(idx('costo'), 6); extra_cost_i = None
        cost_label = 'Costo'
    else:
        # En el catálogo MO v3 el código está en columna A (indice 0).
        # No usar `idx(...) or 1`: el índice 0 es válido y se perdía,
        # haciendo que el código quedara igual a la descripción.
        code_i = pick(idx('codigo','código'), 0); desc_i = pick(idx('descripcion completa','descripción completa','descripcion'), 1)
        unit_i = pick(idx('unidad'), 2); type_i = pick(idx('tipo'), 3)
        # Para MO: usar Salario Real cuando exista, pero para cuadrillas el costo suele venir en CostoTotal.
        cost_i = idx('salario real')
        extra_cost_i = idx('costototal','costo total')
        if cost_i is None:
            cost_i = extra_cost_i if extra_cost_i is not None else (idx('costo','salario integrado') or 9)
        cost_label = 'Salario Real / CostoTotal'
    for rnum, row in enumerate(rows[1:], start=2):
        code = _clean_text(row[code_i] if code_i is not None and code_i < len(row) else '')
        desc = _clean_text(row[desc_i] if desc_i is not None and desc_i < len(row) else '')
        if not code and not desc: continue
        unit = _clean_text(row[unit_i] if unit_i is not None and unit_i < len(row) else '')
        idtipo = row[type_i] if type_i is not None and type_i < len(row) else None
        cost = _step08_safe_float(row[cost_i] if cost_i is not None and cost_i < len(row) else None)
        if cost is None and extra_cost_i is not None and extra_cost_i < len(row):
            cost = _step08_safe_float(row[extra_cost_i])
        tokens = _step09_expand_tokens(_tokenize_text(f'{code} {desc}'))
        records.append({
            'domain': domain,
            'codigo': code,
            'codigo_norm': _step08_norm_code(code),
            'descripcion': desc,
            'descripcion_larga': desc,
            'unidad': unit,
            'unidad_norm': normalize_unit(unit),
            'id_tipo': idtipo,
            'tipo_insumo': {'material':'MATERIALES','mano_obra':'MANO DE OBRA','equipo':'EQUIPO Y HERRAMIENTA'}.get(domain, domain),
            'costo': cost,
            'cost_label': cost_label,
            'tokens': tokens,
            'norm': _normalize_text(f'{code} {desc}'),
            'source_file': path.name,
            'source_row': rnum,
        })
    wb.close()
    return records


def _step09_direct_catalog_match(item, catalogs, domain):
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    if code_norm and (domain, code_norm) in by_code:
        candidates = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)] if isinstance(c.get('costo'), (int,float))]
        if candidates:
            return candidates[0], candidates
    domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
    token_index = (catalogs.get('token_index') or {}).get(domain, {})
    item_tokens = set(_step08_item_tokens(item))
    idxs = set()
    for tok in item_tokens:
        idxs.update(token_index.get(tok, set()))
    prelim = []
    item_has_calhidra = bool({'calhidra','hidratada'} & item_tokens) or ('cal hidratada' in _normalize_text(item.get('descripcion')))
    for idx in idxs:
        if idx >= len(domain_rows): continue
        cand = domain_rows[idx]
        if not isinstance(cand.get('costo'), (int,float)):
            continue
        cand_tokens = set(cand.get('tokens') or [])
        cand_desc_norm = _normalize_text(cand.get('descripcion'))
        if _step09_is_cuadrilla(item) and 'cuadrilla' not in cand_desc_norm:
            continue
        sc = _step08_score_catalog_candidate(item, cand)
        # Alias fuerte CAL HIDRATADA: solo bonifica candidatos CALHIDRA, no cualquier token "cal".
        if item_has_calhidra:
            if 'calhidra' in cand_tokens or 'calhidra' in cand_desc_norm:
                sc = max(sc, 0.92 if normalize_unit(item.get('unidad')) == cand.get('unidad_norm') else 0.78)
            elif 'cal' in cand_tokens and 'calhidra' not in cand_tokens:
                sc = min(sc, 0.28)
        if _step09_is_cuadrilla(item) and 'cuadrilla' in cand_desc_norm:
            sc = max(sc, 0.65)
        if sc >= 0.18:
            c = dict(cand); c['score'] = sc; c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
            if item_has_calhidra and ('calhidra' in cand_tokens or 'calhidra' in cand_desc_norm):
                c['match_reason'] = 'alias técnico CAL HIDRATADA -> CALHIDRA'
            prelim.append(c)
    prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
    selected = prelim[0] if prelim and (prelim[0].get('score') or 0) >= 0.34 else None
    return selected, prelim[:8]


def _parse_component_row(row):
    c0 = _clean_text(row[0] if len(row) > 0 else "")
    c1 = _clean_text(row[1] if len(row) > 1 else "")
    if not c0 and not c1: return None
    if c0.upper().startswith("SUBTOTAL"): return None
    if _section_kind(c0) or _section_kind(c1): return None
    txt = _normalize_text(c0 + " " + c1)
    skip_words = ["precio unitario", "costo directo", "indirect", "utilidad", "subtotal1", "subtotal2", "financiamiento", "importe total", "total"]
    if any(w in txt for w in skip_words): return None
    unidad = _clean_text(row[2] if len(row) > 2 else "")
    precio_base = _as_float(row[3] if len(row) > 3 else None)
    op = _clean_text(row[4] if len(row) > 4 else "*") or "*"
    factor = _as_float(row[5] if len(row) > 5 else None)
    importe = _as_float(row[6] if len(row) > 6 else None)
    if precio_base is None and factor is None and importe is None: return None
    return {"codigo": c0, "descripcion": c1 or c0, "unidad": unidad, "unidad_norm": normalize_unit(unidad), "precio_base": precio_base, "factor": factor, "importe": importe, "op": op, "operacion": op, "attributes": _extract_material_attributes(c1 or c0, unidad)}

# Paso 09c - preferencia de aliases específicos sobre genéricos (costal/rafia vs saco genérico).
_STEP09_DIRECT_CATALOG_MATCH_PREV = _step09_direct_catalog_match

def _step09_direct_catalog_match(item, catalogs, domain):
    selected, candidates = _STEP09_DIRECT_CATALOG_MATCH_PREV(item, catalogs, domain)
    if domain == 'material':
        item_tokens = set(_step08_item_tokens(item))
        if {'costal', 'rafia'} & item_tokens:
            adjusted = []
            for c in candidates or []:
                ct = set(c.get('tokens') or [])
                cn = _normalize_text(c.get('descripcion'))
                sc = float(c.get('score') or 0)
                if {'costal', 'rafia'} & ct or 'costal' in cn or 'rafia' in cn:
                    sc = max(sc, 0.86 if normalize_unit(item.get('unidad')) == c.get('unidad_norm') else 0.74)
                elif 'saco' in ct and not ({'costal', 'rafia'} & ct):
                    sc = min(sc, 0.30)
                cc = dict(c); cc['score'] = round(sc, 4)
                if sc >= 0.18:
                    adjusted.append(cc)
            adjusted.sort(key=lambda x: x.get('score') or 0, reverse=True)
            if adjusted:
                return (adjusted[0] if adjusted[0].get('score',0) >= 0.34 else None), adjusted[:8]
    return selected, candidates

# Paso 09d - corrección final: para COSTAL/RAFIA exigir texto explícito, no solo alias saco expandido.
_STEP09_DIRECT_CATALOG_MATCH_PREV2 = _step09_direct_catalog_match

def _step09_direct_catalog_match(item, catalogs, domain):
    selected, candidates = _STEP09_DIRECT_CATALOG_MATCH_PREV2(item, catalogs, domain)
    if domain == 'material':
        item_norm = _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')}")
        if 'costal' in item_norm or 'rafia' in item_norm:
            adjusted = []
            for c in candidates or []:
                cn = _normalize_text(c.get('descripcion'))
                sc = float(c.get('score') or 0)
                if 'costal' in cn or 'rafia' in cn:
                    sc = max(sc, 0.88 if normalize_unit(item.get('unidad')) == c.get('unidad_norm') else 0.76)
                else:
                    sc = min(sc, 0.29)
                if sc >= 0.18:
                    cc = dict(c); cc['score'] = round(sc,4)
                    if 'costal' in cn or 'rafia' in cn:
                        cc['match_reason'] = 'alias técnico COSTAL/RAFIA'
                    adjusted.append(cc)
            adjusted.sort(key=lambda x: x.get('score') or 0, reverse=True)
            if adjusted:
                return (adjusted[0] if adjusted[0].get('score',0) >= 0.34 else None), adjusted[:8]
    return selected, candidates

# Paso 09e - performance: en arquitectura granular no cargar construdata_matrices.xlsx salvo que se pida tab 3 real.
_EXTRACT_PRECIOS_NACIONAL_STEP09_PREV = extract_precios_nacional

def extract_precios_nacional(filepath: str) -> dict:
    try:
        name = Path(filepath).name.lower()
    except Exception:
        name = ''
    if name == 'construdata_matrices.xlsx' and os.getenv('ENABLE_STEP08_MATRIX_BASE', '0') != '1':
        return {
            '__benchmark_kind__': 'construdata_matrices_skipped_runtime',
            '__construdata_matrices__': {},
            '__materials_index__': [],
            '__skip_reason__': 'Arquitectura granular activa; matriz base CD se omite salvo ENABLE_STEP08_MATRIX_BASE=1.',
        }
    return _EXTRACT_PRECIOS_NACIONAL_STEP09_PREV(filepath)

# Paso 09f - matching granular optimizado sin wrappers anidados.
_STEP09_GENERIC_TOKENS = {'saco','kg','pza','m','m2','m3','lt','ton','pieza','unidad','material','equipo','herramienta','de','para','con','tipo'}

def _step09_direct_catalog_match(item, catalogs, domain):
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    if code_norm and (domain, code_norm) in by_code:
        exact = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)] if isinstance(c.get('costo'), (int,float))]
        if exact:
            return exact[0], exact[:8]

    domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
    token_index = (catalogs.get('token_index') or {}).get(domain, {})
    item_tokens = set(_step08_item_tokens(item))
    meaningful = [t for t in item_tokens if t not in _STEP09_GENERIC_TOKENS]
    if not meaningful:
        meaningful = list(item_tokens)[:4]

    counts = defaultdict(int)
    for tok in meaningful[:10]:
        for idx in token_index.get(tok, set()):
            counts[idx] += 1
    # Si no hay candidatos, no escanear todo el catálogo.
    if not counts:
        return None, []
    # Limitar candidatos para rendimiento: más tokens compartidos primero.
    candidate_idxs = [idx for idx, _ in sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:700]]

    prelim = []
    item_norm = _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')}")
    item_has_calhidra = ('cal hidratada' in item_norm) or ('calhidra' in item_norm) or bool({'calhidra','hidratada'} & item_tokens)
    item_is_costal = 'costal' in item_norm or 'rafia' in item_norm
    for idx in candidate_idxs:
        if idx >= len(domain_rows):
            continue
        cand = domain_rows[idx]
        if not isinstance(cand.get('costo'), (int,float)):
            continue
        cn = _normalize_text(cand.get('descripcion'))
        ct = set(cand.get('tokens') or [])
        if _step09_is_cuadrilla(item) and 'cuadrilla' not in cn:
            continue
        sc = _step08_score_catalog_candidate(item, cand)
        if item_has_calhidra:
            if 'calhidra' in ct or 'calhidra' in cn:
                sc = max(sc, 0.92 if normalize_unit(item.get('unidad')) == cand.get('unidad_norm') else 0.78)
            elif 'cal' in ct and 'calhidra' not in ct:
                sc = min(sc, 0.28)
        if item_is_costal:
            if 'costal' in cn or 'rafia' in cn:
                sc = max(sc, 0.88 if normalize_unit(item.get('unidad')) == cand.get('unidad_norm') else 0.76)
            else:
                sc = min(sc, 0.29)
        if _step09_is_cuadrilla(item) and 'cuadrilla' in cn:
            sc = max(sc, 0.65)
        if sc >= 0.18:
            c = dict(cand); c['score'] = round(sc,4); c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
            if item_has_calhidra and ('calhidra' in ct or 'calhidra' in cn): c['match_reason'] = 'alias técnico CAL HIDRATADA -> CALHIDRA'
            if item_is_costal and ('costal' in cn or 'rafia' in cn): c['match_reason'] = 'alias técnico COSTAL/RAFIA'
            prelim.append(c)
    prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
    selected = prelim[0] if prelim and (prelim[0].get('score') or 0) >= 0.34 else None
    return selected, prelim[:8]

# Paso 09g - para renglones porcentuales, mostrar base de mercado como costo mercado.
_STEP09_APPLY_MARKET_CATALOG_PRICING_PREV = _step08_apply_market_catalog_pricing

def _step08_apply_market_catalog_pricing(conceptos):
    conceptos = _STEP09_APPLY_MARKET_CATALOG_PRICING_PREV(conceptos)
    for _, c in (conceptos or {}).items():
        for item in c.get('granular_market_items') or []:
            if item.get('is_percent_item') and isinstance(item.get('base_mercado'), (int, float)):
                item['precio_mercado_catalogo'] = item.get('precio_mercado')
                item['precio_mercado'] = item.get('base_mercado')
                prov = item.get('precio_proveedor')
                mkt = item.get('precio_mercado')
                item['delta_precio_pct'] = ((float(prov) - float(mkt)) / float(mkt)) if isinstance(prov,(int,float)) and isinstance(mkt,(int,float)) and mkt else None
                item['match_reason'] = (item.get('match_reason') or '') + ' | display: costo mercado = base mercado porcentual'
    return conceptos

# Paso 10 - Correcciones críticas de detalle granular.
# 1) No confundir renglones como "TRANSPORTE DIARIO DE EQUIPO Y HERRAMIENTA" con encabezados de sección.
# 2) Todo elemento porcentual (%MO1, %MO2, %MO3, %MO5, %mo8, EPP/SHE, SHE, etc.) se calcula sobre subtotal MO de mercado.
# 3) Ignorar renglones sin costo/factor/importe porque suelen ser títulos o encabezados dentro de la matriz del contratista.

def _step10_row_has_numeric_matrix_values(row):
    precio_base = _as_float(row[3] if len(row) > 3 else None)
    factor = _as_float(row[5] if len(row) > 5 else None)
    importe = _as_float(row[6] if len(row) > 6 else None)
    return precio_base is not None or factor is not None or importe is not None


def _parse_component_row(row):
    c0 = _clean_text(row[0] if len(row) > 0 else "")
    c1 = _clean_text(row[1] if len(row) > 1 else "")
    if not c0 and not c1:
        return None
    if c0.upper().startswith("SUBTOTAL"):
        return None

    unidad = _clean_text(row[2] if len(row) > 2 else "")
    precio_base = _as_float(row[3] if len(row) > 3 else None)
    op = _clean_text(row[4] if len(row) > 4 else "*") or "*"
    factor = _as_float(row[5] if len(row) > 5 else None)
    importe = _as_float(row[6] if len(row) > 6 else None)

    # Si no trae valores de costo/cantidad/importe, se trata como título o renglón descriptivo y no entra al cálculo.
    if precio_base is None and factor is None and importe is None:
        return None

    # Solo descartar encabezados de sección cuando el renglón no tiene valores numéricos.
    # Esto evita perder insumos reales como "TRANSPORTE DIARIO DE EQUIPO Y HERRAMIENTA".
    if (_section_kind(c0) or _section_kind(c1)) and not _step10_row_has_numeric_matrix_values(row):
        return None

    txt = _normalize_text(c0 + " " + c1)
    skip_phrases = [
        "precio unitario",
        "costo directo",
        "indirecto",
        "indirectos",
        "utilidad",
        "subtotal1",
        "subtotal2",
        "financiamiento",
        "importe total",
    ]
    if any(p in txt for p in skip_phrases):
        return None

    return {
        "codigo": c0,
        "descripcion": c1 or c0,
        "unidad": unidad,
        "unidad_norm": normalize_unit(unidad),
        "precio_base": precio_base,
        "factor": factor,
        "importe": importe,
        "op": op,
        "operacion": op,
        "attributes": _extract_material_attributes(c1 or c0, unidad),
    }


def _collect_components(rows, start, end):
    components = {"materiales": [], "mano_obra": [], "equipo": [], "basicos": []}
    current = None
    for i in range(start, min(end, len(rows))):
        row = rows[i]
        t0 = _clean_text(row[0] if len(row) > 0 else "")
        t1 = _clean_text(row[1] if len(row) > 1 else "")
        has_values = _step10_row_has_numeric_matrix_values(row)
        sec = (_section_kind(t0) or _section_kind(t1)) if not has_values else None
        if sec:
            current = sec
            continue
        if current is None:
            continue
        txt = _row_text(row)
        if any(p in txt for p in ["costo directo", "indirect", "utilidad", "precio unitario", "subtotal1", "subtotal2", "financiamiento"]):
            continue
        item = _parse_component_row(row)
        if item:
            components[current].append(item)
    return components


def _step09_percent_base_kind(item):
    # Regla acordada: todos los conceptos porcentuales del contratista se calculan con base en subtotal de mano de obra de mercado.
    return 'mano_obra'


def _step10_is_percent_like(item):
    code = _normalize_text(item.get('codigo'))
    desc = _normalize_text(item.get('descripcion'))
    unit = normalize_unit(item.get('unidad'))
    text = f"{code} {desc}"
    if _step08_is_percent_item(item):
        return True
    percent_aliases = ['epp', 'she', 'seguridad', 'herramienta menor', 'transporte diario', 'equipo de proteccion personal']
    return unit == '%' or any(a in text for a in percent_aliases)


def _step09_market_importe_for_item(item, market_price, subtotals):
    qty = item.get('factor') if isinstance(item.get('factor'), (int, float)) else item.get('cantidad')
    op = item.get('op') or item.get('operacion') or '*'
    if _step10_is_percent_like(item):
        base = subtotals.get('mano_obra')
        if isinstance(base, (int, float)) and isinstance(qty, (int, float)):
            return float(base) * float(qty), base
        provider_base = item.get('precio_base')
        if isinstance(provider_base, (int, float)) and isinstance(qty, (int, float)):
            return float(provider_base) * float(qty), provider_base
        return item.get('importe'), provider_base
    return _step09_apply_operation(market_price, qty, op), market_price


_STEP10_APPLY_MARKET_CATALOG_PRICING_PREV = _step08_apply_market_catalog_pricing

def _step08_apply_market_catalog_pricing(conceptos):
    conceptos = _STEP10_APPLY_MARKET_CATALOG_PRICING_PREV(conceptos)
    for _, c in (conceptos or {}).items():
        for item in c.get('granular_market_items') or []:
            if _step10_is_percent_like(item) and isinstance(item.get('base_mercado'), (int, float)):
                item['is_percent_item'] = True
                item['percent_base_kind'] = 'mano_obra'
                item['precio_mercado_catalogo'] = item.get('precio_mercado_catalogo', item.get('precio_mercado'))
                item['precio_mercado'] = item.get('base_mercado')
                prov = item.get('precio_proveedor')
                mkt = item.get('precio_mercado')
                item['delta_precio_pct'] = ((float(prov) - float(mkt)) / float(mkt)) if isinstance(prov,(int,float)) and isinstance(mkt,(int,float)) and mkt else None
                reason = item.get('match_reason') or ''
                item['match_reason'] = (reason + ' | ' if reason else '') + 'porcentual calculado sobre subtotal MO de mercado'
    return conceptos

# Paso 10b - hacer que el primer pase de cálculo reconozca EPP/SHE y otros porcentuales antes de calcular subtotales.
def _step08_is_percent_item(item):
    code_raw = str(item.get('codigo') or '').strip()
    unit_raw = str(item.get('unidad') or '').strip()
    desc_raw = str(item.get('descripcion') or '').strip()
    code = code_raw.upper()
    unit = normalize_unit(unit_raw)
    desc = _normalize_text(desc_raw)
    text = _normalize_text(f"{code_raw} {desc_raw}")
    if code.startswith('%') or unit_raw == '%' or unit == '%' or 'porcentaje' in desc:
        return True
    aliases = ['epp', 'she', 'equipo de proteccion personal', 'herramienta menor', 'transporte diario', 'seguridad']
    return any(a in text for a in aliases)

# -----------------------------------------------------------------------------
# Paso 11 - Recuperar matriz base CD + ajuste fino de porcentuales MO + catalogo MO v3
# -----------------------------------------------------------------------------
# Contexto:
# - En Paso 09 se omitio cargar construdata_matrices.xlsx para acelerar la arquitectura
#   granular. Eso dejo el tab "Matriz base CD" sin matches por concepto.
# - Para Paso 11 se restaura la carga de la base de matrices para el tercer tab, sin
#   afectar que el mercado principal siga calculandose por insumo granular.
# - Se corrige la deteccion de porcentuales para no clasificar como porcentaje a MO real
#   como SUPERVISOR DE SEGURIDAD solo por contener la palabra "seguridad".
# - Todos los porcentuales (%MO1, %MO2, %MO3, %MO5, EPP, SHE, unidad %) se calculan sobre
#   la suma completa de mano de obra de mercado del servicio.

STEP08_CATALOG_FILES['mano_obra'] = 'construdata-manodeobra-052026-3.xlsx'

# Restaurar carga real de construdata_matrices.xlsx para recuperar homologacion concepto -> matriz.
# _EXTRACT_PRECIOS_NACIONAL_STEP09_PREV apunta a la version anterior al skip de Paso 09.
def extract_precios_nacional(filepath: str) -> dict:
    return _EXTRACT_PRECIOS_NACIONAL_STEP09_PREV(filepath)


def _step11_text(item):
    return _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')} {item.get('unidad','')}")


def _step11_is_labor_role_item(item):
    text = _step11_text(item)
    role_tokens = [
        'supervisor', 'segurista', 'residente', 'cabo', 'oficial', 'albanil', 'albañil',
        'ayudante', 'peon', 'peón', 'operador', 'soldador', 'pintor', 'fierrero',
        'carpintero', 'electricista', 'plomero', 'maestro', 'maniobrista', 'topografo',
        'topógrafo', 'chofer', 'tecnico', 'técnico'
    ]
    return any(tok in text for tok in role_tokens)


def _step08_is_percent_item(item):
    """Detecta porcentuales reales sin convertir MO humana en porcentaje.

    Casos SI porcentuales:
      - Codigo %MO1/%MO2/%MO3/%MO5/%...
      - Unidad %
      - Descripcion contiene porcentaje explícito
      - EPP / SHE / herramienta menor / equipo de proteccion personal cuando vienen como cargo porcentual
    Casos NO porcentuales:
      - SUPERVISOR DE SEGURIDAD u otros roles humanos, aunque contengan "seguridad".
      - Transporte/logistica si no viene con unidad % o codigo porcentual.
    """
    code_raw = str(item.get('codigo') or '').strip()
    unit_raw = str(item.get('unidad') or '').strip()
    desc_raw = str(item.get('descripcion') or '').strip()
    code = code_raw.upper()
    unit = normalize_unit(unit_raw)
    text = _normalize_text(f"{code_raw} {desc_raw}")

    if _step11_is_labor_role_item(item):
        return False
    if code.startswith('%') or code.startswith('MO%') or code.startswith('%MO'):
        return True
    if unit_raw == '%' or unit == '%':
        return True
    if 'porcentaje' in text or re.search(r'\b\d+(\.\d+)?\s*%\b', text):
        return True
    percent_aliases = [
        'herramienta menor',
        'equipo de proteccion personal', 'equipo de protección personal',
        'epp', 'she',
        'seguridad e higiene', 'seguridad industrial'
    ]
    if any(a in text for a in percent_aliases):
        return True
    return False


def _step10_is_percent_like(item):
    return _step08_is_percent_item(item)


def _step09_percent_base_kind(item):
    # Regla acordada: todo cargo porcentual del contratista se calcula sobre subtotal de MO de mercado.
    return 'mano_obra'


def _step11_operation_amount(price, qty, op):
    if not isinstance(price, (int, float)) or not isinstance(qty, (int, float)):
        return None
    opn = str(op or '*').strip()
    if opn == '/':
        return None if qty == 0 else float(price) / float(qty)
    return float(price) * float(qty)


def _step09_market_importe_for_item(item, market_price, subtotals):
    qty = item.get('factor') if isinstance(item.get('factor'), (int, float)) else item.get('cantidad')
    op = item.get('op') or item.get('operacion') or '*'
    if _step08_is_percent_item(item):
        base = subtotals.get('mano_obra')
        if isinstance(base, (int, float)) and isinstance(qty, (int, float)):
            return float(base) * float(qty), base
        provider_base = item.get('precio_base')
        if isinstance(provider_base, (int, float)) and isinstance(qty, (int, float)):
            return float(provider_base) * float(qty), provider_base
        return item.get('importe'), provider_base
    return _step11_operation_amount(market_price, qty, op), market_price


# Recalculo final del mercado granular usando los nuevos criterios de porcentuales y MO v3.
def _step08_apply_market_catalog_pricing(conceptos):
    try:
        catalogs = load_step08_market_catalogs(force=True)
    except Exception as exc:
        for c in (conceptos or {}).values():
            c['granular_market_items'] = []
            c['granular_market_error'] = str(exc)
        return conceptos

    for _, c in (conceptos or {}).items():
        raw_items = list(_step08_iter_provider_items(c))
        first_pass = []
        for item in raw_items:
            match = _step08_match_market_item(item, catalogs)
            selected = match.get('selected') or {}
            provider_price = item.get('precio_base')
            market_price = selected.get('costo') if isinstance(selected.get('costo'), (int, float)) else provider_price
            fallback_used = not bool(selected)
            first_pass.append((item, match, selected, market_price, fallback_used))

        # La base de porcentuales es la suma de TODOS los renglones de mano de obra de mercado
        # del servicio, excluyendo solo cargos porcentuales. Incluye supervisores, seguristas,
        # operadores y cualquier MO declarada con costo/cantidad real.
        subtotals = {'material': 0.0, 'mano_obra': 0.0, 'equipo': 0.0}
        for item, match, selected, market_price, fallback_used in first_pass:
            if _step08_is_percent_item(item):
                continue
            imp = _step11_operation_amount(market_price, item.get('factor'), item.get('op') or item.get('operacion') or '*')
            domain = item.get('_domain') or 'material'
            if isinstance(imp, (int, float)):
                subtotals[domain] = subtotals.get(domain, 0.0) + float(imp)

        enriched = []
        direct_market = 0.0
        provider_direct = 0.0
        matched_count = 0
        total_count = 0
        for item, match, selected, market_price, fallback_used in first_pass:
            total_count += 1
            provider_price = item.get('precio_base')
            provider_importe = item.get('importe')
            market_importe, market_base = _step09_market_importe_for_item(item, market_price, subtotals)
            is_percent = _step08_is_percent_item(item)
            display_market_price = market_base if is_percent and isinstance(market_base, (int, float)) else market_price
            delta_price = None
            if isinstance(provider_price, (int, float)) and isinstance(display_market_price, (int, float)) and display_market_price:
                delta_price = (float(provider_price) - float(display_market_price)) / float(display_market_price)
            if isinstance(market_importe, (int, float)):
                direct_market += float(market_importe)
            if isinstance(provider_importe, (int, float)):
                provider_direct += float(provider_importe)
            if selected:
                matched_count += 1
            status = match.get('status') if selected else 'sin_match_fallback_contratista'
            reason = selected.get('match_reason') if selected else 'sin match en catalogo; costo mercado igual al costo contratista como fallback'
            if is_percent:
                reason = (reason + ' | ' if reason else '') + 'porcentual calculado sobre suma total MO de mercado del servicio'
            enriched.append({
                'section': item.get('_section'),
                'domain': item.get('_domain'),
                'codigo': item.get('codigo'),
                'descripcion': item.get('descripcion'),
                'unidad': item.get('unidad'),
                'cantidad': item.get('factor'),
                'op': item.get('op') or item.get('operacion') or '*',
                'precio_proveedor': provider_price,
                'importe_proveedor': provider_importe,
                'codigo_mercado': selected.get('codigo'),
                'descripcion_mercado': selected.get('descripcion'),
                'unidad_mercado': selected.get('unidad') if selected else item.get('unidad'),
                'precio_mercado': display_market_price,
                'precio_mercado_catalogo': market_price,
                'importe_mercado': market_importe,
                'base_mercado': market_base if is_percent else None,
                'delta_precio_pct': delta_price,
                'match_status': status,
                'confidence': match.get('confidence') if selected else 0.0,
                'match_reason': reason,
                'source_file': selected.get('source_file'),
                'source_row': selected.get('source_row'),
                'is_percent_item': is_percent,
                'percent_base_kind': 'mano_obra' if is_percent else None,
                'fallback_contratista': fallback_used,
                'top_candidates': match.get('candidates') or [],
            })
        c['granular_market_items'] = enriched
        c['granular_market_direct'] = round(direct_market, 6) if direct_market else None
        c['granular_provider_direct'] = round(provider_direct, 6) if provider_direct else None
        c['granular_market_match_coverage'] = (matched_count / total_count) if total_count else None
        c['granular_market_subtotals'] = {k: round(v, 6) for k, v in subtotals.items()}
    return conceptos


def _step11_ai_commentary_for_concept(concepto):
    """Comentario humano deterministico; si Claude esta disponible, la homologacion previa ya trae razon IA."""
    comp = concepto.get('construdata_matrix_comparison') or {}
    match = concepto.get('construdata_concept_match') or {}
    coverage = concepto.get('granular_market_match_coverage')
    market_direct = concepto.get('granular_market_direct')
    provider_direct = concepto.get('granular_provider_direct')
    parts = []
    if match:
        status = match.get('status') or 'sin_match'
        score = match.get('score')
        codes = []
        for sm in match.get('selected_matches') or []:
            cd = sm.get('concept') or {}
            if cd.get('codigo') or cd.get('clave'):
                codes.append(cd.get('codigo') or cd.get('clave'))
        if codes:
            parts.append(f"Match matriz: {', '.join(codes)} ({status}, score {score:.2f}).") if isinstance(score, (int,float)) else parts.append(f"Match matriz: {', '.join(codes)} ({status}).")
        else:
            parts.append(f"Sin match confiable de matriz Construdata ({status}).")
    if isinstance(coverage, (int,float)):
        parts.append(f"Cobertura de precios unitarios de mercado: {coverage*100:.1f}%.")
    if isinstance(provider_direct, (int,float)) and isinstance(market_direct, (int,float)) and market_direct:
        diff = (provider_direct - market_direct) / market_direct
        parts.append(f"Costo directo declarado vs mercado granular: {diff*100:.1f}%.")
    findings = comp.get('findings') or []
    if findings:
        parts.append('Hallazgos matriz: ' + '; '.join((h.get('message') or '') for h in findings[:3] if h.get('message')))
    if not parts:
        parts.append('Analisis generado con matching granular de insumos; sin suficientes señales para comentario adicional.')
    return ' '.join(parts)


# Override del reporte single-provider para agregar comentarios humanos en Matriz base CD sin alterar los 2 primeros tabs.
_STEP11_BUILD_SINGLE_PROVIDER_PREV = _build_single_provider_pmd

def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP11_BUILD_SINGLE_PROVIDER_PREV(output, meta, dato, provider_analysis)
    try:
        wb = openpyxl.load_workbook(output)
        # Asegurar que Matriz base CD incluya una columna de comentario analitico si existe.
        if 'Matriz base CD' in wb.sheetnames:
            ws = wb['Matriz base CD']
            max_col = ws.max_column + 1
            ws.cell(1, max_col, 'Comentario IA / Analisis humano')
            ws.cell(1, max_col).fill = _fill(C_AZUL_OSC)
            ws.cell(1, max_col).font = _fnt(bold=True, color=C_BLANCO, size=8)
            ws.column_dimensions[get_column_letter(max_col)].width = 70
            # Mapear por servicio si la primera columna contiene servicio.
            comments = {}
            for clave, c in (dato.get('conceptos') or {}).items():
                comments[str(clave)] = _step11_ai_commentary_for_concept(c)
            for r in range(2, ws.max_row + 1):
                service = str(ws.cell(r, 1).value or '').split('|')[0].replace('Análisis:', '').strip()
                if service in comments:
                    ws.cell(r, max_col, comments[service])
                    ws.cell(r, max_col).alignment = _aln('left','top', wrap=True)
        wb.save(output)
    except Exception:
        pass
    return result

# -----------------------------------------------------------------------------
# Paso 12 - IA activa para insumos y concepto/matriz + Excel ajustado
# -----------------------------------------------------------------------------
# Cambios acumulados:
# 1) La homologacion granular de materiales/MO/equipo usa IA cuando existe
#    ANTHROPIC_API_KEY y/o VOYAGE_API_KEY. La IA NO calcula precios: decide entre
#    candidatos del catalogo Construdata actual y deja trazabilidad.
# 2) Se fuerza nuevamente la homologacion concepto -> matriz CD para alimentar
#    el tab 3. Si Claude/Voyage estan configurados, se usan dentro de
#    match_concepto_construdata(); si no, queda fallback por reglas con umbral
#    bajo pero marcado como revision.
# 3) Comparativa se ordena de mayor a menor costo y se colorea azul el grupo de
#    servicios que acumula el 80% o el tramo mas cercano.
# 4) Detalle conserva la matriz del contratista y agrega renglones de Costo
#    directo, Indirecto, Financiamiento, Utilidad y Total PU para contratista
#    vs mercado. Mercado = costo directo granular + 25% indirecto.

try:
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    def get_column_letter(idx):
        letters = ''
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

try:
    from material_ai_matcher import get_material_ai_status
except Exception:  # pragma: no cover
    def get_material_ai_status():
        return {'active': False, 'mode': 'disabled'}


def _step12_ai_enabled_for_materials():
    return os.getenv('ENABLE_AI_MATERIAL_MATCHING', '1') != '0'


def _step12_to_ai_candidate(c):
    return {
        'clave': c.get('codigo'),
        'codigo': c.get('codigo'),
        'descripcion': c.get('descripcion'),
        'descripcion_corta': c.get('descripcion'),
        'unidad': c.get('unidad'),
        'precio': c.get('costo'),
        'tipo_insumo': c.get('tipo_insumo'),
        'tipo_kind': c.get('domain'),
        'score': c.get('score'),
        'fuente_precio': c.get('source_file') or 'catalogo_construdata_granular',
        'source_row': c.get('source_row'),
        'tokens': c.get('tokens') or [],
    }


def _step12_from_ai_evidence(evidence, candidates):
    if not evidence:
        return None
    ev_code = _step08_norm_code(evidence.get('codigo') or evidence.get('clave'))
    ev_desc = _normalize_text(evidence.get('descripcion'))
    ev_unit = normalize_unit(evidence.get('unidad'))
    for c in candidates or []:
        if ev_code and ev_code == _step08_norm_code(c.get('codigo')):
            return c
    best = None
    best_score = 0.0
    for c in candidates or []:
        score = 0.0
        if ev_desc and ev_desc == _normalize_text(c.get('descripcion')):
            score += 0.7
        if ev_unit and ev_unit == normalize_unit(c.get('unidad')):
            score += 0.2
        if score > best_score:
            best, best_score = c, score
    return best if best_score >= 0.7 else None


_STEP12_MATCH_MARKET_ITEM_PREV = _step08_match_market_item


def _step08_match_market_item(item, catalogs):
    """Match granular de insumo con IA opcional.

    Cascada:
    - reglas/codigo/tokens por dominio
    - Voyage sobre candidatos reducidos y Claude para elegir candidato, si hay API keys
    - fallback deterministico si IA no esta configurada o no confirma candidato
    """
    base = _STEP12_MATCH_MARKET_ITEM_PREV(item, catalogs)
    candidates = list(base.get('candidates') or [])
    selected = base.get('selected')
    if not _step12_ai_enabled_for_materials() or not candidates:
        return base

    try:
        ai_candidates = [_step12_to_ai_candidate(c) for c in candidates[:8] if isinstance(c.get('costo'), (int, float))]
        if not ai_candidates:
            return base
        ai_item = {
            'codigo': item.get('codigo'),
            'descripcion': item.get('descripcion'),
            'unidad': item.get('unidad'),
            'precio_base': item.get('precio_base'),
        }
        resolution = select_material_market_price(ai_item, {'top_candidates': ai_candidates, 'selected': ai_candidates[0]})
        evidence = (resolution or {}).get('evidence') or {}
        confidence = float((resolution or {}).get('confidence_pct') or 0.0) / 100.0
        ai_selected = _step12_from_ai_evidence(evidence, candidates)
        # Solo reemplaza si IA encontro candidato comparable. Si no, conserva reglas.
        if ai_selected and confidence >= float(os.getenv('AI_MATERIAL_MIN_CONFIDENCE', '0.60')):
            ai_selected = dict(ai_selected)
            ai_selected['score'] = max(float(ai_selected.get('score') or 0), confidence)
            ai_selected['match_reason'] = (resolution or {}).get('decision_criterion') or ai_selected.get('match_reason') or 'validado por IA'
            ai_selected['ai_status'] = (resolution or {}).get('ai_status')
            ai_selected['voyage_used'] = bool((resolution or {}).get('voyage_used'))
            ai_selected['claude_used'] = bool((resolution or {}).get('claude_used'))
            ai_selected['confidence_pct_ai'] = (resolution or {}).get('confidence_pct')
            return {
                'selected': ai_selected,
                'candidates': candidates,
                'status': 'match_ia' if ai_selected.get('claude_used') or ai_selected.get('voyage_used') else 'match',
                'confidence': ai_selected.get('score'),
                'ai_resolution': resolution,
            }
        # Mantener metadatos IA en top candidate aunque no reemplace.
        if selected and isinstance(selected, dict):
            selected = dict(selected)
            selected['ai_status'] = (resolution or {}).get('ai_status')
            selected['match_reason'] = (selected.get('match_reason') or '') + ' | IA no reemplazo candidato por confianza insuficiente'
            base['selected'] = selected
            base['ai_resolution'] = resolution
    except Exception as exc:
        # La IA nunca debe romper el calculo deterministico.
        base['ai_error'] = type(exc).__name__ + ': ' + str(exc)[:250]
    return base


_STEP12_MATCH_CONCEPTO_PREV = match_concepto_construdata


def match_concepto_construdata(provider_key, provider_concept, construdata_concepts, min_score=0.26, concept_index=None):
    """Paso 12: mantiene IA conceptual y baja el umbral de candidato revisable.

    El tab Matriz base CD no debe quedar vacio si existe un candidato tecnico
    razonable; cuando el score es bajo se marca como revision, no como match fuerte.
    """
    result = _STEP12_MATCH_CONCEPTO_PREV(provider_key, provider_concept, construdata_concepts, min_score=min_score, concept_index=concept_index)
    if (result or {}).get('selected_matches'):
        return result
    candidates = (result or {}).get('candidates') or []
    best = candidates[0] if candidates else None
    if best and best.get('_concept') and float(best.get('score') or 0) >= min_score:
        cd = best.get('_concept')
        return {
            'status': 'match_debil_revision',
            'tipo_match': 'match_parcial',
            'selected': cd,
            'selected_matches': [{'concept': cd, 'peso': 1.0, 'motivo': 'mejor candidato por reglas; requiere revision'}],
            'score': float(best.get('score') or 0),
            'reason': 'Candidato recuperado para tab Matriz base CD; no usar como decision automatica sin revision.',
            'requires_review': True,
            'candidates': candidates[:10],
            'ai_status': get_concept_ai_status(),
        }
    return result


def _step12_ensure_matrix_base_match(conceptos, precios_nacional):
    """Reintenta homologacion de matriz si el tab 3 quedo sin filas."""
    cds = _get_construdata_concepts(precios_nacional)
    if not cds:
        return conceptos
    index = _build_construdata_token_index(cds)
    for key, c in (conceptos or {}).items():
        comp = c.get('construdata_matrix_comparison') or {}
        if comp.get('rows'):
            continue
        match = match_concepto_construdata(key, c, cds, min_score=float(os.getenv('CONCEPT_MATRIX_MIN_SCORE', '0.26')), concept_index=index)
        comparison = compare_provider_matrix_vs_construdata(key, c, match)
        c['construdata_concept_match'] = match
        c['construdata_matrix_comparison'] = comparison
        c['construdata_matrix_findings'] = comparison.get('findings') or []
    return conceptos


_STEP12_ATTACH_PREV = attach_construdata_matrix_analysis


def attach_construdata_matrix_analysis(conceptos, precios_nacional):
    conceptos = _STEP12_ATTACH_PREV(conceptos, precios_nacional)
    return _step12_ensure_matrix_base_match(conceptos, precios_nacional)


def _step12_market_direct(c):
    val = c.get('granular_market_direct')
    return float(val) if isinstance(val, (int, float)) else 0.0


def _step12_provider_direct(c):
    val = c.get('costo_directo')
    if isinstance(val, (int, float)):
        return float(val)
    val = c.get('granular_provider_direct')
    return float(val) if isinstance(val, (int, float)) else 0.0


def _step12_provider_indirect(c):
    amt, pct = _concept_indirect_amount_and_pct(c)
    return float(amt) if isinstance(amt, (int, float)) else None, pct


def _step12_selected_80_rows(conceptos_items):
    total = sum(float(c.get('total') or 0) for _, c in conceptos_items if isinstance(c.get('total'), (int, float)))
    selected = set()
    acc = 0.0
    for clave, c in conceptos_items:
        if total <= 0:
            break
        if acc < 0.80 or not selected:
            selected.add(clave)
            acc += float(c.get('total') or 0) / total
    return selected, acc


def _step12_write_financial_summary_rows(ws_det, start_row, c):
    """Agrega al Detalle los conceptos financieros que forman el PU."""
    market_direct = _step12_market_direct(c)
    market_indirect = market_direct * MARKET_INDIRECT_PCT if market_direct else 0.0
    market_total = market_direct + market_indirect if market_direct else None
    provider_direct = _step12_provider_direct(c)
    provider_indirect, provider_indirect_pct = _step12_provider_indirect(c)
    utilidad = c.get('utilidad') if isinstance(c.get('utilidad'), (int, float)) else 0.0
    financiamiento = c.get('financiamiento') if isinstance(c.get('financiamiento'), (int, float)) else 0.0
    pu_provider = c.get('pu') if isinstance(c.get('pu'), (int, float)) else None
    rows = [
        ('COSTO DIRECTO', provider_direct, market_direct, 'Suma de matriz contratista vs suma con precios mercado.'),
        ('COSTO INDIRECTO', provider_indirect, market_indirect, f'Mercado calculado siempre con {MARKET_INDIRECT_PCT*100:.0f}% sobre costo directo.'),
        ('FINANCIAMIENTO', financiamiento, 0.0, 'Se muestra lo declarado por contratista; mercado granular no agrega financiamiento.'),
        ('UTILIDAD / CARGOS ADICIONALES', utilidad, 0.0, 'Se muestra lo declarado por contratista si fue detectado.'),
        ('TOTAL COSTO UNITARIO', pu_provider, market_total, 'Total contratista vs mercado granular + indirecto 25%.'),
    ]
    row = start_row
    for label, prov, market, note in rows:
        vals = ['', label, '', prov, '', '', prov, market, None, market, '', 'RESUMEN FINANCIERO', '', None, '', note]
        if isinstance(prov, (int, float)) and isinstance(market, (int, float)) and market:
            vals[8] = (float(prov) - float(market)) / float(market)
        for col, val in enumerate(vals, 1):
            cell = ws_det.cell(row, col, val)
            cell.border = _brd(); cell.font = _fnt(bold=True if label == 'TOTAL COSTO UNITARIO' else False, size=8)
            cell.alignment = _aln('left','center', wrap=True) if col in (2,16) else _aln('center','center')
            if col in (4,7,8,10) and isinstance(val, (int, float)): cell.number_format = '$#,##0.00'
            if col == 9 and isinstance(val, (int, float)): cell.number_format = '0.00%'
            if col in (8,9): cell.fill = _fill(C_AMARILLO_PMD)
            if label == 'TOTAL COSTO UNITARIO': cell.fill = _fill('D9EAF7') if col not in (8,9) else _fill(C_AMARILLO_PMD)
        row += 1
    return row


# Writer final Paso 12: tres tabs principales, Comparativa ordenada 80/20, Detalle completo, Matriz base CD.
def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comparativa'
    ws_det = wb.create_sheet('Detalle')
    conceptos_items = list((dato.get('conceptos') or {}).items())
    conceptos_items.sort(key=lambda kv: float(kv[1].get('total') or 0), reverse=True)
    total_contratista = round(sum(v.get('total') or 0 for _, v in conceptos_items if isinstance(v.get('total'), (int, float))), 2)
    blue_keys, blue_pct = _step12_selected_80_rows(conceptos_items)

    for col, width in {'A':12,'B':58,'C':10,'D':12,'E':13,'F':14,'G':14,'H':14,'I':12,'J':14,'K':14,'L':28}.items():
        ws.column_dimensions[col].width = width
    ws.merge_cells('A1:L1')
    ws['A1'] = f"COMPARATIVO DE COTIZACIÓN — {meta.get('proyecto','')}"
    ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10); ws['A1'].alignment = _aln('center','center')
    ws['A2'] = 'Criterio azul'; ws['B2'] = f'Servicios ordenados por costo; azul = acumulado {blue_pct*100:.2f}% de la propuesta.'
    ws['A2'].font = _fnt(bold=True, size=8); ws['B2'].font = _fnt(size=8)
    headers = ['Servicio','Descripción','Unidad','Cantidad','PU Contratista','Importe Contratista','PU Mercado +25%','Importe Mercado','Dif %','Cobertura mercado','Estado','Nota']
    for col,h in enumerate(headers,1):
        cell=ws.cell(4,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (7,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (7,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    row=5; total_market=0.0
    for clave,c in conceptos_items:
        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'),(int,float)) and c.get('cantidad') else 1
        pu = c.get('pu') if isinstance(c.get('pu'),(int,float)) else None
        importe = _concept_total_amount(c)
        market_pu = _market_pu_with_indirect(c)
        market_total = market_pu*cantidad if isinstance(market_pu,(int,float)) else None
        if isinstance(market_total,(int,float)): total_market += market_total
        diff_pct = ((importe-market_total)/market_total) if isinstance(importe,(int,float)) and isinstance(market_total,(int,float)) and market_total else None
        nota = 'Mercado por insumos declarados + 25% indirecto. IA activa si hay Claude/Voyage configurados.'
        vals=[clave,c.get('desc'),c.get('unidad'),cantidad,pu,importe,market_pu,market_total,diff_pct,c.get('granular_market_match_coverage'),_step08_market_status(diff_pct),nota]
        is_blue = clave in blue_keys
        for col,val in enumerate(vals,1):
            cell=ws.cell(row,col,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True) if col in (2,12) else _aln('center','center')
            if col in (5,6,7,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
            if col in (9,10) and isinstance(val,(int,float)): cell.number_format='0.00%'
            if is_blue: cell.fill = _fill('D9EAF7')
            if col in (7,9): cell.fill=_fill(C_AMARILLO_PMD)
        row+=1
    row+=1
    for col,val in [(1,'TOTAL'),(6,total_contratista),(8,total_market),(9,((total_contratista-total_market)/total_market if total_market else None))]:
        cell=ws.cell(row,col,val); cell.font=_fnt(bold=True,size=8); cell.border=_brd()
        if col in (6,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
        if col==9 and isinstance(val,(int,float)): cell.number_format='0.00%'

    # Detalle completo
    widths={'A':15,'B':58,'C':10,'D':13,'E':7,'F':10,'G':14,'H':16,'I':14,'J':14,'K':14,'L':18,'M':44,'N':12,'O':20,'P':42}
    for col,width in widths.items(): ws_det.column_dimensions[col].width=width
    ws_det.merge_cells('A1:P1'); ws_det['A1']='MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO'; ws_det['A1'].fill=_fill(C_AZUL_OSC); ws_det['A1'].font=_fnt(bold=True,color=C_BLANCO,size=9); ws_det['A1'].alignment=_aln('center','center')
    headers_det=['Código','Concepto / Insumo','Unidad','Costo Contratista','Op.','Cantidad','Importe Contratista','Costo Mercado','Dif % Costo','Importe Mercado','Base Mercado %','Tipo','Match Construdata','Conf.','Estado','Nota']
    for col,h in enumerate(headers_det,1):
        cell=ws_det.cell(3,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (8,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (8,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    det_row=4
    for clave,c in conceptos_items:
        market_pu = _market_pu_with_indirect(c)
        title = f'Análisis: {clave} | {c.get("desc","")} | PU Mercado +25%: {market_pu:,.2f}' if isinstance(market_pu,(int,float)) else f'Análisis: {clave} | {c.get("desc","")}'
        ws_det.cell(det_row,1,title).fill=_fill(C_GRIS_SEC); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8); ws_det.cell(det_row,1).alignment=_aln('left','center',wrap=True)
        ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
        det_row+=1
        current=None
        for item in c.get('granular_market_items') or []:
            if item.get('section') != current:
                current=item.get('section')
                ws_det.cell(det_row,1,current).fill=_fill(C_GRIS_FIL); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
                ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
                det_row+=1
            diff=item.get('delta_precio_pct')
            if item.get('fallback_contratista'):
                nota='Sin match: costo mercado = costo contratista como fallback.'
            elif item.get('is_percent_item'):
                nota='Porcentual calculado sobre la suma total de MO de mercado del servicio.'
            else:
                nota=item.get('match_reason') or ''
            vals=[item.get('codigo'),item.get('descripcion'),item.get('unidad'),item.get('precio_proveedor'),item.get('op') or '*',item.get('cantidad'),item.get('importe_proveedor'),item.get('precio_mercado'),diff,item.get('importe_mercado'),item.get('base_mercado'),item.get('section'),f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(' -'),item.get('confidence'),item.get('match_status'),nota]
            for col,val in enumerate(vals,1):
                cc=ws_det.cell(det_row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,13,16) else _aln('center','center')
                if col in (4,7,8,10,11) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
                if col in (9,14) and isinstance(val,(int,float)): cc.number_format='0.00%'
                if col in (8,9): cc.fill=_fill(C_AMARILLO_PMD)
            det_row+=1
        det_row = _step12_write_financial_summary_rows(ws_det, det_row, c) + 1

    _step08_build_matrix_base_sheet(wb, dato)
    # Anexar columna de comentario IA/humano al tab 3.
    if 'Matriz base CD' in wb.sheetnames:
        ws3 = wb['Matriz base CD']
        col = ws3.max_column + 1
        ws3.cell(1, col, 'Comentario IA / análisis humano')
        ws3.cell(1, col).fill = _fill(C_AZUL_OSC); ws3.cell(1, col).font = _fnt(bold=True, color=C_BLANCO, size=8)
        ws3.column_dimensions[get_column_letter(col)].width = 70
        comments = {str(k): _step11_ai_commentary_for_concept(v) for k, v in (dato.get('conceptos') or {}).items()}
        for r in range(2, ws3.max_row + 1):
            service = str(ws3.cell(r, 1).value or '').split('|')[0].replace('Análisis:', '').strip()
            if service in comments:
                ws3.cell(r, col, comments[service]); ws3.cell(r, col).alignment = _aln('left','top', wrap=True)
    wb._sheets=[wb[n] for n in ['Comparativa','Detalle','Matriz base CD'] if n in wb.sheetnames]
    wb.save(output)
    return output

# Paso 12b - evitar llamadas externas si no hay Claude/Voyage configurados.
def _step12_ai_enabled_for_materials():
    if os.getenv('ENABLE_AI_MATERIAL_MATCHING', '1') == '0':
        return False
    try:
        st = get_material_ai_status() or {}
        return bool(st.get('anthropic_enabled') or st.get('voyage_enabled'))
    except Exception:
        return bool((os.getenv('ANTHROPIC_API_KEY') or '').strip() or (os.getenv('VOYAGE_API_KEY') or '').strip())

# -----------------------------------------------------------------------------
# Paso 13 - Correcciones solicitadas:
# 1) El 80/20 se calcula por costo, pero NO reordena los servicios en Excel.
#    Los tabs conservan el orden declarado por el contratista.
# 2) Costos financieros porcentuales (indirecto, financiamiento, utilidad si se
#    detecta como porcentaje) se recalculan contra el subtotal mercado previo.
# 3) Matching de CAL HIDRATADA/CALHIDRA: el alias ya no fuerza match automático
#    contra cualquier candidato con token "cal". El alias solo genera candidatos;
#    la selección exige evidencia técnica mínima o validación IA.
# -----------------------------------------------------------------------------


def _step13_declared_order_concept_items(dato):
    """Retorna conceptos en el orden original del archivo del contratista."""
    return list((dato.get('conceptos') or {}).items())


def _step13_selected_80_rows_keep_order(conceptos_items):
    """Selecciona servicios que acumulan 80% usando ranking por costo, sin cambiar el orden de salida."""
    total = sum(float(c.get('total') or 0) for _, c in conceptos_items if isinstance(c.get('total'), (int, float)))
    ranked = sorted(conceptos_items, key=lambda kv: float(kv[1].get('total') or 0), reverse=True)
    selected = set()
    acc = 0.0
    for clave, c in ranked:
        if total <= 0:
            break
        if acc < 0.80 or not selected:
            selected.add(clave)
            acc += float(c.get('total') or 0) / total
    return selected, acc


# Alias tecnico menos agresivo. Importante: "cal" por si solo NO se expande a CALHIDRA.
def _step09_expand_tokens(tokens):
    base = set(tokens or [])
    expanded = set(base)
    conservative_aliases = {
        'calhidra': {'calhidra', 'cal', 'hidratada'},
        'hidratada': {'calhidra', 'cal', 'hidratada'},
        'rafia': {'costal', 'saco', 'rafia'},
        'costal': {'costal', 'saco', 'rafia'},
        'saco': {'saco', 'costal', 'bulto'},
        'rompedora': {'rompedora', 'demoledor', 'rotomartillo', 'martillo'},
        'electrica': {'electrica', 'electrico', 'electric'},
        'electric': {'electrica', 'electrico', 'electric'},
        'peon': {'peon', 'ayudante'},
        'ayudante': {'ayudante', 'peon'},
        'albanil': {'albanil', 'oficial'},
    }
    for tok in list(base):
        # No expandir el token generico "cal" hacia CALHIDRA.
        if tok == 'cal':
            continue
        expanded.update(conservative_aliases.get(tok, set()))
    # Solo si coexisten cal + hidratada, permitir CALHIDRA.
    if {'cal', 'hidratada'}.issubset(base):
        expanded.update({'calhidra', 'cal', 'hidratada'})
    return [t for t in expanded if t]


def _step13_is_cal_hidratada_query(item):
    text = _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')}")
    toks = set(_tokenize_text(text))
    return ('cal hidratada' in text) or ('calhidra' in text) or {'cal', 'hidratada'}.issubset(toks)


def _step13_is_calhidra_candidate(cand):
    text = _normalize_text(f"{cand.get('codigo','')} {cand.get('descripcion','')}")
    toks = set(cand.get('tokens') or _tokenize_text(text))
    return ('calhidra' in text) or ('calhidra' in toks) or ('cal hidratada' in text)


# Guardar referencia al matcher anterior para usar IA/fallback existente.
_STEP13_DIRECT_MATCH_PREV = _step09_direct_catalog_match


def _step09_direct_catalog_match(item, catalogs, domain):
    """Matcher granular corregido.

    Para CAL HIDRATADA, el alias se usa para recuperar candidatos, pero no para
    aceptar matches que solo tienen el token generico CAL. Esto evita falsos
    matches de cal contra conceptos no equivalentes.
    """
    by_code = catalogs.get('by_code') or {}
    code_norm = _step08_norm_code(item.get('codigo'))
    if code_norm and (domain, code_norm) in by_code:
        candidates = [dict(c, score=0.99, match_reason='codigo exacto + dominio') for c in by_code[(domain, code_norm)] if isinstance(c.get('costo'), (int, float))]
        if candidates:
            return candidates[0], candidates

    domain_rows = (catalogs.get('by_domain') or {}).get(domain, [])
    token_index = (catalogs.get('token_index') or {}).get(domain, {})
    item_tokens = set(_step08_item_tokens(item))
    meaningful = [t for t in item_tokens if t not in globals().get('_STEP09_GENERIC_TOKENS', set())]
    if not meaningful:
        meaningful = list(item_tokens)[:4]

    idxs = set()
    for tok in meaningful[:10]:
        idxs.update(token_index.get(tok, set()))

    # Para CAL HIDRATADA agregar candidatos CALHIDRA aunque la tokenizacion no los haya traido.
    item_is_calhidra = _step13_is_cal_hidratada_query(item)
    if item_is_calhidra:
        idxs.update(token_index.get('calhidra', set()))
        idxs.update(token_index.get('hidratada', set()))

    prelim = []
    item_is_costal = 'costal' in _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')}") or 'rafia' in _normalize_text(item.get('descripcion'))
    for idx in list(idxs)[:1200]:
        if idx >= len(domain_rows):
            continue
        cand = domain_rows[idx]
        if not isinstance(cand.get('costo'), (int, float)):
            continue
        cn = _normalize_text(cand.get('descripcion'))
        ct = set(cand.get('tokens') or [])
        if _step09_is_cuadrilla(item) and 'cuadrilla' not in cn:
            continue
        sc = _step08_score_catalog_candidate(item, cand)

        if item_is_calhidra:
            if _step13_is_calhidra_candidate(cand):
                sc = max(sc, 0.88 if normalize_unit(item.get('unidad')) == cand.get('unidad_norm') else 0.74)
            else:
                # Penaliza cualquier match que venga solo por token generico cal.
                sc = min(sc, 0.30)

        if item_is_costal:
            if 'costal' in cn or 'rafia' in cn:
                sc = max(sc, 0.88 if normalize_unit(item.get('unidad')) == cand.get('unidad_norm') else 0.76)
            else:
                sc = min(sc, 0.29)

        if _step09_is_cuadrilla(item) and 'cuadrilla' in cn:
            sc = max(sc, 0.65)

        if sc >= 0.18:
            c = dict(cand)
            c['score'] = round(sc, 4)
            c['match_reason'] = 'similitud descripcion/unidad dentro del mismo dominio'
            if item_is_calhidra and _step13_is_calhidra_candidate(cand):
                c['match_reason'] = 'alias tecnico controlado CAL HIDRATADA -> CALHIDRA; requiere evidencia por codigo/descripcion/unidad'
            elif item_is_calhidra:
                c['match_reason'] = 'candidato bajo: contiene tokens parciales pero no confirma CALHIDRA'
            if item_is_costal and ('costal' in cn or 'rafia' in cn):
                c['match_reason'] = 'alias tecnico COSTAL/RAFIA'
            prelim.append(c)

    prelim.sort(key=lambda x: x.get('score') or 0, reverse=True)
    selected = prelim[0] if prelim and (prelim[0].get('score') or 0) >= 0.34 else None

    # Si la consulta es CAL HIDRATADA, no aceptar automaticamente un candidato que no sea CALHIDRA.
    if item_is_calhidra and selected and not _step13_is_calhidra_candidate(selected):
        selected = None
    return selected, prelim[:8]


def _step13_financial_rows(c):
    """Calcula renglones financieros encadenados con los porcentajes declarados por el contratista.

    Mercado:
    - costo directo mercado = suma granular previa
    - indirecto mercado = % indirecto declarado * costo directo mercado
    - financiamiento mercado = % financiamiento declarado * subtotal mercado previo
    - utilidad mercado = % utilidad declarado * subtotal mercado previo, si puede deducirse
    """
    provider_direct = _step12_provider_direct(c)
    market_direct = _step12_market_direct(c)
    pu_provider = c.get('pu') if isinstance(c.get('pu'), (int, float)) else None

    provider_indirect, provider_indirect_pct = _step12_provider_indirect(c)
    if not isinstance(provider_indirect_pct, (int, float)) and provider_direct and isinstance(provider_indirect, (int, float)):
        provider_indirect_pct = provider_indirect / provider_direct if provider_direct else None
    if not isinstance(provider_indirect_pct, (int, float)):
        provider_indirect_pct = MARKET_INDIRECT_PCT

    market_indirect = market_direct * provider_indirect_pct if market_direct else 0.0

    provider_after_indirect = provider_direct + (provider_indirect or 0.0)
    market_after_indirect = market_direct + market_indirect

    financiamiento = c.get('financiamiento') if isinstance(c.get('financiamiento'), (int, float)) else 0.0
    financing_pct = (financiamiento / provider_after_indirect) if provider_after_indirect else 0.0
    market_financing = market_after_indirect * financing_pct if market_after_indirect else 0.0

    provider_after_financing = provider_after_indirect + financiamiento
    market_after_financing = market_after_indirect + market_financing

    utilidad = c.get('utilidad') if isinstance(c.get('utilidad'), (int, float)) else 0.0
    utility_pct = (utilidad / provider_after_financing) if provider_after_financing else 0.0
    market_utility = market_after_financing * utility_pct if market_after_financing else 0.0

    market_total = market_after_financing + market_utility
    return [
        ('COSTO DIRECTO', provider_direct, market_direct, None, 'Suma de operaciones previas de matriz: contratista vs precios mercado.'),
        ('COSTO INDIRECTO', provider_indirect, market_indirect, provider_indirect_pct, 'Mercado usa el % declarado por contratista sobre costo directo mercado.'),
        ('FINANCIAMIENTO', financiamiento, market_financing, financing_pct, 'Mercado usa el % declarado por contratista sobre subtotal mercado previo.'),
        ('UTILIDAD / CARGOS ADICIONALES', utilidad, market_utility, utility_pct, 'Mercado usa el % declarado por contratista sobre subtotal mercado previo, si fue detectado.'),
        ('TOTAL COSTO UNITARIO', pu_provider, market_total, None, 'Total contratista vs mercado recalculado con los mismos porcentajes declarados.'),
    ]


def _step13_market_pu(c):
    rows = _step13_financial_rows(c)
    total = rows[-1][2]
    return total if isinstance(total, (int, float)) and total > 0 else _market_pu_with_indirect(c)


# Override para que los totales de mercado usen porcentajes declarados encadenados.
def _market_pu_with_indirect(concepto):
    try:
        return _step13_market_pu(concepto)
    except RecursionError:
        direct = concepto.get('granular_market_direct')
        if isinstance(direct, (int, float)) and direct > 0:
            return direct * (1 + MARKET_INDIRECT_PCT)
        return 0.0
    except Exception:
        direct = concepto.get('granular_market_direct')
        if isinstance(direct, (int, float)) and direct > 0:
            return direct * (1 + MARKET_INDIRECT_PCT)
        return 0.0


def _step12_write_financial_summary_rows(ws_det, start_row, c):
    """Renglones financieros corregidos: mercado usa % declarado sobre subtotal previo."""
    row = start_row
    for label, prov, market, pct, note in _step13_financial_rows(c):
        vals = ['', label, '', prov, '', '', prov, market, None, market, '', 'RESUMEN FINANCIERO', '', pct, '', note]
        if isinstance(prov, (int, float)) and isinstance(market, (int, float)) and market:
            vals[8] = (float(prov) - float(market)) / float(market)
        for col, val in enumerate(vals, 1):
            cell = ws_det.cell(row, col, val)
            cell.border = _brd()
            cell.font = _fnt(bold=True if label == 'TOTAL COSTO UNITARIO' else False, size=8)
            cell.alignment = _aln('left', 'center', wrap=True) if col in (2,16) else _aln('center', 'center')
            if col in (4,7,8,10) and isinstance(val, (int, float)):
                cell.number_format = '$#,##0.00'
            if col in (9,14) and isinstance(val, (int, float)):
                cell.number_format = '0.00%'
            if col in (8,9):
                cell.fill = _fill(C_AMARILLO_PMD)
            if label == 'TOTAL COSTO UNITARIO':
                cell.fill = _fill('D9EAF7') if col not in (8,9) else _fill(C_AMARILLO_PMD)
        row += 1
    return row


# Writer final Paso 13: conserva orden declarado y colorea 80/20 sin reordenar filas.
def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comparativa'
    ws_det = wb.create_sheet('Detalle')
    conceptos_items = _step13_declared_order_concept_items(dato)
    total_contratista = round(sum(v.get('total') or 0 for _, v in conceptos_items if isinstance(v.get('total'), (int, float))), 2)
    blue_keys, blue_pct = _step13_selected_80_rows_keep_order(conceptos_items)

    for col, width in {'A':12,'B':58,'C':10,'D':12,'E':13,'F':14,'G':14,'H':14,'I':12,'J':14,'K':14,'L':28}.items():
        ws.column_dimensions[col].width = width
    ws.merge_cells('A1:L1')
    ws['A1'] = f"COMPARATIVO DE COTIZACION — {meta.get('proyecto','')}"
    ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10); ws['A1'].alignment = _aln('center','center')
    ws['A2'] = 'Criterio azul'
    ws['B2'] = f'Azul = servicios que acumulan {blue_pct*100:.2f}% aprox. del costo al analizarlos de mayor a menor; las filas conservan el orden declarado por contratista.'
    ws['A2'].font = _fnt(bold=True, size=8); ws['B2'].font = _fnt(size=8)
    headers = ['Servicio','Descripcion','Unidad','Cantidad','PU Contratista','Importe Contratista','PU Mercado','Importe Mercado','Dif %','Cobertura mercado','Estado','Nota']
    for col,h in enumerate(headers,1):
        cell=ws.cell(4,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (7,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (7,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    row=5; total_market=0.0
    for clave,c in conceptos_items:
        cantidad = c.get('cantidad') if isinstance(c.get('cantidad'),(int,float)) and c.get('cantidad') else 1
        pu = c.get('pu') if isinstance(c.get('pu'),(int,float)) else None
        importe = _concept_total_amount(c)
        market_pu = _step13_market_pu(c)
        market_total = market_pu*cantidad if isinstance(market_pu,(int,float)) else None
        if isinstance(market_total,(int,float)): total_market += market_total
        diff_pct = ((importe-market_total)/market_total) if isinstance(importe,(int,float)) and isinstance(market_total,(int,float)) and market_total else None
        nota = 'Mercado por insumos declarados + porcentajes financieros del contratista aplicados sobre subtotales mercado previos. IA activa si hay Claude/Voyage configurados.'
        vals=[clave,c.get('desc'),c.get('unidad'),cantidad,pu,importe,market_pu,market_total,diff_pct,c.get('granular_market_match_coverage'),_step08_market_status(diff_pct),nota]
        is_blue = clave in blue_keys
        for col,val in enumerate(vals,1):
            cell=ws.cell(row,col,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True) if col in (2,12) else _aln('center','center')
            if col in (5,6,7,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
            if col in (9,10) and isinstance(val,(int,float)): cell.number_format='0.00%'
            if is_blue: cell.fill = _fill('D9EAF7')
            if col in (7,9): cell.fill=_fill(C_AMARILLO_PMD)
        row+=1
    row+=1
    for col,val in [(1,'TOTAL'),(6,total_contratista),(8,total_market),(9,((total_contratista-total_market)/total_market if total_market else None))]:
        cell=ws.cell(row,col,val); cell.font=_fnt(bold=True,size=8); cell.border=_brd()
        if col in (6,8) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
        if col==9 and isinstance(val,(int,float)): cell.number_format='0.00%'

    widths={'A':15,'B':58,'C':10,'D':13,'E':7,'F':10,'G':14,'H':16,'I':14,'J':14,'K':14,'L':18,'M':44,'N':12,'O':20,'P':42}
    for col,width in widths.items(): ws_det.column_dimensions[col].width=width
    ws_det.merge_cells('A1:P1'); ws_det['A1']='MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO'; ws_det['A1'].fill=_fill(C_AZUL_OSC); ws_det['A1'].font=_fnt(bold=True,color=C_BLANCO,size=9); ws_det['A1'].alignment=_aln('center','center')
    headers_det=['Codigo','Concepto / Insumo','Unidad','Costo Contratista','Op.','Cantidad','Importe Contratista','Costo Mercado','Dif % Costo','Importe Mercado','Base Mercado %','Tipo','Match Construdata','Conf.','Estado','Nota']
    for col,h in enumerate(headers_det,1):
        cell=ws_det.cell(3,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (8,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (8,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    det_row=4
    for clave,c in conceptos_items:
        market_pu = _step13_market_pu(c)
        title = f'Analisis: {clave} | {c.get("desc","")} | PU Mercado: {market_pu:,.2f}' if isinstance(market_pu,(int,float)) else f'Analisis: {clave} | {c.get("desc","")}'
        ws_det.cell(det_row,1,title).fill=_fill(C_GRIS_SEC); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8); ws_det.cell(det_row,1).alignment=_aln('left','center',wrap=True)
        ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
        det_row+=1
        current=None
        for item in c.get('granular_market_items') or []:
            if item.get('section') != current:
                current=item.get('section')
                ws_det.cell(det_row,1,current).fill=_fill(C_GRIS_FIL); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
                ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
                det_row+=1
            diff=item.get('delta_precio_pct')
            if item.get('fallback_contratista'):
                nota='Sin match: costo mercado = costo contratista como fallback.'
            elif item.get('is_percent_item'):
                nota='Porcentual calculado sobre la suma total de MO de mercado del servicio.'
            else:
                nota=item.get('match_reason') or ''
            vals=[item.get('codigo'),item.get('descripcion'),item.get('unidad'),item.get('precio_proveedor'),item.get('op') or '*',item.get('cantidad'),item.get('importe_proveedor'),item.get('precio_mercado'),diff,item.get('importe_mercado'),item.get('base_mercado'),item.get('section'),f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(' -'),item.get('confidence'),item.get('match_status'),nota]
            for col,val in enumerate(vals,1):
                cc=ws_det.cell(det_row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,13,16) else _aln('center','center')
                if col in (4,7,8,10,11) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
                if col in (9,14) and isinstance(val,(int,float)): cc.number_format='0.00%'
                if col in (8,9): cc.fill=_fill(C_AMARILLO_PMD)
            det_row+=1
        det_row = _step12_write_financial_summary_rows(ws_det, det_row, c) + 1

    _step08_build_matrix_base_sheet(wb, dato)
    if 'Matriz base CD' in wb.sheetnames:
        ws3 = wb['Matriz base CD']
        col = ws3.max_column + 1
        ws3.cell(1, col, 'Comentario IA / analisis humano')
        ws3.cell(1, col).fill = _fill(C_AZUL_OSC); ws3.cell(1, col).font = _fnt(bold=True, color=C_BLANCO, size=8)
        ws3.column_dimensions[get_column_letter(col)].width = 70
        comments = {str(k): _step11_ai_commentary_for_concept(v) for k, v in (dato.get('conceptos') or {}).items()}
        for r in range(2, ws3.max_row + 1):
            service = str(ws3.cell(r, 1).value or '').split('|')[0].replace('Analisis:', '').replace('Análisis:', '').strip()
            if service in comments:
                ws3.cell(r, col, comments[service]); ws3.cell(r, col).alignment = _aln('left','top', wrap=True)
    wb._sheets=[wb[n] for n in ['Comparativa','Detalle','Matriz base CD'] if n in wb.sheetnames]
    wb.save(output)
    return output

# -----------------------------------------------------------------------------
# Paso 15 REAL - Resumen PU por contratista para cantidades reales de proyecto
# -----------------------------------------------------------------------------
# Cambio de arquitectura:
# - La matriz del contratista sigue dando la composición del precio unitario.
# - El nuevo archivo Resumen PU da la cantidad real de cada servicio en la cotización.
# - Comparativa usa cantidad real para calcular Importe Contratista e Importe Mercado.
# - El match inicial entre Resumen PU y Matriz es por orden declarado, porque en el
#   ejemplo ambos contienen 63 servicios. Se guarda trazabilidad para revisión.

from difflib import SequenceMatcher as _Step15SequenceMatcher


def _step15_norm_text(value):
    try:
        return _normalize_text(value)
    except Exception:
        import re, unicodedata
        s = unicodedata.normalize('NFKD', str(value or '')).encode('ascii', 'ignore').decode('ascii')
        return re.sub(r'\s+', ' ', re.sub(r'[^a-zA-Z0-9]+', ' ', s.lower())).strip()


def _step15_as_float(value):
    if value is None or value == '':
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        s = str(value).replace('$', '').replace(',', '').replace('%', '').strip()
        if not s:
            return None
        return float(s)
    except Exception:
        return None


def extract_resumen_pu(filepath: str) -> list:
    """Lee el Excel Resumen PU del contratista.

    Formato soportado del ejemplo:
      Código | Concepto | Unidad | Cantidad | P. Unitario | Importe

    Devuelve una lista en el orden declarado, excluyendo encabezados, secciones,
    totales y renglones sin cantidad/precio utilizable.
    """
    rows = []
    if not filepath:
        return rows
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    seq = 0
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = list(row or []) + [None] * 6
        code, desc, unit, qty, pu, amount = vals[:6]
        desc_txt = str(desc or '').strip()
        qty_f = _step15_as_float(qty)
        pu_f = _step15_as_float(pu)
        amount_f = _step15_as_float(amount)
        # Encabezados/secciones/subtotales. Si no hay PU utilizable no es un PU real.
        if not desc_txt or pu_f is None:
            continue
        if desc_txt.upper() in {'CONCEPTO', 'PRESUPUESTO DE OBRA', 'TOTAL'}:
            continue
        if str(pu).strip().upper() == 'TOTAL':
            continue
        if qty_f is None:
            continue
        if amount_f is None:
            amount_f = qty_f * pu_f
        seq += 1
        rows.append({
            'seq': seq,
            'source_row': r_idx,
            'codigo_resumen': code,
            'descripcion_resumen': desc_txt.replace('_x000D_', '\n'),
            'unidad_resumen': unit,
            'cantidad_resumen': qty_f,
            'pu_resumen': pu_f,
            'importe_resumen': amount_f,
            'desc_norm': _step15_norm_text(desc_txt),
            'unidad_norm': normalize_unit(unit),
        })
    return rows


def _step15_similarity(a, b):
    a = _step15_norm_text(a)
    b = _step15_norm_text(b)
    if not a or not b:
        return 0.0
    toks_a = set(a.split())
    toks_b = set(b.split())
    jac = len(toks_a & toks_b) / max(1, len(toks_a | toks_b))
    seq = _Step15SequenceMatcher(None, a[:500], b[:500]).ratio()
    return max(jac, seq * 0.85)


def _step15_find_best_summary_for_concept(concept_key, concept, resumen_rows, used_indexes, order_index):
    # 1) Match prioritario por orden declarado.
    if order_index < len(resumen_rows) and order_index not in used_indexes:
        row = resumen_rows[order_index]
        sim = _step15_similarity(concept.get('desc'), row.get('descripcion_resumen'))
        unit_ok = normalize_unit(concept.get('unidad')) == row.get('unidad_norm')
        # Aunque la descripcion sea larga, conservar por orden; si unidad/PU coinciden, es muy confiable.
        return order_index, row, 'orden_declarado', sim, unit_ok

    # 2) Fallback por similitud descripcion/unidad.
    best = None
    best_score = -1.0
    for idx, row in enumerate(resumen_rows):
        if idx in used_indexes:
            continue
        sim = _step15_similarity(concept.get('desc'), row.get('descripcion_resumen'))
        unit_bonus = 0.12 if normalize_unit(concept.get('unidad')) == row.get('unidad_norm') else 0.0
        pu_bonus = 0.08 if isinstance(concept.get('pu'), (int, float)) and row.get('pu_resumen') and abs(float(concept.get('pu')) - float(row.get('pu_resumen'))) <= max(1.0, float(row.get('pu_resumen')) * 0.01) else 0.0
        score = sim + unit_bonus + pu_bonus
        if score > best_score:
            best = (idx, row, 'descripcion_unidad_pu', sim, normalize_unit(concept.get('unidad')) == row.get('unidad_norm'))
            best_score = score
    if best and best_score >= 0.42:
        return best
    return None, None, 'sin_match_resumen_pu', 0.0, False


def apply_resumen_pu_to_conceptos(conceptos: dict, resumen_rows: list) -> dict:
    """Aplica cantidades reales al dict de conceptos extraído de matriz.

    Actualiza por concepto:
      cantidad = cantidad real de cotización
      pu = PU resumen si existe
      total = cantidad * PU
      resumen_pu_match = trazabilidad del match
    """
    if not conceptos or not resumen_rows:
        return conceptos
    used = set()
    ordered = list(conceptos.items())
    for order_index, (clave, c) in enumerate(ordered):
        idx, row, method, sim, unit_ok = _step15_find_best_summary_for_concept(clave, c, resumen_rows, used, order_index)
        if row is None:
            c['resumen_pu_match'] = {'status': 'sin_match', 'method': method, 'applied': False}
            continue
        used.add(idx)
        qty = row.get('cantidad_resumen')
        pu = row.get('pu_resumen')
        amount = row.get('importe_resumen')
        if isinstance(qty, (int, float)):
            c['cantidad'] = qty
            c['cantidad_real_cotizacion'] = qty
        if isinstance(pu, (int, float)):
            c['pu'] = pu
            c['pu_resumen'] = pu
        if isinstance(amount, (int, float)):
            c['total'] = amount
        elif isinstance(qty, (int, float)) and isinstance(pu, (int, float)):
            c['total'] = qty * pu
        c['resumen_pu_match'] = {
            'status': 'match',
            'applied': True,
            'method': method,
            'source_row': row.get('source_row'),
            'seq': row.get('seq'),
            'codigo_resumen': row.get('codigo_resumen'),
            'descripcion_resumen': row.get('descripcion_resumen'),
            'unidad_resumen': row.get('unidad_resumen'),
            'cantidad_resumen': row.get('cantidad_resumen'),
            'pu_resumen': row.get('pu_resumen'),
            'importe_resumen': row.get('importe_resumen'),
            'similarity': round(float(sim or 0), 4),
            'unit_ok': bool(unit_ok),
        }
    # Guardar no usados para diagnóstico opcional.
    leftovers = [r for i, r in enumerate(resumen_rows) if i not in used]
    if leftovers:
        conceptos['__resumen_pu_no_usados__'] = {'rows': leftovers, 'count': len(leftovers)}
    return conceptos


_STEP15_BUILD_COMPARATIVO_PREV = build_comparativo


def build_comparativo(
    filepaths: list,
    nombres: list,
    output: str,
    meta: dict,
    catalogo_path: str = None,
    nacional_path: str = None,
    resumen_paths: list = None,
):
    """Paso 15: agrega archivo Resumen PU por contratista.

    Para evitar reescribir todo el flujo, se inyectan las cantidades reales justo
    despues de extract_conceptos mediante un override temporal y luego se invoca
    el flujo acumulado anterior.
    """
    resumen_paths = resumen_paths or []
    resumen_by_filepath = {}
    for fp, rp in zip(filepaths or [], resumen_paths):
        if rp:
            try:
                resumen_by_filepath[str(fp)] = extract_resumen_pu(rp)
            except Exception as exc:
                resumen_by_filepath[str(fp)] = {'__error__': str(exc)}

    original_extract = globals().get('extract_conceptos')

    def _extract_conceptos_step15(filepath):
        conceptos = original_extract(filepath)
        resumen = resumen_by_filepath.get(str(filepath))
        if isinstance(resumen, list) and resumen:
            apply_resumen_pu_to_conceptos(conceptos, resumen)
        elif isinstance(resumen, dict) and resumen.get('__error__'):
            for c in conceptos.values():
                if isinstance(c, dict):
                    c['resumen_pu_match'] = {'status': 'error', 'applied': False, 'error': resumen.get('__error__')}
        return conceptos

    globals()['extract_conceptos'] = _extract_conceptos_step15
    try:
        meta2 = dict(meta or {})
        meta2['usa_resumen_pu'] = bool(any(resumen_paths or []))
        return _STEP15_BUILD_COMPARATIVO_PREV(
            filepaths,
            nombres,
            output,
            meta2,
            catalogo_path=catalogo_path,
            nacional_path=nacional_path,
        )
    finally:
        globals()['extract_conceptos'] = original_extract


# Añadir trazabilidad de Resumen PU al tab Detalle sin alterar los cálculos base.
_STEP15_BUILD_SINGLE_PROVIDER_PREV = _build_single_provider_pmd


def _build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP15_BUILD_SINGLE_PROVIDER_PREV(output, meta, dato, provider_analysis)
    try:
        wb = openpyxl.load_workbook(output)
        if 'Detalle' in wb.sheetnames:
            ws = wb['Detalle']
            # Insertar/usar columnas al final para trazabilidad del resumen PU.
            start_col = ws.max_column + 1
            headers = ['Cantidad Resumen PU', 'PU Resumen PU', 'Importe Resumen PU', 'Match Resumen PU']
            for offset, header in enumerate(headers):
                cell = ws.cell(3, start_col + offset, header)
                cell.fill = _fill(C_AZUL_OSC)
                cell.font = _fnt(bold=True, color=C_BLANCO, size=8)
                cell.border = _brd()
                ws.column_dimensions[get_column_letter(start_col + offset)].width = 18 if offset < 3 else 34
            current_key = None
            resumen_map = {}
            for k, c in (dato.get('conceptos') or {}).items():
                if str(k).startswith('__'):
                    continue
                resumen_map[str(k)] = c.get('resumen_pu_match') or {}
            for r in range(4, ws.max_row + 1):
                val = str(ws.cell(r, 1).value or '')
                if val.startswith('Analisis:') or val.startswith('Análisis:'):
                    raw = val.split('|')[0].replace('Analisis:', '').replace('Análisis:', '').strip()
                    current_key = raw
                    info = resumen_map.get(current_key) or {}
                    ws.cell(r, start_col, info.get('cantidad_resumen'))
                    ws.cell(r, start_col + 1, info.get('pu_resumen'))
                    ws.cell(r, start_col + 2, info.get('importe_resumen'))
                    ws.cell(r, start_col + 3, f"{info.get('status','')} / {info.get('method','')} / fila {info.get('source_row','')}")
                    for cc in range(start_col, start_col + 4):
                        ws.cell(r, cc).border = _brd()
                        ws.cell(r, cc).alignment = _aln('center', 'center', wrap=True)
                    for cc in (start_col + 1, start_col + 2):
                        ws.cell(r, cc).number_format = '$#,##0.00'
        # No se escribe ninguna marca de version/paso en el Excel final.
        # La trazabilidad queda solo en las columnas de Resumen PU del tab Detalle.
        wb.save(output)
    except Exception:
        pass
    return result

# -----------------------------------------------------------------------------
# Paso 17 - Correccion real Resumen PU en Excel y Web
# -----------------------------------------------------------------------------
# Objetivo:
# - El Resumen PU debe afectar cantidades e importes reales, no solo trazabilidad.
# - La web debe mostrar totales y porcentajes basados en esas cantidades reales.
# - Para una corrida con 1 proveedor, se usa un flujo directo y controlado:
#   matriz proveedor + resumen PU + mercado granular -> Excel final.

_STEP17_BUILD_COMPARATIVO_PREV = build_comparativo
_STEP17_CALCULATE_EXECUTIVE_FINDINGS_PREV = calculate_executive_findings


def _step17_provider_total(conceptos):
    total = 0.0
    for k, c in (conceptos or {}).items():
        if str(k).startswith('__') or not isinstance(c, dict):
            continue
        v = _concept_total_amount(c)
        if isinstance(v, (int, float)):
            total += float(v)
    return round(total, 2)


def _step17_market_total(conceptos):
    total = 0.0
    for k, c in (conceptos or {}).items():
        if str(k).startswith('__') or not isinstance(c, dict):
            continue
        v = _concept_market_total_amount(c)
        if isinstance(v, (int, float)):
            total += float(v)
    return round(total, 2)


def _step17_apply_resumen_and_market(conceptos, resumen_rows=None, refs=None):
    """Aplica cantidades reales y recalcula mercado granular.

    Se usa tanto para Excel como para preview web. Mantiene el orden declarado.
    """
    if isinstance(resumen_rows, list) and resumen_rows:
        apply_resumen_pu_to_conceptos(conceptos, resumen_rows)
    try:
        _step08_apply_market_catalog_pricing(conceptos)
    except Exception as exc:
        for c in (conceptos or {}).values():
            if isinstance(c, dict):
                c['granular_market_error'] = str(exc)
    if refs:
        try:
            attach_construdata_matrix_analysis(conceptos, refs)
        except Exception:
            pass
    return conceptos


def calculate_executive_findings(provider_name, conceptos, total_proveedor=None, all_provider_totals=None):
    """Hallazgos ejecutivos recalculados con cantidades reales y mercado granular.

    Esta version reemplaza la logica antigua basada en cantidad=1 y matches parciales.
    """
    provider_total = _step17_provider_total(conceptos)
    market_total = _step17_market_total(conceptos)
    concept_rows = []
    drivers = []
    indirectos = []
    for clave, c in (conceptos or {}).items():
        if str(clave).startswith('__') or not isinstance(c, dict):
            continue
        prov = _concept_total_amount(c)
        mkt = _concept_market_total_amount(c)
        qty = c.get('cantidad') if isinstance(c.get('cantidad'), (int, float)) else 1
        pu = c.get('pu') if isinstance(c.get('pu'), (int, float)) else None
        mpu = _market_pu_with_indirect(c)
        delta = ((prov - mkt) / mkt) if isinstance(prov, (int, float)) and isinstance(mkt, (int, float)) and mkt else None
        participation = (prov / provider_total) if provider_total and isinstance(prov, (int, float)) else 0
        concept_rows.append({
            'clave': clave,
            'descripcion': c.get('desc', ''),
            'cantidad': qty,
            'pu_contratista': round(pu, 2) if isinstance(pu, (int, float)) else None,
            'pu_mercado': round(mpu, 2) if isinstance(mpu, (int, float)) else None,
            'total_contratista': round(prov, 2) if isinstance(prov, (int, float)) else None,
            'total_mercado': round(mkt, 2) if isinstance(mkt, (int, float)) else None,
            'participacion': round(participation, 4),
            'delta': round(delta, 4) if isinstance(delta, (int, float)) else None,
            'impacto': round(participation * max(delta or 0, 0), 4) if isinstance(delta, (int, float)) else None,
        })
        if isinstance(delta, (int, float)) and delta > 0.10:
            drivers.append({
                'tipo': 'servicio_sobre_mercado',
                'concepto': f"{clave} - {c.get('desc','')}",
                'proveedor': round(prov, 2) if isinstance(prov, (int, float)) else None,
                'mercado': round(mkt, 2) if isinstance(mkt, (int, float)) else None,
                'delta': round(delta, 4),
                'participacion': round(participation, 4),
                'impacto': round(participation * delta, 4),
                'estado': 'critico' if delta > 0.30 and participation > 0.05 else 'alto',
            })
        # Drivers por insumo con impacto real multiplicado por cantidad del servicio.
        for item in c.get('granular_market_items') or []:
            ip = item.get('importe_proveedor')
            im = item.get('importe_mercado')
            if not isinstance(ip, (int, float)) or not isinstance(im, (int, float)) or im == 0:
                continue
            dlt = (float(ip) - float(im)) / float(im)
            if dlt <= 0.15:
                continue
            prov_item_total = float(ip) * float(qty)
            mkt_item_total = float(im) * float(qty)
            drivers.append({
                'tipo': item.get('domain') or item.get('section') or 'insumo',
                'concepto': f"{clave} · {item.get('descripcion')}",
                'proveedor': round(prov_item_total, 2),
                'mercado': round(mkt_item_total, 2),
                'delta': round(dlt, 4),
                'participacion': round(prov_item_total / provider_total, 4) if provider_total else 0,
                'impacto': round(((prov_item_total / provider_total) if provider_total else 0) * dlt, 4),
                'estado': 'critico' if dlt > 0.50 else 'alto',
            })
        cd = c.get('costo_directo')
        ci = c.get('subtotal2')
        if isinstance(cd, (int, float)) and cd and isinstance(ci, (int, float)):
            pct = ci / cd
            if pct > 0.25:
                indirectos.append({'clave': clave, 'descripcion': c.get('desc', ''), 'proveedor_pct': round(pct, 4), 'mercado_pct': 0.25, 'estado': 'critico' if pct > 0.30 else 'arriba'})
    concept_rows.sort(key=lambda x: x.get('total_contratista') or 0, reverse=True)
    acum = 0.0
    for r in concept_rows:
        acum += r.get('participacion') or 0
        r['acumulado'] = round(acum, 4)
    drivers.sort(key=lambda x: x.get('impacto') or 0, reverse=True)
    ajuste = ((provider_total - market_total) / market_total) if market_total else None
    return {
        'provider_name': provider_name,
        'total_contratista': provider_total,
        'total_mercado': market_total,
        'ajuste_global': round(ajuste, 4) if isinstance(ajuste, (int, float)) else None,
        'top_partidas': concept_rows[:50],
        'drivers': drivers[:50],
        'indirectos_alertas': indirectos[:20],
        'resumen': 'Totales calculados con cantidades reales del Resumen PU cuando fue cargado.',
    }


def build_comparativo(filepaths, nombres, output, meta, catalogo_path=None, nacional_path=None, resumen_paths=None):
    """Paso 17: flujo controlado para 1 proveedor con Resumen PU.

    Evita depender de wrappers previos. Si hay 1 proveedor, genera el Excel directo con:
    - matriz del contratista
    - cantidades reales del Resumen PU
    - mercado granular por insumo
    Para multi proveedor, conserva el flujo acumulado anterior.
    """
    resumen_paths = resumen_paths or []
    if len(filepaths or []) == 1:
        filepath = filepaths[0]
        nombre = (nombres or ['Proveedor 1'])[0] or 'Proveedor 1'
        conceptos = extract_conceptos(filepath)
        resumen_rows = []
        if resumen_paths and resumen_paths[0]:
            resumen_rows = extract_resumen_pu(resumen_paths[0])
        refs = None
        try:
            refs = extract_precios_nacional(nacional_path) if nacional_path else None
        except Exception:
            refs = None
        _step17_apply_resumen_and_market(conceptos, resumen_rows, refs)
        dato = {
            'nombre': nombre,
            'conceptos': conceptos,
            'total_estimado': _step17_provider_total(conceptos),
            'status': 'ok',
        }
        analysis = calculate_executive_findings(nombre, conceptos, dato['total_estimado'], [])
        return _build_single_provider_pmd(output, meta or {}, dato, analysis)
    # Fallback multi proveedor.
    return _STEP17_BUILD_COMPARATIVO_PREV(filepaths, nombres, output, meta, catalogo_path=catalogo_path, nacional_path=nacional_path, resumen_paths=resumen_paths)


# -----------------------------------------------------------------------------
# Paso 18 - Fix definitivo Resumen PU en Comparativa
# -----------------------------------------------------------------------------
# Problema corregido:
# - El Resumen PU podia parsearse correctamente, pero el tab Comparativa no dejaba
#   trazabilidad visible y en algunos flujos el usuario no podia confirmar si la
#   cantidad venia del resumen o de la matriz.
# - Este override fuerza que Comparativa use SIEMPRE, si existe, los campos del
#   resumen_pu_match: cantidad_resumen, pu_resumen e importe_resumen.
# - Si no existe match de resumen, conserva los valores originales de la matriz.

_STEP18_BUILD_COMPARATIVO_PREV = build_comparativo


def _step18_effective_qty_pu_amount(concepto):
    """Valores efectivos para Comparativa.

    Prioridad:
      1) Resumen PU aplicado: cantidad_resumen, pu_resumen, importe_resumen.
      2) Campos normalizados del concepto: cantidad, pu, total.
      3) Fallback: cantidad=1.
    """
    match = concepto.get('resumen_pu_match') or {}
    qty = match.get('cantidad_resumen') if match.get('applied') else None
    pu = match.get('pu_resumen') if match.get('applied') else None
    amount = match.get('importe_resumen') if match.get('applied') else None

    if not isinstance(qty, (int, float)):
        qty = concepto.get('cantidad') if isinstance(concepto.get('cantidad'), (int, float)) else 1
    if not isinstance(pu, (int, float)):
        pu = concepto.get('pu') if isinstance(concepto.get('pu'), (int, float)) else None
    if not isinstance(amount, (int, float)):
        amount = concepto.get('total') if isinstance(concepto.get('total'), (int, float)) else None
    if not isinstance(amount, (int, float)) and isinstance(qty, (int, float)) and isinstance(pu, (int, float)):
        amount = qty * pu
    return qty, pu, amount


def _step18_selected_80_keys(conceptos_items):
    rows = []
    total = 0.0
    for clave, c in conceptos_items:
        if str(clave).startswith('__') or not isinstance(c, dict):
            continue
        _, _, amount = _step18_effective_qty_pu_amount(c)
        if isinstance(amount, (int, float)):
            total += float(amount)
            rows.append((clave, float(amount)))
    rows.sort(key=lambda x: x[1], reverse=True)
    selected = set()
    acc = 0.0
    for clave, amount in rows:
        if total and acc >= 0.80:
            break
        selected.add(clave)
        acc += amount / total if total else 0
    return selected, acc


def _step18_build_single_provider_pmd(output, meta, dato, provider_analysis):
    """Writer limpio para 1 proveedor. Mantiene orden declarado y muestra cantidad real.

    Hojas principales:
      - Comparativa: usa cantidad/PU/importe del Resumen PU cuando existe.
      - Detalle: conserva la matriz del contratista enriquecida con mercado.
      - Matriz base CD: conserva el análisis concepto/matriz si hay match.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comparativa'
    ws_det = wb.create_sheet('Detalle')

    conceptos_items = [(k, v) for k, v in (dato.get('conceptos') or {}).items() if not str(k).startswith('__') and isinstance(v, dict)]
    blue_keys, blue_pct = _step18_selected_80_keys(conceptos_items)

    total_contratista = 0.0
    for _, c in conceptos_items:
        _, _, amt = _step18_effective_qty_pu_amount(c)
        if isinstance(amt, (int, float)):
            total_contratista += float(amt)

    for col, width in {'A':14,'B':58,'C':10,'D':12,'E':13,'F':16,'G':16,'H':18,'I':14,'J':16,'K':14,'L':14,'M':18,'N':42}.items():
        ws.column_dimensions[col].width = width
    ws.merge_cells('A1:N1')
    ws['A1'] = f"COMPARATIVO DE COTIZACION — {meta.get('proyecto','')}"
    ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10); ws['A1'].alignment = _aln('center','center')
    ws['A2'] = 'Criterio azul'
    ws['B2'] = f'Azul = servicios que acumulan {blue_pct*100:.2f}% aprox. del importe contratista. Las filas conservan el orden declarado.'
    ws['A3'] = 'Fuente cantidades'
    resumen_status = 'RECIBIDO' if meta.get('usa_resumen_pu') else 'NO RECIBIDO'
    ws['B3'] = f'Resumen PU: {resumen_status} | renglones leidos: {meta.get("resumen_pu_rows", 0)} | servicios aplicados: {meta.get("resumen_pu_aplicados", 0)}. Cantidad/PU/Importe Contratista deben venir de ese archivo.'
    for cell in ['A2','A3']:
        ws[cell].font = _fnt(bold=True, size=8)
    for cell in ['B2','B3']:
        ws[cell].font = _fnt(size=8); ws[cell].alignment = _aln('left','center',wrap=True)

    headers = [
        'Servicio','Descripcion','Unidad','Cantidad','PU Contratista','Importe Contratista',
        'Cantidad Resumen PU','Fila Resumen PU','PU Mercado','Importe Mercado','Dif %',
        'Cobertura mercado','Estado','Nota'
    ]
    for col,h in enumerate(headers,1):
        cell=ws.cell(5,col,h)
        cell.fill=_fill(C_AMARILLO_PMD if col in (9,11) else C_AZUL_OSC)
        cell.font=_fnt(bold=True,color=('000000' if col in (9,11) else C_BLANCO),size=8)
        cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)

    row=6
    total_market=0.0
    for clave,c in conceptos_items:
        cantidad, pu, importe = _step18_effective_qty_pu_amount(c)
        match = c.get('resumen_pu_match') or {}
        market_pu = _step13_market_pu(c) if '_step13_market_pu' in globals() else _market_pu_with_indirect(c)
        market_total = market_pu * cantidad if isinstance(market_pu,(int,float)) and isinstance(cantidad,(int,float)) else None
        if isinstance(market_total,(int,float)):
            total_market += market_total
        diff_pct = ((importe-market_total)/market_total) if isinstance(importe,(int,float)) and isinstance(market_total,(int,float)) and market_total else None
        nota = 'Cantidad tomada del Resumen PU.' if match.get('applied') else 'Sin match de Resumen PU: cantidad tomada de matriz.'
        vals=[
            clave,c.get('desc'),c.get('unidad'),cantidad,pu,importe,
            match.get('cantidad_resumen') if match.get('applied') else None,
            match.get('source_row') if match.get('applied') else None,
            market_pu,market_total,diff_pct,c.get('granular_market_match_coverage'),
            _step08_market_status(diff_pct) if '_step08_market_status' in globals() else '',nota
        ]
        is_blue = clave in blue_keys
        for col,val in enumerate(vals,1):
            cell=ws.cell(row,col,val); cell.border=_brd(); cell.font=_fnt(size=8)
            cell.alignment=_aln('left','center',wrap=True) if col in (2,14) else _aln('center','center')
            if col in (5,6,9,10) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
            if col in (11,12) and isinstance(val,(int,float)): cell.number_format='0.00%'
            if is_blue: cell.fill=_fill('D9EAF7')
            if col in (9,11): cell.fill=_fill(C_AMARILLO_PMD)
        row+=1

    row+=1
    totals=[(1,'TOTAL'),(6,total_contratista),(10,total_market),(11,((total_contratista-total_market)/total_market if total_market else None))]
    for col,val in totals:
        cell=ws.cell(row,col,val); cell.font=_fnt(bold=True,size=8); cell.border=_brd(); cell.fill=_fill('D9EAF7')
        if col in (6,10) and isinstance(val,(int,float)): cell.number_format='$#,##0.00'
        if col==11 and isinstance(val,(int,float)): cell.number_format='0.00%'

    # Detalle tecnico de matriz, conservando calculos ya enriquecidos por mercado granular.
    widths={'A':15,'B':58,'C':10,'D':13,'E':7,'F':10,'G':14,'H':16,'I':14,'J':14,'K':14,'L':18,'M':44,'N':12,'O':20,'P':42}
    for col,width in widths.items(): ws_det.column_dimensions[col].width=width
    ws_det.merge_cells('A1:P1')
    ws_det['A1']='MATRIZ DEL CONTRATISTA CON PRECIOS DE MERCADO CONSTRUDATA POR INSUMO'
    ws_det['A1'].fill=_fill(C_AZUL_OSC); ws_det['A1'].font=_fnt(bold=True,color=C_BLANCO,size=9); ws_det['A1'].alignment=_aln('center','center')
    headers_det=['Codigo','Concepto / Insumo','Unidad','Costo Contratista','Op.','Cantidad','Importe Contratista','Costo Mercado','Dif % Costo','Importe Mercado','Base Mercado %','Tipo','Match Construdata','Conf.','Estado','Nota']
    for col,h in enumerate(headers_det,1):
        cell=ws_det.cell(3,col,h); cell.fill=_fill(C_AMARILLO_PMD if col in (8,9) else C_AZUL_OSC); cell.font=_fnt(bold=True,color=('000000' if col in (8,9) else C_BLANCO),size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
    det_row=4
    for clave,c in conceptos_items:
        cantidad, pu, importe = _step18_effective_qty_pu_amount(c)
        market_pu = _step13_market_pu(c) if '_step13_market_pu' in globals() else _market_pu_with_indirect(c)
        title = f'Analisis: {clave} | {c.get("desc","")} | Cantidad real: {cantidad} | PU Mercado: {market_pu:,.2f}' if isinstance(market_pu,(int,float)) else f'Analisis: {clave} | {c.get("desc","")} | Cantidad real: {cantidad}'
        ws_det.cell(det_row,1,title).fill=_fill(C_GRIS_SEC); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8); ws_det.cell(det_row,1).alignment=_aln('left','center',wrap=True)
        ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
        det_row+=1
        current=None
        for item in c.get('granular_market_items') or []:
            if item.get('section') != current:
                current=item.get('section')
                ws_det.cell(det_row,1,current).fill=_fill(C_GRIS_FIL); ws_det.cell(det_row,1).font=_fnt(bold=True,size=8)
                ws_det.merge_cells(start_row=det_row,start_column=1,end_row=det_row,end_column=16)
                det_row+=1
            diff=item.get('delta_precio_pct')
            if item.get('fallback_contratista'):
                nota='Sin match: costo mercado = costo contratista como fallback.'
            elif item.get('is_percent_item'):
                nota='Porcentual calculado sobre la suma total de MO de mercado del servicio.'
            else:
                nota=item.get('match_reason') or ''
            vals=[item.get('codigo'),item.get('descripcion'),item.get('unidad'),item.get('precio_proveedor'),item.get('op') or '*',item.get('cantidad'),item.get('importe_proveedor'),item.get('precio_mercado'),diff,item.get('importe_mercado'),item.get('base_mercado'),item.get('section'),f"{item.get('codigo_mercado') or ''} - {item.get('descripcion_mercado') or ''}".strip(' -'),item.get('confidence'),item.get('match_status'),nota]
            for col,val in enumerate(vals,1):
                cc=ws_det.cell(det_row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,13,16) else _aln('center','center')
                if col in (4,7,8,10,11) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
                if col in (9,14) and isinstance(val,(int,float)): cc.number_format='0.00%'
                if col in (8,9): cc.fill=_fill(C_AMARILLO_PMD)
            det_row+=1
        if '_step12_write_financial_summary_rows' in globals():
            det_row = _step12_write_financial_summary_rows(ws_det, det_row, c) + 1

    if '_step08_build_matrix_base_sheet' in globals():
        try:
            _step08_build_matrix_base_sheet(wb, dato)
        except Exception:
            pass
    wb._sheets=[wb[n] for n in ['Comparativa','Detalle','Matriz base CD'] if n in wb.sheetnames]
    wb.save(output)
    return output


def build_comparativo(filepaths, nombres, output, meta, catalogo_path=None, nacional_path=None, resumen_paths=None):
    """Paso 18: flujo controlado con Resumen PU obligatorio si fue cargado.

    Para 1 proveedor genera el Excel usando un writer que lee explícitamente los
    campos del Resumen PU. Esto elimina la ambigüedad de wrappers anteriores.
    """
    resumen_paths = resumen_paths or []
    if len(filepaths or []) == 1:
        filepath = filepaths[0]
        nombre = (nombres or ['Proveedor 1'])[0] or 'Proveedor 1'
        conceptos = extract_conceptos(filepath)
        resumen_rows = []
        resumen_error = None
        resumen_file_received = bool(resumen_paths and resumen_paths[0])
        if resumen_file_received:
            try:
                resumen_rows = extract_resumen_pu(resumen_paths[0])
            except Exception as exc:
                resumen_error = str(exc)
        # En modo 1 proveedor el Resumen PU ya es parte obligatoria del flujo: si no llega,
        # es preferible detener el proceso que generar un Excel con cantidad=1.
        if not resumen_file_received:
            raise ValueError('No se recibio el archivo Resumen PU. Cargalo junto con la matriz para usar cantidades reales.')
        if resumen_error:
            raise ValueError(f'No se pudo leer el Resumen PU: {resumen_error}')
        if not resumen_rows:
            raise ValueError('El archivo Resumen PU fue recibido, pero no se detectaron renglones validos con Codigo/Concepto/Unidad/Cantidad/PU/Importe.')
        refs = None
        try:
            refs = extract_precios_nacional(nacional_path) if nacional_path else None
        except Exception:
            refs = None
        _step17_apply_resumen_and_market(conceptos, resumen_rows, refs)
        if resumen_error:
            for c in conceptos.values():
                if isinstance(c, dict):
                    c['resumen_pu_match'] = {'status': 'error', 'applied': False, 'error': resumen_error}
        dato = {'nombre': nombre, 'conceptos': conceptos, 'total_estimado': _step17_provider_total(conceptos), 'status': 'ok'}
        analysis = calculate_executive_findings(nombre, conceptos, dato['total_estimado'], [])
        applied_count = 0
        for c in conceptos.values():
            if isinstance(c, dict) and (c.get('resumen_pu_match') or {}).get('applied'):
                applied_count += 1
        if resumen_rows and applied_count == 0:
            raise ValueError('Resumen PU leido, pero no se aplico a ningun servicio. Revisa el orden/estructura del archivo.')
        meta2 = dict(meta or {})
        meta2['usa_resumen_pu'] = bool(resumen_rows)
        meta2['resumen_pu_rows'] = len(resumen_rows or [])
        meta2['resumen_pu_aplicados'] = applied_count
        return _step18_build_single_provider_pmd(output, meta2, dato, analysis)
    return _STEP18_BUILD_COMPARATIVO_PREV(filepaths, nombres, output, meta, catalogo_path=catalogo_path, nacional_path=nacional_path, resumen_paths=resumen_paths)


# -----------------------------------------------------------------------------
# Paso 22 - IA conceptual activa + Matriz base CD con base embebida + analisis experto
# -----------------------------------------------------------------------------
# Problema corregido:
# - El flujo de un solo proveedor ya no recibe Neodata manual, por lo que `refs` llegaba None.
# - Como `attach_construdata_matrix_analysis` solo se ejecutaba si refs existia, el tab
#   `Matriz base CD` quedaba sin matches.
# - Claude/Voyage conceptual podia ejecutarse, pero no dejaba trazabilidad en usage.
#
# Nuevo comportamiento:
# - Siempre intenta cargar `data/construdata_matrices.xlsx` para el tab de matriz base.
# - La referencia de mercado principal SIGUE siendo granular por insumo; la matriz CD solo
#   se usa para explicar omisiones/extras y construir matriz base de referencia.
# - Si hay ANTHROPIC_API_KEY/VOYAGE_API_KEY, se usan en `match_concepto_construdata`.
# - Se agrega una hoja `Analisis experto IA` con dictamen humano; usa Claude si esta
#   disponible y cae a analisis deterministico si no.

_STEP22_REFS_CACHE = None
_STEP22_APPLY_PREV = _step17_apply_resumen_and_market
_STEP22_BUILD_PMD_PREV = _step18_build_single_provider_pmd


def _step22_data_dir():
    return _quantia_data_dir()


def _step22_load_matrix_refs():
    """Carga la base completa concepto -> matriz desde data/construdata_matrices.xlsx.

    No se usa para precio de mercado granular; solo para tab Matriz base CD y analisis
    de omisiones/extras. Cachea el resultado para no releer el Excel en cada concepto.
    """
    global _STEP22_REFS_CACHE
    if isinstance(_STEP22_REFS_CACHE, dict):
        return _STEP22_REFS_CACHE
    candidates = []
    env_path = os.getenv('CONSTRUDATA_MATRICES_PATH')
    if env_path:
        candidates.append(Path(env_path))
    candidates.append(_step22_data_dir() / 'construdata_matrices.xlsx')
    candidates.append(Path('data') / 'construdata_matrices.xlsx')
    for path in candidates:
        try:
            if path and path.exists():
                refs = extract_precios_nacional(str(path))
                cds = _get_construdata_concepts(refs)
                if cds:
                    refs['__step22_source_path__'] = str(path)
                    refs['__step22_total_concepts__'] = len(cds)
                    _STEP22_REFS_CACHE = refs
                    return refs
        except Exception as exc:
            _STEP22_REFS_CACHE = {'__step22_error__': f'{type(exc).__name__}: {exc}'}
    if _STEP22_REFS_CACHE is None:
        _STEP22_REFS_CACHE = {'__step22_error__': 'No se encontro data/construdata_matrices.xlsx'}
    return _STEP22_REFS_CACHE


def _step17_apply_resumen_and_market(conceptos, resumen_rows=None, refs=None):
    """Paso 22: aplica Resumen PU + mercado granular + matriz base CD siempre que exista.

    Mantiene el mercado por catalogos granulares, pero reactiva la busqueda concepto/matriz
    usando la base embebida si no se subio una referencia Neodata manual.
    """
    conceptos = _STEP22_APPLY_PREV(conceptos, resumen_rows, refs)
    matrix_refs = refs or _step22_load_matrix_refs()
    try:
        if isinstance(matrix_refs, dict) and _get_construdata_concepts(matrix_refs):
            attach_construdata_matrix_analysis(conceptos, matrix_refs)
        else:
            reason = (matrix_refs or {}).get('__step22_error__') if isinstance(matrix_refs, dict) else 'sin referencia'
            for c in (conceptos or {}).values():
                if isinstance(c, dict):
                    c['construdata_concept_match'] = {
                        'status': 'sin_base_construdata',
                        'score': None,
                        'reason': reason or 'No se cargo construdata_matrices.xlsx',
                        'ai_status': get_concept_ai_status(),
                    }
                    c['construdata_matrix_comparison'] = {'rows': [], 'findings': [], 'reason': reason}
    except Exception as exc:
        for c in (conceptos or {}).values():
            if isinstance(c, dict):
                c['construdata_concept_match'] = {
                    'status': 'error_matrix_cd',
                    'score': None,
                    'reason': f'{type(exc).__name__}: {exc}',
                    'ai_status': get_concept_ai_status(),
                }
                c['construdata_matrix_comparison'] = {'rows': [], 'findings': [], 'reason': str(exc)}
    return conceptos


def _step22_ai_status_text():
    try:
        st = get_concept_ai_status() or {}
    except Exception:
        st = {}
    parts = [f"modo={st.get('mode','deterministic')}"]
    parts.append('Claude=' + ('activo' if st.get('anthropic_enabled') else 'inactivo'))
    parts.append('Voyage=' + ('activo' if st.get('voyage_enabled') else 'inactivo'))
    return ' | '.join(parts)


def _step22_concept_match_summary(conceptos, limit=12):
    rows = []
    for clave, c in (conceptos or {}).items():
        if not isinstance(c, dict) or str(clave).startswith('__'):
            continue
        m = c.get('construdata_concept_match') or {}
        comp = c.get('construdata_matrix_comparison') or {}
        selected = m.get('selected') or {}
        selected_matches = m.get('selected_matches') or []
        codigos = []
        motivos = []
        for sm in selected_matches[:4]:
            cd = sm.get('concept') or {}
            if cd.get('codigo'):
                codigos.append(str(cd.get('codigo')))
            if sm.get('motivo'):
                motivos.append(str(sm.get('motivo')))
        if not codigos and selected.get('codigo'):
            codigos = [str(selected.get('codigo'))]
        rows.append({
            'servicio': clave,
            'descripcion': c.get('desc') or '',
            'status': m.get('status') or comp.get('match_type') or 'sin_match',
            'score': m.get('score') if m.get('score') is not None else comp.get('score'),
            'codigos': ', '.join(codigos),
            'reason': m.get('reason') or '; '.join(motivos) or comp.get('reason') or '',
            'matrix_rows': len(comp.get('rows') or []),
            'findings': len(c.get('construdata_matrix_findings') or []),
        })
    return rows[:limit]


def _step22_deterministic_expert_text(dato, provider_analysis=None):
    conceptos = dato.get('conceptos') or {}
    total_provider = _step17_provider_total(conceptos) if '_step17_provider_total' in globals() else dato.get('total_estimado')
    total_market = _step17_market_total(conceptos) if '_step17_market_total' in globals() else None
    delta = ((total_provider - total_market) / total_market) if isinstance(total_provider,(int,float)) and isinstance(total_market,(int,float)) and total_market else None
    matched = 0; weak = 0; no_match = 0; extras = 0; faltantes = 0
    top_drivers = []
    for clave, c in conceptos.items():
        if not isinstance(c, dict) or str(clave).startswith('__'):
            continue
        m = c.get('construdata_concept_match') or {}
        status = str(m.get('status') or '').lower()
        if 'match' in status and 'sin' not in status:
            matched += 1
            if 'debil' in status or 'revision' in status or m.get('requires_review'):
                weak += 1
        else:
            no_match += 1
        comp = c.get('construdata_matrix_comparison') or {}
        for r in comp.get('rows') or []:
            if r.get('estado') == 'sin_match_insumo': extras += 1
            if r.get('estado') == 'faltante_en_proveedor': faltantes += 1
        d = _concept_total_amount(c) if '_concept_total_amount' in globals() else c.get('total')
        mkt = _concept_market_total_amount(c) if '_concept_market_total_amount' in globals() else None
        if isinstance(d,(int,float)) and isinstance(mkt,(int,float)) and mkt:
            top_drivers.append((abs(d-mkt), clave, c.get('desc'), d, mkt, (d-mkt)/mkt))
    top_drivers.sort(reverse=True)
    lines = []
    lines.append('Analisis experto preliminar de precios unitarios')
    lines.append(f'Referencia IA/matriz: {_step22_ai_status_text()}.')
    if isinstance(total_provider,(int,float)) and isinstance(total_market,(int,float)):
        txt = f'Total contratista ${total_provider:,.2f} vs mercado ${total_market:,.2f}'
        if isinstance(delta,(int,float)): txt += f' ({delta:.2%}).'
        lines.append(txt)
    lines.append(f'Se localizaron {matched} servicios con candidato de matriz Construdata; {weak} requieren revision y {no_match} quedaron sin match confiable.')
    lines.append(f'En matriz base CD se detectan {faltantes} elementos recomendados no ubicados en proveedor y {extras} elementos declarados fuera de la matriz base.')
    if top_drivers[:3]:
        lines.append('Principales partidas por impacto economico:')
        for _, clave, desc, prov, mkt, dlt in top_drivers[:3]:
            lines.append(f'- {clave}: {desc or ""} | proveedor ${prov:,.2f} vs mercado ${mkt:,.2f} ({dlt:.2%}).')
    lines.append('Criterio: los precios se evalúan por insumo granular; la matriz Construdata se usa para detectar alcance, omisiones y posibles conceptos compuestos, no para sustituir el cálculo de mercado por catálogo.')
    return '\n'.join(lines)


def _step22_claude_expert_text(dato, provider_analysis=None):
    if not (os.getenv('ANTHROPIC_API_KEY') or '').strip():
        return None, 'Claude inactivo: falta ANTHROPIC_API_KEY'
    try:
        from material_ai_matcher import _anthropic_message
        conceptos = dato.get('conceptos') or {}
        compact = []
        for row in _step22_concept_match_summary(conceptos, limit=15):
            compact.append(row)
        payload = {
            'rol': 'Eres un experto senior en análisis de precios unitarios, presupuestos de obra y revisión PMD.',
            'tarea': 'Redacta un dictamen ejecutivo breve, humano y técnico. No inventes datos; usa solo la información entregada.',
            'reglas': [
                'Explica diferencias relevantes entre contratista y mercado.',
                'Menciona si la matriz Construdata sirve como referencia de alcance o si requiere revisión.',
                'Distingue precio por insumo granular vs matriz base por concepto.',
                'No afirmes que un match es definitivo si aparece como revisión o débil.'
            ],
            'proveedor': dato.get('nombre'),
            'total_proveedor': _step17_provider_total(conceptos) if '_step17_provider_total' in globals() else dato.get('total_estimado'),
            'total_mercado': _step17_market_total(conceptos) if '_step17_market_total' in globals() else None,
            'matches_matriz_cd': compact,
            'hallazgos': (provider_analysis or {}).get('drivers', [])[:10] if isinstance(provider_analysis, dict) else [],
            'respuesta': 'Texto en español, 4 a 8 bullets ejecutivos.'
        }
        txt = _anthropic_message(payload, max_tokens=900, timeout=35)
        if txt:
            return txt.strip(), 'Claude activo'
        return None, 'Claude no devolvio texto'
    except Exception as exc:
        return None, f'Claude error: {type(exc).__name__}: {str(exc)[:200]}'


def _step22_add_expert_ai_sheet(output, dato, provider_analysis=None):
    try:
        wb = openpyxl.load_workbook(output)
        if 'Analisis experto IA' in wb.sheetnames:
            del wb['Analisis experto IA']
        ws = wb.create_sheet('Analisis experto IA')
        ws.sheet_view.showGridLines = False
        ws.column_dimensions['A'].width = 28
        ws.column_dimensions['B'].width = 110
        ws['A1'] = 'ANALISIS EXPERTO IA / PMD'
        ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10)
        ws.merge_cells('A1:B1')
        claude_txt, claude_status = _step22_claude_expert_text(dato, provider_analysis)
        final_txt = claude_txt or _step22_deterministic_expert_text(dato, provider_analysis)
        rows = [
            ('Estado IA', _step22_ai_status_text()),
            ('Claude dictamen', claude_status),
            ('Uso correcto', 'IA ayuda a homologar materiales/conceptos y explicar; los importes se calculan con reglas y catálogos.'),
            ('Dictamen', final_txt),
        ]
        r = 3
        for k, v in rows:
            ws.cell(r,1,k).font = _fnt(bold=True, size=9); ws.cell(r,1).border = _brd(); ws.cell(r,1).alignment = _aln('left','top',wrap=True)
            ws.cell(r,2,v).font = _fnt(size=9); ws.cell(r,2).border = _brd(); ws.cell(r,2).alignment = _aln('left','top',wrap=True)
            if k == 'Dictamen': ws.row_dimensions[r].height = 180
            else: ws.row_dimensions[r].height = 34
            r += 1
        r += 1
        headers = ['Servicio','Descripción','Status match','Score','Código(s) CD','Filas matriz CD','Motivo']
        for cidx, h in enumerate(headers,1):
            cell = ws.cell(r,cidx,h); cell.fill = _fill(C_AZUL_OSC); cell.font = _fnt(bold=True,color=C_BLANCO,size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
        r += 1
        for row in _step22_concept_match_summary(dato.get('conceptos') or {}, limit=80):
            vals = [row.get('servicio'), row.get('descripcion'), row.get('status'), row.get('score'), row.get('codigos'), row.get('matrix_rows'), row.get('reason')]
            for cidx, val in enumerate(vals,1):
                cell = ws.cell(r,cidx,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True) if cidx in (2,7) else _aln('center','center')
                if cidx == 4 and isinstance(val,(int,float)): cell.number_format='0.00%'
            r += 1
        for col,w in {'A':16,'B':58,'C':22,'D':12,'E':22,'F':14,'G':80}.items(): ws.column_dimensions[col].width=w
        # Mantener orden de tabs: Comparativa, Detalle, Matriz base CD, Analisis experto IA.
        desired = ['Comparativa','Detalle','Matriz base CD','Analisis experto IA']
        wb._sheets = [wb[n] for n in desired if n in wb.sheetnames] + [ws for ws in wb._sheets if ws.title not in desired]
        wb.save(output)
    except Exception as exc:
        # No romper el reporte por falla de dictamen IA.
        try:
            wb = openpyxl.load_workbook(output)
            ws = wb.create_sheet('Analisis experto IA')
            ws['A1'] = 'No se pudo generar analisis experto IA'
            ws['A2'] = f'{type(exc).__name__}: {exc}'
            wb.save(output)
        except Exception:
            pass
    return output


def _step18_build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP22_BUILD_PMD_PREV(output, meta, dato, provider_analysis)
    return _step22_add_expert_ai_sheet(result, dato, provider_analysis)

# Paso 22b - comparacion rapida de Matriz base CD para evitar que el reporte se bloquee.
# Mantiene match IA/conceptual, pero genera filas de matriz CD con una validacion ligera
# contra la matriz del proveedor. Esto es suficiente para que el tab no quede vacio y
# para mostrar omisiones/extras revisables.
def _step22_item_domain(item):
    raw = _normalize_text(str(item.get('tipo') or item.get('section') or item.get('domain') or ''))
    text = _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')} {raw}")
    if 'mano' in raw or 'mo' == raw or any(x in text for x in ['cuadrilla','peon','albanil','supervisor','oficial','ayudante','operador']):
        return 'mano_obra'
    if 'equipo' in raw or 'herramienta' in raw or any(x in text for x in ['maquinaria','renta','rompedora','camion','transporte']):
        return 'equipo'
    if 'basico' in raw or 'basicos' in raw:
        return 'basico'
    return 'material'


def _step22_provider_item_tokens(provider_concept):
    rows = []
    for item in _provider_matrix_items(provider_concept):
        desc = item.get('descripcion') or item.get('desc') or ''
        code = item.get('codigo') or ''
        toks = set(_tokenize_text(_normalize_text(f'{code} {desc}')))
        rows.append({'item': item, 'tokens': toks, 'domain': _step22_item_domain(item)})
    return rows


def _step22_fast_matrix_comparison(provider_key, provider_concept, match):
    rows = []
    findings = []
    selected_matches = (match or {}).get('selected_matches') or []
    if not selected_matches and (match or {}).get('selected'):
        selected_matches = [{'concept': match.get('selected'), 'peso': 1.0, 'motivo': 'seleccionado'}]
    provider_rows = _step22_provider_item_tokens(provider_concept)
    used_provider = set()
    cd_codes = []
    cd_descs = []
    for sm in selected_matches[:4]:
        cd = sm.get('concept') or {}
        if cd.get('codigo'):
            cd_codes.append(str(cd.get('codigo')))
        if cd.get('desc'):
            cd_descs.append(str(cd.get('desc')))
        for item in cd.get('matriz') or []:
            cd_text = _normalize_text(f"{item.get('codigo','')} {item.get('descripcion','')}")
            cd_tokens = set(_tokenize_text(cd_text))
            cd_domain = _step22_item_domain(item)
            best = None; best_score = 0.0; best_idx = None
            for idx, prow in enumerate(provider_rows):
                # Dominio compatible; evita MO contra equipo/EPP.
                if prow['domain'] != cd_domain:
                    continue
                shared = cd_tokens & prow['tokens']
                denom = max(len(cd_tokens), 1)
                score = len(shared) / denom
                # ayuda por descripcion parecida
                if score < 0.35:
                    score = max(score, SequenceMatcher(None, cd_text[:120], _normalize_text(f"{prow['item'].get('codigo','')} {prow['item'].get('descripcion','')}")[:120]).ratio() * 0.55)
                if score > best_score:
                    best_score = score; best = prow; best_idx = idx
            if best and best_score >= 0.35:
                used_provider.add(best_idx)
                rows.append({
                    'tipo': item.get('tipo') or cd_domain,
                    'cd_codigo': item.get('codigo'),
                    'cd_desc': item.get('descripcion'),
                    'unidad': item.get('unidad'),
                    'cantidad': item.get('cantidad') or item.get('volumen'),
                    'costo': item.get('precio') or item.get('costo'),
                    'importe': item.get('importe'),
                    'estado': 'match',
                    'proveedor_codigo': best['item'].get('codigo'),
                    'proveedor_desc': best['item'].get('descripcion'),
                    'score': round(best_score, 4),
                })
            else:
                rows.append({
                    'tipo': item.get('tipo') or cd_domain,
                    'cd_codigo': item.get('codigo'),
                    'cd_desc': item.get('descripcion'),
                    'unidad': item.get('unidad'),
                    'cantidad': item.get('cantidad') or item.get('volumen'),
                    'costo': item.get('precio') or item.get('costo'),
                    'importe': item.get('importe'),
                    'estado': 'faltante_en_proveedor',
                    'score': 0,
                })
                findings.append({'tipo': 'faltante_matriz_base', 'descripcion': item.get('descripcion'), 'codigo': item.get('codigo')})
    # Extras del proveedor no cubiertos por la matriz base CD.
    for idx, prow in enumerate(provider_rows):
        if idx in used_provider:
            continue
        it = prow['item']
        # No agregar encabezados sin costo.
        if not any(isinstance(it.get(k), (int, float)) for k in ['precio_base','factor','importe','cantidad']):
            continue
        rows.append({
            'tipo': it.get('_section') or prow['domain'],
            'cd_codigo': '',
            'cd_desc': '',
            'unidad': it.get('unidad'),
            'cantidad': it.get('factor') or it.get('cantidad'),
            'costo': it.get('precio_base'),
            'importe': it.get('importe'),
            'estado': 'sin_match_insumo',
            'proveedor_codigo': it.get('codigo'),
            'proveedor_desc': it.get('descripcion'),
            'score': 0,
        })
    return {
        'rows': rows,
        'findings': findings,
        'match_type': (match or {}).get('tipo_match') or (match or {}).get('status'),
        'score': (match or {}).get('score'),
        'construdata_codigo': ', '.join(cd_codes),
        'construdata_desc': ' + '.join(cd_descs[:3]),
        'reason': (match or {}).get('reason'),
    }


def attach_construdata_matrix_analysis(conceptos, precios_nacional):
    """Paso 22b: homologacion IA/fallback + matriz base CD rapida.

    Se evita depender de comparaciones pesadas. La calidad del match conceptual mejora con
    Voyage/Claude si las API keys estan activas; si no, usa reglas con trazabilidad.
    """
    cds = _get_construdata_concepts(precios_nacional)
    if not cds:
        return conceptos
    index = _build_construdata_token_index(cds)
    min_score = float(os.getenv('CONCEPT_MATRIX_MIN_SCORE', '0.22'))
    for key, c in (conceptos or {}).items():
        if not isinstance(c, dict) or str(key).startswith('__'):
            continue
        try:
            match = match_concepto_construdata(key, c, cds, min_score=min_score, concept_index=index)
            comp = _step22_fast_matrix_comparison(key, c, match)
            c['construdata_concept_match'] = match
            c['construdata_matrix_comparison'] = comp
            c['construdata_matrix_findings'] = comp.get('findings') or []
        except Exception as exc:
            c['construdata_concept_match'] = {'status': 'error_matrix_cd', 'score': None, 'reason': f'{type(exc).__name__}: {exc}', 'ai_status': get_concept_ai_status()}
            c['construdata_matrix_comparison'] = {'rows': [], 'findings': [], 'reason': str(exc)}
    return conceptos


# ============================================================
# STEP28_PERFORMANCE_DIAGNOSTICS
# IA apagada por defecto + tracking por etapas.
# Para reactivar IA: QUANTIA_ENABLE_AI=1.
# Para guardar trace JSONL: QUANTIA_TRACE_FILE=/tmp/job/trace.jsonl.
# ============================================================
import time as _q_time
import json as _q_json
import os as _q_os
import traceback as _q_traceback

# Por defecto apagamos IA para aislar performance. Si el entorno ya trae
# QUANTIA_ENABLE_AI=1, no se toca.
_q_os.environ.setdefault('QUANTIA_ENABLE_AI', '0')
if _q_os.environ.get('QUANTIA_ENABLE_AI', '0').strip().lower() not in {'1','true','yes','on'}:
    _q_os.environ.setdefault('ENABLE_AI_MATERIAL_MATCHING', '0')

def _q_ai_enabled():
    return _q_os.environ.get('QUANTIA_ENABLE_AI', '0').strip().lower() in {'1','true','yes','on'}

def _q_perf_log(event, **data):
    path = _q_os.getenv('QUANTIA_TRACE_FILE')
    payload = {'ts': _q_time.time(), 'event': event}
    payload.update(data)
    if path:
        try:
            with open(path, 'a', encoding='utf-8') as f:
                f.write(_q_json.dumps(payload, ensure_ascii=False, default=str) + '\n')
        except Exception:
            pass
    return payload

def _q_wrap_func(name):
    fn = globals().get(name)
    if not callable(fn) or getattr(fn, '_q_wrapped', False):
        return
    def wrapped(*args, **kwargs):
        t0 = _q_time.perf_counter()
        _q_perf_log('start', function=name)
        try:
            res = fn(*args, **kwargs)
            dt = _q_time.perf_counter() - t0
            extra = {}
            try:
                if name == 'extract_conceptos' and isinstance(res, dict): extra['conceptos'] = len([k for k in res if not str(k).startswith('__')])
                if name == 'extract_resumen_pu' and isinstance(res, list): extra['rows'] = len(res)
                if name == 'build_comparativo': extra['output'] = args[2] if len(args) >= 3 else kwargs.get('output')
            except Exception:
                pass
            _q_perf_log('end', function=name, seconds=round(dt, 4), **extra)
            return res
        except Exception as exc:
            dt = _q_time.perf_counter() - t0
            _q_perf_log('error', function=name, seconds=round(dt, 4), error=f'{type(exc).__name__}: {exc}', traceback=_q_traceback.format_exc()[-2000:])
            raise
    wrapped._q_wrapped = True
    globals()[name] = wrapped

for _q_name in [
    'extract_conceptos', 'extract_resumen_pu', 'apply_resumen_pu_to_conceptos',
    '_step08_apply_market_catalog_pricing', 'attach_construdata_matrix_analysis',
    '_step18_build_single_provider_pmd', '_step22_add_expert_ai_sheet', 'build_comparativo'
]:
    _q_wrap_func(_q_name)

# Evitar llamada Claude del dictamen mientras QUANTIA_ENABLE_AI=0.
try:
    _STEP28_CLAUDE_EXPERT_PREV = _step22_claude_expert_text
    def _step22_claude_expert_text(dato, provider_analysis=None):
        if not _q_ai_enabled():
            _q_perf_log('ai_skipped', provider='anthropic', operation='expert_text', reason='QUANTIA_ENABLE_AI=0')
            return None, 'IA desactivada por QUANTIA_ENABLE_AI=0'
        return _STEP28_CLAUDE_EXPERT_PREV(dato, provider_analysis)
except Exception:
    pass

_q_perf_log('module_loaded', module='processor', ai_enabled=_q_ai_enabled())

# -----------------------------------------------------------------------------
# STEP 30 - Politica profesional de IA: no consumo runtime por defecto
# -----------------------------------------------------------------------------
# Problema corregido:
# - Claude/Voyage se estaban llamando dentro de la generacion normal del Excel.
# - Eso multiplicaba costo y latencia por insumo/concepto, sin mejorar de forma confiable.
# Nueva regla:
# - La corrida normal usa catalogos + matching deterministico.
# - IA runtime solo se ejecuta si QUANTIA_ENABLE_AI=1 y QUANTIA_AI_RUNTIME_MODE=runtime_review
#   y el scope especifico esta habilitado.
try:
    from ai_runtime_policy import runtime_ai_allowed as _q_runtime_ai_allowed, external_ai_block_reason as _q_external_ai_block_reason, ai_policy_status as _q_ai_policy_status
except Exception:
    def _q_runtime_ai_allowed(scope='runtime'):
        return False
    def _q_external_ai_block_reason(scope='runtime'):
        return 'IA runtime desactivada.'
    def _q_ai_policy_status():
        return {'runtime_mode': 'off'}

# Override del helper legado: con esto _step08_match_market_item conserva el resultado deterministico.
def _step12_ai_enabled_for_materials():
    return bool(_q_runtime_ai_allowed('material_matching'))

# Override del dictamen Claude: no debe gastar tokens en la corrida normal.
def _step22_claude_expert_text(dato, provider_analysis=None):
    if not _q_runtime_ai_allowed('expert_text'):
        return None, _q_external_ai_block_reason('expert_text')
    try:
        from material_ai_matcher import _anthropic_message
        conceptos = dato.get('conceptos') or {}
        compact = []
        for row in _step22_concept_match_summary(conceptos, limit=15):
            compact.append(row)
        payload = {
            'rol': 'Eres un experto senior en análisis de precios unitarios, presupuestos de obra y revisión PMD.',
            'tarea': 'Redacta un dictamen ejecutivo breve, humano y técnico. No inventes datos; usa solo la información entregada.',
            'proveedor': dato.get('nombre'),
            'matches_matriz_cd': compact,
            'hallazgos': (provider_analysis or {}).get('drivers', [])[:10] if isinstance(provider_analysis, dict) else [],
            'respuesta': 'Texto en español, 4 a 8 bullets ejecutivos.'
        }
        txt = _anthropic_message(payload, max_tokens=700, timeout=25)
        return (txt.strip() if txt else None), ('Claude activo' if txt else 'Claude no devolvio texto')
    except Exception as exc:
        return None, f'Claude error: {type(exc).__name__}: {str(exc)[:200]}'

# Hook de estado para hojas/debug.
def quantia_ai_policy_status():
    return _q_ai_policy_status()

# -----------------------------------------------------------------------------
# STEP31 - IA profesional controlada: valor sin consumo descontrolado
# -----------------------------------------------------------------------------
# Diagnóstico corregido:
# - La IA no debe ejecutarse por cada insumo ni por cada concepto durante el cálculo.
# - El cálculo principal queda 100% determinístico: catálogos Construdata, Resumen PU,
#   cantidades reales, operaciones del contratista y fallback documentado.
# - La IA se usa solo como revisión experta controlada al final del reporte, con cache,
#   una llamada máxima por reporte y contexto de mayor impacto económico.

try:
    from quantia_ai_review import generate_controlled_ai_review, ai_runtime_policy as _step31_ai_runtime_policy
except Exception:
    generate_controlled_ai_review = None
    def _step31_ai_runtime_policy():
        return {'enabled': False, 'mode': 'unavailable'}

# Reglas duras: material/concept IA granular OFF salvo runtime_review explícito.
def _step12_ai_enabled_for_materials():
    try:
        from ai_runtime_policy import runtime_ai_allowed
        return bool(runtime_ai_allowed('material_matching'))
    except Exception:
        return False

# El dictamen Claude anterior queda sustituido por una revisión controlada batch.
def _step22_claude_expert_text(dato, provider_analysis=None):
    if generate_controlled_ai_review is None:
        return None, 'IA controlada no disponible'
    review = generate_controlled_ai_review(dato, provider_analysis)
    if review.get('used_ai') and review.get('status') == 'ok':
        parsed = review.get('parsed') or {}
        if parsed:
            bullets = []
            for b in parsed.get('resumen_ejecutivo') or []:
                bullets.append(f'- {b}')
            risks = parsed.get('riesgos') or []
            if risks:
                bullets.append('Riesgos principales:')
                for r in risks[:5]:
                    bullets.append(f"- {r.get('severidad','')}: {r.get('servicio','')} — {r.get('comentario','')}")
            recs = parsed.get('recomendaciones') or []
            if recs:
                bullets.append('Recomendaciones:')
                for r in recs[:5]:
                    bullets.append(f'- {r}')
            return '\n'.join(bullets), 'Claude controlado: 1 llamada batch/cached'
        return (review.get('raw_text') or '').strip(), 'Claude controlado: raw text'
    return None, review.get('reason') or 'IA controlada no ejecutada'

# Reemplaza la hoja Analisis experto IA para mostrar política, cache y contexto usado.
def _step31_add_expert_ai_sheet(output, dato, provider_analysis=None):
    try:
        wb = openpyxl.load_workbook(output)
        if 'Analisis experto IA' in wb.sheetnames:
            del wb['Analisis experto IA']
        ws = wb.create_sheet('Analisis experto IA')
        ws.sheet_view.showGridLines = False
        ws.column_dimensions['A'].width = 26
        ws.column_dimensions['B'].width = 100
        ws['A1'] = 'ANÁLISIS EXPERTO IA CONTROLADA'
        ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10)
        ws.merge_cells('A1:B1')

        review = generate_controlled_ai_review(dato, provider_analysis) if generate_controlled_ai_review else {'status':'unavailable','reason':'Modulo no disponible','policy':_step31_ai_runtime_policy()}
        context = review.get('context') or {}
        policy = review.get('policy') or _step31_ai_runtime_policy()
        dictamen, status = _step22_claude_expert_text(dato, provider_analysis)
        if not dictamen:
            dictamen = _step22_deterministic_expert_text(dato, provider_analysis) if '_step22_deterministic_expert_text' in globals() else 'Dictamen determinístico no disponible.'

        rows = [
            ('Estado IA', status),
            ('Modo', policy.get('mode')),
            ('IA usada', 'SI' if review.get('used_ai') else 'NO'),
            ('Cache', 'SI' if review.get('from_cache') else 'NO'),
            ('Política', json.dumps(policy, ensure_ascii=False, default=str)),
            ('Dictamen', dictamen),
        ]
        r = 3
        for k, v in rows:
            ws.cell(r,1,k).fill=_fill(C_GRIS_FIL); ws.cell(r,1).font=_fnt(bold=True,size=8); ws.cell(r,1).border=_brd(); ws.cell(r,1).alignment=_aln('left','top',wrap=True)
            ws.cell(r,2,v).font=_fnt(size=9); ws.cell(r,2).border=_brd(); ws.cell(r,2).alignment=_aln('left','top',wrap=True)
            ws.row_dimensions[r].height = 170 if k == 'Dictamen' else 40
            r += 1

        r += 1
        ws.cell(r,1,'Contexto enviado/revisado').fill=_fill(C_AZUL_OSC); ws.cell(r,1).font=_fnt(bold=True,color=C_BLANCO,size=8)
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
        r += 1
        headers = ['Tipo','Servicio','Descripción / Motivo','Importe / Score / Estado']
        for cidx,h in enumerate(headers,1):
            cell=ws.cell(r,cidx,h); cell.fill=_fill(C_AZUL_OSC); cell.font=_fnt(bold=True,color=C_BLANCO,size=8); cell.border=_brd(); cell.alignment=_aln('center','center',wrap=True)
        ws.column_dimensions['C'].width = 80
        ws.column_dimensions['D'].width = 38
        r += 1
        for c in (context.get('top_conceptos_por_impacto') or [])[:12]:
            vals=['Concepto impacto', c.get('servicio'), c.get('descripcion'), f"Proveedor={c.get('importe_contratista')} | Mercado={c.get('importe_mercado')} | Match={c.get('match_cd_status')}"]
            for cidx,val in enumerate(vals,1):
                cell=ws.cell(r,cidx,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True)
            r += 1
        for m in (context.get('materiales_a_revisar') or [])[:12]:
            vals=['Material revisar', m.get('servicio'), m.get('insumo'), f"Estado={m.get('match_status')} | CD={m.get('candidato_cd')} | Conf={m.get('confianza')}"]
            for cidx,val in enumerate(vals,1):
                cell=ws.cell(r,cidx,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True)
            r += 1
        desired=['Comparativa','Detalle','Matriz base CD','Analisis experto IA']
        wb._sheets=[wb[n] for n in desired if n in wb.sheetnames] + [s for s in wb._sheets if s.title not in desired]
        wb.save(output)
    except Exception as exc:
        _q_perf_log('step31_ai_sheet_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
    return output

# Sustituye wrapper de step22: agrega hoja controlada, sin consumo granular.
_STEP31_BUILD_PMD_PREV = _STEP22_BUILD_PMD_PREV if '_STEP22_BUILD_PMD_PREV' in globals() else globals().get('_step18_build_single_provider_pmd')
def _step18_build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP31_BUILD_PMD_PREV(output, meta, dato, provider_analysis)
    return _step31_add_expert_ai_sheet(result, dato, provider_analysis)

# Estado de política expuesto para debug.
def quantia_ai_policy_status():
    return _step31_ai_runtime_policy()

# -----------------------------------------------------------------------------
# Paso 32 - Interpretacion profesional de matriz PU + unidad compatible estricta
# -----------------------------------------------------------------------------
# Contexto funcional:
# - El precio de mercado granular NO debe usarse si el match de catalogo tiene unidad
#   incompatible con la unidad declarada por el contratista. Ejemplo critico:
#   CAMION DE VOLTEO 7 M3 declarado como VIAJE no puede tomar un catalogo PZA.
# - En ese caso se conserva el costo del contratista como fallback de mercado y se
#   marca SIN_MATCH_UNIDAD para evitar diferencias absurdas.
# - Los renglones financieros se recalculan en cascada usando los porcentajes
#   declarados por el contratista sobre el subtotal previo de mercado.
# - Los porcentuales tipo %MO, EPP, SHE, HERRAMIENTA MENOR se calculan con la suma
#   completa de mano de obra de mercado del servicio, no por cuadrilla individual.

_STEP32_MATCH_MARKET_ITEM_PREV = _step08_match_market_item


def _step32_unit_group(unit):
    u = normalize_unit(unit)
    if not u:
        return ''
    groups = {
        'pieza': {'pza', 'pieza', 'pz', 'pzas'},
        'viaje': {'viaje', 'viajes', 'vje'},
        'jornada': {'jor', 'jorn', 'jornal', 'jornada', 'dia', 'dias'},
        'hora': {'hr', 'hrs', 'hora', 'horas'},
        'metro': {'m', 'ml', 'metro', 'metros'},
        'metro2': {'m2', 'm²', 'metro2', 'metros2'},
        'metro3': {'m3', 'm³', 'metro3', 'metros3'},
        'kg': {'kg', 'kilo', 'kilogramo', 'kilogramos'},
        'ton': {'ton', 'tons', 'tonelada', 'toneladas'},
        'litro': {'lt', 'l', 'litro', 'litros'},
        'saco': {'saco', 'sacos'},
        'porcentaje': {'%', 'porc', 'porcentaje'},
        'lote': {'lote', 'lotes'},
        'servicio': {'serv', 'servicio', 'servicios'},
    }
    for g, vals in groups.items():
        if u in vals:
            return g
    return u


def _step32_units_compatible(provider_unit, market_unit):
    """Compatibilidad conservadora de unidades para evitar falsos matches.

    Regla base: si ambas unidades existen y no pertenecen al mismo grupo, el match
    se descarta. Solo se permiten equivalencias duras y seguras. Conversiones como
    TON<->KG o PZA<->VIAJE NO se convierten automaticamente porque el precio del
    catalogo no necesariamente corresponde al mismo alcance comercial.
    """
    pu = _step32_unit_group(provider_unit)
    mu = _step32_unit_group(market_unit)
    if not pu or not mu:
        # Si falta una unidad, no descartamos por unidad; queda a score/fallback.
        return True, 'unidad_no_informada'
    if pu == mu:
        return True, 'unidad_equivalente'
    # Equivalencias muy seguras por nomenclatura, no por conversion de cantidad.
    safe_pairs = {
        ('metro', 'metro'), ('metro2', 'metro2'), ('metro3', 'metro3'),
        ('jornada', 'jornada'), ('hora', 'hora'), ('porcentaje', 'porcentaje'),
    }
    if (pu, mu) in safe_pairs:
        return True, 'unidad_equivalente_segura'
    return False, f'unidad_incompatible: proveedor={pu} mercado={mu}'


def _step32_reject_incompatible_unit_match(item, match):
    if not isinstance(match, dict):
        return match
    selected = match.get('selected')
    if not selected:
        return match
    if _step08_is_percent_item(item):
        return match
    ok, reason = _step32_units_compatible(item.get('unidad'), selected.get('unidad'))
    if ok:
        if isinstance(selected, dict):
            selected = dict(selected)
            selected['unit_check'] = reason
            match = dict(match)
            match['selected'] = selected
        return match
    rejected = dict(selected)
    rejected['rejected_reason'] = reason
    new_match = dict(match)
    new_match['selected'] = None
    new_match['status'] = 'sin_match_unidad'
    new_match['confidence'] = 0.0
    new_match['unit_rejected_candidate'] = rejected
    cands = list(new_match.get('candidates') or [])
    if cands:
        cands = [rejected] + cands[:7]
    else:
        cands = [rejected]
    new_match['candidates'] = cands
    return new_match


def _step08_match_market_item(item, catalogs):
    base = _STEP32_MATCH_MARKET_ITEM_PREV(item, catalogs)
    return _step32_reject_incompatible_unit_match(item, base)


_STEP32_APPLY_MARKET_PREV = _step08_apply_market_catalog_pricing


def _step32_recalculate_concept_market_totals(c):
    direct_market = 0.0
    direct_provider = 0.0
    matched = 0
    total = 0
    subtotals = {'material': 0.0, 'mano_obra': 0.0, 'equipo': 0.0, 'basicos': 0.0}
    for item in c.get('granular_market_items') or []:
        total += 1
        if not item.get('fallback_contratista'):
            matched += 1
        imp_m = item.get('importe_mercado')
        imp_p = item.get('importe_proveedor')
        if isinstance(imp_m, (int, float)):
            direct_market += float(imp_m)
            dom = item.get('domain') or item.get('section') or 'material'
            dom = _normalize_text(dom)
            if 'mano' in dom:
                k = 'mano_obra'
            elif 'equipo' in dom or 'herramienta' in dom:
                k = 'equipo'
            elif 'basico' in dom:
                k = 'basicos'
            else:
                k = 'material'
            if not item.get('is_percent_item'):
                subtotals[k] = subtotals.get(k, 0.0) + float(imp_m)
        if isinstance(imp_p, (int, float)):
            direct_provider += float(imp_p)
    c['granular_market_direct'] = round(direct_market, 6) if direct_market else None
    c['granular_provider_direct'] = round(direct_provider, 6) if direct_provider else None
    c['granular_market_match_coverage'] = (matched / total) if total else None
    c['granular_market_subtotals'] = {k: round(v, 6) for k, v in subtotals.items()}
    return c


def _step08_apply_market_catalog_pricing(conceptos):
    conceptos = _STEP32_APPLY_MARKET_PREV(conceptos)
    for _, c in (conceptos or {}).items():
        for item in c.get('granular_market_items') or []:
            status = str(item.get('match_status') or '').lower()
            if status == 'sin_match_unidad':
                provider_price = item.get('precio_proveedor')
                qty = item.get('cantidad')
                op = item.get('op') or '*'
                market_importe = _step11_operation_amount(provider_price, qty, op)
                item['precio_mercado'] = provider_price
                item['precio_mercado_catalogo'] = provider_price
                item['importe_mercado'] = market_importe if isinstance(market_importe, (int, float)) else item.get('importe_proveedor')
                item['delta_precio_pct'] = 0.0
                item['fallback_contratista'] = True
                item['confidence'] = 0.0
                rejected = None
                tops = item.get('top_candidates') or []
                if tops:
                    rejected = tops[0]
                obs = 'Match descartado por unidad incompatible; costo mercado = costo contratista.'
                if rejected and rejected.get('unidad'):
                    obs += f" Candidato descartado: {rejected.get('codigo') or ''} {rejected.get('descripcion') or ''} unidad {rejected.get('unidad')}"
                item['match_reason'] = obs
        _step32_recalculate_concept_market_totals(c)
    return conceptos


# Finanzas: mercado en cascada usando porcentajes declarados por el contratista.
def _step32_pct_from_amount(amount, base):
    if isinstance(amount, (int, float)) and isinstance(base, (int, float)) and base:
        pct = float(amount) / float(base)
        if 0 <= pct <= 2.0:
            return pct
    return 0.0


def _step13_financial_rows(c):
    provider_direct = _step12_provider_direct(c)
    market_direct = _step12_market_direct(c)
    pu_provider = c.get('pu') if isinstance(c.get('pu'), (int, float)) else None

    provider_indirect, provider_indirect_pct = _step12_provider_indirect(c)
    if not isinstance(provider_indirect_pct, (int, float)):
        provider_indirect_pct = _step32_pct_from_amount(provider_indirect, provider_direct)
    provider_indirect = float(provider_indirect or 0.0)
    market_indirect = market_direct * provider_indirect_pct if market_direct else 0.0

    provider_after_indirect = provider_direct + provider_indirect
    market_after_indirect = market_direct + market_indirect

    financiamiento = c.get('financiamiento') if isinstance(c.get('financiamiento'), (int, float)) else 0.0
    financing_pct = _step32_pct_from_amount(financiamiento, provider_after_indirect)
    market_financing = market_after_indirect * financing_pct if market_after_indirect else 0.0

    provider_after_financing = provider_after_indirect + financiamiento
    market_after_financing = market_after_indirect + market_financing

    utilidad = c.get('utilidad') if isinstance(c.get('utilidad'), (int, float)) else 0.0
    utility_pct = _step32_pct_from_amount(utilidad, provider_after_financing)
    market_utility = market_after_financing * utility_pct if market_after_financing else 0.0

    # Otros cargos no identificados: si el PU declarado trae remanente positivo,
    # lo mostramos como ajuste/cargo adicional y aplicamos el mismo porcentaje sobre mercado.
    provider_known_total = provider_after_financing + utilidad
    otros = 0.0
    otros_pct = 0.0
    if isinstance(pu_provider, (int, float)) and pu_provider > provider_known_total + 0.01:
        otros = pu_provider - provider_known_total
        otros_pct = _step32_pct_from_amount(otros, provider_known_total)
    market_otros = (market_after_financing + market_utility) * otros_pct if otros_pct else 0.0

    market_total = market_after_financing + market_utility + market_otros
    return [
        ('COSTO DIRECTO', provider_direct, market_direct, None, 'Suma de materiales, mano de obra, equipo/herramienta y básicos. Mercado usa precios Construdata compatibles por unidad.'),
        ('COSTO INDIRECTO', provider_indirect, market_indirect, provider_indirect_pct, 'Mercado usa el % indirecto declarado por contratista sobre costo directo mercado.'),
        ('FINANCIAMIENTO', financiamiento, market_financing, financing_pct, 'Mercado usa el % declarado sobre subtotal mercado previo.'),
        ('UTILIDAD / CARGOS ADICIONALES', utilidad, market_utility, utility_pct, 'Mercado usa el % utilidad declarado sobre subtotal mercado previo.'),
        ('OTROS CARGOS NO IDENTIFICADOS', otros, market_otros, otros_pct, 'Remanente positivo entre PU y conceptos financieros detectados; se replica como % sobre mercado si existe.'),
        ('TOTAL COSTO UNITARIO', pu_provider, market_total, None, 'Total contratista vs mercado recalculado con la misma estructura porcentual declarada.'),
    ]


def _step13_market_pu(c):
    rows = _step13_financial_rows(c)
    total = rows[-1][2]
    return total if isinstance(total, (int, float)) and total > 0 else 0.0


def _market_pu_with_indirect(concepto):
    return _step13_market_pu(concepto)

# Paso 32b - normalizacion robusta de unidades con superindices antes de normalize_unit.
_STEP32_UNIT_GROUP_PREV = _step32_unit_group

def _step32_unit_group(unit):
    raw = str(unit or '').strip().lower()
    raw = raw.replace(' ', '')
    if raw in {'m²', 'm2', 'm^2', 'mt2', 'mts2'}:
        return 'metro2'
    if raw in {'m³', 'm3', 'm^3', 'mt3', 'mts3'}:
        return 'metro3'
    if raw in {'ml', 'm.l.', 'm.l'}:
        return 'metro'
    return _STEP32_UNIT_GROUP_PREV(unit)

# ============================================================
# PASO 33 - Matriz base CD construida por concepto con IA conceptual controlada
# ============================================================
# Objetivo:
# - Sustituir el tab anterior "Matriz base CD" por una matriz base construida
#   desde Construdata a partir de la descripcion del servicio del contratista.
# - Permitir 1 servicio proveedor -> 1 o N conceptos Construdata.
# - Usar IA solo en matching conceptual/matriz, no por insumo.
# - Mantener el calculo de precios de mercado en tab Detalle basado en catalogos
#   granulares de materiales/MO/equipo.


def _step33_match_type_label(comp, match):
    return (comp or {}).get('match_type') or (match or {}).get('tipo_match') or (match or {}).get('status') or 'SIN_MATCH'


def _step33_selected_cd_summary(match):
    matches = (match or {}).get('selected_matches') or []
    if not matches and (match or {}).get('selected'):
        matches = [{'concept': match.get('selected'), 'peso': 1.0, 'motivo': (match or {}).get('reason') or ''}]
    rows = []
    for sm in matches[:6]:
        cd = sm.get('concept') or {}
        rows.append({
            'codigo': cd.get('codigo') or cd.get('clave'),
            'descripcion': cd.get('desc_larga') or cd.get('desc'),
            'unidad': cd.get('unidad'),
            'pu': cd.get('pu_mercado'),
            'peso': sm.get('peso', 1.0),
            'motivo': sm.get('motivo') or '',
            'matriz_count': len(cd.get('matriz') or []),
        })
    return rows


def _step33_matrix_base_findings_for_service(c):
    comp = c.get('construdata_matrix_comparison') or {}
    rows = comp.get('rows') or []
    faltantes = sum(1 for r in rows if r.get('estado') == 'faltante_en_proveedor')
    extras = sum(1 for r in rows if r.get('estado') == 'sin_match_insumo')
    matched = sum(1 for r in rows if r.get('estado') == 'match')
    return matched, faltantes, extras


def _step33_build_matrix_base_sheet(wb, dato):
    """Nuevo tab Matriz base CD.

    En lugar de presentarlo como una comparacion plana, este tab muestra:
    1) decision concepto proveedor -> concepto(s) Construdata;
    2) matriz base creada por la union de las matrices Construdata seleccionadas;
    3) estatus contra lo declarado por el contratista: match, faltante o extra;
    4) motivo/trazabilidad de la seleccion.
    """
    if 'Matriz base CD' in wb.sheetnames:
        del wb['Matriz base CD']
    ws = wb.create_sheet('Matriz base CD')
    ws.sheet_view.showGridLines = False

    title = 'MATRIZ BASE CONSTRUDATA CONSTRUIDA DESDE LA DESCRIPCION DEL SERVICIO'
    ws.merge_cells('A1:R1')
    ws['A1'] = title
    ws['A1'].fill = _fill(C_AZUL_OSC)
    ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=10)
    ws['A1'].alignment = _aln('center', 'center', wrap=True)

    ws.merge_cells('A2:R2')
    ws['A2'] = (
        'Uso del tab: validar alcance tecnico. El precio de mercado se calcula en Detalle por catalogos granulares; '
        'esta matriz base sirve para detectar omisiones, elementos adicionales y conceptos compuestos 1->N.'
    )
    ws['A2'].fill = _fill('EAF2F8')
    ws['A2'].font = _fnt(size=8)
    ws['A2'].alignment = _aln('left', 'center', wrap=True)

    headers = [
        'Servicio proveedor', 'Descripcion proveedor', 'Tipo match', 'Score', 'Requiere revision',
        'Conceptos CD seleccionados', 'Motivo seleccion / IA', 'Tipo fila',
        'Codigo concepto CD', 'Concepto CD', 'Peso concepto',
        'Tipo elemento CD', 'Codigo insumo CD', 'Insumo base CD', 'Unidad',
        'Cantidad/Base', 'Costo/Base', 'Importe ref.', 'Estado vs contratista', 'Nota tecnica'
    ]
    header_row = 4
    for col, h in enumerate(headers, 1):
        cell = ws.cell(header_row, col, h)
        cell.fill = _fill(C_AZUL_OSC)
        cell.font = _fnt(bold=True, color=C_BLANCO, size=8)
        cell.border = _brd()
        cell.alignment = _aln('center', 'center', wrap=True)

    row = header_row + 1
    conceptos = dato.get('conceptos') or {}
    for clave, c in conceptos.items():
        if not isinstance(c, dict) or str(clave).startswith('__'):
            continue
        match = c.get('construdata_concept_match') or {}
        comp = c.get('construdata_matrix_comparison') or {}
        selected_summary = _step33_selected_cd_summary(match)
        selected_codes = ', '.join([str(x.get('codigo')) for x in selected_summary if x.get('codigo')])
        selected_desc = ' + '.join([str(x.get('descripcion'))[:80] for x in selected_summary if x.get('descripcion')])
        reason = match.get('reason') or comp.get('reason') or ''
        if match.get('claude_decision'):
            cd_reason = match.get('claude_decision') or {}
            reason = (cd_reason.get('explicacion') or reason or 'Decision validada por IA conceptual')
        matched, faltantes, extras = _step33_matrix_base_findings_for_service(c)

        # Bloque resumen por servicio.
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        status_txt = _step33_match_type_label(comp, match)
        summary = (
            f'{clave} | {c.get("desc") or ""} | {status_txt} | '
            f'CD: {selected_codes or "SIN CANDIDATO"} | '
            f'Match insumos: {matched}, faltantes base: {faltantes}, extras contratista: {extras}'
        )
        cell = ws.cell(row, 1, summary)
        cell.fill = _fill(C_GRIS_SEC)
        cell.font = _fnt(bold=True, size=8)
        cell.border = _brd()
        cell.alignment = _aln('left', 'center', wrap=True)
        row += 1

        if not selected_summary and not (comp.get('rows') or []):
            vals = [
                clave, c.get('desc'), 'SIN_MATCH', match.get('score'), True,
                '', reason or 'No se encontro concepto/matriz Construdata confiable.', 'SIN_MATRIZ',
                '', '', '', '', '', '', '', '', '', '', 'SIN_MATRIZ_BASE',
                'No usar matriz CD; revisar alcance o construir candidato manual.'
            ]
            for col, val in enumerate(vals, 1):
                cc = ws.cell(row, col, val); cc.border = _brd(); cc.font = _fnt(size=8); cc.alignment = _aln('left','center',wrap=True) if col in (2,7,20) else _aln('center','center')
                if col == 4 and isinstance(val, (int,float)): cc.number_format = '0.00%'
            for col in range(1, len(headers)+1): ws.cell(row, col).fill = _fill('FCE4D6')
            row += 1
            continue

        # Fila de conceptos seleccionados (1->N).
        for idx, s in enumerate(selected_summary, start=1):
            vals = [
                clave, c.get('desc'), status_txt, match.get('score'), bool(match.get('requires_review')),
                selected_codes, reason, 'CONCEPTO_CD',
                s.get('codigo'), s.get('descripcion'), s.get('peso'),
                '', '', '', s.get('unidad'), '', s.get('pu'), '', 'CONCEPTO_BASE',
                s.get('motivo') or 'Concepto seleccionado para construir matriz base.'
            ]
            for col, val in enumerate(vals, 1):
                cc = ws.cell(row, col, val); cc.border = _brd(); cc.font = _fnt(size=8)
                cc.alignment = _aln('left','center',wrap=True) if col in (2,6,7,10,14,20) else _aln('center','center')
                if col in (4,11) and isinstance(val, (int,float)): cc.number_format = '0.00%'
                if col == 17 and isinstance(val, (int,float)): cc.number_format = '$#,##0.00'
            for col in range(1, len(headers)+1): ws.cell(row, col).fill = _fill('D9EAF7')
            row += 1

        # Filas de matriz base CD, ya uniendo 1 o N conceptos seleccionados.
        for r in comp.get('rows') or []:
            estado = r.get('estado') or ''
            nota = ''
            if estado == 'faltante_en_proveedor':
                nota = 'Elemento recomendado por la matriz base Construdata y no localizado en la matriz del contratista.'
            elif estado == 'sin_match_insumo':
                nota = 'Elemento declarado por contratista fuera de la matriz base; revisar si es alcance adicional o concepto compuesto.'
            elif estado == 'match':
                nota = 'Elemento del contratista homologado contra la matriz base Construdata.'
            else:
                nota = 'Fila de referencia de matriz base.'
            vals = [
                clave, c.get('desc'), status_txt, match.get('score'), bool(match.get('requires_review')),
                selected_codes, reason, 'INSUMO_BASE_CD',
                r.get('source_concept_code') or r.get('codigo_concepto_cd') or comp.get('construdata_codigo'),
                r.get('source_concept_desc') or comp.get('construdata_desc'),
                '',
                r.get('tipo'), r.get('cd_codigo') or r.get('provider_codigo'), r.get('cd_desc') or r.get('provider_desc'),
                r.get('cd_unidad') or r.get('provider_unidad') or r.get('unidad'),
                r.get('cd_cantidad') or r.get('provider_cantidad') or r.get('cantidad'),
                r.get('cd_precio') or r.get('provider_precio') or r.get('costo'),
                r.get('cd_importe') or r.get('provider_importe') or r.get('importe'),
                estado,
                nota,
            ]
            for col, val in enumerate(vals, 1):
                cc = ws.cell(row, col, val); cc.border = _brd(); cc.font = _fnt(size=8)
                cc.alignment = _aln('left','center',wrap=True) if col in (2,6,7,10,14,20) else _aln('center','center')
                if col == 4 and isinstance(val, (int,float)): cc.number_format = '0.00%'
                if col in (17,18) and isinstance(val, (int,float)): cc.number_format = '$#,##0.00'
            if estado == 'faltante_en_proveedor':
                for col in range(1, len(headers)+1): ws.cell(row, col).fill = _fill(C_AMARILLO)
            elif estado == 'sin_match_insumo':
                for col in range(1, len(headers)+1): ws.cell(row, col).fill = _fill('FCE4D6')
            row += 1
        row += 1

    widths = {
        'A':16, 'B':58, 'C':22, 'D':10, 'E':14, 'F':28, 'G':70, 'H':16,
        'I':18, 'J':48, 'K':12, 'L':18, 'M':18, 'N':48, 'O':10,
        'P':12, 'Q':14, 'R':14, 'S':20, 'T':70,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    try:
        ws.freeze_panes = 'A5'
        ws.auto_filter.ref = f'A4:T{max(row-1,4)}'
    except Exception:
        pass
    return ws

# Reemplaza el builder usado por _build_single_provider_pmd.
_step08_build_matrix_base_sheet = _step33_build_matrix_base_sheet


# ============================================================
# PASO 34 - Matriz Recomendada de Mercado / Comparacion estructural
# ============================================================
# Principio funcional:
# - La matriz declarada por el contratista NO se modifica por IA.
# - La IA conceptual solo ayuda a seleccionar 1..N conceptos Construdata que representan
#   el servicio; con esas matrices se construye una matriz recomendada de mercado.
# - La comparacion estructural detecta coincidencias, faltantes de mercado y adicionales
#   del contratista. El precio de mercado del tab Detalle sigue saliendo del catalogo
#   granular y de reglas duras de unidad/tipo/fallback.


def _step34_norm_desc(text):
    try:
        return _normalize_text(text or '')
    except Exception:
        import re as _re
        return _re.sub(r'\s+', ' ', str(text or '').lower()).strip()


def _step34_item_domain(item):
    raw = _step34_norm_desc(item.get('tipo_kind') or item.get('tipo') or item.get('TipoInsumo') or '')
    desc = _step34_norm_desc(item.get('descripcion') or item.get('desc') or item.get('provider_desc') or item.get('cd_desc') or '')
    code = _step34_norm_desc(item.get('codigo') or item.get('clave') or item.get('provider_codigo') or item.get('cd_codigo') or '')
    txt = ' '.join([raw, desc, code])
    if any(x in txt for x in ['mano obra','mano_obra','cuadrilla','peon','oficial','ayudante','topografo','supervisor','cadenero','albanil']):
        return 'MO'
    if any(x in txt for x in ['equipo','maquinaria','herramienta','camion','retro','grua','bomba','compresor','transporte']):
        return 'EQUIPO_HERRAMIENTA'
    if any(x in txt for x in ['basico','basicos']):
        return 'BASICO'
    if any(x in txt for x in ['%mo','epp','she','seguridad','herramienta menor']):
        return 'PORCENTAJE_MO'
    return 'MATERIAL'


def _step34_item_signature(item):
    """Firma conservadora: dominio + tokens clave + unidad.
    No se usa IA para sustituir lo declarado; solo para agrupar equivalentes dentro de CD.
    """
    desc = _step34_norm_desc(item.get('descripcion') or item.get('desc') or item.get('provider_desc') or item.get('cd_desc') or '')
    code = _step34_norm_desc(item.get('codigo') or item.get('clave') or item.get('provider_codigo') or item.get('cd_codigo') or '')
    unit = normalize_unit(item.get('unidad') or item.get('provider_unidad') or item.get('cd_unidad') or '') if 'normalize_unit' in globals() else _step34_norm_desc(item.get('unidad') or '')
    domain = _step34_item_domain(item)
    tokens = [t for t in _concept_tokens(desc) if t not in {'incluye','suministro','colocacion','obra','material','equipo','herramienta'}]
    key_tokens = ' '.join(tokens[:7]) or code or desc[:60]
    return f'{domain}|{unit}|{key_tokens}'


def _step34_cd_recommended_items(match):
    """Une matrices CD de los conceptos seleccionados 1..N y elimina duplicados equivalentes."""
    selected = (match or {}).get('selected_matches') or []
    if not selected and (match or {}).get('selected'):
        selected = [{'concept': match.get('selected'), 'peso': 1.0, 'motivo': (match or {}).get('reason') or ''}]
    by_sig = {}
    for sm in selected:
        cd = sm.get('concept') or {}
        cd_code = cd.get('codigo') or cd.get('clave')
        cd_desc = cd.get('desc_larga') or cd.get('desc')
        peso = sm.get('peso', 1.0)
        motivo = sm.get('motivo') or ''
        for raw in cd.get('matriz') or []:
            item = dict(raw)
            item['source_concept_code'] = cd_code
            item['source_concept_desc'] = cd_desc
            item['source_concept_weight'] = peso
            item['source_concept_reason'] = motivo
            item['domain'] = _step34_item_domain(item)
            sig = _step34_item_signature(item)
            existing = by_sig.get(sig)
            if existing is None:
                item['source_concepts'] = [cd_code] if cd_code else []
                by_sig[sig] = item
            else:
                # Consolidar fuentes y conservar la fila con mayor importe/costo como representativa.
                sc = existing.setdefault('source_concepts', [])
                if cd_code and cd_code not in sc:
                    sc.append(cd_code)
                try:
                    new_amt = float(item.get('importe') or 0)
                    old_amt = float(existing.get('importe') or 0)
                    if new_amt > old_amt:
                        item['source_concepts'] = sc
                        by_sig[sig] = item
                except Exception:
                    pass
    return list(by_sig.values())


def _step34_provider_declared_items(concept):
    items = []
    for p in _provider_matrix_items(concept):
        item = dict(p)
        item['domain'] = _step34_item_domain(item)
        items.append(item)
    return items


def _step34_match_provider_to_market(provider_items, market_items):
    used_market = set()
    rows = []
    matched_market_idx = set()
    for p in provider_items:
        best_idx = None; best = None; best_score = 0.0
        for idx, m in enumerate(market_items):
            if idx in used_market:
                continue
            # Dominio fuerte: no empatar MO con equipo, material con MO, etc.
            if _step34_item_domain(p) != _step34_item_domain(m):
                continue
            sc = _component_match_score(p, m) if '_component_match_score' in globals() else 0.0
            if sc > best_score:
                best_score = sc; best_idx = idx; best = m
        if best is not None and best_score >= 0.30:
            used_market.add(best_idx); matched_market_idx.add(best_idx)
            rows.append({'estado':'COINCIDE', 'provider':p, 'market':best, 'score':best_score})
        else:
            rows.append({'estado':'ADICIONAL_CONTRATISTA', 'provider':p, 'market':None, 'score':best_score})
    for idx, m in enumerate(market_items):
        if idx not in matched_market_idx:
            rows.append({'estado':'FALTANTE_MERCADO', 'provider':None, 'market':m, 'score':None})
    return rows


def _step34_amount(item):
    if not item:
        return None
    for k in ['importe','cd_importe','provider_importe']:
        v = item.get(k)
        if isinstance(v, (int, float)):
            return float(v)
    return None


def _step34_structural_summary(rows):
    total_market = 0.0; matched_market = 0.0; market_count = 0; matched_count = 0
    extras = 0; missing = 0
    for r in rows:
        m = r.get('market')
        if m is not None:
            market_count += 1
            amt = _step34_amount(m)
            if isinstance(amt, (int, float)):
                total_market += max(amt, 0)
        if r.get('estado') == 'COINCIDE':
            matched_count += 1
            amt = _step34_amount(m)
            if isinstance(amt, (int, float)):
                matched_market += max(amt, 0)
        elif r.get('estado') == 'FALTANTE_MERCADO':
            missing += 1
        elif r.get('estado') == 'ADICIONAL_CONTRATISTA':
            extras += 1
    cobertura_tecnica = matched_count / market_count if market_count else None
    cobertura_economica = matched_market / total_market if total_market else cobertura_tecnica
    return {
        'mercado_elementos': market_count,
        'coincidencias': matched_count,
        'faltantes': missing,
        'adicionales': extras,
        'cobertura_tecnica': cobertura_tecnica,
        'cobertura_economica': cobertura_economica,
    }


def _step34_build_recommended_market_matrix_sheet(wb, dato):
    # Reemplazo formal del tab anterior.
    for name in ['Matriz base CD', 'Matriz Recomendada Mercado']:
        if name in wb.sheetnames:
            del wb[name]
    ws = wb.create_sheet('Matriz Recomendada Mercado')
    ws.sheet_view.showGridLines = False
    headers = [
        'Servicio', 'Descripcion servicio', 'Tipo fila', 'Tipo match', 'Score concepto',
        'Conceptos CD usados', 'Motivo seleccion', 'Cobertura tecnica', 'Cobertura economica',
        'Estado estructural', 'Dominio', 'Contratista codigo', 'Contratista elemento', 'Contratista unidad',
        'Contratista cantidad', 'Contratista costo', 'Contratista importe',
        'CD codigo', 'CD elemento recomendado', 'CD unidad', 'CD cantidad/base', 'CD costo/base', 'CD importe ref.',
        'Concepto fuente CD', 'Score insumo', 'Nota experta'
    ]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(1,1,'MATRIZ RECOMENDADA DE MERCADO - COMPARACION ESTRUCTURAL CONTRA CONTRATISTA')
    ws.cell(1,1).fill=_fill(C_AZUL_OSC); ws.cell(1,1).font=_fnt(bold=True,color=C_BLANCO,size=10); ws.cell(1,1).alignment=_aln('center','center',wrap=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.cell(2,1,'Este tab NO sustituye los precios calculados en Detalle. Construye una matriz recomendada desde 1..N conceptos Construdata y detecta coincidencias, faltantes y elementos adicionales del contratista.')
    ws.cell(2,1).fill=_fill('EAF2F8'); ws.cell(2,1).font=_fnt(size=8); ws.cell(2,1).alignment=_aln('left','center',wrap=True)
    header_row=4
    for col,h in enumerate(headers,1):
        c=ws.cell(header_row,col,h); c.fill=_fill(C_AZUL_OSC); c.font=_fnt(bold=True,color=C_BLANCO,size=8); c.border=_brd(); c.alignment=_aln('center','center',wrap=True)
    row=header_row+1
    for clave, c in (dato.get('conceptos') or {}).items():
        if not isinstance(c, dict) or str(clave).startswith('__'):
            continue
        match = c.get('construdata_concept_match') or {}
        market_items = _step34_cd_recommended_items(match)
        provider_items = _step34_provider_declared_items(c)
        comparison_rows = _step34_match_provider_to_market(provider_items, market_items) if market_items else []
        summary = _step34_structural_summary(comparison_rows)
        selected = (match or {}).get('selected_matches') or []
        if not selected and (match or {}).get('selected'):
            selected = [{'concept': match.get('selected'), 'peso':1.0, 'motivo':match.get('reason') or ''}]
        codes = []
        reasons = []
        for sm in selected[:6]:
            cd=sm.get('concept') or {}
            code=cd.get('codigo') or cd.get('clave')
            if code: codes.append(str(code))
            if sm.get('motivo'): reasons.append(str(sm.get('motivo')))
        code_txt=', '.join(codes)
        reason = (match.get('reason') or '; '.join(reasons) or 'Seleccion por reglas/candidatos Construdata')
        tipo_match = match.get('tipo_match') or match.get('status') or 'SIN_MATCH'
        cov_t = summary.get('cobertura_tecnica')
        cov_e = summary.get('cobertura_economica')
        # resumen por servicio
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        txt=(f'{clave} | {c.get("desc") or ""} | {tipo_match} | CD={code_txt or "SIN CANDIDATO"} | '
             f'Cobertura tecnica={cov_t:.1%}' if isinstance(cov_t,(int,float)) else f'{clave} | {c.get("desc") or ""} | {tipo_match} | CD={code_txt or "SIN CANDIDATO"}')
        ws.cell(row,1,txt); ws.cell(row,1).fill=_fill(C_GRIS_SEC); ws.cell(row,1).font=_fnt(bold=True,size=8); ws.cell(row,1).border=_brd(); ws.cell(row,1).alignment=_aln('left','center',wrap=True)
        row += 1
        if not market_items:
            vals=[clave,c.get('desc'),'SIN_MATRIZ_RECOMENDADA',tipo_match,match.get('score'),code_txt,reason,cov_t,cov_e,'SIN_MATCH','','','','','','','','','','','','','','','', 'No se pudo construir matriz recomendada; requiere seleccion manual o mejorar base CD.']
            for col,val in enumerate(vals,1):
                cc=ws.cell(row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,7,13,19,26) else _aln('center','center',wrap=True)
            for col in range(1,len(headers)+1): ws.cell(row,col).fill=_fill('FCE4D6')
            row += 2
            continue
        # filas concepto CD seleccionados
        for sm in selected[:6]:
            cd=sm.get('concept') or {}
            vals=[clave,c.get('desc'),'CONCEPTO_CD',tipo_match,match.get('score'),code_txt,reason,cov_t,cov_e,'CONCEPTO_USADO','','','','','','','','',cd.get('desc_larga') or cd.get('desc'),cd.get('unidad'),'','',cd.get('pu_mercado'),cd.get('codigo') or cd.get('clave'),'',sm.get('motivo') or 'Concepto usado para construir la matriz recomendada.']
            for col,val in enumerate(vals,1):
                cc=ws.cell(row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,7,13,19,26) else _aln('center','center',wrap=True)
                if col in (8,9) and isinstance(val,(int,float)): cc.number_format='0.0%'
                if col in (23,) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
            for col in range(1,len(headers)+1): ws.cell(row,col).fill=_fill('D9EAF7')
            row += 1
        # filas estructurales
        for rr in comparison_rows:
            p=rr.get('provider') or {}; m=rr.get('market') or {}; estado=rr.get('estado')
            if estado == 'COINCIDE':
                nota='Elemento declarado por contratista coincide con la matriz recomendada de mercado.'
                fill=None
            elif estado == 'FALTANTE_MERCADO':
                nota='Elemento recomendado por Construdata que no fue localizado en la matriz del contratista.'
                fill=C_AMARILLO
            else:
                nota='Elemento declarado por el contratista que no aparece en la matriz recomendada; revisar si es alcance adicional, requisito particular o sobreintegracion.'
                fill='FCE4D6'
            vals=[
                clave,c.get('desc'),'COMPARACION_ESTRUCTURAL',tipo_match,match.get('score'),code_txt,reason,cov_t,cov_e,estado,
                _step34_item_domain(p or m), p.get('codigo'), p.get('descripcion'), p.get('unidad'), p.get('factor') if p.get('factor') is not None else p.get('cantidad'), p.get('precio_base') if p.get('precio_base') is not None else p.get('precio'), p.get('importe'),
                m.get('codigo') or m.get('clave'), m.get('descripcion'), m.get('unidad'), m.get('cantidad'), m.get('precio'), m.get('importe'), ', '.join(m.get('source_concepts') or [str(m.get('source_concept_code') or '')]), rr.get('score'), nota
            ]
            for col,val in enumerate(vals,1):
                cc=ws.cell(row,col,val); cc.border=_brd(); cc.font=_fnt(size=8); cc.alignment=_aln('left','center',wrap=True) if col in (2,7,13,19,24,26) else _aln('center','center',wrap=True)
                if col in (8,9) and isinstance(val,(int,float)): cc.number_format='0.0%'
                if col in (16,17,22,23) and isinstance(val,(int,float)): cc.number_format='$#,##0.00'
                if col in (15,21) and isinstance(val,(int,float)): cc.number_format='0.0000'
            if fill:
                for col in range(1,len(headers)+1): ws.cell(row,col).fill=_fill(fill)
            row += 1
        row += 1
    # formato
    widths={
        'A':16,'B':55,'C':24,'D':20,'E':11,'F':28,'G':60,'H':12,'I':12,'J':22,'K':18,'L':18,'M':42,'N':10,'O':12,'P':14,'Q':14,
        'R':18,'S':44,'T':10,'U':12,'V':14,'W':14,'X':22,'Y':12,'Z':70
    }
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    try:
        ws.freeze_panes='A5'; ws.auto_filter.ref=f'A4:Z{max(row-1,4)}'
    except Exception: pass
    return ws

# Sustituye el tab anterior por la Matriz Recomendada de Mercado.
_step08_build_matrix_base_sheet = _step34_build_recommended_market_matrix_sheet

# Asegurar orden de tabs con el nuevo nombre cuando el writer final ya fue ejecutado por wrappers previos.
try:
    _STEP34_BUILD_SINGLE_PREV = _build_single_provider_pmd
    def _build_single_provider_pmd(output, meta=None, dato=None, provider_analysis=None):
        # STEP35: keep the original 4-argument signature. Step34 accidentally
        # replaced it with a 3-argument function, which broke/short-circuited
        # the final workbook post-processing in some deployments.
        output = _STEP34_BUILD_SINGLE_PREV(output, meta or {}, dato or {}, provider_analysis)
        try:
            wb = openpyxl.load_workbook(output)
            if 'Matriz base CD' in wb.sheetnames or 'Matriz Recomendada Mercado' not in wb.sheetnames:
                _step34_build_recommended_market_matrix_sheet(wb, dato or {})
            desired=['Comparativa','Detalle','Matriz Recomendada Mercado','Analisis experto IA']
            wb._sheets=[wb[n] for n in desired if n in wb.sheetnames] + [sh for sh in wb._sheets if sh.title not in desired and sh.title != 'Matriz base CD']
            wb.save(output)
        except Exception as exc:
            # Do not fail silently: create a visible diagnostic sheet in the Excel.
            try:
                wb = openpyxl.load_workbook(output)
                if 'Matriz base CD' in wb.sheetnames:
                    del wb['Matriz base CD']
                if 'Matriz Recomendada Mercado' in wb.sheetnames:
                    del wb['Matriz Recomendada Mercado']
                ws = wb.create_sheet('Matriz Recomendada Mercado')
                ws['A1'] = 'MATRIZ RECOMENDADA DE MERCADO - ERROR DE CONSTRUCCION'
                ws['A2'] = f'{type(exc).__name__}: {exc}'
                ws['A4'] = 'La comparativa principal y Detalle no se afectan. Este tab requiere revisar logs/trace del job.'
                for cell in ('A1','A2','A4'):
                    ws[cell].alignment = Alignment(wrap_text=True, vertical='top')
                ws.column_dimensions['A'].width = 120
                wb.save(output)
            except Exception:
                pass
            _q_perf_log('step34_matrix_recommended_sheet_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
        return output
except Exception:
    pass

# ============================================================
# PASO 36 - Motor de interpretacion de matrices v2
# ============================================================
# Este paso corrige la base del calculo: las matrices reales no son listas planas
# de insumos, sino arboles de calculo con headers, controles, subtotales,
# volumenes y financieros. La parte del CONTRATISTA se lee de la matriz; la parte
# de MERCADO se calcula replicando la estructura porcentual/volumetrica declarada.

try:
    _STEP36_EXTRACT_CONCEPTOS_PREV = extract_conceptos
except Exception:
    _STEP36_EXTRACT_CONCEPTOS_PREV = None

try:
    _STEP36_PARSE_COMPONENT_ROW_PREV = _parse_component_row
except Exception:
    _STEP36_PARSE_COMPONENT_ROW_PREV = None


def _step36_as_float(v):
    try:
        return _as_float(v)
    except Exception:
        try:
            if v is None or v == '':
                return None
            return float(str(v).replace(',', '').replace('$', '').replace('%', '').strip())
        except Exception:
            return None


def _step36_is_zeroish(v):
    n = _step36_as_float(v)
    return n is None or abs(float(n)) < 1e-12


def _step36_row_values(row, max_len=20):
    vals = list(row or [])
    if len(vals) < max_len:
        vals += [None] * (max_len - len(vals))
    return vals


def _step36_section_key(label):
    txt = _normalize_text(label or '')
    if 'mano' in txt or txt == 'mo':
        return 'mano_obra'
    if 'equipo' in txt or 'herramienta' in txt or 'maquinaria' in txt:
        return 'equipo'
    if 'basico' in txt or 'auxiliar' in txt:
        return 'basicos'
    if 'material' in txt:
        return 'material'
    return None


def _step36_section_label_from_key(key):
    return {
        'material': 'MATERIALES',
        'mano_obra': 'MANO DE OBRA',
        'equipo': 'EQUIPO Y HERRAMIENTA',
        'basicos': 'BASICOS',
    }.get(key or '', key or '')


def _step36_row_text(row):
    try:
        return _row_text(row)
    except Exception:
        return ' '.join(str(x) for x in (row or []) if x is not None).lower()


def _step36_read_workbook_rows(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    # Preferir hoja con mas filas utiles.
    best_ws = None
    best_count = -1
    for ws in wb.worksheets:
        count = 0
        for r in ws.iter_rows(values_only=True):
            if any(v not in (None, '') for v in r):
                count += 1
            if count > 20 and best_count > 0:
                break
        if count > best_count:
            best_count = count
            best_ws = ws
    ws = best_ws or wb[wb.sheetnames[0]]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _step36_analysis_ranges(rows):
    starts = [i for i, r in enumerate(rows or []) if _looks_like_analysis_row(r)]
    if not starts:
        return []
    starts.append(len(rows))
    return [(starts[i], starts[i+1]) for i in range(len(starts)-1)]


def _step36_row_amount(row):
    vals = _step36_row_values(row, 8)
    # En estos formatos el importe suele estar en col G (idx 6).
    for idx in (6, 7, 5, 4, 3):
        n = _step36_as_float(vals[idx] if idx < len(vals) else None)
        if n is not None:
            return n
    return None


def _step36_row_pct(row):
    vals = _step36_row_values(row, 8)
    # En estos formatos el porcentaje financiero suele estar en col F (idx 5)
    # o como proporcion 0.30 / 0.24. Ignorar el 1 de participacion CD.
    candidates = []
    for idx in (5, 7, 4, 3, 6):
        n = _step36_as_float(vals[idx] if idx < len(vals) else None)
        if n is None:
            continue
        x = float(n)
        if x > 1.5:
            x = x / 100.0
        if 0 < x < 1.5:
            candidates.append(x)
    # Elegir el primer porcentaje plausible distinto de 1.0.
    for x in candidates:
        if abs(x - 1.0) > 1e-9:
            return x
    return None


def _step36_parse_matrix_block(rows, start, end):
    controls = {
        'material': {}, 'mano_obra': {}, 'equipo': {}, 'basicos': {}
    }
    financial = {
        'cost_direct': None,
        'price_unit': None,
        'charges': [],
        'subtotals': [],
        'rows': [],
    }
    current_section = None
    last_section_with_importe = None

    for i in range(start, min(end, len(rows))):
        row = rows[i]
        vals = _step36_row_values(row, 10)
        c0 = _clean_text(vals[0])
        c1 = _clean_text(vals[1])
        txt = _normalize_text(' '.join(str(v) for v in vals[:8] if v not in (None, '')))

        # Seccion visual: materiales, mano de obra, equipo, basicos.
        sec = None
        if c0 and not any(_step36_as_float(vals[j]) is not None for j in (3,5,6)):
            sec = _step36_section_key(c0)
        if not sec and c1 and not any(_step36_as_float(vals[j]) is not None for j in (3,5,6)):
            sec = _step36_section_key(c1)
        if sec:
            current_section = sec
            continue

        # Controles de bloque: Importe, Volumen y Subtotal.
        if 'importe' in txt and 'precio unitario' not in txt and not 'total' in txt:
            amount = _step36_as_float(vals[6])
            if current_section and amount is not None:
                controls.setdefault(current_section, {})['importe_base'] = amount
                controls[current_section]['importe_row'] = i + 1
                last_section_with_importe = current_section
            continue

        if 'volumen' in txt:
            volumen = _step36_as_float(vals[5])
            amount = _step36_as_float(vals[6])
            target = current_section or last_section_with_importe
            if target:
                if volumen is not None:
                    controls.setdefault(target, {})['volumen'] = volumen
                if amount is not None:
                    controls.setdefault(target, {})['subtotal_volumen'] = amount
                controls.setdefault(target, {})['volumen_row'] = i + 1
            continue

        if c0.upper().startswith('SUBTOTAL') or txt.startswith('subtotal'):
            sec_label = c1 or c0
            sec_key = _step36_section_key(sec_label) or current_section
            amount = _step36_as_float(vals[6])
            pct = _step36_as_float(vals[7])
            if sec_key and amount is not None:
                controls.setdefault(sec_key, {})['subtotal_declared'] = amount
                controls[sec_key]['subtotal_pct_declared'] = pct
                controls[sec_key]['subtotal_row'] = i + 1
                current_section = sec_key
            continue

        # Financieros.
        amount = _step36_as_float(vals[6])
        pct = _step36_row_pct(vals)
        if 'costo directo' in txt or re.search(r'\bcd\b', txt):
            if amount is not None:
                financial['cost_direct'] = amount
                financial['rows'].append({'kind':'cost_direct','label':c1 or c0,'amount':amount,'pct':None,'row':i+1})
            continue
        if 'indirect' in txt or re.search(r'\bci\b', txt):
            if amount is not None or pct is not None:
                financial['charges'].append({'kind':'indirecto','label':c1 or c0 or 'INDIRECTO','amount':amount,'pct':pct,'row':i+1})
                financial['rows'].append({'kind':'indirecto','label':c1 or c0 or 'INDIRECTO','amount':amount,'pct':pct,'row':i+1})
            continue
        if 'financiamiento' in txt or re.search(r'\bcf\b', txt):
            if amount is not None or pct is not None:
                financial['charges'].append({'kind':'financiamiento','label':c1 or c0 or 'FINANCIAMIENTO','amount':amount,'pct':pct,'row':i+1})
                financial['rows'].append({'kind':'financiamiento','label':c1 or c0 or 'FINANCIAMIENTO','amount':amount,'pct':pct,'row':i+1})
            continue
        if 'utilidad' in txt or re.search(r'\bcu\b', txt):
            if amount is not None or pct is not None:
                financial['charges'].append({'kind':'utilidad','label':c1 or c0 or 'UTILIDAD','amount':amount,'pct':pct,'row':i+1})
                financial['rows'].append({'kind':'utilidad','label':c1 or c0 or 'UTILIDAD','amount':amount,'pct':pct,'row':i+1})
            continue
        if 'subtotal1' in txt or 'subtotal2' in txt or (txt.startswith('subtotal') and amount is not None):
            if amount is not None:
                financial['subtotals'].append({'kind':'subtotal','label':c1 or c0 or 'SUBTOTAL','amount':amount,'row':i+1})
                financial['rows'].append({'kind':'subtotal','label':c1 or c0 or 'SUBTOTAL','amount':amount,'pct':None,'row':i+1})
            continue
        if 'precio unitario' in txt:
            if amount is not None:
                financial['price_unit'] = amount
                financial['rows'].append({'kind':'price_unit','label':c1 or c0 or 'PRECIO UNITARIO','amount':amount,'pct':None,'row':i+1})
            continue

    # Derivar multiplicadores por seccion cuando exista Importe + Volumen/Subtotal.
    for sec, ctl in controls.items():
        imp = ctl.get('importe_base')
        subtotal = ctl.get('subtotal_declared') or ctl.get('subtotal_volumen')
        vol = ctl.get('volumen')
        multiplier = 1.0
        if isinstance(vol, (int, float)) and vol not in (0, 1):
            multiplier = float(vol)
        elif isinstance(imp, (int, float)) and isinstance(subtotal, (int, float)) and imp:
            multiplier = float(subtotal) / float(imp)
        ctl['multiplier'] = multiplier

    return {'controls': controls, 'financial': financial}


def _step36_apply_matrix_v2_to_concepts(filepath, conceptos):
    if not isinstance(conceptos, dict) or not conceptos:
        return conceptos
    try:
        rows = _step36_read_workbook_rows(filepath)
        ranges = _step36_analysis_ranges(rows)
    except Exception as exc:
        for c in conceptos.values():
            if isinstance(c, dict):
                c['matrix_v2_error'] = f'{type(exc).__name__}: {exc}'
        return conceptos
    keys = [k for k in conceptos.keys() if not str(k).startswith('__')]
    # Alinear por orden de aparicion; si las claves coinciden, reforzar por clave.
    for idx, key in enumerate(keys):
        if idx >= len(ranges):
            continue
        start, end = ranges[idx]
        parsed = _step36_parse_matrix_block(rows, start, end)
        c = conceptos.get(key)
        if not isinstance(c, dict):
            continue
        c['matrix_v2'] = parsed
        fin = parsed.get('financial') or {}
        ctl = parsed.get('controls') or {}
        # Forzar lectura del contratista desde la matriz, no recalcular.
        if isinstance(fin.get('cost_direct'), (int, float)):
            c['costo_directo'] = float(fin['cost_direct'])
        if isinstance(fin.get('price_unit'), (int, float)):
            c['pu'] = float(fin['price_unit'])
        for ch in fin.get('charges') or []:
            kind = ch.get('kind')
            if kind == 'indirecto':
                if isinstance(ch.get('amount'), (int, float)):
                    c['indirecto'] = float(ch['amount'])
                if isinstance(ch.get('pct'), (int, float)):
                    c['indirecto_pct'] = float(ch['pct'])
            elif kind == 'financiamiento':
                if isinstance(ch.get('amount'), (int, float)):
                    c['financiamiento'] = float(ch['amount'])
                if isinstance(ch.get('pct'), (int, float)):
                    c['financiamiento_pct'] = float(ch['pct'])
            elif kind == 'utilidad':
                if isinstance(ch.get('amount'), (int, float)):
                    c['utilidad'] = float(ch['amount'])
                if isinstance(ch.get('pct'), (int, float)):
                    c['utilidad_pct'] = float(ch['pct'])
        # Subtotales declarados por seccion.
        secmap = {'material':'materiales', 'mano_obra':'mano_obra', 'equipo':'equipo', 'basicos':'basicos'}
        for sec, attr in secmap.items():
            val = (ctl.get(sec) or {}).get('subtotal_declared')
            if isinstance(val, (int, float)):
                c[attr] = float(val)
        c['matrix_v2_summary'] = {
            'controls_detected': {k:v for k,v in ctl.items() if any(x in v for x in ('subtotal_declared','volumen','importe_base'))},
            'financial_rows': fin.get('rows') or [],
        }
    return conceptos


def extract_conceptos(filepath):
    conceptos = _STEP36_EXTRACT_CONCEPTOS_PREV(filepath) if _STEP36_EXTRACT_CONCEPTOS_PREV else {}
    return _step36_apply_matrix_v2_to_concepts(filepath, conceptos)


def _parse_component_row(row):
    vals = _step36_row_values(row, 8)
    c0 = _clean_text(vals[0])
    c1 = _clean_text(vals[1])
    txt = _normalize_text(f'{c0} {c1}')
    price = _step36_as_float(vals[3])
    op = vals[4]
    qty = _step36_as_float(vals[5])
    amount = _step36_as_float(vals[6])
    # Headers visuales de cuadrilla/agrupador: no tienen costo efectivo.
    if (price is None or abs(price) < 1e-12) and (qty is None or abs(qty) < 1e-12) and (amount is None or abs(amount) < 1e-12):
        return None
    if 'importe' in txt or 'volumen' in txt or txt.startswith('subtotal'):
        return None
    # Paso38: renglones de control tipo "Rendimiento: PZA/JOR" no son insumos.
    # En SIMSA estos renglones ya vienen reflejados en subtotales/controles y
    # si se tratan como insumo se duplica el importe de mercado.
    if 'rendimiento' in txt and (not c0 or str(c0).strip().lower().startswith('rendimiento')):
        return None
    return _STEP36_PARSE_COMPONENT_ROW_PREV(row) if _STEP36_PARSE_COMPONENT_ROW_PREV else None


try:
    _STEP36_APPLY_MARKET_PREV = _step08_apply_market_catalog_pricing
except Exception:
    _STEP36_APPLY_MARKET_PREV = None


def _step36_item_section_key(item):
    sec = item.get('section') or item.get('_section') or item.get('domain') or item.get('_domain')
    return _step36_section_key(sec) or ('mano_obra' if 'mano' in _normalize_text(sec) else 'equipo' if 'equipo' in _normalize_text(sec) else 'material')


def _step36_item_price_market(item):
    # precio_mercado puede haberse cambiado a base porcentual para display; para no porcentuales sirve.
    if item.get('is_percent_item'):
        return item.get('precio_mercado_catalogo') if isinstance(item.get('precio_mercado_catalogo'), (int, float)) else item.get('precio_mercado')
    return item.get('precio_mercado') if isinstance(item.get('precio_mercado'), (int, float)) else item.get('precio_mercado_catalogo')


def _step36_qty(item):
    return item.get('cantidad') if isinstance(item.get('cantidad'), (int, float)) else item.get('factor')


def _step36_op(item):
    return item.get('op') or item.get('operacion') or '*'


def _step36_recalculate_structured_market(c):
    items = c.get('granular_market_items') or []
    if not items:
        return c
    controls = ((c.get('matrix_v2') or {}).get('controls') or {})
    # Primer pase: importes de no porcentuales sin volumen de seccion.
    raw_nonpercent = {'material':0.0, 'mano_obra':0.0, 'equipo':0.0, 'basicos':0.0}
    raw_amounts = {}
    for idx, item in enumerate(items):
        sec = _step36_item_section_key(item)
        is_pct = bool(item.get('is_percent_item') or _step08_is_percent_item(item))
        if is_pct:
            continue
        price = _step36_item_price_market(item)
        qty = _step36_qty(item)
        amt = _step11_operation_amount(price, qty, _step36_op(item)) if '_step11_operation_amount' in globals() else None
        if not isinstance(amt, (int, float)):
            amt = item.get('importe_mercado') if isinstance(item.get('importe_mercado'), (int, float)) else item.get('importe_proveedor')
        if isinstance(amt, (int, float)):
            raw_amounts[idx] = float(amt)
            raw_nonpercent[sec] = raw_nonpercent.get(sec, 0.0) + float(amt)

    # Calcular porcentuales de MO sobre MO no porcentual cruda, luego aplicar volumen de MO.
    mo_base_raw = raw_nonpercent.get('mano_obra', 0.0)
    raw_section_total = dict(raw_nonpercent)
    for idx, item in enumerate(items):
        if not (item.get('is_percent_item') or _step08_is_percent_item(item)):
            continue
        sec = _step36_item_section_key(item)
        qty = _step36_qty(item)
        if not isinstance(qty, (int, float)):
            continue
        if sec == 'mano_obra':
            base = mo_base_raw
        elif sec == 'equipo':
            # Cargos porcentuales de equipo/herramienta/EPP/SHE/consumibles normalmente se calculan sobre subtotal MO final.
            mo_mult = ((controls.get('mano_obra') or {}).get('multiplier') or 1.0)
            base = raw_section_total.get('mano_obra', 0.0) * float(mo_mult)
        else:
            base = raw_nonpercent.get(sec, 0.0)
        amt = float(base) * float(qty)
        raw_amounts[idx] = amt
        raw_section_total[sec] = raw_section_total.get(sec, 0.0) + amt
        item['is_percent_item'] = True
        item['base_mercado'] = base
        item['percent_base_kind'] = 'mano_obra' if sec in {'mano_obra','equipo'} else sec
        item['match_reason'] = (item.get('match_reason') or '') + (' | ' if item.get('match_reason') else '') + 'Paso36: porcentaje aplicado sobre la base estructural correcta de la matriz.'

    # Aplicar volumen/multiplicador por seccion.
    final_totals = {'material':0.0, 'mano_obra':0.0, 'equipo':0.0, 'basicos':0.0}
    for idx, item in enumerate(items):
        sec = _step36_item_section_key(item)
        raw_amt = raw_amounts.get(idx)
        if not isinstance(raw_amt, (int, float)):
            continue
        mult = ((controls.get(sec) or {}).get('multiplier') or 1.0)
        final_amt = float(raw_amt) * float(mult)
        item['importe_mercado_base_sin_volumen'] = raw_amt
        item['section_multiplier'] = mult
        item['importe_mercado'] = final_amt
        # Para porcentaje, mostrar costo mercado como base usada; precio catalogo queda aparte.
        if item.get('is_percent_item'):
            item['precio_mercado'] = item.get('base_mercado')
        final_totals[sec] = final_totals.get(sec, 0.0) + final_amt

    market_direct = sum(final_totals.values())
    c['granular_market_direct'] = round(market_direct, 6) if market_direct else None
    c['granular_market_subtotals'] = {k: round(v, 6) for k,v in final_totals.items() if isinstance(v, (int,float))}
    # Contratista directo: leer de matriz si existe. No recalcular.
    fin = ((c.get('matrix_v2') or {}).get('financial') or {})
    if isinstance(fin.get('cost_direct'), (int, float)):
        c['granular_provider_direct'] = float(fin['cost_direct'])
    return c


def _step08_apply_market_catalog_pricing(conceptos):
    conceptos = _STEP36_APPLY_MARKET_PREV(conceptos) if _STEP36_APPLY_MARKET_PREV else conceptos
    for c in (conceptos or {}).values():
        if isinstance(c, dict):
            _step36_recalculate_structured_market(c)
    return conceptos


def _step36_financial_rows(c):
    fin = ((c.get('matrix_v2') or {}).get('financial') or {})
    provider_direct = fin.get('cost_direct') if isinstance(fin.get('cost_direct'), (int, float)) else _step12_provider_direct(c)
    market_direct = c.get('granular_market_direct') if isinstance(c.get('granular_market_direct'), (int, float)) else _step12_market_direct(c)
    rows = []
    rows.append(('COSTO DIRECTO', provider_direct, market_direct, None, 'Contratista: leído de la matriz. Mercado: suma de insumos con precios mercado respetando estructura, volumenes y porcentajes.'))
    provider_running = float(provider_direct or 0.0)
    market_running = float(market_direct or 0.0)
    for ch in (fin.get('charges') or []):
        label = ch.get('label') or ch.get('kind') or 'CARGO'
        kind = ch.get('kind') or ''
        provider_amt = ch.get('amount') if isinstance(ch.get('amount'), (int, float)) else 0.0
        pct = ch.get('pct') if isinstance(ch.get('pct'), (int, float)) else None
        if pct is None:
            base = provider_running if provider_running else provider_direct
            pct = _step32_pct_from_amount(provider_amt, base) if '_step32_pct_from_amount' in globals() else 0.0
        pretty = 'COSTO INDIRECTO' if kind == 'indirecto' else 'FINANCIAMIENTO' if kind == 'financiamiento' else 'UTILIDAD / CARGOS ADICIONALES' if kind == 'utilidad' else str(label).upper()
        # Paso38: el mercado SIEMPRE usa 25% de indirecto, aunque el contratista
        # declare 30%, menos o más. Financiamiento/utilidad siguen usando los %
        # declarados por el contratista sobre el subtotal mercado previo.
        if kind == 'indirecto':
            market_pct = MARKET_INDIRECT_PCT if 'MARKET_INDIRECT_PCT' in globals() else 0.25
            market_amt = market_running * float(market_pct)
            note = f'Mercado usa indirecto fijo de {market_pct*100:.2f}% sobre costo directo mercado. El % del contratista se conserva solo como referencia ({(pct or 0)*100:.2f}%).'
            display_pct = market_pct
        else:
            market_amt = market_running * float(pct or 0.0)
            note = f'Mercado usa el porcentaje declarado por contratista ({(pct or 0)*100:.2f}%) sobre el subtotal mercado previo.'
            display_pct = pct
        rows.append((pretty, provider_amt, market_amt, display_pct, note))
        provider_running += float(provider_amt or 0.0)
        market_running += float(market_amt or 0.0)
    pu_provider = fin.get('price_unit') if isinstance(fin.get('price_unit'), (int, float)) else c.get('pu')
    # Solo agregar remanente si es relevante y no hubo precio unitario coincidente por redondeo.
    gap = None
    if isinstance(pu_provider, (int, float)):
        gap = float(pu_provider) - provider_running
    if isinstance(gap, (int, float)) and gap > max(1.0, abs(float(pu_provider or 0))*0.005):
        pct = gap / provider_running if provider_running else 0.0
        market_gap = market_running * pct if pct else 0.0
        rows.append(('OTROS CARGOS DECLARADOS NO CLASIFICADOS', gap, market_gap, pct, 'Remanente real entre PU y cargos detectados; se replica como porcentaje solo si supera tolerancia.'))
        provider_running += gap
        market_running += market_gap
    rows.append(('TOTAL COSTO UNITARIO', pu_provider, market_running, None, 'PU contratista leído de matriz; PU mercado recalculado con la misma estructura financiera.'))
    return rows


def _step13_financial_rows(c):
    return _step36_financial_rows(c)


def _step13_market_pu(c):
    rows = _step36_financial_rows(c)
    total = rows[-1][2] if rows else None
    return total if isinstance(total, (int, float)) and total > 0 else 0.0


def _market_pu_with_indirect(concepto):
    return _step13_market_pu(concepto)


def _step12_provider_direct(c):
    fin = ((c.get('matrix_v2') or {}).get('financial') or {})
    if isinstance(fin.get('cost_direct'), (int, float)):
        return float(fin['cost_direct'])
    val = c.get('costo_directo')
    if isinstance(val, (int, float)):
        return float(val)
    val = c.get('granular_provider_direct')
    return float(val) if isinstance(val, (int, float)) else 0.0


def _step12_market_direct(c):
    val = c.get('granular_market_direct')
    return float(val) if isinstance(val, (int, float)) else 0.0


def _step12_write_financial_summary_rows(ws_det, start_row, c):
    rows = _step36_financial_rows(c)
    row = start_row
    for label, prov, market, pct, note in rows:
        vals = ['', label, '', prov, '', '', prov, market, None, market, '', 'RESUMEN FINANCIERO V2', '', pct, '', note]
        if isinstance(prov, (int, float)) and isinstance(market, (int, float)) and market:
            vals[8] = (float(prov) - float(market)) / float(market)
        for col, val in enumerate(vals, 1):
            cell = ws_det.cell(row, col, val)
            cell.border = _brd(); cell.font = _fnt(bold=True if label == 'TOTAL COSTO UNITARIO' else False, size=8)
            cell.alignment = _aln('left','center', wrap=True) if col in (2,16) else _aln('center','center')
            if col in (4,7,8,10) and isinstance(val, (int, float)): cell.number_format = '$#,##0.00'
            if col in (9,14) and isinstance(val, (int, float)): cell.number_format = '0.00%'
            if col in (8,9): cell.fill = _fill(C_AMARILLO_PMD)
            if label == 'TOTAL COSTO UNITARIO': cell.fill = _fill('D9EAF7') if col not in (8,9) else _fill(C_AMARILLO_PMD)
        row += 1
    # Agregar controles detectados para trazabilidad de volumen/subtotales.
    controls = ((c.get('matrix_v2') or {}).get('controls') or {})
    if any((v or {}).get('multiplier') not in (None, 1, 1.0) for v in controls.values()):
        ws_det.cell(row, 2, 'CONTROLES DE MATRIZ DETECTADOS').font = _fnt(bold=True, size=8)
        ws_det.cell(row, 2).fill = _fill(C_GRIS_SEC)
        row += 1
        for sec, ctl in controls.items():
            if not isinstance(ctl, dict):
                continue
            if not any(k in ctl for k in ('importe_base','volumen','subtotal_declared','multiplier')):
                continue
            note = f"{_step36_section_label_from_key(sec)} | Importe base={ctl.get('importe_base')} | Volumen={ctl.get('volumen')} | Subtotal declarado={ctl.get('subtotal_declared')} | Multiplicador aplicado={ctl.get('multiplier')}"
            ws_det.cell(row, 2, note).alignment = _aln('left','center', wrap=True)
            ws_det.cell(row, 2).font = _fnt(size=8)
            row += 1
    return row

# Registrar version funcional visible en runtime/debug si existe.
QUANTIA_MATRIX_ENGINE_VERSION = 'step38_market_indirect_25_ignore_rendimiento_controls'

# Paso 36 hotfix: orden correcto de deteccion financiera.
# "PRECIO UNITARIO (CD+CI)" no debe ser detectado como Costo Directo por contener CD.
# "SUBTOTAL1/SUBTOTAL2" no son subtotales de seccion, son controles financieros.
def _step36_parse_matrix_block(rows, start, end):
    controls = {'material': {}, 'mano_obra': {}, 'equipo': {}, 'basicos': {}}
    financial = {'cost_direct': None, 'price_unit': None, 'charges': [], 'subtotals': [], 'rows': []}
    current_section = None
    last_section_with_importe = None
    for i in range(start, min(end, len(rows))):
        row = rows[i]
        vals = _step36_row_values(row, 10)
        c0 = _clean_text(vals[0])
        c1 = _clean_text(vals[1])
        txt = _normalize_text(' '.join(str(v) for v in vals[:8] if v not in (None, '')))
        amount = _step36_as_float(vals[6])
        pct = _step36_row_pct(vals)

        # 1) Financieros explicitos primero.
        if 'precio unitario' in txt:
            if amount is not None:
                financial['price_unit'] = amount
                financial['rows'].append({'kind':'price_unit','label':c1 or c0 or 'PRECIO UNITARIO','amount':amount,'pct':None,'row':i+1})
            continue
        if 'subtotal1' in txt or 'subtotal2' in txt:
            if amount is not None:
                financial['subtotals'].append({'kind':'subtotal','label':c1 or c0 or 'SUBTOTAL','amount':amount,'row':i+1})
                financial['rows'].append({'kind':'subtotal','label':c1 or c0 or 'SUBTOTAL','amount':amount,'pct':None,'row':i+1})
            continue
        if 'costo directo' in txt or re.search(r'\bcd\b', txt):
            if amount is not None:
                financial['cost_direct'] = amount
                financial['rows'].append({'kind':'cost_direct','label':c1 or c0 or 'COSTO DIRECTO','amount':amount,'pct':None,'row':i+1})
            continue
        if 'indirect' in txt or re.search(r'\bci\b', txt):
            if amount is not None or pct is not None:
                financial['charges'].append({'kind':'indirecto','label':c1 or c0 or 'INDIRECTO','amount':amount,'pct':pct,'row':i+1})
                financial['rows'].append({'kind':'indirecto','label':c1 or c0 or 'INDIRECTO','amount':amount,'pct':pct,'row':i+1})
            continue
        if 'financiamiento' in txt or re.search(r'\bcf\b', txt):
            if amount is not None or pct is not None:
                financial['charges'].append({'kind':'financiamiento','label':c1 or c0 or 'FINANCIAMIENTO','amount':amount,'pct':pct,'row':i+1})
                financial['rows'].append({'kind':'financiamiento','label':c1 or c0 or 'FINANCIAMIENTO','amount':amount,'pct':pct,'row':i+1})
            continue
        if 'utilidad' in txt or re.search(r'\bcu\b', txt):
            if amount is not None or pct is not None:
                financial['charges'].append({'kind':'utilidad','label':c1 or c0 or 'UTILIDAD','amount':amount,'pct':pct,'row':i+1})
                financial['rows'].append({'kind':'utilidad','label':c1 or c0 or 'UTILIDAD','amount':amount,'pct':pct,'row':i+1})
            continue

        # 2) Subtotales reales de seccion: SUBTOTAL: MATERIALES/MANO DE OBRA/EQUIPO/BASICOS.
        if c0.upper().startswith('SUBTOTAL') or txt.startswith('subtotal'):
            sec_label = c1 or c0
            sec_key = _step36_section_key(sec_label)
            if sec_key and amount is not None:
                controls.setdefault(sec_key, {})['subtotal_declared'] = amount
                controls[sec_key]['subtotal_pct_declared'] = _step36_as_float(vals[7])
                controls[sec_key]['subtotal_row'] = i + 1
                current_section = sec_key
            continue

        # 3) Encabezados de seccion visual.
        sec = None
        has_numeric_matrix = any(_step36_as_float(vals[j]) is not None for j in (3,5,6))
        if c0 and not has_numeric_matrix:
            sec = _step36_section_key(c0)
        if not sec and c1 and not has_numeric_matrix:
            sec = _step36_section_key(c1)
        if sec:
            current_section = sec
            continue

        # 4) Controles internos: Importe / Volumen.
        if 'importe' in txt and 'total' not in txt:
            if current_section and amount is not None:
                controls.setdefault(current_section, {})['importe_base'] = amount
                controls[current_section]['importe_row'] = i + 1
                last_section_with_importe = current_section
            continue
        if 'volumen' in txt:
            volumen = _step36_as_float(vals[5])
            target = current_section or last_section_with_importe
            if target:
                if volumen is not None:
                    controls.setdefault(target, {})['volumen'] = volumen
                if amount is not None:
                    controls.setdefault(target, {})['subtotal_volumen'] = amount
                controls.setdefault(target, {})['volumen_row'] = i + 1
            continue

    for sec, ctl in controls.items():
        imp = ctl.get('importe_base')
        subtotal = ctl.get('subtotal_declared') or ctl.get('subtotal_volumen')
        vol = ctl.get('volumen')
        multiplier = 1.0
        if isinstance(vol, (int, float)) and abs(float(vol)) > 1e-12 and abs(float(vol)-1.0) > 1e-12:
            multiplier = float(vol)
        elif isinstance(imp, (int, float)) and isinstance(subtotal, (int, float)) and abs(float(imp)) > 1e-12:
            multiplier = float(subtotal) / float(imp)
        ctl['multiplier'] = multiplier
    return {'controls': controls, 'financial': financial}

# -----------------------------------------------------------------------------
# STEP37 - Rediseño profesional del tab Analisis experto IA
# -----------------------------------------------------------------------------
# Objetivo:
# - El tab deja de ser un volcado de texto por servicio.
# - Se convierte en un informe ejecutivo de costos para analistas PU.
# - Claude participa solo como redactor experto sobre métricas ya calculadas.
# - Si Claude esta apagado o falla, se genera un resumen deterministico util.

def _step37_safe_num(v, default=0.0):
    try:
        if v is None or v == '':
            return default
        return float(v)
    except Exception:
        return default


def _step37_money(v):
    try:
        return float(v or 0.0)
    except Exception:
        return 0.0


def _step37_concepts(dato):
    conceptos = (dato or {}).get('conceptos') or {}
    return [(k, v) for k, v in conceptos.items() if isinstance(v, dict) and not str(k).startswith('__')]


def _step37_concept_provider_total(c):
    for key in ('importe_resumen_pu', 'importe_contratista', 'importe', 'total'):
        if isinstance(c.get(key), (int, float)):
            return float(c.get(key) or 0.0)
    qty = _step37_safe_num(c.get('cantidad'), 1.0) or 1.0
    pu = _step37_safe_num(c.get('pu'), 0.0)
    return qty * pu


def _step37_concept_market_pu(c):
    # Priorizar motor financiero actual.
    try:
        if '_step13_market_pu' in globals():
            v = _step13_market_pu(c)
            if isinstance(v, (int, float)) and v > 0:
                return float(v)
    except Exception:
        pass
    for key in ('pu_mercado', 'market_pu'):
        if isinstance(c.get(key), (int, float)):
            return float(c.get(key) or 0.0)
    direct = c.get('granular_market_direct')
    if isinstance(direct, (int, float)):
        return float(direct or 0.0)
    return 0.0


def _step37_concept_market_total(c):
    for key in ('importe_mercado', 'market_total', 'total_mercado'):
        if isinstance(c.get(key), (int, float)):
            return float(c.get(key) or 0.0)
    qty = _step37_safe_num(c.get('cantidad'), 1.0) or 1.0
    return qty * _step37_concept_market_pu(c)


def _step37_item_domain(item):
    txt = _normalize_text(' '.join(str(item.get(k) or '') for k in ('section','domain','tipo','descripcion','codigo')))
    if 'mano' in txt or 'mo' == txt or 'operario' in txt or 'ayudante' in txt or 'supervisor' in txt:
        return 'Mano de obra'
    if 'equipo' in txt or 'maquinaria' in txt or 'herramienta' in txt or 'camion' in txt or 'grua' in txt:
        return 'Equipo / herramienta'
    if 'basico' in txt:
        return 'Básicos'
    return 'Materiales'


def _step37_collect_executive_metrics(dato, provider_analysis=None):
    rows = []
    issues = []
    no_match_unit = 0
    no_match = 0
    total_items = 0
    domain_provider = {}
    domain_market = {}
    concepts = _step37_concepts(dato)
    for clave, c in concepts:
        prov_total = _step37_concept_provider_total(c)
        market_total = _step37_concept_market_total(c)
        diff = prov_total - market_total
        diff_pct = (diff / market_total) if market_total else None
        qty = _step37_safe_num(c.get('cantidad'), 1.0) or 1.0
        rows.append({
            'servicio': clave,
            'descripcion': c.get('desc') or c.get('descripcion') or '',
            'cantidad': qty,
            'pu_contratista': _step37_safe_num(c.get('pu'), 0.0),
            'pu_mercado': _step37_concept_market_pu(c),
            'importe_contratista': prov_total,
            'importe_mercado': market_total,
            'diferencia': diff,
            'diferencia_pct': diff_pct,
        })
        for item in c.get('granular_market_items') or []:
            if not isinstance(item, dict):
                continue
            total_items += 1
            dom = _step37_item_domain(item)
            prov_imp = _step37_money(item.get('importe_proveedor') or item.get('importe_contratista') or item.get('importe'))
            market_imp = _step37_money(item.get('importe_mercado'))
            domain_provider[dom] = domain_provider.get(dom, 0.0) + prov_imp
            domain_market[dom] = domain_market.get(dom, 0.0) + market_imp
            status = str(item.get('match_status') or item.get('estado_match') or '').upper()
            if 'UNIDAD' in status:
                no_match_unit += 1
            if 'SIN_MATCH' in status or 'NO_MATCH' in status:
                no_match += 1
            price_delta = item.get('delta_precio_pct')
            if not isinstance(price_delta, (int, float)):
                pc = _step37_safe_num(item.get('precio_proveedor') or item.get('precio_base'), None)
                pm = _step37_safe_num(item.get('precio_mercado'), None)
                if pc is not None and pm:
                    price_delta = (pc - pm) / pm
            if 'SIN_MATCH' in status or 'UNIDAD' in status or (isinstance(price_delta, (int, float)) and abs(price_delta) >= 0.25):
                issues.append({
                    'servicio': clave,
                    'tipo': dom,
                    'elemento': item.get('descripcion') or item.get('insumo') or item.get('codigo'),
                    'estado': status or 'REVISAR',
                    'impacto': abs(prov_imp - market_imp) if market_imp else abs(prov_imp),
                    'nota': item.get('match_reason') or item.get('observacion') or '',
                })
    total_provider = sum(r['importe_contratista'] for r in rows)
    total_market = sum(r['importe_mercado'] for r in rows)
    diff_total = total_provider - total_market
    diff_pct_total = diff_total / total_market if total_market else None
    rows_sorted_impact = sorted(rows, key=lambda r: abs(r['diferencia']), reverse=True)
    rows_sorted_value = sorted(rows, key=lambda r: r['importe_contratista'], reverse=True)
    issues = sorted(issues, key=lambda r: r['impacto'], reverse=True)
    # Pareto del diferencial absoluto.
    total_abs_diff = sum(abs(r['diferencia']) for r in rows) or 0.0
    pareto = []
    acc = 0.0
    for r in rows_sorted_impact:
        acc += abs(r['diferencia'])
        rr = dict(r)
        rr['pareto_acumulado_diff'] = acc / total_abs_diff if total_abs_diff else 0.0
        pareto.append(rr)
    # Matriz recomendada si existe en datos.
    recommended_summary = []
    for clave, c in concepts:
        cov = c.get('recommended_market_matrix_summary') or c.get('matrix_recommended_summary') or {}
        if cov:
            recommended_summary.append({'servicio': clave, **cov})
    return {
        'total_servicios': len(rows),
        'total_contratista': total_provider,
        'total_mercado': total_market,
        'diferencia_total': diff_total,
        'diferencia_pct_total': diff_pct_total,
        'top_diferenciales': rows_sorted_impact[:10],
        'top_importe': rows_sorted_value[:10],
        'pareto_diferencial': pareto[:12],
        'issues': issues[:15],
        'domain_provider': domain_provider,
        'domain_market': domain_market,
        'no_match': no_match,
        'no_match_unit': no_match_unit,
        'total_items': total_items,
        'recommended_summary': recommended_summary[:12],
    }


def _step37_ai_dictamen(metrics):
    """Claude redacta solo conclusiones ejecutivas. No calcula ni decide precios."""
    try:
        import quantia_ai_review as qair
        policy = qair.ai_runtime_policy()
        if not policy.get('expert_review_enabled') or not policy.get('anthropic_key_set'):
            return None, policy, 'Claude desactivado por política o sin API key'
        payload = {
            'rol': 'Eres un analista senior de precios unitarios. Redacta para un PM/gerente de costos y compras. Sé concreto, ejecutivo y accionable. No repitas descripciones largas; referencia servicios por clave y explica impacto económico/técnico.',
            'tarea': 'Genera un resumen ejecutivo profesional e impactante a partir de métricas calculadas. No inventes datos. No recalcules. No menciones que eres IA. Prioriza riesgos, dinero, acciones de negociación y revisión técnica PMD.',
            'formato_json': {
                'conclusion': 'parrafo breve',
                'riesgos_prioritarios': [{'servicio':'clave', 'riesgo':'texto corto', 'impacto':'texto corto'}],
                'recomendaciones': ['accion concreta'],
                'lectura_ejecutiva': ['bullet corto']
            },
            'datos': {
                'totales': {k: metrics.get(k) for k in ('total_servicios','total_contratista','total_mercado','diferencia_total','diferencia_pct_total','no_match','no_match_unit','total_items')},
                'top_diferenciales': metrics.get('top_diferenciales')[:8],
                'issues': metrics.get('issues')[:10],
                'domain_provider': metrics.get('domain_provider'),
                'domain_market': metrics.get('domain_market'),
            }
        }
        raw = qair._anthropic_message(payload, max_tokens=int(os.getenv('QUANTIA_AI_EXPERT_MAX_TOKENS','1200')), timeout=int(os.getenv('QUANTIA_AI_EXPERT_TIMEOUT','35')))
        parsed = None
        if raw:
            start, end = raw.find('{'), raw.rfind('}')
            if start >= 0 and end > start:
                try:
                    parsed = json.loads(raw[start:end+1])
                except Exception:
                    parsed = None
        return {'raw': raw, 'parsed': parsed}, policy, 'Claude ejecutivo activo'
    except Exception as exc:
        return None, {}, f'Claude ejecutivo error: {type(exc).__name__}: {str(exc)[:160]}'


def _step37_write_section_header(ws, row, title, cols=8, color=C_AZUL_OSC):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    c = ws.cell(row,1,title)
    c.fill = _fill(color); c.font = _fnt(bold=True, color=C_BLANCO, size=10)
    c.alignment = _aln('left','center',wrap=True)
    return row + 1


def _step37_write_table(ws, row, headers, data, widths=None, money_cols=None, pct_cols=None, color_header=C_AZUL_OSC):
    for idx, h in enumerate(headers, 1):
        cell = ws.cell(row, idx, h)
        cell.fill = _fill(color_header); cell.font = _fnt(bold=True, color=C_BLANCO, size=8)
        cell.border = _brd(); cell.alignment = _aln('center','center',wrap=True)
    row += 1
    for vals in data:
        for idx, val in enumerate(vals, 1):
            cell = ws.cell(row, idx, val)
            cell.border = _brd(); cell.font = _fnt(size=8)
            cell.alignment = _aln('left','center',wrap=True) if idx in (1,2,3,4,5) else _aln('center','center',wrap=True)
            if money_cols and idx in money_cols and isinstance(val, (int,float)):
                cell.number_format = '$#,##0.00'
            if pct_cols and idx in pct_cols and isinstance(val, (int,float)):
                cell.number_format = '0.00%'
        row += 1
    if widths:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    return row


def _step37_add_executive_ai_sheet(output, dato, provider_analysis=None):
    try:
        wb = openpyxl.load_workbook(output)
        if 'Analisis experto IA' in wb.sheetnames:
            del wb['Analisis experto IA']
        ws = wb.create_sheet('Analisis experto IA')
        ws.sheet_view.showGridLines = False
        metrics = _step37_collect_executive_metrics(dato or {}, provider_analysis)
        ai_result, policy, ai_status = _step37_ai_dictamen(metrics)
        parsed = (ai_result or {}).get('parsed') or {}

        # Titulo
        ws.merge_cells('A1:H1')
        ws['A1'] = 'RESUMEN EJECUTIVO DE COSTOS Y PRECIOS UNITARIOS'
        ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=14)
        ws['A1'].alignment = _aln('center','center')
        ws.merge_cells('A2:H2')
        ws['A2'] = 'Lectura profesional para revisión PMD: impacto económico, riesgos, cobertura de mercado y acciones recomendadas.'
        ws['A2'].font = _fnt(italic=True, color=C_GRIS_TXT, size=9); ws['A2'].alignment=_aln('center','center')

        # KPIs
        total_provider = metrics['total_contratista']; total_market = metrics['total_mercado']; diff = metrics['diferencia_total']; diff_pct = metrics['diferencia_pct_total']
        kpis = [
            ('Servicios', metrics['total_servicios']),
            ('Contratista', total_provider),
            ('Mercado', total_market),
            ('Diferencia', diff),
            ('Diferencia %', diff_pct if diff_pct is not None else ''),
            ('Items sin match', metrics['no_match']),
            ('Unidad incompatible', metrics['no_match_unit']),
            ('Estado IA', ai_status),
        ]
        row = 4
        for i, (label, value) in enumerate(kpis):
            col = 1 + (i % 4) * 2
            if i and i % 4 == 0:
                row += 2
            ws.cell(row,col,label).fill=_fill(C_GRIS_FIL); ws.cell(row,col).font=_fnt(bold=True,size=8); ws.cell(row,col).border=_brd(); ws.cell(row,col).alignment=_aln('center','center')
            ws.cell(row+1,col,value).font=_fnt(bold=True,size=10); ws.cell(row+1,col).border=_brd(); ws.cell(row+1,col).alignment=_aln('center','center')
            if isinstance(value,(int,float)) and label in ('Contratista','Mercado','Diferencia'):
                ws.cell(row+1,col).number_format='$#,##0.00'
            if isinstance(value,(int,float)) and label == 'Diferencia %':
                ws.cell(row+1,col).number_format='0.00%'
            if label == 'Diferencia' and isinstance(value,(int,float)):
                ws.cell(row+1,col).fill=_fill('FCE4D6' if value > 0 else 'E2F0D9')
            else:
                ws.cell(row+1,col).fill=_fill('FFFFFF')
        row = 9

        # Semáforo ejecutivo
        band = 'SIN DATOS'
        band_color = 'D9E1F2'
        if diff_pct is not None:
            ad = abs(diff_pct)
            if ad <= 0.05:
                band, band_color = 'BAJO RIESGO / ALINEADO A MERCADO', 'C6EFCE'
            elif ad <= 0.15:
                band, band_color = 'REVISIÓN RECOMENDADA', 'FFEB9C'
            else:
                band, band_color = 'ALTA PRIORIDAD DE REVISIÓN', 'FFC7CE'
        ws.merge_cells(start_row=row, start_column=1, end_row=row+1, end_column=8)
        ws.cell(row,1, f'SEMÁFORO EJECUTIVO: {band}')
        ws.cell(row,1).fill = _fill(band_color)
        ws.cell(row,1).font = _fnt(bold=True, color=C_NEGRO_TXT, size=13)
        ws.cell(row,1).alignment = _aln('center','center',wrap=True)
        ws.cell(row,1).border = _brd()
        row += 3

        # Conclusión ejecutiva
        row = _step37_write_section_header(ws, row, '1. Conclusión ejecutiva', cols=8)
        conclusion = parsed.get('conclusion') if isinstance(parsed, dict) else None
        if not conclusion:
            if diff_pct is None:
                conclusion = 'No fue posible calcular una diferencia porcentual global porque el importe de mercado es cero o no disponible. Revisar calidad de catálogos y matches.'
            elif diff_pct > 0:
                conclusion = f'La cotización se ubica {diff_pct:.2%} por arriba de la referencia de mercado calculada. Priorizar la revisión de los servicios con mayor impacto económico antes de negociar.'
            else:
                conclusion = f'La cotización se ubica {abs(diff_pct):.2%} por debajo de la referencia de mercado calculada. Validar posibles omisiones, rendimientos agresivos o alcances no comparables.'
        ws.merge_cells(start_row=row, start_column=1, end_row=row+2, end_column=8)
        ws.cell(row,1,conclusion); ws.cell(row,1).font=_fnt(size=10); ws.cell(row,1).alignment=_aln('left','top',wrap=True); ws.cell(row,1).border=_brd()
        row += 4

        # Top diferenciales
        row = _step37_write_section_header(ws, row, '2. Servicios con mayor diferencial económico', cols=8)
        data=[]
        for r in metrics['top_diferenciales'][:10]:
            data.append([r['servicio'], (r['descripcion'] or '')[:80], r['cantidad'], r['importe_contratista'], r['importe_mercado'], r['diferencia'], r['diferencia_pct'] if r['diferencia_pct'] is not None else '', 'Revisar' if abs(r['diferencia']) > 0 else 'OK'])
        row = _step37_write_table(ws, row, ['Servicio','Descripción corta','Cant.','Importe contratista','Importe mercado','Diferencia','Dif %','Acción'], data, widths={'A':16,'B':48,'C':10,'D':16,'E':16,'F':16,'G':12,'H':18}, money_cols={4,5,6}, pct_cols={7}) + 1

        # Riesgos prioritarios
        row = _step37_write_section_header(ws, row, '3. Riesgos prioritarios / hallazgos accionables', cols=8)
        risk_rows=[]
        ai_risks = parsed.get('riesgos_prioritarios') if isinstance(parsed, dict) else None
        if ai_risks:
            for rr in ai_risks[:10]:
                risk_rows.append([rr.get('servicio',''), rr.get('riesgo',''), rr.get('impacto',''), 'IA experta'])
        else:
            for issue in metrics['issues'][:10]:
                risk_rows.append([issue.get('servicio'), issue.get('elemento'), f"{issue.get('estado')} | Impacto aprox. ${issue.get('impacto',0):,.2f}", issue.get('nota') or 'Revisar match/unidad/alcance'])
        row = _step37_write_table(ws, row, ['Servicio','Riesgo / elemento','Impacto','Comentario'], risk_rows, widths={'A':16,'B':52,'C':28,'D':55}, money_cols=set(), pct_cols=set(), color_header='9E480E') + 1

        # Lectura por familias de costo
        row = _step37_write_section_header(ws, row, '4. Lectura por familia de costo', cols=8)
        domains = sorted(set(metrics['domain_provider']) | set(metrics['domain_market']))
        fam_rows=[]
        for d in domains:
            pv = metrics['domain_provider'].get(d,0.0); mv = metrics['domain_market'].get(d,0.0); dd=pv-mv; pp=dd/mv if mv else None
            fam_rows.append([d,pv,mv,dd,pp if pp is not None else '', 'Mayor atención' if abs(dd) > 0 else 'OK'])
        row = _step37_write_table(ws, row, ['Familia','Contratista','Mercado','Diferencia','Dif %','Lectura'], fam_rows, widths={'A':26,'B':16,'C':16,'D':16,'E':12,'F':24}, money_cols={2,3,4}, pct_cols={5}) + 1

        # Recomendaciones
        row = _step37_write_section_header(ws, row, '5. Recomendaciones para revisión PMD', cols=8)
        recs = parsed.get('recomendaciones') if isinstance(parsed, dict) else None
        if not recs:
            recs = [
                'Validar primero los servicios con mayor diferencial económico, no los de mayor porcentaje aislado.',
                'Revisar manualmente todos los renglones con unidad incompatible o sin match de mercado.',
                'Confirmar que los porcentajes de indirecto, financiamiento y utilidad usados en mercado correspondan a los declarados por el contratista.',
                'Usar la Matriz Recomendada de Mercado como análisis estructural, no como sustituto automático de la matriz declarada.',
            ]
        for i, rec in enumerate(recs[:8], 1):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(row,1,f'{i}. {rec}'); ws.cell(row,1).border=_brd(); ws.cell(row,1).font=_fnt(size=9); ws.cell(row,1).alignment=_aln('left','center',wrap=True)
            row += 1

        # Nota de gobernanza IA
        row += 1
        row = _step37_write_section_header(ws, row, '6. Gobernanza de IA', cols=8, color='595959')
        gov = [
            ['Uso de IA', 'Claude redacta conclusiones sobre métricas ya calculadas; no calcula precios ni modifica costos.'],
            ['Materiales / MO / equipo', 'Matching determinístico por catálogo, unidad y fallback. No Claude por insumo.'],
            ['Política activa', json.dumps(policy, ensure_ascii=False, default=str)[:900]],
        ]
        row = _step37_write_table(ws, row, ['Tema','Criterio'], gov, widths={'A':24,'B':110}, color_header='595959')

        # Estilo general
        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = ws.column_dimensions[get_column_letter(col)].width or 16
        ws.freeze_panes = 'A9'
        # Orden de tabs
        desired=['Comparativa','Detalle','Matriz Recomendada Mercado','Analisis experto IA']
        wb._sheets=[wb[n] for n in desired if n in wb.sheetnames] + [sh for sh in wb._sheets if sh.title not in desired]
        wb.save(output)
    except Exception as exc:
        try:
            wb = openpyxl.load_workbook(output)
            if 'Analisis experto IA' in wb.sheetnames:
                del wb['Analisis experto IA']
            ws = wb.create_sheet('Analisis experto IA')
            ws['A1'] = 'ERROR AL GENERAR RESUMEN EJECUTIVO DE COSTOS'
            ws['A2'] = f'{type(exc).__name__}: {exc}'
            ws.column_dimensions['A'].width = 120
            wb.save(output)
        except Exception:
            pass
        _q_perf_log('step37_executive_ai_sheet_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
    return output

# Sustituir el wrapper vigente para que el tab final sea el resumen ejecutivo profesional.
_STEP37_BUILD_PMD_PREV = _step18_build_single_provider_pmd

def _step18_build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP37_BUILD_PMD_PREV(output, meta, dato, provider_analysis)
    return _step37_add_executive_ai_sheet(result, dato, provider_analysis)

QUANTIA_EXPERT_REPORT_VERSION = 'step37_executive_cost_report_ai_controlled'


# ============================================================
# STEP40 HOTFIX - Credibilidad de matching + resumen ejecutivo estable
# ============================================================
# 1) Materiales críticos: COSTAL DE YUTE no puede empatar con COSTAL DE RAFIA.
#    Se agregan términos obligatorios que deben preservarse entre proveedor y catálogo.
# 2) Cuadrillas / MO agrupada: si el costo/importe de mercado queda mayor que el
#    declarado por contratista, se usa el contratista como fallback conservador.
# 3) El tab "Analisis experto IA" se simplifica a un dictamen humano ejecutivo,
#    sin tablas pesadas ni descripciones kilométricas de servicios.

_STEP40_CRITICAL_TERMS = {
    'yute', 'rafia', 'pvc', 'cpvc', 'cobre', 'galvanizado', 'inox', 'inoxidable',
    'hidratada', 'hidraulica', 'negro', 'blanco', 'aluminio', 'acero', 'nylon'
}
_STEP40_MUTUALLY_EXCLUSIVE = [
    ({'yute'}, {'rafia'}),
    ({'pvc'}, {'cpvc'}),
    ({'hidratada'}, {'hidraulica'}),
    ({'cobre'}, {'galvanizado'}),
]


def _step40_tokens_from_text(*parts):
    txt = _normalize_text(' '.join(str(p or '') for p in parts))
    try:
        toks = set(_tokenize_text(txt))
    except Exception:
        toks = set(txt.split())
    # Normalizaciones simples para plurales/acentos ya normalizados por _normalize_text.
    if 'inoxidable' in toks:
        toks.add('inox')
    return toks


def _step40_required_terms_ok(item, cand):
    """Valida que términos materiales críticos del proveedor no se pierdan.

    Ejemplo: COSTAL DE YUTE puede empatar con otro YUTE, pero nunca con RAFIA.
    Si no se valida, se descarta el candidato y se usa fallback contratista.
    """
    it = _step40_tokens_from_text(item.get('codigo'), item.get('descripcion'), item.get('insumo'))
    ct = set(cand.get('tokens') or []) | _step40_tokens_from_text(cand.get('codigo'), cand.get('descripcion'))
    # Si el proveedor declara un término crítico, el candidato debe conservarlo.
    for t in sorted(_STEP40_CRITICAL_TERMS & it):
        if t not in ct:
            return False, f'termino critico no preservado: {t}'
    # Reglas de exclusión explícita.
    for a, b in _STEP40_MUTUALLY_EXCLUSIVE:
        if (it & a and ct & b) or (it & b and ct & a):
            return False, f'incompatibilidad semantica: {sorted((it&a)|(it&b))} vs {sorted((ct&a)|(ct&b))}'
    return True, 'ok'

try:
    _STEP40_DIRECT_MATCH_PREV = _step09_direct_catalog_match
except Exception:
    _STEP40_DIRECT_MATCH_PREV = None


def _step09_direct_catalog_match(item, catalogs, domain):
    selected, candidates = _STEP40_DIRECT_MATCH_PREV(item, catalogs, domain) if _STEP40_DIRECT_MATCH_PREV else (None, [])
    if domain != 'material':
        return selected, candidates
    filtered = []
    rejection_reason = None
    for cand in candidates or []:
        ok, reason = _step40_required_terms_ok(item or {}, cand or {})
        if ok:
            filtered.append(cand)
        else:
            rejection_reason = reason
    if selected:
        ok, reason = _step40_required_terms_ok(item or {}, selected or {})
        if ok:
            return selected, filtered[:8]
        rejection_reason = reason
    if rejection_reason:
        # Dejar candidatos filtrados para trazabilidad, pero no seleccionar si todos fallan.
        for cand in filtered:
            cand.setdefault('match_reason', 'candidato conserva terminos criticos')
        return (filtered[0] if filtered and (filtered[0].get('score') or 0) >= 0.50 else None), filtered[:8]
    return selected, candidates


def _step40_is_individual_labor_role(item):
    """Roles humanos individuales que NO deben caer en fallback de cuadrilla.

    En Neodata/Construdata una cuadrilla/brigada es una MO agrupada.
    Supervisor, segurista, residente, operador, oficial, ayudante, etc. son
    renglones de MO individuales cuando aparecen como concepto propio.
    """
    txt = _normalize_text(' '.join(str(item.get(k) or '') for k in ('codigo','descripcion','section','domain','tipo')))
    if 'cuadrilla' in txt or 'brigada' in txt or 'mano obra agrupada' in txt or 'mano de obra agrupada' in txt:
        return False
    individual_terms = [
        'supervisor de seguridad', 'supervisor seguridad', 'supervisor de obra', 'supervisor obra',
        'segurista', 'residente', 'cabo de oficios', 'oficial', 'ayudante', 'peon',
        'operador', 'soldador', 'tubero', 'albanil', 'electricista', 'topografo'
    ]
    return any(t in txt for t in individual_terms)


def _step40_is_cuadrilla_like(item):
    """Detecta solo MO agrupada real para fallback conservador.

    No basta con que el texto contenga 'supervisor' o 'supervision'.
    Esos renglones pueden existir como mano de obra individual en Construdata
    y deben usar match directo al catalogo.
    """
    if _step40_is_individual_labor_role(item):
        return False
    txt = _normalize_text(' '.join(str(item.get(k) or '') for k in ('codigo','descripcion','section','domain','tipo','match_reason')))
    return any(x in txt for x in [
        'cuadrilla', 'brigada', 'personal tecnico',
        'mano obra agrupada', 'mano de obra agrupada'
    ])


def _step40_num(v):
    try:
        if v is None or v == '':
            return None
        return float(v)
    except Exception:
        return None


def _step40_provider_price(item):
    for k in ('precio_proveedor', 'precio_base', 'precio', 'costo_proveedor', 'costo'):
        v = _step40_num(item.get(k))
        if v is not None and v > 0:
            return v
    return None


def _step40_market_price(item):
    for k in ('precio_mercado', 'precio_mercado_catalogo', 'costo_mercado'):
        v = _step40_num(item.get(k))
        if v is not None and v > 0:
            return v
    return None


def _step40_provider_amount(item):
    for k in ('importe_proveedor', 'importe_contratista', 'importe'):
        v = _step40_num(item.get(k))
        if v is not None and v >= 0:
            return v
    return None


def _step40_market_amount(item):
    v = _step40_num(item.get('importe_mercado'))
    return v if v is not None else None


def _step40_apply_cuadrilla_conservative_fallback(conceptos):
    """Para cuadrillas/MO agrupada, el mercado solo sustituye si es menor.

    Si el catálogo o cálculo de roles produce un valor superior al declarado por el
    contratista, se conserva el costo/importe del contratista para no inflar mercado.
    """
    for c in (conceptos or {}).values():
        if not isinstance(c, dict):
            continue
        changed = False
        for item in c.get('granular_market_items') or []:
            if not isinstance(item, dict) or not _step40_is_cuadrilla_like(item):
                continue
            pa = _step40_provider_amount(item)
            ma = _step40_market_amount(item)
            pp = _step40_provider_price(item)
            mp = _step40_market_price(item)
            if pa is not None and ma is not None and pa > 0 and ma > pa:
                item['importe_mercado_original'] = ma
                item['importe_mercado'] = pa
                if pp is not None:
                    item['precio_mercado_original'] = item.get('precio_mercado')
                    item['precio_mercado'] = pp
                item['match_status'] = 'FALLBACK_CUADRILLA_CONSERVADOR'
                item['match_reason'] = (item.get('match_reason') or '') + ' | STEP40: cuadrilla/MO agrupada conserva contratista porque mercado calculado excede lo declarado.'
                changed = True
            elif pp is not None and mp is not None and mp > pp:
                item['precio_mercado_original'] = mp
                item['precio_mercado'] = pp
                item['match_status'] = 'FALLBACK_CUADRILLA_CONSERVADOR'
                item['match_reason'] = (item.get('match_reason') or '') + ' | STEP40: costo unitario de cuadrilla conserva contratista porque mercado excede lo declarado.'
                changed = True
        if changed:
            # Recalcular directo de mercado por seccion si existe el helper de estructura.
            try:
                if '_step36_recalculate_structured_market' in globals():
                    _step36_recalculate_structured_market(c)
                else:
                    c['granular_market_direct'] = sum(_step40_market_amount(x) or 0.0 for x in c.get('granular_market_items') or [])
            except Exception:
                c['granular_market_direct'] = sum(_step40_market_amount(x) or 0.0 for x in c.get('granular_market_items') or [])
    return conceptos

try:
    _STEP40_APPLY_MARKET_PREV = _step08_apply_market_catalog_pricing
except Exception:
    _STEP40_APPLY_MARKET_PREV = None


def _step08_apply_market_catalog_pricing(conceptos):
    conceptos = _STEP40_APPLY_MARKET_PREV(conceptos) if _STEP40_APPLY_MARKET_PREV else conceptos
    return _step40_apply_cuadrilla_conservative_fallback(conceptos)


def _step40_service_key(s):
    return str(s or '').strip().split('|')[0][:24]


def _step40_executive_context(metrics):
    top = []
    for r in (metrics.get('top_diferenciales') or [])[:6]:
        top.append({
            'servicio': _step40_service_key(r.get('servicio')),
            'importe_contratista': r.get('importe_contratista'),
            'importe_mercado': r.get('importe_mercado'),
            'diferencia': r.get('diferencia'),
            'diferencia_pct': r.get('diferencia_pct'),
        })
    issues = []
    for it in (metrics.get('issues') or [])[:10]:
        issues.append({
            'servicio': _step40_service_key(it.get('servicio')),
            'tipo': it.get('tipo'),
            'estado': it.get('estado'),
            'impacto': it.get('impacto'),
            'elemento': str(it.get('elemento') or '')[:80],
        })
    return {
        'total_servicios': metrics.get('total_servicios'),
        'total_contratista': metrics.get('total_contratista'),
        'total_mercado': metrics.get('total_mercado'),
        'diferencia_total': metrics.get('diferencia_total'),
        'diferencia_pct_total': metrics.get('diferencia_pct_total'),
        'items_sin_match': metrics.get('no_match'),
        'unidad_incompatible': metrics.get('no_match_unit'),
        'top_diferenciales': top,
        'hallazgos': issues,
        'familias_contratista': metrics.get('domain_provider'),
        'familias_mercado': metrics.get('domain_market'),
    }


def _step40_fallback_executive_text(metrics):
    total_provider = metrics.get('total_contratista') or 0
    total_market = metrics.get('total_mercado') or 0
    diff = metrics.get('diferencia_total') or 0
    pct = metrics.get('diferencia_pct_total')
    top = metrics.get('top_diferenciales') or []
    issues = metrics.get('issues') or []
    direction = 'por arriba' if diff > 0 else 'por debajo' if diff < 0 else 'alineada'
    pct_txt = f'{abs(pct):.2%}' if isinstance(pct, (int, float)) else 'sin porcentaje confiable'
    lines = []
    lines.append(f'La cotización se observa {direction} de la referencia de mercado en {pct_txt}. El diferencial debe revisarse por impacto económico, no solo por variación porcentual aislada.')
    if top[:3]:
        claves = ', '.join(_step40_service_key(r.get('servicio')) for r in top[:3])
        lines.append(f'La revisión prioritaria debe concentrarse en {claves}, ya que estos servicios explican la mayor parte del diferencial económico calculado.')
    if issues:
        unit = sum(1 for i in issues if 'UNIDAD' in str(i.get('estado') or '').upper())
        nom = sum(1 for i in issues if 'SIN_MATCH' in str(i.get('estado') or '').upper() or 'NO_MATCH' in str(i.get('estado') or '').upper())
        lines.append(f'Se detectan {len(issues)} hallazgos relevantes en insumos. De ellos, {unit} están asociados a unidad incompatible y {nom} a falta de match confiable; estos casos deben conservar fallback al contratista hasta validar catálogo.')
    lines.append('El mercado se calculó con reglas determinísticas: precios por catálogo granular, validación de unidad, fallback conservador e indirecto de mercado fijo cuando aplica. Claude no modifica precios ni importes.')
    lines.append('Recomendación: usar este reporte para negociar primero los servicios de mayor impacto, validar renglones sin match/unidad incompatible y revisar cualquier cuadrilla donde el mercado calculado supere el costo declarado por el contratista.')
    return '\n\n'.join(lines)


def _step40_claude_executive_text(metrics):
    try:
        import quantia_ai_review as qair
        policy = qair.ai_runtime_policy()
        if not policy.get('expert_review_enabled') or not policy.get('anthropic_key_set'):
            return None, 'Claude desactivado por politica'
        payload = {
            'rol': 'Actua como analista senior de precios unitarios y costos de obra. Tu lector es un PM o comprador tecnico.',
            'instrucciones': [
                'Redacta un dictamen ejecutivo en español, humano y accionable.',
                'No repitas nombres largos de servicios; usa solo claves/codigos.',
                'No uses tablas. No menciones que eres IA.',
                'No inventes datos. No calcules precios. Interpreta las metricas calculadas.',
                'Maximo 7 parrafos cortos. Enfocate en dinero, riesgo y acciones PMD.',
            ],
            'contexto_calculado': _step40_executive_context(metrics),
            'salida': 'Texto plano ejecutivo, sin markdown pesado.'
        }
        raw = qair._anthropic_message(payload, max_tokens=int(os.getenv('QUANTIA_AI_EXPERT_MAX_TOKENS','900')), timeout=int(os.getenv('QUANTIA_AI_EXPERT_TIMEOUT','30')))
        if raw and str(raw).strip():
            return str(raw).strip(), 'Claude activo'
        return None, 'Claude no devolvio texto'
    except Exception as exc:
        return None, f'Claude error: {type(exc).__name__}: {str(exc)[:160]}'


def _step37_add_executive_ai_sheet(output, dato, provider_analysis=None):
    """STEP40: hoja de dictamen simple y robusta.

    No usa estilos incompatibles, no genera tablas extensas y no rompe el Excel si
    Claude falla. Produce texto útil para un analista PU.
    """
    try:
        wb = openpyxl.load_workbook(output)
        if 'Analisis experto IA' in wb.sheetnames:
            del wb['Analisis experto IA']
        ws = wb.create_sheet('Analisis experto IA')
        ws.sheet_view.showGridLines = False
        metrics = _step37_collect_executive_metrics(dato or {}, provider_analysis)
        claude_text, status = _step40_claude_executive_text(metrics)
        final_text = claude_text or _step40_fallback_executive_text(metrics)
        ws.merge_cells('A1:F1')
        ws['A1'] = 'DICTAMEN EJECUTIVO DE PRECIOS UNITARIOS'
        ws['A1'].fill = _fill(C_AZUL_OSC)
        ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=14)
        ws['A1'].alignment = _aln('center','center',wrap=True)
        ws.merge_cells('A2:F2')
        ws['A2'] = 'Lectura ejecutiva enfocada en impacto económico, riesgos de costo y acciones de revisión PMD.'
        ws['A2'].font = _fnt(color=C_GRIS_TXT, size=9)
        ws['A2'].alignment = _aln('center','center',wrap=True)
        # KPIs mínimos, no tabla pesada.
        total_provider = metrics.get('total_contratista') or 0
        total_market = metrics.get('total_mercado') or 0
        diff = metrics.get('diferencia_total') or 0
        pct = metrics.get('diferencia_pct_total')
        kpis = [
            ('Servicios', metrics.get('total_servicios')),
            ('Contratista', total_provider),
            ('Mercado', total_market),
            ('Diferencia', diff),
            ('Dif. %', pct if isinstance(pct,(int,float)) else ''),
            ('Estado Claude', status),
        ]
        row = 4
        for i,(label,value) in enumerate(kpis):
            col = 1 + i
            ws.cell(row,col,label).fill = _fill(C_GRIS_FIL)
            ws.cell(row,col).font = _fnt(bold=True,size=8)
            ws.cell(row,col).border = _brd()
            ws.cell(row+1,col,value).font = _fnt(bold=True,size=10)
            ws.cell(row+1,col).border = _brd()
            ws.cell(row,col).alignment = _aln('center','center',wrap=True)
            ws.cell(row+1,col).alignment = _aln('center','center',wrap=True)
            if label in ('Contratista','Mercado','Diferencia'):
                ws.cell(row+1,col).number_format = '$#,##0.00'
            if label == 'Dif. %' and isinstance(value,(int,float)):
                ws.cell(row+1,col).number_format = '0.00%'
        ws.merge_cells('A7:F18')
        ws['A7'] = final_text
        ws['A7'].alignment = _aln('left','top',wrap=True)
        ws['A7'].font = _fnt(size=11, color=C_NEGRO_TXT)
        ws['A7'].border = _brd()
        ws['A7'].fill = _fill('FFFFFF')
        # Top servicios solo como referencia breve por clave.
        top_keys = ', '.join(_step40_service_key(r.get('servicio')) for r in (metrics.get('top_diferenciales') or [])[:6])
        ws.merge_cells('A20:F21')
        ws['A20'] = f'Servicios sugeridos para revision prioritaria: {top_keys or "sin diferenciales relevantes"}'
        ws['A20'].alignment = _aln('left','center',wrap=True)
        ws['A20'].font = _fnt(bold=True,size=10)
        ws['A20'].fill = _fill('D9EAF7')
        ws['A20'].border = _brd()
        ws.merge_cells('A23:F24')
        ws['A23'] = 'Gobernanza: este dictamen no modifica precios ni importes. Los costos provienen de reglas deterministicas, catalogos Construdata, validacion de unidad y fallback conservador.'
        ws['A23'].alignment = _aln('left','center',wrap=True)
        ws['A23'].font = _fnt(size=9, color=C_GRIS_TXT)
        for col in range(1,7):
            ws.column_dimensions[get_column_letter(col)].width = 22
        ws.row_dimensions[7].height = 220
        desired = ['Comparativa','Detalle','Matriz Recomendada Mercado','Analisis experto IA']
        wb._sheets = [wb[n] for n in desired if n in wb.sheetnames] + [sh for sh in wb._sheets if sh.title not in desired]
        wb.save(output)
    except Exception as exc:
        try:
            wb = openpyxl.load_workbook(output)
            if 'Analisis experto IA' in wb.sheetnames:
                del wb['Analisis experto IA']
            ws = wb.create_sheet('Analisis experto IA')
            ws['A1'] = 'DICTAMEN EJECUTIVO DE PRECIOS UNITARIOS'
            ws['A2'] = _step40_fallback_executive_text(_step37_collect_executive_metrics(dato or {}, provider_analysis))
            ws['A4'] = f'Nota tecnica: se uso fallback deterministico por error de formato ({type(exc).__name__}: {exc}).'
            ws.column_dimensions['A'].width = 120
            ws['A2'].alignment = _aln('left','top',wrap=True)
            wb.save(output)
        except Exception:
            pass
    return output

QUANTIA_EXPERT_REPORT_VERSION = 'step40_hotfix_human_executive_summary'
QUANTIA_MATCHING_HOTFIX_VERSION = 'step42_mo_roles_no_cuadrilla_codigo_mo_v3'

# ============================================================
# STEP43 CONSOLIDADO - DataDir robusto + diagnóstico catálogos + Detalle sin amarillo indiscriminado
# ============================================================
# Contexto profesional:
# - En Docker/Railway DATA_DIR puede apuntar a /data. Si ese volumen existe pero está
#   vacío, el motor anterior dejaba de usar la carpeta data empacada y todos los
#   insumos caían en fallback contratista, aun cuando los catálogos sí estaban en el repo.
# - En el tab Detalle se pintaban en amarillo todas las columnas de mercado; ahora
#   solo se resalta Costo Mercado cuando es numéricamente distinto al costo del contratista.

_STEP43_CATALOG_PATTERNS = {
    'material': ('construdata-materiales*.xlsx',),
    'mano_obra': ('construdata-manodeobra*.xlsx',),
    'equipo': ('construdata-maquinaria*.xlsx',),
}

try:
    _STEP43_QUANTIA_DATA_DIR_PREV = _quantia_data_dir
except Exception:
    _STEP43_QUANTIA_DATA_DIR_PREV = None


def _step43_catalog_file_score(path):
    """Devuelve cuantos dominios de catálogo existen en path."""
    try:
        p = Path(path)
        if not p.exists() or not p.is_dir():
            return 0, {}
        found = {}
        for domain, patterns in _STEP43_CATALOG_PATTERNS.items():
            files = []
            for pat in patterns:
                files.extend(sorted(p.glob(pat)))
            found[domain] = [str(x.name) for x in files if x.is_file()]
        score = sum(1 for files in found.values() if files)
        return score, found
    except Exception:
        return 0, {}


def _quantia_data_dir() -> Path:
    """Resolver DATA_DIR de forma segura.

    Si DATA_DIR=/data existe pero está vacío, no debe bloquear el fallback a ./data
    empacado. Solo se fuerza el directorio configurado vacío si STRICT_DATA_DIR=1.
    """
    strict = str(os.getenv('STRICT_DATA_DIR', '0')).strip().lower() in {'1','true','yes','si','sí','on'}
    candidates = []
    env_candidates = []
    for key in ('DATA_DIR', 'QUANTIA_DATA_DIR'):
        raw = os.getenv(key)
        if raw:
            env_candidates.append(Path(raw).expanduser())
    candidates.extend(env_candidates)
    try:
        candidates.append(Path(__file__).resolve().parent / 'data')
        candidates.append(Path(__file__).resolve().parent.parent / 'data')
    except Exception:
        pass
    candidates.append(Path.cwd() / 'data')
    candidates.append(Path('data'))

    # Deduplicar preservando orden.
    dedup = []
    seen = set()
    for c in candidates:
        try:
            key = str(c.resolve()) if c.exists() else str(c)
        except Exception:
            key = str(c)
        if key not in seen:
            dedup.append(c); seen.add(key)

    # Preferir directorios con catálogos reales.
    best = None
    best_score = -1
    for c in dedup:
        score, _found = _step43_catalog_file_score(c)
        if score > best_score:
            best, best_score = c, score
        if score >= 3:
            return c

    # Si el usuario fuerza un DATA_DIR, devolverlo aunque no tenga catálogos.
    if strict and env_candidates:
        return env_candidates[0]
    if best is not None and best.exists():
        return best
    if _STEP43_QUANTIA_DATA_DIR_PREV:
        try:
            return _STEP43_QUANTIA_DATA_DIR_PREV()
        except Exception:
            pass
    return Path('data')


def quantia_market_catalog_diagnostics(force=False):
    """Diagnóstico compacto de catálogos granulares Construdata usados por Detalle."""
    diag = {
        'version': 'step43_datadir_catalog_diagnostics',
        'env_DATA_DIR': os.getenv('DATA_DIR'),
        'env_QUANTIA_DATA_DIR': os.getenv('QUANTIA_DATA_DIR'),
        'strict_data_dir': os.getenv('STRICT_DATA_DIR', '0'),
        'resolved_data_dir': str(_quantia_data_dir()),
        'files': {},
        'record_counts': {},
        'status': 'unknown',
        'warning': None,
    }
    try:
        cats = load_step08_market_catalogs(force=force)
        files = cats.get('files') or {}
        by_domain = cats.get('by_domain') or {}
        for domain, path in files.items():
            p = Path(path)
            diag['files'][domain] = {'path': str(p), 'exists': p.exists(), 'name': p.name}
        diag['record_counts'] = {domain: len(rows or []) for domain, rows in by_domain.items()}
        total = sum(diag['record_counts'].values())
        diag['status'] = 'ok' if total > 0 else 'sin_registros'
        if total == 0:
            diag['warning'] = 'No se cargaron registros de catálogos; Detalle usará fallback contratista.'
        missing = [d for d in ('material','mano_obra','equipo') if not diag['record_counts'].get(d)]
        if missing and total > 0:
            diag['warning'] = 'Catálogos parciales sin registros en: ' + ', '.join(missing)
    except Exception as exc:
        diag['status'] = 'error'
        diag['warning'] = f'{type(exc).__name__}: {exc}'
    return diag


def _step43_num(v):
    try:
        if v is None or v == '':
            return None
        return float(v)
    except Exception:
        return None


def _step43_find_header_columns(ws, header_row=3):
    cols = {}
    for cell in ws[header_row]:
        txt = _normalize_text(cell.value or '')
        if txt:
            cols[txt] = cell.column
    return cols


def _step43_apply_detalle_market_highlight(output_path):
    """Quita amarillo global de mercado y resalta solo costo mercado diferente."""
    try:
        wb = openpyxl.load_workbook(output_path)
        if 'Detalle' not in wb.sheetnames:
            wb.save(output_path); return output_path
        ws = wb['Detalle']
        cols = _step43_find_header_columns(ws, 3)
        c_provider = cols.get('costo contratista') or 4
        c_market = cols.get('costo mercado') or 8
        c_diff = cols.get('dif costo') or cols.get('dif % costo') or 9
        c_importe_market = cols.get('importe mercado') or 10
        c_status = cols.get('estado') or 15

        # Encabezados de mercado dejan de ser amarillos por defecto.
        for c in (c_market, c_diff, c_importe_market):
            cell = ws.cell(3, c)
            cell.fill = _fill(C_AZUL_OSC)
            cell.font = _fnt(bold=True, color=C_BLANCO, size=8)
            cell.alignment = _aln('center','center',wrap=True)
            cell.border = _brd()

        for r in range(4, ws.max_row + 1):
            # Limpiar pintura anterior en mercado. Se preservan secciones/merged por valor vacío.
            for c in (c_market, c_diff, c_importe_market):
                ws.cell(r, c).fill = _fill(C_BLANCO)
            provider = _step43_num(ws.cell(r, c_provider).value)
            market = _step43_num(ws.cell(r, c_market).value)
            # Solo filas con costo unitario comparable.
            if provider is None or market is None:
                continue
            tol = max(0.01, abs(provider) * 0.0001)
            if abs(provider - market) > tol:
                ws.cell(r, c_market).fill = _fill(C_AMARILLO_PMD)
                # La diferencia puede quedar sin amarillo para cumplir literal: solo costo mercado.
            # Resaltar estados sin match de forma suave en Estado, no en costo.
            st = str(ws.cell(r, c_status).value or '').lower()
            if st and ('sin_match' in st or 'fallback' in st or 'unidad' in st):
                ws.cell(r, c_status).fill = _fill('FCE4D6')
        wb.save(output_path)
    except Exception as exc:
        try:
            _q_perf_log('step43_detalle_highlight_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
        except Exception:
            pass
    return output_path


def _step43_write_catalog_diagnostic_sheet(output_path):
    try:
        wb = openpyxl.load_workbook(output_path)
        if 'Diagnóstico Catálogos' in wb.sheetnames:
            del wb['Diagnóstico Catálogos']
        ws = wb.create_sheet('Diagnóstico Catálogos')
        ws.sheet_view.showGridLines = False
        ws.merge_cells('A1:H1')
        ws['A1'] = 'DIAGNÓSTICO DE CATÁLOGOS CONSTRUDATA / MATCH DETALLE'
        ws['A1'].fill = _fill(C_AZUL_OSC); ws['A1'].font = _fnt(bold=True, color=C_BLANCO, size=12); ws['A1'].alignment = _aln('center','center')
        diag = quantia_market_catalog_diagnostics(force=False)
        rows = [
            ('Estado', diag.get('status')),
            ('DATA_DIR env', diag.get('env_DATA_DIR')),
            ('QUANTIA_DATA_DIR env', diag.get('env_QUANTIA_DATA_DIR')),
            ('Directorio resuelto', diag.get('resolved_data_dir')),
            ('STRICT_DATA_DIR', diag.get('strict_data_dir')),
            ('Advertencia', diag.get('warning')),
        ]
        r = 3
        for k, v in rows:
            ws.cell(r,1,k); ws.cell(r,2,v)
            ws.cell(r,1).fill = _fill(C_GRIS_FIL); ws.cell(r,1).font = _fnt(bold=True,size=8); ws.cell(r,1).border=_brd()
            ws.cell(r,2).font = _fnt(size=8); ws.cell(r,2).alignment = _aln('left','center',wrap=True); ws.cell(r,2).border=_brd()
            ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=8)
            r += 1
        r += 1
        headers = ['Dominio','Archivo','Existe','Registros cargados','Comentario']
        for c,h in enumerate(headers,1):
            cell = ws.cell(r,c,h); cell.fill=_fill(C_AZUL_OSC); cell.font=_fnt(bold=True,color=C_BLANCO,size=8); cell.border=_brd(); cell.alignment=_aln('center','center')
        r += 1
        for domain in ('material','mano_obra','equipo'):
            info = (diag.get('files') or {}).get(domain) or {}
            count = (diag.get('record_counts') or {}).get(domain, 0)
            vals = [domain, info.get('path'), bool(info.get('exists')), count, 'OK' if count else 'Sin registros: revisar DATA_DIR o archivo fuente']
            for c,val in enumerate(vals,1):
                cell = ws.cell(r,c,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True)
                if c == 5 and not count: cell.fill=_fill('FCE4D6')
            r += 1
        # Resumen de estados existentes en Detalle.
        r += 2
        ws.cell(r,1,'Estados detectados en Detalle'); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
        ws.cell(r,1).fill=_fill(C_GRIS_SEC); ws.cell(r,1).font=_fnt(bold=True,size=9); ws.cell(r,1).border=_brd(); r += 1
        for c,h in enumerate(['Tipo','Estado','Renglones'],1):
            cell=ws.cell(r,c,h); cell.fill=_fill(C_AZUL_OSC); cell.font=_fnt(bold=True,color=C_BLANCO,size=8); cell.border=_brd()
        r += 1
        counts = defaultdict(int)
        if 'Detalle' in wb.sheetnames:
            ws_det = wb['Detalle']
            cols = _step43_find_header_columns(ws_det, 3)
            c_tipo = cols.get('tipo') or 12
            c_estado = cols.get('estado') or 15
            for rr in range(4, ws_det.max_row + 1):
                tipo = ws_det.cell(rr,c_tipo).value
                estado = ws_det.cell(rr,c_estado).value
                if tipo and estado:
                    counts[(str(tipo), str(estado))] += 1
        for (tipo, estado), n in sorted(counts.items(), key=lambda x: (-x[1], x[0]))[:30]:
            for c,val in enumerate([tipo, estado, n],1):
                cell=ws.cell(r,c,val); cell.border=_brd(); cell.font=_fnt(size=8); cell.alignment=_aln('left','center',wrap=True)
            r += 1
        for col, width in {'A':22,'B':80,'C':10,'D':18,'E':58,'F':14,'G':14,'H':14}.items():
            ws.column_dimensions[col].width = width
        try:
            ws.freeze_panes = 'A10'
            ws.auto_filter.ref = f'A10:E{max(r-1,10)}'
        except Exception:
            pass
        # Orden sugerido: dejar diagnóstico cerca de Detalle.
        desired = ['Comparativa','Detalle','Diagnóstico Catálogos','Matriz Recomendada Mercado','Analisis experto IA']
        wb._sheets = [wb[n] for n in desired if n in wb.sheetnames] + [sh for sh in wb._sheets if sh.title not in desired]
        wb.save(output_path)
    except Exception as exc:
        try:
            _q_perf_log('step43_catalog_diagnostic_sheet_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
        except Exception:
            pass
    return output_path

try:
    _STEP43_BUILD_PMD_PREV = _step18_build_single_provider_pmd
except Exception:
    _STEP43_BUILD_PMD_PREV = None


def _step18_build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP43_BUILD_PMD_PREV(output, meta, dato, provider_analysis) if _STEP43_BUILD_PMD_PREV else output
    _step43_apply_detalle_market_highlight(result)
    _step43_write_catalog_diagnostic_sheet(result)
    return result

QUANTIA_STEP43_CONSOLIDADO_VERSION = 'v0.2.3_datadir_detalle_highlight_catalog_diagnostics'

# ============================================================
# STEP44 - Mano de obra individual no debe seleccionar cuadrillas del catálogo
# ============================================================
# Ejemplo detectado: AYU-GRAL / Ayudante General podía seleccionar primero
# "CUADRILLA No 3 (1 AYUDANTE GENERAL)" porque tenía el mismo score que
# "MO021 AYUDANTE GENERAL". Para lógica tipo Neodata, si el proveedor declara
# un rol individual, se debe preferir el rol individual del catálogo.

try:
    _STEP44_DIRECT_MATCH_PREV = _step09_direct_catalog_match
except Exception:
    _STEP44_DIRECT_MATCH_PREV = None


def _step44_catalog_candidate_is_cuadrilla(cand):
    txt = _normalize_text(' '.join(str((cand or {}).get(k) or '') for k in ('codigo','descripcion','tipo_insumo','norm')))
    return 'cuadrilla' in txt or 'brigada' in txt or 'mano obra agrupada' in txt or 'mano de obra agrupada' in txt


def _step44_provider_item_is_explicit_cuadrilla(item):
    txt = _normalize_text(' '.join(str((item or {}).get(k) or '') for k in ('codigo','descripcion','section','domain','tipo')))
    return 'cuadrilla' in txt or 'brigada' in txt or 'mano obra agrupada' in txt or 'mano de obra agrupada' in txt


def _step44_provider_item_is_individual_labor(item):
    try:
        if _step40_is_individual_labor_role(item or {}):
            return True
    except Exception:
        pass
    txt = _normalize_text(' '.join(str((item or {}).get(k) or '') for k in ('codigo','descripcion')))
    if _step44_provider_item_is_explicit_cuadrilla(item):
        return False
    terms = [
        'ayudante general', 'ayudante', 'peon', 'oficial', 'soldador', 'argonero', 'tubero',
        'supervisor', 'segurista', 'residente', 'operador', 'electricista', 'albanil',
        'topografo', 'maniobrista', 'tecnico', 'maestro'
    ]
    return any(t in txt for t in terms)


def _step09_direct_catalog_match(item, catalogs, domain):
    selected, candidates = _STEP44_DIRECT_MATCH_PREV(item, catalogs, domain) if _STEP44_DIRECT_MATCH_PREV else (None, [])
    if domain != 'mano_obra':
        return selected, candidates
    if not _step44_provider_item_is_individual_labor(item or {}):
        return selected, candidates
    # Si el proveedor declaró un rol individual, no seleccionar una cuadrilla por empate de tokens.
    non_cuad = [dict(c) for c in (candidates or []) if not _step44_catalog_candidate_is_cuadrilla(c)]
    cuad = [dict(c) for c in (candidates or []) if _step44_catalog_candidate_is_cuadrilla(c)]
    if not non_cuad:
        return selected, candidates
    non_cuad.sort(key=lambda x: x.get('score') or 0, reverse=True)
    best = non_cuad[0]
    if (best.get('score') or 0) >= 0.34:
        best['match_reason'] = (best.get('match_reason') or 'similitud descripcion/unidad dentro del mismo dominio') + ' | STEP44: rol MO individual; se descarta cuadrilla catalogada como candidato principal.'
        # Conservar cuadrillas al final solo para trazabilidad, no como selección.
        return best, (non_cuad + cuad)[:8]
    return selected, candidates

QUANTIA_STEP44_MO_INDIVIDUAL_OVER_CUADRILLA = 'v0.2.3_step44_mo_individual_prefer_role_catalog'

# STEP45 - desempate fino para roles MO individuales (AYU-GRAL -> MO021, no PEON/cuadrilla)
try:
    _STEP45_DIRECT_MATCH_PREV = _step09_direct_catalog_match
except Exception:
    _STEP45_DIRECT_MATCH_PREV = None

_STEP45_ROLE_PHRASES = [
    'ayudante general', 'supervisor de seguridad', 'supervisor seguridad',
    'supervisor de obra', 'supervisor obra', 'oficial soldador', 'soldador',
    'oficial tubero', 'tubero', 'operador', 'electricista', 'albanil', 'peon'
]
_STEP45_CRITICAL_ROLE_TOKENS = {
    'ayudante', 'general', 'peon', 'supervisor', 'seguridad', 'obra', 'oficial',
    'soldador', 'argonero', 'tubero', 'operador', 'electricista', 'albanil', 'residente', 'segurista'
}


def _step45_role_adjusted_score(item, cand):
    base = float((cand or {}).get('score') or 0.0)
    item_txt = _normalize_text(' '.join(str((item or {}).get(k) or '') for k in ('codigo','descripcion')))
    cand_txt = _normalize_text(' '.join(str((cand or {}).get(k) or '') for k in ('codigo','descripcion')))
    score = base
    for phrase in _STEP45_ROLE_PHRASES:
        if phrase in item_txt and phrase in cand_txt:
            score += 0.35
    it = set(item_txt.split()) & _STEP45_CRITICAL_ROLE_TOKENS
    ct = set(cand_txt.split()) & _STEP45_CRITICAL_ROLE_TOKENS
    score += 0.05 * len(it & ct)
    missing = it - ct
    # Penalizaciones de sustitución que pueden cambiar el rol laboral.
    if 'ayudante' in it and 'ayudante' not in ct:
        score -= 0.25
    if 'peon' in it and 'peon' not in ct:
        score -= 0.18
    for t in ('tubero','soldador','seguridad','obra','operador','electricista','albanil'):
        if t in it and t not in ct:
            score -= 0.18
    if 'general' in it and 'ayudante' in it and 'general' not in ct:
        score -= 0.15
    if _step44_catalog_candidate_is_cuadrilla(cand):
        score -= 0.30
    return score


def _step09_direct_catalog_match(item, catalogs, domain):
    selected, candidates = _STEP45_DIRECT_MATCH_PREV(item, catalogs, domain) if _STEP45_DIRECT_MATCH_PREV else (None, [])
    if domain != 'mano_obra' or not _step44_provider_item_is_individual_labor(item or {}):
        return selected, candidates
    viable = [dict(c) for c in (candidates or []) if not _step44_catalog_candidate_is_cuadrilla(c)]
    if not viable:
        return selected, candidates
    for c in viable:
        c['_role_adjusted_score'] = _step45_role_adjusted_score(item, c)
    viable.sort(key=lambda x: x.get('_role_adjusted_score') or 0, reverse=True)
    best = viable[0]
    if (best.get('_role_adjusted_score') or 0) >= 0.34:
        # Mantener score base para no inflar confianza financiera; dejar ajuste como trazabilidad.
        best['match_reason'] = (best.get('match_reason') or 'similitud descripcion/unidad dentro del mismo dominio') + f" | STEP45: desempate rol MO individual (score_rol={best.get('_role_adjusted_score'):.3f})."
        others = [c for c in (candidates or []) if c not in viable]
        return best, (viable + others)[:8]
    return selected, candidates

QUANTIA_STEP45_MO_ROLE_TIEBREAK = 'v0.2.3_step45_labor_role_tiebreak'

# ============================================================
# STEP46 - Presentacion para analistas: ocultar hojas tecnicas por configuracion
# ============================================================
# Solo modifica estado/nombre de pestañas del Excel final. No cambia cálculos,
# fórmulas, importes ni datos generados por el motor.

try:
    _STEP46_BUILD_PMD_PREV = _step18_build_single_provider_pmd
except Exception:
    _STEP46_BUILD_PMD_PREV = None


def _step46_env_bool(name, default=False):
    raw = os.getenv(name)
    if raw is None:
        return bool(default)
    return str(raw).strip().lower() in {'1','true','yes','si','sí','on'}


def _step46_hide_internal_sheets_enabled():
    return _step46_env_bool('EXCEL_HIDE_INTERNAL_SHEETS', True) and not _step46_env_bool('EXCEL_SHOW_INTERNAL_SHEETS', False)


def _step46_internal_sheet_names():
    raw = os.getenv('EXCEL_INTERNAL_SHEETS', '').strip()
    if raw:
        names = [x.strip() for x in raw.split(',') if x.strip()]
    else:
        names = [
            'Alcance y Riesgos',
            'Diagnóstico Catálogos',
            'Diagnostico Catalogos',
            'Trazabilidad Técnica',
        ]
    if 'Trazabilidad Técnica' not in names:
        names.append('Trazabilidad Técnica')
    return names


def _step46_apply_analyst_excel_presentation(output_path):
    try:
        wb = openpyxl.load_workbook(output_path)
        changed = False
        for name in ('Resumen Profesional', 'Comparativa', 'Detalle', 'Analisis experto IA'):
            if name in wb.sheetnames:
                wb[name].sheet_state = 'visible'
                changed = True
        if _step46_hide_internal_sheets_enabled():
            for name in _step46_internal_sheet_names():
                if name in wb.sheetnames:
                    wb[name].sheet_state = 'hidden'
                    changed = True
        visible = [ws for ws in wb.worksheets if ws.sheet_state == 'visible']
        if not visible:
            for candidate in ('Resumen Profesional', 'Comparativa', 'Detalle', 'Analisis experto IA'):
                if candidate in wb.sheetnames:
                    wb[candidate].sheet_state = 'visible'
                    visible = [wb[candidate]]
                    changed = True
                    break
        if visible:
            try:
                wb.active = wb.worksheets.index(visible[0])
            except Exception:
                pass
        if changed:
            wb.save(output_path)
    except Exception as exc:
        try:
            _q_perf_log('step46_analyst_presentation_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
        except Exception:
            pass
    return output_path


def _step18_build_single_provider_pmd(output, meta, dato, provider_analysis):
    result = _STEP46_BUILD_PMD_PREV(output, meta, dato, provider_analysis) if _STEP46_BUILD_PMD_PREV else output
    return _step46_apply_analyst_excel_presentation(result)


QUANTIA_STEP46_ANALYST_PRESENTATION = 'v0.2.4_analyst_excel_presentation_only'

# ============================================================
# STEP41 PMD - Reglas de analisis experto PMD para precios unitarios
# ============================================================
# Este bloque NO modifica calculos, PU, importes ni matching. Solo agrega
# metricas y texto de analisis al tab "Analisis experto IA" y al contexto
# que se envia a Claude cuando esta habilitado.

PMD_MARKET_INDIRECT_PCT = float(os.getenv('PMD_MARKET_INDIRECT_PCT', os.getenv('MARKET_INDIRECT_PCT', '0.25')) or 0.25)
PMD_LABOR_MARKET_YEAR = int(float(os.getenv('PMD_LABOR_MARKET_YEAR', '2026') or 2026))
PMD_LABOR_DOF_ADJUSTMENT_PCT = float(os.getenv('PMD_LABOR_DOF_ADJUSTMENT_PCT', '0.13') or 0.13)
PMD_TOOL_MINOR_MARKET_PCT = float(os.getenv('PMD_TOOL_MINOR_MARKET_PCT', '0.05') or 0.05)
PMD_EPP_MARKET_PCT = float(os.getenv('PMD_EPP_MARKET_PCT', '0.10') or 0.10)
PMD_PARETO_TARGET_PCT = float(os.getenv('PMD_PARETO_TARGET_PCT', '0.80') or 0.80)


def _pmd_num(v, default=None):
    try:
        if v is None or v == '':
            return default
        return float(v)
    except Exception:
        return default


def _pmd_pct(v, default=None):
    n = _pmd_num(v, default)
    if n is None:
        return None
    if n > 1.5:
        n = n / 100.0
    return n


def _pmd_txt(*parts):
    try:
        return _normalize_text(' '.join(str(p or '') for p in parts))
    except Exception:
        return ' '.join(str(p or '') for p in parts).lower()


def _pmd_item_desc(item):
    return str(item.get('descripcion') or item.get('insumo') or item.get('concepto') or item.get('codigo') or '').strip()


def _pmd_item_code(item):
    return str(item.get('codigo') or item.get('code') or '').strip()


def _pmd_item_provider_price(item):
    for k in ('precio_proveedor','precio_base','costo_proveedor','costo_contratista','precio','costo'):
        v = _pmd_num(item.get(k))
        if v is not None and v > 0:
            return v
    return None


def _pmd_item_market_price(item):
    for k in ('precio_mercado','precio_mercado_catalogo','costo_mercado'):
        v = _pmd_num(item.get(k))
        if v is not None and v > 0:
            return v
    return None


def _pmd_item_provider_amount(item):
    for k in ('importe_proveedor','importe_contratista','importe'):
        v = _pmd_num(item.get(k))
        if v is not None and v >= 0:
            return v
    price = _pmd_item_provider_price(item)
    qty = _pmd_num(item.get('cantidad') or item.get('factor') or item.get('volumen'))
    if price is not None and qty is not None:
        return price * qty
    return None


def _pmd_item_market_amount(item):
    v = _pmd_num(item.get('importe_mercado'))
    if v is not None and v >= 0:
        return v
    price = _pmd_item_market_price(item)
    qty = _pmd_num(item.get('cantidad') or item.get('factor') or item.get('volumen'))
    if price is not None and qty is not None:
        return price * qty
    return None


def _pmd_labor_role_label(text):
    t = _pmd_txt(text)
    if 'supervisor' in t and 'seguridad' in t:
        return 'Supervisor de Seguridad'
    if 'supervisor' in t and ('obra' in t or 'obras' in t):
        return 'Supervisor Obra'
    if 'soldador' in t or 'argonero' in t:
        return 'Oficial Soldador/argonero'
    if 'tubero' in t:
        return 'Oficial Tubero'
    if 'electricista' in t or 'electrico' in t or 'eléctrico' in t:
        return 'Oficial Eléctrico'
    if 'medio oficial' in t:
        return 'Medio Oficial'
    if 'ayudante' in t or 'peon' in t or 'peón' in t:
        return 'Ayudante General'
    if 'operador' in t:
        return 'Operador'
    if 'cuadrilla supervision' in t or 'cuadrilla de supervision' in t or 'cuadrilla de supervisión' in t:
        return 'Cuadrilla de Supervisión'
    if 'oficial' in t:
        return 'Oficial especializado'
    return None


def _pmd_collect_rules(dato, metrics=None, provider_analysis=None):
    concepts = _step37_concepts(dato or {}) if '_step37_concepts' in globals() else []
    total_provider = _pmd_num((metrics or {}).get('total_contratista'), 0.0) or 0.0
    if not total_provider:
        total_provider = sum(_step37_concept_provider_total(c) for _, c in concepts) if '_step37_concept_provider_total' in globals() else 0.0

    # Partidas azules / concentracion 80-85% por importe contratista.
    rows = []
    for clave, c in concepts:
        imp = _step37_concept_provider_total(c) if '_step37_concept_provider_total' in globals() else _pmd_num(c.get('total'), 0.0)
        rows.append({'servicio': clave, 'descripcion': c.get('desc') or c.get('descripcion') or '', 'importe': imp})
    rows = sorted(rows, key=lambda r: r.get('importe') or 0.0, reverse=True)
    acc = 0.0
    blue_rows = []
    for r in rows:
        acc += r.get('importe') or 0.0
        rr = dict(r)
        rr['acumulado_pct'] = (acc / total_provider) if total_provider else 0.0
        blue_rows.append(rr)
        if total_provider and rr['acumulado_pct'] >= PMD_PARETO_TARGET_PCT:
            break
    blue_pct = blue_rows[-1]['acumulado_pct'] if blue_rows else 0.0

    # Indirecto global declarado por contratista.
    indirect_pcts = []
    for _, c in concepts:
        pct = None
        try:
            if '_concept_indirect_amount_and_pct' in globals():
                _amt, pct = _concept_indirect_amount_and_pct(c)
        except Exception:
            pct = None
        if pct is None:
            pct = _pmd_pct(c.get('indirecto_pct'))
        if pct is not None and 0 <= pct <= 1.5:
            indirect_pcts.append(float(pct))
    contractor_indirect_pct = (sum(indirect_pcts) / len(indirect_pcts)) if indirect_pcts else None
    indirect_status = 'No determinado'
    if contractor_indirect_pct is not None:
        if contractor_indirect_pct > PMD_MARKET_INDIRECT_PCT:
            indirect_status = 'Arriba de mercado'
        elif contractor_indirect_pct < PMD_MARKET_INDIRECT_PCT:
            indirect_status = 'Bajo / competitivo'
        else:
            indirect_status = 'Alineado'

    labor_findings = []
    participation_findings = []
    equipment_findings = []
    material_findings = []
    no_reference_findings = []
    labor_provider_total = 0.0
    labor_market_total = 0.0
    equipment_provider_total = 0.0
    equipment_market_total = 0.0
    material_provider_total = 0.0
    material_market_total = 0.0

    for clave, c in concepts:
        for item in c.get('granular_market_items') or []:
            if not isinstance(item, dict):
                continue
            desc = _pmd_item_desc(item)
            code = _pmd_item_code(item)
            text = _pmd_txt(code, desc, item.get('tipo'), item.get('section'), item.get('domain'))
            dom = _step37_item_domain(item) if '_step37_item_domain' in globals() else ''
            prov_price = _pmd_item_provider_price(item)
            mkt_price = _pmd_item_market_price(item)
            prov_amt = _pmd_item_provider_amount(item) or 0.0
            mkt_amt = _pmd_item_market_amount(item) or 0.0
            status = str(item.get('match_status') or item.get('estado_match') or '').lower()
            if 'mano' in dom.lower():
                labor_provider_total += prov_amt
                labor_market_total += mkt_amt
                role = _pmd_labor_role_label(' '.join([code, desc]))
                if role and prov_price and mkt_price:
                    diff_pct = (prov_price - mkt_price) / mkt_price if mkt_price else None
                    if diff_pct is not None and diff_pct > 0.10:
                        labor_findings.append({
                            'servicio': clave,
                            'rol': role,
                            'codigo': code,
                            'contratista': prov_price,
                            'mercado': mkt_price,
                            'diff_pct': diff_pct,
                            'estado': 'alto' if diff_pct > 0.30 else ('revision' if diff_pct > 0.20 else 'observacion')
                        })
            elif 'equipo' in dom.lower() or 'herramienta' in dom.lower():
                equipment_provider_total += prov_amt
                equipment_market_total += mkt_amt
            else:
                material_provider_total += prov_amt
                material_market_total += mkt_amt

            # Herramienta menor / EPP como porcentaje sobre MO.
            qty = _pmd_pct(item.get('cantidad') or item.get('factor') or item.get('volumen'))
            if 'herramienta menor' in text or '%herr' in text or '%mo1' in text:
                pct = qty if qty is not None else None
                if pct is not None and pct > PMD_TOOL_MINOR_MARKET_PCT:
                    participation_findings.append({
                        'tipo': 'Herramienta Menor',
                        'servicio': clave,
                        'contratista_pct': pct,
                        'mercado_pct': PMD_TOOL_MINOR_MARKET_PCT,
                        'estado': 'elevado' if pct > 0.07 else 'revisar'
                    })
            if 'equipo de proteccion' in text or 'equipo de protección' in text or 'epp' in text or '%mo5' in text:
                pct = qty if qty is not None else None
                if pct is not None and pct > PMD_EPP_MARKET_PCT:
                    participation_findings.append({
                        'tipo': 'EPP',
                        'servicio': clave,
                        'contratista_pct': pct,
                        'mercado_pct': PMD_EPP_MARKET_PCT,
                        'estado': 'revisar'
                    })

            # Equipos criticos / rentas especializadas.
            equipment_terms = ['montacargas', 'grua', 'grúa', 'manlift', 'plataforma', 'andamio', 'tijera']
            if any(t in text for t in equipment_terms):
                diff_pct = (prov_price - mkt_price) / mkt_price if prov_price and mkt_price else None
                participation = (prov_amt / total_provider) if total_provider and prov_amt else 0.0
                if (diff_pct is not None and diff_pct > 0.20) or participation >= 0.05:
                    equipment_findings.append({
                        'servicio': clave,
                        'equipo': desc or code,
                        'contratista': prov_price,
                        'mercado': mkt_price,
                        'diff_pct': diff_pct,
                        'participacion': participation,
                        'estado': 'critico' if participation >= 0.30 or (diff_pct is not None and diff_pct > 1.0) else 'revisar'
                    })

            # Materiales con diferencia relevante.
            if 'material' in dom.lower() and prov_price and mkt_price:
                diff_pct = (prov_price - mkt_price) / mkt_price if mkt_price else None
                if diff_pct is not None and diff_pct > 0.25:
                    material_findings.append({
                        'servicio': clave,
                        'material': desc or code,
                        'contratista': prov_price,
                        'mercado': mkt_price,
                        'diff_pct': diff_pct,
                        'impacto': abs(prov_amt - mkt_amt)
                    })

            if 'sin_match' in status or 'no_match' in status or (not mkt_price and not mkt_amt):
                if prov_amt or prov_price:
                    no_reference_findings.append({
                        'servicio': clave,
                        'elemento': desc or code,
                        'tipo': dom,
                        'importe_contratista': prov_amt,
                        'estado': status or 'sin referencia'
                    })

    # Desduplicacion por rol, tomando el mayor diferencial.
    labor_by_role = {}
    for f in sorted(labor_findings, key=lambda x: x.get('diff_pct') or 0, reverse=True):
        labor_by_role.setdefault(f['rol'], f)
    labor_findings = list(labor_by_role.values())

    # Supervision 1:5: diagnostico textual si hay supervisor en hallazgos o importes.
    supervision_present = any(f['rol'] in ('Supervisor Obra','Supervisor de Seguridad','Cuadrilla de Supervisión') for f in labor_findings)
    mo_excess = labor_provider_total - labor_market_total
    total_diff = (metrics or {}).get('diferencia_total')
    mo_share_of_gap = (mo_excess / total_diff) if total_diff and total_diff > 0 else None

    return {
        'market_indirect_pct': PMD_MARKET_INDIRECT_PCT,
        'contractor_indirect_pct': contractor_indirect_pct,
        'indirect_status': indirect_status,
        'labor_market_year': PMD_LABOR_MARKET_YEAR,
        'labor_dof_adjustment_pct': PMD_LABOR_DOF_ADJUSTMENT_PCT,
        'blue_items_pct': blue_pct,
        'blue_items_count': len(blue_rows),
        'blue_items': blue_rows[:15],
        'labor_findings': labor_findings[:12],
        'participation_findings': participation_findings[:12],
        'equipment_findings': sorted(equipment_findings, key=lambda x: (x.get('participacion') or 0, x.get('diff_pct') or 0), reverse=True)[:12],
        'material_findings': sorted(material_findings, key=lambda x: x.get('impacto') or 0, reverse=True)[:12],
        'no_reference_findings': sorted(no_reference_findings, key=lambda x: x.get('importe_contratista') or 0, reverse=True)[:12],
        'supervision_present': supervision_present,
        'supervision_rule': 'Referencia PMD: 1 supervisor de obra + 1 supervisor de seguridad por cada 5 trabajadores.',
        'labor_provider_total': labor_provider_total,
        'labor_market_total': labor_market_total,
        'labor_excess': mo_excess,
        'labor_share_of_gap': mo_share_of_gap,
        'material_provider_total': material_provider_total,
        'material_market_total': material_market_total,
        'equipment_provider_total': equipment_provider_total,
        'equipment_market_total': equipment_market_total,
    }

try:
    _STEP41_COLLECT_EXEC_METRICS_PREV = _step37_collect_executive_metrics
except Exception:
    _STEP41_COLLECT_EXEC_METRICS_PREV = None


def _step37_collect_executive_metrics(dato, provider_analysis=None):
    metrics = _STEP41_COLLECT_EXEC_METRICS_PREV(dato, provider_analysis) if _STEP41_COLLECT_EXEC_METRICS_PREV else {}
    try:
        metrics['pmd_rules'] = _pmd_collect_rules(dato, metrics, provider_analysis)
    except Exception as exc:
        metrics['pmd_rules'] = {'error': f'{type(exc).__name__}: {str(exc)[:160]}'}
    return metrics

try:
    _STEP41_AI_DICTAMEN_PREV = _step37_ai_dictamen
except Exception:
    _STEP41_AI_DICTAMEN_PREV = None


def _step37_ai_dictamen(metrics):
    """Claude redacta con contexto PMD. No calcula ni modifica precios."""
    try:
        import quantia_ai_review as qair
        policy = qair.ai_runtime_policy()
        if not policy.get('expert_review_enabled') or not policy.get('anthropic_key_set'):
            return None, policy, 'Claude desactivado por política o sin API key'
        pmd = metrics.get('pmd_rules') or {}
        payload = {
            'rol': 'Eres un analista senior PMD de precios unitarios, APU/Neodata y revision de cotizaciones industriales.',
            'tarea': (
                'Redacta un dictamen profesional con estilo PMD. No inventes datos, no recalcules y no modifiques precios. '
                'Usa las reglas PMD enviadas: concentracion de partidas azules, indirecto vs 25%, MO por rol individual, '
                'actualizacion MO 2026/DOF 13%, supervision 1:5, herramienta menor 5%, EPP, equipos criticos y servicios sin referencia.'
            ),
            'reglas_pmd': {
                'mercado_indirecto': pmd.get('market_indirect_pct', PMD_MARKET_INDIRECT_PCT),
                'mo_year': pmd.get('labor_market_year', PMD_LABOR_MARKET_YEAR),
                'mo_dof_increment': pmd.get('labor_dof_adjustment_pct', PMD_LABOR_DOF_ADJUSTMENT_PCT),
                'herramienta_menor_ref': PMD_TOOL_MINOR_MARKET_PCT,
                'epp_ref': PMD_EPP_MARKET_PCT,
                'supervision': pmd.get('supervision_rule'),
            },
            'formato_json': {
                'conclusion': 'parrafo breve con lectura PMD',
                'riesgos_prioritarios': [{'servicio':'clave', 'riesgo':'texto PMD', 'impacto':'texto corto'}],
                'recomendaciones': ['accion concreta para analista PMD'],
                'lectura_ejecutiva': ['bullet PMD']
            },
            'datos': {
                'totales': {k: metrics.get(k) for k in ('total_servicios','total_contratista','total_mercado','diferencia_total','diferencia_pct_total','no_match','no_match_unit','total_items')},
                'top_diferenciales': (metrics.get('top_diferenciales') or [])[:8],
                'issues': (metrics.get('issues') or [])[:10],
                'domain_provider': metrics.get('domain_provider'),
                'domain_market': metrics.get('domain_market'),
                'pmd_rules': pmd,
            }
        }
        raw = qair._anthropic_message(payload, max_tokens=int(os.getenv('QUANTIA_AI_EXPERT_MAX_TOKENS','1500')), timeout=int(os.getenv('QUANTIA_AI_EXPERT_TIMEOUT','35')))
        parsed = None
        if raw:
            start, end = raw.find('{'), raw.rfind('}')
            if start >= 0 and end > start:
                try:
                    parsed = json.loads(raw[start:end+1])
                except Exception:
                    parsed = None
        return {'raw': raw, 'parsed': parsed}, policy, 'Claude ejecutivo PMD activo'
    except Exception as exc:
        return None, {}, f'Claude ejecutivo PMD error: {type(exc).__name__}: {str(exc)[:160]}'


def _pmd_fmt_money(v):
    return f'${v:,.2f}' if isinstance(v, (int, float)) else 'N/D'


def _pmd_fmt_pct(v):
    return f'{v*100:.2f}%' if isinstance(v, (int, float)) else 'N/D'

try:
    _STEP41_ADD_EXEC_SHEET_PREV = _step37_add_executive_ai_sheet
except Exception:
    _STEP41_ADD_EXEC_SHEET_PREV = None


def _step41_append_pmd_rules_sheet(output, dato, provider_analysis=None):
    """Agrega secciones PMD al tab Analisis experto IA sin tocar calculos."""
    try:
        wb = openpyxl.load_workbook(output)
        if 'Analisis experto IA' not in wb.sheetnames:
            return output
        ws = wb['Analisis experto IA']
        metrics = _step37_collect_executive_metrics(dato or {}, provider_analysis)
        pmd = metrics.get('pmd_rules') or {}
        row = ws.max_row + 2

        row = _step37_write_section_header(ws, row, '7. Reglas PMD aplicadas al análisis', cols=8, color='305496')
        rule_rows = []
        if pmd.get('blue_items_pct'):
            rule_rows.append(['Partidas principales', f"Las partidas principales representan {_pmd_fmt_pct(pmd.get('blue_items_pct'))} de la propuesta.", 'Priorizar revisión de partidas azules / mayor impacto.'])
        if pmd.get('contractor_indirect_pct') is not None:
            rule_rows.append(['Indirecto global', f"Contratista {_pmd_fmt_pct(pmd.get('contractor_indirect_pct'))} vs mercado {_pmd_fmt_pct(pmd.get('market_indirect_pct'))}.", pmd.get('indirect_status') or 'Revisar'])
        else:
            rule_rows.append(['Indirecto global', f"No se detectó indirecto explícito; referencia PMD de mercado {_pmd_fmt_pct(PMD_MARKET_INDIRECT_PCT)}.", 'Revisar si el contratista lo declara en su matriz.'])
        rule_rows.append(['Mano de obra mercado', f"Costos MO de mercado actualizados a {pmd.get('labor_market_year', PMD_LABOR_MARKET_YEAR)} considerando incremento {_pmd_fmt_pct(pmd.get('labor_dof_adjustment_pct', PMD_LABOR_DOF_ADJUSTMENT_PCT))} DOF.", 'Usar como base de comparación por rol individual.'])
        if pmd.get('supervision_present'):
            rule_rows.append(['Supervisión', pmd.get('supervision_rule'), 'Validar proporción de supervisión frente a cuadrilla ejecutora.'])
        rule_rows.append(['Herramienta menor', f"Referencia PMD {_pmd_fmt_pct(PMD_TOOL_MINOR_MARKET_PCT)}.", 'Alertar si contratista supera la referencia.'])
        rule_rows.append(['EPP', f"Referencia PMD {_pmd_fmt_pct(PMD_EPP_MARKET_PCT)}.", 'Alertar solo si contratista supera la referencia.'])
        row = _step37_write_table(ws, row, ['Regla', 'Lectura calculada', 'Acción PMD'], rule_rows, widths={'A':28,'B':70,'C':70}, color_header='305496') + 1

        # Hallazgos de mano de obra.
        if pmd.get('labor_findings'):
            row = _step37_write_section_header(ws, row, '8. Mano de obra: roles con costo por arriba de mercado', cols=8, color='9E480E')
            labor_rows = []
            for f in pmd.get('labor_findings')[:10]:
                labor_rows.append([f.get('rol'), _pmd_fmt_money(f.get('contratista')), _pmd_fmt_money(f.get('mercado')), _pmd_fmt_pct(f.get('diff_pct')), f.get('estado'), 'Solicitar apertura de tarifa, jornal y rendimiento; comparar contra rol individual, no cuadrilla genérica.'])
            row = _step37_write_table(ws, row, ['Rol MO', 'Contratista', 'Mercado', 'Dif %', 'Estado', 'Recomendación'], labor_rows, widths={'A':28,'B':16,'C':16,'D':12,'E':16,'F':70}, color_header='9E480E') + 1

        # Participacion %MO, herramienta, EPP.
        if pmd.get('participation_findings'):
            row = _step37_write_section_header(ws, row, '9. Porcentajes sobre MO: herramienta menor / EPP', cols=8, color='9E480E')
            part_rows = []
            for f in pmd.get('participation_findings')[:10]:
                part_rows.append([f.get('tipo'), f.get('servicio'), _pmd_fmt_pct(f.get('contratista_pct')), _pmd_fmt_pct(f.get('mercado_pct')), f.get('estado'), 'Validar si el porcentaje corresponde al alcance real y evitar duplicidad con equipo incluido.'])
            row = _step37_write_table(ws, row, ['Concepto', 'Servicio', 'Contratista', 'Mercado', 'Estado', 'Recomendación'], part_rows, widths={'A':24,'B':16,'C':14,'D':14,'E':16,'F':70}, color_header='9E480E') + 1

        # Equipos criticos.
        if pmd.get('equipment_findings'):
            row = _step37_write_section_header(ws, row, '10. Equipo crítico / renta especializada', cols=8, color='9E480E')
            eq_rows = []
            for f in pmd.get('equipment_findings')[:10]:
                eq_rows.append([f.get('servicio'), f.get('equipo'), _pmd_fmt_money(f.get('contratista')), _pmd_fmt_money(f.get('mercado')), _pmd_fmt_pct(f.get('participacion')), f.get('estado'), 'Solicitar cotización soporte y alcance incluido: renta, seguro, flete/porte, operador y consumibles.'])
            row = _step37_write_table(ws, row, ['Servicio', 'Equipo', 'Contratista', 'Mercado', '% propuesta', 'Estado', 'Recomendación'], eq_rows, widths={'A':16,'B':52,'C':16,'D':16,'E':14,'F':16,'G':70}, color_header='9E480E') + 1

        # Hallazgos sin referencia.
        if pmd.get('no_reference_findings'):
            row = _step37_write_section_header(ws, row, '11. Servicios / insumos sin referencia confiable', cols=8, color='595959')
            nr_rows = []
            for f in pmd.get('no_reference_findings')[:10]:
                nr_rows.append([f.get('servicio'), f.get('elemento'), f.get('tipo'), _pmd_fmt_money(f.get('importe_contratista')), 'No concluir caro/barato sin soporte; solicitar desglose o referencia externa.'])
            row = _step37_write_table(ws, row, ['Servicio', 'Elemento', 'Tipo', 'Importe contratista', 'Recomendación'], nr_rows, widths={'A':16,'B':62,'C':24,'D':18,'E':80}, color_header='595959') + 1

        # Ajustes visuales.
        for c in range(1, 9):
            ws.column_dimensions[get_column_letter(c)].width = max(ws.column_dimensions[get_column_letter(c)].width or 12, 14)
        wb.save(output)
    except Exception as exc:
        _q_perf_log('step41_pmd_append_error', error=f'{type(exc).__name__}: {exc}') if '_q_perf_log' in globals() else None
    return output


def _step37_add_executive_ai_sheet(output, dato, provider_analysis=None):
    result = _STEP41_ADD_EXEC_SHEET_PREV(output, dato, provider_analysis) if _STEP41_ADD_EXEC_SHEET_PREV else output
    return _step41_append_pmd_rules_sheet(result, dato, provider_analysis)

QUANTIA_EXPERT_REPORT_VERSION = 'step41_pmd_rules_expert_report'
