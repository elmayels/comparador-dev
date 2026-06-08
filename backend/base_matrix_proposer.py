"""
Generador de Matriz Propuesta IA para catalogos Nestle.

V0.2.6: flujo independiente de cotizaciones. Lee un catalogo de conceptos,
encuentra matrices Construdata candidatas, compone una matriz propuesta estilo
APU/Neodata y escribe un Excel con un tab visible de presupuesto totalizador,
un tab visible de matriz detalle y un tab oculto de trazabilidad.

La IA de esta version es un analizador local deterministico (auditable) basado en
reglas de alcance, tokens tecnicos y seleccion de matrices. El modulo queda listo
para conectar un LLM externo sin cambiar el contrato del endpoint.
"""
from __future__ import annotations

import math
import os
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


@dataclass
class ServiceItem:
    area: str
    part: str
    qty: float
    unit: str
    description: str
    brand: str = ""
    source_sheet: str = ""
    source_row: int = 0
    family: str = ""


@dataclass
class MatrixRow:
    matrix_code: str
    matrix_name: str
    matrix_desc: str
    matrix_unit: str
    official_pu: float
    insumo_code: str
    insumo_name: str
    insumo_desc: str
    insumo_unit: str
    tipo: str
    unit_cost: float
    quantity: float
    is_yield: bool
    raw_import: float
    order: int = 0


@dataclass
class MatrixCandidate:
    code: str
    name: str
    desc: str
    unit: str
    official_pu: float
    score: float
    decision: str
    role: str
    reason: str


STOPWORDS = {
    "de", "del", "la", "el", "los", "las", "en", "y", "o", "con", "para", "por", "a", "al",
    "incluye", "suministro", "instalacion", "instalación", "fabricacion", "fabricación", "montaje",
    "todo", "necesario", "correcta", "ejecucion", "ejecución", "trabajos", "obra", "materiales",
    "mano", "herramienta", "equipo", "limpieza", "acarreos", "fletes", "transporte", "ver", "plano",
    "marca", "seguridad", "nestle", "nestlé", "contratistas", "requerimientos"
}

FAMILY_HINTS = [
    ("trazo", ["trazo", "nivelacion", "nivelación", "topograf"]),
    ("muro_block", ["muro", "block"]),
    ("castillo", ["castillo"]),
    ("cadena", ["cadena", "dala"]),
    ("sardinel", ["sardinel"]),
    ("aplanado", ["aplanado"]),
    ("pintura", ["pintura", "recubrimiento", "elastomer"]),
    ("limpieza", ["limpieza"]),
    ("conduit", ["conduit", "tubo"]),
    ("cople", ["cople"]),
    ("condulet", ["condulet"]),
    ("cable_cobre", ["cable", "cobre"]),
    ("estructura_metalica", ["estructura", "metalic", "metálic", "acero", "ir", "ipr", "placa", "columna", "viga", "alfarda"]),
    ("demolicion", ["demolicion", "demolición", "desmantel", "retiro"]),
    ("barandal", ["barandal", "pasamanos"]),
    ("escalon", ["escalon", "escalón", "escalera"]),
    ("ancla", ["ancla", "hilti", "epox"]),
    ("tapial", ["tapial", "polietileno"]),
]


def _strip_accents(value: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", value or "") if not unicodedata.combining(ch))


def normalize_text(value: Any) -> str:
    text = _strip_accents(str(value or "")).lower()
    text = re.sub(r"[^a-z0-9#.%/\-\s\"]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokens(value: str) -> List[str]:
    out = []
    for t in re.findall(r"[a-z0-9#]+", normalize_text(value)):
        if len(t) < 2 or t in STOPWORDS:
            continue
        out.append(t)
    return out


def _as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            if math.isfinite(float(value)):
                return float(value)
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("$", "").replace(",", "").replace(" ", "")
    try:
        return float(text)
    except Exception:
        return None


def _eval_simple_formula(ws, formula: str, _seen: Optional[set] = None, _depth: int = 0) -> Optional[float]:
    """Evalua formulas simples del catalogo (=11.9+2.2, =C98).
    No intenta formulas multi-hoja o complejas para evitar ciclos de plantillas.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return _as_float(formula)
    if _depth > 10 or "!" in formula:
        return None
    _seen = _seen or set()
    expr = formula[1:].strip()
    def repl(m):
        ref = m.group(0)
        if ref in _seen:
            return "0"
        _seen.add(ref)
        val = ws[ref].value
        if isinstance(val, str) and val.startswith("="):
            got = _eval_simple_formula(ws, val, _seen, _depth + 1)
        else:
            got = _as_float(val)
        return str(got if got is not None else 0)
    expr = re.sub(r"\b[A-Z]{1,3}\d+\b", repl, expr)
    if not re.fullmatch(r"[0-9.\-+*/() ]+", expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return None


def infer_family(description: str, part: str = "") -> str:
    norm = normalize_text(f"{part} {description}")
    # Prioridades para evitar que palabras genericas como limpieza/equipo dominen el alcance.
    if "trazo" in norm and ("nivelacion" in norm or "topograf" in norm):
        return "trazo"
    if "ucrete" in norm:
        return "recubrimiento_especial"
    if "hilti" in norm or "cortafuego" in norm:
        return "sellado_especial"
    if "tapial" in norm or "polietileno" in norm:
        return "tapial"
    for family, hints in FAMILY_HINTS:
        if all(normalize_text(h) in norm for h in hints[:2]) if family in {"muro_block", "cable_cobre"} else any(normalize_text(h) in norm for h in hints):
            return family
    return "general"


def parse_nestle_catalog(path: str, only_sheet: Optional[str] = None, start_row: Optional[int] = None, end_row: Optional[int] = None) -> List[ServiceItem]:
    wb = load_workbook(path, data_only=False, read_only=False)
    items: List[ServiceItem] = []
    skip_names = {"portada", "general", "anexo a", "anexo b", "anexo c", "anexo d", "anexo e", "anexo f", "auto-evaluacion", "auto-evaluación"}
    for ws in wb.worksheets:
        if only_sheet and normalize_text(ws.title) != normalize_text(only_sheet):
            continue
        if normalize_text(ws.title) in skip_names:
            continue
        header_row = None
        cols = {}
        for r in range(1, min(ws.max_row, 120) + 1):
            row_vals = [normalize_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 15) + 1)]
            joined = " | ".join(row_vals)
            if "part" in joined and "cant" in joined and ("uni" in joined or "unidad" in joined) and ("desc" in joined or "d e s c" in joined):
                header_row = r
                for c, val in enumerate(row_vals, start=1):
                    if "part" in val and "partida" not in val:
                        cols["part"] = c
                    elif "cant" in val:
                        cols["qty"] = c
                    elif val in {"uni", "unidad"} or "uni" in val:
                        cols["unit"] = c
                    elif "desc" in val or "d e s c" in val:
                        cols["desc"] = c
                    elif "marca" in val:
                        cols["brand"] = c
                break
        if not header_row or not {"part", "qty", "unit", "desc"}.issubset(cols):
            continue
        r0 = max(header_row + 1, start_row or (header_row + 1))
        r1 = min(ws.max_row, end_row or ws.max_row)
        current_section = ""
        for r in range(r0, r1 + 1):
            part = ws.cell(r, cols["part"]).value
            desc = ws.cell(r, cols["desc"]).value
            qty_raw = ws.cell(r, cols["qty"]).value
            unit = ws.cell(r, cols["unit"]).value
            # Filas seccion sin partida
            if not part and desc and not qty_raw:
                d = str(desc).strip()
                if len(d) < 80:
                    current_section = d
                continue
            if not part or not desc:
                continue
            part_s = str(part).strip()
            if len(part_s) > 25 or normalize_text(part_s) in {"part", "partida"}:
                continue
            qty = _eval_simple_formula(ws, str(qty_raw)) if isinstance(qty_raw, str) and qty_raw.startswith("=") else _as_float(qty_raw)
            if qty is None:
                continue
            brand = ws.cell(r, cols.get("brand", 0)).value if cols.get("brand") else ""
            item = ServiceItem(
                area=ws.title,
                part=part_s,
                qty=qty,
                unit=str(unit or "").strip(),
                description=str(desc or "").strip(),
                brand=str(brand or "").strip(),
                source_sheet=ws.title,
                source_row=r,
                family=infer_family(str(desc or ""), part_s),
            )
            items.append(item)
    return items


class ConstrudataMatrixBase:
    def __init__(self, path: str):
        self.path = str(path)
        self.loaded = False
        self.matrices: Dict[str, Dict[str, Any]] = {}
        self.rows: Dict[str, List[MatrixRow]] = {}

    def _iter_xlsx_rows_fast(self, path: Path):
        """Lee sheet1.xml directamente. El export Construdata viene como XLSX crudo sin estilos complejos.
        Esto evita tiempos muy altos de openpyxl en archivos de 90k+ filas.
        """
        def col_idx(ref: str) -> int:
            letters = re.sub(r"[^A-Z]", "", ref.upper())
            n = 0
            for ch in letters:
                n = n * 26 + (ord(ch) - 64)
            return n
        with zipfile.ZipFile(str(path)) as z:
            names = set(z.namelist())
            shared = []
            if "xl/sharedStrings.xml" in names:
                for event, elem in ET.iterparse(z.open("xl/sharedStrings.xml"), events=("end",)):
                    if elem.tag.endswith("}si"):
                        txt = []
                        for node in elem.iter():
                            if node.tag.endswith("}t") and node.text:
                                txt.append(node.text)
                        shared.append("".join(txt)); elem.clear()
            sheet_name = "xl/worksheets/sheet1.xml" if "xl/worksheets/sheet1.xml" in names else next((n for n in names if n.startswith("xl/worksheets/sheet")), None)
            if not sheet_name:
                return
            for event, row in ET.iterparse(z.open(sheet_name), events=("end",)):
                if not row.tag.endswith("}row"):
                    continue
                vals = [None] * 24
                for c in row:
                    if not c.tag.endswith("}c"):
                        continue
                    ref = c.attrib.get("r", "A1")
                    idx = col_idx(ref) - 1
                    if idx >= 24:
                        continue
                    typ = c.attrib.get("t")
                    val = None
                    for child in c:
                        if child.tag.endswith("}v"):
                            val = child.text; break
                        if child.tag.endswith("}is"):
                            parts = []
                            for tnode in child.iter():
                                if tnode.tag.endswith("}t") and tnode.text:
                                    parts.append(tnode.text)
                            val = "".join(parts); typ = "str"; break
                    if typ == "s" and val is not None:
                        try: val = shared[int(val)]
                        except Exception: pass
                    vals[idx] = val
                row.clear()
                yield vals

    def load(self) -> None:
        if self.loaded:
            return
        p = Path(self.path)
        if not p.exists():
            raise FileNotFoundError(f"No se encontro construdata_matrices: {p}")
        for vals in self._iter_xlsx_rows_fast(p):
            if not vals or len(vals) < 24:
                continue
            code = vals[4]
            insumo = vals[13]
            if not code or not insumo:
                continue
            mcode = str(code).strip()
            row = MatrixRow(
                matrix_code=mcode,
                matrix_name=str(vals[5] or "").strip(),
                matrix_desc=str(vals[6] or "").strip(),
                matrix_unit=str(vals[7] or "").strip(),
                official_pu=_as_float(vals[10]) or 0.0,
                insumo_code=str(vals[13] or "").strip(),
                insumo_name=str(vals[14] or "").strip(),
                insumo_desc=str(vals[15] or "").strip(),
                insumo_unit=str(vals[16] or "").strip(),
                tipo=str(vals[18] or "").strip(),
                unit_cost=_as_float(vals[19]) or 0.0,
                quantity=_as_float(vals[20]) or 0.0,
                is_yield=str(vals[21]).strip().lower() in {"true", "1", "verdadero"},
                raw_import=_as_float(vals[23]) or 0.0,
                order=int(_as_float(vals[12]) or 0),
            )
            self.rows.setdefault(mcode, []).append(row)
            if mcode not in self.matrices:
                text = " ".join([row.matrix_name, row.matrix_desc])
                self.matrices[mcode] = {
                    "code": mcode,
                    "name": row.matrix_name,
                    "desc": row.matrix_desc,
                    "unit": row.matrix_unit,
                    "official_pu": row.official_pu,
                    "tokens": set(tokens(text)),
                    "text": normalize_text(text),
                }
            if len(self.matrices[mcode]["tokens"]) < 80:
                self.matrices[mcode]["tokens"].update(tokens(f"{row.insumo_name} {row.insumo_desc} {row.insumo_code}"))
        self.loaded = True

    def search(self, item: ServiceItem, max_candidates: int = 10) -> List[MatrixCandidate]:
        self.load()
        qtext = f"{item.part} {item.description} {item.brand} {item.family}"
        qtokens = set(tokens(qtext))
        family = item.family
        scored: List[Tuple[float, Dict[str, Any]]] = []
        for m in self.matrices.values():
            mt = m["tokens"]
            overlap = len(qtokens & mt)
            if overlap <= 0:
                continue
            score = overlap / max(6.0, len(qtokens) ** 0.65)
            text = m["text"]
            unit_bonus = 0.0
            if normalize_text(item.unit) and normalize_text(item.unit) == normalize_text(m.get("unit")):
                unit_bonus = 0.18
            # Bonos tecnicos por familia para seleccionar matrices mas parecidas.
            if family == "trazo" and "trazo" in text and "topogra" in text:
                score += 0.75
            elif family == "muro_block" and "muro" in text and "block" in text:
                score += 0.85
            elif family == "castillo" and "castillo" in text:
                score += 0.75
            elif family == "cadena" and ("cadena" in text or "dala" in text):
                score += 0.75
            elif family == "aplanado" and "aplanado" in text:
                score += 0.75
            elif family == "conduit" and "conduit" in text and ("grues" in text or "galv" in text):
                score += 0.65
            elif family == "cople" and "cople" in text:
                score += 0.65
            elif family == "condulet" and "condulet" in text:
                score += 0.80
            elif family == "cable_cobre" and "cable" in text and "cobre" in text:
                score += 0.85
            elif family == "estructura_metalica" and any(x in text for x in ["estructura metal", "perfil", "ipr", "placa", "acero"]):
                score += 0.55
            elif family == "demolicion" and any(x in text for x in ["demolicion", "demolic", "desmant"]):
                score += 0.50
            elif family == "limpieza" and "limpieza" in text:
                score += 0.45
            elif family == "barandal" and "barandal" in text:
                score += 0.55
            elif family == "escalon" and ("escal" in text or "escalon" in text):
                score += 0.45
            elif family == "pintura" and ("pintura" in text or "recubr" in text):
                score += 0.45
            score += unit_bonus
            scored.append((score, m))
        scored.sort(key=lambda x: x[0], reverse=True)
        out: List[MatrixCandidate] = []
        for idx, (score, m) in enumerate(scored[:max_candidates]):
            role = "Principal" if idx == 0 and score >= 0.55 else "Candidata"
            decision = "incluir" if idx == 0 and score >= 0.72 else ("revisar" if score >= 0.45 else "descartar")
            reason = _candidate_reason(item, m, score, decision, idx)
            out.append(MatrixCandidate(m["code"], m["name"], m["desc"], m["unit"], m["official_pu"], min(score, 1.0), decision, role, reason))
        return out


def _candidate_reason(item: ServiceItem, m: Dict[str, Any], score: float, decision: str, idx: int) -> str:
    if decision == "incluir":
        return "Matriz principal propuesta por coincidencia tecnica de descripcion, familia y unidad cuando aplica."
    if decision == "revisar":
        return "Candidato relacionado; requiere validacion de alcance/unidad antes de uso automatico."
    return "Candidato de baja similitud; se conserva solo para trazabilidad."


def _service_factor(item: ServiceItem, candidate: MatrixCandidate) -> float:
    desc = normalize_text(item.description)
    if normalize_text(item.unit) in {"pza", "pieza"} and normalize_text(candidate.unit) in {"m", "ml", "metro"}:
        # Tramos declarados de 3 m.
        if re.search(r"tramo\s+de\s+3\s*m", desc) or re.search(r"3\.00\s*m", desc):
            return 3.0
    return 1.0


def _additional_candidates(item: ServiceItem, base: ConstrudataMatrixBase, selected: List[MatrixCandidate], all_candidates: List[MatrixCandidate]) -> List[MatrixCandidate]:
    """Complementos simples y auditables: cople para conduit ambos extremos, castillo/cadena/aplanado para muro."""
    selected_codes = {c.code for c in selected}
    extra: List[MatrixCandidate] = []
    desc = normalize_text(item.description)
    if item.family == "conduit" and ("ambos extremos" in desc or "coples" in desc or "acoplamiento" in desc):
        fake = ServiceItem(item.area, item.part, item.qty, "pza", "cople conduit gruesa galvanizado 2 pulgadas", item.brand, item.source_sheet, item.source_row, "cople")
        for c in base.search(fake, 5):
            if c.code not in selected_codes and c.score >= 0.55:
                c.role = "Complementaria"
                c.decision = "incluir"
                c.reason = "Complementaria: el alcance menciona acoplamiento/coples en extremos."
                extra.append(c)
                break
    if item.family == "muro_block":
        for label, fam in [("castillo 15x15 concreto", "castillo"), ("cadena 15x20 concreto", "cadena"), ("aplanado fino mortero", "aplanado")]:
            fake = ServiceItem(item.area, item.part, item.qty, "m", label, item.brand, item.source_sheet, item.source_row, fam)
            for c in base.search(fake, 5):
                if c.code not in selected_codes and c.score >= 0.55:
                    c.role = "Complementaria"
                    c.decision = "incluir"
                    c.reason = f"Complementaria: el alcance de muro menciona {fam}/acabados."
                    extra.append(c); selected_codes.add(c.code)
                    break
    return extra


def choose_matrices(item: ServiceItem, base: ConstrudataMatrixBase, max_candidates: int = 10) -> Tuple[List[MatrixCandidate], List[MatrixCandidate], str, str]:
    candidates = base.search(item, max_candidates)
    selected: List[MatrixCandidate] = []
    if candidates and candidates[0].score >= 0.62:
        c0 = candidates[0]
        c0.decision = "incluir"
        c0.role = "Principal"
        selected.append(c0)
    selected.extend(_additional_candidates(item, base, selected, candidates))
    if not selected:
        status = "Sin matriz propuesta"
        obs = "No se encontro una matriz Construdata suficientemente confiable para proponer APU automatico."
    elif len(selected) == 1:
        status = "Matriz unica propuesta" if selected[0].score >= 0.72 else "Propuesta parcial"
        obs = selected[0].reason
    else:
        status = "Matriz compuesta propuesta"
        obs = "Se compone una sola matriz propuesta con matriz principal y complementarias explicitas del alcance."
    # Casos especializados: no forzar automatico aunque haya candidatos genericos.
    # No basta con que una descripcion mencione tapiales/seguridad de forma incidental;
    # solo se bloquea cuando el concepto principal sea especializado.
    low_auto_text = normalize_text(item.description)
    specialized = (
        item.family in {"tapial", "recubrimiento_especial", "sellado_especial"}
        or any(t in low_auto_text for t in ["ucrete", "hilti", "cortafuego", "axalta", "inoxidable"])
    )
    if specialized and item.family not in {"conduit", "condulet"}:
        if selected:
            candidates = selected + [c for c in candidates if c.code not in {s.code for s in selected}]
        selected = []
        status = "Requiere revision"
        obs = "Alcance/marca especializada; se muestran candidatos en trazabilidad, pero no se genera matriz automatica."
    return selected, candidates, status, obs


def calculate_matrix_rows(base: ConstrudataMatrixBase, item: ServiceItem, selected: List[MatrixCandidate]) -> Tuple[List[Dict[str, Any]], Dict[str, float]]:
    detail: List[Dict[str, Any]] = []
    totals = {"MATERIALES": 0.0, "MANO DE OBRA": 0.0, "EQUIPO Y HERRAMIENTA": 0.0, "BASICOS": 0.0, "OTROS": 0.0}
    for cand in selected:
        rows = base.rows.get(cand.code, [])
        factor = _service_factor(item, cand)
        if cand.role == "Complementaria" and "cople" in normalize_text(cand.name + " " + cand.desc) and ("ambos extremos" in normalize_text(item.description) or "coples" in normalize_text(item.description)):
            factor = 2.0
        # Subtotal MO para porcentuales calculado por cada matriz/factor.
        mo_subtotal = 0.0
        prelim = []
        for r in rows:
            if r.insumo_code.startswith("%MO") or normalize_text(r.insumo_unit) == "%":
                prelim.append((r, None))
                continue
            if r.is_yield and r.quantity:
                imp = (r.unit_cost / r.quantity) * factor
            else:
                imp = (r.unit_cost * r.quantity) * factor
            prelim.append((r, imp))
            if normalize_text(r.tipo) == normalize_text("MANO DE OBRA"):
                mo_subtotal += imp
        for r, imp in prelim:
            if imp is None:
                imp = mo_subtotal * (r.quantity or 0.0) * factor if r.quantity < 1 else mo_subtotal * (r.quantity / 100.0) * factor
            tipo = r.tipo or "OTROS"
            key = tipo if tipo in totals else "OTROS"
            totals[key] += imp
            detail.append({
                "part": item.part,
                "service": item.description,
                "unit_service": item.unit,
                "role": cand.role,
                "matrix_code": cand.code,
                "matrix_name": cand.name,
                "tipo": tipo,
                "insumo_code": r.insumo_code,
                "insumo": r.insumo_desc or r.insumo_name,
                "unit": r.insumo_unit,
                "qty": r.quantity,
                "factor": factor,
                "unit_cost": r.unit_cost,
                "importe": imp,
                "decision": "Incluir",
                "observacion": cand.reason,
            })
    return detail, totals


def _money(v: Any) -> float:
    try:
        return float(v or 0)
    except Exception:
        return 0.0


def write_matrix_proposal_excel(
    output_path: str,
    services: List[ServiceItem],
    base: ConstrudataMatrixBase,
    project_meta: Optional[Dict[str, str]] = None,
    indirect_pct: float = 0.25,
    max_services: Optional[int] = None,
) -> Dict[str, Any]:
    """Escribe el workbook de matriz propuesta.

    Salida homologada con los demas casos de uso:
      1. Resumen Profesional
      2. Comparativa
      3. Detalle
      4. Analisis experto IA
    La trazabilidad tecnica queda oculta por configuracion.
    """
    project_meta = project_meta or {}
    wb = Workbook()
    budget = wb.active
    budget.title = "Resumen Profesional"
    comp = wb.create_sheet("Comparativa")
    ws = wb.create_sheet("Detalle")
    ai = wb.create_sheet("Analisis experto IA")
    tr = wb.create_sheet("Trazabilidad Técnica")
    tr.sheet_state = "hidden" if os.getenv("EXCEL_HIDE_INTERNAL_SHEETS", "1").strip().lower() not in {"0", "false", "no"} else "visible"

    # Styles
    navy = "1F4E78"; blue = "D9EAF7"; dark = "243B53"; green = "E2F0D9"; yellow = "FFF2CC"; gray = "E7E6E6"; red = "F4CCCC"
    light_header = "D9EAF7"; total_fill = "D9EAD3"
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_font = Font(bold=True, size=14, color="FFFFFF")
    header_font = Font(bold=True, color="FFFFFF")
    small_bold = Font(bold=True, size=10)

    for sh in [budget, comp, ws, ai, tr]:
        sh.sheet_view.showGridLines = False

    # ------------------------------------------------------------------
    # Hoja 1: Presupuesto Propuesto (totalizador tipo catalogo)
    # ------------------------------------------------------------------
    budget.merge_cells("A1:Q1")
    budget["A1"] = "PRESUPUESTO PROPUESTO IA - Catálogo base del proyecto"
    budget["A1"].fill = PatternFill("solid", fgColor=navy)
    budget["A1"].font = title_font
    budget["A1"].alignment = Alignment(horizontal="center")
    budget["A2"] = "Proyecto"; budget["B2"] = project_meta.get("proyecto") or "Catálogo base"
    budget["D2"] = "Cliente"; budget["E2"] = project_meta.get("cliente") or "Nestlé"
    budget["G2"] = "Ubicación"; budget["H2"] = project_meta.get("ubicacion") or "—"
    budget["J2"] = "Indirectos"; budget["K2"] = indirect_pct; budget["K2"].number_format = "0.00%"
    budget["A3"] = "Fuente técnica"; budget["B3"] = "Construdata matrices + análisis IA local auditable"
    budget["D3"] = "Uso"; budget["E3"] = "Presupuesto base propuesto; revisar partidas con estado de revisión antes de cotizar."
    for cell in ["A2","D2","G2","J2","A3","D3"]:
        budget[cell].font = small_bold

    budget_headers = [
        "Área", "Part.", "Cant.", "Uni.", "Descripción", "Marca",
        "Materiales", "Mano de Obra", "Maquinaria/Equipo", "Básicos",
        "Costo Directo", "Indirectos 25%", "Financiamiento", "Utilidad", "Cargos",
        "PU Propuesto", "Total Propuesto", "Estado", "Matriz principal", "Complementarias", "Observación IA"
    ]
    budget_header_row = 5
    for j, h in enumerate(budget_headers, 1):
        cell = budget.cell(budget_header_row, j, h)
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ------------------------------------------------------------------
    # Hoja 2: Comparativa (resumen ejecutivo de propuesta por partida)
    # ------------------------------------------------------------------
    comp.merge_cells("A1:K1")
    comp["A1"] = "COMPARATIVA - PRESUPUESTO BASE IA"
    comp["A1"].fill = PatternFill("solid", fgColor=navy)
    comp["A1"].font = title_font
    comp["A1"].alignment = Alignment(horizontal="center")
    comp["A2"] = "Uso"; comp["B2"] = "Resumen por partida de la matriz/presupuesto propuesto. El desglose APU esta en Detalle."
    comp["A2"].font = small_bold
    comp.merge_cells("B2:K2")
    comp_headers = ["Part.", "Descripción", "Uni.", "Cant.", "PU Propuesto", "Total Propuesto", "Estado", "Matriz principal", "Complementarias", "Observación IA", "Acción sugerida"]
    comp_header_row = 4
    for j, h in enumerate(comp_headers, 1):
        cell = comp.cell(comp_header_row, j, h)
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    comp_row = comp_header_row + 1

    # ------------------------------------------------------------------
    # Hoja 3: Detalle (detalle estilo APU)
    # ------------------------------------------------------------------
    ws.merge_cells("A1:K1")
    ws["A1"] = "MATRIZ PROPUESTA IA - Detalle APU por servicio"
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A2"] = "Proyecto"; ws["B2"] = project_meta.get("proyecto") or "Catálogo base"
    ws["D2"] = "Cliente"; ws["E2"] = project_meta.get("cliente") or "Nestlé"
    ws["G2"] = "Indirectos"; ws["H2"] = indirect_pct
    ws["H2"].number_format = "0.00%"
    ws["A3"] = "Fuente"; ws["B3"] = "Construdata matrices + análisis IA local auditable"
    ws["D3"] = "Nota"; ws["E3"] = "La IA propone alcance/matrices; los costos se calculan desde insumos Construdata."

    trace_headers = ["Part.", "Servicio", "Unidad", "Cantidad", "Candidato", "Matriz", "Unidad CD", "Score", "Decision", "Rol", "Motivo"]
    tr.append(trace_headers)
    for cell in tr[1]:
        cell.fill = PatternFill("solid", fgColor=dark); cell.font = header_font; cell.border = border

    row = 5
    budget_row = budget_header_row + 1
    summary = {"services": 0, "with_matrix": 0, "requires_review": 0, "details": 0, "budget_total": 0.0}
    visible_services = services[:max_services] if max_services else services
    for item in visible_services:
        selected, candidates, status, obs = choose_matrices(item, base)
        detail_rows, totals = calculate_matrix_rows(base, item, selected)
        direct = sum(totals.values())
        indirect = direct * indirect_pct
        financing = 0.0
        utility = 0.0
        charges = 0.0
        pu = direct + indirect + financing + utility + charges
        total = pu * item.qty
        summary["services"] += 1
        if selected:
            summary["with_matrix"] += 1
        if "revision" in normalize_text(status) or not selected:
            summary["requires_review"] += 1
        summary["details"] += len(detail_rows)
        summary["budget_total"] += total

        # Budget/catalog row
        budget_values = [
            item.area, item.part, item.qty, item.unit, item.description, item.brand or "—",
            totals.get("MATERIALES", 0), totals.get("MANO DE OBRA", 0), totals.get("EQUIPO Y HERRAMIENTA", 0), totals.get("BASICOS", 0),
            direct, indirect, financing, utility, charges, pu, total, status,
            selected[0].code if selected else "—", ", ".join(c.code for c in selected[1:]) or "—", obs
        ]
        for j, v in enumerate(budget_values, 1):
            cell = budget.cell(budget_row, j, v)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(j in {5, 21}))
            if j in {7,8,9,10,11,12,13,14,15,16,17}:
                cell.number_format = '$#,##0.00'
            if j == 3:
                cell.number_format = '0.0000'
            if j == 18:
                nstatus = normalize_text(str(v))
                if "sin matriz" in nstatus or "revision" in nstatus:
                    cell.fill = PatternFill("solid", fgColor=yellow)
                elif "parcial" in nstatus:
                    cell.fill = PatternFill("solid", fgColor=light_header)
                else:
                    cell.fill = PatternFill("solid", fgColor=green)
                cell.font = Font(bold=True)
        budget_row += 1

        # Comparativa row
        comp_values = [
            item.part, item.description, item.unit, item.qty, pu, total, status,
            selected[0].code if selected else "—", ", ".join(c.code for c in selected[1:]) or "—", obs,
            "Validar alcance y partidas en revisión" if ("revision" in normalize_text(status) or not selected) else "Usar como referencia base de cotización"
        ]
        for j, v in enumerate(comp_values, 1):
            cell = comp.cell(comp_row, j, v)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(j in {2,10,11}))
            if j in {5,6}:
                cell.number_format = '$#,##0.00'
            if j == 4:
                cell.number_format = '0.0000'
            if j == 7:
                nstatus = normalize_text(str(v))
                if "sin matriz" in nstatus or "revision" in nstatus:
                    cell.fill = PatternFill("solid", fgColor=yellow)
                elif "parcial" in nstatus:
                    cell.fill = PatternFill("solid", fgColor=light_header)
                else:
                    cell.fill = PatternFill("solid", fgColor=green)
                cell.font = Font(bold=True)
        comp_row += 1

        # Title section in detail sheet
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
        c = ws.cell(row, 1, f"PARTIDA {item.part} | {item.description[:180]}")
        c.fill = PatternFill("solid", fgColor=dark); c.font = Font(bold=True, color="FFFFFF", size=11)
        row += 1
        meta = [
            ("Área", item.area), ("Unidad", item.unit), ("Cantidad", item.qty), ("Marca", item.brand or "—"),
            ("Estado", status), ("Matriz principal", selected[0].code if selected else "—"), ("Complementarias", ", ".join(c.code for c in selected[1:]) or "—"),
        ]
        for i, (k, v) in enumerate(meta):
            col = 1 + i * 2
            if col + 1 > 11: break
            ws.cell(row, col, k).font = small_bold
            ws.cell(row, col + 1, v)
        row += 1
        ws.cell(row, 1, "Observación IA").font = small_bold
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=11)
        ws.cell(row, 2, obs).alignment = Alignment(wrap_text=True, vertical="top")
        row += 2

        # Economic summary
        econ_headers = ["Materiales", "Mano de Obra", "Maquinaria/Equipo", "Básicos", "Costo Directo", "Indirectos 25%", "Financiamiento", "Utilidad", "Cargos", "PU Propuesto", "Total"]
        for j, h in enumerate(econ_headers, 1):
            cell = ws.cell(row, j, h); cell.fill = PatternFill("solid", fgColor=navy); cell.font = header_font; cell.border = border; cell.alignment = Alignment(horizontal="center")
        row += 1
        econ_vals = [totals.get("MATERIALES", 0), totals.get("MANO DE OBRA", 0), totals.get("EQUIPO Y HERRAMIENTA", 0), totals.get("BASICOS", 0), direct, indirect, financing, utility, charges, pu, total]
        for j, v in enumerate(econ_vals, 1):
            cell = ws.cell(row, j, v); cell.number_format = '$#,##0.00'; cell.border = border
            if j in {5, 10, 11}: cell.font = Font(bold=True)
        row += 2

        # Coverage small block
        ws.cell(row, 1, "Cobertura de alcance").fill = PatternFill("solid", fgColor=green if selected else yellow)
        ws.cell(row, 1).font = small_bold
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=11)
        if selected:
            coverage = "Cubierto parcialmente por matrices Construdata seleccionadas. Validar elementos especializados, marcas, acabados, pruebas y cantidades no expresadas."
        else:
            coverage = "Sin matriz automática; revisar candidatos en trazabilidad y/o crear matriz manual."
        ws.cell(row, 2, coverage).alignment = Alignment(wrap_text=True)
        row += 2

        # Detail table
        headers = ["Bloque", "Código", "Descripción", "Unidad", "Cantidad/Rend.", "Factor", "P.U.", "Importe", "Fuente matriz", "Rol", "Observación IA"]
        for j, h in enumerate(headers, 1):
            cell = ws.cell(row, j, h); cell.fill = PatternFill("solid", fgColor=dark); cell.font = header_font; cell.border = border
        row += 1
        if detail_rows:
            for d in detail_rows:
                vals = [d["tipo"], d["insumo_code"], d["insumo"], d["unit"], d["qty"], d["factor"], d["unit_cost"], d["importe"], f'{d["matrix_code"]} - {d["matrix_name"]}', d["role"], d["observacion"]]
                for j, v in enumerate(vals, 1):
                    cell = ws.cell(row, j, v); cell.border = border; cell.alignment = Alignment(vertical="top", wrap_text=(j in {3,9,11}))
                    if j in {7,8}: cell.number_format = '$#,##0.00'
                    if j in {5,6}: cell.number_format = '0.0000'
                    if j == 1:
                        fill = green if "MATERIAL" in str(v).upper() else (blue if "MANO" in str(v).upper() else gray)
                        cell.fill = PatternFill("solid", fgColor=fill)
                row += 1
        else:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
            ws.cell(row, 1, "No se genera matriz automática para esta partida. Ver trazabilidad.").fill = PatternFill("solid", fgColor=yellow)
            row += 1
        row += 2

        # Trace rows
        for cnd in candidates:
            tr.append([item.part, item.description, item.unit, item.qty, cnd.code, cnd.name, cnd.unit, cnd.score, cnd.decision, cnd.role, cnd.reason])

    # Totalizador del presupuesto propuesto
    total_row = budget_row
    budget.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=10)
    budget.cell(total_row, 1, "TOTAL PRESUPUESTO PROPUESTO IA").fill = PatternFill("solid", fgColor=total_fill)
    budget.cell(total_row, 1).font = Font(bold=True, size=11)
    for cidx in range(1, len(budget_headers)+1):
        budget.cell(total_row, cidx).border = border
        if cidx in {11,12,13,14,15,17}:
            col = get_column_letter(cidx)
            budget.cell(total_row, cidx, f"=SUM({col}{budget_header_row+1}:{col}{budget_row-1})")
            budget.cell(total_row, cidx).number_format = '$#,##0.00'
            budget.cell(total_row, cidx).font = Font(bold=True)
            budget.cell(total_row, cidx).fill = PatternFill("solid", fgColor=total_fill)
        elif cidx == 16:
            budget.cell(total_row, cidx, "—")
            budget.cell(total_row, cidx).font = Font(bold=True)
            budget.cell(total_row, cidx).fill = PatternFill("solid", fgColor=total_fill)
    budget.cell(total_row, 18, f"{summary['with_matrix']} con matriz / {summary['requires_review']} revisión")
    budget.cell(total_row, 18).fill = PatternFill("solid", fgColor=total_fill)
    budget.cell(total_row, 18).font = Font(bold=True)

    # Hoja 4: Analisis experto IA
    ai.merge_cells("A1:F1")
    ai["A1"] = "ANALISIS EXPERTO IA - PRESUPUESTO BASE PROPUESTO"
    ai["A1"].fill = PatternFill("solid", fgColor=navy)
    ai["A1"].font = title_font
    ai["A1"].alignment = Alignment(horizontal="center")
    ai.merge_cells("A2:F2")
    ai["A2"] = "Lectura generada desde el catalogo base, matrices candidatas Construdata y calculo deterministico de costos."
    ai["A2"].fill = PatternFill("solid", fgColor=blue)
    ai["A2"].alignment = Alignment(wrap_text=True)
    ai_headers = ["Indicador", "Valor", "Lectura", "Riesgo", "Acción sugerida", "Fuente"]
    for j, h in enumerate(ai_headers, 1):
        cell = ai.cell(4, j, h)
        cell.fill = PatternFill("solid", fgColor=dark); cell.font = header_font; cell.border = border
    ai_rows = [
        ["Servicios detectados", summary.get("services"), "Conceptos procesados desde el catalogo base.", "Bajo", "Validar alcance técnico", "Catálogo base"],
        ["Con matriz", summary.get("with_matrix"), "Partidas con matriz Construdata principal asignada.", "Bajo" if summary.get("with_matrix") else "Alto", "Revisar partidas sin matriz", "Construdata matrices"],
        ["Requieren revisión", summary.get("requires_review"), "Partidas con cobertura parcial, sin matriz o con observación técnica.", "Medio" if summary.get("requires_review") else "Bajo", "Revisar Comparativa y Detalle", "Motor IA local"],
        ["Total presupuesto", summary.get("budget_total"), "Total propuesto antes de cotización de contratistas.", "Informativo", "Usar como base de comparación", "Resumen Profesional"],
    ]
    for rr, vals in enumerate(ai_rows, 5):
        for j, v in enumerate(vals, 1):
            cell = ai.cell(rr, j, v); cell.border = border; cell.alignment = Alignment(wrap_text=True, vertical="top")
            if j == 2 and isinstance(v, (int, float)) and vals[0] == "Total presupuesto":
                cell.number_format = '$#,##0.00'

    # Widths and formatting
    budget_widths = [18, 14, 12, 10, 55, 18, 14, 14, 16, 12, 15, 15, 14, 12, 12, 15, 16, 24, 18, 22, 60]
    for i, w in enumerate(budget_widths, 1):
        budget.column_dimensions[get_column_letter(i)].width = w
    comp_widths = [14, 60, 10, 12, 16, 18, 24, 18, 24, 60, 42]
    for i, w in enumerate(comp_widths, 1):
        comp.column_dimensions[get_column_letter(i)].width = w
    widths = [14, 18, 55, 12, 14, 12, 14, 14, 42, 16, 55]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for i, w in enumerate([22, 18, 58, 16, 18, 60], 1):
        ai.column_dimensions[get_column_letter(i)].width = w
    for i, w in enumerate([14, 55, 12, 12, 16, 50, 12, 10, 14, 16, 60], 1):
        tr.column_dimensions[get_column_letter(i)].width = w
    for sh in [budget, comp, ws, ai]:
        for r in range(1, sh.max_row + 1):
            sh.row_dimensions[r].height = 20
    budget.freeze_panes = "A6"
    comp.freeze_panes = "A5"
    ws.freeze_panes = "A5"
    tr.freeze_panes = "A2"
    budget.auto_filter.ref = f"A{budget_header_row}:{get_column_letter(len(budget_headers))}{max(budget_row-1, budget_header_row)}"
    comp.auto_filter.ref = f"A{comp_header_row}:K{max(comp_row-1, comp_header_row)}"
    wb._sheets = [budget, comp, ws, ai, tr]
    wb.save(output_path)
    return summary

def generate_matrix_proposal_workbook(
    catalog_path: str,
    output_path: str,
    construdata_matrices_path: str,
    project_meta: Optional[Dict[str, str]] = None,
    indirect_pct: float = 0.25,
    only_sheet: Optional[str] = None,
    start_row: Optional[int] = None,
    end_row: Optional[int] = None,
) -> Dict[str, Any]:
    services = parse_nestle_catalog(catalog_path, only_sheet=only_sheet, start_row=start_row, end_row=end_row)
    if not services:
        raise ValueError("No se detectaron conceptos con Part., Cant., Uni. y Descripción en el catálogo cargado.")
    base = ConstrudataMatrixBase(construdata_matrices_path)
    base.load()
    summary = write_matrix_proposal_excel(output_path, services, base, project_meta=project_meta, indirect_pct=indirect_pct)
    summary.update({"catalog_services": len(services), "output_path": output_path})
    return summary
