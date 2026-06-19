from __future__ import annotations

import re
import time
import unicodedata
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


def _txt(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _num(v: Any) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace("%", "")
    s = s.replace(" ", "")
    if s.count(",") and s.count("."):
        # Colombian/Latam human format: 1.234.567,89 or 1,234,567.89.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z0-9%]+", " ", s).lower().strip()
    return re.sub(r"\s+", " ", s)



def canonical_key(code: str = "", description: str = "") -> str:
    code = str(code or "").strip().lower()
    if code:
        return f"code:{code}"
    desc = str(description or "").strip().lower()
    return "desc:" + " ".join(desc.split())[:120]


def _tokens(s: str) -> set[str]:
    stop = {"de", "del", "la", "el", "los", "las", "y", "en", "con", "para", "por", "a", "un", "una", "m", "ml", "kg", "pza", "und"}
    return {t for t in _norm(s).split() if len(t) > 2 and t not in stop}


@dataclass
class CanonicalConcept:
    code: str = ""
    description: str = ""
    unit: str = ""
    quantity: float | None = None
    unit_price: float | None = None
    amount: float | None = None
    market_unit_price: float | None = None
    market_amount: float | None = None
    market_source: str = ""
    market_state: str = ""
    family: str = ""
    hierarchy_level: int = 0
    original_order: int = 0
    is_executable: bool = False
    weight_pct: float | None = None
    cumulative_pct: float | None = None
    is_pareto_80: bool = False
    source_row: int | None = None


@dataclass
class CanonicalApuItem:
    code: str = ""
    concept_key: str = ""
    description: str = ""
    unit: str = ""
    section: str = ""
    unit_price: float | None = None
    operator: str = "*"
    quantity: float | None = None
    amount: float | None = None
    percent: float | None = None
    market_unit_price: float | None = None
    market_operator: str = "*"
    market_quantity: float | None = None
    market_amount: float | None = None
    market_deviation: float | None = None
    state: str = ""
    observation: str = ""
    source_row: int | None = None


@dataclass
class CanonicalProvider:
    name: str
    concepts_file: str = ""
    matrix_file: str = ""
    concepts: list[CanonicalConcept] = field(default_factory=list)
    apu_items: list[CanonicalApuItem] = field(default_factory=list)
    # Provider-scoped 80/20. These keys represent the concepts that explain
    # the first 80% of this provider's own catalog amount (P.U. total *
    # quantity). They are intentionally not stored on the shared catalog spine
    # because every provider can have a different Pareto set.
    pareto_concept_keys: list[str] = field(default_factory=list)
    validations: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class CanonicalRun:
    run_id: str
    kind: str
    project_name: str
    providers: list[CanonicalProvider] = field(default_factory=list)
    base_concepts: list[CanonicalConcept] = field(default_factory=list)
    base_apu_items: list[CanonicalApuItem] = field(default_factory=list)
    validations: list[dict[str, Any]] = field(default_factory=list)

    def as_dict(self):
        return asdict(self)


class ReferenceCatalog:
    def __init__(self, data_dir: Path):
        self.items: list[dict[str, Any]] = []
        self.index: dict[str, list[dict[str, Any]]] = {}
        self.load(data_dir)

    def load(self, data_dir: Path):
        for path in data_dir.glob("*.xlsx"):
            lname = path.name.lower()
            if "matrices" in lname:
                continue
            kind = "MATERIAL" if "material" in lname else "MO" if "mano" in lname or "obra" in lname else "MAQUINARIA" if "maquinaria" in lname else "REF"
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                for ws in wb.worksheets[:3]:
                    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 3000), values_only=True):
                        vals = list(row)
                        texts = [_txt(v) for v in vals if _txt(v)]
                        nums = [_num(v) for v in vals if _num(v) is not None]
                        desc = max(texts, key=len) if texts else ""
                        # prefer plausible unit price: positive, not tiny, not row index-like
                        price = next((n for n in reversed(nums) if n and n > 1), None)
                        if desc and price:
                            item = {"description": desc, "price": price, "kind": kind, "source": path.name, "tokens": _tokens(desc)}
                            self.items.append(item)
                            for tok in item["tokens"]:
                                self.index.setdefault(tok, []).append(item)
                wb.close()
            except Exception:
                continue

    def match(self, description: str, section: str = "") -> dict[str, Any] | None:
        q = _tokens(description)
        if not q:
            return None
        wanted = ""
        sn = _norm(section)
        if "material" in sn:
            wanted = "MATERIAL"
        elif "mano" in sn or sn == "mo":
            wanted = "MO"
        elif "maquinaria" in sn or "equipo" in sn:
            wanted = "MAQUINARIA"
        best = None
        best_score = 0.0
        candidates: dict[int, dict[str, Any]] = {}
        for tok in q:
            for item in self.index.get(tok, []):
                candidates[id(item)] = item
        # Fallback to full scan only for very sparse/rare token queries.
        iterable = candidates.values() if candidates else self.items
        for item in iterable:
            if wanted and item["kind"] != wanted:
                continue
            it = item["tokens"]
            if not it:
                continue
            inter = len(q & it)
            score = inter / max(len(q), 1)
            if score > best_score:
                best_score = score
                best = item
        if best and best_score >= 0.35:
            return {**best, "confidence": round(best_score, 2)}
        return None


def _sheet_score(ws) -> int:
    return min(ws.max_row or 0, 5000) * min(ws.max_column or 0, 50)


def _best_sheet(path: Path, preferred_names: Iterable[str] | None = None):
    wb = load_workbook(path, read_only=True, data_only=True)
    if preferred_names:
        wanted = {_norm(x) for x in preferred_names}
        for ws in wb.worksheets:
            if _norm(ws.title) in wanted:
                return wb, ws
    ws = max(wb.worksheets, key=_sheet_score)
    return wb, ws


def _hierarchy_level(code: str, executable: bool = False) -> int:
    c = _txt(code)
    if not c:
        return 0
    # 1 -> level 1, 1.1 -> level 2, 1.1.1 -> level 3.
    parts = [p for p in re.split(r"[.]", c) if p != ""]
    return max(1, len(parts)) if parts else (3 if executable else 1)


def _apply_pareto_80(concepts: list[CanonicalConcept]) -> None:
    execs = [c for c in concepts if c.is_executable and (c.amount or 0) > 0]
    total = sum(float(c.amount or 0) for c in execs)
    if total <= 0:
        return
    ordered = sorted(execs, key=lambda c: float(c.amount or 0), reverse=True)
    cumulative = 0.0
    for c in ordered:
        c.weight_pct = float(c.amount or 0) / total
        cumulative += c.weight_pct
        c.cumulative_pct = cumulative
        # Include the concept that crosses the 80% threshold. This is the
        # professional 80/20 selection for negotiation priority.
        if cumulative <= 0.80 or (cumulative - (c.weight_pct or 0)) < 0.80:
            c.is_pareto_80 = True



APU_SECTION_TERMS = {
    "materiales", "mano de obra", "equipo y herramienta", "equipo", "maquinaria",
    "basicos", "básicos", "seccion financiera", "sección financiera",
    "subtotal materiales", "subtotal mano de obra", "subtotal maquinaria",
    "subtotal equipo", "subtotal basicos", "subtotal básicos", "costo directo",
    "costo indirecto", "indirectos", "utilidad", "financiamiento",
    "precio unitario", "total costo unitario", "total por servicio",
}


def _looks_like_apu_matrix_row(code: str, desc: str, unit: str = "", op: str = "") -> bool:
    """Guardrail for Comparativa: catalog parsing must not leak PU/matrix rows.

    Comparativa compares catalog concepts and total P.U. by provider. It must
    never show material/labor/equipment detail rows from the matrix/APU. Those
    rows belong exclusively to Detalle - <Proveedor>.
    """
    ndesc = _norm(desc)
    ncode = _norm(code)
    nunit = _norm(unit)
    nop = _norm(op)
    if not ndesc:
        return False
    if ndesc in APU_SECTION_TERMS or any(term == ndesc for term in APU_SECTION_TERMS):
        return True
    if any(term in ndesc for term in [
        "subtotal materiales", "subtotal mano", "subtotal maquinaria", "subtotal equipo",
        "costo directo", "costo indirecto", "precio unitario", "total costo",
        "seccion financiera", "sección financiera"
    ]):
        return True
    # Common PU detail signals: percentage rows, operators, and section-only rows.
    if ncode.startswith("%") or "%" in code or nop in {"*", "/", "%"} and not re.match(r"^\d+(\.\d+)*$", str(code).strip()):
        if nunit in {"%", "jorn", "hr", "hora", "kg", "m3", "m2", "m", "pza", "l", "lt"} or ncode:
            return True
    if ndesc in {"materiales", "mano de obra", "maquinaria", "equipo", "basicos", "básicos"}:
        return True
    return False


def _detect_header(rows: list[list[Any]], synonyms: dict[str, list[str]]) -> tuple[int, dict[str, int]]:
    best_idx = 0
    best_map: dict[str, int] = {}
    best_score = -1
    for idx, row in enumerate(rows[:25]):
        cells = [_norm(_txt(v)) for v in row]
        colmap: dict[str, int] = {}
        score = 0
        for field, words in synonyms.items():
            for c, text in enumerate(cells):
                if any(w in text for w in words):
                    colmap[field] = c
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_idx = idx
            best_map = colmap
    return best_idx, best_map


def _header_norms(row: list[Any]) -> list[str]:
    return [_norm(_txt(v)) for v in row]


def _detect_repeated_header_blocks(row: list[Any]) -> dict[str, int]:
    """Detect PMD-style provider + market blocks in a single header row.

    Standard manual PMD matrices usually repeat these headers:
    Código | Concepto | Unidad | P. Unitario | Op. | Cantidad | Importe | % | | P. Unitario | Op. | Cantidad | Importe

    The second repeated price/operator/quantity/amount block is the market
    block even when the word "Mercado" is not present in the same row.
    """
    cells = _header_norms(row)
    colmap: dict[str, int] = {}
    def positions(options: set[str] | tuple[str, ...]) -> list[int]:
        out = []
        for i, text in enumerate(cells):
            if any(opt == text or opt in text for opt in options):
                out.append(i)
        return out
    # provider columns
    for i, text in enumerate(cells):
        if text in {"codigo", "clave"}:
            colmap.setdefault("code", i)
        elif "concepto" in text or "descripcion" in text:
            colmap.setdefault("description", i)
        elif text in {"unidad", "und", "udm"}:
            colmap.setdefault("unit", i)
        elif text in {"op", "operador"}:
            colmap.setdefault("operator", i)
        elif "cantidad" in text or text == "cant":
            colmap.setdefault("quantity", i)
        elif "importe" in text or text in {"total", "valor"}:
            colmap.setdefault("amount", i)
        elif text == "%" or "porcentaje" in text:
            colmap.setdefault("percent", i)
        elif text in {"p unitario", "p u", "precio unitario", "pu"}:
            colmap.setdefault("unit_price", i)
    pu_cols = [i for i, t in enumerate(cells) if t in {"p unitario", "p u", "precio unitario", "pu"}]
    op_cols = [i for i, t in enumerate(cells) if t in {"op", "operador"}]
    qty_cols = [i for i, t in enumerate(cells) if "cantidad" in t or t == "cant"]
    amt_cols = [i for i, t in enumerate(cells) if "importe" in t or t in {"total", "valor"}]
    if len(pu_cols) >= 2:
        colmap["market_unit_price"] = pu_cols[1]
    if len(op_cols) >= 2:
        colmap["market_operator"] = op_cols[1]
    if len(qty_cols) >= 2:
        colmap["market_quantity"] = qty_cols[1]
    if len(amt_cols) >= 2:
        colmap["market_amount"] = amt_cols[1]
    return colmap


def _detect_matrix_header(rows: list[list[Any]]) -> tuple[int, dict[str, int]]:
    """Detect matrix/APU columns by semantic headers, including PMD variants."""
    # Variant A: standard PMD repeated block in row 2 or nearby.
    best_idx, best_map, best_score = 0, {}, -1
    for idx, row in enumerate(rows[:40]):
        m = _detect_repeated_header_blocks(row)
        score = len([k for k in ["code", "description", "unit", "unit_price", "quantity", "amount"] if k in m]) + 2 * len([k for k in ["market_unit_price", "market_amount"] if k in m])
        if score > best_score:
            best_idx, best_map, best_score = idx, m, score
    if best_score >= 5:
        # Variant B: wide PMD/TAPIAL layout. It often has headers like:
        # Clave / Descripción / Unidad / Cantidad / P. Unitario / Importe / Cantidad / P. Unitario / Importe
        row = rows[best_idx]
        cells = _header_norms(row)
        if len(row) >= 24 and any("clave" == t for t in cells) and any("descripcion" in t for t in cells):
            # Prefer detected positions but override obvious wide-layout market group.
            # 0-based columns observed in TAPIAL: code C, desc F, unit J, qty N, PU S, amount W, market qty Y, market PU Z, market amount AA.
            best_map.update({
                "code": 2, "description": 5, "unit": 9,
                "quantity": 13, "unit_price": 18, "amount": 22,
                "market_quantity": 24, "market_unit_price": 25, "market_amount": 26,
            })
        return best_idx, best_map
    return _detect_header(rows, MATRIX_SYNONYMS)


def _detect_concept_market_columns(header_row: list[Any]) -> dict[str, int]:
    cells = _header_norms(header_row)
    pu_cols = [i for i,t in enumerate(cells) if t in {"p u", "p unitario", "precio unitario", "pu"}]
    amt_cols = [i for i,t in enumerate(cells) if "importe" in t or t in {"subtotal", "total", "valor"}]
    out = {}
    if len(pu_cols) >= 2:
        out["market_unit_price"] = pu_cols[1]
    if len(amt_cols) >= 2:
        out["market_amount"] = amt_cols[1]
    return out


CONCEPT_SYNONYMS = {
    "code": ["codigo", "clave", "partida", "item"],
    "description": ["descripcion", "concepto", "servicio", "actividad"],
    "unit": ["unidad", "und", "udm"],
    "quantity": ["cantidad", "cant", "volumen"],
    "unit_price": ["p u", "precio unitario", "pu", "p unitario"],
    "amount": ["importe", "total", "valor", "monto"],
}

MATRIX_SYNONYMS = {
    "code": ["codigo", "clave", "insumo", "partida"],
    "description": ["descripcion", "concepto", "insumo", "recurso"],
    "unit": ["unidad", "und", "udm"],
    "unit_price": ["p unitario", "precio unitario", "p u", "pu"],
    "operator": ["op", "operador"],
    "quantity": ["cantidad", "rendimiento", "cant"],
    "amount": ["importe", "total", "valor"],
    "percent": ["%", "porcentaje"],
    "market_unit_price": ["mercado p unitario", "mercado precio unitario", "mercado p u", "mercado pu", "referencia p unitario"],
    "market_operator": ["mercado op", "referencia op"],
    "market_quantity": ["mercado cantidad", "mercado cant", "referencia cantidad"],
    "market_amount": ["mercado importe", "mercado total", "referencia importe"],
}


def parse_concepts(path: Path) -> list[CanonicalConcept]:
    """Parse a contractor/base concept catalog preserving original order and hierarchy.

    Preferred source shape is the CATALOGO sheet used by PU workbooks:
    CODIGO | DESCRIPCION | UNIDAD | CANTIDAD | P U | SUBTOTAL.

    The output intentionally keeps non-executable hierarchy rows so the
    Comparativa can show the catalog exactly as declared, while Pareto 80/20
    is calculated only over executable rows with amount.
    """
    wb, ws = _best_sheet(path, preferred_names=["CATALOGO", "CATÁLOGO", "Catalogo", "CATALOGO DE CONCEPTOS", "Catálogo de conceptos", "Comparativa"])
    try:
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 1500), values_only=True)]
        hidx, cmap = _detect_header(rows, CONCEPT_SYNONYMS)
        # If we have the canonical CATALOGO layout but fuzzy detection missed it,
        # scan explicit header names.
        if len(cmap) < 3:
            for idx, row in enumerate(rows[:40]):
                normed = [_norm(_txt(v)) for v in row]
                if any("codigo" == x for x in normed) and any("descripcion" in x for x in normed):
                    hidx = idx
                    cmap = {}
                    for c, text in enumerate(normed):
                        if text == "codigo": cmap["code"] = c
                        elif "descripcion" in text: cmap["description"] = c
                        elif "unidad" in text: cmap["unit"] = c
                        elif "cantidad" in text: cmap["quantity"] = c
                        elif text in {"p u", "pu", "precio unitario"}: cmap["unit_price"] = c
                        elif "subtotal" in text or "importe" in text: cmap["amount"] = c
                    break
        cmarket = _detect_concept_market_columns(rows[hidx] if hidx < len(rows) else [])
        data: list[CanonicalConcept] = []
        current_family = ""
        order = 0
        for ridx, row in enumerate(rows[hidx + 1:], hidx + 2):
            def val(field: str, default_idx: int | None = None):
                idx = cmap.get(field, default_idx)
                return row[idx] if idx is not None and idx < len(row) else None
            code = _txt(val("code", 0))
            desc = _txt(val("description", 1))
            if not desc and not code:
                texts = [_txt(v) for v in row if _txt(v)]
                # Ignore empty exported padding rows.
                if not texts:
                    continue
                desc = max(texts, key=len)
            if not desc or len(desc) < 3:
                continue
            # Ignore catalog grand-total rows. They are financial summary rows,
            # not catalog concepts to be compared or painted in Pareto 80/20.
            if not code and (_norm(desc) in {"subtotal", "i v a", "iva", "total"} or _num(desc) is not None):
                continue
            unit = _txt(val("unit"))
            op_guess = _txt(row[cmap.get("operator", -1)]) if "operator" in cmap and cmap.get("operator", -1) < len(row) else ""
            if _looks_like_apu_matrix_row(code, desc, unit, op_guess):
                continue
            qty = _num(val("quantity"))
            pu = _num(val("unit_price"))
            amount = _num(val("amount"))
            market_pu = _num(row[cmarket.get("market_unit_price")]) if cmarket.get("market_unit_price") is not None and cmarket.get("market_unit_price") < len(row) else None
            market_amount = _num(row[cmarket.get("market_amount")]) if cmarket.get("market_amount") is not None and cmarket.get("market_amount") < len(row) else None
            if amount is None and qty is not None and pu is not None:
                amount = qty * pu
            if market_amount is None and qty is not None and market_pu is not None:
                market_amount = qty * market_pu
            # Continuation row: PMD catalog descriptions may span multiple visual rows.
            # If a row has no code/unit/amount/PU, append it to the previous concept
            # instead of creating a fake catalog line.
            if (not code) and (not unit) and qty in (None, 0) and pu in (None, 0) and amount in (None, 0) and data and data[-1].is_executable:
                data[-1].description = (data[-1].description + " " + desc).strip()
                continue
            # Executable means it participates economically. Notes and hierarchy rows
            # may have unit/cantidad, but without P.U./importe they should not drive KPIs.
            executable = bool((pu is not None or amount is not None) and _norm(unit) not in {"nota"})
            level = _hierarchy_level(code, executable)
            if not executable:
                current_family = desc
            order += 1
            data.append(CanonicalConcept(
                code=code,
                description=desc,
                unit=unit,
                quantity=qty,
                unit_price=pu,
                amount=amount,
                market_unit_price=market_pu,
                market_amount=market_amount,
                market_source="Comparativa manual" if market_pu is not None or market_amount is not None else "",
                market_state="Mercado leído de comparativa" if market_pu is not None or market_amount is not None else "",
                family=current_family,
                hierarchy_level=level,
                original_order=order,
                is_executable=executable,
                source_row=ridx,
            ))
        # Final contract guardrail: this is a concept catalog, not PU detail.
        # Keep hierarchy/concept rows, drop any matrix/APU detail that may have
        # leaked in from a wrong sheet or ambiguous workbook.
        data = [c for c in data if not _looks_like_apu_matrix_row(c.code, c.description, c.unit)]
        _apply_pareto_80(data)
        return data[:500]
    finally:
        wb.close()


def _classify_section(text: str, current: str) -> str:
    n = _norm(text)
    if not n:
        return current
    if "material" in n:
        return "MATERIALES"
    if "mano de obra" in n or n == "mo" or "jornal" in n:
        return "MANO DE OBRA"
    if "maquinaria" in n or "equipo" in n or "herramienta" in n:
        return "MAQUINARIA"
    if "financier" in n or "indirect" in n or "utilidad" in n or "costo directo" in n:
        return "SECCION FINANCIERA"
    return current


def _percent_base(description: str, section: str) -> str:
    n = _norm(description + " " + section)
    if "material" in n:
        return "% SOBRE MATERIALES"
    if "mano" in n or "mo" in n or "herramienta" in n or "epp" in n:
        return "% SOBRE MO"
    if "maquinaria" in n or "equipo" in n:
        return "% SOBRE MAQUINARIA"
    if "directo ind" in n or "directo indirect" in n or "financ" in n:
        return "% SOBRE DIRECTO+IND"
    if "directo" in n or "indirect" in n or "utilidad" in n:
        return "% SOBRE DIRECTO"
    return section


def parse_matrix(path: Path, catalog: ReferenceCatalog | None = None) -> list[CanonicalApuItem]:
    wb, ws = _best_sheet(path, preferred_names=["PU", "APU", "MATRIZ", "MATRICES", "Detalle - P1", "Detalle - P2"])
    try:
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20000), values_only=True)]
        hidx, cmap = _detect_matrix_header(rows)
        current_section = ""
        current_concept_key = ""
        out: list[CanonicalApuItem] = []

        def cell(row: list[Any], field: str):
            idx = cmap.get(field)
            return row[idx] if idx is not None and idx < len(row) else None

        for ridx, row in enumerate(rows[hidx + 1:], hidx + 2):
            texts = [_txt(v) for v in row if _txt(v)]
            longest = max(texts, key=len) if texts else ""
            # Only update the current APU section from explicit section/header rows.
            # Do not infer a new section from normal item descriptions such as
            # "Equipo de protección" or "Herramienta menor" because that
            # corrupts subsequent subtotals and financial calculations.
            first_norm = _norm(texts[0]) if texts else ""
            section_header_terms = {"materiales", "mano de obra", "equipo y herramienta", "equipo", "maquinaria", "basicos", "básicos", "seccion financiera", "sección financiera"}
            if texts and (first_norm in section_header_terms or (len(texts) <= 2 and any(t in _norm(" ".join(texts)) for t in section_header_terms))):
                current_section = _classify_section(" ".join(texts[:4]), current_section)
            whole = _norm(" ".join(texts))

            if texts and (_norm(texts[0]) in {"partida", "analisis"} or str(texts[0]).strip().lower().startswith(("partida", "análisis", "analisis"))):
                label0 = _norm(texts[0])
                header_code = ""
                header_unit = ""
                header_qty = None
                header_amount = None
                # Typical matrix layout:
                # A: "Análisis:" | B: concept code | D: unit | F: quantity | G: total PU
                if label0.startswith("analisis") and len(row) >= 7:
                    header_code = _txt(row[1])
                    header_unit = _txt(row[3])
                    header_qty = _num(row[5])
                    header_amount = _num(row[6])
                    # Wide/manual PMD layouts can have the label in C and the
                    # analysis code several columns to the right. Use the first
                    # non-empty value after the label when the standard B cell is blank.
                    if not header_code:
                        label_pos = next((i for i, v in enumerate(row) if _norm(_txt(v)).startswith("analisis")), None)
                        if label_pos is not None:
                            for v in row[label_pos + 1:]:
                                txt = _txt(v)
                                if txt and not _norm(txt).startswith(("partida", "analisis")):
                                    header_code = txt
                                    break
                    if not header_unit:
                        # fallback to mapped unit header if available
                        ui = cmap.get("unit")
                        header_unit = _txt(row[ui]) if ui is not None and ui < len(row) else header_unit
                    if header_code:
                        current_concept_key = canonical_key(header_code, "")
                desc_header = " ".join(texts)
                out.append(CanonicalApuItem(
                    code=header_code,
                    concept_key=current_concept_key,
                    description=desc_header,
                    unit=header_unit,
                    section="PARTIDA",
                    quantity=header_qty,
                    amount=header_amount,
                    state="",
                    observation="Encabezado de análisis detectado",
                    source_row=ridx,
                ))
                continue

            if any(k in whole for k in ["subtotal materiales", "subtotal mano", "subtotal maquinaria", "subtotal equipo", "subtotal basicos", "costo directo", "total costo", "precio unitario", "materiales", "mano de obra", "maquinaria", "equipo y herramienta", "basicos", "seccion financiera", "indirectos", "utilidad"]):
                if len(texts) <= 4 or whole in {"materiales", "mano de obra", "maquinaria", "equipo", "equipo y herramienta", "basicos", "seccion financiera"} or "subtotal" in whole or "total" in whole or "costo directo" in whole or "utilidad" in whole or "precio unitario" in whole:
                    amount_guess = _num(cell(row, "amount"))
                    pct_guess = _num(cell(row, "percent"))
                    nums = [_num(v) for v in row if _num(v) is not None]
                    if amount_guess is None and nums:
                        # Fallback only when the sheet does not expose a mapped Importe column.
                        # Prefer not to take market columns as contractor amount.
                        amount_guess = nums[-2] if len(nums) >= 2 and ("precio" in whole or "total" in whole) else nums[-1]
                    if pct_guess is None and nums:
                        # Financial rows in wide PMD layouts often carry the
                        # percentage in an auxiliary column before the amount.
                        # Prefer the last plausible fraction/percentage-like
                        # number, not necessarily the last numeric cell.
                        candidates = [n for n in nums if abs(n) <= 1]
                        pct_guess = candidates[-1] if candidates else None
                    section = whole.upper()[:28] if ("subtotal" in whole or "total" in whole or "costo" in whole or "utilidad" in whole or "precio" in whole) else "TÍTULO"
                    desc_label = longest.upper()
                    if "subtotal" in whole:
                        target = next((t for t in texts if _norm(t) not in {"subtotal", "subtotal:"} and _num(t) is None), "")
                        target_norm = _norm(target)
                        desc_label = (target if target_norm.startswith("subtotal") else "SUBTOTAL " + target).strip().upper()
                    elif "costo directo" in whole:
                        desc_label = "COSTO DIRECTO"
                    elif "precio unitario" in whole or "total costo" in whole:
                        desc_label = "PRECIO UNITARIO" if "precio" in whole else "TOTAL COSTO UNITARIO"
                    elif "utilidad" in whole:
                        desc_label = "UTILIDAD"
                    mpu = _num(cell(row, "market_unit_price"))
                    mop = _txt(cell(row, "market_operator")) or ""
                    mqty = _num(cell(row, "market_quantity"))
                    mamount = _num(cell(row, "market_amount"))
                    if mamount is None and mpu is not None and mqty is not None:
                        mamount = mpu * mqty if mop != "/" else (mpu / mqty if mqty else None)
                    out.append(CanonicalApuItem(
                        concept_key=current_concept_key,
                        description=desc_label,
                        section=section or "TÍTULO",
                        amount=amount_guess,
                        percent=pct_guess,
                        market_unit_price=mpu,
                        market_operator=mop or "",
                        market_quantity=mqty,
                        market_amount=mamount,
                        state="Mercado declarado en matriz" if mamount is not None or mpu is not None else "",
                        observation="Fila estructural detectada" + (" con mercado declarado" if mamount is not None or mpu is not None else ""),
                        source_row=ridx,
                    ))
                    continue

            desc_idx = cmap.get("description")
            desc = _txt(row[desc_idx]) if desc_idx is not None and desc_idx < len(row) else longest
            if not desc or len(desc) < 3:
                continue
            code = _txt(cell(row, "code"))
            unit = _txt(cell(row, "unit"))
            pu = _num(cell(row, "unit_price"))
            op = _txt(cell(row, "operator")) or "*"
            qty = _num(cell(row, "quantity"))
            amount = _num(cell(row, "amount"))
            pct = _num(cell(row, "percent"))
            mpu = _num(cell(row, "market_unit_price"))
            mop = _txt(cell(row, "market_operator")) or "*"
            mqty = _num(cell(row, "market_quantity"))
            mamount = _num(cell(row, "market_amount"))

            if amount is None and pu is not None and qty is not None:
                amount = pu * qty if op != "/" else (pu / qty if qty else None)
            if mamount is None and mpu is not None and mqty is not None:
                mamount = mpu * mqty if mop != "/" else (mpu / mqty if mqty else None)

            # Concept header rows in a PU matrix define the current concept bucket.
            # All following insumos/subtotals inherit this key until the next header.
            looks_like_concept_header = bool(code and unit and amount is not None and not code.strip().startswith("%") and current_section in {"", "SIN SECCIÓN"})
            if looks_like_concept_header:
                current_concept_key = canonical_key(code, desc)

            section = current_section or "SIN SECCIÓN"
            is_declared_percent = (op == "%") or unit.strip() == "%" or code.strip().startswith("%") or "%" in desc
            if is_declared_percent:
                section = _percent_base(desc, section)
                if qty is None and pct is not None:
                    qty = pct / 100 if pct > 1 else pct

            item = CanonicalApuItem(
                code=code,
                concept_key=current_concept_key,
                description=desc,
                unit=unit,
                section=section,
                unit_price=pu,
                operator=op or "*",
                quantity=qty,
                amount=amount,
                percent=pct,
                market_unit_price=mpu,
                market_operator=mop or "*",
                market_quantity=mqty,
                market_amount=mamount,
                market_deviation=(pu / mpu - 1) if pu is not None and mpu else None,
                state="Mercado declarado en matriz" if (mpu is not None or mamount is not None) else "Pendiente mercado",
                observation="Leído desde matriz/APU" + (" · mercado declarado" if (mpu is not None or mamount is not None) else ""),
                source_row=ridx,
            )
            if (item.market_unit_price is None and item.market_amount is None) and catalog:
                match = catalog.match(desc, section)
                if match:
                    item.market_unit_price = float(match["price"])
                    item.market_quantity = qty
                    item.market_operator = op or "*"
                    item.market_amount = item.market_unit_price * qty if qty is not None else None
                    item.market_deviation = (pu / item.market_unit_price - 1) if pu is not None and item.market_unit_price else None
                    item.state = "Match mercado"
                    item.observation = f"{match['source']} · confianza {match['confidence']}"
                else:
                    item.state = "Sin referencia"
                    item.observation = "No se encontró match granular en data"
            out.append(item)
        _post_process_market_financials(out)
        return out[:6000]
    finally:
        wb.close()


def _canonical_section_name(section: str, description: str = "") -> str:
    sn = _norm(section or "")
    dn = _norm(description or "")
    n = (sn + " " + dn).strip()
    # Canonical percentage bases declared by the parser have priority over
    # words that may appear in the item description, e.g. EPP contains "Equipo"
    # but is commonly declared as % SOBRE MO.
    if "sobre materiales" in sn:
        return "MATERIALES"
    if "sobre mo" in sn or "mano de obra" in sn:
        return "MANO DE OBRA"
    if "sobre maquinaria" in sn or "maquinaria" in sn:
        return "MAQUINARIA"
    if "material" in n:
        return "MATERIALES"
    # Explicit equipment/maquinaria titles win over the word "herramienta".
    if "maquinaria" in n or "equipo y herramienta" in n or sn == "equipo" or "montacargas" in n:
        return "MAQUINARIA"
    if "mano" in n or sn == "mo" or "herramienta menor" in n or "epp" in n:
        return "MANO DE OBRA"
    if "equipo" in n:
        return "MAQUINARIA"
    if "basico" in n:
        return "BASICOS"
    return section or "SIN SECCIÓN"


def _is_structural_item(item: CanonicalApuItem) -> bool:
    s = item.section or ""
    d = _norm(item.description or "")
    return (
        s in {"PARTIDA", "TÍTULO"}
        or s.startswith("SUBTOTAL")
        or any(k in d for k in ["subtotal", "costo directo", "indirecto", "utilidad", "financiamiento", "precio unitario", "total costo", "total por servicio", "importe", "volumen", "rendimiento"])
    )


def _post_process_market_financials(items: list[CanonicalApuItem]) -> None:
    """Calculate market subtotals and financial rows inside the canonical model.

    Matrix files generally contain contractor PU/APU values only. Market values
    must be calculated from matched granular data and then propagated to:
    section subtotals, volume rows, costo directo, indirectos/utilidad and
    precio unitario. The Excel writer only renders these canonical values.
    """
    groups: dict[str, list[CanonicalApuItem]] = {}
    order: list[str] = []
    for it in items:
        key = it.concept_key or "__global__"
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(it)

    for key in order:
        group = groups[key]
        section_market: dict[str, float] = {}
        last_section = ""
        last_importe_market_by_section: dict[str, float] = {}
        subtotal_market_by_section: dict[str, float] = {}
        direct_market: float | None = None
        financial_market_total: float = 0.0

        for it in group:
            dnorm = _norm(it.description or "")
            # Track explicit section titles.
            if it.section == "TÍTULO":
                sec = _canonical_section_name(it.description, it.description)
                if sec in {"MATERIALES", "MANO DE OBRA", "MAQUINARIA", "BASICOS"}:
                    last_section = sec
                continue

            # Regular line item: calculate market amount from matched unit price.
            if not _is_structural_item(it):
                sec = _canonical_section_name(it.section, it.description)
                if sec in {"MATERIALES", "MANO DE OBRA", "MAQUINARIA", "BASICOS"}:
                    last_section = sec
                if it.market_amount is None and it.market_unit_price is not None and it.quantity is not None:
                    if (it.market_operator or "*") == "/":
                        it.market_amount = it.market_unit_price / it.quantity if it.quantity else None
                    else:
                        it.market_amount = it.market_unit_price * it.quantity
                if it.market_deviation is None and it.unit_price is not None and it.market_unit_price:
                    it.market_deviation = it.unit_price / it.market_unit_price - 1
                if it.market_amount is not None:
                    section_market[sec] = section_market.get(sec, 0.0) + float(it.market_amount or 0)
                continue

            # Structural/calculation rows.
            if "importe" in dnorm and "subtotal" not in dnorm:
                sec = last_section or _canonical_section_name(it.section, it.description)
                val = section_market.get(sec)
                if val is not None:
                    it.market_amount = val
                    it.market_unit_price = val
                    it.market_quantity = None
                    it.market_operator = ""
                    it.state = "Cálculo mercado"
                    it.observation = f"Importe mercado calculado desde insumos de {sec}"
                    last_importe_market_by_section[sec] = val
                continue

            if "volumen" in dnorm or "rendimiento" in dnorm:
                sec = last_section or _canonical_section_name(it.section, it.description)
                base = last_importe_market_by_section.get(sec, section_market.get(sec))
                vol = it.quantity
                if base is not None and vol is not None:
                    it.market_unit_price = base
                    # Manual PMD uses both patterns: Volumen multiplies, Rendimiento divides.
                    if "rendimiento" in dnorm:
                        it.market_operator = "/"
                        it.market_amount = base / vol if vol else None
                        it.observation = f"Rendimiento mercado = importe {sec} / rendimiento"
                    else:
                        it.market_operator = "*"
                        it.market_amount = base * vol
                        it.observation = f"Volumen mercado = importe {sec} × volumen"
                    it.market_quantity = vol
                    it.state = "Cálculo mercado"
                    subtotal_market_by_section[sec] = it.market_amount
                continue

            if "subtotal" in dnorm:
                sec = _canonical_section_name(it.description, it.description)
                val = subtotal_market_by_section.get(sec, section_market.get(sec))
                if val is not None:
                    it.market_amount = val
                    it.market_unit_price = val
                    it.market_operator = ""
                    it.market_quantity = None
                    it.state = "Cálculo mercado"
                    it.observation = f"Subtotal mercado calculado para {sec}"
                    subtotal_market_by_section[sec] = val
                continue

            if "costo directo" in dnorm:
                direct_market = sum(float(v or 0) for v in subtotal_market_by_section.values())
                if direct_market:
                    it.market_amount = direct_market
                    it.market_unit_price = direct_market
                    it.market_operator = ""
                    it.market_quantity = None
                    it.state = "Cálculo mercado"
                    it.observation = "Costo directo mercado = suma de subtotales de secciones"
                continue

            if "indirect" in dnorm or "utilidad" in dnorm or "financ" in dnorm:
                base = direct_market
                pct = it.quantity if it.quantity is not None else it.percent
                if base is not None and pct is not None:
                    factor = pct / 100 if pct > 1 else pct
                    it.market_unit_price = base
                    it.market_operator = "*"
                    it.market_quantity = factor
                    it.market_amount = base * factor
                    it.state = "Cálculo mercado"
                    it.observation = "Cargo financiero mercado calculado sobre costo directo"
                    financial_market_total += float(it.market_amount or 0)
                continue

            if "precio unitario" in dnorm or "total costo" in dnorm or "total por servicio" in dnorm:
                base = direct_market or 0.0
                total = base + financial_market_total
                if total:
                    it.market_amount = total
                    it.market_unit_price = total
                    it.market_operator = ""
                    it.market_quantity = None
                    it.state = "Cálculo mercado"
                    it.observation = "Precio unitario mercado = costo directo + cargos financieros"
                continue


def canonical_rows_from_items(items: list[CanonicalApuItem], include_market: bool = True) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    # If the parser already preserved PU structural rows (PARTIDA, TÍTULO,
    # SUBTOTAL, TOTAL), do not synthesize extra section titles. This keeps the
    # output close to the contractor's PU definition.
    has_structural = any(
        (i.section in {"PARTIDA", "TÍTULO"}) or i.section.startswith("SUBTOTAL") or "TOTAL" in i.section or "COSTO" in i.section
        for i in items
    )
    last_section = None
    for item in items:
        structural = (item.section in {"PARTIDA", "TÍTULO"}) or item.section.startswith("SUBTOTAL") or "TOTAL" in item.section or "COSTO" in item.section or "UTILIDAD" in item.section or "PRECIO" in item.section
        if (not has_structural) and item.section != last_section and item.section and not item.section.startswith("%"):
            rows.append({"code":"", "concept": item.section, "unit":"", "section":"TÍTULO", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":""})
            last_section = item.section
        amount = item.amount if item.amount is not None else ""
        market_amount = item.market_amount if item.market_amount is not None else ""
        rows.append({
            "code": item.code,
            "concept": item.description,
            "unit": item.unit,
            "section": item.section,
            "pu": "" if structural else item.unit_price if item.unit_price is not None else "",
            "op": "" if structural else item.operator,
            "qty": "" if structural else item.quantity if item.quantity is not None else "",
            "amount": amount,
            "pct": item.percent / 100 if item.percent and item.percent > 1 else item.percent if item.percent is not None else "",
            "mpu": item.market_unit_price if include_market and item.market_unit_price is not None else "",
            "mop": item.market_operator if include_market and (item.market_unit_price is not None or item.market_amount is not None) else "",
            "mqty": item.market_quantity if include_market and item.market_quantity is not None else "",
            "mamount": market_amount if include_market else "",
            "dev": item.market_deviation if include_market and item.market_deviation is not None else "",
            "state": item.state,
            "obs": item.observation,
        })
    return rows



def apply_provider_market_to_concepts(provider: CanonicalProvider) -> None:
    """Attach market totals to provider concepts from its canonical APU detail.

    This is a canonical model step, not an Excel styling patch. Comparativa
    compares concept-level P.U. totals; Detalle explains the APU. Therefore the
    provider must expose concept-level market values derived from the APU market
    rows or from granular reference matches.
    """
    if not provider.concepts or not provider.apu_items:
        return
    concept_map = {canonical_key(c.code, c.description): c for c in provider.concepts}
    direct: dict[str, CanonicalApuItem] = {}
    aggregates: dict[str, float] = {}
    for item in provider.apu_items:
        key = item.concept_key or canonical_key(item.code, item.description)
        if not key or key not in concept_map:
            continue
        dnorm = _norm(item.description or "")
        # Preferred source: canonical financial row "PRECIO UNITARIO". This is
        # the market P.U. for the catalog concept. The Comparativa then multiplies
        # it by catalog quantity for market importe.
        if ("precio unitario" in dnorm or "total costo unitario" in dnorm) and item.market_amount is not None:
            direct[key] = item
            continue
        # Secondary source: concept header rows if a matrix explicitly declares
        # market values there.
        if item.code and canonical_key(item.code, item.description) == key and (item.market_amount is not None or item.market_unit_price is not None):
            direct[key] = item
        # Aggregate only true line-level market amounts as fallback. Exclude all
        # structural/calculation rows to avoid double counting subtotals, volume,
        # costo directo and precio unitario.
        structural = _is_structural_item(item)
        if not structural and item.market_amount is not None:
            aggregates[key] = aggregates.get(key, 0.0) + float(item.market_amount or 0)
    for key, concept in concept_map.items():
        qty = concept.quantity or 1
        if key in direct:
            item = direct[key]
            # For a final precio unitario row, market_amount represents market P.U.
            mpu = item.market_amount if item.market_amount is not None else item.market_unit_price
            mamount = (float(mpu) * float(qty)) if mpu is not None and qty else mpu
            concept.market_unit_price = mpu
            concept.market_amount = mamount
            concept.market_source = "APU financiero"
            concept.market_state = "Mercado calculado desde detalle"
        elif key in aggregates:
            mpu = aggregates[key]
            concept.market_unit_price = mpu
            concept.market_amount = (mpu * float(qty)) if qty else mpu
            concept.market_source = "APU granular"
            concept.market_state = "Mercado agregado desde insumos"


def run_id(prefix: str) -> str:
    return f"{prefix}-{int(time.time())}"
