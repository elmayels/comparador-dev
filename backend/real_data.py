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
    family: str = ""
    source_row: int | None = None


@dataclass
class CanonicalApuItem:
    code: str = ""
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
                            self.items.append({"description": desc, "price": price, "kind": kind, "source": path.name, "tokens": _tokens(desc)})
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
        for item in self.items:
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


def _best_sheet(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = max(wb.worksheets, key=_sheet_score)
    return wb, ws


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
}


def parse_concepts(path: Path) -> list[CanonicalConcept]:
    wb, ws = _best_sheet(path)
    try:
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 1000), values_only=True)]
        hidx, cmap = _detect_header(rows, CONCEPT_SYNONYMS)
        data = []
        for ridx, row in enumerate(rows[hidx + 1:], hidx + 2):
            desc = _txt(row[cmap.get("description", 1)] if cmap.get("description", 1) < len(row) else "")
            code = _txt(row[cmap.get("code", 0)] if cmap.get("code", 0) < len(row) else "")
            if not desc and not code:
                # fallback: select longest text as description
                texts = [_txt(v) for v in row if _txt(v)]
                desc = max(texts, key=len) if texts else ""
            if not desc or len(desc) < 4:
                continue
            qty = _num(row[cmap["quantity"]]) if "quantity" in cmap and cmap["quantity"] < len(row) else None
            pu = _num(row[cmap["unit_price"]]) if "unit_price" in cmap and cmap["unit_price"] < len(row) else None
            amount = _num(row[cmap["amount"]]) if "amount" in cmap and cmap["amount"] < len(row) else None
            unit = _txt(row[cmap["unit"]]) if "unit" in cmap and cmap["unit"] < len(row) else ""
            # If amount is missing but qty and pu exist, calculate.
            if amount is None and qty is not None and pu is not None:
                amount = qty * pu
            data.append(CanonicalConcept(code=code, description=desc, unit=unit, quantity=qty, unit_price=pu, amount=amount, source_row=ridx))
        return data[:300]
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
    wb, ws = _best_sheet(path)
    try:
        rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 3000), values_only=True)]
        hidx, cmap = _detect_header(rows, MATRIX_SYNONYMS)
        current_section = ""
        out: list[CanonicalApuItem] = []
        for ridx, row in enumerate(rows[hidx + 1:], hidx + 2):
            texts = [_txt(v) for v in row if _txt(v)]
            longest = max(texts, key=len) if texts else ""
            current_section = _classify_section(" ".join(texts[:4]), current_section)
            # Detect title/subtotal rows.
            whole = _norm(" ".join(texts))
            if any(k in whole for k in ["subtotal materiales", "subtotal mano", "subtotal maquinaria", "costo directo", "total costo", "materiales", "mano de obra", "maquinaria", "seccion financiera"]):
                if len(texts) <= 3 or whole in {"materiales", "mano de obra", "maquinaria", "equipo", "seccion financiera"} or "subtotal" in whole or "total" in whole:
                    section = current_section if "subtotal" not in whole and "total" not in whole else whole.upper()[:28]
                    out.append(CanonicalApuItem(description=longest.upper(), section=section or "TÍTULO", state="", observation="Fila estructural detectada", source_row=ridx))
                    continue
            desc_idx = cmap.get("description")
            desc = _txt(row[desc_idx]) if desc_idx is not None and desc_idx < len(row) else longest
            if not desc or len(desc) < 3:
                continue
            code = _txt(row[cmap["code"]]) if "code" in cmap and cmap["code"] < len(row) else ""
            unit = _txt(row[cmap["unit"]]) if "unit" in cmap and cmap["unit"] < len(row) else ""
            pu = _num(row[cmap["unit_price"]]) if "unit_price" in cmap and cmap["unit_price"] < len(row) else None
            op = _txt(row[cmap["operator"]]) if "operator" in cmap and cmap["operator"] < len(row) else "*"
            qty = _num(row[cmap["quantity"]]) if "quantity" in cmap and cmap["quantity"] < len(row) else None
            amount = _num(row[cmap["amount"]]) if "amount" in cmap and cmap["amount"] < len(row) else None
            pct = _num(row[cmap["percent"]]) if "percent" in cmap and cmap["percent"] < len(row) else None
            if amount is None and pu is not None and qty is not None:
                amount = pu * qty if op != "/" else (pu / qty if qty else None)
            section = current_section or "SIN SECCIÓN"
            if op == "%" or (pct is not None) or "%" in desc:
                section = _percent_base(desc, section)
                if qty is None and pct is not None:
                    qty = pct / 100 if pct > 1 else pct
            item = CanonicalApuItem(code=code, description=desc, unit=unit, section=section, unit_price=pu, operator=op or "*", quantity=qty, amount=amount, percent=pct, state="Pendiente mercado", observation="Leído desde matriz/APU", source_row=ridx)
            if catalog:
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
        return out[:800]
    finally:
        wb.close()


def canonical_rows_from_items(items: list[CanonicalApuItem], include_market: bool = True) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    # If parser did not detect subtotal rows, add lightweight section subtotal separators.
    last_section = None
    for item in items:
        if item.section != last_section and item.section and not item.section.startswith("%") and "SUBTOTAL" not in item.section and "TOTAL" not in item.section:
            rows.append({"code":"", "concept": item.section, "unit":"", "section":"TÍTULO", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":""})
            last_section = item.section
        amount = item.amount if item.amount is not None else ""
        market_amount = item.market_amount if item.market_amount is not None else ""
        total_base = None
        rows.append({
            "code": item.code,
            "concept": item.description,
            "unit": item.unit,
            "section": item.section,
            "pu": item.unit_price if item.unit_price is not None else "",
            "op": item.operator,
            "qty": item.quantity if item.quantity is not None else "",
            "amount": amount,
            "pct": item.percent / 100 if item.percent and item.percent > 1 else item.percent if item.percent is not None else "",
            "mpu": item.market_unit_price if include_market and item.market_unit_price is not None else "",
            "mop": item.market_operator if include_market else "",
            "mqty": item.market_quantity if include_market and item.market_quantity is not None else "",
            "mamount": market_amount if include_market else "",
            "dev": item.market_deviation if include_market and item.market_deviation is not None else "",
            "state": item.state,
            "obs": item.observation,
        })
    return rows


def run_id(prefix: str) -> str:
    return f"{prefix}-{int(time.time())}"
