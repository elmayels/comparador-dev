from __future__ import annotations

import json
import os
import time
import urllib.request
import urllib.error
import html
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

from .real_data import (
    CanonicalProvider, CanonicalRun, ReferenceCatalog,
    canonical_rows_from_items, parse_concepts, parse_base_concepts, parse_matrix, run_id, apply_provider_market_to_concepts, classify_xlsx_role, generate_base_budget_from_concepts, canonical_key,
)

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
FRONTEND_DIR = ROOT / "frontend"
RUNTIME_DIR = ROOT / "backend" / "runtime"
REPORTS_DIR = RUNTIME_DIR / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Optional AI analysis layer
# ---------------------------------------------------------------------------
def _load_local_env() -> None:
    """Small .env loader with support for spaces around '='.

    Railway/native environment variables still win. This only helps local runs
    where the user writes lines such as `AI_ANALYSIS_MODEL = claude-haiku-4-5`
    in a .env file and expects the app to pick them up without python-dotenv.
    """
    for env_path in (ROOT / ".env", Path.cwd() / ".env"):
        if not env_path.exists():
            continue
        for raw_line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


_load_local_env()


def _normalize_ai_provider(value: str | None) -> str:
    raw = (value or "").strip().lower()
    aliases = {
        "anthopic": "anthropic",
        "antrophic": "anthropic",
        "claude": "anthropic",
        "open-ai": "openai",
    }
    return aliases.get(raw, raw)


def _resolved_ai_provider() -> str:
    explicit_provider = (os.getenv("AI_ANALYSIS_PROVIDER") or os.getenv("AI_PROVIDER") or "").strip()
    if explicit_provider:
        return _normalize_ai_provider(explicit_provider)
    model_hint = (os.getenv("AI_ANALYSIS_MODEL") or os.getenv("AI_MODEL") or "").strip().lower()
    if model_hint.startswith("claude"):
        return "anthropic"
    if os.getenv("ANTHROPIC_API_KEY") and not os.getenv("OPENAI_API_KEY"):
        return "anthropic"
    return "openai"


# Single feature flag requested for the first real version. When disabled or
# when credentials are not available, the application still generates a
# deterministic expert analysis from the canonical calculated data. No amounts,
# quantities or matches are invented by this layer.
AI_ANALYSIS_ENABLED = os.getenv("ENABLE_AI_ANALYSIS", "0").strip().lower() in {"1", "true", "yes", "on"}
AI_ANALYSIS_PROVIDER = _resolved_ai_provider()
AI_ANALYSIS_MODEL = (os.getenv("AI_ANALYSIS_MODEL") or os.getenv("AI_MODEL") or "").strip()
AI_ANALYSIS_TIMEOUT = float(os.getenv("AI_ANALYSIS_TIMEOUT", "25") or 25)
AI_ANALYSIS_MAX_INPUT_ITEMS = int(os.getenv("AI_ANALYSIS_MAX_INPUT_ITEMS", "12") or 12)
AI_ANALYSIS_LAST_ERROR: str | None = None
AI_ANALYSIS_LAST_PROVIDER_RESPONSE: str | None = None


def _configured_ai_model() -> str:
    """Return a provider-compatible model unless the deploy explicitly overrides it."""
    explicit = (os.getenv("AI_ANALYSIS_MODEL") or os.getenv("AI_MODEL") or "").strip()
    if explicit:
        return explicit
    if AI_ANALYSIS_PROVIDER == "anthropic":
        # Cheap/fast Anthropic default for the current prototype.
        return "claude-haiku-4-5"
    return "gpt-4o-mini"


app = FastAPI(title="Quantia APU Canonical", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

USERS = {
    "superadmin@demo.com": {"id": "u-001", "name": "Super Administrador", "email": "superadmin@demo.com", "role": "SUPER_ADMIN", "password": "demo123", "org": "Global"},
    "admin@demo.com": {"id": "u-002", "name": "Administrador Demo", "email": "admin@demo.com", "role": "ADMIN", "password": "demo123", "org": "Empresa Demo"},
    "analista@demo.com": {"id": "u-003", "name": "Analista Demo", "email": "analista@demo.com", "role": "ANALYST", "password": "demo123", "org": "Empresa Demo"},
}


def _classify_reference_file(path: Path) -> str:
    name = path.name.lower()
    if "matrices" in name:
        return "CONSTRUDATA_MATRICES_BASE_BUDGET"
    if "material" in name:
        return "MATERIALS_REFERENCE"
    if "mano" in name or "obra" in name:
        return "LABOR_REFERENCE"
    if "maquinaria" in name or "equipo" in name:
        return "EQUIPMENT_REFERENCE"
    if "porcentaje" in name or "parametro" in name:
        return "PERCENTAGES_REFERENCE"
    if path.suffix.lower() == ".json":
        return "AUXILIARY_JSON"
    return "AUXILIARY_REFERENCE"


def _inspect_xlsx(path: Path) -> dict:
    info = {"sheets": [], "readable": False}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets[:10]:
            info["sheets"].append({"name": ws.title, "rows": ws.max_row, "columns": ws.max_column})
        wb.close()
        info["readable"] = True
    except Exception as exc:
        info["error"] = f"{type(exc).__name__}: {exc}"
    return info


def _list_data_references() -> list[dict]:
    rows = []
    if not DATA_DIR.exists():
        return rows
    for p in sorted(DATA_DIR.iterdir()):
        if p.is_dir():
            continue
        item = {
            "name": p.name,
            "extension": p.suffix.lower(),
            "sizeBytes": p.stat().st_size,
            "kind": _classify_reference_file(p),
            "canonicalUse": "Presupuesto base" if "matrices" in p.name.lower() else "Detalle APU contratista" if p.suffix.lower() in [".xlsx", ".json"] else "Auxiliar",
        }
        if p.suffix.lower() == ".xlsx":
            item.update(_inspect_xlsx(p))
        rows.append(item)
    return rows


@app.get("/api/health")
def health():
    return {"status": "ok", "version": "0.0.1", "time": datetime.utcnow().isoformat() + "Z"}


@app.post("/api/auth/login")
def login(payload: dict):
    email = (payload.get("email") or "").strip().lower()
    password = payload.get("password") or ""
    user = USERS.get(email)
    if not user or user["password"] != password:
        raise HTTPException(status_code=401, detail="Credenciales inválidas")
    clean = {k: v for k, v in user.items() if k != "password"}
    return {"token": f"mock-token-{clean['id']}", "user": clean}


@app.get("/api/references")
def references():
    refs = _list_data_references()
    return {
        "dataDir": str(DATA_DIR),
        "references": refs,
        "summary": {
            "total": len(refs),
            "baseBudgetMatrix": len([r for r in refs if r["kind"] == "CONSTRUDATA_MATRICES_BASE_BUDGET"]),
            "materials": len([r for r in refs if r["kind"] == "MATERIALS_REFERENCE"]),
            "labor": len([r for r in refs if r["kind"] == "LABOR_REFERENCE"]),
            "equipment": len([r for r in refs if r["kind"] == "EQUIPMENT_REFERENCE"]),
        },
    }


def _validate_xlsx_uploads(files: List[UploadFile]) -> list[str]:
    names = []
    for f in files:
        if not f.filename or not f.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"Archivo no permitido: {f.filename}. Solo se acepta .xlsx")
        names.append(f.filename)
    return names


# Compatibility route: older frontend bundles used /api/base-budgets/mock-run
# and posted the file under the field name "files".  This route is now a real
# base-budget execution path so a cached/old UI cannot generate the legacy mock
# workbook by accident.
@app.post("/api/base-budgets/mock-run")
async def base_budget_mock_run(projectName: str = Form("Presupuesto base real"), files: List[UploadFile] = File(default=[])):
    if not files:
        raise HTTPException(status_code=400, detail="Carga un archivo .xlsx de conceptos para generar el presupuesto base real")
    _validate_xlsx_uploads(files)
    return await _execute_base_budget_real(projectName, files[0])


@app.post("/api/comparisons/mock-run")
async def comparison_mock_run(mode: str = Form("MULTI"), projectName: str = Form("Comparación demo"), files: List[UploadFile] = File(default=[])):
    names = _validate_xlsx_uploads(files)
    if mode.upper() == "MULTI" and len(names) < 2:
        raise HTTPException(status_code=400, detail="La comparación múltiple requiere mínimo dos archivos .xlsx")
    if mode.upper() == "SINGLE" and len(names) < 1:
        raise HTTPException(status_code=400, detail="La comparación individual requiere un archivo .xlsx")
    return {
        "id": f"RUN-{int(time.time())}",
        "mode": mode.upper(),
        "projectName": projectName,
        "status": "COMPLETED_WITH_WARNINGS",
        "uploadedFiles": names,
        "kpis": {
            "contractors": max(1, len(names)),
            "bestOffer": 1180000,
            "criticalItems": 18,
            "globalRisk": "Amarillo",
            "withoutReference": 14,
        },
        "findings": [
            "El comparador funciona sin presupuesto base obligatorio.",
            "El detalle APU por contratista usa matriz propia del contratista y referencias granulares desde data.",
            "No se usa construdata_matrices.xlsx para armar el detalle de contratistas.",
        ],
    }


# -----------------------------------------------------------------------------
# Professional Excel report engine V0
# -----------------------------------------------------------------------------
# The workbook below is deliberately a professional, clean deliverable. The
# historical Excel is only a semantic reference for the Comparativa and Detalle
# tabs. We do not copy its dark visual style or paste a full reference matrix.

BRAND = {
    "navy": "1F4E79",
    "blue": "D9EAF7",
    "blue2": "EDF6FC",
    "green": "E2F0D9",
    "green_text": "0E6B3D",
    "yellow": "FFF2CC",
    "red": "FCE4D6",
    "gray": "F2F4F7",
    "line": "D9E2EC",
    "text": "1F2937",
    "muted": "667085",
}

MONEY_FMT = '$#,##0.00'
PCT_FMT = '0.00%'
INT_FMT = '#,##0'


def _thin_border(color: str = "D9E2EC"):
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def _bottom_border(color: str = "CBD5E1"):
    return Border(bottom=Side(style="thin", color=color))


def _setup_sheet(ws, title: str, subtitle: str, max_col: int):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=17, color=BRAND["navy"])
    ws["A1"].alignment = Alignment(vertical="center")
    ws["A2"] = subtitle
    ws["A2"].font = Font(italic=True, color=BRAND["muted"])
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 34


def _set_widths(ws, widths: dict[str, float]):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _section_label(ws, row: int, label: str, max_col: int, fill: str = "EDF6FC"):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_col)
    cell = ws.cell(row, 1, label)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=True, color=BRAND["navy"])
    cell.alignment = Alignment(vertical="center")
    cell.border = _bottom_border()
    ws.row_dimensions[row].height = 22


def _write_kpi(ws, row: int, col: int, label: str, value, note: str = "", fill: str = "F8FAFC"):
    # 3-row, 2-column KPI card
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)
    ws.merge_cells(start_row=row+1, start_column=col, end_row=row+1, end_column=col+1)
    ws.merge_cells(start_row=row+2, start_column=col, end_row=row+2, end_column=col+1)
    for r in range(row, row+3):
        for c in range(col, col+2):
            ws.cell(r, c).fill = PatternFill("solid", fgColor=fill)
            ws.cell(r, c).border = _thin_border()
    ws.cell(row, col, label).font = Font(bold=True, size=9, color=BRAND["muted"])
    ws.cell(row+1, col, value).font = Font(bold=True, size=15, color=BRAND["navy"])
    ws.cell(row+2, col, note).font = Font(size=9, color=BRAND["muted"])
    for r in range(row, row+3):
        ws.cell(r, col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _header_style(ws, row: int, start_col: int, end_col: int, fill: str = "1F4E79", font_color: str = "FFFFFF"):
    for col in range(start_col, end_col + 1):
        c = ws.cell(row, col)
        c.fill = PatternFill("solid", fgColor=fill)
        c.font = Font(bold=True, color=font_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _thin_border("B6C2CF")
    ws.row_dimensions[row].height = 36


def _body_style(ws, min_row: int, max_row: int, min_col: int, max_col: int):
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            c = ws.cell(row, col)
            c.border = _thin_border("EAECF0")
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.font = Font(color=BRAND["text"])
        if row % 2 == 0:
            for col in range(min_col, max_col + 1):
                ws.cell(row, col).fill = PatternFill("solid", fgColor="FCFCFD")


def _add_table(ws, ref: str, name: str, style: str = "TableStyleMedium2"):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name=style, showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)
    return tab


def _apply_formats(ws, money_cols=None, pct_cols=None, int_cols=None, start_row=1, end_row=100):
    money_cols = money_cols or []
    pct_cols = pct_cols or []
    int_cols = int_cols or []
    for row in range(start_row, end_row + 1):
        for col in money_cols:
            ws.cell(row, col).number_format = MONEY_FMT
        for col in pct_cols:
            ws.cell(row, col).number_format = PCT_FMT
        for col in int_cols:
            ws.cell(row, col).number_format = INT_FMT


def _status_fill(value: str):
    value = str(value or "").lower()
    if "alto" in value or "rojo" in value or "sobrecosto" in value or "error" in value:
        return "FCE4D6"
    if "medio" in value or "amarillo" in value or "revisar" in value or "advert" in value:
        return "FFF2CC"
    if "verde" in value or "ok" in value or "alineado" in value:
        return "E2F0D9"
    return "F8FAFC"




def _safe_provider_names(provider_names: Optional[List[str]] = None) -> List[str]:
    """Normalize visible provider names for web/Excel.

    V0 rule: the short visible name comes from the UI textbox and is capped
    at 10 characters so it fits cleanly in Comparativa blocks and detail
    sheet tabs. These names are the canonical display names for summaries,
    tables and Excel sheets.
    """
    raw = provider_names or []
    cleaned = []
    seen = set()
    for idx, name in enumerate(raw, 1):
        short = (name or '').strip()[:10]
        if not short:
            short = f"Prov {idx}"
        # Excel sheet names cannot contain these characters.
        for ch in ['\\', '/', '?', '*', '[', ']', ':']:
            short = short.replace(ch, '-')
        base = short[:10]
        candidate = base
        n = 2
        while candidate.lower() in seen:
            suffix = str(n)
            candidate = (base[:10-len(suffix)] + suffix)[:10]
            n += 1
        seen.add(candidate.lower())
        cleaned.append(candidate)
    return cleaned or ['Prov A', 'Prov B']


def _sample_contractors(provider_names: Optional[List[str]] = None):
    names = _safe_provider_names(provider_names)
    defaults = names + [f"Prov {chr(65+i)}" for i in range(len(names), 3)]
    return [
        {"rank": 1, "contractor": defaults[1] if len(defaults) > 1 else defaults[0], "amount": 1180000, "diff_min": 0.00, "avg_dev": -0.078, "risk": "Bajo", "traffic": "Verde"},
        {"rank": 2, "contractor": defaults[0], "amount": 1250000, "diff_min": 0.059, "avg_dev": 0.024, "risk": "Medio", "traffic": "Amarillo"},
        {"rank": 3, "contractor": defaults[2] if len(defaults) > 2 else f"{defaults[0]} C", "amount": 1410000, "diff_min": 0.195, "avg_dev": 0.143, "risk": "Alto", "traffic": "Rojo"},
    ]


def _write_resumen_ejecutivo(ws, mode: str, provider_names: Optional[List[str]] = None):
    _setup_sheet(ws, "Resumen Ejecutivo", "Vista gerencial: KPIs, ranking, hallazgos y navegación interna del reporte.", 10)
    _set_widths(ws, {"A": 20, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18, "J": 18})
    ws.freeze_panes = "A12"

    _write_kpi(ws, 4, 1, "Mejor oferta", 1180000, (_safe_provider_names(provider_names)[1] if len(_safe_provider_names(provider_names)) > 1 else _safe_provider_names(provider_names)[0]), "F8FAFC")
    _write_kpi(ws, 4, 3, "Riesgo global", "Amarillo", "Con advertencias", "FFF7E6")
    _write_kpi(ws, 4, 5, "Partidas críticas", 18, "Prioridad alta/media", "F8FAFC")
    _write_kpi(ws, 4, 7, "Sin referencia", 14, "Revisar alcance", "F8FAFC")
    _write_kpi(ws, 4, 9, "Modo", "Multi" if mode == "comparison" else "Base", "V0 mock", "F8FAFC")
    for cell in ["A5", "C5", "E5", "G5", "I5"]:
        if isinstance(ws[cell].value, (int, float)):
            ws[cell].number_format = MONEY_FMT if cell == "A5" else INT_FMT

    _section_label(ws, 9, "Ranking económico", 10)
    headers = ["Posición", "Contratista", "Monto total", "Dif. vs menor", "Desv. promedio", "Riesgo", "Semáforo"]
    for i, h in enumerate(headers, 1):
        ws.cell(10, i, h)
    _header_style(ws, 10, 1, len(headers), fill="1F4E79")
    for r, item in enumerate(_sample_contractors(provider_names), 11):
        vals = [item["rank"], item["contractor"], item["amount"], item["diff_min"], item["avg_dev"], item["risk"], item["traffic"]]
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v)
        ws.cell(r, 6).fill = PatternFill("solid", fgColor=_status_fill(item["risk"]))
        ws.cell(r, 7).fill = PatternFill("solid", fgColor=_status_fill(item["traffic"]))
    _body_style(ws, 11, 13, 1, len(headers))
    _apply_formats(ws, money_cols=[3], pct_cols=[4, 5], start_row=11, end_row=13)
    _add_table(ws, "A10:G13", "ExecutiveRanking", "TableStyleMedium2")

    _section_label(ws, 16, "Hallazgos y recomendaciones", 10)
    findings = [
        ["Hallazgo", "Cinco partidas explican la mayor parte de la diferencia económica entre contratistas."],
        ["Riesgo", "Existen insumos y porcentajes de referencia que deben verificarse en Detalle APU antes de negociar."],
        ["Recomendación", "Priorizar revisión por impacto económico, no solo por desviación porcentual."],
        ["Limitación", "V0 usa datos mockeados; V1 debe leer la matriz/APU real de cada contratista."],
    ]
    for r, row in enumerate(findings, 17):
        ws.cell(r, 1, row[0]).font = Font(bold=True, color=BRAND["navy"])
        ws.cell(r, 2, row[1])
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        for c in range(1, 11):
            ws.cell(r, c).border = _thin_border("EAECF0")
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[17].height = 30
    ws.row_dimensions[18].height = 30
    ws.row_dimensions[19].height = 30
    ws.row_dimensions[20].height = 30

    _section_label(ws, 23, "Navegación", 10)
    names = _safe_provider_names(provider_names)
    links = [("Comparativa", "#'Comparativa'!A1")]
    links += [(f"Detalle {name}", f"#'Detalle - {name}'!A1") for name in names[:4]]
    links += [("Partidas Críticas", "#'Partidas Críticas'!A1"), ("Insumos Críticos", "#'Insumos Críticos'!A1"), ("Validaciones", "#'Validaciones'!A1")]
    for idx, (text, target) in enumerate(links, 1):
        cell = ws.cell(24, idx, text)
        cell.hyperlink = target
        cell.style = "Hyperlink"
        cell.alignment = Alignment(horizontal="center")


def _provider_group_header(ws, row: int, start_col: int, end_col: int, title: str, fill: str):
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = ws.cell(row, start_col, title)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(start_col, end_col + 1):
        ws.cell(row, c).border = _thin_border("B6C2CF")
    ws.row_dimensions[row].height = 24


def _write_comparativa_professional(ws, provider_names: Optional[List[str]] = None):
    """Horizontal economic comparison by provider.

    This sheet is intentionally horizontal because it compares providers on the
    same commercial concepts. The provider block names come from the UI textbox
    (max 10 chars) and are propagated to summaries and Excel detail tabs.
    """
    names = _safe_provider_names(provider_names)
    ws.sheet_view.showGridLines = False
    base_cols = 4
    provider_cols = 6
    total_cols = base_cols + provider_cols * len(names)
    widths = {"A": 14, "B": 54, "C": 12, "D": 11}
    for idx, _ in enumerate(names):
        start = base_cols + idx * provider_cols + 1
        for offset, width in enumerate([14, 16, 12, 12, 16, 16]):
            widths[get_column_letter(start + offset)] = width
    _set_widths(ws, widths)
    ws.freeze_panes = "E3"
    _provider_group_header(ws, 1, 1, 4, "Servicios / Cotización", "475467")
    palette = ["1F4E79", "0E6B3D", "7C3AED", "B54708"]
    for idx, name in enumerate(names):
        start_col = base_cols + idx * provider_cols + 1
        _provider_group_header(ws, 1, start_col, start_col + provider_cols - 1, name, palette[idx % len(palette)])
    headers = ["Partida", "Descripción", "Unidad", "Cantidad"]
    for _ in names:
        headers += ["P.U.", "Importe", "% Part.", "% ajuste", "Mercado P.U.", "Mercado Importe"]
    for c, h in enumerate(headers, 1):
        ws.cell(2, c, h)
    _header_style(ws, 2, 1, 4, fill="475467")
    for idx, _ in enumerate(names):
        start_col = base_cols + idx * provider_cols + 1
        _header_style(ws, 2, start_col, start_col + provider_cols - 1, fill=palette[idx % len(palette)])

    base_rows = [
        ["FLEX41.11", "Instalación de bomba centrífuga. Incluye materiales, mano de obra, herramientas, equipos, traslados, limpieza y puesta en sitio.", "PZA", 1],
        ["FLEX41.12", "Suministro e instalación de guarda protectora en acero inoxidable 304.", "PZA", 1],
        ["FLEX41.20", "Instalación de válvula doble asiento. Incluye soportes, anclajes, izajes, acarreo y limpieza.", "PZA", 1],
        ["FLEX41.25", "Instalación de válvula check. Incluye soportería, anclajes, maniobras, mano de obra y limpieza final.", "PZA", 1],
    ]
    provider_prices = [
        [19187.17, 16437.07, 38976.81, 39030.70],
        [15842.49, 13678.38, 19016.66, 19066.29],
        [20220.00, 17105.00, 35500.00, 36550.00],
        [18440.00, 14990.00, 28620.00, 30100.00],
    ]
    market_prices = [
        [7238.26, 3523.68, 6727.88, 6684.68],
        [6618.23, 3208.53, 6208.82, 6168.96],
        [7001.55, 3412.20, 6555.30, 6501.10],
        [6900.00, 3300.00, 6440.00, 6400.00],
    ]
    for r_idx, row in enumerate(base_rows, 3):
        for c, v in enumerate(row, 1):
            ws.cell(r_idx, c, v)
        for p_idx, _ in enumerate(names):
            start_col = base_cols + p_idx * provider_cols + 1
            prices = provider_prices[p_idx % len(provider_prices)]
            markets = market_prices[p_idx % len(market_prices)]
            qty_cell = f"D{r_idx}"
            pu = prices[r_idx - 3]
            market = markets[r_idx - 3]
            ws.cell(r_idx, start_col, pu)
            ws.cell(r_idx, start_col + 1, f"={qty_cell}*{get_column_letter(start_col)}{r_idx}")
            total_ref = f"${get_column_letter(start_col+1)}$7"
            ws.cell(r_idx, start_col + 2, f"={get_column_letter(start_col+1)}{r_idx}/{total_ref}")
            ws.cell(r_idx, start_col + 3, f"=IF({get_column_letter(start_col+4)}{r_idx}=0,0,{get_column_letter(start_col)}{r_idx}/{get_column_letter(start_col+4)}{r_idx}-1)")
            ws.cell(r_idx, start_col + 4, market)
            ws.cell(r_idx, start_col + 5, f"={qty_cell}*{get_column_letter(start_col+4)}{r_idx}")
    total_row = 7
    ws.cell(total_row, 1, "TOTAL")
    for p_idx, _ in enumerate(names):
        start_col = base_cols + p_idx * provider_cols + 1
        ws.cell(total_row, start_col, f"=AVERAGE({get_column_letter(start_col)}3:{get_column_letter(start_col)}6)")
        ws.cell(total_row, start_col+1, f"=SUM({get_column_letter(start_col+1)}3:{get_column_letter(start_col+1)}6)")
        ws.cell(total_row, start_col+4, f"=AVERAGE({get_column_letter(start_col+4)}3:{get_column_letter(start_col+4)}6)")
        ws.cell(total_row, start_col+5, f"=SUM({get_column_letter(start_col+5)}3:{get_column_letter(start_col+5)}6)")
    _body_style(ws, 3, total_row, 1, total_cols)
    for c in range(1, total_cols + 1):
        ws.cell(total_row, c).font = Font(bold=True, color=BRAND["navy"])
        ws.cell(total_row, c).fill = PatternFill("solid", fgColor="F8FAFC")
        ws.cell(total_row, c).border = Border(top=Side(style="medium", color="1F4E79"), bottom=Side(style="thin", color="D9E2EC"))
    money_cols, pct_cols = [], []
    for idx in range(len(names)):
        start_col = base_cols + idx * provider_cols + 1
        money_cols += [start_col, start_col+1, start_col+4, start_col+5]
        pct_cols += [start_col+2, start_col+3]
        ws.conditional_formatting.add(f"{get_column_letter(start_col+3)}3:{get_column_letter(start_col+3)}6", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    _apply_formats(ws, money_cols=money_cols, pct_cols=pct_cols, start_row=3, end_row=total_row)
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}{total_row}"

    _section_label(ws, 10, "Resumen individual por proveedor", total_cols)
    row = 11
    for idx, name in enumerate(names):
        ws.cell(row, 1, name).font = Font(bold=True, color=BRAND["navy"], size=12)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
        row += 1
        bullets = [
            "El nombre visible proviene del textbox de carga y se usa en resúmenes, Comparativa y tabs de Detalle.",
            "La comparación de mercado por proveedor se muestra en columnas propias dentro de su bloque.",
            "Las desviaciones deben validarse en el tab Detalle individual, porque cada matriz/APU puede tener estructura diferente.",
        ]
        for bullet in bullets:
            ws.cell(row, 1, "• " + bullet)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
            ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[row].height = 28
            row += 1
        row += 1

def _canonical_detail_headers(include_market: bool) -> list[str]:
    """Canonical APU detail columns used by both base budgets and provider comparisons.

    The same canonical layout is reused for:
    - BASE_BUDGET detail: generated from engineering concepts + Construdata matrices.
      Market columns are omitted because this detail is already the market/base matrix.
    - CONTRACTOR_APU detail: generated from each provider's own matrix/APU.
      Market columns are included to compare against granular data references.
    """
    base = [
        "Código", "Concepto / insumo", "Unidad", "Sección", "P. Unitario",
        "Op.", "Cantidad", "Importe", "%"
    ]
    market = ["Mercado P.U.", "Mercado Op.", "Mercado Cant.", "Mercado Importe", "Diferencia %"]
    tail = ["Estado", "Observación"]
    return base + (market if include_market else []) + tail


def _canonical_apu_rows(variant: str = "A") -> list[dict]:
    """Return V0 canonical APU rows.

    V0 still uses controlled mock values, but the row model mirrors the real
    canonical structure that V1 must populate from either:
    - a generated base-budget matrix, or
    - the provider's uploaded matrix/APU.
    """
    factor = 1 if variant == "A" else 0.88
    mat_rate = 0.08 if variant == "A" else 0.06
    mo_tool_rate = 0.09 if variant == "A" else 0.05
    epp_rate = 0.09 if variant == "A" else 0.07
    machine_rate = 0.05 if variant == "A" else 0.04
    indirect_rate = 0.25
    finance_rate = 0.0285 if variant == "A" else 0.00
    return [
        {"code":"FLEX41.11", "concept":"INSTALACIÓN DE BOMBA CENTRÍFUGA FRISTAM MODELO FPR 3531-155", "unit":"PZA", "section":"PARTIDA", "pu":"", "op":"", "qty":1, "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":"Matriz canónica generada/normalizada desde la fuente del proceso"},
        {"code":"", "concept":"MATERIALES", "unit":"", "section":"TÍTULO", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":""},
        {"code":"MAT-BOMBA", "concept":"Materiales menores para montaje de bomba", "unit":"LOTE", "section":"MATERIALES", "pu":4500*factor, "op":"*", "qty":1, "amount":"=E5*G5", "pct":"=H5/$H$25", "mpu":4500, "mop":"*", "mqty":1, "mamount":"=J5*L5", "dev":"=IF(J5=0,0,E5/J5-1)", "state":"OK", "obs":"Referencia granular: materiales data"},
        {"code":"ANCL-M10", "concept":"Anclaje mecánico acero inoxidable M10", "unit":"PZA", "section":"MATERIALES", "pu":120*factor, "op":"*", "qty":8, "amount":"=E6*G6", "pct":"=H6/$H$25", "mpu":120, "mop":"*", "mqty":8, "mamount":"=J6*L6", "dev":"=IF(J6=0,0,E6/J6-1)", "state":"OK", "obs":"Referencia granular: materiales data"},
        {"code":"SOL-INOX", "concept":"Consumible soldadura acero inoxidable", "unit":"KG", "section":"MATERIALES", "pu":285*factor, "op":"*", "qty":1.5, "amount":"=E7*G7", "pct":"=H7/$H$25", "mpu":300.8, "mop":"*", "mqty":1.5, "mamount":"=J7*L7", "dev":"=IF(J7=0,0,E7/J7-1)", "state":"Revisar" if variant == "A" else "OK", "obs":"Precio validado contra data/materiales"},
        {"code":"%CONS-MAT", "concept":"Consumibles declarados como % sobre subtotal de MATERIALES", "unit":"%", "section":"% SOBRE MATERIALES", "pu":"=SUM(H5:H7)", "op":"%", "qty":mat_rate, "amount":"=E8*G8", "pct":"=H8/$H$25", "mpu":"=SUM(M5:M7)", "mop":"%", "mqty":mat_rate, "mamount":"=J8*L8", "dev":"=IF(J8=0,0,E8/J8-1)", "state":"OK", "obs":"El % aplica solo sobre subtotal materiales"},
        {"code":"", "concept":"SUBTOTAL MATERIALES", "unit":"", "section":"SUBTOTAL MATERIALES", "pu":"=SUM(H5:H7)", "op":"", "qty":"", "amount":"=SUM(H5:H8)", "pct":"=H9/$H$25", "mpu":"=SUM(M5:M7)", "mop":"", "mqty":"", "mamount":"=SUM(M5:M8)", "dev":"=IF(M9=0,0,H9/M9-1)", "state":"", "obs":"Incluye insumos + % aplicables a materiales"},
        {"code":"", "concept":"MANO DE OBRA", "unit":"", "section":"TÍTULO", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":""},
        {"code":"SUP-O", "concept":"Supervisor Obra", "unit":"JOR", "section":"MO", "pu":2643.19*factor, "op":"*", "qty":0.1, "amount":"=E11*G11", "pct":"=H11/$H$25", "mpu":1757.33, "mop":"*", "mqty":0.1, "mamount":"=J11*L11", "dev":"=IF(J11=0,0,E11/J11-1)", "state":"Revisar", "obs":"Referencia granular: mano de obra data"},
        {"code":"OF-SOL", "concept":"Oficial soldador / argonero", "unit":"JOR", "section":"MO", "pu":2120.45*factor, "op":"*", "qty":1, "amount":"=E12*G12", "pct":"=H12/$H$25", "mpu":1326.14, "mop":"*", "mqty":1, "mamount":"=J12*L12", "dev":"=IF(J12=0,0,E12/J12-1)", "state":"Revisar", "obs":"Referencia granular: mano de obra data"},
        {"code":"AYU-GRAL", "concept":"Ayudante general", "unit":"JOR", "section":"MO", "pu":1074.96*factor, "op":"*", "qty":1, "amount":"=E13*G13", "pct":"=H13/$H$25", "mpu":777.88, "mop":"*", "mqty":1, "mamount":"=J13*L13", "dev":"=IF(J13=0,0,E13/J13-1)", "state":"OK", "obs":"Referencia granular: mano de obra data"},
        {"code":"%HERR", "concept":"Herramienta menor declarada como % sobre subtotal MO", "unit":"%", "section":"% SOBRE MO", "pu":"=SUM(H11:H13)", "op":"%", "qty":mo_tool_rate, "amount":"=E14*G14", "pct":"=H14/$H$25", "mpu":"=SUM(M11:M13)", "mop":"%", "mqty":mo_tool_rate, "mamount":"=J14*L14", "dev":"=IF(J14=0,0,E14/J14-1)", "state":"OK", "obs":"El % aplica solo sobre subtotal mano de obra"},
        {"code":"%EPP", "concept":"Equipo de protección personal declarado como % sobre subtotal MO", "unit":"%", "section":"% SOBRE MO", "pu":"=SUM(H11:H13)", "op":"%", "qty":epp_rate, "amount":"=E15*G15", "pct":"=H15/$H$25", "mpu":"=SUM(M11:M13)", "mop":"%", "mqty":epp_rate, "mamount":"=J15*L15", "dev":"=IF(J15=0,0,E15/J15-1)", "state":"OK", "obs":"El % aplica solo sobre subtotal mano de obra"},
        {"code":"", "concept":"SUBTOTAL MANO DE OBRA", "unit":"", "section":"SUBTOTAL MO", "pu":"=SUM(H11:H13)", "op":"", "qty":"", "amount":"=SUM(H11:H15)", "pct":"=H16/$H$25", "mpu":"=SUM(M11:M13)", "mop":"", "mqty":"", "mamount":"=SUM(M11:M15)", "dev":"=IF(M16=0,0,H16/M16-1)", "state":"", "obs":"Incluye MO + % aplicables a MO"},
        {"code":"", "concept":"MAQUINARIA / EQUIPO", "unit":"", "section":"TÍTULO", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":""},
        {"code":"EQ-MEZ", "concept":"Equipo menor / maquinaria auxiliar", "unit":"HR", "section":"MAQUINARIA", "pu":350*factor, "op":"*", "qty":2, "amount":"=E18*G18", "pct":"=H18/$H$25", "mpu":330, "mop":"*", "mqty":2, "mamount":"=J18*L18", "dev":"=IF(J18=0,0,E18/J18-1)", "state":"OK", "obs":"Referencia granular: maquinaria data"},
        {"code":"%EQ", "concept":"Porcentaje sobre subtotal de MAQUINARIA", "unit":"%", "section":"% SOBRE MAQUINARIA", "pu":"=H18", "op":"%", "qty":machine_rate, "amount":"=E19*G19", "pct":"=H19/$H$25", "mpu":"=M18", "mop":"%", "mqty":machine_rate, "mamount":"=J19*L19", "dev":"=IF(J19=0,0,E19/J19-1)", "state":"OK", "obs":"El % aplica solo sobre subtotal maquinaria"},
        {"code":"", "concept":"SUBTOTAL MAQUINARIA", "unit":"", "section":"SUBTOTAL MAQUINARIA", "pu":"=SUM(H18:H19)", "op":"", "qty":"", "amount":"=SUM(H18:H19)", "pct":"=H20/$H$25", "mpu":"=SUM(M18:M19)", "mop":"", "mqty":"", "mamount":"=SUM(M18:M19)", "dev":"=IF(M20=0,0,H20/M20-1)", "state":"", "obs":"Incluye maquinaria + % aplicables a maquinaria"},
        {"code":"", "concept":"SECCIÓN FINANCIERA", "unit":"", "section":"TÍTULO", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"", "obs":""},
        {"code":"", "concept":"COSTO DIRECTO", "unit":"", "section":"TOTAL DIRECTO", "pu":"=H9+H16+H20", "op":"", "qty":"", "amount":"=E22", "pct":"=H22/$H$25", "mpu":"=M9+M16+M20", "mop":"", "mqty":"", "mamount":"=J22", "dev":"=IF(M22=0,0,H22/M22-1)", "state":"", "obs":"Materiales + MO + maquinaria"},
        {"code":"IND", "concept":"Costo indirecto declarado como % sobre COSTO DIRECTO", "unit":"%", "section":"% SOBRE DIRECTO", "pu":"=H22", "op":"%", "qty":indirect_rate, "amount":"=E23*G23", "pct":"=H23/$H$25", "mpu":"=M22", "mop":"%", "mqty":indirect_rate, "mamount":"=J23*L23", "dev":"=IF(J23=0,0,E23/J23-1)", "state":"OK", "obs":"Porcentaje financiero declarado por la fuente"},
        {"code":"FIN", "concept":"Financiamiento declarado como % sobre directo + indirecto", "unit":"%", "section":"% SOBRE DIRECTO+IND", "pu":"=H22+H23", "op":"%", "qty":finance_rate, "amount":"=E24*G24", "pct":"=H24/$H$25", "mpu":"=M22+M23", "mop":"%", "mqty":finance_rate, "mamount":"=J24*L24", "dev":"=IF(J24=0,0,E24/J24-1)", "state":"OK", "obs":"Base = costo directo + indirecto"},
        {"code":"", "concept":"TOTAL COSTO UNITARIO", "unit":"", "section":"TOTAL", "pu":"=H22+H23+H24", "op":"", "qty":"", "amount":"=E25", "pct":1, "mpu":"=M22+M23+M24", "mop":"", "mqty":"", "mamount":"=J25", "dev":"=IF(M25=0,0,H25/M25-1)", "state":"", "obs":"Este total debe reconciliar con Comparativa"},
    ]


def _strip_market_formula(value):
    """Convert a market-based formula to blank when rendering base-budget detail."""
    if isinstance(value, str) and any(token in value for token in ["J", "L", "M"]):
        return ""
    return value


def _write_canonical_apu_detail_sheet(ws, display_name: str, theme_color: str, *, variant: str = "A", include_market: bool = True, source_type: str = "CONTRACTOR_APU"):
    """Write a canonical APU detail sheet.

    This is the single shared layout for details generated by:
    - BaseBudgetEngine: include_market=False, source_type='BASE_BUDGET'
    - ContractorMatrixDetailEngine: include_market=True, source_type='CONTRACTOR_APU'

    Reusing this routine prevents base-budget and comparison details from
    drifting apart. The base detail excludes only market columns because its
    values already represent the generated market/base matrix.
    """
    headers = _canonical_detail_headers(include_market)
    max_col = len(headers)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    base_widths = [14, 42, 12, 18, 14, 11, 12, 16, 12]
    market_widths = [16, 11, 12, 16, 14] if include_market else []
    tail_widths = [18, 32]
    for idx, width in enumerate(base_widths + market_widths + tail_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    title = f"Detalle APU - {display_name}" if source_type == "CONTRACTOR_APU" else "Detalle Base - Matriz presupuestada"
    _provider_group_header(ws, 1, 1, max_col, title, theme_color)
    for c, h in enumerate(headers, 1):
        ws.cell(2, c, h)
    _header_style(ws, 2, 1, max_col, fill=theme_color)

    keys = ["code", "concept", "unit", "section", "pu", "op", "qty", "amount", "pct"]
    if include_market:
        keys += ["mpu", "mop", "mqty", "mamount", "dev"]
    keys += ["state", "obs"]

    for r, row in enumerate(_canonical_apu_rows(variant), 3):
        for c, key in enumerate(keys, 1):
            value = row.get(key, "")
            if not include_market and key in {"state", "obs"}:
                # Keep regular values.
                pass
            elif not include_market:
                value = _strip_market_formula(value)
            ws.cell(r, c, value)
        section = row["section"]
        fill = "FFFFFF"; bold = False
        if section in {"PARTIDA", "TÍTULO"}:
            fill = "EAF2FF" if section == "PARTIDA" else "F2F4F7"; bold = True
        elif section.startswith("SUBTOTAL") or section in {"TOTAL", "TOTAL DIRECTO"}:
            fill = "FFF2CC"; bold = True
        elif section.startswith("%"):
            fill = "F8FAFC"
        for c in range(1, max_col + 1):
            ws.cell(r, c).fill = PatternFill("solid", fgColor=fill)
            ws.cell(r, c).border = _thin_border("EAECF0")
            ws.cell(r, c).alignment = Alignment(vertical="top", wrap_text=True)
            if bold:
                ws.cell(r, c).font = Font(bold=True, color=BRAND["text"])

    # Number formats. Base columns are stable; market columns only exist when requested.
    money_cols = [5, 8]
    pct_cols = [9]
    if include_market:
        money_cols += [10, 13]
        pct_cols += [14]
    _apply_formats(ws, money_cols=money_cols, pct_cols=pct_cols, start_row=3, end_row=25)
    for col in [7, 12] if include_market else [7]:
        for row in range(3, 26):
            if ws.cell(row, col-1).value == "%":
                ws.cell(row, col).number_format = PCT_FMT
    ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}25"
    if include_market:
        dev_col = get_column_letter(14)
        ws.conditional_formatting.add(f"{dev_col}3:{dev_col}25", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    for r in range(5, 9): ws.row_dimensions[r].outlineLevel = 1
    for r in range(11, 16): ws.row_dimensions[r].outlineLevel = 1
    for r in range(18, 20): ws.row_dimensions[r].outlineLevel = 1
    for r in range(23, 25): ws.row_dimensions[r].outlineLevel = 1

    note_row = 28
    ws.cell(note_row, 1, "Regla canónica de generación")
    ws.cell(note_row, 1).font = Font(bold=True, color=BRAND["navy"])
    if include_market:
        note = "Esta hoja se genera desde el archivo matriz/APU propio del proveedor seleccionado y compara contra referencias granulares de data. No usa construdata_matrices.xlsx."
    else:
        note = "Esta hoja usa el mismo layout canónico del Detalle APU de comparativa, pero excluye columnas de mercado porque representa la matriz base/de mercado generada desde conceptos de ingeniería + matrices Construdata."
    ws.cell(note_row, 2, note)
    ws.merge_cells(start_row=note_row, start_column=2, end_row=note_row, end_column=max_col)
    ws.cell(note_row, 2).alignment = Alignment(wrap_text=True)


def _write_contractor_detail_sheet(ws, contractor_name: str, theme_color: str, variant: str = "A"):
    _write_canonical_apu_detail_sheet(ws, contractor_name, theme_color, variant=variant, include_market=True, source_type="CONTRACTOR_APU")

def _write_partidas_criticas(ws):
    _setup_sheet(ws, "Partidas Críticas", "Priorización de partidas por impacto económico, desviación y riesgo de negociación.", 12)
    _set_widths(ws, {"A": 10, "B": 16, "C": 16, "D": 36, "E": 18, "F": 16, "G": 14, "H": 14, "I": 18, "J": 26, "K": 22, "L": 18})
    headers = ["Prioridad", "Proveedor", "Partida", "Descripción", "Familia", "Impacto", "Desv. %", "Peso %", "Motivo", "Acción sugerida", "Estado revisión", "Responsable"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, len(headers))
    rows = [
        [1, "Proveedor C", "CIV-001", "Concreto f'c 250", "Concreto", 95000, 0.24, 0.18, "Alto impacto", "Solicitar desglose APU", "Pendiente", "Ingeniería"],
        [2, "Proveedor A", "CIV-002", "Acero de refuerzo", "Acero", 62000, 0.20, 0.12, "Precio sobre referencia", "Negociar precio unitario", "Pendiente", "Compras"],
        [3, "Proveedor C", "ELE-001", "Luminarias LED", "Instalaciones", 48000, 0.22, 0.09, "Ficha técnica no validada", "Solicitar aclaración", "Pendiente", "Ingeniería"],
        [4, "Proveedor B", "GEN-001", "Partida sin referencia directa", "Especial", 15000000, 0, 0.20, "Sin referencia", "Cotización externa", "Pendiente", "Costos"],
    ]
    for r, row in enumerate(rows, 6):
        for c, v in enumerate(row, 1): ws.cell(r, c, v)
    _body_style(ws, 6, 9, 1, len(headers))
    _apply_formats(ws, money_cols=[6], pct_cols=[7, 8], start_row=6, end_row=9)
    _add_table(ws, "A5:L9", "PartidasCriticasTable", "TableStyleMedium2")
    ws.freeze_panes = "A6"
    ws.conditional_formatting.add("F6:F9", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5", showValue=True))
    ws.conditional_formatting.add("G6:G9", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    dv = DataValidation(type="list", formula1='"Pendiente,Revisado,Aceptado,Solicitar aclaración,Negociar"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("K6:K50")


def _write_insumos_criticos(ws):
    _setup_sheet(ws, "Insumos Críticos", "Insumos, mano de obra, maquinaria y porcentajes que explican variaciones relevantes.", 13)
    _set_widths(ws, {"A": 16, "B": 16, "C": 18, "D": 28, "E": 10, "F": 14, "G": 16, "H": 16, "I": 14, "J": 16, "K": 22, "L": 20, "M": 28})
    headers = ["Contratista", "Partida", "Tipo insumo", "Insumo", "Unidad", "Cantidad", "P.U. contratista", "P.U. referencia", "Desv. %", "Impacto", "Fuente data", "Alerta", "Observación"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, len(headers))
    rows = [
        ["Contratista A", "FLEX41.11", "MO", "Supervisor Obra", "JOR", 0.10, 2643.19, 1757.33, "=(G6-H6)/H6", "=MAX((G6-H6)*F6,0)", "data/mano_obra", "Revisar", "Tarifa superior a referencia"],
        ["Contratista A", "FLEX41.11", "MO", "Oficial soldador/argonero", "JOR", 1, 2120.45, 1326.14, "=(G7-H7)/H7", "=MAX((G7-H7)*F7,0)", "data/mano_obra", "Revisar", "Validar rendimiento"],
        ["Contratista C", "CIV-001", "Material", "Concreto premezclado", "m3", 40, 2850000, 2300000, "=(G8-H8)/H8", "=MAX((G8-H8)*F8,0)", "data/materiales", "Alto impacto", "Negociar precio base"],
        ["Contratista C", "ELE-001", "Material", "Luminaria LED", "pza", 35, 1200000, 980000, "=(G9-H9)/H9", "=MAX((G9-H9)*F9,0)", "data/materiales", "Ficha técnica", "Confirmar especificación"],
    ]
    for r, row in enumerate(rows, 6):
        for c, v in enumerate(row, 1): ws.cell(r, c, v)
    _body_style(ws, 6, 9, 1, len(headers))
    _apply_formats(ws, money_cols=[7, 8, 10], pct_cols=[9], start_row=6, end_row=9)
    _add_table(ws, "A5:M9", "InsumosCriticosTable", "TableStyleMedium2")
    ws.conditional_formatting.add("I6:I9", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    ws.conditional_formatting.add("J6:J9", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5", showValue=True))
    ws.freeze_panes = "A6"


def _write_validaciones(ws):
    _setup_sheet(ws, "Validaciones", "Bitácora de advertencias, inconsistencias y puntos que requieren revisión técnica.", 10)
    _set_widths(ws, {"A": 14, "B": 24, "C": 16, "D": 16, "E": 18, "F": 24, "G": 36, "H": 26, "I": 28, "J": 18})
    headers = ["Severidad", "Tipo", "Contratista", "Partida", "Campo", "Valor detectado", "Problema", "Impacto potencial", "Acción sugerida", "Estado"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, len(headers))
    rows = [
        ["Alta", "Total no cuadra", "Contratista A", "FLEX41.11", "P.U.", "31524.55", "El total calculado debe reconciliar con Comparativa", "Puede afectar ranking", "Revisar matriz APU", "Pendiente"],
        ["Media", "Unidad dudosa", "Contratista C", "ELE-001", "Unidad", "pza", "La especificación técnica de luminaria no es suficiente", "Precio no comparable", "Solicitar ficha técnica", "Pendiente"],
        ["Media", "Insumo no encontrado", "Contratista B", "GEN-001", "Referencia", "N/A", "No existe referencia granular en data", "Requiere cotización externa", "Solicitar soporte", "Pendiente"],
        ["Baja", "Match medio", "Contratista A", "CIV-002", "Insumo", "Acero", "Match por descripción, no por código", "Revisión documental", "Validar código", "Pendiente"],
    ]
    for r, row in enumerate(rows, 6):
        for c, v in enumerate(row, 1): ws.cell(r, c, v)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=_status_fill(row[0]))
    _body_style(ws, 6, 9, 1, len(headers))
    _add_table(ws, "A5:J9", "ValidacionesTable", "TableStyleMedium2")
    dv = DataValidation(type="list", formula1='"Pendiente,Revisado,Resuelto,No aplica"')
    ws.add_data_validation(dv)
    dv.add("J6:J100")
    ws.freeze_panes = "A6"


def _write_parametros(ws):
    _setup_sheet(ws, "Parámetros", "Trazabilidad del reporte: fuentes, reglas, versiones y limitaciones metodológicas.", 8)
    _set_widths(ws, {"A": 28, "B": 70, "C": 20, "D": 22, "E": 22, "F": 22, "G": 22, "H": 22})
    params = [
        ["Versión sistema", "V0 canónica", ""],
        ["Fecha generación", datetime.now().strftime("%Y-%m-%d %H:%M"), ""],
        ["Motor Excel", "ExcelReportEngineV1 - versión profesional inicial", ""],
        ["Presupuesto base", "Usa data/construdata_matrices.xlsx solamente para módulo independiente", "Regla canónica"],
        ["Detalle APU", "Usa matriz/APU del contratista + referencias granulares de data", "Regla canónica"],
        ["Materiales", "data/construdata-materiales-052026.xlsx", "Referencia granular"],
        ["Mano de obra", "data/construdata-manodeobra-052026*.xlsx", "Referencia granular"],
        ["Maquinaria", "data/construdata-maquinaria-052026.xlsx", "Referencia granular"],
        ["IA", "Interpreta resultados calculados; no calcula montos", "Regla canónica"],
        ["Limitación V0", "Datos mockeados para validar diseño y estructura. V1 conectará lectura real de matrices/APU.", "Importante"],
    ]
    headers = ["Parámetro", "Valor", "Tipo"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, 3)
    for r, row in enumerate(params, 6):
        for c, v in enumerate(row, 1): ws.cell(r, c, v)
    _body_style(ws, 6, 15, 1, 3)
    _add_table(ws, "A5:C15", "ParametrosTable", "TableStyleMedium2")


def _short_codes(items: list[Any], limit: int = 5) -> str:
    codes = []
    for x in items:
        code = str(getattr(x, "code", "") or "").strip()
        if code and code not in codes:
            codes.append(code)
        if len(codes) >= limit:
            break
    return ", ".join(codes) if codes else "—"


def _safe_amount(obj: Any) -> float:
    amount = getattr(obj, "amount", None)
    if amount is not None:
        try:
            return float(amount or 0)
        except Exception:
            return 0.0
    qty = getattr(obj, "quantity", None)
    pu = getattr(obj, "unit_price", None)
    if qty is not None and pu is not None:
        try:
            return float(qty or 0) * float(pu or 0)
        except Exception:
            return 0.0
    return 0.0



def _money_text(value: Any) -> str:
    try:
        return f"${float(value or 0):,.2f}"
    except Exception:
        return "$0.00"


def _pct_text(value: Any) -> str:
    try:
        return f"{float(value or 0):.1f}%"
    except Exception:
        return "0.0%"


def _short_desc(text: str, limit: int = 72) -> str:
    t = " ".join(str(text or "").split())
    if len(t) <= limit:
        return t
    return t[:limit - 1].rstrip() + "…"


def _apu_section_name(section: str) -> str:
    s = str(section or "").upper()
    if "MATERIAL" in s:
        return "Materiales"
    if "MANO" in s or s in {"MO", "SUBTOTAL MO"}:
        return "Mano de obra"
    if "MAQUIN" in s or "EQUIPO" in s or "HERRAM" in s:
        return "Maquinaria / equipo"
    if "INDIRECT" in s:
        return "Indirectos"
    if "BASIC" in s or "BÁSIC" in s:
        return "Básicos"
    if "DIRECT" in s:
        return "Costo directo"
    return section or "Sin sección"


def _item_amount_value(item: Any, market: bool = False) -> float:
    val = getattr(item, "market_amount", None) if market else getattr(item, "amount", None)
    try:
        return float(val or 0)
    except Exception:
        return 0.0


def _sum_scaled_items(items: list[Any], qty_by_key: dict[str, float], sections: set[str] | None = None, market: bool = False) -> float:
    total = 0.0
    for item in items or []:
        sec = str(getattr(item, "section", "") or "").upper()
        if sections and sec not in sections:
            continue
        qty = float(qty_by_key.get(getattr(item, "concept_key", ""), 1) or 1)
        total += _item_amount_value(item, market=market) * qty
    return total


def _top_apu_overcost_items(items: list[Any], qty_by_key: dict[str, float], limit: int = 8) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for it in items or []:
        if getattr(it, "market_amount", None) is None or getattr(it, "amount", None) is None:
            continue
        amount = float(getattr(it, "amount", 0) or 0)
        market = float(getattr(it, "market_amount", 0) or 0)
        delta = amount - market
        if delta <= 0:
            continue
        qty = float(qty_by_key.get(getattr(it, "concept_key", ""), 1) or 1)
        rows.append({
            "code": getattr(it, "code", "") or "",
            "description": _short_desc(getattr(it, "description", "") or "", 90),
            "section": _apu_section_name(getattr(it, "section", "") or ""),
            "amount": round(amount * qty, 2),
            "market_amount": round(market * qty, 2),
            "delta": round(delta * qty, 2),
            "delta_pct": round((delta / market * 100) if market else 0, 2),
            "match": _short_desc(((getattr(it, "matched_reference_code", "") or "") + " - " + (getattr(it, "matched_reference_description", "") or "")).strip(" -"), 100),
        })
    return sorted(rows, key=lambda r: r["delta"], reverse=True)[:limit]


def _base_section_breakdown_sequential(items: list[Any]) -> dict[str, float]:
    totals = {"materials": 0.0, "labor": 0.0, "equipment": 0.0, "basics": 0.0, "direct": 0.0, "indirect": 0.0, "total_service": 0.0}
    current_qty = 1.0
    for item in items or []:
        sec = str(getattr(item, "section", "") or "").upper()
        desc = str(getattr(item, "description", "") or "").upper()
        if sec == "PARTIDA":
            try:
                current_qty = float(getattr(item, "quantity", None) or 1)
            except Exception:
                current_qty = 1.0
            continue
        amount = _item_amount_value(item, market=False)
        scaled = amount * current_qty
        if sec == "SUBTOTAL MATERIALES":
            totals["materials"] += scaled
        elif sec in {"SUBTOTAL MANO DE OBRA", "SUBTOTAL MO"}:
            totals["labor"] += scaled
        elif sec in {"SUBTOTAL MAQUINARIA", "SUBTOTAL EQUIPO", "SUBTOTAL EQUIPO Y HERRAMIENTA"}:
            totals["equipment"] += scaled
        elif sec in {"SUBTOTAL BASICOS", "SUBTOTAL BÁSICOS"}:
            totals["basics"] += scaled
        elif sec == "COSTO DIRECTO":
            totals["direct"] += scaled
        elif sec in {"INDIRECTO", "COSTO INDIRECTO"}:
            totals["indirect"] += scaled
        elif sec == "TOTAL POR SERVICIO":
            totals["total_service"] += amount
    return {k: round(v, 2) for k, v in totals.items()}


def _section_breakdown_for_run(run: CanonicalRun) -> dict[str, Any]:
    if run.kind == "base":
        return _base_section_breakdown_sequential(run.base_apu_items)
    return {}


def _provider_section_breakdown(provider: CanonicalProvider) -> dict[str, Any]:
    qty_by_key = {canonical_key(c.code, c.description): float(c.quantity or 0) for c in provider.concepts if _is_valid_comparativa_concept(c)}
    items = provider.apu_items or []
    contractor = {
        "materials": _sum_scaled_items(items, qty_by_key, {"SUBTOTAL MATERIALES"}, market=False),
        "labor": _sum_scaled_items(items, qty_by_key, {"SUBTOTAL MANO DE OBRA", "SUBTOTAL MO"}, market=False),
        "equipment": _sum_scaled_items(items, qty_by_key, {"SUBTOTAL MAQUINARIA", "SUBTOTAL EQUIPO", "SUBTOTAL EQUIPO Y HERRAMIENTA"}, market=False),
        "direct": _sum_scaled_items(items, qty_by_key, {"COSTO DIRECTO"}, market=False),
        "indirect": _sum_scaled_items(items, qty_by_key, {"INDIRECTO", "COSTO INDIRECTO"}, market=False),
    }
    market = {
        "materials": _sum_scaled_items(items, qty_by_key, {"SUBTOTAL MATERIALES"}, market=True),
        "labor": _sum_scaled_items(items, qty_by_key, {"SUBTOTAL MANO DE OBRA", "SUBTOTAL MO"}, market=True),
        "equipment": _sum_scaled_items(items, qty_by_key, {"SUBTOTAL MAQUINARIA", "SUBTOTAL EQUIPO", "SUBTOTAL EQUIPO Y HERRAMIENTA"}, market=True),
        "direct": _sum_scaled_items(items, qty_by_key, {"COSTO DIRECTO"}, market=True),
        "indirect": _sum_scaled_items(items, qty_by_key, {"INDIRECTO", "COSTO INDIRECTO"}, market=True),
    }
    return {"contractor": {k: round(v,2) for k,v in contractor.items()}, "market": {k: round(v,2) for k,v in market.items()}}


def _analysis_feature_status() -> dict[str, Any]:
    if not AI_ANALYSIS_ENABLED:
        return {"enabled": False, "mode": "DISABLED_LOCAL_EXPERT", "provider": "local", "model": "deterministic", "hasKey": False, "lastError": None}
    has_key = False
    if AI_ANALYSIS_PROVIDER == "anthropic":
        has_key = bool(os.getenv("ANTHROPIC_API_KEY"))
    elif AI_ANALYSIS_PROVIDER == "openai":
        has_key = bool(os.getenv("OPENAI_API_KEY"))
    return {
        "enabled": True,
        "mode": "EXTERNAL_AI" if has_key else "LOCAL_EXPERT_FALLBACK_NO_KEY",
        "provider": AI_ANALYSIS_PROVIDER,
        "model": _configured_ai_model(),
        "hasKey": has_key,
        "timeoutSeconds": AI_ANALYSIS_TIMEOUT,
        "maxInputItems": AI_ANALYSIS_MAX_INPUT_ITEMS,
        "lastError": AI_ANALYSIS_LAST_ERROR,
        "lastProviderResponsePreview": AI_ANALYSIS_LAST_PROVIDER_RESPONSE,
    }


def _run_analysis_context(run: CanonicalRun | None, summary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Build a compact, auditable context for AI/expert interpretation.

    The context intentionally contains only values already calculated by the
    canonical run. The AI layer may interpret these fields, but it must not
    calculate new amounts, modify prices or create facts outside this JSON.
    """
    if not run:
        return {"kind": "unknown", "available": False}

    if run.kind == "base":
        exec_concepts = [c for c in run.base_concepts if getattr(c, "is_executable", False)]
        total = sum(_safe_amount(c) for c in exec_concepts)
        validations = run.validations or []
        matched = len([c for c in exec_concepts if str(getattr(c, "market_state", "")).startswith("Match")])
        unmatched = max(0, len(exec_concepts) - matched)
        qty_by_key = {canonical_key(c.code, c.description): float(c.quantity or 0) for c in exec_concepts}
        section_breakdown = _section_breakdown_for_run(run)
        direct = float(section_breakdown.get("direct", 0) or 0)
        indirect = float(section_breakdown.get("indirect", 0) or 0)
        if not direct:
            direct = max(0.0, total / 1.25) if total else 0.0
        if not indirect and direct:
            indirect = direct * 0.25
        top = []
        for c in sorted(exec_concepts, key=_safe_amount, reverse=True)[:AI_ANALYSIS_MAX_INPUT_ITEMS]:
            amount = _safe_amount(c)
            top.append({
                "code": c.code or "",
                "description": _short_desc(c.description, 90),
                "unit": c.unit or "",
                "quantity": float(c.quantity or 0),
                "unit_price": round(float(c.unit_price or 0), 2),
                "amount": round(amount, 2),
                "weight_pct": round((amount / total * 100) if total else 0, 2),
                "state": c.market_state or "",
            })
        severity_counts: dict[str, int] = {}
        for v in validations:
            sev = str(v.get("severity", "Info") or "Info")
            severity_counts[sev] = severity_counts.get(sev, 0) + 1
        validation_samples = []
        for v in validations[:AI_ANALYSIS_MAX_INPUT_ITEMS]:
            validation_samples.append({
                "severity": v.get("severity", "Info"),
                "type": v.get("type", ""),
                "message": _short_desc(str(v.get("message", "")), 130),
            })
        return {
            "kind": "base_budget",
            "project_name": run.project_name,
            "run_id": run.run_id,
            "total_amount": round(total, 2),
            "direct_cost": round(direct, 2),
            "indirect_cost": round(indirect, 2),
            "indirect_pct": 25.0,
            "section_breakdown": {k: round(v, 2) for k, v in section_breakdown.items()},
            "concepts_read": len(run.base_concepts),
            "concepts_executable": len(exec_concepts),
            "detail_rows": len(run.base_apu_items),
            "matched_matrices": matched,
            "unmatched_matrices": unmatched,
            "coverage_pct": round((matched / len(exec_concepts) * 100) if exec_concepts else 0, 2),
            "validations_count": len(validations),
            "validation_severity_counts": severity_counts,
            "validation_samples": validation_samples,
            "top_concepts": top,
            "top_overcost_items": _top_apu_overcost_items(run.base_apu_items, qty_by_key, AI_ANALYSIS_MAX_INPUT_ITEMS),
            "summary": summary or {},
        }

    providers_ctx = []
    totals = []
    market_totals = []
    for p in run.providers or []:
        valid_concepts = [c for c in p.concepts if _is_valid_comparativa_concept(c)]
        amount = _canonical_amount_from_concepts(valid_concepts)
        market_amount = sum(float(getattr(c, "market_amount", 0) or 0) for c in valid_concepts)
        if not market_amount:
            market_amount = 0.0
        totals.append(amount)
        market_totals.append(market_amount)
        no_ref = len([i for i in p.apu_items if str(i.state).upper().startswith("SIN REFERENCIA") or str(i.state).upper() == "SIN REFERENCIA"])
        fallback_items = len([i for i in p.apu_items if getattr(i, "market_unit_price_is_fallback", False) or getattr(i, "market_amount_is_fallback", False)])
        with_reference = len([i for i in p.apu_items if getattr(i, "matched_reference_code", "") or getattr(i, "matched_reference_description", "")])
        top = []
        for c in sorted(valid_concepts, key=_concept_amount, reverse=True)[:AI_ANALYSIS_MAX_INPUT_ITEMS]:
            c_amount = _concept_amount(c)
            m_amount = float(getattr(c, "market_amount", 0) or 0)
            delta = c_amount - m_amount if m_amount else 0
            top.append({
                "code": c.code or "",
                "description": _short_desc(c.description, 90),
                "unit": c.unit or "",
                "quantity": float(c.quantity or 0),
                "unit_price": round(float(c.unit_price or 0), 2),
                "amount": round(c_amount, 2),
                "market_amount": round(m_amount, 2),
                "overcost": round(delta, 2),
                "overcost_pct": round((delta / m_amount * 100) if m_amount else 0, 2),
                "weight_pct": round((c_amount / amount * 100) if amount else 0, 2),
                "state": getattr(c, "market_state", "") or "",
            })
        section_breakdown = _provider_section_breakdown(p)
        qty_by_key = {canonical_key(c.code, c.description): float(c.quantity or 0) for c in valid_concepts}
        providers_ctx.append({
            "name": p.name,
            "total_amount": round(amount, 2),
            "market_total_amount": round(market_amount, 2),
            "overcost_amount": round(amount - market_amount, 2) if market_amount else 0,
            "overcost_pct": round(((amount / market_amount) - 1) * 100, 2) if market_amount else 0,
            "concepts_executable": len(valid_concepts),
            "apu_rows": len(p.apu_items),
            "reference_items": with_reference,
            "fallback_items": fallback_items,
            "no_reference_items": no_ref,
            "section_breakdown": section_breakdown,
            "top_concepts": top,
            "top_overcost_items": _top_apu_overcost_items(p.apu_items, qty_by_key, AI_ANALYSIS_MAX_INPUT_ITEMS),
            "validations_count": len(p.validations or []),
        })
    best_idx = min(range(len(totals)), key=lambda i: totals[i]) if totals else None
    worst_idx = max(range(len(totals)), key=lambda i: totals[i]) if totals else None
    spread = ((totals[worst_idx] / totals[best_idx]) - 1) if best_idx is not None and worst_idx is not None and totals[best_idx] else 0
    return {
        "kind": "comparison",
        "project_name": run.project_name,
        "run_id": run.run_id,
        "providers_count": len(run.providers or []),
        "providers": providers_ctx,
        "best_provider": providers_ctx[best_idx]["name"] if best_idx is not None else "",
        "worst_provider": providers_ctx[worst_idx]["name"] if worst_idx is not None else "",
        "economic_spread_pct": round(spread * 100, 2),
        "min_amount": round(min(totals), 2) if totals else 0,
        "max_amount": round(max(totals), 2) if totals else 0,
        "total_no_reference_items": sum(int(p.get("no_reference_items", 0) or 0) for p in providers_ctx),
        "total_fallback_items": sum(int(p.get("fallback_items", 0) or 0) for p in providers_ctx),
        "summary": summary or {},
    }


def _top_list_text(rows: list[dict[str, Any]], limit: int = 5, *, amount_key: str = "amount", pct_key: str = "weight_pct") -> str:
    parts: list[str] = []
    for row in (rows or [])[:limit]:
        code = str(row.get("code") or "").strip() or _short_desc(str(row.get("description", "")), 18)
        desc = _short_desc(str(row.get("description", "")), 34)
        amount = _money_text(row.get(amount_key, 0))
        pct = _pct_text(row.get(pct_key, 0)) if row.get(pct_key) is not None else ""
        label = f"{code}"
        if desc and desc.upper() != code.upper():
            label += f" ({desc})"
        parts.append(f"{label}: {amount}" + (f" · {pct}" if pct else ""))
    return "; ".join(parts) or "sin datos priorizables calculados"


def _section_mix_text(values: dict[str, Any], total: float | None = None) -> str:
    total = float(total or 0)
    order = [("materials", "Materiales"), ("labor", "MO"), ("equipment", "Maquinaria/equipo"), ("basics", "Básicos"), ("direct", "Costo directo"), ("indirect", "Indirectos")]
    chunks: list[str] = []
    for key, label in order:
        val = float((values or {}).get(key, 0) or 0)
        if val == 0:
            continue
        pct = f" ({_pct_text(val/total*100)})" if total else ""
        chunks.append(f"{label}: {_money_text(val)}{pct}")
    return "; ".join(chunks) or "sin desglose por sección disponible"


def _market_alert_text(rows: list[dict[str, Any]], limit: int = 5) -> str:
    alerts: list[str] = []
    for row in (rows or [])[:limit]:
        code = str(row.get("code") or "").strip() or _short_desc(str(row.get("description", "")), 20)
        section = str(row.get("section") or "Sin sección")
        delta = float(row.get("delta", 0) or 0)
        pct = float(row.get("delta_pct", 0) or 0)
        match = _short_desc(str(row.get("match") or ""), 45)
        suffix = f" contra {match}" if match else " contra mercado"
        alerts.append(f"{code} · {section}: sobrecosto {_money_text(delta)} ({_pct_text(pct)}){suffix}")
    return "; ".join(alerts) or "no se identificaron sobrecostos monetarios contra referencias de mercado en los insumos priorizados"


def _risk_level_from_values(coverage: float = 0, unmatched: int = 0, fallback: int = 0, no_ref: int = 0, spread: float = 0, overcost_pct: float = 0) -> str:
    if coverage and coverage < 70:
        return "alto"
    if unmatched >= 5 or no_ref >= 20 or fallback >= 30 or spread >= 20 or overcost_pct >= 20:
        return "alto"
    if unmatched > 0 or no_ref > 0 or fallback > 0 or spread >= 8 or overcost_pct >= 8 or (coverage and coverage < 90):
        return "medio"
    return "bajo"


def _local_expert_analysis(context: dict[str, Any]) -> list[dict[str, str]]:
    """Deterministic expert narrative with numbers, priorities and market alerts.

    This fallback is intentionally strict: it only uses calculated canonical data,
    but it should read like a professional APU review, not a generic executive
    paragraph. The external AI receives the same structure and is asked to match
    this level of specificity.
    """
    kind = context.get("kind")
    if kind == "base_budget":
        top = context.get("top_concepts") or []
        total = float(context.get("total_amount", 0) or 0)
        direct = float(context.get("direct_cost", 0) or 0)
        indirect = float(context.get("indirect_cost", 0) or 0)
        coverage = float(context.get("coverage_pct", 0) or 0)
        unmatched = int(context.get("unmatched_matrices", 0) or 0)
        matched = int(context.get("matched_matrices", 0) or 0)
        detail_rows = int(context.get("detail_rows", 0) or 0)
        executable = int(context.get("concepts_executable", 0) or 0)
        read = int(context.get("concepts_read", 0) or 0)
        breakdown = context.get("section_breakdown") or {}
        over = context.get("top_overcost_items") or []
        risk = _risk_level_from_values(coverage=coverage, unmatched=unmatched)
        top_text = _top_list_text(top, 6)
        section_text = _section_mix_text(breakdown, total)
        alert_text = _market_alert_text(over, 5)
        top1 = top[0] if top else {}
        concentration = float(top1.get("weight_pct", 0) or 0)
        concentration_note = f"La mayor concentración está en {top1.get('code')} con {_money_text(top1.get('amount'))} ({_pct_text(concentration)} del total). " if top1 else ""
        review_note = f"{unmatched} conceptos quedan en revisión por no contar con matriz directa; deben validarse técnicamente antes de cerrar el presupuesto." if unmatched else "No quedan conceptos sin matriz directa dentro de los ejecutables calculados."
        return [
            {"section": "Resumen ejecutivo APU", "content": f"Presupuesto base de {context.get('project_name') or 'la corrida'} por {_money_text(total)}. Se leyeron {read} conceptos, {executable} fueron ejecutables y se generaron {detail_rows} filas APU. El costo directo es {_money_text(direct)} y el indirecto 25% equivale a {_money_text(indirect)}. Cobertura Construdata: {_pct_text(coverage)} ({matched} con match, {unmatched} en revisión)."},
            {"section": "Estructura del costo", "content": f"Desglose calculado: {section_text}. Esta separación permite ubicar si la presión económica está en materiales, MO, maquinaria/equipo o indirectos. El indirecto se mantiene fijo al 25%, por lo que el riesgo financiero se concentra en la correcta integración del costo directo."},
            {"section": "Concentración y partidas críticas", "content": f"{concentration_note}Top de impacto: {top_text}. La revisión debe iniciar por estas claves, porque concentran el presupuesto y cualquier ajuste de rendimiento, alcance o matriz modifica de forma material el resultado."},
            {"section": "Referencias de mercado y alertas", "content": f"{alert_text}. {review_note} Las filas sin referencia plena o sin referencia no deben considerarse validación plena de mercado; requieren soporte del analista o sustitución por una matriz Construdata más representativa."},
            {"section": "Riesgo económico", "content": f"Riesgo preliminar {risk}. La cobertura de {_pct_text(coverage)} permite usar el resultado como base de control, pero el cierre depende de validar las matrices asignadas a los conceptos críticos y documentar los casos en revisión."},
            {"section": "Acciones recomendadas", "content": "Primero validar las partidas críticas por monto; después revisar insumos con sobrecosto contra mercado y finalmente confirmar conceptos sin matriz directa. No negociar por porcentaje aislado: priorizar desviación monetaria, trazabilidad Construdata y consistencia técnica de rendimientos."},
        ]
    if kind == "comparison":
        providers = context.get("providers") or []
        best = context.get("best_provider") or "—"
        worst = context.get("worst_provider") or "—"
        spread = float(context.get("economic_spread_pct", 0) or 0)
        no_ref = int(context.get("total_no_reference_items", 0) or 0)
        fallback = int(context.get("total_fallback_items", 0) or 0)
        provider_lines = []
        over_lines = []
        section_lines = []
        top_concept_lines = []
        max_over_pct = 0.0
        for p in providers:
            name = p.get("name", "Proveedor")
            total = float(p.get("total_amount", 0) or 0)
            market = float(p.get("market_total_amount", 0) or 0)
            over = float(p.get("overcost_amount", 0) or 0)
            over_pct = float(p.get("overcost_pct", 0) or 0)
            max_over_pct = max(max_over_pct, over_pct)
            if market:
                sign = "sobrecosto" if over >= 0 else "ahorro"
                provider_lines.append(f"{name}: {_money_text(total)} vs mercado {_money_text(market)}; {sign} {_money_text(abs(over))} ({_pct_text(abs(over_pct))})")
            else:
                provider_lines.append(f"{name}: {_money_text(total)}; sin total mercado consolidado")
            tops = p.get("top_overcost_items") or []
            if tops:
                over_lines.append(f"{name}: {_market_alert_text(tops, 3)}")
            tc = p.get("top_concepts") or []
            if tc:
                top_concept_lines.append(f"{name}: {_top_list_text(tc, 4)}")
            sb = p.get("section_breakdown") or {}
            contractor = sb.get("contractor") or {}
            market_sb = sb.get("market") or {}
            if contractor:
                section_lines.append(f"{name} declarado: {_section_mix_text(contractor, total)}")
            if market_sb and any(float(v or 0) for v in market_sb.values()):
                market_total = float(p.get("market_total_amount", 0) or 0)
                section_lines.append(f"{name} mercado: {_section_mix_text(market_sb, market_total)}")
        mode_text = "análisis individual contra mercado" if len(providers) <= 1 else f"comparativa de {len(providers)} contratistas"
        risk = _risk_level_from_values(fallback=fallback, no_ref=no_ref, spread=spread, overcost_pct=max_over_pct)
        provider_text = "; ".join(provider_lines[:6]) or "sin proveedores válidos"
        over_text = " | ".join(over_lines[:4]) or "no se identificaron sobrecostos monetarios priorizados con referencia de mercado"
        section_text = " | ".join(section_lines[:6]) or "sin desglose por sección disponible"
        top_text = " | ".join(top_concept_lines[:4]) or "sin top de conceptos calculado"
        if len(providers) <= 1:
            p0 = providers[0] if providers else {}
            pname = p0.get("name", "Proveedor")
            total = float(p0.get("total_amount", 0) or 0)
            market = float(p0.get("market_total_amount", 0) or 0)
            over = float(p0.get("overcost_amount", 0) or 0)
            over_pct = float(p0.get("overcost_pct", 0) or 0)
            market_phrase = f" frente a mercado {_money_text(market)}, con diferencia {_money_text(over)} ({_pct_text(over_pct)})" if market else " sin total de mercado consolidado"
            return [
                {"section": "Resumen ejecutivo APU", "content": f"Análisis individual del proveedor {pname}: monto ofertado {_money_text(total)}{market_phrase}. Al existir un solo proveedor, la lectura correcta es competitividad contra mercado, trazabilidad de referencias y concentración del costo."},
                {"section": "Sobrecostos contra mercado", "content": f"Alertas priorizadas: {over_text}. La revisión debe enfocarse en desviación monetaria, trazabilidad de mercado y partidas de mayor impacto."},
                {"section": "Resumen por sección", "content": f"{section_text}. El diagnóstico separa Materiales, Mano de obra, Maquinaria/equipo e Indirectos para ubicar si la presión económica viene de insumos, rendimientos, equipos o estructura financiera."},
                {"section": "Partidas críticas", "content": f"Conceptos de mayor impacto: {top_text}. Estas partidas explican la mayor parte del monto y deben revisarse antes que diferencias pequeñas o aisladas."},
                {"section": "Referencias y trazabilidad", "content": f"Se detectan {no_ref} insumos sin referencia y {fallback} valores fallback. Todo valor sin referencia plena debe leerse como ausencia de validación Construdata, no como precio de mercado confirmado."},
                {"section": "Riesgo y acciones", "content": f"Riesgo preliminar {risk}. Validar matches Construdata en partidas críticas, revisar rendimiento de MO/maquinaria, confirmar indirectos y documentar insumos sin referencia antes de usar la propuesta como base de negociación."},
            ]
        return [
            {"section": "Resumen ejecutivo APU", "content": f"Corrida de {mode_text}. Mejor posición económica: {best}; mayor monto: {worst}; brecha entre extremos {_pct_text(spread)}. Totales evaluados: {provider_text}."},
            {"section": "Sobrecostos contra mercado", "content": f"Alertas por contratista: {over_text}. Estos hallazgos deben revisarse por desviación monetaria y no solo por porcentaje, porque las partidas de bajo importe pueden distorsionar la prioridad real."},
            {"section": "Resumen por sección", "content": f"{section_text}. Separar Materiales, MO, Maquinaria/equipo e Indirectos permite identificar si la diferencia viene de precios de insumos, rendimientos, equipos o estructura financiera."},
            {"section": "Partidas críticas", "content": f"Conceptos de mayor impacto: {top_text}. La negociación debe concentrarse en el 80% económico y en conceptos con sobrecosto frente a mercado, no en diferencias menores o aisladas."},
            {"section": "Referencias y trazabilidad", "content": f"Se detectan {no_ref} insumos sin referencia y {fallback} valores fallback. Cuando el mercado usa un valor sin referencia plena, el valor no representa validación Construdata; solo evita inventar un precio y debe quedar sujeto a revisión."},
            {"section": "Riesgo y acciones", "content": f"Riesgo preliminar {risk}. Revisar primero al contratista con mayor sobrecosto, validar matches Construdata de partidas críticas, solicitar soporte de rendimientos y separar negociación de Materiales, MO, Maquinaria/equipo e Indirectos."},
        ]
    return [
        {"section": "Resumen ejecutivo", "content": "No hay datos calculados suficientes para emitir un análisis profesional."},
        {"section": "Limitaciones", "content": "El análisis no genera importes ni completa datos ausentes."},
    ]


def _ai_system_prompt() -> str:
    return (
        "Eres un experto senior en análisis de precios unitarios (APU), Neodata y Construdata. "
        "Redacta en español natural, profesional y accionable para un analista de precios unitarios. "
        "Usa únicamente los datos del JSON proporcionado. Está prohibido inventar montos, porcentajes, contratistas, partidas, causas, matches o referencias. "
        "Tu análisis debe usar cifras concretas cuando existan: nombres de contratistas, monto ofertado, monto mercado, sobrecosto monetario y porcentual, cobertura Construdata, Materiales, Mano de obra, Maquinaria/equipo, Básicos e Indirectos. "
        "Señala alertas de mercado: valores sin referencia plena, sin referencia, conceptos sin matriz directa, sobrecostos por partida/insumo, partidas de mayor impacto y concentración 80/20. "
        "No copies nombres largos completos de servicios; usa códigos y descripciones cortas. No repitas la misma idea entre secciones. "
        "La IA no calcula ni corrige importes; solo interpreta valores ya calculados. Si un dato no está en el JSON, omítelo. "
        "Devuelve SOLO un objeto JSON válido. No uses markdown, no uses backticks, no agregues texto antes o después del JSON. "
        "La forma exacta debe ser: {\"sections\":[{\"section\":\"...\",\"content\":\"...\"}]}. "
        "Usa exactamente 6 secciones con estos nombres: Resumen ejecutivo APU, Sobrecostos contra mercado, Resumen por sección, Partidas críticas, Referencias y trazabilidad, Riesgo y acciones. "
        "Cada content debe tener entre 80 y 160 palabras, con cifras concretas si están disponibles. "
        "Evita frases genéricas como 'revisar partidas importantes' sin indicar cuáles, cuánto representan y por qué son relevantes."
    )


def _extract_json_object(text: str) -> dict[str, Any] | None:
    """Parse provider output even if it arrives wrapped in markdown or with prose.

    Anthropic and some OpenAI models may return ```json fences or explanatory
    text despite instructions. This parser prevents JSONDecodeError from killing
    the run and allows the fallback to take over only when extraction truly fails.
    """
    raw = (text or "").strip()
    if not raw:
        return None
    candidates = [raw]
    if "```" in raw:
        parts = raw.split("```")
        for part in parts:
            cleaned = part.strip()
            if cleaned.lower().startswith("json"):
                cleaned = cleaned[4:].strip()
            if cleaned.startswith("{") and cleaned.endswith("}"):
                candidates.insert(0, cleaned)
    first = raw.find("{")
    last = raw.rfind("}")
    if first >= 0 and last > first:
        candidates.append(raw[first:last + 1])
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            continue
    return None


def _normalize_ai_sections(parsed: dict[str, Any] | None) -> list[dict[str, str]]:
    if not isinstance(parsed, dict):
        return []
    sections = parsed.get("sections")
    if sections is None and isinstance(parsed.get("analysis"), list):
        sections = parsed.get("analysis")
    if sections is None and isinstance(parsed.get("items"), list):
        sections = parsed.get("items")
    if isinstance(sections, dict):
        sections = [{"section": k, "content": v} for k, v in sections.items()]
    if not isinstance(sections, list):
        return []
    clean: list[dict[str, str]] = []
    for item in sections[:6]:
        if not isinstance(item, dict):
            continue
        section = str(item.get("section") or item.get("title") or "").strip()[:80]
        content = str(item.get("content") or item.get("text") or item.get("body") or "").strip()
        if section and content:
            clean.append({"section": section, "content": content})
    return clean


def _call_external_ai_analysis(context: dict[str, Any]) -> list[dict[str, str]] | None:
    global AI_ANALYSIS_LAST_ERROR, AI_ANALYSIS_LAST_PROVIDER_RESPONSE
    AI_ANALYSIS_LAST_ERROR = None
    AI_ANALYSIS_LAST_PROVIDER_RESPONSE = None
    status = _analysis_feature_status()
    if not status.get("enabled") or status.get("mode") != "EXTERNAL_AI":
        return None
    payload_context = json.dumps(context, ensure_ascii=False, default=str)[:26000]
    model = _configured_ai_model()
    try:
        if AI_ANALYSIS_PROVIDER == "anthropic":
            key = os.getenv("ANTHROPIC_API_KEY")
            if not key:
                return None
            body = {
                "model": model,
                "max_tokens": 1800,
                "temperature": 0.05,
                "system": _ai_system_prompt(),
                "messages": [{"role": "user", "content": payload_context}],
            }
            req = urllib.request.Request(
                "https://api.anthropic.com/v1/messages",
                data=json.dumps(body).encode("utf-8"),
                headers={"content-type":"application/json", "x-api-key":key, "anthropic-version":"2023-06-01"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=AI_ANALYSIS_TIMEOUT) as resp:
                raw_response = resp.read().decode("utf-8", errors="replace")
            AI_ANALYSIS_LAST_PROVIDER_RESPONSE = raw_response[:4000]
            try:
                data = json.loads(raw_response)
            except Exception as exc:
                AI_ANALYSIS_LAST_ERROR = f"Anthropic returned non-JSON HTTP payload ({type(exc).__name__}). Preview: {raw_response[:500]}"
                return None
            chunks: list[str] = []
            for part in data.get("content", []) or []:
                if isinstance(part, dict):
                    if part.get("type") == "text" and part.get("text"):
                        chunks.append(str(part.get("text")))
                    elif part.get("text"):
                        chunks.append(str(part.get("text")))
            text = "".join(chunks).strip()
        elif AI_ANALYSIS_PROVIDER == "openai":
            key = os.getenv("OPENAI_API_KEY")
            if not key:
                return None
            body = {
                "model": model,
                "temperature": 0.05,
                "response_format": {"type": "json_object"},
                "messages": [
                    {"role": "system", "content": _ai_system_prompt()},
                    {"role": "user", "content": payload_context},
                ],
            }
            req = urllib.request.Request(
                "https://api.openai.com/v1/chat/completions",
                data=json.dumps(body).encode("utf-8"),
                headers={"content-type":"application/json", "authorization":f"Bearer {key}"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=AI_ANALYSIS_TIMEOUT) as resp:
                raw_response = resp.read().decode("utf-8", errors="replace")
            AI_ANALYSIS_LAST_PROVIDER_RESPONSE = raw_response[:4000]
            try:
                data = json.loads(raw_response)
            except Exception as exc:
                AI_ANALYSIS_LAST_ERROR = f"OpenAI returned non-JSON HTTP payload ({type(exc).__name__}). Preview: {raw_response[:500]}"
                return None
            text = str(data.get("choices", [{}])[0].get("message", {}).get("content", "") or "").strip()
        else:
            AI_ANALYSIS_LAST_ERROR = f"Proveedor IA no soportado: {AI_ANALYSIS_PROVIDER}"
            return None

        if text:
            AI_ANALYSIS_LAST_PROVIDER_RESPONSE = str(text)[:4000]
        parsed = _extract_json_object(text)
        clean = _normalize_ai_sections(parsed)
        if not clean:
            preview = (text or "").replace("\n", " ")[:500]
            AI_ANALYSIS_LAST_ERROR = f"AI provider returned non-JSON or no valid sections. Preview: {preview}"
        return clean or None
    except urllib.error.HTTPError as exc:
        try:
            body = exc.read().decode("utf-8", errors="replace")[:1600]
        except Exception:
            body = ""
        AI_ANALYSIS_LAST_ERROR = f"HTTP {exc.code} from {AI_ANALYSIS_PROVIDER} using model {model}: {body}"
        return None
    except Exception as exc:
        AI_ANALYSIS_LAST_ERROR = f"{type(exc).__name__}: {exc}"
        return None


def generate_expert_ai_analysis(run: CanonicalRun | None, summary: dict[str, Any] | None = None) -> dict[str, Any]:
    context = _run_analysis_context(run, summary)
    status = _analysis_feature_status()
    sections = _call_external_ai_analysis(context)
    mode = status.get("mode")
    if not sections:
        sections = _local_expert_analysis(context)
        if status.get("enabled") and status.get("mode") == "EXTERNAL_AI":
            mode = "LOCAL_EXPERT_FALLBACK_AI_ERROR"
        else:
            mode = status.get("mode")
    return {
        "mode": mode,
        "provider": status.get("provider"),
        "model": _configured_ai_model(),
        "enabled": bool(status.get("enabled")),
        "sections": sections,
        "context": context,
        "error": AI_ANALYSIS_LAST_ERROR,
    }




# ---------------------------------------------------------------------------
# Professional diagnostic layer (commercial output)
# ---------------------------------------------------------------------------
def _status_from_severity(sev: str) -> tuple[str, str]:
    s = (sev or "").upper()
    if s in {"HIGH", "CRITICAL", "CRITICO", "CRÍTICO"}:
        return "Crítico", "red"
    if s in {"MEDIUM", "REVIEW", "WARNING", "MEDIO", "REVISAR"}:
        return "Revisar", "yellow"
    return "OK", "green"


def _diagnostic_run_type(context: dict[str, Any]) -> str:
    if context.get("kind") == "base_budget":
        return "base_budget"
    count = int(context.get("providers_count", 0) or 0)
    return "single_provider_comparison" if count <= 1 else "multi_provider_comparison"


def _diagnostic_run_label(run_type: str) -> str:
    return {
        "base_budget": "Presupuesto base",
        "single_provider_comparison": "Comparativa individual contra mercado",
        "multi_provider_comparison": "Comparativa múltiple de contratistas",
    }.get(run_type, "Diagnóstico profesional")


def _status_label(code: str) -> str:
    return {"OK":"Aceptable", "REVIEW":"Revisar", "CRITICAL":"Alto riesgo"}.get(str(code or "").upper(), "Revisar")


def _traffic_from_status(code: str) -> str:
    return {"OK":"green", "REVIEW":"yellow", "CRITICAL":"red"}.get(str(code or "").upper(), "yellow")


def _section_display_name(section: str) -> str:
    s = (section or "").upper()
    if "MATERIAL" in s:
        return "Materiales"
    if "MANO" in s or s == "MO" or "CUADRILLA" in s:
        return "Mano de obra"
    if "MAQUIN" in s or "EQUIPO" in s or "HERRAM" in s:
        return "Maquinaria / equipo"
    if "BASIC" in s:
        return "Básicos"
    if "INDIRECT" in s or "FINAN" in s or "TOTAL" in s or "DIRECT" in s:
        return "Indirectos / financiero"
    return str(section or "Otros")[:60]


def _section_key(section: str) -> str:
    label = _section_display_name(section)
    if label == "Materiales": return "materials"
    if label == "Mano de obra": return "labor"
    if label == "Maquinaria / equipo": return "equipment"
    if label == "Básicos": return "basics"
    if label == "Indirectos / financiero": return "financial"
    return "other"


def _is_real_apu_item(item: Any) -> bool:
    desc = str(getattr(item, "description", "") or "").strip()
    code = str(getattr(item, "code", "") or "").strip()
    sec = str(getattr(item, "section", "") or "").upper()
    label = (desc or code).upper()
    if not (desc or code):
        return False
    blocked = ["SUBTOTAL", "COSTO DIRECTO", "TOTAL POR SERVICIO", "PRECIO UNITARIO", "SECCION FINANCIERA", "SECCIÓN FINANCIERA"]
    if any(b in label for b in blocked):
        return False
    if any(b in sec for b in ["SUBTOTAL", "COSTO DIRECTO", "TOTAL POR SERVICIO"]):
        return False
    return getattr(item, "amount", None) is not None or getattr(item, "market_amount", None) is not None


def _impact_multiplier(item: Any, qty_by_key: dict[str, float]) -> float:
    key = str(getattr(item, "concept_key", "") or "")
    qty = float(qty_by_key.get(key, 1) or 1)
    return qty if qty > 0 else 1.0


def _item_row(item: Any, qty_by_key: dict[str, float], *, source: str = "comparison") -> dict[str, Any]:
    mult = _impact_multiplier(item, qty_by_key)
    amount = float(getattr(item, "amount", 0) or 0) * mult
    market_amount_raw = getattr(item, "market_amount", None)
    market_amount = float(market_amount_raw or 0) * mult if market_amount_raw is not None else None
    unit_price = float(getattr(item, "unit_price", 0) or 0)
    market_unit_price = getattr(item, "market_unit_price", None)
    if market_unit_price is not None:
        market_unit_price = float(market_unit_price or 0)
    diff = None
    diff_pct = None
    if market_amount is not None and market_amount:
        diff = amount - market_amount
        diff_pct = diff / market_amount * 100
    ref_code = str(getattr(item, "matched_reference_code", "") or "").strip()
    ref_desc = str(getattr(item, "matched_reference_description", "") or "").strip()
    return {
        "code": str(getattr(item, "code", "") or ""),
        "description": _short_desc(str(getattr(item, "description", "") or ""), 70),
        "unit": str(getattr(item, "unit", "") or ""),
        "operator": str(getattr(item, "operator", "") or getattr(item, "market_operator", "") or ""),
        "quantity": float(getattr(item, "quantity", 0) or 0),
        "contractor_unit_price": round(unit_price, 2),
        "market_unit_price": round(market_unit_price, 2) if market_unit_price is not None else None,
        "contractor_amount": round(amount, 2),
        "market_amount": round(market_amount, 2) if market_amount is not None else None,
        "difference_amount": round(diff, 2) if diff is not None else None,
        "difference_pct": round(diff_pct, 2) if diff_pct is not None else None,
        "impact_amount": round(amount, 2),
        "section": _section_display_name(str(getattr(item, "section", "") or "")),
        "market_reference": (ref_code + " - " + ref_desc).strip(" -"),
        "state": str(getattr(item, "state", "") or ""),
        "source": source,
    }


def _items_by_section(items: list[Any], qty_by_key: dict[str, float], source: str = "comparison") -> dict[str, list[dict[str, Any]]]:
    buckets = {"materials": [], "labor": [], "equipment": [], "basics": [], "financial": [], "other": []}
    for item in items or []:
        if not _is_real_apu_item(item):
            continue
        key = _section_key(str(getattr(item, "section", "") or ""))
        row = _item_row(item, qty_by_key, source=source)
        if key == "financial":
            continue
        buckets.setdefault(key, []).append(row)
    for key in buckets:
        buckets[key] = sorted(buckets[key], key=lambda r: float(r.get("impact_amount") or 0), reverse=True)
    return buckets


def _difference_status(diff_pct: Any, no_ref: bool = False) -> str:
    if no_ref:
        return "REVIEW"
    if diff_pct is None:
        return "REVIEW"
    try:
        v = abs(float(diff_pct or 0))
    except Exception:
        return "REVIEW"
    if v >= 20:
        return "CRITICAL"
    if v >= 8:
        return "REVIEW"
    return "OK"


def _recommended_action(row: dict[str, Any]) -> str:
    ref = row.get("market_reference")
    diff_pct = row.get("difference_pct")
    desc = row.get("description") or row.get("code") or "insumo"
    if not ref:
        return "Validar referencia Construdata o solicitar soporte del precio."
    if diff_pct is None:
        return "Revisar cantidad, rendimiento y referencia asociada."
    if float(diff_pct or 0) > 20:
        return "Revisar sobrecosto y negociar contra referencia de mercado."
    if float(diff_pct or 0) < -20:
        return "Validar que el precio bajo no omita alcance, rendimiento o insumos."
    return "Confirmar que la referencia corresponda técnicamente al alcance."


def _concept_row_from_context(item: dict[str, Any], total: float) -> dict[str, Any]:
    amount = float(item.get("amount", 0) or item.get("contractor_amount", 0) or 0)
    market = item.get("market_amount")
    market = float(market or 0) if market is not None else None
    diff = amount - market if market else None
    diff_pct = diff / market * 100 if market else None
    return {
        "concept_code": item.get("code", ""),
        "description": _short_desc(str(item.get("description", "") or ""), 70),
        "contractor_amount": round(amount, 2),
        "market_amount": round(market, 2) if market is not None else None,
        "difference_amount": round(diff, 2) if diff is not None else None,
        "difference_pct": round(diff_pct, 2) if diff_pct is not None else None,
        "participation_pct": round(amount / total * 100, 2) if total else 0,
        "probable_cause": "Alta concentración del importe" if total and amount / total >= 0.15 else "Partida relevante por impacto económico",
        "priority_action": "Revisar matriz, cantidades, rendimientos y referencia de mercado.",
    }


def _build_base_diagnostic_context(run: CanonicalRun, context: dict[str, Any]) -> dict[str, Any]:
    exec_concepts = [c for c in run.base_concepts if getattr(c, "is_executable", False)]
    qty_by_key = {canonical_key(c.code, c.description): float(c.quantity or 0) for c in exec_concepts}
    total = float(context.get("total_amount", 0) or 0)
    items = _items_by_section(run.base_apu_items, qty_by_key, source="base")
    top_concepts = [_concept_row_from_context(c, total) for c in (context.get("top_concepts") or [])[:10]]
    section_breakdown = context.get("section_breakdown") or {}
    return {
        "run_type": "base_budget",
        "run_label": "Presupuesto base",
        "name": run.project_name or "Presupuesto base",
        "date": datetime.utcnow().strftime("%Y-%m-%d"),
        "total_amount": total,
        "market_amount": None,
        "direct_cost": context.get("direct_cost", 0),
        "indirect_cost": context.get("indirect_cost", 0),
        "indirect_pct": context.get("indirect_pct", 25),
        "concepts_processed": context.get("concepts_executable", 0),
        "with_reference": context.get("matched_matrices", 0),
        "without_full_reference": context.get("unmatched_matrices", 0),
        "coverage_pct": context.get("coverage_pct", 0),
        "section_values": section_breakdown,
        "top_materials": items.get("materials", [])[:10],
        "top_labor": items.get("labor", [])[:10],
        "top_equipment": items.get("equipment", [])[:10],
        "critical_concepts": top_concepts,
        "validations": context.get("validation_samples", []),
    }


def _build_comparison_diagnostic_context(run: CanonicalRun, context: dict[str, Any]) -> dict[str, Any]:
    providers = context.get("providers") or []
    provider = providers[0] if providers else {}
    provider_obj = (run.providers or [None])[0] if (run.providers or []) else None
    valid_concepts = [c for c in getattr(provider_obj, "concepts", []) if _is_valid_comparativa_concept(c)] if provider_obj else []
    qty_by_key = {canonical_key(c.code, c.description): float(c.quantity or 0) for c in valid_concepts}
    total = float(provider.get("total_amount", context.get("max_amount", 0)) or 0)
    market_total = float(provider.get("market_total_amount", 0) or 0) or None
    items = _items_by_section(getattr(provider_obj, "apu_items", []) if provider_obj else [], qty_by_key, source="comparison")
    # Critical concepts: for one provider use its top concepts; for multi, aggregate best available rows by provider.
    crit = []
    if len(providers) <= 1:
        crit = [_concept_row_from_context(c, total) for c in (provider.get("top_concepts") or [])[:10]]
    else:
        for pctx in providers:
            for c in (pctx.get("top_concepts") or [])[:4]:
                row = _concept_row_from_context(c, float(pctx.get("total_amount", 0) or 0))
                row["provider"] = pctx.get("name", "")
                crit.append(row)
        crit = sorted(crit, key=lambda r: float(r.get("contractor_amount") or 0), reverse=True)[:10]
    br = (provider.get("section_breakdown") or {}).get("contractor", {}) if provider else {}
    br_market = (provider.get("section_breakdown") or {}).get("market", {}) if provider else {}
    return {
        "run_type": "single_provider_comparison" if len(providers) <= 1 else "multi_provider_comparison",
        "run_label": "Comparativa individual contra mercado" if len(providers) <= 1 else "Comparativa múltiple de contratistas",
        "name": run.project_name or "Comparativa APU",
        "date": datetime.utcnow().strftime("%Y-%m-%d"),
        "providers": providers,
        "total_amount": total,
        "market_amount": market_total,
        "difference_amount": round(total - market_total, 2) if market_total is not None else None,
        "difference_pct": round(((total / market_total) - 1) * 100, 2) if market_total else None,
        "direct_cost": br.get("direct"),
        "indirect_cost": br.get("indirect"),
        "indirect_pct": 25,
        "concepts_processed": provider.get("concepts_executable", 0),
        "with_reference": provider.get("reference_items", 0),
        "without_full_reference": int(provider.get("fallback_items", 0) or 0) + int(provider.get("no_reference_items", 0) or 0),
        "coverage_pct": round((float(provider.get("reference_items",0) or 0) / max(1, len(getattr(provider_obj, "apu_items", []) if provider_obj else []))) * 100, 2),
        "section_values": br,
        "section_market_values": br_market,
        "top_materials": items.get("materials", [])[:10],
        "top_labor": items.get("labor", [])[:10],
        "top_equipment": items.get("equipment", [])[:10],
        "critical_concepts": crit,
        "validations": [],
    }


def _build_diagnostic_context(run: CanonicalRun | None, summary: dict[str, Any] | None = None) -> dict[str, Any]:
    base_context = _run_analysis_context(run, summary)
    if not run:
        return {"run_type":"unknown", "run_label":"Diagnóstico profesional", "name":"", "date":datetime.utcnow().strftime("%Y-%m-%d")}
    if base_context.get("kind") == "base_budget":
        return _build_base_diagnostic_context(run, base_context)
    return _build_comparison_diagnostic_context(run, base_context)


def _section_summary_from_diag(ctx: dict[str, Any]) -> list[dict[str, Any]]:
    total = float(ctx.get("total_amount", 0) or 0)
    values = ctx.get("section_values") or {}
    mvalues = ctx.get("section_market_values") or {}
    mapping = [("materials","Materiales"),("labor","Mano de obra"),("equipment","Maquinaria / equipo"),("basics","Básicos"),("indirect","Indirectos / financiero")]
    rows = []
    for key, label in mapping:
        amount = float(values.get(key, 0) or 0)
        mamount = mvalues.get(key)
        if mamount is not None:
            mamount = float(mamount or 0)
        if amount == 0 and not mamount:
            continue
        diff = amount - mamount if mamount else None
        diff_pct = diff / mamount * 100 if mamount else None
        status = _difference_status(diff_pct)
        if total and amount / total >= 0.35:
            status = "REVIEW" if status == "OK" else status
        rows.append({
            "section": label,
            "contractor_amount": round(amount,2),
            "market_amount": round(mamount,2) if mamount is not None else None,
            "difference_amount": round(diff,2) if diff is not None else None,
            "difference_pct": round(diff_pct,2) if diff_pct is not None else None,
            "participation_pct": round(amount/total*100,2) if total else 0,
            "status": status,
            "comment": "Alta participación; revisar drivers" if total and amount/total >= 0.30 else "Sin alerta mayor por participación",
        })
    return rows


def _kpis_from_diag(ctx: dict[str, Any], alerts_count: int = 0, priorities_count: int = 0) -> list[dict[str, Any]]:
    kpis = [
        {"label":"Monto total", "value":ctx.get("total_amount"), "format":"currency", "status":"neutral"},
        {"label":"Monto mercado", "value":ctx.get("market_amount"), "format":"currency", "status":"neutral"},
        {"label":"Diferencia contra mercado", "value":ctx.get("difference_amount"), "format":"currency", "status":"review"},
        {"label":"% diferencia contra mercado", "value":ctx.get("difference_pct"), "format":"percent", "status":"review"},
        {"label":"Costo directo", "value":ctx.get("direct_cost"), "format":"currency", "status":"neutral"},
        {"label":"Indirecto", "value":ctx.get("indirect_cost"), "format":"currency", "status":"neutral"},
        {"label":"% indirecto aplicado", "value":ctx.get("indirect_pct"), "format":"percent", "status":"neutral"},
        {"label":"Conceptos procesados", "value":ctx.get("concepts_processed"), "format":"number", "status":"neutral"},
        {"label":"Con referencia", "value":ctx.get("with_reference"), "format":"number", "status":"ok"},
        {"label":"Sin referencia plena", "value":ctx.get("without_full_reference"), "format":"number", "status":"review"},
        {"label":"Cobertura de mercado", "value":ctx.get("coverage_pct"), "format":"percent", "status":"ok" if float(ctx.get("coverage_pct",0) or 0) >= 90 else "review"},
        {"label":"Alertas críticas", "value":alerts_count, "format":"number", "status":"critical" if alerts_count else "ok"},
        {"label":"Partidas prioritarias", "value":priorities_count, "format":"number", "status":"review" if priorities_count else "ok"},
    ]
    return [k for k in kpis if k.get("value") is not None]


def _build_market_alerts(ctx: dict[str, Any], section_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    alerts = []
    total = float(ctx.get("total_amount", 0) or 0)
    for row in section_rows:
        if row.get("status") in {"REVIEW", "CRITICAL"}:
            alerts.append({"severity":"HIGH" if row.get("status") == "CRITICAL" else "MEDIUM", "alert_type":"Sección con desviación o alta participación", "item":row.get("section"), "section":row.get("section"), "contractor_value":row.get("contractor_amount"), "market_value":row.get("market_amount"), "deviation_pct":row.get("difference_pct"), "analyst_check":"Validar composición, rendimientos y referencias principales de la sección."})
    for concept in (ctx.get("critical_concepts") or [])[:10]:
        if float(concept.get("participation_pct") or 0) >= 15:
            alerts.append({"severity":"HIGH", "alert_type":"Partida con alta concentración", "item":concept.get("concept_code"), "section":"Catálogo", "contractor_value":concept.get("contractor_amount"), "market_value":concept.get("market_amount"), "deviation_pct":concept.get("difference_pct"), "analyst_check":"Revisar matriz completa, cantidad, rendimiento y referencia asociada."})
    for key, label in [("top_materials","Material crítico con diferencia relevante"),("top_labor","Mano de obra con rendimiento sensible"),("top_equipment","Maquinaria con peso elevado")]:
        for item in (ctx.get(key) or [])[:10]:
            diff_pct = item.get("difference_pct")
            no_ref = not item.get("market_reference")
            if no_ref or (diff_pct is not None and abs(float(diff_pct or 0)) >= 15):
                alerts.append({"severity":"HIGH" if diff_pct is not None and abs(float(diff_pct or 0)) >= 25 else "MEDIUM", "alert_type":label if not no_ref else "Sin referencia Construdata", "item":item.get("code") or item.get("description"), "section":item.get("section"), "contractor_value":item.get("contractor_unit_price"), "market_value":item.get("market_unit_price"), "deviation_pct":diff_pct, "analyst_check":_recommended_action(item)})
    return alerts[:20]


def _build_review_plan(ctx: dict[str, Any], alerts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    plan = []
    for idx, concept in enumerate((ctx.get("critical_concepts") or [])[:5], 1):
        code = concept.get("concept_code") or "partida prioritaria"
        plan.append({"priority":idx, "what_to_review":f"Revisar partida {code}", "why_it_matters":f"Participa con {_pct_text(concept.get('participation_pct',0))} del total y tiene impacto de {_money_text(concept.get('contractor_amount',0))}.", "where_to_check":"Comparativa y Detalle del proveedor / Detalle Base", "decision_needed":"Confirmar matriz, alcance, cantidad, rendimiento y referencia de mercado."})
    start = len(plan) + 1
    for alert in alerts[:5]:
        if len(plan) >= 10:
            break
        plan.append({"priority":start, "what_to_review":str(alert.get("item") or alert.get("alert_type") or "alerta"), "why_it_matters":str(alert.get("alert_type") or "alerta contra mercado"), "where_to_check":"Tabla de alertas y columna Match Construdata en Detalle", "decision_needed":str(alert.get("analyst_check") or "Validar técnicamente antes de cierre.")})
        start += 1
    return plan


def _local_professional_diagnostic(ctx: dict[str, Any]) -> dict[str, Any]:
    section_rows = _section_summary_from_diag(ctx)
    alerts = _build_market_alerts(ctx, section_rows)
    plan = _build_review_plan(ctx, alerts)
    critical_count = len([a for a in alerts if a.get("severity") == "HIGH"])
    coverage = float(ctx.get("coverage_pct", 0) or 0)
    diff_pct = ctx.get("difference_pct")
    high_concentration = any(float(c.get("participation_pct") or 0) >= 30 for c in (ctx.get("critical_concepts") or []))
    if critical_count >= 3 or (diff_pct is not None and abs(float(diff_pct or 0)) >= 20):
        status = "CRITICAL"
    elif critical_count or high_concentration or coverage < 90 or int(ctx.get("without_full_reference",0) or 0) > 0:
        status = "REVIEW"
    else:
        status = "OK"
    headline = {"title":"Diagnóstico profesional APU", "run_type":ctx.get("run_type"), "general_status":status, "traffic_light":_traffic_from_status(status), "executive_line":"Revisar primero las partidas de mayor impacto y las referencias sin trazabilidad plena." if status != "OK" else "Resultado consistente; mantener revisión de trazabilidad en partidas principales."}
    # Add actions to top items.
    for key in ["top_materials", "top_labor", "top_equipment"]:
        for row in ctx.get(key, []) or []:
            row["recommended_action"] = _recommended_action(row)
    diagnosis = {
        "risk_summary":"El riesgo principal se concentra en partidas de alto impacto, desviaciones contra mercado y conceptos sin referencia plena.",
        "main_cost_driver":"La prioridad se determina por participación económica y por desviación contra mercado, no por número de observaciones.",
        "market_traceability":"Los valores sin referencia plena deben validarse contra Construdata o soporte documental antes del cierre.",
        "recommendation":"Usar el diagnóstico como guía de revisión y negociación; no liberar versión final sin atender las prioridades marcadas.",
    }
    decision = {"verdict":"ACCEPTABLE" if status == "OK" else "REVIEW_REQUIRED" if status == "REVIEW" else "HIGH_RISK", "main_reason":headline["executive_line"], "next_action":plan[0]["what_to_review"] if plan else "Mantener control de trazabilidad.", "priority":"HIGH" if status == "CRITICAL" else "MEDIUM" if status == "REVIEW" else "LOW"}
    return {"headline":headline, "kpis":_kpis_from_diag(ctx, critical_count, len(plan)), "section_summary":section_rows, "top_materials":ctx.get("top_materials", [])[:10], "top_labor":ctx.get("top_labor", [])[:10], "top_equipment":ctx.get("top_equipment", [])[:10], "critical_concepts":ctx.get("critical_concepts", [])[:10], "market_alerts":alerts, "analyst_review_plan":plan, "professional_diagnosis":diagnosis, "final_decision":decision}


def _professional_diagnostic_prompt() -> str:
    return (
        "Actúa como experto senior en análisis de precios unitarios, Neodata y Construdata. "
        "Recibirás un JSON con datos ya calculados. No calcules importes nuevos ni inventes datos. "
        "Devuelve únicamente JSON válido, sin markdown y sin texto fuera del JSON. "
        "No menciones IA, proveedor, modelo, motor, prompt, fallback ni detalles técnicos. "
        "El objetivo es producir un diagnóstico comercial, visual y accionable para un analista APU. "
        "Si run_type es single_provider_comparison, NO hagas ranking; usa lectura individual contra mercado. "
        "Si run_type es multi_provider_comparison, sí puedes comparar contratistas. "
        "Usa tablas y acciones, no párrafos largos. Las descripciones deben ser cortas. "
        "Respeta exactamente la estructura: headline, kpis, section_summary, top_materials, top_labor, top_equipment, critical_concepts, market_alerts, analyst_review_plan, professional_diagnosis, final_decision. "
        "Cada acción del plan debe decir qué revisar, por qué importa, dónde buscarlo y qué decisión tomar. "
        "Si falta evidencia, usa null, [] o 'requiere validación'."
    )


def _normalize_professional_diagnostic(parsed: dict[str, Any] | None, fallback: dict[str, Any]) -> dict[str, Any] | None:
    if not isinstance(parsed, dict):
        return None
    required = ["headline", "kpis", "section_summary", "critical_concepts", "market_alerts", "analyst_review_plan", "professional_diagnosis", "final_decision"]
    if not all(k in parsed for k in required):
        return None
    # Ensure arrays exist.
    for key in ["kpis", "section_summary", "top_materials", "top_labor", "top_equipment", "critical_concepts", "market_alerts", "analyst_review_plan"]:
        if not isinstance(parsed.get(key), list):
            parsed[key] = fallback.get(key, [])
    for key in ["headline", "professional_diagnosis", "final_decision"]:
        if not isinstance(parsed.get(key), dict):
            parsed[key] = fallback.get(key, {})
    return parsed


def _call_external_professional_diagnostic(ctx: dict[str, Any], fallback: dict[str, Any]) -> dict[str, Any] | None:
    global AI_ANALYSIS_LAST_ERROR, AI_ANALYSIS_LAST_PROVIDER_RESPONSE
    status = _analysis_feature_status()
    if not status.get("enabled") or status.get("mode") != "EXTERNAL_AI":
        return None
    model = _configured_ai_model()
    payload_context = json.dumps({"context": ctx, "fallback_schema": fallback}, ensure_ascii=False, default=str)[:36000]
    try:
        if AI_ANALYSIS_PROVIDER == "anthropic":
            key = os.getenv("ANTHROPIC_API_KEY")
            if not key:
                return None
            body = {"model": model, "max_tokens": 3500, "temperature": 0.05, "system": _professional_diagnostic_prompt(), "messages": [{"role":"user", "content": payload_context}]}
            req = urllib.request.Request("https://api.anthropic.com/v1/messages", data=json.dumps(body).encode("utf-8"), headers={"content-type":"application/json", "x-api-key":key, "anthropic-version":"2023-06-01"}, method="POST")
            with urllib.request.urlopen(req, timeout=AI_ANALYSIS_TIMEOUT) as resp:
                raw_response = resp.read().decode("utf-8", errors="replace")
            AI_ANALYSIS_LAST_PROVIDER_RESPONSE = raw_response[:4000]
            data = json.loads(raw_response)
            text = "".join(str(part.get("text", "")) for part in data.get("content", []) if isinstance(part, dict)).strip()
        elif AI_ANALYSIS_PROVIDER == "openai":
            key = os.getenv("OPENAI_API_KEY")
            if not key:
                return None
            body = {"model": model, "temperature": 0.05, "response_format": {"type":"json_object"}, "messages":[{"role":"system", "content":_professional_diagnostic_prompt()}, {"role":"user", "content": payload_context}]}
            req = urllib.request.Request("https://api.openai.com/v1/chat/completions", data=json.dumps(body).encode("utf-8"), headers={"content-type":"application/json", "authorization":f"Bearer {key}"}, method="POST")
            with urllib.request.urlopen(req, timeout=AI_ANALYSIS_TIMEOUT) as resp:
                raw_response = resp.read().decode("utf-8", errors="replace")
            AI_ANALYSIS_LAST_PROVIDER_RESPONSE = raw_response[:4000]
            data = json.loads(raw_response)
            text = str(data.get("choices", [{}])[0].get("message", {}).get("content", "") or "").strip()
        else:
            return None
        parsed = _extract_json_object(text)
        if text:
            AI_ANALYSIS_LAST_PROVIDER_RESPONSE = text[:4000]
        diag = _normalize_professional_diagnostic(parsed, fallback)
        if not diag:
            AI_ANALYSIS_LAST_ERROR = "El proveedor no devolvió la estructura de diagnóstico esperada."
        return diag
    except urllib.error.HTTPError as exc:
        try:
            body = exc.read().decode("utf-8", errors="replace")[:1200]
        except Exception:
            body = ""
        AI_ANALYSIS_LAST_ERROR = f"HTTP {exc.code}: {body}"
        return None
    except Exception as exc:
        AI_ANALYSIS_LAST_ERROR = f"{type(exc).__name__}: {exc}"
        return None


def generate_professional_diagnostic(run: CanonicalRun | None, summary: dict[str, Any] | None = None) -> dict[str, Any]:
    ctx = _build_diagnostic_context(run, summary)
    fallback = _local_professional_diagnostic(ctx)
    external = _call_external_professional_diagnostic(ctx, fallback)
    diag = external or fallback
    # Never expose technical mode/provider/model in the commercial payload.
    diag["_context"] = ctx
    return diag


def _diag_value(value: Any, fmt: str | None = None) -> str:
    if value is None or value == "":
        return "N/A"
    if fmt == "currency":
        return _money_text(value)
    if fmt == "percent":
        return _pct_text(value)
    return str(value)


def _write_diag_section_title(ws, row: int, title: str, max_col: int = 10) -> int:
    _section_label(ws, row, title, max_col)
    return row + 1


def _write_table_block(ws, row: int, title: str, headers: list[str], rows: list[list[Any]], *, money_cols: list[int] | None = None, pct_cols: list[int] | None = None, int_cols: list[int] | None = None, table_name: str = "DiagTable") -> int:
    row = _write_diag_section_title(ws, row, title, max(10, len(headers)))
    for c, h in enumerate(headers, 1):
        ws.cell(row, c, h)
    _header_style(ws, row, 1, len(headers))
    start = row + 1
    if not rows:
        ws.cell(start, 1, "Sin datos calculados para esta sección")
        ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=len(headers))
        start += 1
    else:
        for vals in rows:
            for c, v in enumerate(vals, 1):
                ws.cell(start, c, v)
            start += 1
    _body_style(ws, row + 1, start - 1, 1, len(headers))
    _apply_formats(ws, money_cols=money_cols or [], pct_cols=pct_cols or [], int_cols=int_cols or [], start_row=row+1, end_row=start-1)
    try:
        _add_table(ws, f"A{row}:{get_column_letter(len(headers))}{start-1}", table_name, "TableStyleMedium2")
    except Exception:
        pass
    return start + 2


def _write_analisis_ia(ws, run: CanonicalRun | None = None):
    diag = generate_professional_diagnostic(run)
    _setup_sheet(ws, "Análisis IA", "Diagnóstico profesional para revisión de precios unitarios.", 10)
    _set_widths(ws, {"A": 20, "B": 42, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 20, "I": 22, "J": 32})
    headline = diag.get("headline", {})
    decision = diag.get("final_decision", {})
    row = 3
    ws.cell(row, 1, "Tipo de corrida"); ws.cell(row, 2, _diagnostic_run_label(headline.get("run_type", "")))
    ws.cell(row, 4, "Estado general"); ws.cell(row, 5, _status_label(headline.get("general_status", "REVIEW")))
    ws.cell(row + 1, 1, "Lectura ejecutiva"); ws.cell(row + 1, 2, headline.get("executive_line", ""))
    ws.merge_cells(start_row=row+1, start_column=2, end_row=row+1, end_column=10)
    for c in [1,4]: ws.cell(row, c).font = Font(bold=True, color=BRAND["navy"])
    ws.cell(row+1, 1).font = Font(bold=True, color=BRAND["navy"])
    ws.cell(row+1, 2).alignment = Alignment(wrap_text=True)
    row += 4

    # KPI table
    kpi_rows = [[k.get("label"), _diag_value(k.get("value"), k.get("format")), k.get("status", "")] for k in diag.get("kpis", [])]
    row = _write_table_block(ws, row, "KPIs principales", ["Indicador", "Valor", "Estado"], kpi_rows, table_name="DiagKpis")

    section_rows = []
    for s in diag.get("section_summary", []):
        section_rows.append([s.get("section"), s.get("contractor_amount"), s.get("market_amount"), s.get("difference_amount"), (float(s.get("difference_pct") or 0)/100 if s.get("difference_pct") is not None else None), (float(s.get("participation_pct") or 0)/100), _status_label(s.get("status")), s.get("comment")])
    row = _write_table_block(ws, row, "Resumen por secciones APU", ["Sección", "Importe", "Mercado", "Diferencia $", "Diferencia %", "% total", "Estado", "Comentario"], section_rows, money_cols=[2,3,4], pct_cols=[5,6], table_name="DiagSections")

    mat_rows = [[i.get("code"), i.get("description"), i.get("unit"), i.get("quantity"), i.get("contractor_unit_price"), i.get("market_unit_price"), i.get("difference_amount"), (float(i.get("difference_pct") or 0)/100 if i.get("difference_pct") is not None else None), i.get("impact_amount"), i.get("market_reference"), i.get("recommended_action")] for i in diag.get("top_materials", [])[:10]]
    row = _write_table_block(ws, row, "Top 10 materiales por impacto", ["Código", "Descripción", "Unidad", "Cantidad", "P.U.", "P.U. mercado", "Dif. $", "Dif. %", "Importe", "Referencia", "Acción"], mat_rows, money_cols=[5,6,7,9], pct_cols=[8], table_name="DiagMaterials")

    lab_rows = [[i.get("code"), i.get("description"), i.get("unit"), i.get("operator"), i.get("quantity"), i.get("contractor_unit_price"), i.get("market_unit_price"), i.get("difference_amount"), (float(i.get("difference_pct") or 0)/100 if i.get("difference_pct") is not None else None), i.get("impact_amount"), i.get("recommended_action")] for i in diag.get("top_labor", [])[:10]]
    row = _write_table_block(ws, row, "Top 10 mano de obra / cuadrillas", ["Código", "Descripción", "Unidad", "Op.", "Rend./Cant.", "P.U.", "P.U. mercado", "Dif. $", "Dif. %", "Importe", "Acción"], lab_rows, money_cols=[6,7,8,10], pct_cols=[9], table_name="DiagLabor")

    eq_rows = [[i.get("code"), i.get("description"), i.get("unit"), i.get("quantity"), i.get("contractor_unit_price"), i.get("market_unit_price"), i.get("difference_amount"), (float(i.get("difference_pct") or 0)/100 if i.get("difference_pct") is not None else None), i.get("impact_amount"), i.get("recommended_action")] for i in diag.get("top_equipment", [])[:10]]
    row = _write_table_block(ws, row, "Top 10 maquinaria / equipo", ["Código", "Descripción", "Unidad", "Cantidad", "P.U.", "P.U. mercado", "Dif. $", "Dif. %", "Importe", "Acción"], eq_rows, money_cols=[5,6,7,9], pct_cols=[8], table_name="DiagEquipment")

    concept_rows = [[c.get("concept_code"), c.get("description"), c.get("contractor_amount"), c.get("market_amount"), c.get("difference_amount"), (float(c.get("difference_pct") or 0)/100 if c.get("difference_pct") is not None else None), (float(c.get("participation_pct") or 0)/100), c.get("probable_cause"), c.get("priority_action")] for c in diag.get("critical_concepts", [])[:10]]
    row = _write_table_block(ws, row, "Partidas críticas del catálogo", ["Partida", "Descripción", "Importe", "Mercado", "Dif. $", "Dif. %", "% total", "Causa probable", "Acción prioritaria"], concept_rows, money_cols=[3,4,5], pct_cols=[6,7], table_name="DiagConcepts")

    alert_rows = [[a.get("severity"), a.get("alert_type"), a.get("item"), a.get("section"), a.get("contractor_value"), a.get("market_value"), (float(a.get("deviation_pct") or 0)/100 if a.get("deviation_pct") is not None else None), a.get("analyst_check")] for a in diag.get("market_alerts", [])[:20]]
    row = _write_table_block(ws, row, "Alertas contra mercado", ["Severidad", "Tipo", "Partida/Insumo", "Sección", "Valor", "Mercado", "Desviación", "Qué revisar"], alert_rows, money_cols=[5,6], pct_cols=[7], table_name="DiagAlerts")

    plan_rows = [[p.get("priority"), p.get("what_to_review"), p.get("why_it_matters"), p.get("where_to_check"), p.get("decision_needed")] for p in diag.get("analyst_review_plan", [])[:10]]
    row = _write_table_block(ws, row, "Plan de revisión para el analista", ["Prioridad", "Qué revisar", "Por qué importa", "Dónde buscar", "Decisión requerida"], plan_rows, int_cols=[1], table_name="DiagPlan")

    pd = diag.get("professional_diagnosis", {})
    diagnosis_rows = [["Riesgo principal", pd.get("risk_summary")], ["Driver de costo", pd.get("main_cost_driver")], ["Trazabilidad", pd.get("market_traceability")], ["Recomendación", pd.get("recommendation")]]
    row = _write_table_block(ws, row, "Diagnóstico profesional breve", ["Tema", "Lectura"], diagnosis_rows, table_name="DiagBrief")

    final_rows = [[_status_label(headline.get("general_status", "REVIEW")), decision.get("main_reason"), decision.get("next_action"), decision.get("priority")]]
    row = _write_table_block(ws, row, "Conclusión ejecutiva", ["Dictamen", "Motivo principal", "Próxima acción", "Prioridad"], final_rows, table_name="DiagDecision")
    ws.freeze_panes = "A7"

def _write_base_budget_report(wb):
    """Generate the independent base-budget workbook.

    Important V0.7 correction: the base-budget detail uses the same canonical
    APU detail writer as contractor comparisons. Only market columns are
    removed, because the generated base matrix is itself the market/reference
    matrix. This keeps both processes attached to the same canonical data model.
    """
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    _setup_sheet(ws, "Presupuesto Base", "Presupuesto independiente generado desde conceptos de ingeniería + matrices Construdata.", 10)
    _set_widths(ws, {"A": 22, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22, "G": 22, "H": 22, "I": 22, "J": 22})
    _write_kpi(ws, 4, 1, "Monto estimado", 56520000, "Base mock", "F8FAFC")
    _write_kpi(ws, 4, 3, "Conceptos", 128, "Archivo ingeniería", "F8FAFC")
    _write_kpi(ws, 4, 5, "Con match", 104, "Construdata matrices", "E2F0D9")
    _write_kpi(ws, 4, 7, "En revisión", 17, "Match medio", "FFF2CC")
    _write_kpi(ws, 4, 9, "Sin match", 7, "Cotizar", "FCE4D6")
    for cell in ["A5"]: ws[cell].number_format = MONEY_FMT
    _section_label(ws, 9, "Regla de modelo canónico", 10)
    ws["A10"] = "El presupuesto base es independiente del comparador, pero comparte el mismo modelo canónico de matriz/APU y el mismo escritor de detalle."
    ws.merge_cells(start_row=10, start_column=1, end_row=10, end_column=10)
    ws["A10"].alignment = Alignment(wrap_text=True)

    comp = wb.create_sheet("Comparativa")
    _setup_sheet(comp, "Comparativa presupuesto base", "Conceptos presupuestados contra matrices Construdata. No es flujo obligatorio de licitación.", 10)
    _set_widths(comp, {"A":14,"B":42,"C":12,"D":12,"E":42,"F":16,"G":18,"H":14,"I":16,"J":28})
    headers = ["Código", "Concepto ingeniería", "Unidad", "Cantidad", "Concepto Construdata", "P.U. referencia", "Importe base", "Confianza", "Estado", "Observación"]
    for c,h in enumerate(headers,1): comp.cell(5,c,h)
    _header_style(comp,5,1,10)
    rows = [
        ["001", "Excavación manual", "m3", 120, "Excavación manual material común", 88000, "=D6*F6", "Alta", "Con precio", "Match por descripción/unidad"],
        ["002", "Concreto f'c 250", "m3", 40, "Concreto 250 kg/cm2", 420000, "=D7*F7", "Alta", "Con precio", "Match confiable"],
        ["003", "Acero de refuerzo", "kg", 1800, "Acero fy 4200", 6200, "=D8*F8", "Media", "Revisar", "Validar especificación"],
        ["004", "Partida especial", "gl", 1, "", 0, "=D9*F9", "Baja", "Sin referencia", "Requiere cotización externa"],
    ]
    for r,row in enumerate(rows,6):
        for c,v in enumerate(row,1): comp.cell(r,c,v)
    _body_style(comp,6,9,1,10)
    _apply_formats(comp, money_cols=[6,7], start_row=6, end_row=9)
    _add_table(comp,"A5:J9","BaseBudgetComparativaTable","TableStyleMedium2")

    # Shared canonical detail writer. Same format as provider detail, excluding market columns.
    detail = wb.create_sheet("Detalle Base")
    _write_canonical_apu_detail_sheet(detail, "BASE", "1F4E79", variant="A", include_market=False, source_type="BASE_BUDGET")

def _write_comparison_report(wb, provider_names: Optional[List[str]] = None):
    names = _safe_provider_names(provider_names)
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    _write_resumen_ejecutivo(ws, "comparison", names)
    _write_comparativa_professional(wb.create_sheet("Comparativa"), names)
    palette = ["1F4E79", "0E6B3D", "7C3AED", "B54708"]
    for idx, name in enumerate(names):
        _write_contractor_detail_sheet(wb.create_sheet(f"Detalle - {name}"), name, palette[idx % len(palette)], "A" if idx % 2 == 0 else "B")
    _write_partidas_criticas(wb.create_sheet("Partidas Críticas"))
    _write_insumos_criticos(wb.create_sheet("Insumos Críticos"))
    _write_validaciones(wb.create_sheet("Validaciones"))
    _write_analisis_ia(wb.create_sheet("Análisis IA"))


def build_report(kind: str, provider_names: Optional[List[str]] = None) -> Path:
    wb = Workbook()
    if kind == "comparison":
        _write_comparison_report(wb, provider_names)
    else:
        _write_base_budget_report(wb)

    # Workbook-level finishing touches
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        # Make first tab active by default and keep tab colors subtle.
        if sheet.title in {"Resumen Ejecutivo", "Presupuesto Base"}:
            sheet.sheet_properties.tabColor = BRAND["navy"]
        elif "Detalle" in sheet.title:
            sheet.sheet_properties.tabColor = "70AD47"
        elif "Validaciones" in sheet.title:
            sheet.sheet_properties.tabColor = "FFC000"
        elif "IA" in sheet.title:
            sheet.sheet_properties.tabColor = "5B9BD5"

    filename = f"apu_v0_{kind}_professional_{int(time.time())}.xlsx"
    out = REPORTS_DIR / filename
    wb.save(out)
    return out


@app.get("/api/reports/{kind}")
def report(kind: str, providers: Optional[str] = Query(default=None)):
    if kind not in {"base", "comparison"}:
        raise HTTPException(status_code=400, detail="kind debe ser base o comparison")
    if kind == "base":
        latest_base = REAL_REPORTS.get("LATEST_BASE")
        if latest_base and latest_base.exists():
            return FileResponse(latest_base, filename=latest_base.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        raise HTTPException(status_code=409, detail="No hay presupuesto base real generado en esta sesión. Usa /api/base-budgets/real-run con un .xlsx.")
    provider_names = [p.strip() for p in providers.split(",") if p.strip()] if providers else None
    path = build_report(kind, provider_names)
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



# -----------------------------------------------------------------------------
# V1: real-data ingestion through the canonical model
# -----------------------------------------------------------------------------
# These endpoints keep the V0 professional UI/Excel shell, but start replacing
# static demo paths with parsed XLSX content. This version is intentionally
# tolerant: it detects headers by synonyms and maps rows into the canonical
# model. Unsupported structures are surfaced as validations instead of being
# silently forced into a fake format.

UPLOADS_DIR = RUNTIME_DIR / "uploads"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
REAL_RUNS: dict[str, CanonicalRun] = {}
REAL_REPORTS: dict[str, Path] = {}


BASE_RUN_SUMMARIES: dict[str, dict[str, Any]] = {}


def _short_code_prefix(code: str) -> str:
    code = (code or '').strip().upper()
    if not code:
        return 'SIN CODIGO'
    # Prefer the first alphabetic/number prefix before separators. Examples: PR01 -> PR, 1.1.1 -> 1, DD-01 -> DD
    import re as _re
    m = _re.match(r'([A-Z]+)', code)
    if m:
        return m.group(1)[:8]
    return code.split('.')[0].split('-')[0][:8]


def _base_run_summary(run: CanonicalRun, report_path: Path | None = None, source_file: str = '') -> dict[str, Any]:
    exec_concepts = [c for c in run.base_concepts if getattr(c, 'is_executable', False)]
    section_totals = _base_section_breakdown_sequential(run.base_apu_items)
    total_amount = sum(float(c.amount or 0) for c in exec_concepts)
    direct_cost = float(section_totals.get('direct', 0) or 0)
    indirect_cost = float(section_totals.get('indirect', 0) or 0)

    cost_breakdown = {
        'materials': {'label': 'Materiales', 'amount': round(float(section_totals.get('materials', 0) or 0), 2)},
        'labor': {'label': 'Mano de obra', 'amount': round(float(section_totals.get('labor', 0) or 0), 2)},
        'equipment': {'label': 'Maquinaria / equipo', 'amount': round(float(section_totals.get('equipment', 0) or 0), 2)},
        'basics': {'label': 'Básicos', 'amount': round(float(section_totals.get('basics', 0) or 0), 2)},
        'indirect': {'label': 'Indirecto 25%', 'amount': round(indirect_cost, 2)},
    }

    matched = len([c for c in exec_concepts if str(getattr(c, 'market_state', '')).startswith('Match')])
    unmatched = max(0, len(exec_concepts) - matched)
    coverage_pct = (matched / len(exec_concepts) * 100) if exec_concepts else 0

    segs: dict[str, float] = {}
    for c in exec_concepts:
        prefix = _short_code_prefix(c.code)
        segs[prefix] = segs.get(prefix, 0.0) + float(c.amount or 0)
    segments = [
        {'code': k, 'amount': round(v, 2), 'weightPct': round((v / total_amount * 100) if total_amount else 0, 2)}
        for k, v in sorted(segs.items(), key=lambda kv: kv[1], reverse=True)[:10]
    ]

    top_concepts = []
    for c in sorted(exec_concepts, key=lambda x: float(x.amount or 0), reverse=True)[:10]:
        top_concepts.append({
            'code': c.code or '',
            'unit': c.unit or '',
            'quantity': float(c.quantity or 0),
            'unitPrice': round(float(c.unit_price or 0), 2),
            'amount': round(float(c.amount or 0), 2),
            'weightPct': round((float(c.amount or 0) / total_amount * 100) if total_amount else 0, 2),
            'state': c.market_state or '',
        })

    findings = []
    findings.append(f'Se procesaron {len(exec_concepts)} conceptos ejecutables con {len(run.base_apu_items)} filas de Detalle Base.')
    if total_amount:
        findings.append(f'El monto total estimado es {total_amount:,.0f}; el indirecto aplicado corresponde al 25% del costo directo.')
    findings.append(f'La cobertura de matrices Construdata es {coverage_pct:.1f}% ({matched} con match, {unmatched} en revisión).')
    if top_concepts:
        findings.append('Las partidas de mayor impacto deben revisarse primero por su peso económico relativo.')
    if unmatched:
        findings.append('Los conceptos sin matriz directa requieren validación técnica antes de emitir una versión final.')

    analysis = generate_expert_ai_analysis(run)
    findings = [s.get('content', '') for s in analysis.get('sections', []) if s.get('content')]

    return {
        'runId': run.run_id,
        'type': 'base_budget',
        'projectName': run.project_name,
        'status': 'COMPLETED_WITH_WARNINGS' if unmatched else 'COMPLETED',
        'sourceFile': source_file,
        'downloadUrl': f'/api/real-runs/{run.run_id}/report',
        'aiReportUrl': f'/api/real-runs/{run.run_id}/ai-report',
        'reportFile': report_path.name if report_path else '',
        'conceptsRead': len(run.base_concepts),
        'conceptsExecutable': len(exec_concepts),
        'detailRows': len(run.base_apu_items),
        'validations': len(run.validations or []),
        'totalAmount': round(total_amount, 2),
        'directCost': round(direct_cost, 2),
        'indirectCost': round(indirect_cost, 2),
        'coverage': {
            'matched': matched,
            'unmatched': unmatched,
            'estimated': 0,
            'coveragePct': round(coverage_pct, 2),
        },
        'costBreakdown': cost_breakdown,
        'segments': segments,
        'topConcepts': top_concepts,
        'executiveFindings': findings,
    }


def _safe_filename(name: str) -> str:
    keep = []
    for ch in (name or "archivo.xlsx"):
        if ch.isalnum() or ch in {".", "-", "_"}:
            keep.append(ch)
        else:
            keep.append("_")
    return "".join(keep)[:140] or "archivo.xlsx"


async def _save_upload(upload: UploadFile, target_dir: Path) -> Path:
    if not upload.filename or not upload.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail=f"Archivo no permitido: {upload.filename}. Solo se acepta .xlsx")
    target_dir.mkdir(parents=True, exist_ok=True)
    path = target_dir / _safe_filename(upload.filename)
    content = await upload.read()
    path.write_bytes(content)
    return path


def _canonical_amount_from_concepts(concepts: list[Any]) -> float:
    total = 0.0
    for c in concepts:
        if getattr(c, "amount", None) is not None:
            total += float(c.amount or 0)
        elif getattr(c, "quantity", None) is not None and getattr(c, "unit_price", None) is not None:
            total += float(c.quantity or 0) * float(c.unit_price or 0)
    return total


def _canonical_concept_union(providers: list[CanonicalProvider]) -> list[Any]:
    if not providers:
        return []
    # Canonical comparison spine: union of catalog concepts across providers.
    # Never use matrix/APU rows here. Preserve provider order and original catalog
    # order; append concepts that do not exist in previous providers. This keeps
    # Comparativa concept-only while allowing providers with non-identical catalogs.
    seen: set[str] = set()
    spine: list[Any] = []
    for provider in providers:
        for concept in provider.concepts:
            # Comparativa contract: only concepts declared in the contractor's
            # concept/catalog file with unit and quantity > 0. Never use APU
            # insumos or hierarchy rows in this tab.
            if not _is_valid_comparativa_concept(concept):
                continue
            key = _concept_key(concept)
            if not key or key in seen:
                continue
            seen.add(key)
            spine.append(concept)
    return spine


def _concept_key(c: Any) -> str:
    code = str(getattr(c, "code", "") or "").strip().lower()
    if code:
        return f"code:{code}"
    desc = str(getattr(c, "description", "") or "").strip().lower()
    return "desc:" + " ".join(desc.split())[:120]


def _is_valid_comparativa_concept(c: Any) -> bool:
    unit = str(getattr(c, "unit", "") or "").strip()
    qty = getattr(c, "quantity", None)
    try:
        qty_ok = qty is not None and float(qty) > 0
    except Exception:
        qty_ok = False
    return bool(getattr(c, "is_executable", False) and unit and qty_ok)


def _concept_amount(c: Any) -> float:
    amount = getattr(c, "amount", None)
    if amount is not None:
        try:
            return float(amount or 0)
        except Exception:
            return 0.0
    pu = getattr(c, "unit_price", None)
    qty = getattr(c, "quantity", None)
    if pu is not None and qty is not None:
        try:
            return float(pu or 0) * float(qty or 0)
        except Exception:
            return 0.0
    return 0.0


def _provider_pareto_keys(provider: CanonicalProvider) -> set[str]:
    """Return concept keys that explain the first 80% of a provider amount.

    Canonical rule: Pareto 80/20 belongs to each provider because it is
    calculated from that provider's own catalog amount (P.U. total * quantity).
    It must never be promoted to the shared catalog/base columns in
    Comparativa; the writer may only shade the provider's own block.
    """
    rows = [c for c in provider.concepts if _is_valid_comparativa_concept(c) and _concept_amount(c) > 0]
    total = sum(_concept_amount(c) for c in rows)
    if total <= 0:
        return set()
    selected: set[str] = set()
    cumulative = 0.0
    for c in sorted(rows, key=_concept_amount, reverse=True):
        previous = cumulative
        cumulative += _concept_amount(c) / total
        if cumulative <= 0.80 or previous < 0.80:
            selected.add(_concept_key(c))
    # Persist the provider-scoped Pareto keys in the canonical provider object
    # so downstream writers, KPIs and IA all consume the same semantic result.
    try:
        provider.pareto_concept_keys = sorted(selected)
    except Exception:
        pass
    return selected


def _write_real_comparativa(ws, providers: list[CanonicalProvider]):
    """Render the canonical comparison using the minimum accepted workbook contract.

    Output contract restored from Comparativo_cotizacion_2_proveedores.xlsx:
    - A:D = frozen base catalog/concept columns.
    - One 6-column block per provider to the right.
    - Market values live inside each provider block, not in a separate global block.
    - Row order follows the declared catalog spine.
    - Pareto 80/20 is provider-scoped. It is calculated from each provider's
      own concept amount and highlighted only inside that provider block.
      The shared Servicios/Cotización columns A:D are never shaded by Pareto.
    """
    spine = _canonical_concept_union(providers)
    names = [p.name for p in providers] or ["Proveedor"]
    ws.sheet_view.showGridLines = False

    base_cols = 4
    provider_cols = 6
    total_cols = base_cols + provider_cols * len(names)
    ws.freeze_panes = "E3"  # freeze row headers and A:D base catalog columns

    widths = {"A": 16, "B": 72, "C": 12, "D": 12}
    for idx, _ in enumerate(names):
        start = base_cols + idx * provider_cols + 1
        for offset, width in enumerate([15, 16, 12, 12, 16, 18]):
            widths[get_column_letter(start + offset)] = width
    _set_widths(ws, widths)

    _provider_group_header(ws, 1, 1, base_cols, "Servicios / Cotización", "475467")
    palette = ["1F4E79", "0E6B3D", "7C3AED", "B54708", "344054", "6941C6"]
    for idx, name in enumerate(names):
        start_col = base_cols + idx * provider_cols + 1
        _provider_group_header(ws, 1, start_col, start_col + provider_cols - 1, name, palette[idx % len(palette)])

    headers = ["Partida", "Descripción", "Unidad", "Cantidad"]
    for _ in names:
        headers += ["P.U.", "Importe", "% Part.", "% ajuste", "Mercado P.U.", "Mercado Importe"]
    for c, h in enumerate(headers, 1):
        ws.cell(2, c, h)
    _header_style(ws, 2, 1, base_cols, fill="475467")
    for idx, _ in enumerate(names):
        start_col = base_cols + idx * provider_cols + 1
        _header_style(ws, 2, start_col, start_col + provider_cols - 1, fill=palette[idx % len(palette)])

    concept_maps: list[dict[str, Any]] = []
    pareto_maps: list[set[str]] = []
    totals: list[float] = []
    for p in providers:
        m = {_concept_key(c): c for c in p.concepts}
        concept_maps.append(m)
        pareto_maps.append(_provider_pareto_keys(p))
        totals.append(sum(_concept_amount(c) for c in p.concepts if _is_valid_comparativa_concept(c)))

    if not spine:
        ws.cell(3, 1, "SIN-DATA")
        ws.cell(3, 2, "No se detectaron conceptos en los archivos cargados")
        spine = []
        max_rows = 1
    else:
        max_rows = min(len(spine), 500)

    for i in range(max_rows):
        row_idx = 3 + i
        base = spine[i] if spine and i < len(spine) else None
        code = getattr(base, "code", "") if base else ""
        desc = getattr(base, "description", "") if base else ""
        unit = getattr(base, "unit", "") if base else ""
        qty = getattr(base, "quantity", "") if base else ""
        executable = _is_valid_comparativa_concept(base) if base else False
        level = int(getattr(base, "hierarchy_level", 0) or 0) if base else 0
        indent = "   " * max(level - 1, 0)
        key = _concept_key(base) if base else ""

        ws.cell(row_idx, 1, code)
        ws.cell(row_idx, 2, f"{indent}{desc}")
        ws.cell(row_idx, 3, unit)
        ws.cell(row_idx, 4, qty)

        for pidx, _ in enumerate(names):
            start = base_cols + pidx * provider_cols + 1
            c = concept_maps[pidx].get(key) if pidx < len(concept_maps) and key else None
            pu = getattr(c, "unit_price", None) if c else None
            amount = _concept_amount(c) if c else None
            market_pu = getattr(c, "market_unit_price", None) if c else None
            market_amount = getattr(c, "market_amount", None) if c else None
            ws.cell(row_idx, start, pu if pu is not None else "")
            ws.cell(row_idx, start + 1, amount if amount not in (None, 0.0) else "")
            ws.cell(row_idx, start + 2, f"={get_column_letter(start+1)}{row_idx}/${get_column_letter(start+1)}${3+max_rows}" if executable else "")
            ws.cell(row_idx, start + 3, f"=IF({get_column_letter(start+4)}{row_idx}=0,0,{get_column_letter(start)}{row_idx}/{get_column_letter(start+4)}{row_idx}-1)" if executable else "")
            ws.cell(row_idx, start + 4, market_pu if market_pu is not None else "")
            ws.cell(row_idx, start + 5, market_amount if market_amount is not None else "")
            if key in (pareto_maps[pidx] if pidx < len(pareto_maps) else set()):
                # Pareto 80/20 is scoped to this provider. Only this provider's
                # six-column block is shaded. The base catalog columns A:D
                # remain neutral because they are shared Servicios/Cotización.
                for col in range(start, start + provider_cols):
                    ws.cell(row_idx, col).fill = PatternFill("solid", fgColor="DCEBFF")

        if not executable:
            base_fill = "F2F4F7"
            font = Font(bold=True, color=BRAND["text"])
        else:
            base_fill = "FFFFFF"
            font = Font(color=BRAND["text"])
        for col in range(1, base_cols + 1):
            ws.cell(row_idx, col).fill = PatternFill("solid", fgColor=base_fill)
            ws.cell(row_idx, col).font = font
        for col in range(1, total_cols + 1):
            cell = ws.cell(row_idx, col)
            if cell.fill.fgColor.rgb in ("00000000", "000000") or cell.fill.fill_type is None:
                cell.fill = PatternFill("solid", fgColor=("F2F4F7" if not executable else "FFFFFF"))
            cell.border = _thin_border("EAECF0")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if not executable:
                cell.font = Font(bold=True, color=BRAND["text"])

    total_row = 3 + max_rows
    ws.cell(total_row, 1, "TOTAL")
    ws.cell(total_row, 2, "")
    for pidx, _ in enumerate(names):
        start = base_cols + pidx * provider_cols + 1
        first_data = 3
        last_data = total_row - 1
        ws.cell(total_row, start, f"=AVERAGE({get_column_letter(start)}{first_data}:{get_column_letter(start)}{last_data})")
        ws.cell(total_row, start + 1, f"=SUM({get_column_letter(start+1)}{first_data}:{get_column_letter(start+1)}{last_data})")
        ws.cell(total_row, start + 4, f"=AVERAGE({get_column_letter(start+4)}{first_data}:{get_column_letter(start+4)}{last_data})")
        ws.cell(total_row, start + 5, f"=SUM({get_column_letter(start+5)}{first_data}:{get_column_letter(start+5)}{last_data})")
    for c in range(1, total_cols + 1):
        cell = ws.cell(total_row, c)
        cell.font = Font(bold=True, color=BRAND["navy"])
        cell.fill = PatternFill("solid", fgColor="F8FAFC")
        cell.border = Border(top=Side(style="medium", color="1F4E79"), bottom=Side(style="thin", color="D9E2EC"))
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    money_cols, pct_cols = [], []
    for idx, _ in enumerate(names):
        start = base_cols + idx * provider_cols + 1
        money_cols += [start, start + 1, start + 4, start + 5]
        pct_cols += [start + 2, start + 3]
        if total_row > 3:
            ws.conditional_formatting.add(f"{get_column_letter(start+3)}3:{get_column_letter(start+3)}{total_row-1}", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    _apply_formats(ws, money_cols=money_cols, pct_cols=pct_cols, start_row=3, end_row=total_row)
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}{total_row}"

    # Executive notes below, matching the accepted workbook spirit while making
    # the canonical contract explicit.
    note_row = total_row + 3
    _section_label(ws, note_row, "Resumen individual por proveedor", total_cols)
    row = note_row + 1
    for idx, name in enumerate(names):
        ws.cell(row, 1, name).font = Font(bold=True, color=BRAND["navy"], size=12)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
        row += 1
        bullets = [
            "Comparativa renderizada desde el catálogo canónico: se conserva el orden declarado y se congelan A:D.",
            "El bloque del proveedor contiene P.U., Importe, % Part., % ajuste, Mercado P.U. y Mercado Importe.",
            "El sombreado azul es independiente por proveedor y solo aparece dentro del bloque de ese proveedor; A:D permanece neutral.",
        ]
        for bullet in bullets:
            ws.cell(row, 1, "• " + bullet)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)
            ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[row].height = 28
            row += 1
        row += 1

def _write_real_canonical_detail(ws, title: str, rows: list[dict[str, Any]], *, include_market: bool, theme_color: str = "1F4E79"):
    """Render a PU-style detail sheet using the accepted Excel contract.

    Contract:
    - Contractor detail: A:H provider matrix, I separator, J:M market comparison.
    - Base detail: same A:H provider/base matrix, no market block.
    - No horizontal multi-provider detail; each provider has its own sheet.
    - Structural PU rows are preserved as section titles/subtotals/totals.
    """
    base_headers = ["Código", "Concepto", "Unidad", "P. Unitario", "Op.", "Cantidad", "Importe", "%"]
    market_headers = ["Mercado P. Unitario", "Mercado Op.", "Mercado Cantidad", "Mercado Importe", "Match Construdata"]
    headers = base_headers + ([""] + market_headers if include_market else [])
    max_col = len(headers)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    widths = [16, 64, 12, 16, 9, 12, 16, 11]
    if include_market:
        widths += [4, 18, 12, 16, 18, 46]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    _provider_group_header(ws, 1, 1, 8, title, theme_color)
    if include_market:
        ws.cell(1, 9, "")
        ws.cell(1, 9).fill = PatternFill("solid", fgColor="FFFFFF")
        _provider_group_header(ws, 1, 10, 14, "Mercado / Referencia", "475467")
    for c, h in enumerate(headers, 1):
        ws.cell(2, c, h)
    _header_style(ws, 2, 1, 8, fill=theme_color)
    if include_market:
        ws.cell(2, 9).fill = PatternFill("solid", fgColor="FFFFFF")
        ws.cell(2, 9).border = _thin_border("FFFFFF")
        _header_style(ws, 2, 10, 14, fill="475467")

    if not rows:
        rows = [{"code":"", "concept":"No se detectaron filas de matriz/APU", "unit":"", "section":"VALIDACIÓN", "pu":"", "op":"", "qty":"", "amount":"", "pct":"", "mpu":"", "mop":"", "mqty":"", "mamount":"", "dev":"", "state":"Sin datos", "obs":"Revisar estructura del archivo cargado"}]

    # Reuse style objects. Creating a fresh fill/border/font per cell makes
    # large PMD matrices extremely slow to save and inflates the XLSX style table.
    detail_fills = {
        "EEF4FF": PatternFill("solid", fgColor="EEF4FF"),
        "F2F4F7": PatternFill("solid", fgColor="F2F4F7"),
        "F8FAFC": PatternFill("solid", fgColor="F8FAFC"),
        "EAF2FF": PatternFill("solid", fgColor="EAF2FF"),
        "FFFFFF": PatternFill("solid", fgColor="FFFFFF"),
    }
    detail_border = _thin_border("EAECF0")
    sep_border = _thin_border("FFFFFF")
    normal_font = Font(bold=False, color=BRAND["text"])
    bold_font = Font(bold=True, color=BRAND["text"])
    top_alignment = Alignment(vertical="top", wrap_text=True)

    for ridx, row in enumerate(rows[:6000], 3):
        section = str(row.get("section", ""))
        concept_text = str(row.get("concept", "") or "")
        upper = concept_text.upper()
        is_partida = section == "PARTIDA"
        is_title = section == "TÍTULO"
        is_subtotal = section.startswith("SUBTOTAL") or upper.startswith("SUBTOTAL")
        is_financial = any(k in upper or k in section for k in ["TOTAL", "COSTO", "UTILIDAD", "FINANCIAMIENTO", "PRECIO"])

        values = [
            row.get("code", ""),
            row.get("concept", ""),
            row.get("unit", ""),
            row.get("pu", ""),
            row.get("op", ""),
            row.get("qty", ""),
            row.get("amount", ""),
            row.get("pct", ""),
        ]
        if include_market:
            values += [
                "",
                row.get("mpu", ""),
                row.get("mop", ""),
                row.get("mqty", ""),
                row.get("mamount", ""),
                row.get("match_ref", ""),
            ]
        for cidx, value in enumerate(values, 1):
            ws.cell(ridx, cidx, value)

        # Sobrio: el detalle debe ser técnico y limpio, no decorativo.
        if is_partida:
            fill = "EEF4FF"; bold = True
        elif is_title:
            fill = "F2F4F7"; bold = True
        elif is_subtotal:
            fill = "F8FAFC"; bold = True
        elif is_financial:
            fill = "EAF2FF"; bold = True
        elif str(row.get("op", "")) == "%" or str(row.get("unit", "")) == "%" or str(row.get("code", "")).startswith("%"):
            fill = "FFFFFF"; bold = False
        else:
            fill = "FFFFFF"; bold = False

        fill_obj = detail_fills.get(fill, detail_fills["FFFFFF"])
        font_obj = bold_font if bold else normal_font
        market_ref_cols = {
            10: bool(row.get("mpu_ref")),
            11: bool(row.get("mop_ref")),
            12: bool(row.get("mqty_ref")),
            13: bool(row.get("mamount_ref")),
        }
        for c in range(1, max_col + 1):
            cell = ws.cell(ridx, c)
            if include_market and c == 9:
                cell.fill = detail_fills["FFFFFF"]
                cell.border = sep_border
            else:
                cell.fill = fill_obj
                cell.border = detail_border
            cell.alignment = top_alignment
            # Market columns J:M are emphasized only when the canonical row says
            # the market value is reference-based or different from contractor.
            # Fallback values copied from contractor stay visually normal.
            if include_market and c in market_ref_cols and market_ref_cols[c]:
                cell.font = bold_font
            else:
                cell.font = font_obj

    last = min(2 + len(rows), 6002)
    money_cols = [4, 7]
    pct_cols = [8]
    if include_market:
        money_cols += [10, 13]
    _apply_formats(ws, money_cols=money_cols, pct_cols=pct_cols, start_row=3, end_row=last)
    ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{last}"

    # Fine visual cue: keep market block separated without adding extra audit
    # columns to the accepted detail layout. Validations remain in Validaciones.
    note_row = last + 3
    ws.cell(note_row, 1, "Modelo canónico")
    ws.cell(note_row, 1).font = Font(bold=True, color=BRAND["navy"])
    msg = "Este detalle se renderiza desde el mismo modelo canónico APU. En proveedor se agrega bloque J:M de mercado; en base se omite porque la matriz ya es mercado/base."
    ws.cell(note_row, 2, msg)
    end_note_col = max_col
    ws.merge_cells(start_row=note_row, start_column=2, end_row=note_row, end_column=end_note_col)
    ws.cell(note_row, 2).alignment = Alignment(wrap_text=True)

def _write_real_validaciones(ws, run: CanonicalRun):
    _setup_sheet(ws, "Validaciones", "Advertencias reales del proceso de carga y parseo inicial.", 10)
    headers = ["Severidad", "Tipo", "Proveedor", "Archivo", "Campo", "Valor", "Problema", "Acción sugerida", "Estado", "Origen"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, 10)
    rows = []
    for p in run.providers:
        if not p.concepts:
            rows.append(["Alta", "Parser conceptos", p.name, p.concepts_file, "conceptos", "0", "No se detectaron conceptos", "Validar encabezados del archivo", "Pendiente", "V1"])
        if not p.apu_items:
            rows.append(["Alta", "Parser matriz", p.name, p.matrix_file, "matriz", "0", "No se detectaron insumos APU", "Validar estructura de matriz/APU", "Pendiente", "V1"])
        no_match = len([i for i in p.apu_items if i.state == "Sin referencia"])
        if no_match:
            rows.append(["Media", "Mercado", p.name, p.matrix_file, "referencia", no_match, "Insumos sin match granular en data", "Revisar descripción/unidad o cargar referencia complementaria", "Pendiente", "V1"])
    if not rows:
        rows.append(["Baja", "Carga", "—", "—", "general", "OK", "No se generaron validaciones críticas", "Continuar revisión técnica", "Revisado", "V1"])
    for r, row in enumerate(rows, 6):
        for c, v in enumerate(row, 1): ws.cell(r, c, v)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=_status_fill(row[0]))
    _body_style(ws, 6, 5 + len(rows), 1, 10)
    _add_table(ws, f"A5:J{5+len(rows)}", "RealValidacionesTable", "TableStyleMedium2")
    _set_widths(ws, {"A":14,"B":22,"C":16,"D":28,"E":18,"F":12,"G":48,"H":42,"I":18,"J":18})


def build_real_comparison_report(run: CanonicalRun) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    _setup_sheet(ws, "Resumen Ejecutivo", "Reporte generado desde archivos XLSX reales y modelo canónico.", 10)
    names = [p.name for p in run.providers]
    totals = [_canonical_amount_from_concepts(p.concepts) for p in run.providers]
    best_idx = min(range(len(totals)), key=lambda i: totals[i]) if totals else 0
    _write_kpi(ws, 4, 1, "Proveedores", len(run.providers), "Carga real", "F8FAFC")
    _write_kpi(ws, 4, 3, "Mejor oferta", totals[best_idx] if totals else 0, names[best_idx] if names else "—", "E2F0D9")
    _write_kpi(ws, 4, 5, "Conceptos leídos", sum(len(p.concepts) for p in run.providers), "Catálogos", "F8FAFC")
    _write_kpi(ws, 4, 7, "Filas APU", sum(len(p.apu_items) for p in run.providers), "Matrices", "F8FAFC")
    _write_kpi(ws, 4, 9, "Motor", "V1", "Data real", "FFF2CC")
    for cell in ["C5"]: ws[cell].number_format = MONEY_FMT
    _section_label(ws, 9, "Ranking por importes detectados en catálogo de conceptos", 10)
    headers = ["Posición", "Proveedor", "Conceptos", "Filas matriz", "Monto detectado", "Observación"]
    for c, h in enumerate(headers, 1): ws.cell(10, c, h)
    _header_style(ws, 10, 1, len(headers))
    ranking = sorted([(totals[i], p) for i, p in enumerate(run.providers)], key=lambda x: x[0])
    for idx, (total, p) in enumerate(ranking, 11):
        ws.cell(idx, 1, idx-10); ws.cell(idx, 2, p.name); ws.cell(idx, 3, len(p.concepts)); ws.cell(idx, 4, len(p.apu_items)); ws.cell(idx, 5, total); ws.cell(idx, 6, "Monto calculado desde columnas detectadas")
    _body_style(ws, 11, 10 + len(ranking), 1, len(headers))
    _apply_formats(ws, money_cols=[5], start_row=11, end_row=10 + len(ranking))
    if ranking:
        _add_table(ws, f"A10:F{10+len(ranking)}", "RealExecutiveRanking", "TableStyleMedium2")
    _section_label(ws, 15 + len(ranking), "Nota de alcance V1", 10)
    ws.cell(16 + len(ranking), 1, "Esta versión procesa archivos reales y conserva trazabilidad. La homologación semántica avanzada queda como mejora evolutiva.")
    ws.merge_cells(start_row=16+len(ranking), start_column=1, end_row=16+len(ranking), end_column=10)

    _write_real_comparativa(wb.create_sheet("Comparativa"), run.providers)
    palette = ["1F4E79", "0E6B3D", "7C3AED", "B54708", "344054"]
    for idx, p in enumerate(run.providers):
        sheet_name = f"Detalle - {p.name}"[:31]
        rows = canonical_rows_from_items(p.apu_items, include_market=True)
        _write_real_canonical_detail(wb.create_sheet(sheet_name), f"Detalle APU - {p.name}", rows, include_market=True, theme_color=palette[idx % len(palette)])
    _write_real_validaciones(wb.create_sheet("Validaciones"), run)
    _write_analisis_ia(wb.create_sheet("Análisis IA"), run)
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
    out = REPORTS_DIR / f"apu_v1_real_comparison_{run.run_id}.xlsx"
    wb.save(out)
    return out


def build_real_base_report(run: CanonicalRun) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    _setup_sheet(ws, "Presupuesto Base", "Presupuesto generado desde conceptos reales y expresado en el modelo canónico.", 10)
    total = _canonical_amount_from_concepts(run.base_concepts)
    _write_kpi(ws, 4, 1, "Monto detectado", total, "Conceptos base", "F8FAFC")
    exec_concepts = [c for c in run.base_concepts if getattr(c, "is_executable", False)]
    _write_kpi(ws, 4, 3, "Conceptos", len(exec_concepts), "Unidad + cantidad > 0", "F8FAFC")
    _write_kpi(ws, 4, 5, "Detalle", len(run.base_apu_items), "Matriz base", "E2F0D9")
    _write_kpi(ws, 4, 7, "Fuente", "Construdata", "Matrices/ref.", "F8FAFC")
    _write_kpi(ws, 4, 9, "Motor", "V1", "Data real", "FFF2CC")
    ws["A5"].number_format = MONEY_FMT
    comp = wb.create_sheet("Comparativa")
    _setup_sheet(comp, "Comparativa presupuesto base", "Conceptos reales leídos desde el catálogo base de ingeniería.", 8)
    headers = ["Código", "Concepto", "Unidad", "Cantidad", "P.U.", "Importe", "Estado", "Observación"]
    for c, h in enumerate(headers, 1): comp.cell(5, c, h)
    _header_style(comp, 5, 1, 8)
    base_rows = [c for c in run.base_concepts if getattr(c, "is_executable", False)]
    for r, cpt in enumerate(base_rows[:250], 6):
        vals = [cpt.code, cpt.description, cpt.unit, cpt.quantity or "", cpt.unit_price or "", cpt.amount or "", cpt.market_state or "Generado", f"Fila origen {cpt.source_row}"]
        for c, v in enumerate(vals, 1): comp.cell(r, c, v)
    if base_rows:
        _body_style(comp, 6, 5 + min(len(base_rows), 250), 1, 8)
        _apply_formats(comp, money_cols=[5,6], start_row=6, end_row=5+min(len(base_rows),250))
        _add_table(comp, f"A5:H{5+min(len(base_rows),250)}", "RealBaseComparativaTable", "TableStyleMedium2")
    _set_widths(comp, {"A":16,"B":58,"C":12,"D":12,"E":16,"F":18,"G":16,"H":28})
    detail_rows = canonical_rows_from_items(run.base_apu_items, include_market=False)
    _write_real_canonical_detail(wb.create_sheet("Detalle Base"), "Detalle Base - matriz canónica", detail_rows, include_market=False, theme_color="1F4E79")
    # Base-budget validations: show match confidence and missing matrices without
    # depending on provider-specific validation writer.
    val = wb.create_sheet("Validaciones")
    _setup_sheet(val, "Validaciones presupuesto base", "Trazabilidad de matches contra construdata_matrices.xlsx.", 8)
    headers_v = ["Severidad", "Tipo", "Mensaje", "Estado"]
    for c, h in enumerate(headers_v, 1): val.cell(5, c, h)
    _header_style(val, 5, 1, 4)
    rows_v = run.validations or [{"severity":"Baja", "type":"Base", "message":"Sin validaciones"}]
    for r, v in enumerate(rows_v[:500], 6):
        vals = [v.get("severity", "Info"), v.get("type", ""), v.get("message", ""), "Revisar" if v.get("severity") in {"Alta", "Media"} else "Informativo"]
        for c, value in enumerate(vals, 1): val.cell(r, c, value)
        val.cell(r, 1).fill = PatternFill("solid", fgColor=_status_fill(vals[0]))
    if rows_v:
        _body_style(val, 6, 5 + min(len(rows_v), 500), 1, 4)
        _add_table(val, f"A5:D{5+min(len(rows_v),500)}", "BaseValidacionesTable", "TableStyleMedium2")
    _set_widths(val, {"A":14,"B":24,"C":80,"D":18})
    _write_analisis_ia(wb.create_sheet("Análisis IA"), run)
    out = REPORTS_DIR / f"apu_v1_real_base_{run.run_id}.xlsx"
    wb.save(out)
    return out


@app.post("/api/comparisons/real-run")
async def comparison_real_run(
    projectName: str = Form("Comparación real"),
    provider_names: List[str] = Form(...),
    concept_files: List[UploadFile] = File(...),
    matrix_files: List[UploadFile] = File(...),
):
    if not (len(provider_names) == len(concept_files) == len(matrix_files)):
        raise HTTPException(status_code=400, detail="Cada proveedor debe tener nombre, archivo de conceptos y archivo matriz/APU")
    if len(provider_names) < 1:
        raise HTTPException(status_code=400, detail="Debe cargar al menos un proveedor")
    names = _safe_provider_names(provider_names)
    rid = run_id("REAL-CMP")
    run_dir = UPLOADS_DIR / rid
    catalog = ReferenceCatalog(DATA_DIR)
    providers: list[CanonicalProvider] = []
    for idx, name in enumerate(names):
        concepts_path = await _save_upload(concept_files[idx], run_dir / name)
        matrix_path = await _save_upload(matrix_files[idx], run_dir / name)
        provider = CanonicalProvider(name=name, concepts_file=concept_files[idx].filename or "", matrix_file=matrix_files[idx].filename or "")
        # Determine the real semantic role by workbook structure, not by the
        # upload field name. Some vendors send the PU/APU in the "concepts"
        # slot and the totals-by-concept workbook in the "matrix" slot.
        concept_role = classify_xlsx_role(concepts_path)
        matrix_role = classify_xlsx_role(matrix_path)
        concept_source = concepts_path
        matrix_source = matrix_path
        if concept_role == "apu_detail" and matrix_role != "apu_detail":
            concept_source = matrix_path
            matrix_source = concepts_path
            provider.validations.append({"severity":"Media", "type":"Rol de archivo", "message":"Se detectó que los archivos venían invertidos; se usó el archivo con totales como conceptos y el PU/APU como matriz."})
        try:
            provider.concepts = parse_concepts(concept_source)
            # Fallback: if strict catalog parsing finds no valid executable
            # concepts, try the other uploaded workbook before failing silently.
            if not [c for c in provider.concepts if _is_valid_comparativa_concept(c)]:
                alt = matrix_path if concept_source == concepts_path else concepts_path
                alt_concepts = parse_concepts(alt)
                if [c for c in alt_concepts if _is_valid_comparativa_concept(c)]:
                    provider.concepts = alt_concepts
                    provider.validations.append({"severity":"Media", "type":"Conceptos", "message":"Se usó el archivo alterno porque el primero no contenía conceptos válidos con unidad y cantidad > 0."})
        except Exception as exc:
            provider.validations.append({"severity":"Alta", "type":"Parser conceptos", "message":f"{type(exc).__name__}: {exc}"})
        try:
            provider.apu_items = parse_matrix(matrix_source, catalog)
            # Fallback symmetrical to the concept detection.
            if len(provider.apu_items) < 5:
                alt = concepts_path if matrix_source == matrix_path else matrix_path
                alt_items = parse_matrix(alt, catalog)
                if len(alt_items) > len(provider.apu_items):
                    provider.apu_items = alt_items
                    provider.validations.append({"severity":"Media", "type":"Matriz", "message":"Se usó el archivo alterno porque el primero no contenía una matriz/APU estructurada."})
            apply_provider_market_to_concepts(provider)
        except Exception as exc:
            provider.validations.append({"severity":"Alta", "type":"Parser matriz", "message":f"{type(exc).__name__}: {exc}"})
        providers.append(provider)
    run = CanonicalRun(run_id=rid, kind="comparison", project_name=projectName, providers=providers)
    REAL_RUNS[rid] = run
    report_path = build_real_comparison_report(run)
    REAL_REPORTS[rid] = report_path
    summary = _comparison_run_summary(run, report_path)
    BASE_RUN_SUMMARIES[rid] = summary
    return {
        "id": rid,
        "status": "COMPLETED_WITH_WARNINGS",
        "projectName": projectName,
        "providers": [{"name": p.name, "concepts": len(p.concepts), "apuItems": len(p.apu_items), "conceptsFile": p.concepts_file, "matrixFile": p.matrix_file} for p in providers],
        "downloadUrl": f"/api/real-runs/{rid}/report",
        "aiReportUrl": f"/api/real-runs/{rid}/ai-report",
        "summaryUrl": f"/api/real-runs/{rid}/summary",
        "summary": summary,
        "note": "Datos reales parseados con modelo canónico; homologación avanzada en evolución."
    }



def _comparison_run_summary(run: CanonicalRun, report_path: Path | None = None) -> dict[str, Any]:
    context = _run_analysis_context(run)
    analysis = generate_expert_ai_analysis(run)
    providers = context.get("providers", [])
    total_min = min([float(p.get("total_amount", 0) or 0) for p in providers], default=0)
    total_max = max([float(p.get("total_amount", 0) or 0) for p in providers], default=0)
    return {
        "runId": run.run_id,
        "type": "comparison",
        "projectName": run.project_name,
        "status": "COMPLETED_WITH_WARNINGS",
        "downloadUrl": f"/api/real-runs/{run.run_id}/report",
        "aiReportUrl": f"/api/real-runs/{run.run_id}/ai-report",
        "reportFile": report_path.name if report_path else "",
        "providersCount": len(run.providers or []),
        "bestProvider": context.get("best_provider", ""),
        "worstProvider": context.get("worst_provider", ""),
        "minAmount": round(total_min, 2),
        "maxAmount": round(total_max, 2),
        "economicSpreadPct": context.get("economic_spread_pct", 0),
        "providers": providers,
        "executiveFindings": [s.get("content", "") for s in analysis.get("sections", [])],
    }


async def _execute_base_budget_real(projectName: str, concepts_file: UploadFile):
    if not concepts_file.filename or not concepts_file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail=f"Archivo no permitido: {concepts_file.filename}. Solo se acepta .xlsx")
    rid = run_id("REAL-BASE")
    run_dir = UPLOADS_DIR / rid
    concepts_path = await _save_upload(concepts_file, run_dir)
    base_concepts = parse_base_concepts(concepts_path)
    base_apu_items, base_validations = generate_base_budget_from_concepts(base_concepts, DATA_DIR)
    run = CanonicalRun(run_id=rid, kind="base", project_name=projectName, base_concepts=base_concepts, base_apu_items=base_apu_items, validations=base_validations)
    REAL_RUNS[rid] = run
    report_path = build_real_base_report(run)
    REAL_REPORTS[rid] = report_path
    # Last real base report is also exposed through the legacy report URL for
    # old/cached UI bundles. This eliminates the demo workbook path entirely for
    # base budgets once a real run has been executed.
    REAL_REPORTS["LATEST_BASE"] = report_path
    summary = _base_run_summary(run, report_path, concepts_file.filename or "")
    BASE_RUN_SUMMARIES[rid] = summary
    BASE_RUN_SUMMARIES["LATEST_BASE"] = summary
    executable_count = len([c for c in base_concepts if getattr(c, "is_executable", False)])
    return {
        "id": rid,
        "status": summary.get("status", "COMPLETED_WITH_WARNINGS"),
        "projectName": projectName,
        "concepts": len(base_concepts),
        "executableConcepts": executable_count,
        "apuItems": len(base_apu_items),
        "validations": len(base_validations),
        "downloadUrl": f"/api/real-runs/{rid}/report",
        "aiReportUrl": f"/api/real-runs/{rid}/ai-report",
        "sourceFile": concepts_file.filename,
        "mode": "REAL",
        "summaryUrl": f"/api/real-runs/{rid}/summary",
        "summary": summary,
    }


@app.post("/api/base-budgets/real-run")
async def base_budget_real_run(projectName: str = Form("Presupuesto base real"), concepts_file: UploadFile = File(...), matrix_file: UploadFile | None = File(default=None)):
    return await _execute_base_budget_real(projectName, concepts_file)





def _sample_apu_ai_context() -> dict[str, Any]:
    """Small deterministic context used to test the external AI path without uploading files."""
    return {
        "kind": "base_budget",
        "project_name": "Prueba IA APU",
        "run_id": "AI-SMOKE-TEST",
        "total_amount": 22747686.09,
        "direct_cost": 18198148.87,
        "indirect_cost": 4549537.22,
        "indirect_pct": 25.0,
        "section_breakdown": {
            "materials": 6249711.70,
            "labor": 2347014.99,
            "equipment": 9576096.32,
            "basics": 25325.87,
            "direct": 18198148.87,
            "indirect": 4549537.22,
            "total_service": 22747686.09,
        },
        "concepts_read": 83,
        "concepts_executable": 61,
        "detail_rows": 1458,
        "matched_matrices": 57,
        "unmatched_matrices": 4,
        "coverage_pct": 93.4,
        "validations_count": 4,
        "validation_severity_counts": {"Media": 4},
        "validation_samples": [
            {"severity": "Media", "type": "Matriz", "message": "Concepto sin matriz directa; requiere validación técnica."}
        ],
        "top_concepts": [
            {"code": "DD06", "description": "Partida dominante", "unit": "PZA", "quantity": 1, "unit_price": 12808910.91, "amount": 12808910.91, "weight_pct": 56.3, "state": "Match Construdata"},
            {"code": "EM01", "description": "Partida electromecánica", "unit": "PZA", "quantity": 1, "unit_price": 2099259.68, "amount": 2099259.68, "weight_pct": 9.2, "state": "Match Construdata"},
            {"code": "EM11", "description": "Partida relevante", "unit": "PZA", "quantity": 1, "unit_price": 1228824.33, "amount": 1228824.33, "weight_pct": 5.4, "state": "Match Construdata"},
        ],
        "top_overcost_items": [
            {"code": "EQ-01", "description": "Equipo con presión económica", "section": "EQUIPO Y HERRAMIENTA", "amount": 480000.0, "market_amount": 390000.0, "overcost": 90000.0, "overcost_pct": 23.08, "state": "Match Construdata"}
        ],
    }


def _execute_ai_analysis_for_context(context: dict[str, Any]) -> dict[str, Any]:
    status = _analysis_feature_status()
    sections = _call_external_ai_analysis(context)
    mode = status.get("mode")
    if not sections:
        sections = _local_expert_analysis(context)
        if status.get("enabled") and status.get("mode") == "EXTERNAL_AI":
            mode = "LOCAL_EXPERT_FALLBACK_AI_ERROR"
    return {
        "mode": mode,
        "provider": status.get("provider"),
        "model": _configured_ai_model(),
        "enabled": bool(status.get("enabled")),
        "hasKey": bool(status.get("hasKey")),
        "error": AI_ANALYSIS_LAST_ERROR,
        "sections": sections,
        "contextKeys": list(context.keys()),
    }


@app.get("/api/ai-analysis/status")
def ai_analysis_status():
    status = _analysis_feature_status()
    # Never expose secrets, only whether the configured provider has a key.
    return status




@app.post("/api/ai-analysis/test")
def ai_analysis_test(payload: dict[str, Any] | None = None):
    """Smoke test for the AI analysis layer.

    With ENABLE_AI_ANALYSIS=1 and a valid provider API key, this endpoint calls
    the external model using either the supplied payload or a compact APU sample
    context. Without a key, it returns the local expert fallback and explains the
    mode, so the integration can be verified before running a full workbook.
    """
    context = payload if payload else _sample_apu_ai_context()
    if not isinstance(context, dict):
        raise HTTPException(status_code=400, detail="El payload debe ser un JSON object")
    return _execute_ai_analysis_for_context(context)



def _html_money(value: Any) -> str:
    try:
        return f"${float(value or 0):,.2f}"
    except Exception:
        return "$0.00"


def _html_pct(value: Any) -> str:
    try:
        return f"{float(value or 0):,.2f}%"
    except Exception:
        return "0.00%"


def _html_escape(value: Any) -> str:
    return html.escape(str(value if value is not None else ""))


def _html_kpi(title: str, value: str, subtitle: str = "") -> str:
    return '<div class="kpi"><div class="kpi-t">{}</div><div class="kpi-v">{}</div><div class="kpi-s">{}</div></div>'.format(_html_escape(title), _html_escape(value), _html_escape(subtitle))


def _html_table(headers: list[str], rows: list[list[Any]]) -> str:
    head = ''.join('<th>{}</th>'.format(_html_escape(h)) for h in headers)
    body_rows = []
    for row in rows:
        cells = ''.join('<td>{}</td>'.format(_html_escape(v)) for v in row)
        body_rows.append('<tr>{}</tr>'.format(cells))
    body = ''.join(body_rows)
    return '<div class="table-wrap"><table><thead><tr>{}</tr></thead><tbody>{}</tbody></table></div>'.format(head, body)


def _html_bar(label: str, value: float, total: float) -> str:
    pct = (value / total * 100) if total else 0
    pct = max(0, min(100, pct))
    return '<div class="bar"><div><b>{}</b><span>{} / {}</span></div><i><em style="width:{:.2f}%"></em></i></div>'.format(_html_escape(label), _html_money(value), _html_pct(pct), pct)



def _html_status_chip(status: str) -> str:
    s = str(status or "REVIEW").upper()
    cls = "ok" if s in {"OK", "LOW"} else "bad" if s in {"CRITICAL", "HIGH", "HIGH_RISK"} else "warn"
    return f"<span class='chip {cls}'>{_html_escape(_status_label(s))}</span>"


def _html_value(value: Any, fmt: str | None = None) -> str:
    return _html_escape(_diag_value(value, fmt))


def _html_diag_table(headers: list[str], rows: list[list[Any]], classes: str = "") -> str:
    head = ''.join(f'<th>{_html_escape(h)}</th>' for h in headers)
    if not rows:
        body = f'<tr><td colspan="{len(headers)}" class="muted">Sin datos calculados para esta sección</td></tr>'
    else:
        body = ''
        for row in rows:
            body += '<tr>' + ''.join(f'<td>{v if isinstance(v, str) and v.startswith("<") else _html_escape(v)}</td>' for v in row) + '</tr>'
    return f'<div class="table-wrap {classes}"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def _run_ai_report_html(run: CanonicalRun, summary: dict[str, Any] | None = None) -> str:
    diag = generate_professional_diagnostic(run, summary)
    headline = diag.get("headline", {})
    decision = diag.get("final_decision", {})
    ctx = diag.get("_context", {})
    title = "Diagnóstico profesional"
    run_label = _diagnostic_run_label(headline.get("run_type", ctx.get("run_type", "")))
    status = str(headline.get("general_status", "REVIEW") or "REVIEW").upper()
    traffic = headline.get("traffic_light", _traffic_from_status(status))
    executive = headline.get("executive_line", "Revisar partidas de mayor impacto y referencias sin trazabilidad plena.")

    kpi_cards = ''
    for k in diag.get("kpis", [])[:13]:
        kpi_cards += "<article class='kpi'><span>{}</span><strong>{}</strong><small>{}</small></article>".format(_html_escape(k.get("label", "")), _html_value(k.get("value"), k.get("format")), _html_escape(k.get("status", "")))

    sec_rows = []
    for s in diag.get("section_summary", []):
        sec_rows.append([s.get("section"), _html_money(s.get("contractor_amount",0)), _html_money(s.get("market_amount",0)) if s.get("market_amount") is not None else "N/A", _html_money(s.get("difference_amount",0)) if s.get("difference_amount") is not None else "N/A", _html_pct(s.get("difference_pct",0)) if s.get("difference_pct") is not None else "N/A", _html_pct(s.get("participation_pct",0)), _html_status_chip(s.get("status")), s.get("comment", "")])
    section_table = _html_diag_table(["Sección", "Importe", "Mercado", "Dif. $", "Dif. %", "% total", "Estado", "Comentario"], sec_rows)

    bars = ''
    total = float(ctx.get("total_amount", 0) or 0)
    for s in diag.get("section_summary", []):
        amount = float(s.get("contractor_amount", 0) or 0)
        if amount:
            bars += _html_bar(str(s.get("section", "")), amount, total)

    def top_item_rows(items: list[dict[str, Any]], kind: str) -> list[list[Any]]:
        rows = []
        for i in (items or [])[:10]:
            if kind == 'labor':
                rows.append([i.get("code"), i.get("description"), i.get("unit"), i.get("operator"), i.get("quantity"), _html_money(i.get("contractor_unit_price",0)), _html_money(i.get("market_unit_price",0)) if i.get("market_unit_price") is not None else "N/A", _html_money(i.get("difference_amount",0)) if i.get("difference_amount") is not None else "N/A", _html_pct(i.get("difference_pct",0)) if i.get("difference_pct") is not None else "N/A", _html_money(i.get("impact_amount",0)), i.get("recommended_action", "")])
            else:
                rows.append([i.get("code"), i.get("description"), i.get("unit"), i.get("quantity"), _html_money(i.get("contractor_unit_price",0)), _html_money(i.get("market_unit_price",0)) if i.get("market_unit_price") is not None else "N/A", _html_money(i.get("difference_amount",0)) if i.get("difference_amount") is not None else "N/A", _html_pct(i.get("difference_pct",0)) if i.get("difference_pct") is not None else "N/A", _html_money(i.get("impact_amount",0)), i.get("market_reference", ""), i.get("recommended_action", "")])
        return rows

    materials_table = _html_diag_table(["Código", "Descripción", "Unidad", "Cant.", "P.U.", "P.U. mercado", "Dif. $", "Dif. %", "Importe", "Referencia", "Acción"], top_item_rows(diag.get("top_materials", []), 'materials'))
    labor_table = _html_diag_table(["Código", "Descripción", "Unidad", "Op.", "Rend./Cant.", "P.U.", "P.U. mercado", "Dif. $", "Dif. %", "Importe", "Acción"], top_item_rows(diag.get("top_labor", []), 'labor'))
    equipment_table = _html_diag_table(["Código", "Descripción", "Unidad", "Cant.", "P.U.", "P.U. mercado", "Dif. $", "Dif. %", "Importe", "Referencia", "Acción"], top_item_rows(diag.get("top_equipment", []), 'equipment'))

    concept_rows = []
    for c in diag.get("critical_concepts", [])[:10]:
        concept_rows.append([c.get("concept_code"), c.get("description"), _html_money(c.get("contractor_amount",0)), _html_money(c.get("market_amount",0)) if c.get("market_amount") is not None else "N/A", _html_money(c.get("difference_amount",0)) if c.get("difference_amount") is not None else "N/A", _html_pct(c.get("difference_pct",0)) if c.get("difference_pct") is not None else "N/A", _html_pct(c.get("participation_pct",0)), c.get("probable_cause", ""), c.get("priority_action", "")])
    concepts_table = _html_diag_table(["Partida", "Descripción", "Importe", "Mercado", "Dif. $", "Dif. %", "% total", "Causa probable", "Acción"], concept_rows)

    alert_rows = []
    for a in diag.get("market_alerts", [])[:20]:
        sev = str(a.get("severity", "MEDIUM")).upper()
        chip = "<span class='chip {}'>{}</span>".format("bad" if sev == "HIGH" else "warn" if sev == "MEDIUM" else "ok", _html_escape(sev))
        alert_rows.append([chip, a.get("alert_type"), a.get("item"), a.get("section"), _html_money(a.get("contractor_value",0)) if a.get("contractor_value") is not None else "N/A", _html_money(a.get("market_value",0)) if a.get("market_value") is not None else "N/A", _html_pct(a.get("deviation_pct",0)) if a.get("deviation_pct") is not None else "N/A", a.get("analyst_check")])
    alerts_table = _html_diag_table(["Severidad", "Tipo", "Partida/Insumo", "Sección", "Valor", "Mercado", "Desviación", "Qué revisar"], alert_rows)

    plan_rows = []
    for pitem in diag.get("analyst_review_plan", [])[:10]:
        plan_rows.append([pitem.get("priority"), pitem.get("what_to_review"), pitem.get("why_it_matters"), pitem.get("where_to_check"), pitem.get("decision_needed")])
    plan_table = _html_diag_table(["Prioridad", "Qué revisar", "Por qué importa", "Dónde buscar", "Decisión requerida"], plan_rows)

    pd = diag.get("professional_diagnosis", {})
    diag_table = _html_diag_table(["Tema", "Lectura"], [["Riesgo principal", pd.get("risk_summary")], ["Driver de costo", pd.get("main_cost_driver")], ["Trazabilidad", pd.get("market_traceability")], ["Recomendación", pd.get("recommendation")]])
    final_table = _html_diag_table(["Dictamen", "Motivo principal", "Próxima acción", "Prioridad"], [[_status_label(status), decision.get("main_reason"), decision.get("next_action"), decision.get("priority")]])

    css = """
    :root{--bg:#07111f;--card:#0f1d33;--card2:#132642;--line:#263850;--text:#eaf1fb;--muted:#9fb0c7;--blue:#67a8ff;--green:#4ade80;--yellow:#fbbf24;--red:#fb7185}*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at top,#142844,#07111f 60%);color:var(--text);font-family:Inter,Segoe UI,Arial,sans-serif}.wrap{max-width:1440px;margin:auto;padding:34px}.hero,.panel{background:linear-gradient(135deg,rgba(19,38,66,.96),rgba(15,29,51,.96));border:1px solid var(--line);border-radius:24px;padding:24px;box-shadow:0 22px 70px rgba(0,0,0,.25)}.hero{display:grid;grid-template-columns:1.3fr .7fr;gap:20px;align-items:center}h1{margin:6px 0 8px;font-size:36px}h2{font-size:20px;margin:0 0 14px}p{color:#c7d2e5;line-height:1.56}.eyebrow{font-size:12px;text-transform:uppercase;letter-spacing:.14em;color:#a8c7ff;font-weight:900}.chip{display:inline-flex;align-items:center;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:900}.chip.ok{background:rgba(74,222,128,.16);color:#86efac}.chip.warn{background:rgba(251,191,36,.16);color:#fde68a}.chip.bad{background:rgba(251,113,133,.16);color:#fecdd3}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:18px 0}.kpi{background:rgba(8,21,37,.9);border:1px solid var(--line);border-radius:18px;padding:16px}.kpi span{display:block;font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.kpi strong{display:block;font-size:23px;margin-top:8px}.kpi small{display:block;color:var(--muted);margin-top:5px}.grid{display:grid;grid-template-columns:.85fr 1.15fr;gap:18px;margin-top:18px}.stack{display:grid;gap:18px;margin-top:18px}.bar{margin:12px 0}.bar div{display:flex;justify-content:space-between;color:var(--muted);font-size:13px}.bar b{color:var(--text)}.bar i{display:block;height:12px;background:#081525;border:1px solid var(--line);border-radius:999px;overflow:hidden;margin-top:6px}.bar em{display:block;height:100%;background:linear-gradient(90deg,var(--blue),var(--green))}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:16px}table{width:100%;border-collapse:collapse;font-size:13px}th,td{border-bottom:1px solid var(--line);padding:10px;text-align:left;vertical-align:top}th{background:rgba(103,168,255,.13);color:#dbeafe;white-space:nowrap}tr:last-child td{border-bottom:0}.muted{color:var(--muted)}.actions{display:flex;gap:12px;margin-top:18px}.btn{background:#2563eb;color:white;text-decoration:none;border-radius:12px;padding:12px 16px;font-weight:900}.btn.secondary{background:#17243a;border:1px solid var(--line)}.decision{border-left:5px solid var(--yellow)}.decision.green{border-left-color:var(--green)}.decision.red{border-left-color:var(--red)}@media(max-width:1000px){.hero,.grid,.kpis{grid-template-columns:1fr}.wrap{padding:18px}}
    """
    decision_cls = "green" if status == "OK" else "red" if status == "CRITICAL" else ""
    return f"""<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{_html_escape(title)}</title><style>{css}</style></head><body><main class='wrap'><section class='hero'><div><div class='eyebrow'>{_html_escape(run_label)}</div><h1>{_html_escape(title)}</h1><p>{_html_escape(executive)}</p><div class='actions'><a class='btn' href='/api/real-runs/{_html_escape(run.run_id)}/report'>Descargar Excel</a><a class='btn secondary' href='#plan'>Ver plan de revisión</a></div></div><div class='panel decision {decision_cls}'><h2>Dictamen</h2>{_html_status_chip(status)}<p><b>{_html_escape(decision.get('next_action','Revisar prioridades del diagnóstico.'))}</b></p><p class='muted'>{_html_escape(decision.get('main_reason',''))}</p></div></section><section class='kpis'>{kpi_cards}</section><div class='grid'><section class='panel'><h2>Participación por sección</h2>{bars or '<p class="muted">Sin desglose disponible.</p>'}</section><section class='panel'><h2>Resumen por secciones APU</h2>{section_table}</section></div><section class='stack'><section class='panel'><h2>Top 10 materiales por impacto</h2>{materials_table}</section><section class='panel'><h2>Top 10 mano de obra / cuadrillas</h2>{labor_table}</section><section class='panel'><h2>Top 10 maquinaria / equipo</h2>{equipment_table}</section><section class='panel'><h2>Partidas críticas del catálogo</h2>{concepts_table}</section><section class='panel'><h2>Alertas contra mercado</h2>{alerts_table}</section><section class='panel' id='plan'><h2>Plan de revisión para el analista</h2>{plan_table}</section><section class='panel'><h2>Diagnóstico profesional breve</h2>{diag_table}</section><section class='panel'><h2>Conclusión ejecutiva</h2>{final_table}</section></section></main></body></html>"""

@app.get("/api/real-runs/{run_id}/summary")
def real_run_summary(run_id: str):
    summary = BASE_RUN_SUMMARIES.get(run_id)
    if not summary:
        raise HTTPException(status_code=404, detail="Resumen de corrida no encontrado")
    return summary


@app.get("/api/real-runs/latest-base/summary")
def latest_base_run_summary():
    summary = BASE_RUN_SUMMARIES.get("LATEST_BASE")
    if not summary:
        raise HTTPException(status_code=404, detail="No hay presupuesto base real generado en esta sesión")
    return summary


@app.get("/api/real-runs/{run_id}/ai-report", response_class=HTMLResponse)
def real_run_ai_report(run_id: str):
    run = REAL_RUNS.get(run_id)
    if not run:
        raise HTTPException(status_code=404, detail="Corrida no encontrada")
    summary = BASE_RUN_SUMMARIES.get(run_id)
    return HTMLResponse(_run_ai_report_html(run, summary))


@app.get("/api/real-runs/{run_id}/report")
def real_run_report(run_id: str):
    path = REAL_REPORTS.get(run_id)
    if not path or not path.exists():
        raise HTTPException(status_code=404, detail="Reporte no encontrado")
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Static frontend must be mounted last.
app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
