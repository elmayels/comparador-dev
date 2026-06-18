from __future__ import annotations

import json
import os
import time
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
FRONTEND_DIR = ROOT / "frontend"
RUNTIME_DIR = ROOT / "backend" / "runtime"
REPORTS_DIR = RUNTIME_DIR / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="APU Canonical Platform V0", version="0.0.1")
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


@app.post("/api/base-budgets/mock-run")
async def base_budget_mock_run(projectName: str = Form("Presupuesto base demo"), files: List[UploadFile] = File(default=[])):
    names = _validate_xlsx_uploads(files)
    refs = _list_data_references()
    matrix = next((r for r in refs if r["kind"] == "CONSTRUDATA_MATRICES_BASE_BUDGET"), None)
    return {
        "id": f"BB-{int(time.time())}",
        "projectName": projectName,
        "status": "COMPLETED_WITH_WARNINGS",
        "uploadedFiles": names,
        "construdataMatrix": matrix["name"] if matrix else None,
        "kpis": {
            "estimatedTotal": 56520000,
            "totalConcepts": 128,
            "matchedConcepts": 104,
            "reviewConcepts": 17,
            "unmatchedConcepts": 7,
            "duplicatedConcepts": 3,
            "estimationRisk": "Medio",
        },
        "findings": [
            "La matriz base se generó como presupuesto independiente; no se asocia obligatoriamente a una licitación.",
            "17 conceptos tienen match medio y requieren revisión técnica antes de usar el presupuesto.",
            "7 conceptos no encontraron referencia directa en matrices Construdata.",
        ],
    }


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


def _write_analisis_ia(ws):
    _setup_sheet(ws, "Análisis IA", "Narrativa ejecutiva generada sobre datos calculados. La IA no inventa importes ni modifica cálculos.", 6)
    _set_widths(ws, {"A": 24, "B": 95, "C": 18, "D": 18, "E": 18, "F": 18})
    sections = [
        ["Resumen ejecutivo", "La comparación identifica una propuesta competitiva, pero con partidas críticas que requieren validación técnica antes de negociar o adjudicar."],
        ["Riesgos", "El riesgo principal se concentra en partidas con alto impacto económico y en diferencias de insumos de mano de obra y materiales."],
        ["Recomendaciones", "Solicitar desglose APU, validar rendimiento, comparar porcentajes financieros y negociar partidas con mayor impacto económico."],
        ["Preguntas al contratista", "¿Qué alcance está incluido en las partidas bajo mercado? ¿Qué tarifas y rendimientos soportan mano de obra? ¿Qué fichas técnicas respaldan los insumos críticos?"],
        ["Limitaciones", "V0 usa datos mockeados. El motor real debe leer la matriz/APU del contratista y poblar Detalle APU desde esa fuente."],
    ]
    headers = ["Sección", "Contenido"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, 2)
    for r, row in enumerate(sections, 6):
        ws.cell(r, 1, row[0]).font = Font(bold=True, color=BRAND["navy"])
        ws.cell(r, 2, row[1])
        ws.row_dimensions[r].height = 45
    _body_style(ws, 6, 10, 1, 2)
    _add_table(ws, "A5:B10", "AnalisisIATable", "TableStyleMedium2")


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
    provider_names = [p.strip() for p in providers.split(",") if p.strip()] if providers else None
    path = build_report(kind, provider_names)
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# Static frontend must be mounted last.
app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
