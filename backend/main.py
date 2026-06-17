from __future__ import annotations

import json
import os
import time
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
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


def _sample_contractors():
    return [
        {"rank": 1, "contractor": "Contratista B", "amount": 1180000, "diff_min": 0.00, "avg_dev": -0.078, "risk": "Bajo", "traffic": "Verde"},
        {"rank": 2, "contractor": "Contratista A", "amount": 1250000, "diff_min": 0.059, "avg_dev": 0.024, "risk": "Medio", "traffic": "Amarillo"},
        {"rank": 3, "contractor": "Contratista C", "amount": 1410000, "diff_min": 0.195, "avg_dev": 0.143, "risk": "Alto", "traffic": "Rojo"},
    ]


def _write_resumen_ejecutivo(ws, mode: str):
    _setup_sheet(ws, "Resumen Ejecutivo", "Vista gerencial: KPIs, ranking, hallazgos y navegación interna del reporte.", 10)
    _set_widths(ws, {"A": 20, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18, "J": 18})
    ws.freeze_panes = "A12"

    _write_kpi(ws, 4, 1, "Mejor oferta", 1180000, "Contratista B", "F8FAFC")
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
    for r, item in enumerate(_sample_contractors(), 11):
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
    links = [
        ("Comparativa", "#'Comparativa'!A1"),
        ("Detalle APU", "#'Detalle APU'!A1"),
        ("Partidas Críticas", "#'Partidas Críticas'!A1"),
        ("Insumos Críticos", "#'Insumos Críticos'!A1"),
        ("Validaciones", "#'Validaciones'!A1"),
        ("Parámetros", "#'Parámetros'!A1"),
    ]
    for idx, (text, target) in enumerate(links, 1):
        cell = ws.cell(24, idx, text)
        cell.hyperlink = target
        cell.style = "Hyperlink"
        cell.alignment = Alignment(horizontal="center")


def _write_comparativa_professional(ws):
    _setup_sheet(ws, "Comparativa", "Tabla económica principal. Vista normalizada, filtrable y auditable por partida/contratista.", 16)
    _set_widths(ws, {
        "A": 14, "B": 42, "C": 16, "D": 11, "E": 11, "F": 16, "G": 16, "H": 16,
        "I": 16, "J": 16, "K": 16, "L": 12, "M": 16, "N": 12, "O": 16, "P": 30
    })
    headers = ["Partida", "Descripción", "Familia", "Unidad", "Cantidad", "Contratista", "P.U. ofertado", "Importe ofertado", "P.U. referencia", "Importe referencia", "Diferencia", "Desviación %", "Impacto", "Peso %", "Estado", "Observación"]
    for c, h in enumerate(headers, 1):
        ws.cell(5, c, h)
    _header_style(ws, 5, 1, len(headers), fill="1F4E79")
    rows = [
        ["FLEX41.11", "Instalación de bomba centrífuga", "Mecánica", "PZA", 1, "Contratista A", 31524.55, "=E6*G6", 6020.40, "=E6*I6", "=H6-J6", "=K6/J6", "=MAX(K6,0)", "=H6/SUM($H$6:$H$14)", "Sobrecosto", "Validar alcance y matriz APU"],
        ["FLEX41.12", "Guarda protectora para bomba", "Mecánica", "PZA", 1, "Contratista A", 24937.16, "=E7*G7", 6020.40, "=E7*I7", "=H7-J7", "=K7/J7", "=MAX(K7,0)", "=H7/SUM($H$6:$H$14)", "Sobrecosto", "Revisar precio unitario"],
        ["FLEX41.29", "Tubería sanitaria acero inoxidable 316L 3 pulgadas", "Tubería", "M", 41, "Contratista A", 3496.73, "=E8*G8", 6020.40, "=E8*I8", "=H8-J8", "=K8/J8", "=MAX(K8,0)", "=H8/SUM($H$6:$H$14)", "Bajo mercado", "Confirmar que no haya omisión de alcance"],
        ["FLEX41.40", "Tubería sanitaria acero inoxidable 316L 2 1/2 pulgadas", "Tubería", "M", 38, "Contratista B", 3496.73, "=E9*G9", 6020.40, "=E9*I9", "=H9-J9", "=K9/J9", "=MAX(K9,0)", "=H9/SUM($H$6:$H$14)", "Bajo mercado", "Validar especificación técnica"],
        ["FLEX41.62", "Ensayo de boroscopia en tuberías", "Ensayos", "SERV", 1, "Contratista C", 75226.36, "=E10*G10", 6020.40, "=E10*I10", "=H10-J10", "=K10/J10", "=MAX(K10,0)", "=H10/SUM($H$6:$H$14)", "Sobrecosto", "Negociar o justificar servicio"],
        ["CIV-001", "Concreto f'c 250", "Concreto", "m3", 40, "Contratista C", 2850000, "=E11*G11", 2300000, "=E11*I11", "=H11-J11", "=K11/J11", "=MAX(K11,0)", "=H11/SUM($H$6:$H$14)", "Sobrecosto", "Alto impacto"],
        ["CIV-002", "Acero de refuerzo", "Acero", "kg", 1800, "Contratista A", 42000, "=E12*G12", 35000, "=E12*I12", "=H12-J12", "=K12/J12", "=MAX(K12,0)", "=H12/SUM($H$6:$H$14)", "Sobrecosto", "Revisar insumo base"],
        ["ELE-001", "Luminarias LED", "Instalaciones", "pza", 35, "Contratista C", 1200000, "=E13*G13", 980000, "=E13*I13", "=H13-J13", "=K13/J13", "=MAX(K13,0)", "=H13/SUM($H$6:$H$14)", "Sobrecosto", "Solicitar ficha técnica"],
        ["GEN-001", "Partida sin referencia directa", "Especial", "gl", 1, "Contratista B", 15000000, "=E14*G14", 0, "=E14*I14", "=IF(I14=0,0,H14-J14)", "=IF(I14=0,0,K14/J14)", "=IF(I14=0,H14,MAX(K14,0))", "=H14/SUM($H$6:$H$14)", "Sin referencia", "Requiere cotización externa"],
    ]
    for r, row in enumerate(rows, 6):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    _body_style(ws, 6, 14, 1, len(headers))
    _apply_formats(ws, money_cols=[7, 8, 9, 10, 11, 13], pct_cols=[12, 14], int_cols=[5], start_row=6, end_row=14)
    _add_table(ws, "A5:P14", "ComparativaTable", "TableStyleMedium2")
    ws.auto_filter.ref = "A5:P14"
    ws.freeze_panes = "A6"
    # Conditional formatting
    ws.conditional_formatting.add("L6:L14", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    ws.conditional_formatting.add("M6:M14", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5", showValue=True))
    ws.conditional_formatting.add("O6:O14", FormulaRule(formula=['$O6="Sobrecosto"'], fill=PatternFill("solid", fgColor="FCE4D6")))
    ws.conditional_formatting.add("O6:O14", FormulaRule(formula=['$O6="Bajo mercado"'], fill=PatternFill("solid", fgColor="E2F0D9")))
    ws.conditional_formatting.add("O6:O14", FormulaRule(formula=['$O6="Sin referencia"'], fill=PatternFill("solid", fgColor="E7E6E6")))
    ws["A16"] = "Regla canónica"
    ws["A16"].font = Font(bold=True, color=BRAND["navy"])
    ws["A17"] = "Comparativa resume importes y desviaciones. El Detalle APU explica de dónde viene cada precio desde la matriz propia del contratista."
    ws.merge_cells("A17:P17")
    ws["A17"].alignment = Alignment(wrap_text=True)


def _write_detalle_apu_professional(ws):
    _setup_sheet(ws, "Detalle APU", "Reconstrucción ordenada de la matriz/APU del contratista, enriquecida con referencias granulares de data.", 18)
    _set_widths(ws, {
        "A": 16, "B": 26, "C": 16, "D": 38, "E": 12, "F": 10, "G": 13, "H": 16, "I": 18,
        "J": 16, "K": 16, "L": 14, "M": 14, "N": 18, "O": 18, "P": 16, "Q": 18, "R": 34
    })
    headers = ["Contratista", "Partida", "Bloque APU", "Concepto / insumo", "Unidad", "Op.", "Cantidad", "P.U. contratista", "Importe contratista", "P.U. referencia", "Importe referencia", "Diferencia", "Desv. %", "Fuente referencia", "Estado match", "Tipo fila", "Alerta", "Regla / observación"]
    for c, h in enumerate(headers, 1):
        ws.cell(5, c, h)
    _header_style(ws, 5, 1, len(headers), fill="1F4E79")
    rows = [
        ["Contratista A", "FLEX41.11", "PARTIDA", "Instalación de bomba centrífuga", "PZA", "*", 1, 31524.55, "=G6*H6", 6020.40, "=G6*J6", "=I6-K6", "=IF(K6=0,0,L6/K6)", "Comparativa", "Leído", "PARTIDA", "Revisar", "Partida principal leída desde archivo del contratista"],
        ["Contratista A", "FLEX41.11", "MATERIALES", "Consumibles para corte y soldadura", "%", "%", 0.08, 22453.38, "=G7*H7", 22453.38, "=G7*J7", "=I7-K7", "=IF(K7=0,0,L7/K7)", "data/materiales", "Match medio", "INSUMO", "OK", "Porcentaje sobre base de materiales"],
        ["Contratista A", "FLEX41.11", "MANO DE OBRA", "Supervisor Obra", "JOR", "*", 0.10, 2643.19, "=G8*H8", 1757.33, "=G8*J8", "=I8-K8", "=IF(K8=0,0,L8/K8)", "data/mano_obra", "Match alto", "INSUMO", "Revisar", "Precio MO superior a referencia"],
        ["Contratista A", "FLEX41.11", "MANO DE OBRA", "Oficial soldador/argonero", "JOR", "*", 1, 2120.45, "=G9*H9", 1326.14, "=G9*J9", "=I9-K9", "=IF(K9=0,0,L9/K9)", "data/mano_obra", "Match alto", "INSUMO", "Revisar", "Validar rendimiento y tarifa"],
        ["Contratista A", "FLEX41.11", "EQUIPO Y HERRAMIENTA", "Herramienta menor", "%", "%", 0.09, "=SUM(I8:I9)", "=G10*H10", "=SUM(K8:K9)", "=G10*J10", "=I10-K10", "=IF(K10=0,0,L10/K10)", "data/porcentajes", "Referencia", "PORCENTAJE", "OK", "%HERR sobre mano de obra"],
        ["Contratista A", "FLEX41.11", "EQUIPO Y HERRAMIENTA", "Equipo de protección personal", "%", "%", 0.09, "=SUM(I8:I9)", "=G11*H11", "=SUM(K8:K9)", "=G11*J11", "=I11-K11", "=IF(K11=0,0,L11/K11)", "data/porcentajes", "Referencia", "PORCENTAJE", "OK", "%EPP sobre mano de obra"],
        ["Contratista A", "FLEX41.11", "SUBTOTAL", "Costo directo", "", "", 1, "=SUM(I7:I11)", "=H12", "=SUM(K7:K11)", "=J12", "=I12-K12", "=IF(K12=0,0,L12/K12)", "Cálculo", "Calculado", "SUBTOTAL", "OK", "Suma materiales + MO + equipo"],
        ["Contratista A", "FLEX41.11", "FINANCIERO", "Costo indirecto", "%", "%", 0.25, "=I12", "=G13*H13", "=K12", "=G13*J13", "=I13-K13", "=IF(K13=0,0,L13/K13)", "data/porcentajes", "Referencia", "PORCENTAJE", "Revisar", "Comparar porcentaje aplicado vs rango permitido"],
        ["Contratista A", "FLEX41.11", "TOTAL", "Total costo unitario", "", "", 1, "=I12+I13", "=H14", "=K12+K13", "=J14", "=I14-K14", "=IF(K14=0,0,L14/K14)", "Cálculo", "Calculado", "TOTAL", "Reconciliar", "Debe reconciliar contra P.U. en Comparativa"],
    ]
    for r, row in enumerate(rows, 6):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
        # Fill by type
        row_type = row[15]
        fill = "FFFFFF"
        if row_type == "PARTIDA": fill = "EAF2FF"
        elif row_type in {"SUBTOTAL", "TOTAL"}: fill = "FFF2CC"
        elif row_type == "PORCENTAJE": fill = "F8FAFC"
        for c in range(1, len(headers)+1):
            ws.cell(r, c).fill = PatternFill("solid", fgColor=fill)
            ws.cell(r, c).border = _thin_border("EAECF0")
            ws.cell(r, c).alignment = Alignment(vertical="top", wrap_text=True)
            if row_type in {"PARTIDA", "SUBTOTAL", "TOTAL"}:
                ws.cell(r, c).font = Font(bold=True, color=BRAND["text"])
    _apply_formats(ws, money_cols=[8, 9, 10, 11, 12], pct_cols=[7, 13], start_row=6, end_row=14)
    _add_table(ws, "A5:R14", "DetalleApuTable", "TableStyleMedium4")
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = "A5:R14"
    ws.conditional_formatting.add("M6:M14", ColorScaleRule(start_type="min", start_color="E2F0D9", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="FCE4D6"))
    ws.conditional_formatting.add("L6:L14", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5", showValue=True))
    # group detail rows visually
    for r in range(7, 12):
        ws.row_dimensions[r].outlineLevel = 1
    for r in range(13, 14):
        ws.row_dimensions[r].outlineLevel = 1
    ws["A17"] = "Regla canónica"
    ws["A17"].font = Font(bold=True, color=BRAND["navy"])
    ws["A18"] = "Esta hoja NO se crea copiando matrices Construdata. V1 debe leer el APU/matriz del contratista, identificar bloques y cruzar insumos contra materiales, mano de obra, maquinaria y porcentajes desde data."
    ws.merge_cells("A18:R18")
    ws["A18"].alignment = Alignment(wrap_text=True)


def _write_partidas_criticas(ws):
    _setup_sheet(ws, "Partidas Críticas", "Priorización de partidas por impacto económico, desviación y riesgo de negociación.", 12)
    _set_widths(ws, {"A": 10, "B": 16, "C": 16, "D": 36, "E": 18, "F": 16, "G": 14, "H": 14, "I": 18, "J": 26, "K": 22, "L": 18})
    headers = ["Prioridad", "Contratista", "Partida", "Descripción", "Familia", "Impacto", "Desv. %", "Peso %", "Motivo", "Acción sugerida", "Estado revisión", "Responsable"]
    for c, h in enumerate(headers, 1): ws.cell(5, c, h)
    _header_style(ws, 5, 1, len(headers))
    rows = [
        [1, "Contratista C", "CIV-001", "Concreto f'c 250", "Concreto", 95000, 0.24, 0.18, "Alto impacto", "Solicitar desglose APU", "Pendiente", "Ingeniería"],
        [2, "Contratista A", "CIV-002", "Acero de refuerzo", "Acero", 62000, 0.20, 0.12, "Precio sobre referencia", "Negociar precio unitario", "Pendiente", "Compras"],
        [3, "Contratista C", "ELE-001", "Luminarias LED", "Instalaciones", 48000, 0.22, 0.09, "Ficha técnica no validada", "Solicitar aclaración", "Pendiente", "Ingeniería"],
        [4, "Contratista B", "GEN-001", "Partida sin referencia directa", "Especial", 15000000, 0, 0.20, "Sin referencia", "Cotización externa", "Pendiente", "Costos"],
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
    # Professional base-budget version; kept separate from contractor comparison.
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
    detail = wb.create_sheet("Detalle Base")
    _setup_sheet(detail, "Detalle presupuesto base", "Cruce conceptual entre conceptos de ingeniería y matrices Construdata.", 9)
    _set_widths(detail,{"A":22,"B":44,"C":44,"D":12,"E":12,"F":16,"G":18,"H":18,"I":32})
    headers = ["Concepto base", "Descripción ingeniería", "Matriz Construdata", "Unidad", "Cantidad", "P.U. ref.", "Importe", "Estado", "Nota"]
    for c,h in enumerate(headers,1): detail.cell(5,c,h)
    _header_style(detail,5,1,9)
    rows = [
        ["Excavación", "Excavación manual", "Excavación manual material común", "m3", 120, 88000, "=E6*F6", "Con precio", "Usado para presupuesto base independiente"],
        ["Acero", "Acero de refuerzo", "Acero fy 4200", "kg", 1800, 6200, "=E7*F7", "Revisar", "Match medio por descripción"],
    ]
    for r,row in enumerate(rows,6):
        for c,v in enumerate(row,1): detail.cell(r,c,v)
    _body_style(detail,6,7,1,9)
    _apply_formats(detail, money_cols=[6,7], start_row=6, end_row=7)
    _add_table(detail,"A5:I7","BaseBudgetDetalleTable","TableStyleMedium4")
    _write_parametros(wb.create_sheet("Parámetros"))


def _write_comparison_report(wb):
    ws = wb.active
    ws.title = "Resumen Ejecutivo"
    _write_resumen_ejecutivo(ws, "comparison")
    _write_comparativa_professional(wb.create_sheet("Comparativa"))
    _write_detalle_apu_professional(wb.create_sheet("Detalle APU"))
    _write_partidas_criticas(wb.create_sheet("Partidas Críticas"))
    _write_insumos_criticos(wb.create_sheet("Insumos Críticos"))
    _write_validaciones(wb.create_sheet("Validaciones"))
    _write_parametros(wb.create_sheet("Parámetros"))
    _write_analisis_ia(wb.create_sheet("Análisis IA"))


def build_report(kind: str) -> Path:
    wb = Workbook()
    if kind == "comparison":
        _write_comparison_report(wb)
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
def report(kind: str):
    if kind not in {"base", "comparison"}:
        raise HTTPException(status_code=400, detail="kind debe ser base o comparison")
    path = build_report(kind)
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# Static frontend must be mounted last.
app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
