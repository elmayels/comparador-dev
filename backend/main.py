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


def _style_sheet(ws, title: str):
    ws.freeze_panes = "A4"
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="0B1020")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    ws["A2"] = "Versión 0 canónica - datos mockeados, estructura preparada para motor real"
    ws["A2"].font = Font(italic=True, color="667085")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)


def _write_table(ws, start_row: int, headers: list[str], rows: list[list]):
    header_fill = PatternFill("solid", fgColor="17213D")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D0D5DD")
    for col, h in enumerate(headers, start=1):
        c = ws.cell(start_row, col, h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")
        c.border = Border(bottom=thin)
    for r_i, row in enumerate(rows, start=start_row + 1):
        for c_i, val in enumerate(row, start=1):
            cell = ws.cell(r_i, c_i, val)
            cell.border = Border(bottom=Side(style="hair", color="EAECF0"))
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
    for idx, _ in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = 18



def _money(cell):
    cell.number_format = '$#,##0.00'


def _pct(cell):
    cell.number_format = '0.00%'


def _dark_header(ws, row: int, col_start: int, col_end: int):
    fill = PatternFill("solid", fgColor="000000")
    font = Font(bold=True, color="FF0000")
    border = Border(bottom=Side(style="thin", color="00B050"))
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _dark_body_row(ws, row: int, col_start: int, col_end: int, section: bool = False):
    fill = PatternFill("solid", fgColor="000000")
    border = Border(top=Side(style="thin", color="008000"), bottom=Side(style="thin", color="008000"))
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.font = Font(color="0000FF", bold=section)
        cell.border = border
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def _configure_legacy_visual_sheet(ws):
    ws.sheet_view.showGridLines = False
    widths = {
        'A': 14, 'B': 70, 'C': 12, 'D': 14, 'E': 10, 'F': 14, 'G': 14,
        'H': 12, 'I': 4, 'J': 16, 'K': 12, 'L': 16, 'M': 16
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _write_comparativa_like_legacy(ws):
    ws.title = "Comparativa"
    _configure_legacy_visual_sheet(ws)
    ws.freeze_panes = "A3"
    ws["A1"] = "Servicios / Cotización"
    ws["E1"] = "Matrices Pu FLEXIBILIDAD TQ 41 Y 42"
    for rng in ["A1:D1", "E1:J1"]:
        ws.merge_cells(rng)
    for cell in [ws["A1"], ws["E1"]]:
        cell.fill = PatternFill("solid", fgColor="000000")
        cell.font = Font(bold=True, color="0000FF")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(bottom=Side(style="thin", color="008000"))

    headers = ["Partida", "Descripción", "Unidad", "Cantidad", "P.U.", "Importe", "% Part.", "% ajuste", "Mercado P.U.", "Mercado Importe"]
    for idx, header in enumerate(headers, 1):
        ws.cell(2, idx, header)
    _dark_header(ws, 2, 1, 10)

    rows = [
        ["FLEX41.11", "Instalación de bomba centrifuga Fristam, modelo FPR 3531-155 con motor de 2HP para transferencia de producto a placa. Incluye materiales, mano de obra, herramientas, equipos y todo lo necesario para su correcta instalación.", "PZA", 1, 31524.55, 31524.55, 0.028268, 4.236288, 6020.40, 6020.40],
        ["FLEX41.12", "Suministro e instalación de guarda protectora para bomba centrifuga en acero inoxidable 304. Incluye materiales, mano de obra, herramientas, equipos y limpieza final.", "PZA", 1, 24937.16, 24937.16, 0.022361, 3.142110, 6020.40, 6020.40],
        ["FLEX41.29", "Suministro e instalación de tubería sanitaria de acero inoxidable 316L 3 pulgadas. Incluye soportes, cortes, pulidos y limpieza del área de trabajo.", "M", 41, 3496.73, 143365.93, 0.128557, -0.419186, 6020.40, 246836.40],
        ["FLEX41.40", "Suministro e instalación de tubería sanitaria de acero inoxidable 316L 2-1/2 pulgadas. Incluye materiales, mano de obra y equipos.", "M", 38, 3496.73, 132875.74, 0.11915, -0.419186, 6020.40, 228775.20],
        ["FLEX41.62", "Servicio de control de aseguramiento de calidad a través de ensayo de boroscopia en tuberías de acero inoxidable.", "SERVICIO", 1, 75226.36, 75226.36, 0.067456, 11.495093, 6020.40, 6020.40],
    ]
    for r, row in enumerate(rows, 3):
        for c, val in enumerate(row, 1):
            ws.cell(r, c, val)
        _dark_body_row(ws, r, 1, 10)
        for col in [5, 6, 9, 10]:
            _money(ws.cell(r, col))
        for col in [7, 8]:
            _pct(ws.cell(r, col))
        ws.row_dimensions[r].height = 40

    total_row = 3 + len(rows)
    ws.cell(total_row, 1, "TOTAL")
    ws.cell(total_row, 5, "=AVERAGE(E3:E7)")
    ws.cell(total_row, 6, "=SUM(F3:F7)")
    ws.cell(total_row, 9, "=AVERAGE(I3:I7)")
    ws.cell(total_row, 10, "=SUM(J3:J7)")
    _dark_body_row(ws, total_row, 1, 10, section=True)
    for col in [5, 6, 9, 10]:
        _money(ws.cell(total_row, col))

    summary_start = total_row + 2
    bullets = [
        ["Resumen individual por contratista"],
        ["Matrices Pu FLEXIBILIDAD TQ 41 Y 42"],
        ["• Las partidas sombreadas en azul concentran el mayor impacto económico del contratista."],
        ["• El total cotizado se compara contra mercado calculado usando los precios base de materiales, mano de obra y maquinaria desde data."],
        ["• Las partidas con mayor desviación deben revisarse contra alcance, exclusiones y suficiencia técnica."],
        ["• Esta hoja replica la lógica histórica del tab Comparativa: partida, precio ofertado, importe, participación, ajuste y mercado."],
    ]
    for i, row in enumerate(bullets, summary_start):
        ws.cell(i, 1, row[0])
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
        _dark_body_row(ws, i, 1, 10, section=i in [summary_start, summary_start + 1])


def _write_detalle_like_legacy(ws):
    ws.title = "Detalle"
    _configure_legacy_visual_sheet(ws)
    ws.freeze_panes = "A2"
    headers = ["Código", "Concepto", "Unidad", "P. Unitario", "Op.", "Cantidad", "Importe", "%", "", "Mercado P. Unitario", "Mercado Op.", "Mercado Cantidad", "Mercado Importe"]
    for idx, header in enumerate(headers, 1):
        ws.cell(1, idx, header)
    _dark_header(ws, 1, 1, 13)

    rows = [
        ["FLEX41.11", "Instalación de bomba centrifuga Fristam, modelo FPR 3531-155 con motor de 2HP para transferencia de producto a placa. Incluye materiales, mano de obra, herramientas y equipos.", "PZA", 31524.55, None, 1, 31524.55, 0.028268, None, 6020.40, "*", 1, 6020.40, "concept"],
        [None, "MATERIALES", None, None, None, None, None, None, None, None, None, None, None, "section"],
        ["consu-mec", "Consumibles para corte y soldadura", "%", 22453.38, "*", 0.08, 1796.27, 0.001611, None, 22453.38, "*", 0.08, 1796.27, "item"],
        [None, "SUBTOTAL MATERIALES", None, 1796.27, None, None, 1796.27, 0.001611, None, None, None, None, None, "subtotal"],
        [None, "MANO DE OBRA", None, None, None, None, None, None, None, None, None, None, None, "section"],
        ["SUP-O", "Supervisor Obra", "JOR", 2643.19, "*", 0.10, 264.32, 0.000237, None, 1757.33, "*", 0.10, 175.73, "item"],
        ["SUP-SEG", "Supervisor de Seguridad", "JOR", 2287.73, "*", 0.333333, 762.58, 0.000684, None, 1757.33, "*", 0.33, 579.92, "item"],
        ["OF-SOL", "Oficial Soldador/argonero", "JOR", 2120.45, "*", 1, 2120.45, 0.001901, None, 1326.14, "*", 1, 1326.14, "item"],
        ["OF-TUB", "Oficial Tubero", "JOR", 2120.45, "*", 1, 2120.45, 0.001901, None, 1221.95, "*", 1, 1221.95, "item"],
        ["AYU-GRAL", "Ayudante General", "JOR", 1074.96, "*", 1, 1074.96, 0.000964, None, 777.88, "*", 1, 777.88, "item"],
        [None, "SUBTOTAL MANO DE OBRA", None, 6342.76, None, None, 6342.76, 0.005688, None, 4081.62, None, None, 4081.62, "subtotal"],
        [None, "EQUIPO Y HERRAMIENTA", None, None, None, None, None, None, None, None, None, None, None, "section"],
        ["%HERR", "Herramienta Menor", "%", 6342.76, "%", 0.09, 570.85, 0.000512, None, 4081.62, "%", 0.09, 367.35, "item"],
        ["%EPP", "Equipo de proteccion personal Basico", "%", 6342.76, "%", 0.09, 570.85, 0.000512, None, 4081.62, "%", 0.09, 367.35, "item"],
        [None, "SUBTOTAL EQUIPO Y HERRAMIENTA", None, 1141.70, None, None, 1141.70, 0.001024, None, 734.70, None, None, 734.70, "subtotal"],
        [None, "SECCION FINANCIERA", None, None, None, None, None, None, None, None, None, None, None, "section"],
        [None, "COSTO DIRECTO", None, 9280.73, None, None, 9280.73, 0.021745, None, 4816.32, None, None, 4816.32, "subtotal"],
        [None, "COSTO INDIRECTO", None, 7274.90, "%", 0.25, 7274.90, 0.25, None, 4816.32, "%", 0.25, 1204.08, "subtotal"],
        [None, "TOTAL COSTO UNITARIO", None, 31524.55, None, None, 31524.55, 0.028268, None, 6020.40, None, None, 6020.40, "total"],
        [None, "TOTAL POR SERVICIO", None, 31524.55, None, 1, 31524.55, 0.028268, None, 6020.40, None, 1, 6020.40, "total"],
        [None, None, None, None, None, None, None, None, None, None, None, None, None, "blank"],
        ["FLEX41.12", "Suministro e instalación de guarda protectora para bomba centrifuga en acero inoxidable 304. Incluye materiales, mano de obra, herramientas, equipos y limpieza final.", "PZA", 24937.16, None, 1, 24937.16, 0.022361, None, 6020.40, "*", 1, 6020.40, "concept"],
        [None, "MATERIALES", None, None, None, None, None, None, None, None, None, None, None, "section"],
        ["flex4.16", "Lamina lisa en acero inoxidable calibre 20", "M2", 3016.00, "*", 1, 3016.00, 0.002704, None, 3016.00, "*", 1, 3016.00, "item"],
        ["consu-mec", "Consumibles para corte y soldadura", "%", 3016.00, "%", 0.08, 241.28, 0.001074, None, 3016.00, "%", 0.08, 241.28, "item"],
        [None, "SUBTOTAL MATERIALES", None, 3257.28, None, None, 3257.28, 0.002921, None, None, None, None, None, "subtotal"],
    ]
    for r, row in enumerate(rows, 2):
        kind = row[-1]
        for c, val in enumerate(row[:-1], 1):
            ws.cell(r, c, val)
        _dark_body_row(ws, r, 1, 13, section=kind in {"section", "subtotal", "total"})
        if kind == "blank":
            ws.row_dimensions[r].height = 8
        elif kind == "concept":
            ws.row_dimensions[r].height = 48
        for col in [4, 7, 10, 13]:
            if isinstance(ws.cell(r, col).value, (int, float)):
                _money(ws.cell(r, col))
        for col in [8]:
            if isinstance(ws.cell(r, col).value, (int, float)):
                _pct(ws.cell(r, col))


def build_report(kind: str) -> Path:
    wb = Workbook()
    ws = wb.active
    if kind == "comparison":
        _write_comparativa_like_legacy(ws)
        detail = wb.create_sheet("Detalle")
        _write_detalle_like_legacy(detail)
        ai = wb.create_sheet("Analisis IA")
        ai["A1"] = "Analisis IA"
        ai["A1"].font = Font(bold=True, size=14)
        ai_rows = [
            "1. Veredicto general",
            "La propuesta se evalúa contra mercado granular usando referencias de data para materiales, mano de obra y maquinaria.",
            "2. Principales riesgos detectados",
            "Validar partidas bajo mercado para descartar omisiones de alcance y revisar porcentajes financieros/indirectos.",
            "3. Limitaciones",
            "V0 usa datos mockeados; el motor real deberá leer la matriz/APU del contratista para poblar Detalle.",
        ]
        for i, text in enumerate(ai_rows, 3):
            ai.cell(i, 1, text)
            ai.column_dimensions['A'].width = 120
            ai.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="top")
    else:
        ws.title = "Comparativa"
        _style_sheet(ws, "Comparativa")
        headers = ["Código", "Concepto ingeniería", "Unidad", "Cantidad", "Concepto Construdata", "PU referencia", "Importe base", "Confianza", "Estado"]
        rows = [
            ["001", "Excavación manual", "m3", 120, "Excavación manual material común", 88000, 10560000, "Alta", "Con precio"],
            ["002", "Concreto f'c 250", "m3", 40, "Concreto 250 kg/cm2", 420000, 16800000, "Alta", "Con precio"],
            ["003", "Acero de refuerzo", "kg", 1800, "Acero fy 4200", 6200, 11160000, "Media", "Revisar"],
            ["004", "Partida especial", "gl", 1, "", 0, 0, "Baja", "Sin referencia"],
        ]
        _write_table(ws, 4, headers, rows)
        detail = wb.create_sheet("Detalle")
        _style_sheet(detail, "Detalle")
        headers = ["Concepto base", "Matriz Construdata", "Unidad", "Cantidad", "PU ref.", "Importe", "Operador", "Estado", "Nota"]
        rows = [
            ["Excavación manual", "Excavación manual material común", "m3", 120, 88000, 10560000, "*", "Con precio", "Usado para presupuesto base independiente"],
            ["Acero de refuerzo", "Acero fy 4200", "kg", 1800, 6200, 11160000, "*", "Revisar", "Match medio por descripción"],
        ]
        _write_table(detail, 4, headers, rows)

    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
    filename = f"apu_v0_{kind}_{int(time.time())}.xlsx"
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
