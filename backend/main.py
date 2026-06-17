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


def build_report(kind: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativa"
    _style_sheet(ws, "Comparativa")
    if kind == "base":
        headers = ["Código", "Concepto ingeniería", "Unidad", "Cantidad", "Concepto Construdata", "PU referencia", "Importe base", "Confianza", "Estado"]
        rows = [
            ["001", "Excavación manual", "m3", 120, "Excavación manual material común", 88000, 10560000, "Alta", "Con precio"],
            ["002", "Concreto f'c 250", "m3", 40, "Concreto 250 kg/cm2", 420000, 16800000, "Alta", "Con precio"],
            ["003", "Acero de refuerzo", "kg", 1800, "Acero fy 4200", 6200, 11160000, "Media", "Revisar"],
            ["004", "Partida especial", "gl", 1, "", 0, 0, "Baja", "Sin referencia"],
        ]
    else:
        headers = ["Contratista", "Monto total", "Ranking", "Dif vs menor", "Desv. promedio", "Partidas críticas", "Semáforo", "Riesgo", "Observación"]
        rows = [
            ["Contratista B", 1180000, 1, 0, -7.8, 8, "Verde", "Bajo", "Oferta más competitiva"],
            ["Contratista A", 1250000, 2, 5.9, 2.4, 18, "Amarillo", "Medio", "Revisar acero y acabados"],
            ["Contratista C", 1410000, 3, 19.5, 14.3, 26, "Rojo", "Alto", "Sobrecostos en concreto"],
        ]
    _write_table(ws, 4, headers, rows)

    detail = wb.create_sheet("Detalle")
    _style_sheet(detail, "Detalle")
    if kind == "base":
        headers = ["Concepto base", "Matriz Construdata", "Unidad", "Cantidad", "PU ref.", "Importe", "Operador", "Estado", "Nota"]
        rows = [
            ["Excavación manual", "Excavación manual material común", "m3", 120, 88000, 10560000, "*", "Con precio", "Usado para presupuesto base independiente"],
            ["Acero de refuerzo", "Acero fy 4200", "kg", 1800, 6200, 11160000, "*", "Revisar", "Match medio por descripción"],
        ]
    else:
        headers = ["Contratista", "Concepto", "Tipo insumo", "Insumo", "Unidad", "Cantidad", "Precio contratista", "Precio base data", "Dif %", "Alerta"]
        rows = [
            ["Contratista A", "Concreto f'c 250", "Material", "Cemento", "kg", 320, 580, 560, 3.57, "OK"],
            ["Contratista A", "Concreto f'c 250", "MO", "Oficial", "hr", 1.2, 18000, 17500, 2.85, "OK"],
            ["Contratista C", "Concreto f'c 250", "Maquinaria", "Mezcladora", "hr", 0.4, 42000, 34000, 23.52, "Revisar"],
            ["Contratista C", "Acabados", "%", "Indirectos", "%", 1, 18, 15, 20, "Porcentaje superior"],
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
