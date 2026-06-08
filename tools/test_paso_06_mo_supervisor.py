"""Prueba rapida del fix Paso 06.

Ejecutar desde la raiz del proyecto:
    python tools/test_paso_06_mo_supervisor.py

Debe confirmar que SUPERVISOR DE SEGURIDAD (mano de obra) no se homologa
contra EQUIPO DE SEGURIDAD (equipo/herramienta), aunque compartan el token
"seguridad".
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import processor  # noqa: E402

provider = {
    "tipo_kind": "mano_obra",
    "codigo": "MO-SUP-SEG",
    "descripcion": "SUPERVISOR DE SEGURIDAD",
    "unidad": "JOR",
}

cd_equipo_seguridad = {
    "tipo_kind": "equipo",
    "codigo": "%MO5",
    "descripcion": "EQUIPO DE SEGURIDAD",
    "unidad": "%",
    "formula_kind": "porcentaje_mo",
}

cd_supervisor = {
    "tipo_kind": "mano_obra",
    "codigo": "SUPSEG",
    "descripcion": "SUPERVISOR DE SEGURIDAD",
    "unidad": "JOR",
}

compatible, reason = processor._component_domains_compatible(provider, cd_equipo_seguridad)
assert compatible is False, "ERROR: supervisor de seguridad no debe ser compatible con equipo de seguridad"
assert processor._component_match_score(provider, cd_equipo_seguridad) == 0.0
idx, cd, score = processor._match_matrix_component(provider, [cd_equipo_seguridad], min_score=0.1)
assert idx is None and cd is None, "ERROR: se genero un match falso contra equipo de seguridad"

idx, cd, score = processor._match_matrix_component(provider, [cd_supervisor], min_score=0.1)
assert idx == 0 and cd is cd_supervisor and score >= 0.80, "ERROR: supervisor si debe empatar con mano de obra equivalente"

print("OK Paso 06: MO-SUP-SEG no empata con EQUIPO DE SEGURIDAD y si empata con MO equivalente.")
print("Motivo rechazo esperado:", reason)
