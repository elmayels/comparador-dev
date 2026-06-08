"""Hotfix Step42: roles individuales de mano de obra no son cuadrillas.

Valida tres cosas criticas para matriz tipo Neodata/Construdata:
1. MO-SUP-SEG y MO-SUP-SEG. normalizan al mismo codigo.
2. SUPERVISOR DE SEGURIDAD es MO individual, no cuadrilla/brigada.
3. Una cuadrilla explicita sigue siendo cuadrilla para fallback conservador.

Ejecutar desde la raiz del proyecto:
    python tools/test_step42_mo_roles_no_cuadrilla.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import processor  # noqa: E402

assert processor._step08_norm_code('MO-SUP-SEG') == processor._step08_norm_code('MO-SUP-SEG.')
assert processor._step08_norm_code('MO-SUP-SEG') == 'MOSUPSEG'

supervisor = {
    'codigo': 'MO-SUP-SEG',
    'descripcion': 'SUPERVISOR DE SEGURIDAD',
    'unidad': 'JOR',
    'section': 'MANO DE OBRA',
    'tipo': 'MANO DE OBRA',
}
cuadrilla = {
    'codigo': '2A',
    'descripcion': 'CUADRILLA (2 AYUDANTE GENERAL)',
    'unidad': 'JOR',
    'section': 'MANO DE OBRA',
    'tipo': 'MANO DE OBRA',
}

assert processor._component_domain(supervisor) == 'mano_obra'
assert processor._step09_is_cuadrilla(supervisor) is False
assert processor._step40_is_individual_labor_role(supervisor) is True
assert processor._step40_is_cuadrilla_like(supervisor) is False, 'Supervisor no debe activar fallback de cuadrilla'

assert processor._step09_is_cuadrilla(cuadrilla) is True
assert processor._step40_is_cuadrilla_like(cuadrilla) is True, 'Cuadrilla explicita si debe activar fallback conservador'

# Simula el caso peligroso: si mercado supervisor > proveedor, NO debe conservar proveedor por regla de cuadrilla.
conceptos = {
    'SVC': {
        'granular_market_items': [
            {
                **supervisor,
                'precio_proveedor': 1000.0,
                'precio_mercado': 1500.0,
                'importe_proveedor': 100.0,
                'importe_mercado': 150.0,
                'match_status': 'match',
            }
        ]
    }
}
processor._step40_apply_cuadrilla_conservative_fallback(conceptos)
item = conceptos['SVC']['granular_market_items'][0]
assert item['precio_mercado'] == 1500.0
assert item['importe_mercado'] == 150.0
assert item.get('match_status') == 'match'

print('OK Step42: supervisor/roles MO individuales no activan fallback cuadrilla; codigos con punto matchean.')
