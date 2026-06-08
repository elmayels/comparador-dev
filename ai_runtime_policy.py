"""
Politica central de IA para Quantia.

Objetivo:
- Evitar consumo accidental de Claude/Voyage durante la generacion normal del Excel.
- Separar busqueda deterministica/rapida de tareas IA costosas.

Variables:
- QUANTIA_ENABLE_AI: bandera general. Default 0.
- QUANTIA_AI_RUNTIME_MODE:
    off              -> no se llama ninguna API externa en comparativas.
    index_only       -> reservado para usar indices locales/precalculados sin llamar APIs.
    runtime_review   -> permite llamadas runtime granulares, solo para pruebas controladas.
    controlled_review -> permite UNA revision Claude controlada al final del reporte; no llamadas por insumo/concepto.
    admin_reindex    -> permite llamadas Voyage en endpoints admin de reindexado.
- ENABLE_AI_MATERIAL_MATCHING: compatibilidad legacy; solo se respeta si runtime_mode=runtime_review.
"""
from __future__ import annotations

import os
from typing import Dict

_TRUE = {"1", "true", "yes", "on", "si", "sí"}


def env_true(name: str, default: str = "0") -> bool:
    return str(os.getenv(name, default)).strip().lower() in _TRUE


def ai_mode() -> str:
    return str(os.getenv("QUANTIA_AI_RUNTIME_MODE", "off")).strip().lower() or "off"


def ai_globally_enabled() -> bool:
    return env_true("QUANTIA_ENABLE_AI", "0")


def runtime_ai_allowed(scope: str = "runtime") -> bool:
    """True solo cuando se habilitan llamadas externas dentro de una corrida normal."""
    if not ai_globally_enabled():
        return False
    mode = ai_mode()
    if scope == "admin_reindex":
        return mode in {"admin_reindex", "runtime_review"}
    if scope == "material_matching":
        # Material matching con IA por renglón queda deshabilitado salvo modo runtime_review explícito.
        return mode == "runtime_review" and env_true("ENABLE_AI_MATERIAL_MATCHING", "0")
    if scope == "concept_matching":
        # PASO 33: permitir IA conceptual controlada.
        # No se usa para insumos/materiales; solo para concepto -> matriz CD.
        return mode in {"controlled_review", "runtime_review"} and env_true("ENABLE_AI_CONCEPT_MATCHING", "0")
    if scope == "expert_text":
        # controlled_review permite una revision experta al final del reporte.
        return mode in {"controlled_review", "runtime_review"} and env_true("ENABLE_AI_EXPERT_TEXT", "1")
    if scope == "expert_review":
        return mode in {"controlled_review", "runtime_review"} and env_true("ENABLE_AI_EXPERT_TEXT", "1")
    return mode == "runtime_review"


def external_ai_block_reason(scope: str = "runtime") -> str:
    return (
        f"IA runtime desactivada para scope={scope}. "
        f"QUANTIA_ENABLE_AI={os.getenv('QUANTIA_ENABLE_AI','0')}, "
        f"QUANTIA_AI_RUNTIME_MODE={os.getenv('QUANTIA_AI_RUNTIME_MODE','off')}."
    )


def ai_policy_status() -> Dict[str, object]:
    return {
        "quantia_enable_ai": "on" if ai_globally_enabled() else "off",
        "runtime_mode": ai_mode(),
        "material_runtime_allowed": runtime_ai_allowed("material_matching"),
        "concept_runtime_allowed": runtime_ai_allowed("concept_matching"),
        "expert_text_allowed": runtime_ai_allowed("expert_text"),
        "admin_reindex_allowed": runtime_ai_allowed("admin_reindex"),
    }
