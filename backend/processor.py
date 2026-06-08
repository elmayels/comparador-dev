"""Compatibility shim for deployments that import backend.processor.

The real implementation lives at the project root: /app/processor.py.
Railway/startup variants sometimes execute backend/main.py directly, causing
Python to resolve `processor` from /app/backend first. This shim forces loading
and re-exporting the root processor module so imports remain consistent.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
_REAL_PROCESSOR = _ROOT / "processor.py"

if not _REAL_PROCESSOR.exists():
    raise ImportError(f"Root processor.py not found at {_REAL_PROCESSOR}")

spec = importlib.util.spec_from_file_location("_quantia_root_processor", _REAL_PROCESSOR)
if spec is None or spec.loader is None:
    raise ImportError(f"Cannot load root processor.py from {_REAL_PROCESSOR}")

module = importlib.util.module_from_spec(spec)
sys.modules.setdefault("_quantia_root_processor", module)
spec.loader.exec_module(module)

for name in dir(module):
    if not name.startswith("__"):
        globals()[name] = getattr(module, name)
