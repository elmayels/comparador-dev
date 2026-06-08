"""Smoke check rapido para Quantia Comparador APU v0.2.

Uso:
  DATA_DIR=$PWD/data python tools/smoke_mvp_v02.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(root))
    os.environ.setdefault("DATA_DIR", str(root / "data"))
    from fastapi.testclient import TestClient
    from backend.main import app

    client = TestClient(app)
    checks = ["/health", "/benchmark/status"]
    for path in checks:
        res = client.get(path)
        print(f"{path}: {res.status_code}")
        if res.status_code >= 400:
            print(res.text[:500])
            return 1
    print("OK: app importada, rutas basicas disponibles, DATA_DIR configurado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
