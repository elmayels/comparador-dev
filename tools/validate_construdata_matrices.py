"""
Valida que data/construdata_matrices.xlsx tenga estructura concepto -> matriz.

Uso:
  python tools/validate_construdata_matrices.py data/construdata_matrices.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from processor import load_construdata_matrices  # noqa: E402


def main():
    path = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "data" / "construdata_matrices.xlsx")
    if not path.exists():
        print(f"ERROR: no existe {path}")
        sys.exit(1)

    refs = load_construdata_matrices(str(path))
    conceptos = refs.get("__construdata_matrices__", {}) if isinstance(refs, dict) else {}
    materials = refs.get("__materials_index__", []) if isinstance(refs, dict) else []

    print("Archivo:", path)
    print("Tipo benchmark:", refs.get("__benchmark_kind__"))
    print("Hoja matrices:", refs.get("__matrix_sheet__"))
    print("Conceptos cargados:", len(conceptos))
    print("Insumos/materiales indexados:", len(materials))

    with_matrix = [(k, v) for k, v in conceptos.items() if v.get("matriz")]
    print("Conceptos con matriz:", len(with_matrix))

    if not with_matrix:
        print("ERROR: no se detecto ninguna matriz por concepto.")
        sys.exit(2)

    sample_key, sample = with_matrix[0]
    print("\nEjemplo concepto:")
    print("  Codigo:", sample_key)
    print("  Descripcion:", sample.get("desc"))
    print("  Unidad:", sample.get("unidad"))
    print("  PU:", sample.get("pu_mercado"))
    print("  Renglones:", len(sample.get("matriz") or []))
    print("  Primer insumo:", sample.get("matriz", [None])[0])


if __name__ == "__main__":
    main()
