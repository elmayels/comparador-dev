from pathlib import Path
import tempfile
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from processor import build_comparativo

ROOT = Path(__file__).resolve().parents[1]
provider = ROOT.parent / '334.00-26-NESTLE-EGG REPORTE DE MATRICES 20260430.xlsx'
if not provider.exists():
    print('Coloca el archivo proveedor junto al proyecto para probar:', provider.name)
    raise SystemExit(0)

out = Path(tempfile.gettempdir()) / 'test_paso_07.xlsx'
build_comparativo([str(provider)], ['Proveedor prueba'], str(out), {'cliente':'TEST','proyecto':'PASO 07'}, nacional_path=str(ROOT/'data'/'construdata_matrices.xlsx'))
print(out)
