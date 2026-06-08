"""
Exporta la base maestra Construdata/Neodata como MATRICES POR CONCEPTO.

Uso recomendado:
  1) Adjunta el MDF en SQL Server / SQL Server Express.
  2) Ejecuta:
     python tools/export_construdata_matrices_sqlserver.py \
       --server localhost\\SQLEXPRESS \
       --database Construbase_Mayo_2026_LDB2019 \
       --out data/construdata_matrices.xlsx

Autenticacion SQL:
     python tools/export_construdata_matrices_sqlserver.py \
       --server localhost\\SQLEXPRESS \
       --database Construbase_Mayo_2026_LDB2019 \
       --user sa --password TU_PASSWORD \
       --out data/construdata_matrices.xlsx

Notas:
- Este archivo NO compara proveedores. Solo genera la referencia correcta.
- El comparador puede usar el Excel resultante como benchmark Neodata.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
import pyodbc


def connect(args):
    if args.user:
        cs = (
            f"DRIVER={{{args.driver}}};SERVER={args.server};DATABASE={args.database};"
            f"UID={args.user};PWD={args.password};TrustServerCertificate=yes;"
        )
    else:
        cs = (
            f"DRIVER={{{args.driver}}};SERVER={args.server};DATABASE={args.database};"
            "Trusted_Connection=yes;TrustServerCertificate=yes;"
        )
    return pyodbc.connect(cs, autocommit=True)


def read_sql(conn, sql, params=None):
    return pd.read_sql(sql, conn, params=params or [])


def table_exists(conn, table_name: str) -> bool:
    q = "SELECT 1 AS ok FROM sys.tables WHERE name = ?"
    return len(read_sql(conn, q, [table_name])) > 0


def export_schema(conn, out_dir: Path) -> pd.DataFrame:
    schema = read_sql(conn, """
        SELECT
            t.name AS tabla,
            c.column_id,
            c.name AS columna,
            ty.name AS tipo,
            c.max_length,
            c.precision,
            c.scale,
            c.is_nullable
        FROM sys.tables t
        INNER JOIN sys.columns c ON c.object_id = t.object_id
        INNER JOIN sys.types ty ON ty.user_type_id = c.user_type_id
        WHERE t.name LIKE 'Pu%'
        ORDER BY t.name, c.column_id
    """)
    schema.to_csv(out_dir / "schema_pu_tables.csv", index=False, encoding="utf-8-sig")
    return schema


def choose_budget(conn, requested_id: int | None) -> tuple[int, pd.DataFrame]:
    presupuestos = read_sql(conn, "SELECT * FROM PuPresupuestos ORDER BY IdPresupuesto")
    if requested_id is not None:
        return int(requested_id), presupuestos
    if len(presupuestos) == 1:
        return int(presupuestos.iloc[0]["IdPresupuesto"]), presupuestos
    print("Presupuestos encontrados:")
    print(presupuestos.head(50).to_string())
    print("\nEjecuta de nuevo con --id-presupuesto <ID>.")
    sys.exit(1)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--server", required=True)
    parser.add_argument("--database", required=True)
    parser.add_argument("--driver", default="ODBC Driver 17 for SQL Server")
    parser.add_argument("--user")
    parser.add_argument("--password")
    parser.add_argument("--id-presupuesto", type=int)
    parser.add_argument("--id-moneda", type=int, default=1)
    parser.add_argument("--out", default="data/construdata_matrices.xlsx")
    args = parser.parse_args()

    out_path = Path(args.out).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    conn = connect(args)
    schema = export_schema(conn, out_path.parent)

    required_tables = [
        "PuPresupuestos",
        "PuPresupuestosPartidas",
        "PuPresupuestosConceptos",
        "PuPresupuestosConceptosPrecios",
        "PuExpIns",
        "PuCatalogo",
        "PuUnidades",
        "PuMatrices",
        "PuExpInsCostos",
        "PuTipoInsumos",
    ]
    missing = [t for t in required_tables if not table_exists(conn, t)]
    if missing:
        print("Faltan tablas esperadas:", ", ".join(missing))
        print(f"Se genero {out_path.parent / 'schema_pu_tables.csv'} para ajustar nombres.")
        sys.exit(2)

    id_presupuesto, presupuestos = choose_budget(conn, args.id_presupuesto)
    id_moneda = args.id_moneda

    conceptos_sql = """
        SELECT
            PC.IdPresupuesto,
            PC.IdPresupuestoConcepto,
            PC.IdPresupuestoPartida,
            PC.Renglon AS ConceptoRenglon,
            PC.Cantidad,
            PC.CantidadTotal,
            PC.IdExpIns AS IdExpInsConcepto,
            C.IdCodigo AS IdCodigoConcepto,
            C.Codigo AS CodigoConcepto,
            C.Descripcion AS Concepto,
            C.DescripcionLarga AS ConceptoDescripcionLarga,
            U.Unidad AS ConceptoUnidad,
            PCP.Precio AS ConceptoPrecioUnitario
        FROM PuPresupuestosConceptos PC
        INNER JOIN PuExpIns E ON E.IdExpIns = PC.IdExpIns
        INNER JOIN PuCatalogo C ON C.IdCodigo = E.IdCodigo
        LEFT JOIN PuUnidades U ON U.IdUnidad = C.IdUnidad
        LEFT JOIN PuPresupuestosConceptosPrecios PCP
            ON PCP.IdPresupuestoConcepto = PC.IdPresupuestoConcepto
           AND PCP.IdMoneda = ?
        WHERE PC.IdPresupuesto = ?
        ORDER BY C.Codigo, PC.Renglon
    """
    conceptos = read_sql(conn, conceptos_sql, [id_moneda, id_presupuesto])

    matrices_sql = """
        SELECT
            CM.IdPresupuesto,
            M.IdMatriz,
            M.IdCodigoMatriz AS IdExpInsConcepto,
            CC.Codigo AS CodigoConcepto,
            CC.Descripcion AS Concepto,
            M.Renglon AS MatrizRenglon,
            M.Dividir,
            M.Volumen AS CantidadMatriz,
            M.Expresion,
            II.IdExpIns AS IdExpInsInsumo,
            CI.IdCodigo AS IdCodigoInsumo,
            CI.Codigo AS CodigoInsumo,
            CI.Descripcion AS Insumo,
            CI.DescripcionLarga AS InsumoDescripcionLarga,
            UI.Unidad AS UnidadInsumo,
            CI.IdTipo,
            TI.Descripcion AS TipoInsumo,
            CI.EsAgrupador,
            CI.EsPorcentaje,
            II.Expins,
            II.EsCostoHorario,
            II.InsumoIntegrado,
            EC.Costo AS CostoInsumo,
            CASE
                WHEN M.Dividir = 1 AND NULLIF(M.Volumen, 0) IS NOT NULL
                    THEN EC.Costo / NULLIF(M.Volumen, 0)
                ELSE EC.Costo * M.Volumen
            END AS ImporteCalculado
        FROM PuMatrices M
        INNER JOIN PuExpIns CM ON CM.IdExpIns = M.IdCodigoMatriz
        INNER JOIN PuCatalogo CC ON CC.IdCodigo = CM.IdCodigo
        INNER JOIN PuExpIns II ON II.IdExpIns = M.IdCodigoInsumo
        INNER JOIN PuCatalogo CI ON CI.IdCodigo = II.IdCodigo
        LEFT JOIN PuUnidades UI ON UI.IdUnidad = CI.IdUnidad
        LEFT JOIN PuTipoInsumos TI ON TI.IdTipo = CI.IdTipo
        LEFT JOIN PuExpInsCostos EC
            ON EC.IdExpIns = II.IdExpIns
           AND EC.IdMoneda = ?
        WHERE CM.IdPresupuesto = ?
        ORDER BY CC.Codigo, M.Renglon
    """
    matrices = read_sql(conn, matrices_sql, [id_moneda, id_presupuesto])

    partidas = read_sql(
        conn,
        "SELECT * FROM PuPresupuestosPartidas WHERE IdPresupuesto = ? ORDER BY IdPresupuestoPartida",
        [id_presupuesto],
    )

    catalogo_sql = """
        SELECT
            E.IdPresupuesto,
            E.IdExpIns,
            C.IdCodigo,
            C.Codigo,
            C.Descripcion,
            C.DescripcionLarga,
            U.Unidad,
            C.IdTipo,
            TI.Descripcion AS TipoInsumo,
            C.EsAgrupador,
            C.EsPorcentaje,
            E.Expins,
            E.EsCostoHorario,
            E.InsumoIntegrado,
            EC.Costo
        FROM PuExpIns E
        INNER JOIN PuCatalogo C ON C.IdCodigo = E.IdCodigo
        LEFT JOIN PuUnidades U ON U.IdUnidad = C.IdUnidad
        LEFT JOIN PuTipoInsumos TI ON TI.IdTipo = C.IdTipo
        LEFT JOIN PuExpInsCostos EC
            ON EC.IdExpIns = E.IdExpIns
           AND EC.IdMoneda = ?
        WHERE E.IdPresupuesto = ?
        ORDER BY C.Codigo
    """
    catalogo = read_sql(conn, catalogo_sql, [id_moneda, id_presupuesto])

    base_comparacion = conceptos.merge(
        matrices,
        on=["IdPresupuesto", "IdExpInsConcepto", "CodigoConcepto", "Concepto"],
        how="left",
    )

    resumen = matrices.groupby(["CodigoConcepto", "Concepto"], dropna=False).agg(
        RenglonesMatriz=("CodigoInsumo", "count"),
        TotalMatriz=("ImporteCalculado", "sum"),
    ).reset_index()
    resumen = conceptos[["CodigoConcepto", "Concepto", "ConceptoUnidad", "ConceptoPrecioUnitario"]].merge(
        resumen,
        on=["CodigoConcepto", "Concepto"],
        how="left",
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        presupuestos.to_excel(writer, "presupuestos", index=False)
        partidas.to_excel(writer, "partidas", index=False)
        conceptos.to_excel(writer, "conceptos", index=False)
        matrices.to_excel(writer, "matrices_detalle", index=False)
        resumen.to_excel(writer, "resumen_matrices", index=False)
        catalogo.to_excel(writer, "catalogo_insumos", index=False)
        base_comparacion.to_excel(writer, "base_comparacion", index=False)
        schema.to_excel(writer, "schema", index=False)

    print(f"OK: exportado {out_path}")
    print(f"Presupuesto: {id_presupuesto}; moneda: {id_moneda}")
    print(f"Conceptos: {len(conceptos):,}")
    print(f"Renglones de matrices: {len(matrices):,}")
    print(f"Conceptos con matriz: {matrices['CodigoConcepto'].nunique():,}")


if __name__ == "__main__":
    main()
