"""Pasaje directo fcmenu.mdb -> PostgreSQL, sin CSV intermedio.

Reusa TODO el motor ya probado de `load_csv_to_postgres.py` (config de
tablas, TRUNCATE automático, corte por años en --modo prueba, clientes/
artículos "fantasma" para huérfanos, manejo de las 3 FK) — lo único que
cambia es de dónde vienen los datos de origen: acá se leen fila por fila
directo del `.mdb` vía ODBC (`access_odbc.py`) en vez de un CSV ya
exportado.

Requiere el driver ODBC de Access instalado (ver access_odbc.py) y que
el esquema de Postgres ya esté aplicado (`alembic upgrade head` — lo
hace `migrar_completo.py` antes de llamar acá).

Uso:
    .venv\\Scripts\\python.exe migration\\etl\\load_mdb_to_postgres.py --mdb C:\\FCMenu\\fcmenu.mdb --modo completo
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

from sqlalchemy.orm import Session

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from migration.db import SessionLocal, engine  # noqa: E402
from migration.etl.access_odbc import conectar_mdb, convertir_valor, leer_tabla  # noqa: E402
from migration.etl.load_csv_to_postgres import (  # noqa: E402
    ANIOS_PRUEBA_DEFAULT,
    TABLAS,
    ConfigTabla,
    _fecha_corte,
    _generar_articulos_fantasma,
    _generar_clientes_fantasma,
    _recrear_fks,
    _soltar_fks,
    _truncar_tabla,
)


def _cargar_tabla_desde_mdb(db: Session, conn, config: ConfigTabla, *, fecha_corte=None) -> int:
    _truncar_tabla(db, config.modelo)

    if config.excluir_en_modo_prueba and fecha_corte is not None:
        print(f"  [VACÍA] {config.modelo.__tablename__} excluida en modo prueba (se recarga completa en el corte final)")
        return 0

    columnas = {c.name: c for c in config.modelo.__table__.columns if c.name != "id"}
    columna_fecha = config.columna_fecha if fecha_corte is not None else None
    nombre_tabla_access = config.modelo.__tablename__

    total = 0
    descartadas_por_fecha = 0
    lote = []
    LOTE_SIZE = 2000

    for fila in leer_tabla(conn, nombre_tabla_access):
        if columna_fecha is not None:
            valor_fecha = convertir_valor(fila.get(columna_fecha), columnas[columna_fecha])
            if valor_fecha is not None and valor_fecha < fecha_corte:
                descartadas_por_fecha += 1
                continue

        valores = {}
        for nombre_col, columna in columnas.items():
            valores[nombre_col] = convertir_valor(fila.get(nombre_col), columna)
        lote.append(valores)
        if len(lote) >= LOTE_SIZE:
            db.bulk_insert_mappings(config.modelo, lote)
            db.commit()
            total += len(lote)
            lote = []

    if lote:
        db.bulk_insert_mappings(config.modelo, lote)
        db.commit()
        total += len(lote)

    if descartadas_por_fecha:
        print(f"  ({descartadas_por_fecha} filas anteriores a {fecha_corte} descartadas — modo prueba)")

    return total


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--mdb", required=True, help=r"Ruta al fcmenu.mdb, ej. C:\FCMenu\fcmenu.mdb")
    parser.add_argument("--modo", choices=["completo", "prueba"], default="completo")
    parser.add_argument("--anios", type=int, default=ANIOS_PRUEBA_DEFAULT)
    parser.add_argument("--tablas", nargs="+", default=None)
    args = parser.parse_args()

    fecha_corte = _fecha_corte(args.anios) if args.modo == "prueba" else None

    tablas = TABLAS
    if args.tablas:
        pedidas = set(args.tablas)
        tablas = [c for c in TABLAS if c.modelo.__tablename__ in pedidas]

    print(f"Conectando a Access: {args.mdb}")
    conn = conectar_mdb(args.mdb)
    print(f"Conectando a Postgres: {engine.url}")
    print(f"Modo: {args.modo}" + (f" (últimos {args.anios} años, desde {fecha_corte})" if fecha_corte else ""))

    db = SessionLocal()
    resumen = []
    try:
        _soltar_fks(db)
        try:
            for config in tablas:
                print(f"Cargando {config.modelo.__tablename__} ...")
                n = _cargar_tabla_desde_mdb(db, conn, config, fecha_corte=fecha_corte)
                print(f"  {n} filas insertadas")
                resumen.append((config.modelo.__tablename__, n))

            n_clientes = _generar_clientes_fantasma(db)
            if n_clientes:
                print(f"\n{n_clientes} Cliente(s) fantasma repuesto(s)")
            n_articulos = _generar_articulos_fantasma(db)
            if n_articulos:
                print(f"{n_articulos} Artículo(s) fantasma repuesto(s)")
        finally:
            _recrear_fks(db)
    finally:
        db.close()
        conn.close()

    print("\nResumen:")
    for tabla, n in resumen:
        print(f"  {tabla}: {n}")
    print("Listo.")


if __name__ == "__main__":
    main()
