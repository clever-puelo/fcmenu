"""Orquestador de un pasaje completo Access -> PostgreSQL, con
regeneración limpia de la base (DROP+CREATE, no sólo TRUNCATE de tablas)
y verificación automática al final.

Camino por default: lectura DIRECTA del `.mdb` vía ODBC (`--mdb`,
`load_mdb_to_postgres.py`) — sin CSV intermedio, sin PowerShell aparte.
Requiere el driver "Microsoft Access Database Engine 2016
Redistributable" (64-bit) instalado en esta PC — ver `access_odbc.py`.

Camino alternativo (`--fuente csv`): si por algún motivo el driver ODBC
de 64 bits no se puede instalar en esta PC (ej. ya hay Office de 32 bits
instalado ahí), usa el camino viejo — CSV ya generado aparte con
`export_access_to_csv.ps1` en 32-bit, copiado a `migration/etl/data/`.

Variables de entorno:
  DATABASE_URL   conexión de la app (rol fcmenu_app) contra la base a
                 (re)crear, ej.:
                 postgresql+psycopg2://fcmenu_app:XXX@localhost:5432/fcmenu_prod
  PG_ADMIN_URL   conexión de superusuario, para poder DROP/CREATE
                 DATABASE y el rol si hace falta, ej.:
                 postgresql+psycopg2://postgres:YYY@localhost:5432/postgres

Uso:
    .venv\\Scripts\\python.exe migration\\etl\\migrar_completo.py --mdb C:\\FCMenu\\fcmenu.mdb --modo completo
    .venv\\Scripts\\python.exe migration\\etl\\migrar_completo.py --mdb C:\\FCMenu\\fcmenu.mdb --modo prueba --anios 10
    .venv\\Scripts\\python.exe migration\\etl\\migrar_completo.py --fuente csv --modo completo  # camino viejo, con CSV ya generado
"""

from __future__ import annotations

import argparse
import os
import sys
from decimal import Decimal
from pathlib import Path
from typing import Optional

import psycopg2
from sqlalchemy import text
from sqlalchemy.engine import make_url

# Empaquetado como .exe único (PyInstaller `--onefile`): congelado,
# `__file__` apunta adentro de la carpeta temporal de extracción
# (`sys._MEIPASS`), no al repo real — ahí es donde hay que buscar
# `alembic.ini`/`migration/alembic/` (bundleados como datos, ver
# `build_exe.py`). Corriendo desde código fuente (no congelado), sigue
# siendo el repo de siempre.
if getattr(sys, "frozen", False):
    REPO_ROOT = Path(sys._MEIPASS)  # type: ignore[attr-defined]
else:
    REPO_ROOT = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(REPO_ROOT))


def _regenerar_base(app_url: str, admin_url: str) -> None:
    """DROP + CREATE del rol y la base de aplicación — limpieza total,
    como si Postgres se acabara de instalar. Corta antes cualquier
    conexión activa a la base (pg_terminate_backend) para que el DROP
    no falle por "database is being accessed by other users"."""
    app = make_url(app_url)
    admin = make_url(admin_url)

    conn = psycopg2.connect(
        host=admin.host or "localhost", port=admin.port or 5432,
        user=admin.username, password=admin.password,
        dbname=admin.database or "postgres",
    )
    conn.autocommit = True  # DROP/CREATE DATABASE no puede ir dentro de una transacción
    cur = conn.cursor()

    print(f"Rol {app.username!r} ...")
    cur.execute("SELECT 1 FROM pg_roles WHERE rolname = %s", (app.username,))
    if cur.fetchone():
        cur.execute(f'ALTER ROLE "{app.username}" WITH PASSWORD %s', (app.password,))
    else:
        cur.execute(f'CREATE ROLE "{app.username}" LOGIN PASSWORD %s', (app.password,))

    print(f"Eliminando base {app.database!r} si existe ...")
    cur.execute(
        "SELECT pg_terminate_backend(pid) FROM pg_stat_activity "
        "WHERE datname = %s AND pid <> pg_backend_pid()",
        (app.database,),
    )
    cur.execute(f'DROP DATABASE IF EXISTS "{app.database}"')

    print(f"Creando base {app.database!r} (owner {app.username!r}) ...")
    cur.execute(f'CREATE DATABASE "{app.database}" OWNER "{app.username}"')

    cur.close()
    conn.close()


def _aplicar_esquema() -> None:
    from alembic import command
    from alembic.config import Config

    print("Aplicando esquema (alembic upgrade head) ...")
    cfg = Config(str(REPO_ROOT / "alembic.ini"))
    command.upgrade(cfg, "head")


def _cargar_datos(fuente: str, mdb: Optional[str], modo: str, anios: int) -> None:
    # Llamada directa a la función, NO subproceso (`sys.executable` +
    # ruta a otro .py) — un .exe congelado no tiene un intérprete ni
    # archivos .py sueltos para invocar así. También evita el costo de
    # levantar un segundo proceso Python + reconectar a Postgres aparte.
    print(f"Cargando datos (fuente={fuente}, modo={modo}) ...")
    if fuente == "mdb":
        from migration.etl.load_mdb_to_postgres import cargar

        cargar(mdb, modo, anios)
    else:
        from migration.etl.load_csv_to_postgres import cargar

        cargar(modo, anios)


def _verificar(fuente: str, mdb: Optional[str], modo: str) -> None:
    from migration.db import SessionLocal
    from migration.etl.load_csv_to_postgres import TABLAS
    from migration.repository import RepositoryFactory

    conn_mdb = None
    if fuente == "mdb":
        from migration.etl.access_odbc import conectar_mdb
        conn_mdb = conectar_mdb(mdb)

    db = SessionLocal()
    try:
        origen = "Access (.mdb)" if fuente == "mdb" else "CSV"
        print(f"\n== Conteo de filas por tabla (Postgres vs. {origen}) ==")
        for config in TABLAS:
            n_db = db.execute(text(f'SELECT COUNT(*) FROM "{config.modelo.__tablename__}"')).scalar()
            if fuente == "mdb":
                cur = conn_mdb.cursor()
                cur.execute(f"SELECT COUNT(*) FROM [{config.modelo.__tablename__}]")
                n_origen = cur.fetchone()[0]
                cur.close()
            else:
                from migration.etl.load_csv_to_postgres import DATA_DIR
                ruta = DATA_DIR / config.archivo
                n_origen = (sum(1 for _ in ruta.open(encoding="utf-8-sig")) - 1) if ruta.exists() else None
            marca = ""
            if modo == "completo" and n_origen is not None and n_origen != n_db and not config.excluir_en_modo_prueba:
                marca = f"  <-- DIFIERE DE {origen.upper()}"
            print(f"  {config.modelo.__tablename__:12s} Postgres={n_db:>8}  {origen}={n_origen if n_origen is not None else '—':>8}{marca}")

        if fuente == "mdb":
            conn_mdb.close()

        if modo != "completo":
            print("\n(Cruce de saldos omitido en --modo prueba: el corte por años hace que "
                  "Cliente.DEUDA y el saldo recalculado NO deban coincidir a propósito.)")
            return

        print("\n== Cliente.DEUDA (cacheado, viene de Access tal cual) vs. saldo recalculado sumando Ctasctes ==")
        repos = RepositoryFactory(db)
        clientes = repos.cliente().todos_ordenados_por_nombre()
        diferencias = 0
        fantasma = 0
        for c in clientes:
            calculado = repos.ctascte().deuda_cliente(c.CODIGO)["deuda"]
            cacheado = c.DEUDA if c.DEUDA is not None else Decimal("0")
            if abs(calculado - cacheado) > Decimal("0.01"):
                if c.NOMB and c.NOMB.startswith("[BAJA]"):
                    fantasma += 1  # cliente recuperado de un huérfano — no tenía DEUDA en Access, es esperable
                    continue
                diferencias += 1
                print(f"  Cliente {c.CODIGO} ({c.NOMB}): Access DEUDA={cacheado}  Recalculado={calculado}  <-- DIFIERE")
        if fantasma:
            print(f"  ({fantasma} cliente(s) fantasma con diferencia esperada, no cuentan como error)")
        if diferencias:
            print(f"\n{diferencias} cliente(s) con diferencia real de saldo — revisar antes de dar la migración por buena.")
        else:
            print(f"\nOK — los {len(clientes) - fantasma} clientes reales coinciden exactamente entre Access y el saldo recalculado.")
    finally:
        db.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--mdb", default=None, help=r"Ruta al fcmenu.mdb (requerido salvo --fuente csv), ej. C:\FCMenu\fcmenu.mdb")
    parser.add_argument("--fuente", choices=["mdb", "csv"], default="mdb", help="mdb (default): lectura directa vía ODBC. csv: camino viejo con CSV ya exportado.")
    parser.add_argument("--modo", choices=["completo", "prueba"], default="completo")
    parser.add_argument("--anios", type=int, default=10, help="Sólo aplica con --modo prueba.")
    parser.add_argument("--skip-regen", action="store_true", help="No borra/recrea la base, sólo recarga datos sobre la existente.")
    parser.add_argument("--yes", action="store_true", help="No pedir confirmación antes del DROP DATABASE.")
    args = parser.parse_args()

    if args.fuente == "mdb" and not args.mdb:
        parser.error("--mdb es obligatorio con --fuente mdb (default). Usá --fuente csv si querés el camino con CSV ya exportado.")

    app_url = os.environ["DATABASE_URL"]

    if not args.skip_regen:
        admin_url = os.environ["PG_ADMIN_URL"]
        destino = make_url(app_url)
        print(f"Se va a BORRAR y RECREAR la base {destino.database!r} en {destino.host}:{destino.port or 5432}.")
        if not args.yes:
            if input("Escribí SI para confirmar: ").strip().upper() != "SI":
                print("Cancelado.")
                return
        _regenerar_base(app_url, admin_url)
        _aplicar_esquema()

    _cargar_datos(args.fuente, args.mdb, args.modo, args.anios)
    _verificar(args.fuente, args.mdb, args.modo)
    print("\nMigración completa.")


if __name__ == "__main__":
    main()
