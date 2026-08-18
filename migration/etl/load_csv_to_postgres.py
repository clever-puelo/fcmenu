"""Carga los CSV generados por export_access_to_csv.ps1 a PostgreSQL.

Uso:
    DATABASE_URL=postgresql+psycopg2://fcmenu_app:***@localhost:5432/fcmenu_dev \
        .venv/Scripts/python.exe migration/etl/load_csv_to_postgres.py

Requiere que el schema ya esté aplicado (`alembic upgrade head`).

Conversión de tipos: se apoya en las columnas SQLAlchemy de cada modelo
(migration/models.py) para decidir cómo parsear cada valor de texto del
CSV (Decimal para Numeric, date para Date/columnas fecha reales, int para
Integer, str para String/Text). Ver data_layer_progress.md para el detalle
de los casos especiales ya auditados (Imputacion.TIPO/TIPOI texto,
Cotizacion.FECHA texto "dd/mm/yyyy" sin convertir).
"""

from __future__ import annotations

import csv
import sys
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from sqlalchemy import Date, Integer, Numeric
from sqlalchemy.orm import Session

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from migration.db import SessionLocal, engine  # noqa: E402
from migration.models import (  # noqa: E402
    Articulo,
    Banco,
    Cheque,
    Cliente,
    Cotizacion,
    Ctascte,
    Despacho,
    DtoxClte,
    Efectivo,
    Fcestad1,
    FcivaVta,
    Fctabla1,
    Imputacion,
    MovimVS,
    MovStock,
    NotaClte,
    Parametro,
    Proveed,
    Stock,
    Totales,
)

DATA_DIR = Path(__file__).resolve().parent / "data"

# Orden de carga: sin dependencias FK explícitas en el schema actual, pero
# se respeta un orden lógico (catálogos/maestros antes que movimientos)
# por prolijidad y para facilitar diagnóstico si algo falla a mitad de camino.
TABLAS = [
    ("Clientes.csv", Cliente),
    ("Articulo.csv", Articulo),
    ("STOCK.csv", Stock),
    ("Fctabla1.csv", Fctabla1),
    ("Parametro.csv", Parametro),
    ("Proveed.csv", Proveed),
    ("Bancos.csv", Banco),
    ("Cotizacion.csv", Cotizacion),
    ("Ctasctes.csv", Ctascte),
    ("Imputacion.csv", Imputacion),
    ("FcivaVta.csv", FcivaVta),
    ("Fcestad1.csv", Fcestad1),
    ("NOTACLTE.csv", NotaClte),
    ("Cheques.csv", Cheque),
    ("DtoxClte.csv", DtoxClte),
    ("Despachos.csv", Despacho),
    ("MovStock.csv", MovStock),
    ("Totales.csv", Totales),
    ("Efectivo.csv", Efectivo),
    ("MovimVS.csv", MovimVS),
]


def _convertir_valor(raw: str, columna):
    if raw is None or raw == "":
        return None
    tipo = columna.type
    try:
        if isinstance(tipo, Numeric):
            return Decimal(raw)
        if isinstance(tipo, Date):
            return datetime.strptime(raw, "%Y-%m-%d").date()
        if isinstance(tipo, Integer):
            # Access a veces guarda "1" para dbBoolean/dbByte con .0 de por
            # medio si vino de un cálculo — Int() de más no rompe nada.
            return int(float(raw))
        return raw  # String/Text: tal cual
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"No se pudo convertir '{raw}' con columna {columna.name} ({tipo}): {exc}")


def cargar_tabla(db: Session, archivo: str, modelo) -> int:
    ruta = DATA_DIR / archivo
    if not ruta.exists():
        print(f"  [SKIP] {archivo} no existe (¿corriste export_access_to_csv.ps1?)")
        return 0

    columnas = {c.name: c for c in modelo.__table__.columns if c.name != "id"}

    total = 0
    lote = []
    LOTE_SIZE = 2000

    with ruta.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for fila in reader:
            valores = {}
            for nombre_col, columna in columnas.items():
                raw = fila.get(nombre_col)
                valores[nombre_col] = _convertir_valor(raw, columna)
            lote.append(valores)
            if len(lote) >= LOTE_SIZE:
                db.bulk_insert_mappings(modelo, lote)
                db.commit()
                total += len(lote)
                lote = []

    if lote:
        db.bulk_insert_mappings(modelo, lote)
        db.commit()
        total += len(lote)

    return total


def main():
    print(f"Conectando a: {engine.url}")
    db = SessionLocal()
    try:
        for archivo, modelo in TABLAS:
            print(f"Cargando {archivo} -> {modelo.__tablename__} ...")
            n = cargar_tabla(db, archivo, modelo)
            print(f"  {n} filas insertadas")
    finally:
        db.close()
    print("Listo.")


if __name__ == "__main__":
    main()
