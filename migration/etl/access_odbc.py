"""Lectura directa de fcmenu.mdb vía ODBC (pyodbc) — reemplaza al par
`export_access_to_csv.ps1` (DAO, 32-bit) + CSV intermedio.

Requiere el driver "Microsoft Access Database Engine 2016 Redistributable"
instalado en la PC donde se corre esto, EN LA MISMA ARQUITECTURA que el
Python (o el .exe) que lo usa — de 64 bits, si Python/el .exe son de 64
bits (que es lo normal en una PC nueva sin Office instalado; no hace
falta la versión de 32 bits salvo que ya haya Office 32-bit en esa PC,
en cuyo caso instalar la de 32 bits y correr con un Python de 32 bits).

Descarga: buscar "Microsoft Access Database Engine 2016 Redistributable"
en el sitio de Microsoft, elegir AccessDatabaseEngine_X64.exe.

A diferencia del camino CSV (que necesitaba formatear fechas a ISO y
forzar cultura invariante para los decimales a mano, ver el .ps1),
pyodbc devuelve los valores ya tipados (date/datetime/bool/Decimal) —
no hace falta ningún parseo de texto.
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import Any, Iterator

import pyodbc
from sqlalchemy import Date, Integer, Numeric


def conectar_mdb(ruta_mdb: str) -> pyodbc.Connection:
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={ruta_mdb};"
    )
    return pyodbc.connect(conn_str, autocommit=True)


def leer_tabla(conn: pyodbc.Connection, nombre_tabla: str) -> Iterator[dict[str, Any]]:
    """Itera las filas de una tabla de Access como dicts {nombre_campo: valor_nativo}."""
    cur = conn.cursor()
    cur.execute(f"SELECT * FROM [{nombre_tabla}]")
    nombres = [col[0] for col in cur.description]
    for fila in cur:
        yield dict(zip(nombres, fila))
    cur.close()


def convertir_valor(raw: Any, columna) -> Any:
    """Adapta el valor nativo que devuelve pyodbc al tipo Python que
    espera la columna SQLAlchemy destino — equivalente al `_convertir_
    valor`/`Format-Valor` de los caminos anteriores (CSV), pero mucho
    más simple porque acá no hay que parsear texto: pyodbc ya entrega
    fechas como date/datetime y booleanos como bool."""
    if raw is None:
        return None
    tipo = columna.type

    if isinstance(tipo, Numeric):
        return raw if isinstance(raw, Decimal) else Decimal(str(raw))

    if isinstance(tipo, Date):
        if isinstance(raw, datetime):
            return raw.date()
        if isinstance(raw, date):
            return raw
        return raw  # no debería llegar acá para una columna Date real

    if isinstance(tipo, Integer):
        if isinstance(raw, bool):  # dbBoolean (Articulo.MCALIS, Fctabla1.MCA1-3, ...)
            return int(raw)
        return int(raw)

    # String/Text (incluye Cotizacion.FECHA, que es TEXT "dd/mm/yyyy" en
    # el origen, no una fecha real — pyodbc ya la entrega como str, sin
    # necesitar el caso especial que sí hacía falta en el camino CSV).
    return str(raw) if not isinstance(raw, str) else raw
