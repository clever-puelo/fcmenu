"""Carga los CSV generados por export_access_to_csv.ps1 a PostgreSQL.

Uso (modo completo, igual que antes — corte final a producción real):
    DATABASE_URL=postgresql+psycopg2://fcmenu_app:***@localhost:5432/fcmenu_dev \
        .venv/Scripts/python.exe migration/etl/load_csv_to_postgres.py

Uso (modo prueba — rápido, para re-correr muchas veces durante el
desarrollo, pedido del usuario 2026-08-19: "la migración... se hará
varias veces mientras esté en etapa de pruebas. Debe ser automática,
rápida y segura"):
    ... migration/etl/load_csv_to_postgres.py --modo prueba

El export previo (`export_access_to_csv.ps1 -ModoPrueba`, 32-bit) tiene
el mismo switch — salta Fcestad1/MovStock también del lado de Access,
no sólo acá — para no pagar tampoco el tiempo de exportarlas.

`--modo prueba` hace 2 cosas para acortar el ciclo de prueba (~482.000
de las ~714.000 filas reales están en `Fcestad1`/`MovStock`, ver
`pyqt6_ui_progress.md`):
1. `Fcestad1`/`MovStock` se dejan VACÍAS (no se cargan — son las 2
   tablas más grandes por lejos, y el corte final las carga completas de
   nuevo de todos modos).
2. Las tablas de MOVIMIENTOS/historial (con columna de fecha real) se
   recortan a los últimos `--anios` (default 10) — reduce volumen y, de
   yapa, reduce muchísimo los "huérfanos" reales de cada tabla respecto
   de Clientes/Articulo (verificado contra `fcmenu_dev`: Ctasctes
   30→3, FcivaVta 6→1). Las tablas de CATÁLOGO/CONFIG (Clientes,
   Articulo, Fctabla1, Parametro, DtoxClte, etc. — estado ACTUAL, no
   historial) se cargan siempre completas en los 2 modos: un Cliente
   activo desde hace 15 años tiene que seguir estando en las pruebas
   aunque su historial de Cta.Cte. se recorte.

**TRUNCATE automático antes de cada tabla, en los 2 modos** (antes había
que acordarse de truncar a mano — la causa real del incidente de
duplicación documentado en `cutover_migration_tool.md`) — cada corrida
de este script es idempotente: se puede re-correr las veces que haga
falta sin duplicar nada ni dejar datos de una corrida anterior mezclados
con la nueva. `RESTART IDENTITY CASCADE`: además de la propia tabla,
arrastra (y vuelve a dejar en 0) cualquier tabla que la referencie por
FK — inofensivo acá porque este script igual recarga TODAS las tablas
en la misma corrida.

Requiere que el schema ya esté aplicado (`alembic upgrade head`).

Conversión de tipos: se apoya en las columnas SQLAlchemy de cada modelo
(migration/models.py) para decidir cómo parsear cada valor de texto del
CSV (Decimal para Numeric, date para Date/columnas fecha reales, int para
Integer, str para String/Text). Ver data_layer_progress.md para el detalle
de los casos especiales ya auditados (Imputacion.TIPO/TIPOI texto,
Cotizacion.FECHA texto "dd/mm/yyyy" sin convertir).
"""

from __future__ import annotations

import argparse
import csv
import sys
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Optional

from sqlalchemy import Date, Integer, Numeric, text
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
ANIOS_PRUEBA_DEFAULT = 10

# Las mismas 3 FK reales agregadas en Alembic `bad21413e66c` (NOT VALID
# — toleran el historial ya inconsistente, pero validan todo INSERT/
# UPDATE nuevo). **Se sueltan antes del TRUNCATE+recarga masiva y se
# vuelven a crear al final** (bug real encontrado probando este mismo
# script: un TRUNCATE+INSERT completo reinserta TAMBIÉN las filas
# históricamente huérfanas — para Postgres son inserts nuevos, no datos
# "de antes de la constraint", así que `NOT VALID` no las protege ahí y
# la carga completa rompía a mitad de camino). Soltar/recrear alrededor
# de la recarga masiva es el patrón estándar para este caso.
FKS_CLIENTES = [
    ("fk_ctasctes_clte_clientes", "Ctasctes", "CLTE"),
    ("fk_fcivavta_clte_clientes", "FcivaVta", "CLTE"),
    ("fk_dtoxclte_clte_clientes", "DtoxClte", "CLTE"),
]

# "Clientes/Artículos fantasma" (pedido del usuario, 2026-08-19: en vez
# de tolerar el huérfano con NOT VALID para siempre, o borrar el
# movimiento histórico real que le da origen — "inventemos el cliente y
# el artículo, colocando una descripción de aviso") — después de cada
# recarga masiva, se repone un Cliente/Artículo con este código exacto
# para cualquier CLTE/(COD1,COD2) que algún movimiento real siga
# referenciando pero que ya no exista en el catálogo actual. Con el
# hueco tapado, la FK deja de ser sólo NOT VALID — se puede VALIDAR de
# verdad (ver `_recrear_fks`), y ningún reporte agregado por rango de
# fechas (`ListadosService`) pierde esos importes reales si algún día
# se vuelve a correr sobre ese período.
NOMBRE_CLIENTE_FANTASMA = "[BAJA] Cliente inexistente — recuperado de movimientos históricos"
DESCRI_ARTICULO_FANTASMA = "[BAJA] Artículo discontinuado — recuperado de Despachos históricos"


def _generar_clientes_fantasma(db: Session) -> int:
    codigos = db.execute(
        text(
            '''
            SELECT DISTINCT "CLTE" FROM "Ctasctes"
                WHERE "CLTE" IS NOT NULL AND NOT EXISTS (SELECT 1 FROM "Clientes" cl WHERE cl."CODIGO" = "Ctasctes"."CLTE")
            UNION
            SELECT DISTINCT "CLTE" FROM "FcivaVta"
                WHERE "CLTE" IS NOT NULL AND NOT EXISTS (SELECT 1 FROM "Clientes" cl WHERE cl."CODIGO" = "FcivaVta"."CLTE")
            UNION
            SELECT DISTINCT "CLTE" FROM "DtoxClte"
                WHERE NOT EXISTS (SELECT 1 FROM "Clientes" cl WHERE cl."CODIGO" = "DtoxClte"."CLTE")
            '''
        )
    ).scalars().all()
    for codigo in codigos:
        # CIVA=3 (Consumidor Final) — mismo default que usa el resto de
        # la app cuando un Cliente no tiene Condición de IVA cargada
        # (`civa or 3`, ver FacturadorWindow/etc.); el resto de los
        # campos quedan en blanco, no hace falta inventar más que eso.
        db.execute(
            text('INSERT INTO "Clientes" ("CODIGO", "NOMB", "CIVA") VALUES (:codigo, :nomb, 3)'),
            {"codigo": codigo, "nomb": NOMBRE_CLIENTE_FANTASMA},
        )
    if codigos:
        db.commit()
    return len(codigos)


def _generar_articulos_fantasma(db: Session) -> int:
    """Sólo para `Despachos` — `Fcestad1` queda deliberadamente afuera
    (no es un huérfano a tapar, es el ítem libre "**"/líneas que usan la
    descripción de Sección por diseño, ver `factura_renglon.
    SECCION_ITEM_LIBRE` — inventarle un Artículo ahí estaría mal, no
    sólo de más)."""
    pares = db.execute(
        text(
            '''
            SELECT DISTINCT trim("COD1"), trim("COD2") FROM "Despachos" d
            WHERE NOT EXISTS (
                SELECT 1 FROM "Articulo" a WHERE trim(a."COD1") = trim(d."COD1") AND trim(a."COD2") = trim(d."COD2")
            )
            '''
        )
    ).all()
    for cod1, cod2 in pares:
        # PREC=0 a propósito (no "sin cargar"): `DetalleGrid._aplicar_
        # articulo` ya rechaza facturar un Artículo con precio <= 0 — un
        # Artículo fantasma queda así protegido de venderse por error,
        # sin necesitar una columna de "activo/inactivo" que este schema
        # no tiene.
        db.execute(
            text(
                'INSERT INTO "Articulo" ("COD1", "COD2", "DESCRI", "PREC") VALUES (:cod1, :cod2, :descri, 0)'
            ),
            {"cod1": cod1, "cod2": cod2 or "", "descri": DESCRI_ARTICULO_FANTASMA},
        )
    if pares:
        db.commit()
    return len(pares)


def _soltar_fks(db: Session) -> None:
    for nombre, tabla, _columna in FKS_CLIENTES:
        db.execute(text(f'ALTER TABLE "{tabla}" DROP CONSTRAINT IF EXISTS "{nombre}"'))
    db.commit()


def _recrear_fks(db: Session) -> None:
    """Recrea las 3 FK como `NOT VALID` y, ya con los Clientes fantasma
    puestos (`_generar_clientes_fantasma`, llamado ANTES que esto — ver
    `main()`), intenta VALIDARLAS de verdad en el mismo momento: si no
    queda ningún huérfano real, `VALIDATE CONSTRAINT` las deja tan
    sólidas como una FK agregada desde el día uno (no sólo "protege para
    adelante"). Si por lo que sea todavía quedara alguno (ej. una fecha
    corrida entre que se generaron los fantasma y se recreó la FK, no
    debería pasar en una corrida normal), se avisa pero NO se aborta —
    la FK sigue protegiendo escrituras nuevas igual, sólo queda NOT VALID
    un poco más.

    `DROP CONSTRAINT IF EXISTS` antes de cada `ADD` — sin esto, llamar
    esta función sin haber pasado antes por `_soltar_fks()` (ej. a mano,
    para regenerar los fantasma sobre una base que ya tenía las 3 FK
    puestas) tira `DuplicateObject` y no llega a validar nada; con esto
    es segura de llamar en cualquier estado previo, no sólo dentro del
    flujo normal de `main()`."""
    for nombre, tabla, columna in FKS_CLIENTES:
        db.execute(text(f'ALTER TABLE "{tabla}" DROP CONSTRAINT IF EXISTS "{nombre}"'))
        db.execute(
            text(
                f'ALTER TABLE "{tabla}" ADD CONSTRAINT "{nombre}" '
                f'FOREIGN KEY ("{columna}") REFERENCES "Clientes" ("CODIGO") NOT VALID'
            )
        )
    db.commit()

    for nombre, tabla, _columna in FKS_CLIENTES:
        try:
            db.execute(text(f'ALTER TABLE "{tabla}" VALIDATE CONSTRAINT "{nombre}"'))
            db.commit()
        except Exception as exc:  # noqa: BLE001 — no aborta la migración por esto
            db.rollback()
            print(f"  Aviso: no se pudo validar {nombre} de una ({exc}) — queda NOT VALID, protege igual.")


@dataclass(frozen=True)
class ConfigTabla:
    archivo: str
    modelo: type
    # Columna FECHA/FECENT real (tipo Date, no texto) para el recorte de
    # `--modo prueba` — `None` para catálogos/config (se cargan siempre
    # completos) o para el único caso de fecha-como-texto (`Cotizacion`).
    columna_fecha: Optional[str] = None
    # `Fcestad1`/`MovStock` — ver docstring del módulo.
    excluir_en_modo_prueba: bool = False


# Orden de carga: los catálogos/maestros (Clientes, Articulo) van
# primero — ahora SÍ hace falta respetar este orden (a diferencia del
# comentario viejo, "sin dependencias FK explícitas"): desde 2026-08-19
# hay `ForeignKey()` reales (`Ctascte.CLTE`/`FcivaVta.CLTE`/
# `DtoxClte.CLTE` -> `Clientes.CODIGO`, ver Alembic `bad21413e66c`) que
# exigen que el Cliente ya exista antes de insertar lo que lo referencia.
TABLAS = [
    ConfigTabla("Clientes.csv", Cliente),
    ConfigTabla("Articulo.csv", Articulo),
    ConfigTabla("STOCK.csv", Stock),
    ConfigTabla("Fctabla1.csv", Fctabla1),
    ConfigTabla("Parametro.csv", Parametro),
    ConfigTabla("Proveed.csv", Proveed),
    ConfigTabla("Bancos.csv", Banco),
    ConfigTabla("Cotizacion.csv", Cotizacion),  # FECHA es texto "dd/mm/yyyy" — no se filtra
    ConfigTabla("Ctasctes.csv", Ctascte, columna_fecha="FECHA"),
    ConfigTabla("Imputacion.csv", Imputacion, columna_fecha="FECHA"),
    ConfigTabla("FcivaVta.csv", FcivaVta, columna_fecha="FECHA"),
    ConfigTabla("Fcestad1.csv", Fcestad1, excluir_en_modo_prueba=True),
    ConfigTabla("NOTACLTE.csv", NotaClte),
    ConfigTabla("Cheques.csv", Cheque, columna_fecha="FECEMI"),
    ConfigTabla("DtoxClte.csv", DtoxClte),
    ConfigTabla("Despachos.csv", Despacho, columna_fecha="FECENT"),
    ConfigTabla("MovStock.csv", MovStock, excluir_en_modo_prueba=True),
    ConfigTabla("Totales.csv", Totales, columna_fecha="FECHA"),
    ConfigTabla("Efectivo.csv", Efectivo, columna_fecha="FECHA"),
    ConfigTabla("MovimVS.csv", MovimVS, columna_fecha="FECHA"),
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


def _truncar_tabla(db: Session, modelo) -> None:
    """TRUNCATE automático (siempre, en los 2 modos) — ver docstring del
    módulo: es lo que evita el incidente real de duplicación de la vez
    que se corrió el loader completo sin truncar antes."""
    tabla = modelo.__tablename__
    db.execute(text(f'TRUNCATE TABLE "{tabla}" RESTART IDENTITY CASCADE'))
    db.commit()


def cargar_tabla(db: Session, config: ConfigTabla, *, fecha_corte: Optional[date] = None) -> int:
    """Trunca y recarga una tabla. `fecha_corte`: si se pasa (modo
    prueba) y la tabla tiene `columna_fecha`, se descartan las filas
    anteriores a esa fecha ANTES de insertar — no se cargan de más para
    filtrarlas después."""
    _truncar_tabla(db, config.modelo)

    if config.excluir_en_modo_prueba and fecha_corte is not None:
        print(f"  [VACÍA] {config.archivo} excluida en modo prueba (se recarga completa en el corte final)")
        return 0

    ruta = DATA_DIR / config.archivo
    if not ruta.exists():
        print(f"  [SKIP] {config.archivo} no existe (¿corriste export_access_to_csv.ps1?)")
        return 0

    columnas = {c.name: c for c in config.modelo.__table__.columns if c.name != "id"}
    columna_fecha = config.columna_fecha if fecha_corte is not None else None

    total = 0
    descartadas_por_fecha = 0
    lote = []
    LOTE_SIZE = 2000

    with ruta.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for fila in reader:
            if columna_fecha is not None:
                valor_fecha = _convertir_valor(fila.get(columna_fecha), columnas[columna_fecha])
                if valor_fecha is not None and valor_fecha < fecha_corte:
                    descartadas_por_fecha += 1
                    continue

            valores = {}
            for nombre_col, columna in columnas.items():
                raw = fila.get(nombre_col)
                valores[nombre_col] = _convertir_valor(raw, columna)
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


def _fecha_corte(anios: int) -> date:
    hoy = date.today()
    try:
        return hoy.replace(year=hoy.year - anios)
    except ValueError:
        # 29/feb sin año bisiesto en el año de corte — un día antes.
        return hoy.replace(year=hoy.year - anios, day=28)


def cargar(modo: str = "completo", anios: int = ANIOS_PRUEBA_DEFAULT, tablas_pedidas: Optional[list[str]] = None) -> list[tuple[str, int]]:
    """Cuerpo real de `main()`, extraído para poder llamarse directo
    desde Python (`migrar_completo.py`) sin pasar por un subproceso +
    argparse — necesario para poder empaquetar el orquestador como un
    único .exe (PyInstaller): un subprocess que invoca `sys.executable`
    sobre ESTE archivo no funciona dentro de un .exe congelado, no hay
    intérprete ni archivo .py sueltos ahí adentro."""
    fecha_corte = _fecha_corte(anios) if modo == "prueba" else None

    tablas = TABLAS
    if tablas_pedidas:
        pedidas = set(tablas_pedidas)
        tablas = [c for c in TABLAS if c.modelo.__tablename__ in pedidas]
        encontradas = {c.modelo.__tablename__ for c in tablas}
        faltantes = pedidas - encontradas
        if faltantes:
            print(f"Aviso: no se reconocen estas tablas, se ignoran: {sorted(faltantes)}")

    print(f"Conectando a: {engine.url}")
    print(f"Modo: {modo}" + (f" (últimos {anios} años, desde {fecha_corte})" if fecha_corte else ""))
    db = SessionLocal()
    resumen: list[tuple[str, int]] = []
    try:
        # Sueltas ANTES de la recarga masiva, recreadas SIEMPRE al final
        # (éxito o error) — ver docstring de `FKS_CLIENTES` más arriba.
        _soltar_fks(db)
        try:
            for config in tablas:
                print(f"Cargando {config.archivo} -> {config.modelo.__tablename__} ...")
                n = cargar_tabla(db, config, fecha_corte=fecha_corte)
                print(f"  {n} filas insertadas")
                resumen.append((config.modelo.__tablename__, n))

            n_clientes = _generar_clientes_fantasma(db)
            if n_clientes:
                print(f"\n{n_clientes} Cliente(s) fantasma repuesto(s) — \"{NOMBRE_CLIENTE_FANTASMA}\"")
            n_articulos = _generar_articulos_fantasma(db)
            if n_articulos:
                print(f"{n_articulos} Artículo(s) fantasma repuesto(s) — \"{DESCRI_ARTICULO_FANTASMA}\"")
        finally:
            _recrear_fks(db)
    finally:
        db.close()

    print("\nResumen:")
    for tabla, n in resumen:
        print(f"  {tabla}: {n}")
    print("Listo.")
    return resumen


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--modo",
        choices=["completo", "prueba"],
        default="completo",
        help='"completo" (default): todo, sin recortes — corte final a producción real. '
        '"prueba": Fcestad1/MovStock vacías + últimos N años en el resto de las tablas de movimientos.',
    )
    parser.add_argument(
        "--anios",
        type=int,
        default=ANIOS_PRUEBA_DEFAULT,
        help=f"Sólo aplica con --modo prueba. Default: {ANIOS_PRUEBA_DEFAULT}.",
    )
    parser.add_argument(
        "--tablas",
        nargs="+",
        default=None,
        help="Nombres de tabla (__tablename__, ej. Ctasctes Articulo) para recargar sólo un subconjunto.",
    )
    args = parser.parse_args()
    cargar(args.modo, args.anios, args.tablas)


if __name__ == "__main__":
    main()
