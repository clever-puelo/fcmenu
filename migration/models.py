from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import Optional

from sqlalchemy import Date, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


class Cliente(Base):
    __tablename__ = "Clientes"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CODIGO: Mapped[int] = mapped_column(Integer, nullable=False)
    NOMB: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    DIR: Mapped[Optional[str]] = mapped_column(String(120), nullable=True)
    LOC: Mapped[Optional[str]] = mapped_column(String(80), nullable=True)
    CP: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    PCIA: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    TEL1: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TEL2: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    FAX: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    EMAIL: Mapped[Optional[str]] = mapped_column(String(120), nullable=True)
    ZONA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    VEND: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CVTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DTO1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DTO2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DTO3: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DTO4: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DTO5: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CIVA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CUIT: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    NGAN: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    IB: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    CORR1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CORR2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CANAL: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PROIND: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    CREDIT: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DEUDA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    FALTA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FACTUAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Articulo(Base):
    __tablename__ = "Articulo"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD1: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    COD2: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    DESCRI: Mapped[Optional[str]] = mapped_column(String(120), nullable=True)
    STOCK: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PREC: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    VTA1: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    VTA2: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PCOS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DTO1: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    STMIN: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MCALIS: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    STANT: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    GPTIPO: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    GPDSD: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    GPHST: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FACTUAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Stock(Base):
    __tablename__ = "STOCK"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD1: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    COD2: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    STMIN: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    STMAX: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    STREP: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DEP1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    EST1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ESTT1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DEP2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    EST2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ESTT2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    STANT: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PULG: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    MTR: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    STUNID: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    ENTMES: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    SALMES: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    AJEMES: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    AJSMES: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FACTUAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Fctabla1(Base):
    __tablename__ = "Fctabla1"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CTAB: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    COD: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    DESCRI: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    ALF1: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF2: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF3: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF4: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF5: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF6: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF7: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    ALF8: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    NUMSD1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUMSD2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUMSD3: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUMCD1: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    NUMCD2: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    NUMCD3: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    FECHA1: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FECHA2: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    MCA1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MCA2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MCA3: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FECACT: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Parametro(Base):
    __tablename__ = "Parametro"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CLAVE: Mapped[Optional[str]] = mapped_column(String(20), nullable=False)
    NUME1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME2: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME3: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME4: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME5: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME6: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME7: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME8: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME9: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME10: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME11: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME12: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME13: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME14: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME15: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME16: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME17: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME18: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME19: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NUME20: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    IVAINS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    IVANI: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FECACT: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    LISFORD: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    FECFORD: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    LISVW: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    FECVW: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    LIMVTA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PTOVTA: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    NOMEMPR: Mapped[Optional[str]] = mapped_column(String(80), nullable=True)
    MCAIB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)


class Ctascte(Base):
    __tablename__ = "Ctasctes"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    LETRA: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    TIPO: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PREFIJO: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    IMPUT1: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    IMPUT2: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    IMPUT3: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    IMPUT4: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    IMPUT5: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    IMPUT6: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    DEBE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    IMPTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    BON: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    FECVTO: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    CVTA: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    TIPO9: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    MOTI: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    USUAR: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)


class Imputacion(Base):
    """Detalle histórico de aplicación de pagos (recibos) a comprobantes.

    Schema confirmado vía DAO contra fcmenu.mdb real (52.227 filas en
    producción) — no inferido del VB6. OJO: a diferencia de Ctasctes.TIPO
    (numérico, dbByte), acá TIPO y TIPOI son TEXTO (dbText, 2 chars) —
    VB6 los puebla asignando un entero a un campo texto (EmiRec.frm:1119,
    1140), por eso se mapean como str, no int.

    Se inserta un registro por cada aplicación real de un recibo
    (TIPO="4") y, si hubo descuento por pronto pago, otro más (TIPO="6")
    — ver EmiRec.frm:1103-1157.
    """

    __tablename__ = "Imputacion"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    TIPO: Mapped[Optional[str]] = mapped_column(String(2), nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    IMPTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    TIPOI: Mapped[Optional[str]] = mapped_column(String(2), nullable=True)
    CPBTEI: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CORR: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FECHAI: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FECVTOI: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(6), nullable=True)


class FcivaVta(Base):
    __tablename__ = "FcivaVta"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    TIPO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    LETRA: Mapped[Optional[str]] = mapped_column(String(10), nullable=True)
    PTOVTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NOMB: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    PCIA: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    CVTA: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    MOTI: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    CIVA: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    CUIT: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    VEND: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ZONA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    TOTCAN: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    GRINS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    GRNOINS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    IVAINS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    IVANOINS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    EXENTO: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PORCIB: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    TOTIB: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    TOTCOS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    BON: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    ITEMS: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    COMIS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    NOIMPR: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)


class Fcestad1(Base):
    __tablename__ = "Fcestad1"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD1: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    COD2: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TIPO: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    LETRA: Mapped[Optional[str]] = mapped_column(String(10), nullable=True)
    PTOVTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ITEM: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    GPTIPO: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    GPDSD: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    GPHST: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    PULG: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    MTR: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    MILIM: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    TELAS: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CANT: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PCOS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PVTA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESP: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    BON: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    IVA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    IMPTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    TIPO9: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)


class NotaClte(Base):
    """Notas libres por cliente.

    TITULO1-4/NOTA1-4/PIE1-4 son campos Memo (sin límite fijo) en el
    fcmenu.mdb real (confirmado vía DAO: type=12, size=0) — se mapean a
    Text, no a String acotado, para no truncar notas largas al migrar.
    """

    __tablename__ = "NOTACLTE"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=False)
    TITULO1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO2: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA2: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE2: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO3: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA3: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE3: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO4: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA4: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE4: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FECACT: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class NotaArticulo(Base):
    """Notas libres por artículo (tabla real `Notartic`).

    No estaba entre las 14 tablas migradas en la fase de datos original
    — se agregó recién en la fase de UI (2026-08-15) al migrar
    `Notartic.frm`. Esquema auditado vía DAO 32-bit contra fcmenu.mdb
    real: COD1 text(5), COD2 text(10), TITULO1-5/NOTA1-5/PIE1-5 Memo
    (type=12, sin límite — igual que NotaClte). A diferencia de
    NotaClte son 5 juegos (1-5), no 4, y no tiene USUARIO/FECACT.
    `Notartic.frm` sólo usa NOTA3 ("Ventas") y NOTA4 ("Especial") — el
    resto queda disponible para una futura función de impresión.
    Clave real compuesta COD1+COD2 (Sección+Código del artículo, igual
    que `Articulo` — ver nota de padding en `ArticuloRepository`).
    """

    __tablename__ = "Notartic"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD1: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    COD2: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TITULO1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO2: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA2: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE2: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO3: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA3: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE3: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO4: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA4: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE4: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    TITULO5: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    NOTA5: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    PIE5: Mapped[Optional[str]] = mapped_column(Text, nullable=True)


class DtoxClte(Base):
    """Descuentos por Sección y por Cliente (tabla real `DtoxClte`,
    editada desde `DtosxClte.frm`). Fuente real de los 5 niveles de
    descuento en cascada que ya usa
    `FacturaService.calcular_bonificacion_cascada()` (citada ahí como
    "EmiFact.frm:1821-1827 (DxClte + bucle...)" — `DxClte` en
    FCMENU.bas:605 es la sub que carga estas filas en memoria al armar
    una factura).

    No estaba entre las 14 tablas migradas en la fase de datos original
    — se agregó recién acá (2026-08-15). Esquema auditado vía DAO
    32-bit: `CLTE` Integer, `SECCION` text(5), `DTO1-5`/`RGO1-3` Double,
    `USUARIO` text(6), `FACTUAL` Date. 12.035 filas reales.

    **`RGO1-3` ("Rango") están confirmados muertos, no sólo en el código
    (`Rgos()` en FCMENU.bas se carga pero nunca se lee en ningún
    cálculo, búsqueda exhaustiva) sino también en los datos reales:
    las 12.035 filas tienen `RGO1=RGO2=RGO3=0`, sin una sola excepción.**
    Se mantienen en el modelo por fidelidad de esquema pero no se
    exponen en la UI — ver `dtos_cliente_dialog.py`.
    """

    __tablename__ = "DtoxClte"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=False)
    SECCION: Mapped[Optional[str]] = mapped_column(String(5), nullable=True)
    DTO1: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DTO2: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DTO3: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DTO4: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    DTO5: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    RGO1: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    RGO2: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    RGO3: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FACTUAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Despacho(Base):
    """Lotes/despachos con stock disponible, por Artículo (tabla real
    `Despachos`, editada desde `StockDespa.frm`/`VerDesp.frm` y consultada
    desde `DetFact.frm` al cargar un ítem de factura — permite elegir un
    lote puntual con `STOCK > 0` en vez de descontar del stock general).

    No estaba entre las 14+1 tablas migradas hasta ahora — se agregó acá
    (2026-08-16, fase Facturador) porque `DetFact.frm` la usa activamente.
    Esquema auditado vía DAO 32-bit contra `fcmenu.mdb` real: `COD1`
    text(5), `COD2` text(10), `NRODESP` text(20), `CPBTE` Long, `ENTRADA`/
    `SALIDAS`/`STOCK` Double, `FECENT`/`FECUSAL` Date. **6.161 filas
    reales, con `STOCK > 0` genuino** (ej. lote "18001IC04080837W", 62 u.
    disponibles) — a diferencia de `Notartic`, acá el usuario pidió migrar
    los datos reales.

    Mismo patrón de padding de Access en COD1/COD2 que `Articulo` — ver
    `DespachoRepository`.
    """

    __tablename__ = "Despachos"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD1: Mapped[Optional[str]] = mapped_column(String(5), nullable=True)
    COD2: Mapped[Optional[str]] = mapped_column(String(10), nullable=True)
    NRODESP: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ENTRADA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    SALIDAS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    STOCK: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    FECENT: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FECUSAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class MovStock(Base):
    """Movimientos de stock (tabla real `MovStock`, consultada desde
    `Stock.frm` como historial de movimientos por artículo). No estaba
    entre las tablas migradas hasta ahora — se agregó acá (2026-08-16,
    fase Facturador, servicio de emisión) porque `EmiFact.frm Sub Graba`
    inserta una fila por cada renglón facturado.

    Esquema auditado vía DAO 32-bit contra `fcmenu.mdb` real: `COD1`
    text(5), `COD2` text(10), `TIPO` text(2) (código de movimiento, ej.
    "21"=Fact.'A', "22"=Fact.'B' — se graba como texto, no numérico),
    `PTOVTA`/`ITEM`/`MILIM`/`TELAS` enteros chicos, `CPBTE` Long,
    `PULG`/`MTR`/`CANT` Currency. **244.310 filas reales** — migradas
    completas (decisión del usuario, mismo criterio que Ctasctes/
    Fcestad1/Despacho: mantener continuidad histórica real para
    `Stock.frm`).
    """

    __tablename__ = "MovStock"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD1: Mapped[Optional[str]] = mapped_column(String(5), nullable=True)
    COD2: Mapped[Optional[str]] = mapped_column(String(10), nullable=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    TIPO: Mapped[Optional[str]] = mapped_column(String(2), nullable=True)
    PTOVTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ITEM: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PULG: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    MTR: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    MILIM: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    TELAS: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CANT: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(10), nullable=True)


class Totales(Base):
    """Contador diario agregado de facturación (tabla real `Totales`, PK
    real por `FECHA`, una fila por día — consultada desde `TotFact.frm`
    y `VerFact.frm` como reporte de totales del período). No estaba
    entre las tablas migradas hasta ahora — se agregó acá (2026-08-16,
    fase Facturador, servicio de emisión).

    Esquema auditado vía DAO 32-bit contra `fcmenu.mdb` real: todos los
    contadores (`FACA`/`FACB`/... `UNIDA`/`UNIDB`/...) son Long, los
    importes (`PVTA`/`PCOS`/`PESP`/`PESPA`/`PESPB`/...) son Currency.
    Se mapean TODAS las columnas reales por fidelidad de esquema, aunque
    el servicio de emisión actual (alcance Factura A/B) sólo escribe
    `FACA`/`FACB`/`UNIDA`/`UNIDB`/`PESPA`/`PESPB`/`PVTA`/`PCOS`/`PESP` —
    el resto (Notas de Créd./Déb., Mostrador, Exportación, Devolución)
    corresponde a circuitos todavía no migrados. **3.730 filas reales**
    (~10 años de historial diario) — migradas completas (decisión del
    usuario, mismo criterio que las demás tablas de historial).
    """

    __tablename__ = "Totales"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FACA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FACB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FACC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NDA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NDB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NDC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MOSTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MOSTB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MOSTC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCMOSTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCMOSTB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCMOSTC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNIDA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNIDB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNIDC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DEV: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    MOSTDEV: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNMOSTA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNMOSTB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNMOSTC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PVTA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PCOS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESP: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPB: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPC: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPMOSTA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPMOSTB: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPMOSTC: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    EXPA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    EXPB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCEXPA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NCEXPB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PESPEXPA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PESPEXPB: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    UNEXPA: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNEXPB: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    UNEXPC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    EXPDEV: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ACUCPRAS: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(6), nullable=True)
    FACTUAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Proveed(Base):
    __tablename__ = "Proveed"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CODIGO: Mapped[int] = mapped_column(Integer, nullable=False)
    RSOCIAL: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    CLASIF: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    DIREC: Mapped[Optional[str]] = mapped_column(String(120), nullable=True)
    LOCAL: Mapped[Optional[str]] = mapped_column(String(80), nullable=True)
    PCIA: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    CODPOS: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    CONTACT: Mapped[Optional[str]] = mapped_column(String(80), nullable=True)
    TEL1: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TEL2: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TEL3: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TEL4: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TEL5: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    TIPO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    CPAGO: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    CIVA: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    CUIT: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    CORR1: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FUCPRA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    UCPBTE: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    UIMPTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    # REFER1 es Memo real (type=12) en fcmenu.mdb — Text, no String acotado.
    REFER1: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    # REFER2 real es TEXT(50) — el String(30) previo era más chico que el dato real.
    REFER2: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    MCACONT: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    PROVREF: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    COMPRADOR: Mapped[Optional[str]] = mapped_column(String(30), nullable=True)
    HABLIN: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    USUAR: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FACTUAL: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Banco(Base):
    __tablename__ = "Bancos"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    COD: Mapped[int] = mapped_column(Integer, nullable=False)
    SUC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NOMBRE: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    NOMSUC: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    DIR: Mapped[Optional[str]] = mapped_column(String(120), nullable=True)
    LOC: Mapped[Optional[str]] = mapped_column(String(80), nullable=True)
    CP: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    CLEARING: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    PROPIO: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)


class Cheque(Base):
    __tablename__ = "Cheques"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    NROCHEQ: Mapped[int] = mapped_column(Integer, nullable=False)
    BCOSUC: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBCO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    DATOSAD: Mapped[Optional[str]] = mapped_column(String(40), nullable=True)
    TIPING: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBING: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CONCEP: Mapped[Optional[str]] = mapped_column(String(60), nullable=True)
    IMPORTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    PROPIO: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    ORDEN: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    TIPEGR: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBEGR: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    DESTINO: Mapped[Optional[str]] = mapped_column(String(60), nullable=True)
    OBSERV: Mapped[Optional[str]] = mapped_column(String(120), nullable=True)
    ESTADO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    MONEDA: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    COTIZ: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    MOTRECH: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FECEMI: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FECVTO: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FECEGR: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    FECRECH: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    NROOPE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    NROOPS: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FECING: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Cotizacion(Base):
    """Cotización diaria del dólar.

    OJO: a diferencia de todos los demás campos FECHA del sistema (dbDate),
    Cotizacion.FECHA está guardado como TEXT(10) en el fcmenu.mdb real
    (confirmado vía DAO: type=10, size=10 — probablemente "dd/mm/yyyy").
    Se mapea como String, no Date; conversión a date real queda pendiente
    para cuando se migren los datos (parsear con el formato real de la
    tabla antes de insertar en PostgreSQL).
    """

    __tablename__ = "Cotizacion"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    FECHA: Mapped[Optional[str]] = mapped_column(String(10), nullable=False)
    DOLAR: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    DOLARA: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    DOLARB: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    DOLARC: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    DOLARD: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    DOLARE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 4), nullable=True)
    USUARIO: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    FECACT: Mapped[Optional[date]] = mapped_column(Date, nullable=True)


class Efectivo(Base):
    """Pago en efectivo de un Recibo (`EmiRec.frm Sub Graba()`, bloque
    "Efectivo", líneas 1350-1370).

    No estaba entre las tablas ya migradas — se agregó acá (2026-08-15,
    fase Cuenta Corriente/Recibo). Esquema auditado vía DAO 32-bit
    contra `fcmenu.mdb` real: `FECHA` Date, `TIPO` Byte, `CPBTE`
    Integer, `IMPTE` Currency — sin columna de identidad propia (Access
    no tenía PK en esta tabla, igual que `Ctasctes`/`Imputacion`). 6.164
    filas reales migradas.

    `TIPO` siempre vale 4 en los datos reales (mismo código que
    `Ctascte.TIPO=4`, "Recibo") — se mapea igual como Integer por
    fidelidad de esquema, aunque el emisor de Recibos sólo escribe ese
    valor.
    """

    __tablename__ = "Efectivo"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    TIPO: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    IMPTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)


class MovimVS(Base):
    """Anticipos/Retenciones/Tarjetas/Transferencias/Baja Incobrable de
    un Recibo — la contraparte de `Efectivo`/`Cheques` para las demás
    formas de pago del combo "Operación" de `DetPago.frm` (Retenc.Gan.,
    Retenc.IB, Retenc.IVA, Retenc.SUS, Dep./Transf., Tarj.Déb/Cr.,
    Cheques Vs., Cheq.Electr., Baja Incobr. — `TIPREG` 1-9, mismo orden
    que `Combo1` de `DetPago.frm`). Fuente: `EmiRec.frm Sub Graba()`,
    bloque "aAnticipos" (líneas ~1318-1343).

    No estaba entre las tablas ya migradas — se agregó acá (2026-08-15,
    fase Cuenta Corriente/Recibo). Esquema auditado vía DAO 32-bit:
    `CLTE` Integer, `FECHA` Date, `TIPO` **text(2)**, `CPBTE` Integer,
    `TIPREG` **text(2)**, `IMPTE` Currency, `BCOSUC` text(6),
    `DATOSAD`/`CONCEP`/`OBSERV` text(50). 46.929 filas reales migradas.

    OJO: `TIPO`/`TIPREG` parecen códigos numéricos (`TIPO` siempre "4" en
    los datos reales, mismo código que `Ctascte.TIPO=4`; `TIPREG` es el
    código de "Operación" 1-9 del combo de `DetPago.frm`) pero el DAO
    real los reporta `type=10` (dbText), no numérico — igual quirk ya
    visto en `Imputacion.TIPO`/`TIPOI`. Confirmado con datos reales: hay
    filas con `TIPREG` en blanco (`"  "`), que un `int()` directo rompe
    — se mapean como `String`, no `Integer`.
    """

    __tablename__ = "MovimVS"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    CLTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    FECHA: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    TIPO: Mapped[Optional[str]] = mapped_column(String(2), nullable=True)
    CPBTE: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    TIPREG: Mapped[Optional[str]] = mapped_column(String(2), nullable=True)
    IMPTE: Mapped[Optional[Decimal]] = mapped_column(Numeric(12, 2), nullable=True)
    BCOSUC: Mapped[Optional[str]] = mapped_column(String(6), nullable=True)
    DATOSAD: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    CONCEP: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    OBSERV: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)


__all__ = [
    "Base",
    "Cliente",
    "Articulo",
    "Stock",
    "Fctabla1",
    "Parametro",
    "Ctascte",
    "Imputacion",
    "FcivaVta",
    "Fcestad1",
    "NotaClte",
    "Proveed",
    "Banco",
    "Cheque",
    "Cotizacion",
    "Efectivo",
    "MovimVS",
]
