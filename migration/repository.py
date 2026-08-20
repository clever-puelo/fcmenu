from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import Any, Generic, List, Optional, Type, TypeVar

from sqlalchemy import Numeric, case, cast, func
from sqlalchemy.orm import Session

from .models import (
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
    NotaArticulo,
    NotaClte,
    Parametro,
    Proveed,
    Stock,
    Totales,
)

T = TypeVar("T")


class BaseRepository(Generic[T]):
    """Repositorio genérico CRUD para cualquier modelo SQLAlchemy."""

    def __init__(self, db: Session, model: Type[T]):
        self.db = db
        self.model = model

    def create(self, obj_in: dict[str, Any]) -> T:
        """Crea un nuevo registro."""
        db_obj = self.model(**obj_in)
        self.db.add(db_obj)
        self.db.commit()
        self.db.refresh(db_obj)
        return db_obj

    def read(self, obj_id: int) -> Optional[T]:
        """Lee un registro por ID."""
        return self.db.query(self.model).filter(self.model.id == obj_id).first()

    def read_all(self, skip: int = 0, limit: int = 100) -> List[T]:
        """Lee todos los registros con paginación."""
        return self.db.query(self.model).offset(skip).limit(limit).all()

    def update(self, obj_id: int, obj_in: dict[str, Any]) -> Optional[T]:
        """Actualiza un registro."""
        db_obj = self.read(obj_id)
        if db_obj:
            for key, value in obj_in.items():
                setattr(db_obj, key, value)
            self.db.commit()
            self.db.refresh(db_obj)
        return db_obj

    def delete(self, obj_id: int) -> bool:
        """Elimina un registro."""
        db_obj = self.read(obj_id)
        if db_obj:
            self.db.delete(db_obj)
            self.db.commit()
            return True
        return False


class ClienteRepository(BaseRepository[Cliente]):
    """Repositorio específico para Clientes."""

    def __init__(self, db: Session):
        super().__init__(db, Cliente)

    def by_codigo(self, codigo: int) -> Optional[Cliente]:
        """Busca cliente por CODIGO."""
        return self.db.query(Cliente).filter(Cliente.CODIGO == codigo).first()

    def by_nombre(self, nombre: str) -> List[Cliente]:
        """Busca clientes por nombre (LIKE)."""
        return self.db.query(Cliente).filter(Cliente.NOMB.ilike(f"%{nombre}%")).all()

    def by_cuit(self, cuit: str) -> Optional[Cliente]:
        """Busca cliente por CUIT."""
        return self.db.query(Cliente).filter(Cliente.CUIT == cuit).first()

    def activos(self) -> List[Cliente]:
        """Retorna clientes activos (USUARIO != NULL)."""
        return self.db.query(Cliente).filter(Cliente.USUARIO.isnot(None)).all()

    def desde_codigo(self, codigo: int, limite: int = 200) -> List[Cliente]:
        """Clientes con CODIGO >= codigo, ordenados de menor a mayor.

        Usado por `ClienteBusquedaWindow` cuando el operador tipea un
        código de cliente que no existe (feedback del usuario, 2026-08-15):
        en vez de un simple "no encontrado", se le ofrece la lista real
        ordenada arrancando justo desde el número que ingresó, para que
        pueda ubicar el código correcto sin tener que adivinar un filtro
        de texto. `limite` evita traer la tabla entera de un tirón."""
        return (
            self.db.query(Cliente)
            .filter(Cliente.CODIGO >= codigo)
            .order_by(Cliente.CODIGO.asc())
            .limit(limite)
            .all()
        )

    def by_vendedor(self, vend: int) -> List[Cliente]:
        """Clientes asignados a un Vendedor (Fctabla1 CTAB='VD') — usado para
        validar la baja del Vendedor en el ABM de Tablas."""
        return self.db.query(Cliente).filter(Cliente.VEND == vend).all()

    def by_zona(self, zona: int) -> List[Cliente]:
        """Clientes asignados a una Zona (Fctabla1 CTAB='ZN')."""
        return self.db.query(Cliente).filter(Cliente.ZONA == zona).all()

    def by_pcia(self, letra: str) -> List[Cliente]:
        """Clientes de una Provincia (código de 1 letra, ver
        `migration.provincias`) — réplica de `Listados.frm SelClientes`
        ("Por Provincia")."""
        return self.db.query(Cliente).filter(func.trim(Cliente.PCIA) == letra.strip()).all()

    def todos_ordenados_por_nombre(self) -> List[Cliente]:
        """Todos los clientes, orden alfabético — réplica de
        `Listados.frm` ("Todos", `ORDER BY NOMB`), usado por varios de
        los listados (Clientes/Saldos/Estado de Cuenta/Planilla de
        Cobranzas)."""
        return self.db.query(Cliente).order_by(Cliente.NOMB.asc()).all()

    def by_cond_venta(self, cvta: int) -> List[Cliente]:
        """Clientes con una Condición de Venta (Fctabla1 CTAB='CV')."""
        return self.db.query(Cliente).filter(Cliente.CVTA == cvta).all()

    def proximo_codigo(self) -> int:
        """Próximo código de cliente para Alta.

        Fuente: Abmclte.frm Sub Blanquea() —
        "SELECT TOP 1 CODIGO FROM Clientes WHERE CODIGO < 99900 ORDER BY
        CODIGO DESC" + 1. Los códigos >= 99900 quedan reservados (clientes
        genéricos/de prueba tipo "Consumidor Final") y no participan del
        autonumerado.
        """
        ultimo = (
            self.db.query(Cliente)
            .filter(Cliente.CODIGO < 99900)
            .order_by(Cliente.CODIGO.desc())
            .first()
        )
        return (ultimo.CODIGO + 1) if ultimo else 1

    @staticmethod
    def proximo_correlativo(cliente: Cliente) -> int:
        """Próximo valor de `Cliente.CORR1` — el "ciclo" 1-99 que
        `EmisionReciboService` usa para etiquetar juntos un Recibo y los
        comprobantes que paga (ver `IMPUT1-6` de `Ctasctes` y `Corr` de
        `Imputacion`). Réplica de `EmiRec.frm Sub Graba()` línea 1035-1036:
        `Corr = Val(CORR1) + 1`, y si da 0 o pasa de 99 vuelve a 1."""
        corr = (cliente.CORR1 or 0) + 1
        if corr <= 0 or corr > 99:
            corr = 1
        return corr


class NotaClteRepository(BaseRepository[NotaClte]):
    """Repositorio específico para Notas de Cliente (NotaClte.frm).

    Sólo `NOTA1`/`NOTA2` ("Principal"/"Adicional") se editan desde ese
    formulario — TITULO1-4/PIE1-4/NOTA3-4 existen en el modelo (son
    campos Memo reales de `fcmenu.mdb`) pero ninguna pantalla ya migrada
    los usa; quedan disponibles para una futura función de impresión.
    """

    def __init__(self, db: Session):
        super().__init__(db, NotaClte)

    def by_cliente(self, clte: int) -> Optional[NotaClte]:
        """Nota de un cliente puntual (relación 1 a 1 real: el legacy
        siempre hace `SELECT TOP 1 ... WHERE CLTE = ...` y opera sobre esa
        única fila)."""
        return self.db.query(NotaClte).filter(NotaClte.CLTE == clte).first()


class DtoxClteRepository(BaseRepository[DtoxClte]):
    """Repositorio específico para Descuentos por Sección y Cliente
    (DtosxClte.frm). Ver docstring de `DtoxClte` en models.py — RGO1-3
    existen en el esquema pero están confirmados muertos (código y
    datos reales), no se exponen acá con getters propios."""

    def __init__(self, db: Session):
        super().__init__(db, DtoxClte)

    def by_cliente(self, clte: int) -> List[DtoxClte]:
        """Todas las filas de descuento de un cliente, una por Sección."""
        return self.db.query(DtoxClte).filter(DtoxClte.CLTE == clte).all()

    def by_cliente_seccion(self, clte: int, seccion: str) -> Optional[DtoxClte]:
        """El descuento de un cliente para una Sección puntual — usado por
        `EmisionFacturaService` para recalcular la cascada de bonificación
        en el momento de emitir (réplica de `DxClte` llamado fresco desde
        `EmiFact.frm Sub Graba`, no reutiliza el % ya mostrado en pantalla
        durante la carga del renglón)."""
        return (
            self.db.query(DtoxClte)
            .filter(
                (DtoxClte.CLTE == clte) & (func.trim(DtoxClte.SECCION) == seccion.strip())
            )
            .first()
        )

    def reemplazar_para_cliente(self, clte: int, filas: list[dict]) -> int:
        """Reemplaza todos los descuentos de un cliente por los de `filas`
        (cada dict con al menos SECCION/DTO1-5).

        Réplica de `DtosxClte.frm Sub Grabar()`: borra todo lo existente
        para el cliente y vuelve a insertar. **No replica un bug real del
        legacy**: si el operador repetía una Sección en la grilla, el
        `Exit Sub` de la línea 856 abandonaba la transacción completa sin
        `CommitTrans` ni `RollbackTrans` (nada se guardaba, ni siquiera lo
        ya insertado esa vuelta). Acá se deduplica por Sección
        (última fila del formulario gana) en vez de abortar todo.
        """
        existentes = self.by_cliente(clte)
        for fila in existentes:
            self.db.delete(fila)

        por_seccion: dict[str, dict] = {}
        for fila in filas:
            seccion = (fila.get("SECCION") or "").strip()
            if seccion:
                por_seccion[seccion] = fila

        for seccion, datos in por_seccion.items():
            nueva = DtoxClte(CLTE=clte, **{k: v for k, v in datos.items() if k != "SECCION"}, SECCION=seccion)
            self.db.add(nueva)

        self.db.commit()
        return len(por_seccion)


class ArticuloRepository(BaseRepository[Articulo]):
    """Repositorio específico para Artículos."""

    def __init__(self, db: Session):
        super().__init__(db, Articulo)

    def by_cod1(self, cod1: str) -> Optional[Articulo]:
        """Busca artículo por COD1 (código primario).

        Compara con `func.trim()`: en los datos reales migrados desde
        Access, COD1/COD2 quedaron con el padding fijo original (p.ej.
        `'GPN  '`, `'0300      '`) — ver nota de padding en
        `FctablasRepository.by_ctab_cod`, mismo fenómeno acá.
        """
        return self.db.query(Articulo).filter(func.trim(Articulo.COD1) == cod1.strip()).first()

    def by_cod2(self, cod2: str) -> Optional[Articulo]:
        """Busca artículo por COD2 (código secundario). Ver nota de padding
        en `by_cod1`."""
        return self.db.query(Articulo).filter(func.trim(Articulo.COD2) == cod2.strip()).first()

    def by_seccion(self, cod1: str) -> List[Articulo]:
        """Todos los artículos de una Sección, ordenados por COD2. Usado
        por la ventana de Búsqueda de Artículos: elegir una Sección lista
        todos sus artículos abajo."""
        return (
            self.db.query(Articulo)
            .filter(func.trim(Articulo.COD1) == cod1.strip())
            .order_by(Articulo.COD2)
            .all()
        )

    def todos(self) -> List[Articulo]:
        """Todos los artículos del catálogo (modo "Todos" de ModPrec.frm:
        "SELECT * FROM Articulo ORDER BY COD1, COD2 DESC")."""
        return (
            self.db.query(Articulo)
            .order_by(Articulo.COD1, Articulo.COD2.desc())
            .all()
        )

    def by_descri(self, descri: str) -> List[Articulo]:
        """Busca artículos por descripción (LIKE)."""
        return self.db.query(Articulo).filter(Articulo.DESCRI.ilike(f"%{descri}%")).all()

    def by_cod1_cod2(self, cod1: str, cod2: str) -> Optional[Articulo]:
        """Busca artículo por su clave real compuesta (Sección + Código).

        Fuente: AbmArt.frm Grabacion() — "WHERE COD1 = ... and COD2 = ...".
        COD1 por sí solo NO es la clave del artículo (es la Sección, la
        comparten muchos artículos); la clave real es el par COD1+COD2.
        Ver nota de padding en `by_cod1`.
        """
        return (
            self.db.query(Articulo)
            .filter(
                (func.trim(Articulo.COD1) == cod1.strip())
                & (func.trim(Articulo.COD2) == cod2.strip())
            )
            .first()
        )

    def bajo_stock(self) -> List[Articulo]:
        """Retorna artículos bajo stock (STOCK < STMIN).

        AVISO (verificado contra VB6): Articulo.STOCK/STMIN son campos que
        solo se editan a mano desde AbmArt.frm — la venta (EmiFact.frm) NUNCA
        los actualiza, solo toca Stock.STUNID en la tabla STOCK. Este método
        puede devolver datos desactualizados. Para stock crítico real, usar
        StockRepository.criticos() / StockService.
        """
        return self.db.query(Articulo).filter(Articulo.STOCK < Articulo.STMIN).all()


class NotaArticuloRepository(BaseRepository[NotaArticulo]):
    """Repositorio específico para Notas de Artículo (Notartic.frm).

    Sólo `NOTA3`/`NOTA4` ("Ventas"/"Especial") se editan desde ese
    formulario — ver docstring de `NotaArticulo` para el resto de los
    campos (disponibles pero sin UI que los use todavía).
    """

    def __init__(self, db: Session):
        super().__init__(db, NotaArticulo)

    def by_articulo(self, cod1: str, cod2: str) -> Optional[NotaArticulo]:
        """Nota de un artículo puntual (clave real COD1+COD2, ver nota de
        padding en `ArticuloRepository.by_cod1_cod2`)."""
        return (
            self.db.query(NotaArticulo)
            .filter(
                (func.trim(NotaArticulo.COD1) == cod1.strip())
                & (func.trim(NotaArticulo.COD2) == cod2.strip())
            )
            .first()
        )


class StockRepository(BaseRepository[Stock]):
    """Repositorio específico para Stock."""

    def __init__(self, db: Session):
        super().__init__(db, Stock)

    def by_cod1(self, cod1: str) -> Optional[Stock]:
        """Busca stock por COD1."""
        return self.db.query(Stock).filter(Stock.COD1 == cod1).first()

    def by_cod2(self, cod2: str) -> Optional[Stock]:
        """Busca stock por COD2."""
        return self.db.query(Stock).filter(Stock.COD2 == cod2).first()

    def by_cod1_cod2(self, cod1: str, cod2: str) -> Optional[Stock]:
        """Busca el registro de stock de un artículo puntual (clave real
        COD1+COD2). Compara con `func.trim()` — mismo bug de padding fijo
        de Access confirmado en `STOCK` (ej. `'A    '`/`'47        '`),
        mismo patrón que `ArticuloRepository`/`DespachoRepository` — usado
        por el servicio de emisión (`EmisionFacturaService`) para el
        upsert de stock equivalente a `EmiFact.frm Sub Graba`."""
        return (
            self.db.query(Stock)
            .filter(
                (func.trim(Stock.COD1) == cod1.strip())
                & (func.trim(Stock.COD2) == cod2.strip())
            )
            .first()
        )

    def criticos(self) -> List[Stock]:
        """Retorna items de stock crítico (STUNID < STMIN).

        Fuente de verdad confirmada contra VB6 (EmiFact.frm): la venta
        descuenta de Stock.STUNID, no de EST1/EST2 (esquema multi-depósito
        legacy que la facturación no toca). Por eso se filtra por STUNID.
        """
        return self.db.query(Stock).filter(Stock.STUNID < Stock.STMIN).all()


class DespachoRepository(BaseRepository[Despacho]):
    """Repositorio específico para Despachos/Lotes — consultado desde
    `DetFact.frm Sub VerSiBusca` al cargar un ítem de factura, y desde
    `VerDesp.frm` (consulta de despachos) y `Stock.frm` (carga de
    movimientos por importación). `StockDespa.frm` (ABM aparte de esta
    tabla) está confirmado MUERTO — no figura en `FCMENU.vbp`.

    Compara COD1/COD2 con `func.trim()` — mismo patrón de padding de
    Access que `ArticuloRepository` (ver nota ahí)."""

    def __init__(self, db: Session):
        super().__init__(db, Despacho)

    def con_stock(self, cod1: str, cod2: str) -> List[Despacho]:
        """Despachos con stock disponible (`STOCK > 0`) para un artículo,
        para el selector de lote de `DetFact.frm` (columna FG2 del legacy).
        """
        return (
            self.db.query(Despacho)
            .filter(
                (func.trim(Despacho.COD1) == cod1.strip())
                & (func.trim(Despacho.COD2) == cod2.strip())
                & (Despacho.STOCK > 0)
            )
            .order_by(Despacho.FECENT)
            .all()
        )

    def by_nrodesp(self, cod1: str, cod2: str, nrodesp: str) -> Optional[Despacho]:
        """Un despacho puntual por su número (clave real usada al descontar
        stock en `EmiFact.frm Sub Graba`)."""
        return (
            self.db.query(Despacho)
            .filter(
                (func.trim(Despacho.COD1) == cod1.strip())
                & (func.trim(Despacho.COD2) == cod2.strip())
                & (Despacho.NRODESP == nrodesp)
            )
            .first()
        )

    def resumen_lotes(self) -> List[dict]:
        """Lista de lotes/despachos reales (NRODESP+FECENT), con la
        cantidad de artículos que tiene cada uno — réplica de
        `VerDesp.frm Sub CargaDespa` (`GROUP BY NRODESP, FECENT`), orden
        real por fecha descendente (`FG2.Sort = flexSortGenericDescending`
        sobre la columna Fecha).

        `cpbte` (pedido del usuario, 2026-08-19: "agregar comprobante
        despues del despacho") — el Nº de comprobante de ENTRADA de ese
        lote (`Despachos.CPBTE`), en teoría constante para todas las
        filas de un mismo NRODESP+FECENT (mismo lote de importación);
        `func.max()` por si algún caso real no lo fuera, para no romper
        el `GROUP BY` con una columna no agregada."""
        filas = (
            self.db.query(Despacho.NRODESP, Despacho.FECENT, func.count(Despacho.id), func.max(Despacho.CPBTE))
            .group_by(Despacho.NRODESP, Despacho.FECENT)
            .order_by(Despacho.FECENT.desc())
            .all()
        )
        return [
            {"nrodesp": nrodesp, "fecent": fecent, "cantidad": cantidad, "cpbte": cpbte}
            for nrodesp, fecent, cantidad, cpbte in filas
        ]

    def articulos_de_lote(self, nrodesp: str, fecent: Optional[date] = None) -> List[Despacho]:
        """Todos los renglones (uno por artículo) de un lote/despacho
        puntual — réplica de `VerDesp.frm Sub DoVer` (`WHERE NRODESP = ...`,
        sin filtrar por artículo).

        **Bug real encontrado con datos reales** (2026-08-19, "22001IC05005447M
        dice que tiene 2 artículos y al desplegar muestra más de 15
        renglones"): un mismo `NRODESP` puede repetirse en fechas de
        entrada DISTINTAS (confirmado contra `fcmenu_dev`: ese lote tiene
        27 filas el 17/03/2022 y 2 más el 05/01/2023 — 29 en total bajo
        el mismo número) — `resumen_lotes()` ya agrupa por NRODESP+FECENT
        (son lotes de importación distintos que reusan el número), pero
        antes acá se filtraba sólo por NRODESP: al abrir CUALQUIERA de
        los 2 renglones del resumen se traían las 29 filas juntas, sin
        importar cuál se clickeó. `fecent` (opcional sólo para no romper
        el único caller de antes de este fix) ahora filtra también por
        la fecha exacta del renglón del resumen elegido."""
        filtros = [Despacho.NRODESP == nrodesp]
        if fecent is not None:
            filtros.append(Despacho.FECENT == fecent)
        return (
            self.db.query(Despacho)
            .filter(*filtros)
            .order_by(Despacho.COD1.asc(), Despacho.COD2.asc())
            .all()
        )


class MovStockRepository(BaseRepository[MovStock]):
    """Repositorio específico para Movimientos de Stock (`Stock.frm`,
    historial de movimientos por artículo — ver `MovStock` en models.py).
    Usado por `EmisionFacturaService` para insertar el movimiento de cada
    renglón facturado (equivalente a `EmiFact.frm Sub Graba`)."""

    def __init__(self, db: Session):
        super().__init__(db, MovStock)

    def by_cod1_cod2(self, cod1: str, cod2: str) -> List[MovStock]:
        """Historial de movimientos de un artículo puntual, más recientes
        primero — mismo `func.trim()` de padding que `ArticuloRepository`."""
        return (
            self.db.query(MovStock)
            .filter(
                (func.trim(MovStock.COD1) == cod1.strip())
                & (func.trim(MovStock.COD2) == cod2.strip())
            )
            .order_by(MovStock.FECHA.desc())
            .all()
        )

    def existe_comprobante(self, tipo: str, cpbte: int) -> bool:
        """Réplica de `Stock.frm Sub BuscaCPBTE()` (líneas 822-868): antes
        de grabar un movimiento nuevo, verifica que no exista ya uno con
        el mismo `TIPO`+`CPBTE` — evita cargar dos veces el mismo
        comprobante. `tipo` es el código real de 2 dígitos (ver
        `StockMovimientoService.CODIGOS_ENTRADA/CODIGOS_SALIDA`), no el
        `EntrSali`/`CodEntr` de la UI."""
        return (
            self.db.query(MovStock)
            .filter(MovStock.TIPO == tipo, MovStock.CPBTE == cpbte)
            .first()
            is not None
        )


class TotalesRepository(BaseRepository[Totales]):
    """Repositorio específico para el contador diario agregado de
    facturación (`TotFact.frm`/`VerFact.frm`, PK real por FECHA — ver
    `Totales` en models.py). Usado por `EmisionFacturaService` para el
    upsert diario equivalente a `EmiFact.frm Sub Graba` (`AlTotal:`)."""

    def __init__(self, db: Session):
        super().__init__(db, Totales)

    def by_fecha(self, fecha: date) -> Optional[Totales]:
        """Fila del día, si ya existe (no hay padding que tolerar acá:
        FECHA es una columna Date real, no texto, a diferencia de
        `Cotizacion.FECHA`)."""
        return self.db.query(Totales).filter(Totales.FECHA == fecha).first()

    def entre_fechas(
        self, fecha_desde: date, fecha_hasta_inclusiva: date, limite: Optional[int] = None
    ) -> List[Totales]:
        """Filas entre `fecha_desde` y `fecha_hasta_inclusiva` (ambas
        inclusive), ordenadas por FECHA ascendente — réplica de
        `TotFact.frm Sub DoVer2`/`CargaGrilla` ("Muestra desde ... los
        primeros ...", mismo mecanismo ya migrado en `VerFact.frm`).
        `limite=None` trae todas las filas (usado por el resumen
        mensual, que no tiene tope de N — `Sub MuestraTotal` no usa
        `Combo1`)."""
        query = (
            self.db.query(Totales)
            .filter((Totales.FECHA >= fecha_desde) & (Totales.FECHA <= fecha_hasta_inclusiva))
            .order_by(Totales.FECHA)
        )
        if limite is not None:
            query = query.limit(limite)
        return query.all()


class ParametroRepository(BaseRepository[Parametro]):
    """Repositorio específico para Parámetros."""

    def __init__(self, db: Session):
        super().__init__(db, Parametro)

    def by_clave(self, clave: str) -> Optional[Parametro]:
        """Busca parámetro por CLAVE.

        Compara con `func.trim()`: en los datos reales migrados desde
        Access, CLAVE quedó con el padding fijo original (`'1 '`, no
        `'1'`) — mismo fenómeno de padding que `Fctabla1.CTAB` y
        `Articulo.COD1/COD2` (ver notas de padding en esos repositorios).
        Sin trim, `get_config()` devolvía `None` contra la base real pese
        a que la fila de configuración sí existe — rompía silenciosamente
        todo lo que depende de ella (IVAINS/IVANI en FacturaService,
        deuda en ClienteService, etc.).
        """
        return self.db.query(Parametro).filter(func.trim(Parametro.CLAVE) == clave.strip()).first()

    def get_config(self) -> Optional[Parametro]:
        """Obtiene la fila única de configuración global (CLAVE = '1').

        Confirmado en FCMENU.bas:333: todo el sistema arranca leyendo
        "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'" y de ahí toma
        IVAINS, IVANI, NOMEMPR, NUME20 (días de deuda), LIMVTA (monto de
        deuda) y MCAIB (flag de percepción IIBB). No es una tabla de
        parámetros por clave arbitraria: es una sola fila de config.
        """
        return self.by_clave("1")

    def get_ivains(self) -> Optional[Parametro]:
        """Obtiene la fila de configuración (para leer .IVAINS / .IVANI).

        NOTA: antes buscaba CLAVE='IVAINS', que no existe — corregido para
        usar la fila de config real (CLAVE='1'), ver get_config().
        """
        return self.get_config()

    def get_punto_venta(self) -> Optional[Parametro]:
        """Obtiene la fila de configuración (para leer .PTOVTA)."""
        return self.get_config()


# Etiquetas de Ctascte.TIPO para grillas/reportes — réplica exacta de
# `CtaCte.frm:1962-1975` (fuente canónica, usada en 2 lugares distintos
# del propio legacy). `DetRec.frm` tenía un label DISTINTO y equivocado
# para TIPO=7 ("Ctdo." en vez de "RecAn.") — no se replica ese bug, ver
# `pyqt6_ui_progress.md` (feedback del usuario, 2026-08-15).
ETIQUETAS_TIPO_CTASCTE = {
    0: "Sdo.A.",
    1: "Fact.",
    2: "NCréd.",
    3: "NDéb.",
    4: "Recibo",
    5: "PACta.",
    6: "Descto",
    7: "RecAn.",
    8: "NCInt",
    9: "Anul.",
}

# Etiquetas de FcivaVta.TIPO — réplica exacta de `VerFact.frm Sub
# CargaGrilla` (líneas 1149-1172). PARECIDAS pero NO IDÉNTICAS a
# ETIQUETAS_TIPO_CTASCTE: comparten significado en 0-6 y 9, pero
# divergen en 7/8 porque son dos enumeraciones de dos tablas distintas
# (FCIVAVTA = subdiario de ventas, con "Factura Mostrador"/"Devolución
# Mostrador" como tipos propios; Ctascte = movimientos de cuenta
# corriente, con "Recibo Anulado"/"Nota de Crédito Interna") — no es
# una inconsistencia a resolver, no mezclar los dos diccionarios.
ETIQUETAS_TIPO_FCIVAVTA = {
    0: "Sdo.A.",
    1: "Fact.",
    2: "NCréd.",
    3: "NDéb.",
    4: "Recibo",
    5: "PACta.",
    6: "Descto",
    7: "F.Most.",
    8: "Dev.M.",
    9: "Anul.",
}


class CtascteRepository(BaseRepository[Ctascte]):
    """Repositorio específico para Cuentas Corrientes."""

    def __init__(self, db: Session):
        super().__init__(db, Ctascte)

    def by_cliente(self, clte: int) -> List[Ctascte]:
        """Retorna movimientos de cuenta corriente de un cliente."""
        return self.db.query(Ctascte).filter(Ctascte.CLTE == clte).all()

    # TIPO de CtasCtes que suman al saldo deudor del cliente (algoritmo real
    # extraído de CtaCte.frm:2278-2289 y CtaCte.frm:1941-1950):
    #   0 = Saldo Anterior, 1 = Factura, 3 = Nota de Débito, 7 = Recibo Anulado
    # Cualquier otro TIPO (2=NC, 4=Recibo, 5=Pago a Cta, 6=Descuento,
    # 8=NC Interés, 9=Anulación) resta.
    TIPOS_DEBE = {0, 1, 3, 7}

    def deuda_cliente(self, clte: int) -> dict[str, Any]:
        """Calcula el saldo de cuenta corriente de un cliente sumando
        Ctascte.IMPTE según TIPO (no confundir con el campo Ctascte.DEBE,
        que es el saldo pendiente puntual de una factura/ND, usado solo
        para el seguimiento de imputaciones parciales — ver CuentaCorrienteService).

        NOTA: previamente este método sumaba mal Ctascte.DEBE como si fuera
        el "debe" contable; se corrigió tras auditar CtaCte.frm (regla de
        oro del proyecto: reportar diferencias en saldos de CC antes de asumir).
        """
        movimientos = self.by_cliente(clte)
        saldo = Decimal("0")
        for m in movimientos:
            importe = m.IMPTE or Decimal("0")
            saldo += importe if m.TIPO in self.TIPOS_DEBE else -importe
        return {"cliente": clte, "saldo": saldo, "deuda": saldo}

    def by_comprobante(self, tipo: int, letra: str, prefijo: int, cpbte: int) -> List[Ctascte]:
        """Busca movimientos de CC por comprobante."""
        return (
            self.db.query(Ctascte)
            .filter(
                (Ctascte.TIPO == tipo)
                & (Ctascte.LETRA == letra)
                & (Ctascte.PREFIJO == prefijo)
                & (Ctascte.CPBTE == cpbte)
            )
            .all()
        )

    # TIPO que puede aparecer como "pendiente de cobro" en un Recibo nuevo —
    # réplica exacta de DetRec.frm Sub DoVer3() (líneas 510-512): TIPO NOT IN
    # (0=Sdo.Ant., 2=NC, 4=Recibo, 6=Descuento, 8=NC Interés, 9=Anulación),
    # o sea TIPO IN (1=Factura, 3=ND, 5=Pago a Cta, 7=Recibo Anulado) — un
    # Recibo Anulado sí puede reaparecer legítimamente como deuda a cobrar
    # (ver docstring de EmisionReciboService).
    TIPOS_PENDIENTES_COBRO = {1, 3, 5, 7}

    def pendientes_cobro(self, clte: int) -> List[Ctascte]:
        """Comprobantes con saldo pendiente (`DEBE > 0`) que se le pueden
        aplicar a un Recibo nuevo, ordenados por fecha (igual que la
        grilla "Pendiente" de `DetRec.frm`). `TIPO9 <> '1'` excluye
        comprobantes tipo Cotización (Letra "C", fuera de alcance)."""
        return (
            self.db.query(Ctascte)
            .filter(
                Ctascte.CLTE == clte,
                Ctascte.TIPO.in_(self.TIPOS_PENDIENTES_COBRO),
                Ctascte.DEBE > 0,
                (Ctascte.TIPO9.is_(None)) | (Ctascte.TIPO9 != "1"),
            )
            .order_by(Ctascte.FECHA.asc())
            .all()
        )

    def factura_mas_antigua_impaga(self, clte: int) -> Optional[Ctascte]:
        """La Factura (TIPO=1) impaga más vieja de un cliente (`DEBE>1`,
        la más antigua por `FECHA`) — réplica de `BusClte.frm Function
        TieneDeuda()` (líneas 1531-1538), usada para el "Aviso de Deuda"
        del Facturador (`ConDeuda.frm`)."""
        return (
            self.db.query(Ctascte)
            .filter(Ctascte.CLTE == clte, Ctascte.TIPO == 1, Ctascte.DEBE > 1)
            .order_by(Ctascte.FECHA.asc(), Ctascte.CPBTE.asc(), Ctascte.TIPO.asc())
            .first()
        )

    def facturas_nd_impagas(self, clte: int) -> List[Ctascte]:
        """Facturas/ND (`TIPO IN (1,3)`) con `DEBE > 1` de un cliente —
        réplica de `VerCobra.frm Sub CargaGrilla` (líneas 340-341,
        "Cobranzas por Zona"). El umbral `>1` (no `>0`) es el mismo ya
        usado en `factura_mas_antigua_impaga()` — descarta residuos de
        centavos de redondeo como si fueran deuda real."""
        return (
            self.db.query(Ctascte)
            .filter(Ctascte.CLTE == clte, Ctascte.TIPO.in_((1, 3)), Ctascte.DEBE > 1)
            .order_by(Ctascte.FECHA.asc(), Ctascte.CPBTE.asc(), Ctascte.TIPO.asc())
            .all()
        )

    def saldo_inicial(self, clte: int, anio: int) -> dict[str, Decimal]:
        """"Saldo Anterior" del extracto anual — acumulado de TODOS los años
        previos a `anio` (no sólo el inmediato anterior), separado en
        Debe/Crédito con la MISMA fórmula real del saldo corriente
        (`TIPOS_DEBE`, ya confirmada en `deuda_cliente()`).

        `CtaCte.frm:1833-1836` (`Sub DoVer3`) usaba acá una fórmula
        DISTINTA e incompleta (sólo TIPO 1/3 como débito y 2/4 como
        crédito, ignorando 0/5/6/7/8/9) — en una sesión previa el usuario
        había confirmado replicarla tal cual por comparabilidad
        histórica, pero al probar contra datos reales (cuenta 407,
        2026-08-16) confirmó CORREGIRLA: esa fórmula incompleta hacía que
        "mostrar todo" (desde 2005) y filtrar "desde 2013" dieran saldos
        finales DISTINTOS para el mismo cliente (diferencia de $412,86,
        exactamente los descuentos TIPO=6 anteriores a 2013 que la
        fórmula vieja no contaba) — inaceptable para un operador
        manejando plata real. Con `TIPOS_DEBE` el Saldo Inicial de
        cualquier año siempre reconcilia con el saldo final de "mostrar
        todo", sin importar qué año se elija como punto de partida."""
        movimientos = (
            self.db.query(Ctascte)
            .filter(Ctascte.CLTE == clte, func.extract("year", Ctascte.FECHA) < anio)
            .all()
        )
        debe = Decimal("0")
        credito = Decimal("0")
        for m in movimientos:
            importe = m.IMPTE or Decimal("0")
            if m.TIPO in self.TIPOS_DEBE:
                debe += importe
            else:
                credito += importe
        return {"debe": debe, "credito": credito}

    def extracto(self, clte: int, anio: int) -> List[Ctascte]:
        """Movimientos de `anio` en adelante, para el extracto anual —
        réplica de CtaCte.frm:1853-1856 (`Sub DoVer3`), orden real
        `FECHA, CPBTE, TIPO`."""
        return (
            self.db.query(Ctascte)
            .filter(Ctascte.CLTE == clte, func.extract("year", Ctascte.FECHA) >= anio)
            .order_by(Ctascte.FECHA.asc(), Ctascte.CPBTE.asc(), Ctascte.TIPO.asc())
            .all()
        )

    def saldo_vencido(self, clte: int, hoy: Optional[date] = None) -> Decimal:
        """Porción vencida del saldo del cliente — campo "Sdo.Venc." que
        `CtaCte.frm` declaraba pero nunca poblaba (ver `pyqt6_ui_progress.md`).
        Fórmula confirmada con el usuario (2026-08-16): suma el `DEBE`
        todavía impago de las Facturas/ND (TIPO 1/3) cuya `FECVTO` (fecha
        de vencimiento real, ya calculada y grabada por
        `EmisionFacturaService` al emitir — `Cliente.CVTA` → días de la
        Cond. de Venta) ya pasó."""
        hoy = hoy or date.today()
        movimientos = (
            self.db.query(Ctascte)
            .filter(
                Ctascte.CLTE == clte,
                Ctascte.TIPO.in_((1, 3)),
                Ctascte.DEBE > 0,
                Ctascte.FECVTO.isnot(None),
                Ctascte.FECVTO < hoy,
            )
            .all()
        )
        return sum((m.DEBE or Decimal("0") for m in movimientos), Decimal("0"))

    def by_tipo_y_rango(self, tipo: int, fecha_desde: date, fecha_hasta_exclusiva: date) -> List[Ctascte]:
        """Movimientos de un TIPO puntual (ej. 4=Recibo) en un rango de
        fechas — réplica de `Listados.frm SelSubCobzasAFIP` ("Subdiario
        de Cobranzas AFIP")."""
        return (
            self.db.query(Ctascte)
            .filter(
                (Ctascte.TIPO == tipo)
                & (Ctascte.FECHA >= fecha_desde)
                & (Ctascte.FECHA < fecha_hasta_exclusiva)
            )
            .order_by(Ctascte.FECHA.asc(), Ctascte.CPBTE.asc())
            .all()
        )

    def cobranzas_y_nd_rechazadas(
        self, clte: int, fecha_desde: date, fecha_hasta_exclusiva: date
    ) -> List[Ctascte]:
        """Recibos (TIPO=4) y Pagos a Cta. (TIPO=5) de un cliente en el
        rango, MÁS las Notas de Débito por cheque rechazado (TIPO=3,
        MOTI=' 2' — código real de `Fctabla1` MT, confirmado contra
        datos reales de `fcmenu_dev`) — réplica de `Listados.frm
        LisComCobNVO` (líneas 3410-3412, "Comisiones por Cobranzas").
        Las ND de cheque rechazado se incluyen en el listado pero NO
        suman al total de comisión (ver `ListadosService.
        comisiones_cobranzas`, que replica esa exclusión)."""
        return (
            self.db.query(Ctascte)
            .filter(
                Ctascte.CLTE == clte,
                Ctascte.FECHA >= fecha_desde,
                Ctascte.FECHA < fecha_hasta_exclusiva,
                (Ctascte.TIPO.in_((4, 5))) | ((Ctascte.TIPO == 3) & (func.trim(Ctascte.MOTI) == "2")),
            )
            .order_by(Ctascte.FECHA.asc())
            .all()
        )

    def saldos_todos_clientes(self) -> List[dict[str, Any]]:
        """Saldo de TODOS los clientes, calculado con la misma fórmula real
        de `deuda_cliente()`/`TIPOS_DEBE` (réplica de CtaCte.frm `CargaFG8`,
        líneas 2259-2297) pero agregado en una sola consulta SQL en vez de
        iterar cliente por cliente en Python — el legacy hacía un `SELECT`
        + loop por cada uno de los ~2774 clientes con `DoEvents` para no
        colgar la UI; acá el mismo resultado sale de un único `GROUP BY`.
        Ordenado por Nombre, igual que `DoVer8` (`ORDER BY NOMB`)."""
        importe = func.coalesce(Ctascte.IMPTE, 0)
        signo = func.sum(
            cast(case((Ctascte.TIPO.in_(tuple(self.TIPOS_DEBE)), importe), else_=-importe), Numeric(14, 2))
        )
        filas = (
            self.db.query(Cliente.CODIGO, Cliente.NOMB, func.coalesce(signo, 0))
            .outerjoin(Ctascte, Ctascte.CLTE == Cliente.CODIGO)
            .group_by(Cliente.CODIGO, Cliente.NOMB)
            .order_by(Cliente.NOMB.asc())
            .all()
        )
        return [{"codigo": codigo, "nombre": nombre, "saldo": Decimal(saldo)} for codigo, nombre, saldo in filas]


class ImputacionRepository(BaseRepository[Imputacion]):
    """Repositorio específico para el detalle histórico de imputaciones.

    OJO: TIPO/TIPOI son texto (ver Imputacion en models.py) — los filtros
    acá comparan contra str, no int.
    """

    def __init__(self, db: Session):
        super().__init__(db, Imputacion)

    def by_cliente(self, clte: int) -> List[Imputacion]:
        """Retorna todas las imputaciones (aplicaciones de pago) de un cliente."""
        return self.db.query(Imputacion).filter(Imputacion.CLTE == clte).all()

    def by_comprobante_original(self, clte: int, tipoi: str, cpbtei: int) -> List[Imputacion]:
        """Retorna las aplicaciones registradas contra un comprobante original
        puntual (ej. qué recibos/descuentos tocaron la Factura Nº cpbtei)."""
        return (
            self.db.query(Imputacion)
            .filter(
                (Imputacion.CLTE == clte)
                & (Imputacion.TIPOI == tipoi)
                & (Imputacion.CPBTEI == cpbtei)
            )
            .all()
        )

    def by_recibo(self, clte: int, tipo: str, cpbte: int) -> List[Imputacion]:
        """Retorna el detalle de aplicaciones que generó un recibo puntual."""
        return (
            self.db.query(Imputacion)
            .filter(
                (Imputacion.CLTE == clte)
                & (Imputacion.TIPO == tipo)
                & (Imputacion.CPBTE == cpbte)
            )
            .all()
        )


class FcivaVtaRepository(BaseRepository[FcivaVta]):
    """Repositorio específico para Facturación IVA (encabezados)."""

    def __init__(self, db: Session):
        super().__init__(db, FcivaVta)

    def by_comprobante(
        self, tipo: str, letra: str, ptovta: int, cpbte: int
    ) -> Optional[FcivaVta]:
        """Busca factura por tipo, letra, punto de venta y comprobante.

        **Bug real encontrado y corregido** (2026-08-19, "desde Cuentas
        Corrientes no funciona el click para ver la factura"):
        `FcivaVta.TIPO` viene de `fcmenu_dev` real con el padding fijo de
        Access (`'1 '`, no `'1'` — confirmado contra datos reales), a
        diferencia de `Ctascte.TIPO`/`Fcestad1.TIPO` (`Integer`, sin este
        problema) — el `==` exacto contra el `str(tipo)` limpio que arma
        `CtaCteWindow._on_click_fila_extracto` nunca encontraba nada,
        `by_comprobante` devolvía `None` en silencio y el clic no hacía
        nada visible. `func.trim()` en TIPO/LETRA, mismo patrón ya usado
        en `Fctabla1.CTAB`/`Articulo.COD1-2`/`Parametro.CLAVE` (4ª vez
        que este bug de padding aparece en una columna de texto migrada
        desde Access)."""
        return (
            self.db.query(FcivaVta)
            .filter(
                (func.trim(FcivaVta.TIPO) == tipo.strip())
                & (func.trim(FcivaVta.LETRA) == letra.strip())
                & (FcivaVta.PTOVTA == ptovta)
                & (FcivaVta.CPBTE == cpbte)
            )
            .first()
        )

    def by_cliente(self, clte: int) -> List[FcivaVta]:
        """Retorna facturas de un cliente."""
        return self.db.query(FcivaVta).filter(FcivaVta.CLTE == clte).all()

    def facturas_de_cliente(self, clte: int) -> List[FcivaVta]:
        """Sólo Facturas (`TIPO='1'`) de un cliente, más recientes
        primero — para elegir la Factura "original" a devolver en la
        Nota de Crédito por Devolución de Mercadería (`Motivo=1`,
        `NotaCreditoMercaderiaWindow`). A diferencia de `by_cliente()`
        no trae NC/ND/Cotizaciones ya emitidas contra este cliente."""
        return (
            self.db.query(FcivaVta)
            .filter((FcivaVta.CLTE == clte) & (FcivaVta.TIPO == "1"))
            .order_by(FcivaVta.FECHA.desc(), FcivaVta.CPBTE.desc())
            .all()
        )

    def by_fecha_rango(self, fecha_desde, fecha_hasta) -> List[FcivaVta]:
        """Retorna facturas en un rango de fechas."""
        return (
            self.db.query(FcivaVta)
            .filter((FcivaVta.FECHA >= fecha_desde) & (FcivaVta.FECHA <= fecha_hasta))
            .all()
        )

    def ultima_venta(self, clte: int) -> Optional[date]:
        """Fecha de la última venta del cliente (`MAX(FECHA)`) — para el
        campo "Última Venta" de la Consulta de Cta.Cte (`CtaCte.frm`
        declaraba "Últ.Vta.Ctd."/"Últ.Vta.C.Cte." por separado, pero no
        hay forma confiable de clasificar Contado/Cta.Cte en los datos
        reales de Cond. de Venta — 45 valores de texto libre sin bandera
        booleana; el usuario confirmó 2026-08-16 unificarlo en un solo
        campo)."""
        return self.db.query(func.max(FcivaVta.FECHA)).filter(FcivaVta.CLTE == clte).scalar()

    def by_fecha_rango_vendedor(
        self, fecha_desde: date, fecha_hasta_exclusiva: date, vend: Optional[int] = None
    ) -> List[FcivaVta]:
        """Comprobantes del subdiario en un rango, opcionalmente
        filtrados por Vendedor — réplica de `Listados.frm
        LisSubVtasCom` ("Subdiario Vtas. para Comisiones"), orden real
        `zona, vend, fecha, tipo, ptovta, cpbte`."""
        query = self.db.query(FcivaVta).filter(
            (FcivaVta.FECHA >= fecha_desde) & (FcivaVta.FECHA < fecha_hasta_exclusiva)
        )
        if vend is not None:
            query = query.filter(FcivaVta.VEND == vend)
        return query.order_by(
            FcivaVta.ZONA.asc(), FcivaVta.VEND.asc(), FcivaVta.FECHA.asc(),
            FcivaVta.TIPO.asc(), FcivaVta.PTOVTA.asc(), FcivaVta.CPBTE.asc(),
        ).all()

    def con_percepcion_ib(self, fecha_desde: date, fecha_hasta_exclusiva: date) -> List[FcivaVta]:
        """Comprobantes con percepción de Ingresos Brutos aplicada
        (`TOTIB > 0`) en un rango — réplica de `Listados.frm
        SelPerARBA` ("Percepciones ARBA")."""
        return (
            self.db.query(FcivaVta)
            .filter(
                (FcivaVta.TOTIB > 0)
                & (FcivaVta.FECHA >= fecha_desde)
                & (FcivaVta.FECHA < fecha_hasta_exclusiva)
            )
            .order_by(FcivaVta.PTOVTA.asc(), FcivaVta.CPBTE.asc())
            .all()
        )

    def resumen_por_provincia_y_tipo(self, fecha_desde: date, fecha_hasta_exclusiva: date) -> List[dict[str, Any]]:
        """Suma de `GRINS`/`IVAINS` agrupada por Provincia+TIPO en un
        rango — réplica de `Listados.frm LisIngBrutos` ("Ingresos
        Brutos"), mismo `GROUP BY PCIA, TIPO ORDER BY PCIA, TIPO` del
        legacy, hecho en SQL en vez de traer todas las filas."""
        filas = (
            self.db.query(
                FcivaVta.PCIA,
                FcivaVta.TIPO,
                func.coalesce(func.sum(FcivaVta.GRINS), 0),
                func.coalesce(func.sum(FcivaVta.IVAINS), 0),
            )
            .filter((FcivaVta.FECHA >= fecha_desde) & (FcivaVta.FECHA < fecha_hasta_exclusiva))
            .group_by(FcivaVta.PCIA, FcivaVta.TIPO)
            .order_by(FcivaVta.PCIA.asc(), FcivaVta.TIPO.asc())
            .all()
        )
        return [
            {"pcia": pcia, "tipo": int(tipo) if tipo is not None and str(tipo).strip().isdigit() else None,
             "grins": Decimal(grins), "ivains": Decimal(ivains)}
            for pcia, tipo, grins, ivains in filas
        ]

    def emitidas_desde(self, fecha_desde: date, fecha_hasta_exclusiva: date, limite: int) -> List[FcivaVta]:
        """Comprobantes emitidos entre `fecha_desde` (inclusive) y
        `fecha_hasta_exclusiva` (exclusive), ordenados por FECHA/CPBTE
        ascendente y recortados a los primeros `limite` — réplica de
        `VerFact.frm Sub DoVer2`/`CargaGrilla` ("Muestra desde ... los
        primeros ...").

        El filtro de "Sucursal" (`Frame1` del .frm: Todas/Cta.Cte./
        Mostrador/Manuales/Empresas/Exportación) NO se replica: estaba
        `Visible=False` en el propio formulario y, aunque se activara,
        las dos ramas de su `If LaSuc = 0` arman exactamente el mismo
        SQL — nunca filtró nada, en ningún build. Confirmado con el
        usuario no migrarlo (2026-08-16)."""
        return (
            self.db.query(FcivaVta)
            .filter((FcivaVta.FECHA >= fecha_desde) & (FcivaVta.FECHA < fecha_hasta_exclusiva))
            .order_by(FcivaVta.FECHA, FcivaVta.CPBTE)
            .limit(limite)
            .all()
        )


class Fcestad1Repository(BaseRepository[Fcestad1]):
    """Repositorio específico para Detalles de Facturación."""

    def __init__(self, db: Session):
        super().__init__(db, Fcestad1)

    def by_comprobante(self, tipo: int, letra: str, ptovta: int, cpbte: int) -> List[Fcestad1]:
        """Retorna renglones de una factura."""
        return (
            self.db.query(Fcestad1)
            .filter(
                (Fcestad1.TIPO == tipo)
                & (Fcestad1.LETRA == letra)
                & (Fcestad1.PTOVTA == ptovta)
                & (Fcestad1.CPBTE == cpbte)
            )
            .all()
        )

    def by_articulo(self, cod1: str) -> List[Fcestad1]:
        """Retorna todos los renglones que contienen artículos de una Sección.

        AVISO: COD1 es la Sección, no la clave de un artículo puntual — esto
        devuelve movimientos de TODOS los artículos de esa sección. Para
        verificar movimientos de un artículo específico usar
        `by_cod1_cod2()`.
        """
        return self.db.query(Fcestad1).filter(func.trim(Fcestad1.COD1) == cod1.strip()).all()

    def by_cod1_cod2(self, cod1: str, cod2: str) -> List[Fcestad1]:
        """Retorna los renglones de facturación de un artículo puntual
        (clave real COD1+COD2, ver ArticuloRepository.by_cod1_cod2 y su
        nota de padding)."""
        return (
            self.db.query(Fcestad1)
            .filter(
                (func.trim(Fcestad1.COD1) == cod1.strip())
                & (func.trim(Fcestad1.COD2) == cod2.strip())
            )
            .all()
        )

    def by_seccion_y_rango(self, cod1: str, fecha_desde: date, fecha_hasta_exclusiva: date) -> List[Fcestad1]:
        """Renglones de una Sección en un rango de fechas `[desde, hasta)`
        — réplica de `ESTADIST.frm Sub DoVer3` ("Ventas de Artículos por
        Cliente"), reusado también por `VTAXART.frm` ("Ventas por
        Artículo"). El límite superior EXCLUSIVO ya viene resuelto por
        el llamador (`EstadisticaVentasService.fin_de_mes_exclusivo()`)."""
        return (
            self.db.query(Fcestad1)
            .filter(
                func.trim(Fcestad1.COD1) == cod1.strip(),
                Fcestad1.FECHA >= fecha_desde,
                Fcestad1.FECHA < fecha_hasta_exclusiva,
            )
            .order_by(Fcestad1.FECHA.asc(), Fcestad1.CPBTE.asc())
            .all()
        )

    def by_cod1_cod2_y_rango(self, cod1: str, cod2: str, fecha_desde: date, fecha_hasta_exclusiva: date) -> List[Fcestad1]:
        """Igual que `by_seccion_y_rango()` pero acotado a UN artículo
        puntual (COD1+COD2) — para `VTAXART.frm` cuando se filtra por
        Sección Y código, no toda la Sección."""
        return (
            self.db.query(Fcestad1)
            .filter(
                func.trim(Fcestad1.COD1) == cod1.strip(),
                func.trim(Fcestad1.COD2) == cod2.strip(),
                Fcestad1.FECHA >= fecha_desde,
                Fcestad1.FECHA < fecha_hasta_exclusiva,
            )
            .order_by(Fcestad1.FECHA.asc(), Fcestad1.CPBTE.asc())
            .all()
        )


class FctablasRepository(BaseRepository[Fctabla1]):
    """Repositorio específico para Tablas genéricas de Facturación."""

    def __init__(self, db: Session):
        super().__init__(db, Fctabla1)

    def by_ctab_cod(self, ctab: str, cod: str) -> Optional[Fctabla1]:
        """Busca entrada en tabla por CTAB y COD.

        Compara con `func.trim()` de ambos lados: `CTAB` en la base real
        (migrada desde Access) quedó con el padding fijo original del
        legacy (p.ej. `'CV   '`, 5 caracteres — ver Abmclte.frm Form_Load:
        `WHERE CTAB = 'CV   '`). Comparar sin trim rompe silenciosamente
        contra datos reales aunque los tests con SQLite en memoria (que
        insertan `'CV'` sin padding) pasen igual.
        """
        return (
            self.db.query(Fctabla1)
            .filter(
                (func.trim(Fctabla1.CTAB) == ctab.strip())
                & (func.trim(Fctabla1.COD) == cod.strip())
            )
            .first()
        )

    def by_ctab(self, ctab: str) -> List[Fctabla1]:
        """Retorna todas las entradas de una tabla específica (ver nota de
        padding en `by_ctab_cod`)."""
        return (
            self.db.query(Fctabla1)
            .filter(func.trim(Fctabla1.CTAB) == ctab.strip())
            .order_by(Fctabla1.COD)
            .all()
        )

    def secciones_que_usan_unidad(self, cod_um: str) -> List[Fctabla1]:
        """Secciones (CTAB='SC') que tienen esta Unidad de Medida asignada
        en cualquiera de sus 7 slots (ALF1-ALF7 — ver Acttabla.frm, bloques
        "Especificación"/"Cálculo de Precio"/"Unid. de Facturación").
        Usado para validar la baja de una Unidad de Medida: si alguna
        Sección la usa, borrarla rompería el armado de código de artículos
        o el cálculo de precio de esa Sección.
        """
        cod_um = cod_um.strip()
        return (
            self.db.query(Fctabla1)
            .filter(func.trim(Fctabla1.CTAB) == "SC")
            .filter(
                (func.trim(Fctabla1.ALF1) == cod_um)
                | (func.trim(Fctabla1.ALF2) == cod_um)
                | (func.trim(Fctabla1.ALF3) == cod_um)
                | (func.trim(Fctabla1.ALF4) == cod_um)
                | (func.trim(Fctabla1.ALF5) == cod_um)
                | (func.trim(Fctabla1.ALF6) == cod_um)
                | (func.trim(Fctabla1.ALF7) == cod_um)
            )
            .all()
        )


class ProveedRepository(BaseRepository[Proveed]):
    """Repositorio específico para Proveedores."""

    def __init__(self, db: Session):
        super().__init__(db, Proveed)

    def by_codigo(self, codigo: int) -> Optional[Proveed]:
        """Busca proveedor por CODIGO."""
        return self.db.query(Proveed).filter(Proveed.CODIGO == codigo).first()

    def by_rsocial(self, rsocial: str) -> List[Proveed]:
        """Busca proveedores por Razón Social (LIKE)."""
        return self.db.query(Proveed).filter(Proveed.RSOCIAL.ilike(f"%{rsocial}%")).all()

    def by_cuit(self, cuit: str) -> Optional[Proveed]:
        """Busca proveedor por CUIT."""
        return self.db.query(Proveed).filter(Proveed.CUIT == cuit).first()


class BancoRepository(BaseRepository[Banco]):
    """Repositorio específico para Bancos."""

    def __init__(self, db: Session):
        super().__init__(db, Banco)

    def by_cod(self, cod: int) -> Optional[Banco]:
        """Busca banco por COD."""
        return self.db.query(Banco).filter(Banco.COD == cod).first()

    def by_codigo_sucursal(self, cod: int, suc: int) -> Optional[Banco]:
        """Busca banco por COD y SUC."""
        return self.db.query(Banco).filter((Banco.COD == cod) & (Banco.SUC == suc)).first()

    def buscar_por_nombre(self, texto: str) -> List[Banco]:
        """Busca bancos cuyo NOMBRE contenga el texto (LIKE, sin
        distinguir mayúsculas) — usado por `BancoBusquedaDialog` cuando
        el operador no sabe el código y deja el campo en blanco en
        `PagoDialog` (no hay ABM de Bancos migrado todavía — `ABMBcos.frm`
        queda pendiente; esto es sólo un selector liviano)."""
        return self.db.query(Banco).filter(Banco.NOMBRE.ilike(f"%{texto}%")).order_by(Banco.NOMBRE).all()

    def todos(self) -> List[Banco]:
        """Todos los bancos, ordenados por Nombre — lista default de
        `BancoBusquedaDialog` sin filtro."""
        return self.db.query(Banco).order_by(Banco.NOMBRE).all()


class ChequeRepository(BaseRepository[Cheque]):
    """Repositorio específico para Cheques."""

    def __init__(self, db: Session):
        super().__init__(db, Cheque)

    def by_nrocheq(self, nrocheq: int) -> Optional[Cheque]:
        """Busca cheque por NROCHEQ."""
        return self.db.query(Cheque).filter(Cheque.NROCHEQ == nrocheq).first()

    # Cheques.ESTADO real (confirmado 2026-08-16 contra VerCheq.frm/EmiRec.frm,
    # no asumido): TEXTO "1"-"4", NO las palabras "COBRADO"/"RECHAZADO" que
    # tenían estos dos métodos antes (nunca se habían usado en la app —
    # `por_cobrar()`/`rechazados()` estaban muertos, sin ningún call site —
    # así que el valor equivocado nunca se manifestó; corregido acá antes
    # de que alguna pantalla nueva los use).
    #   1 = En Cartera (EmiRec.frm:1294, "1-En Cartera")
    #   2 = Egresado    (VerCheq.frm:1094, al completar tipegr/cpbegr/fecegr)
    #   3 = Rechazado, 4 = Otros (VerCheq.frm Option1(3)/Option1(4))
    ESTADO_EN_CARTERA = "1"
    ESTADO_EGRESADO = "2"
    ESTADO_RECHAZADO = "3"
    ESTADO_OTROS = "4"

    def por_cobrar(self) -> List[Cheque]:
        """Retorna cheques en cartera (ESTADO='1'), pendientes de depositar/cobrar."""
        return self.db.query(Cheque).filter(Cheque.ESTADO == self.ESTADO_EN_CARTERA).all()

    def rechazados(self) -> List[Cheque]:
        """Retorna cheques rechazados (ESTADO='3')."""
        return self.db.query(Cheque).filter(Cheque.ESTADO == self.ESTADO_RECHAZADO).all()

    def en_cartera_de_cliente(self, clte: int) -> List[Cheque]:
        """Cheques en cartera (ESTADO='1') de un cliente puntual — réplica
        de `CtaCte.frm Sub DoVer9` (líneas 2009-2056, "Ver Cheques"),
        ordenados por vencimiento igual que la consulta general de
        `VerCheq.frm` (`ORDER BY FECVTO, NROCHEQ`)."""
        return (
            self.db.query(Cheque)
            .filter(Cheque.CLTE == clte, Cheque.ESTADO == self.ESTADO_EN_CARTERA)
            .order_by(Cheque.FECVTO.asc(), Cheque.NROCHEQ.asc())
            .all()
        )

    def by_estado_desde(self, estado: str, fecha_desde: date, limite: int = 50) -> List[Cheque]:
        """Réplica de `VerCheq.frm Sub DoVer2` ("Consulta de Cheques"
        general, todos los clientes): cheques de un Estado con
        `FECING >= fecha_desde`, ordenados por vencimiento, hasta
        `limite` filas. El legacy calcula también un "hasta fin de mes"
        (`FECHAHst`) pero nunca lo usa en el `WHERE` — código muerto
        real, confirmado leyendo el `Sub` completo, no se replica."""
        return (
            self.db.query(Cheque)
            .filter(Cheque.ESTADO == estado, Cheque.FECING >= fecha_desde)
            .order_by(Cheque.FECVTO.asc(), Cheque.NROCHEQ.asc())
            .limit(limite)
            .all()
        )


class EfectivoRepository(BaseRepository[Efectivo]):
    """Repositorio específico para el pago en efectivo de un Recibo."""

    def __init__(self, db: Session):
        super().__init__(db, Efectivo)

    def by_comprobante(self, cpbte: int, tipo: int = 4) -> Optional[Efectivo]:
        """Busca la fila de efectivo de un Recibo puntual — réplica del
        `SELECT * FROM Efectivo WHERE Cpbte = ... AND Tipo = 4` de
        `EmiRec.frm Sub Graba()` (patrón upsert: actualiza si ya existe)."""
        return (
            self.db.query(Efectivo)
            .filter(Efectivo.CPBTE == cpbte, Efectivo.TIPO == tipo)
            .first()
        )


class MovimVSRepository(BaseRepository[MovimVS]):
    """Repositorio específico para Anticipos/Retenciones/Tarjetas/
    Transferencias/Baja Incobrable de un Recibo (ver docstring de
    `MovimVS` en models.py)."""

    def __init__(self, db: Session):
        super().__init__(db, MovimVS)

    def by_comprobante_y_tipreg(self, cpbte: int, tipreg: str) -> Optional[MovimVS]:
        """Réplica del `SELECT * FROM MovimVS WHERE CPBTE = ... AND
        TIPREG = '...'` de `EmiRec.frm Sub Graba()` (guard de idempotencia
        contra un reintento de grabación del mismo Recibo)."""
        return (
            self.db.query(MovimVS)
            .filter(MovimVS.CPBTE == cpbte, MovimVS.TIPREG == tipreg)
            .first()
        )


class CotizacionRepository(BaseRepository[Cotizacion]):
    """Repositorio específico para Cotizaciones.

    OJO: Cotizacion.FECHA es TEXT "dd/mm/yyyy" en fcmenu.mdb (confirmado
    vía DAO), no una fecha nativa — ver models.py. Ordenar/filtrar con
    SQL sobre ese texto da resultados cronológicamente incorrectos (ej.
    "01/12/2026" ordena ANTES que "02/01/2010" en ASCII, a pesar de ser
    16 años posterior — verificado con datos reales). Por eso acá se
    parsea a date en Python en vez de usar ORDER BY FECHA en SQL.
    """

    def __init__(self, db: Session):
        super().__init__(db, Cotizacion)

    @staticmethod
    def _parse_fecha(fecha_str: str) -> date:
        return datetime.strptime(fecha_str, "%d/%m/%Y").date()

    @staticmethod
    def _format_fecha(fecha: date) -> str:
        return fecha.strftime("%d/%m/%Y")

    def ultima(self) -> Optional[Cotizacion]:
        """Retorna la cotización más reciente (parseando FECHA como fecha real)."""
        filas = self.db.query(Cotizacion).all()
        if not filas:
            return None
        return max(filas, key=lambda c: self._parse_fecha(c.FECHA))

    def by_fecha(self, fecha: date) -> Optional[Cotizacion]:
        """Retorna cotización para una fecha específica (date de Python)."""
        return (
            self.db.query(Cotizacion)
            .filter(Cotizacion.FECHA == self._format_fecha(fecha))
            .first()
        )

    def guardar(self, fecha: date, valores: dict) -> Cotizacion:
        """Crea o actualiza la cotización de una fecha (upsert simple por
        FECHA, formateada dd/mm/yyyy como en los datos reales).

        Reemplaza el mecanismo de "fila ancla" 01/01/1900 de Cotizac.frm
        (que además tenía su propia búsqueda por fecha en formato
        yyyy/mm/dd, inconsistente con el formato real dd/mm/yyyy — ver
        docstring de la clase) — decisión confirmada con el usuario
        2026-08-15: la última cotización real se obtiene con `ultima()`,
        no hace falta mantener una fila especial con doble escritura.
        """
        existente = self.by_fecha(fecha)
        datos = {**valores, "FECHA": self._format_fecha(fecha)}
        if existente is None:
            return self.create(datos)
        return self.update(existente.id, datos)


class RepositoryFactory:
    """Factory para crear repositorios sin code duplication."""

    def __init__(self, db: Session):
        self.db = db
        self._repos = {}

    def cliente(self) -> ClienteRepository:
        if "cliente" not in self._repos:
            self._repos["cliente"] = ClienteRepository(self.db)
        return self._repos["cliente"]

    def notaclte(self) -> NotaClteRepository:
        if "notaclte" not in self._repos:
            self._repos["notaclte"] = NotaClteRepository(self.db)
        return self._repos["notaclte"]

    def dtoxclte(self) -> DtoxClteRepository:
        if "dtoxclte" not in self._repos:
            self._repos["dtoxclte"] = DtoxClteRepository(self.db)
        return self._repos["dtoxclte"]

    def articulo(self) -> ArticuloRepository:
        if "articulo" not in self._repos:
            self._repos["articulo"] = ArticuloRepository(self.db)
        return self._repos["articulo"]

    def notaarticulo(self) -> NotaArticuloRepository:
        if "notaarticulo" not in self._repos:
            self._repos["notaarticulo"] = NotaArticuloRepository(self.db)
        return self._repos["notaarticulo"]

    def stock(self) -> StockRepository:
        if "stock" not in self._repos:
            self._repos["stock"] = StockRepository(self.db)
        return self._repos["stock"]

    def despacho(self) -> DespachoRepository:
        if "despacho" not in self._repos:
            self._repos["despacho"] = DespachoRepository(self.db)
        return self._repos["despacho"]

    def movstock(self) -> MovStockRepository:
        if "movstock" not in self._repos:
            self._repos["movstock"] = MovStockRepository(self.db)
        return self._repos["movstock"]

    def totales(self) -> TotalesRepository:
        if "totales" not in self._repos:
            self._repos["totales"] = TotalesRepository(self.db)
        return self._repos["totales"]

    def parametro(self) -> ParametroRepository:
        if "parametro" not in self._repos:
            self._repos["parametro"] = ParametroRepository(self.db)
        return self._repos["parametro"]

    def ctascte(self) -> CtascteRepository:
        if "ctascte" not in self._repos:
            self._repos["ctascte"] = CtascteRepository(self.db)
        return self._repos["ctascte"]

    def imputacion(self) -> ImputacionRepository:
        if "imputacion" not in self._repos:
            self._repos["imputacion"] = ImputacionRepository(self.db)
        return self._repos["imputacion"]

    def fciva_vta(self) -> FcivaVtaRepository:
        if "fciva_vta" not in self._repos:
            self._repos["fciva_vta"] = FcivaVtaRepository(self.db)
        return self._repos["fciva_vta"]

    def fcestad1(self) -> Fcestad1Repository:
        if "fcestad1" not in self._repos:
            self._repos["fcestad1"] = Fcestad1Repository(self.db)
        return self._repos["fcestad1"]

    def fctablas(self) -> FctablasRepository:
        if "fctablas" not in self._repos:
            self._repos["fctablas"] = FctablasRepository(self.db)
        return self._repos["fctablas"]

    def proveed(self) -> ProveedRepository:
        if "proveed" not in self._repos:
            self._repos["proveed"] = ProveedRepository(self.db)
        return self._repos["proveed"]

    def banco(self) -> BancoRepository:
        if "banco" not in self._repos:
            self._repos["banco"] = BancoRepository(self.db)
        return self._repos["banco"]

    def cheque(self) -> ChequeRepository:
        if "cheque" not in self._repos:
            self._repos["cheque"] = ChequeRepository(self.db)
        return self._repos["cheque"]

    def efectivo(self) -> EfectivoRepository:
        if "efectivo" not in self._repos:
            self._repos["efectivo"] = EfectivoRepository(self.db)
        return self._repos["efectivo"]

    def movimvs(self) -> MovimVSRepository:
        if "movimvs" not in self._repos:
            self._repos["movimvs"] = MovimVSRepository(self.db)
        return self._repos["movimvs"]

    def cotizacion(self) -> CotizacionRepository:
        if "cotizacion" not in self._repos:
            self._repos["cotizacion"] = CotizacionRepository(self.db)
        return self._repos["cotizacion"]


__all__ = [
    "BaseRepository",
    "ETIQUETAS_TIPO_CTASCTE",
    "ClienteRepository",
    "ArticuloRepository",
    "StockRepository",
    "DespachoRepository",
    "ParametroRepository",
    "CtascteRepository",
    "ImputacionRepository",
    "FcivaVtaRepository",
    "Fcestad1Repository",
    "FctablasRepository",
    "ProveedRepository",
    "BancoRepository",
    "ChequeRepository",
    "EfectivoRepository",
    "MovimVSRepository",
    "CotizacionRepository",
    "RepositoryFactory",
]
