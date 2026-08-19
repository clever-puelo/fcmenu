"""Capa de Servicios (lógica de negocio) sobre los repositorios de FCMENU.

Toda la lógica acá documentada fue extraída y confirmada contra el código
VB6 original (no asumida) siguiendo la Regla de Oro #2 del CLAUDE.md del
proyecto. Cada método cita el/los .frm de origen. Las decisiones de negocio
que requerían criterio del usuario (fórmula de IVA, fuente de verdad de
stock, fuente de verdad de deuda) fueron validadas explícitamente antes de
escribir este archivo — ver `data_layer_progress.md` en memoria de sesión.

Todos los importes usan `decimal.Decimal` (prohibido `float` para dinero,
por regla del proyecto).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from typing import Optional

from sqlalchemy.orm import Session

from .models import (
    Articulo,
    Cheque,
    Cliente,
    Ctascte,
    Despacho,
    Efectivo,
    Fcestad1,
    FcivaVta,
    Imputacion,
    MovimVS,
    MovStock,
    Parametro,
    Stock,
    Totales,
)
from .decimals import format_decimal
from .provincias import nombre_provincia
from .repository import ETIQUETAS_TIPO_CTASCTE, ETIQUETAS_TIPO_FCIVAVTA, RepositoryFactory

TWO_PLACES = Decimal("0.01")


def _round2(value: Decimal) -> Decimal:
    """Redondeo a 2 decimales, half-up (equivalente al Round() de VB6)."""
    return Decimal(value).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)


def _val_vb6(texto: Optional[str]) -> int:
    """Réplica de la función Val() de VB6: toma los dígitos (con signo
    opcional) desde el principio de la cadena, 0 si no hay ninguno.
    Usada por el filtro DESDE/HASTA de `ModPrec.frm` (`ArticuloService.
    candidatos_modificacion_precio`) y por `VTAXART.frm` (`Estadistica
    VentasService.ventas_articulo_agrupadas`) — mismo patrón "Val(COD2)
    como un único número" en ambas pantallas legacy."""
    texto = (texto or "").strip()
    m = re.match(r"-?\d+", texto)
    return int(m.group()) if m else 0


# ---------------------------------------------------------------------------
# FacturaService
# ---------------------------------------------------------------------------


@dataclass
class TotalFactura:
    """Desglose de una factura, según la fórmula confirmada con el usuario.

    Fuente: EmiFact.frm:1159-1181 (impresión) y EmiFact.frm:2641-2644 (AFIP),
    con la diferencia entre ambas resuelta: TotIVANI/TotSIVA (rama
    ClteCIVA=2, "Responsable No Inscripto") es código legacy muerto —
    categoría de IVA eliminada por AFIP en 2003 — y se excluye del cálculo.
    """

    bruto: Decimal
    descuento: Decimal
    neto_gravado: Decimal
    iva: Decimal
    percepcion_iibb: Decimal
    total: Decimal


class FacturaService:
    """Cálculo de totales de factura: IVA, bonificaciones, percepciones.

    No incluye la emisión/persistencia del comprobante ni la integración
    AFIP (WSFEv1) — eso es Fase 2 (Business Logic Services / PyAfipWs).
    """

    # ClteCIVA < 3 son las únicas condiciones de IVA sobre las que el legacy
    # calcula IVA/percepción IIBB (CabFact.frm Combo1: 1=Insc., 2=No Insc.).
    # EmiFact.frm:1162-1167 aplica TotIVAIns a AMBAS (civa 1 y 2) por igual;
    # solo civa=2 calculaba ADEMÁS TotIVANI, que es la parte muerta que se
    # excluye acá (ver TotalFactura) — la alícuota de IVAIns en sí sigue
    # aplicando igual para civa=1 y civa=2.
    CIVA_GRAVADO_MAX = 3  # ClteCIVA < 3 → sujeto a IVA/percepción en el legacy

    def __init__(self, db: Session):
        self.repos = RepositoryFactory(db)

    def alicuota_iva_inscripto(self) -> Decimal:
        """Alícuota de IVA Inscripto, configurable (no hardcodeada).

        Fuente: FCMENU.bas:333-343 — se lee una vez de Parametro.IVAINS
        (fila única CLAVE='1') y se usa como fracción (IVAINS/100).
        """
        config = self.repos.parametro().get_config()
        if config is None or config.IVAINS is None:
            raise ValueError(
                "Parametro (CLAVE='1').IVAINS no está configurado — "
                "el legacy directamente aborta el arranque en este caso "
                "(FCMENU.bas:337-339)."
            )
        return Decimal(config.IVAINS) / Decimal(100)

    def percepcion_iibb_habilitada(self) -> bool:
        """Replica `ConPercep <> 1` de EmiFact.frm:1177 (Parametro.MCAIB).

        MCAIB nulo → ConPercep = 2 (deshabilitada), igual que FCMENU.bas:350-354.
        """
        config = self.repos.parametro().get_config()
        mcaib = config.MCAIB if config and config.MCAIB is not None else 2
        return mcaib == 1

    def calcular_bonificacion_cascada(
        self, importe: Decimal, porcentajes: list[Decimal]
    ) -> Decimal:
        """Descuento en cascada (hasta 5 niveles), sobre el saldo ya descontado.

        Fuente: EmiFact.frm:1821-1827 (DxClte + bucle For i2 = 1 To 5):
            SumDtos = SumDtos + ((Impte - SumDtos) * (Dto_i / 100))
        Cada porcentaje se aplica sobre lo que queda del importe original
        tras los descuentos previos (no son porcentajes que se suman entre sí).
        """
        importe = Decimal(importe)
        acumulado = Decimal("0")
        for pct in porcentajes[:5]:
            pct = Decimal(pct)
            if pct > 0:
                acumulado += (importe - acumulado) * (pct / Decimal(100))
        return acumulado

    def letra_comprobante(self, civa_cliente: int, provincia: Optional[str]) -> str:
        """Letra de Factura (A o B) — fórmula confirmada con el usuario
        en la fase del Facturador, fuente `CabFact.frm:709-713`:

            If ClteCIVA < 3 Or (ClteCIVA = 4 And CltePCIA = "V ") Then
                LaLetra = "A"
            Else
                LaLetra = "B"

        `provincia` se compara sin el padding fijo de Access (confirmado
        contra datos reales: `Cliente.PCIA` viene como `'V '`, 2 chars,
        mismo patrón que el resto de las columnas TEXT migradas) —
        acá se compara ya con `.strip()` de ambos lados, no hace falta
        que el llamador lo pre-procese."""
        provincia_normalizada = (provincia or "").strip()
        if civa_cliente < self.CIVA_GRAVADO_MAX or (civa_cliente == 4 and provincia_normalizada == "V"):
            return "A"
        return "B"

    def calcular_total(
        self,
        bruto: Decimal,
        descuento: Decimal,
        civa_cliente: int,
        porcentaje_iibb: Decimal = Decimal("0"),
    ) -> TotalFactura:
        """Calcula el total de una factura con la fórmula confirmada.

        bruto: suma de los ítems (antes de descuento global).
        descuento: total de descuentos/bonificaciones ya calculado.
        civa_cliente: Clientes.CIVA del cliente facturado.
        porcentaje_iibb: alícuota de percepción de Ingresos Brutos del cliente
            (Text2.Text en CabFact.frm), en unidades de porcentaje (ej. 3.5).
        """
        bruto = Decimal(bruto)
        descuento = Decimal(descuento)
        neto_gravado = bruto - descuento

        gravado = civa_cliente < self.CIVA_GRAVADO_MAX
        con_percepcion = self.percepcion_iibb_habilitada()

        iva = Decimal("0")
        percepcion = Decimal("0")
        if gravado:
            iva = _round2(neto_gravado * self.alicuota_iva_inscripto())
            if con_percepcion:
                percepcion = _round2(neto_gravado * (Decimal(porcentaje_iibb) / Decimal(100)))

        total = _round2(neto_gravado) + iva + percepcion

        return TotalFactura(
            bruto=_round2(bruto),
            descuento=_round2(descuento),
            neto_gravado=_round2(neto_gravado),
            iva=iva,
            percepcion_iibb=percepcion,
            total=total,
        )


# ---------------------------------------------------------------------------
# CuentaCorrienteService
# ---------------------------------------------------------------------------


@dataclass
class ImputacionResultado:
    comprobante_id: int
    slot_usado: str  # "IMPUT1".."IMPUT6"
    debe_anterior: Decimal
    debe_nuevo: Decimal
    imputacion_id: int
    descuento_imputacion_id: Optional[int] = None


@dataclass
class FilaExtracto:
    """Una fila del extracto anual — réplica de `CtaCte.frm CargaGrilla()`
    (líneas 1898-2003), incluye el saldo corriente ya acumulado columna
    por columna igual que la grilla original."""

    fecha: date
    tipo: int
    tipo_label: str
    cpbte: Optional[int]
    prefijo: Optional[int]
    letra: Optional[str]
    imput1: str
    imput2: str
    imput3: str
    imput4: str
    resto: Optional[Decimal]  # "RESTO": DEBE puntual pendiente, sólo TIPO 1/3
    debe: Decimal
    haber: Decimal
    saldo: Decimal
    comprobante_id: Optional[int] = None  # Ctascte.id, para drill-down


@dataclass
class ExtractoAnual:
    anio: int
    saldo_inicial_debe: Decimal
    saldo_inicial_credito: Decimal
    filas: list[FilaExtracto]
    total_debe: Decimal
    total_haber: Decimal
    saldo_final: Decimal


@dataclass
class ResumenCliente:
    """Datos del panel lateral de la Consulta de Cta.Cte — 5 de los 6
    campos que `CtaCte.frm` declaraba pero nunca poblaba (confirmado con
    grep de todo el proyecto: ningún `.frm`/`.bas` les asigna valor).
    "Contacto" se descartó (sin fuente de datos real, mismo criterio que
    los controles fantasma de `Abmclte.frm`); decisión confirmada con el
    usuario 2026-08-16."""

    telefono: str
    fecha_alta: Optional[date]
    valores_pendientes: Decimal  # total de cheques en cartera del cliente
    ultima_venta: Optional[date]
    saldo_vencido: Decimal


@dataclass
class ComprobanteZona:
    """Un renglón de `CobranzasZona` — Factura o ND impaga de un cliente
    de la Zona consultada."""

    tipo_label: str  # "Fact." o "N/D."
    letra: str
    comprobante: str  # "01-000123" (Prefijo-Cpbte, mismo padding del legacy)
    fecha: Optional[date]
    fecvto: Optional[date]
    importe: Decimal
    debe: Decimal
    dias_vencido: int  # 0 si no está vencido
    vencido: bool


@dataclass
class ClienteZonaResumen:
    cliente_codigo: int
    cliente_nombre: str
    comprobantes: list[ComprobanteZona]
    subtotal_importe: Decimal
    subtotal_debe: Decimal


@dataclass
class CobranzasZona:
    zona: int
    clientes: list[ClienteZonaResumen]
    total_deuda: Decimal
    total_vencido: Decimal


class CuentaCorrienteService:
    """Saldo e imputación de pagos sobre Ctasctes.

    Fuente: CtaCte.frm (saldo), EmiRec.frm:1038-1161 (imputación al confirmar
    un recibo). Ver TIPOS_DEBE en CtascteRepository para el detalle de signos.
    """

    # Máximo de imputaciones parciales por comprobante — límite legacy
    # (IMPUT1..IMPUT6 en Ctasctes) que el usuario confirmó mantener tal cual,
    # no liberar, por compatibilidad con reportes existentes.
    MAX_IMPUTACIONES = 6
    SLOTS_IMPUTACION = [f"IMPUT{i}" for i in range(1, MAX_IMPUTACIONES + 1)]

    # TIPO que suman al DEBE pendiente de un comprobante en vez de restar
    # (EmiRec.frm:1094-1098): NC, Pago a Cuenta, NC por Interés.
    TIPOS_QUE_SUMAN_AL_DEBE = {2, 5, 8}

    def __init__(self, db: Session):
        self.repos = RepositoryFactory(db)

    def saldo_cliente(self, clte: int) -> Decimal:
        """Saldo de cuenta corriente (deudor > 0), recalculado desde el ledger.

        Se recalcula siempre desde Ctasctes en vez de leer Clientes.DEUDA
        (campo cacheado que se mantiene en paralelo en EmiFact.frm/EmiRec.frm
        y puede desincronizarse) — decisión confirmada con el usuario.
        """
        resultado = self.repos.ctascte().deuda_cliente(clte)
        return resultado["saldo"]

    def facturas_pendientes(self, clte: int) -> list:
        """Facturas/ND (TIPO 1 o 3) con saldo pendiente (DEBE > 0) para un cliente.

        Ctascte.DEBE es el saldo puntual de ESE comprobante (no confundir con
        el saldo general del cliente) — ver EmiRec.frm:1094-1098.
        """
        movimientos = self.repos.ctascte().by_cliente(clte)
        return [
            m
            for m in movimientos
            if m.TIPO in (1, 3) and (m.DEBE or Decimal("0")) > Decimal("0")
        ]

    def tiene_deuda_vieja(self, clte: int, hoy: Optional[date] = None) -> bool:
        """"Aviso de Deuda" del Facturador (`ConDeuda.frm`, disparado desde
        `BusClte.frm Function TieneDeuda()`, líneas 1499-1547): True si la
        Factura impaga más vieja del cliente (`DEBE>1`) tiene más de
        `Parametro.NUME20` ("Días Vto.") días desde su FECHA de EMISIÓN
        (no `FECVTO`/vencimiento real — confirmado con el usuario
        2026-08-16 replicar tal cual, aunque mide antigüedad de emisión,
        no días de atraso real).

        Si no hay ninguna Factura en esas condiciones, o si `Parametro`/
        `NUME20` no están configurados, devuelve `False` (el legacy en
        ese caso ni siquiera llega a arrancar — `FCMENU.bas` corta la
        app entera si `Parametro` no existe — acá se prefiere un default
        seguro de "no avisar" antes que romper)."""
        factura = self.repos.ctascte().factura_mas_antigua_impaga(clte)
        if factura is None or factura.FECHA is None:
            return False

        config = self.repos.parametro().get_config()
        if config is None or config.NUME20 is None:
            return False

        hoy = hoy or date.today()
        dias = (hoy - factura.FECHA).days
        return dias >= config.NUME20

    # "Fact. "/"N/D.  " — réplica exacta de VerCobra.frm:357-360 (distinta
    # de ETIQUETAS_TIPO_CTASCTE, que usa "NDéb." para el mismo TIPO=3 en
    # otra pantalla — cada consulta legacy tiene su propia etiqueta, no
    # hay una única fuente de verdad de texto para reusar).
    ETIQUETAS_COBRANZA = {1: "Fact.", 3: "N/D."}

    def cobranzas_por_zona(self, zona: int, hoy: Optional[date] = None) -> CobranzasZona:
        """"Cobranzas por Zona" (`VerCobra.frm Sub DoVer1`/`CargaGrilla`):
        Facturas/ND impagas (`DEBE>1`) de todos los clientes de una Zona,
        agrupadas por cliente con subtotales — réplica de la grilla
        agrupada/con outline del legacy (acá se arma la jerarquía en
        Python; la UI decide cómo mostrarla, ej. `QTreeWidget`).

        `dias_vencido`/`vencido` se calculan contra `FECVTO` (fecha de
        vencimiento real) — a diferencia de `tiene_deuda_vieja()` (que
        mide la antigüedad de EMISIÓN), acá el legacy sí compara contra
        el vencimiento real (`VerCobra.frm:370`), no hay contradicción:
        son dos pantallas distintas con criterios distintos, ambos
        fieles a su fuente."""
        hoy = hoy or date.today()
        clientes = sorted(self.repos.cliente().by_zona(zona), key=lambda c: (c.NOMB or ""))

        resumenes: list[ClienteZonaResumen] = []
        total_deuda = Decimal("0")
        total_vencido = Decimal("0")

        for cliente in clientes:
            movimientos = self.repos.ctascte().facturas_nd_impagas(cliente.CODIGO)
            if not movimientos:
                continue

            comprobantes = []
            subtotal_importe = Decimal("0")
            subtotal_debe = Decimal("0")
            for m in movimientos:
                debe = m.DEBE or Decimal("0")
                importe = m.IMPTE or Decimal("0")
                dias_vencido = 0
                vencido = False
                if m.FECVTO is not None and m.FECVTO < hoy:
                    dias_vencido = (hoy - m.FECVTO).days
                    vencido = True
                    total_vencido += debe

                comprobante_fmt = f"{(m.PREFIJO or 0):02d}-{(m.CPBTE or 0):06d}"
                comprobantes.append(
                    ComprobanteZona(
                        tipo_label=self.ETIQUETAS_COBRANZA.get(m.TIPO, "?"),
                        letra=(m.LETRA or "").strip(),
                        comprobante=comprobante_fmt,
                        fecha=m.FECHA,
                        fecvto=m.FECVTO,
                        importe=importe,
                        debe=debe,
                        dias_vencido=dias_vencido,
                        vencido=vencido,
                    )
                )
                subtotal_importe += importe
                subtotal_debe += debe
                total_deuda += debe

            resumenes.append(
                ClienteZonaResumen(
                    cliente_codigo=cliente.CODIGO,
                    cliente_nombre=(cliente.NOMB or "").strip(),
                    comprobantes=comprobantes,
                    subtotal_importe=subtotal_importe,
                    subtotal_debe=subtotal_debe,
                )
            )

        return CobranzasZona(zona=zona, clientes=resumenes, total_deuda=total_deuda, total_vencido=total_vencido)

    # TIPO de comprobante que genera Imputacion.TIPO — literal fijo "4"
    # (recibo) / "6" (descuento por pronto pago), texto en Imputacion aunque
    # numérico en Ctasctes (EmiRec.frm:1119 y :1140). "2" (Nota de Crédito)
    # agregado para `EmisionNotaCreditoService` — réplica de `EmiFact.frm
    # Sub ImputaFact()` línea 2422 (`RgCCTE!TIPO = 2`), mismo mecanismo de
    # slots que el Recibo, sólo cambia el literal grabado en Imputacion.TIPO.
    TIPO_IMPUTACION_RECIBO = "4"
    TIPO_IMPUTACION_DESCUENTO = "6"
    TIPO_IMPUTACION_NOTA_CREDITO = "2"

    def imputar_pago(
        self,
        comprobante: Ctascte,
        importe_aplicado: Decimal,
        corr: int,
        cpbte_recibo: int,
        usuario: str,
        clte: Optional[int] = None,
        fecha: Optional[date] = None,
        descuento: Decimal = Decimal("0"),
        commit: bool = True,
        tipo_imputacion: str = TIPO_IMPUTACION_RECIBO,
    ) -> ImputacionResultado:
        """Aplica un pago/crédito parcial o total a un comprobante
        (factura/ND) de CtasCtes y deja el registro histórico en
        Imputacion — flujo completo end-to-end.

        Replica EmiRec.frm:1038-1157 (Recibo, `tipo_imputacion` default
        "4") y, con `tipo_imputacion=TIPO_IMPUTACION_NOTA_CREDITO`,
        `EmiFact.frm Sub ImputaFact()` líneas 2351-2437 (Nota de
        Crédito) — mismo mecanismo exacto de slots IMPUT1-6 y de
        actualización de DEBE en ambos casos, sólo cambia qué literal se
        graba en `Imputacion.TIPO` y de dónde sale `cpbte_recibo`/`corr`
        (el propio comprobante que se está emitiendo, no un Recibo):
          1. Busca el primer slot libre entre IMPUT1..IMPUT6 (vacío en VB6
             era "  "/"00"/"0 "; acá se trata None/"" como libre).
          2. Le asigna el correlativo `corr` del comprobante que imputa
             (Clientes.CORR1, ciclo 1-99, gestionado fuera de este método).
          3. Actualiza DEBE: resta el importe aplicado, salvo TIPO en
             TIPOS_QUE_SUMAN_AL_DEBE (NC/Pago a Cta/NC Interés) que suma.
          4. Inserta un registro en Imputacion (TIPO=`tipo_imputacion`) con
             la referencia al comprobante original (TIPOI/CPBTEI/FECHAI/
             FECVTOI).
          5. Si `descuento` > 0, inserta un segundo registro (TIPO="6") —
             sólo aplica al camino de Recibo, la NC nunca pasa `descuento`.

        Lanza ValueError si ya no quedan slots libres (7ma aplicación
        parcial), caso que el VB6 tampoco maneja explícitamente.

        `commit=False` (usado por `EmisionReciboService`/
        `EmisionNotaCreditoService`, que llaman esto dentro de UNA sola
        transacción propia): deja los cambios sólo en la sesión —
        `flush()` para que `Imputacion.id` quede disponible, sin
        confirmar — y el llamador hace un único `commit()`/`rollback()`
        al final. Con el default `commit=True` el método es atómico por
        sí solo, igual que antes.
        """
        importe_aplicado = Decimal(importe_aplicado)
        descuento = Decimal(descuento)
        fecha = fecha or date.today()
        clte = clte if clte is not None else comprobante.CLTE

        slot_libre = None
        for slot in self.SLOTS_IMPUTACION:
            valor = getattr(comprobante, slot)
            if valor in (None, "", "  ", "00", "0 "):
                slot_libre = slot
                break

        if slot_libre is None:
            raise ValueError(
                f"Comprobante {comprobante.CPBTE} ya tiene las "
                f"{self.MAX_IMPUTACIONES} imputaciones parciales completas "
                "(límite legacy IMPUT1..IMPUT6, mantenido por decisión del usuario)."
            )

        # IMPUT1..6 son TEXT(2) reales en Access (confirmado vía DAO) — se
        # castea explícito a str en vez de confiar en la coerción implícita
        # de SQLite, que no es portable a PostgreSQL (VARCHAR no acepta un
        # int como bind param sin adaptador explícito).
        setattr(comprobante, slot_libre, str(corr))

        debe_anterior = comprobante.DEBE or Decimal("0")
        if comprobante.TIPO in self.TIPOS_QUE_SUMAN_AL_DEBE:
            comprobante.DEBE = debe_anterior + importe_aplicado
        else:
            comprobante.DEBE = debe_anterior - importe_aplicado

        # `self.db.add(...)` directo en vez de `ImputacionRepository.create()`
        # a propósito: `create()` hace su propio `commit()` por llamada, lo
        # que rompería la atomicidad de `EmisionReciboService` (que llama
        # esto una vez por cada comprobante aplicado dentro de UN solo
        # Recibo). El commit real queda al final de este método (o del
        # llamador, si `commit=False`).
        registro = Imputacion(
            CLTE=clte,
            FECHA=fecha,
            TIPO=tipo_imputacion,
            CPBTE=cpbte_recibo,
            IMPTE=importe_aplicado,
            TIPOI=str(comprobante.TIPO),
            CPBTEI=comprobante.CPBTE,
            CORR=corr,
            FECHAI=comprobante.FECHA,
            FECVTOI=comprobante.FECVTO,
            USUARIO=usuario[:6] if usuario else None,
        )
        self.repos.db.add(registro)

        descuento_id = None
        if descuento > 0:
            registro_dto = Imputacion(
                CLTE=clte,
                FECHA=fecha,
                TIPO=self.TIPO_IMPUTACION_DESCUENTO,
                CPBTE=cpbte_recibo,
                IMPTE=descuento,
                TIPOI=str(comprobante.TIPO),
                CPBTEI=comprobante.CPBTE,
                CORR=corr,
                FECHAI=comprobante.FECHA,
                FECVTOI=comprobante.FECVTO,
                USUARIO=usuario[:6] if usuario else None,
            )
            self.repos.db.add(registro_dto)

        # flush (no commit) alcanza para que `registro`/`registro_dto`
        # tengan `.id` asignado por Postgres — necesario para el resultado
        # incluso cuando `commit=False` deja la confirmación real en manos
        # del llamador (`EmisionReciboService`).
        self.repos.db.flush()
        if descuento > 0:
            descuento_id = registro_dto.id

        if commit:
            self.repos.db.commit()

        return ImputacionResultado(
            comprobante_id=comprobante.id,
            slot_usado=slot_libre,
            debe_anterior=debe_anterior,
            debe_nuevo=comprobante.DEBE,
            imputacion_id=registro.id,
            descuento_imputacion_id=descuento_id,
        )

    # ------------------------------------------------------------------
    # Consulta (CtaCte.frm)
    # ------------------------------------------------------------------

    def extracto_anual(self, clte: int, anio: int) -> ExtractoAnual:
        """Extracto de cuenta corriente de un año en adelante — réplica
        de `CtaCte.frm Sub DoVer3()`+`CargaGrilla()` (líneas 1816-2003),
        saldo corriente incluido.

        El "Saldo Inicial" (línea `Sdo.Ant.`) usa `CtascteRepository.
        saldo_inicial()`, que a su vez usa la MISMA fórmula real del
        saldo corriente (`TIPOS_DEBE`) — así el saldo final de cualquier
        año elegido siempre reconcilia con el de "mostrar todo" (ver
        docstring de `saldo_inicial()` para el porqué: el legacy usaba
        acá una fórmula distinta e incompleta, confirmada a corregir por
        el usuario tras verla fallar con datos reales, 2026-08-16)."""
        ctascte_repo = self.repos.ctascte()
        inicial = ctascte_repo.saldo_inicial(clte, anio)
        saldo = inicial["debe"] - inicial["credito"]

        filas = [
            FilaExtracto(
                fecha=date(anio - 1, 12, 31),
                tipo=0,
                tipo_label="Sdo.Ant.",
                cpbte=None,
                prefijo=None,
                letra=None,
                imput1="", imput2="", imput3="", imput4="",
                resto=None,
                debe=inicial["debe"],
                haber=inicial["credito"],
                saldo=saldo,
            )
        ]

        debe_total = inicial["debe"]
        haber_total = inicial["credito"]

        for m in ctascte_repo.extracto(clte, anio):
            importe = m.IMPTE or Decimal("0")
            if m.TIPO in ctascte_repo.TIPOS_DEBE:
                debe, haber = importe, Decimal("0")
                saldo += importe
            else:
                debe, haber = Decimal("0"), importe
                saldo -= importe

            resto = None
            if m.TIPO in (1, 3):
                resto = m.DEBE or Decimal("0")

            debe_total += debe
            haber_total += haber

            filas.append(
                FilaExtracto(
                    fecha=m.FECHA,
                    tipo=m.TIPO,
                    tipo_label=ETIQUETAS_TIPO_CTASCTE.get(m.TIPO, "?"),
                    cpbte=m.CPBTE,
                    prefijo=m.PREFIJO,
                    letra=m.LETRA,
                    imput1=(m.IMPUT1 or "").strip(),
                    imput2=(m.IMPUT2 or "").strip(),
                    imput3=(m.IMPUT3 or "").strip(),
                    imput4=(m.IMPUT4 or "").strip(),
                    resto=resto,
                    debe=debe,
                    haber=haber,
                    saldo=saldo,
                    comprobante_id=m.id,
                )
            )

        return ExtractoAnual(
            anio=anio,
            saldo_inicial_debe=inicial["debe"],
            saldo_inicial_credito=inicial["credito"],
            filas=filas,
            total_debe=debe_total,
            total_haber=haber_total,
            saldo_final=saldo,
        )

    def resumen_cliente(self, cliente: Cliente) -> ResumenCliente:
        """Panel lateral de la Consulta — ver `ResumenCliente`."""
        cheques_en_cartera = self.repos.cheque().en_cartera_de_cliente(cliente.CODIGO)
        valores_pendientes = sum((c.IMPORTE or Decimal("0") for c in cheques_en_cartera), Decimal("0"))
        return ResumenCliente(
            telefono=(cliente.TEL1 or "").strip(),
            fecha_alta=cliente.FALTA,
            valores_pendientes=valores_pendientes,
            ultima_venta=self.repos.fciva_vta().ultima_venta(cliente.CODIGO),
            saldo_vencido=self.repos.ctascte().saldo_vencido(cliente.CODIGO),
        )

    def saldos_todos_clientes(self) -> list[dict]:
        """Saldo de todos los clientes — réplica de `CtaCte.frm DoVer8`
        ("Ver Saldos"), ver `CtascteRepository.saldos_todos_clientes()`."""
        return self.repos.ctascte().saldos_todos_clientes()


# ---------------------------------------------------------------------------
# StockService
# ---------------------------------------------------------------------------


class StockService:
    """Consultas y movimientos sobre la tabla Stock (STUNID como fuente de verdad).

    Confirmado con el usuario: Stock.STUNID es el campo que la venta
    realmente actualiza (EmiFact.frm:1864-1882); EST1/EST2/DEP1/DEP2 son un
    esquema multi-depósito legacy que la facturación no toca, y
    Articulo.STOCK es un campo de edición manual (AbmArt.frm) tampoco
    sincronizado por la venta. Ninguno de los dos se usa acá como fuente
    de verdad.
    """

    def __init__(self, db: Session):
        self.repos = RepositoryFactory(db)

    def articulos_criticos(self) -> list:
        """Items con STUNID < STMIN (StockRepository.criticos(), ya corregido)."""
        return self.repos.stock().criticos()

    def a_reponer(self) -> list:
        """Items por debajo del punto de pedido (STUNID < STREP).

        STREP = "Stock de Reposición" (tooltip en AbmArt.frm/Abmitem.frm):
        punto de pedido, distinto del mínimo absoluto (STMIN).
        """
        stock_repo = self.repos.stock()
        return [
            s
            for s in stock_repo.read_all(limit=10_000)
            if s.STREP is not None and (s.STUNID or Decimal("0")) < s.STREP
        ]

    def registrar_movimiento(
        self, cod1: str, cod2: str, cantidad: Decimal, es_salida: bool
    ) -> Stock:
        """Ajusta STUNID y ENTMES/SALMES para un artículo.

        Replica la rama de "unidades" de EmiFact.frm:1864-1872 (la única que
        se usa acá; la rama paralela de "metros" en EmiFact.frm:1874-1882
        tiene un bug de copy-paste confirmado — usa RgSTOK!Mtr como base en
        vez de RgSTOK!STUnid, pisando el descuento ya hecho — por eso NO se
        replica esa rama tal cual. Si el negocio realmente vende por metros
        además de por unidad, avisar para modelarlo como un acumulador
        independiente en vez de compartir STUNID.

        **Bug propio corregido (2026-08-16, al leer `Stock.frm Sub Graba()`
        completo para la fase de Stock)**: la rama de Entrada restaba de
        `SALMES` en vez de sumar a `ENTMES` — nunca se había notado porque
        este método no tenía ningún call site real todavía (sólo sus
        propios tests, que afirmaban `SALMES == -20` para una Entrada, un
        valor sin sentido de negocio que sólo reflejaba el bug). Corregido
        contra la fuente real (`Stock.frm:1141-1147`): Entrada suma a
        `ENTMES`, Salida suma a `SALMES` — nunca se tocan entre sí.
        """
        stock_repo = self.repos.stock()
        item = stock_repo.by_cod1(cod1)
        if item is None or item.COD2 != cod2:
            # Búsqueda exacta por COD1+COD2, como en EmiFact.frm:1835-1837.
            item = next(
                (
                    s
                    for s in stock_repo.read_all(limit=10_000)
                    if s.COD1 == cod1 and s.COD2 == cod2
                ),
                None,
            )
        if item is None:
            raise ValueError(f"No existe stock para COD1={cod1!r} COD2={cod2!r}")

        cantidad = Decimal(cantidad)
        stunid = item.STUNID or Decimal("0")

        if es_salida:
            item.STUNID = stunid - cantidad
            item.SALMES = (item.SALMES or Decimal("0")) + cantidad
        else:
            item.STUNID = stunid + cantidad
            item.ENTMES = (item.ENTMES or Decimal("0")) + cantidad

        self.repos.stock().db.commit()
        return item


@dataclass
class RenglonMovimientoStock:
    """Un renglón de la grilla de `StockMovimientoService.emitir_movimiento()`
    — Sección+código ya resueltos (ver `factura_renglon.resolver_seccion_
    renglon()`/`armar_codigo_renglon()`, reusados por la UI para construir
    `cod2` a partir de los segmentos, igual que en el Facturador) y la
    cantidad tipeada directo por el operador (a diferencia de un renglón
    de Factura, acá NO hay multiplicadores ALF4-6 — es un solo número en
    la unidad de la Sección, UNID/MTRS/KG)."""

    cod1: str
    cod2: str
    cantidad: Decimal


@dataclass
class ResultadoMovimientoStock:
    tipo_codigo: str
    nro_comprobante: int
    cantidad_renglones: int


class StockMovimientoService:
    """Réplica de `Stock.frm` ("Ingreso de Movimientos al Stock") —
    `Sub Graba()` (líneas 1060-1286), incluida la validación de
    comprobante duplicado (`Sub BuscaCPBTE()`, líneas 822-868) y el alta/
    actualización de `Despachos` cuando la Entrada es por Importación.

    **Alcance de Secciones ampliado** (confirmado con el usuario,
    2026-08-16): el legacy sólo permitía cargar movimientos para
    Secciones de código simple (`ALF1` en 'NRO'/'MM1', sin código
    compuesto) — acá se levantó esa restricción reutilizando el mismo
    armador de segmentos ya construido para Artículos/Facturador
    (`factura_renglon.resolver_seccion_renglon`). La restricción de
    UNIDAD (`ALF7` debe ser UNID/MTRS/KG — `Stock.frm Sub BuscaSeccion()`
    líneas 885-894) SÍ se mantiene: es inherente a que `Stock.STUNID` es
    un único acumulador, no una restricción arbitraria de la pantalla.
    """

    # Réplica de `Stock.frm Sub Graba()` (líneas 1168-1192), confirmada
    # contra la tabla de decodificación REAL de `VerStock.frm Sub
    # CargaMOVS` (líneas 1098-1111, ej. `Case "14": TipoMS = "Compra"`) —
    # NO la de `Sub BuscaCPBTE()` (líneas 829-853), que tiene Importación/
    # Compra Local y Ajuste-/Promoción invertidos respecto a Graba() Y a
    # VerStock. Dos fuentes independientes (la que graba, la que
    # reporta) coinciden entre sí y difieren sólo de la que hace el
    # chequeo de duplicado — bug real de copy-paste en BuscaCPBTE(), no
    # ambigüedad de negocio, no hizo falta preguntar (mismo criterio que
    # los bugs de padding de Access ya encontrados).
    CODIGOS_ENTRADA = {
        "fabrica": "13",
        "compra_local": "14",
        "importacion": "15",
        "ajuste_mas": "18",
    }
    CODIGOS_SALIDA = {
        "remito": "23",
        "promocion": "24",
        "ajuste_menos": "25",
        "ajuste_inv_menos": "28",
    }

    UNIDADES_VALIDAS = {"UNID", "MTRS", "KG"}

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def validar_seccion(self, cod_seccion: str) -> None:
        """Réplica de `Stock.frm Sub BuscaSeccion()` (líneas 870-911):
        la Sección debe existir y su Unidad de Facturación (`ALF7`) tiene
        que ser UNID/MTRS/KG — ver docstring de la clase."""
        seccion = self.repos.fctablas().by_ctab_cod("SC", cod_seccion)
        if seccion is None:
            raise ValueError(f"La Sección {cod_seccion} No Existe en la Tabla")
        alf7 = (seccion.ALF7 or "").strip()
        if alf7 not in self.UNIDADES_VALIDAS:
            raise ValueError(f"La Sección '{cod_seccion}' no se vende por unidades")

    def emitir_movimiento(
        self,
        *,
        es_entrada: bool,
        forma: str,
        nro_comprobante: int,
        renglones: list[RenglonMovimientoStock],
        usuario: str,
        nro_despacho: Optional[str] = None,
        fecha: Optional[date] = None,
    ) -> ResultadoMovimientoStock:
        """Graba el movimiento completo en una única transacción atómica
        (Stock/MovStock/Despachos), réplica de `Sub Graba()`.

        **Nº de Despacho gateado explícitamente** (`if nro_despacho`, NO
        siempre como el legacy): `Graba()` toca la tabla `Despachos` para
        TODOS los renglones sin importar el tipo de movimiento, usando
        `Text2.Text` (el campo de cabecera, vacío salvo Entrada+
        Importación) como `NRODESP` — para cualquier otro movimiento
        (Salida, o Entrada que no sea Importación) esto crea/actualiza un
        registro `Despachos` con `NRODESP=''` (lote fantasma) y puede
        dejar su `STOCK` negativo sin ningún control. El selector de lote
        por renglón que sí se ve durante una Salida (`FG2`/columna 4 de
        FG1) tampoco se usa: `Graba()` nunca lee esa columna, sólo
        `Text2.Text`. Confirmado con lectura completa del código, no es
        una suposición: acá `Despachos` sólo se toca cuando el llamador
        pasa un `nro_despacho` real (Entrada+Importación, igual que el
        único caso donde el legacy lo puebla con intención), evitando el
        lote fantasma.
        """
        codigos = self.CODIGOS_ENTRADA if es_entrada else self.CODIGOS_SALIDA
        if forma not in codigos:
            raise ValueError(f"Forma de movimiento inválida: {forma!r}")
        tipo_codigo = codigos[forma]

        if es_entrada and forma == "importacion" and not nro_despacho:
            raise ValueError("Debe colocar el Nº de DESPACHO si es importación")

        if not renglones:
            raise ValueError("No hay renglones para grabar")

        if self.repos.movstock().existe_comprobante(tipo_codigo, nro_comprobante):
            raise ValueError("Ya existe un movimiento con ese NÚMERO")

        fecha = fecha or date.today()
        usuario6 = (usuario or "")[:6]

        try:
            for item_idx, renglon in enumerate(renglones, start=1):
                self._grabar_renglon(
                    renglon, item_idx, es_entrada, tipo_codigo, nro_comprobante, nro_despacho, fecha, usuario6
                )
            self.db.commit()
        except Exception:
            self.db.rollback()
            raise

        return ResultadoMovimientoStock(
            tipo_codigo=tipo_codigo, nro_comprobante=nro_comprobante, cantidad_renglones=len(renglones)
        )

    def _grabar_renglon(
        self,
        renglon: RenglonMovimientoStock,
        item_idx: int,
        es_entrada: bool,
        tipo_codigo: str,
        nro_comprobante: int,
        nro_despacho: Optional[str],
        fecha: date,
        usuario6: str,
    ) -> None:
        cantidad = Decimal(renglon.cantidad)

        # --- Stock (upsert) ---------------------------------------------
        stock = self.repos.stock().by_cod1_cod2(renglon.cod1, renglon.cod2)
        if stock is None:
            stock = Stock(
                COD1=renglon.cod1, COD2=renglon.cod2,
                STMIN=Decimal("0"), STMAX=Decimal("0"), STREP=Decimal("0"),
                DEP1=0, EST1=0, ESTT1=0, DEP2=0, EST2=0, ESTT2=0,
                STUNID=Decimal("0"), PULG=Decimal("0"), MTR=Decimal("0"),
                STANT=Decimal("0"), ENTMES=Decimal("0"), SALMES=Decimal("0"),
                AJEMES=Decimal("0"), AJSMES=Decimal("0"),
            )
            self.db.add(stock)

        if es_entrada:
            stock.STUNID = (stock.STUNID or Decimal("0")) + cantidad
            stock.ENTMES = (stock.ENTMES or Decimal("0")) + cantidad
        else:
            stock.STUNID = (stock.STUNID or Decimal("0")) - cantidad
            stock.SALMES = (stock.SALMES or Decimal("0")) + cantidad
        stock.FACTUAL = fecha
        stock.USUARIO = usuario6

        # --- MovStock (historial, siempre un insert nuevo) ---------------
        self.db.add(
            MovStock(
                COD1=renglon.cod1, COD2=renglon.cod2, FECHA=fecha, TIPO=tipo_codigo,
                PTOVTA=1, CPBTE=nro_comprobante, ITEM=item_idx,
                PULG=Decimal("0"), MTR=Decimal("0"), MILIM=0, TELAS=0,
                CANT=cantidad, USUARIO=usuario6,
            )
        )

        # --- Despacho/Lote (sólo si el llamador pasó un lote real) -------
        if not nro_despacho:
            return
        despacho = self.repos.despacho().by_nrodesp(renglon.cod1, renglon.cod2, nro_despacho)
        if despacho is None:
            despacho = Despacho(
                COD1=renglon.cod1, COD2=renglon.cod2, NRODESP=nro_despacho,
                CPBTE=nro_comprobante, ENTRADA=Decimal("0"), SALIDAS=Decimal("0"),
                STOCK=Decimal("0"), FECENT=fecha, FECUSAL=fecha,
            )
            self.db.add(despacho)
        if es_entrada:
            despacho.STOCK = (despacho.STOCK or Decimal("0")) + cantidad
            despacho.ENTRADA = (despacho.ENTRADA or Decimal("0")) + cantidad
        else:
            despacho.STOCK = (despacho.STOCK or Decimal("0")) - cantidad
            despacho.SALIDAS = (despacho.SALIDAS or Decimal("0")) + cantidad


# ---------------------------------------------------------------------------
# ClienteService
# ---------------------------------------------------------------------------


@dataclass
class EstadisticaFacturacion:
    cliente: int
    desde: date
    hasta: date
    cantidad_comprobantes: int
    total_facturado: Decimal


class ClienteService:
    """Búsquedas y estadísticas de clientes, apoyado en CuentaCorrienteService
    para deuda (fuente canónica: ledger de Ctasctes, no Clientes.DEUDA)."""

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)
        self.cuentas = CuentaCorrienteService(db)

    def buscar(
        self,
        codigo: Optional[int] = None,
        nombre: Optional[str] = None,
        cuit: Optional[str] = None,
    ):
        """Búsqueda combinada: código exacto, nombre parcial o CUIT exacto."""
        cliente_repo = self.repos.cliente()
        if codigo is not None:
            resultado = cliente_repo.by_codigo(codigo)
            return [resultado] if resultado else []
        if cuit is not None:
            resultado = cliente_repo.by_cuit(cuit)
            return [resultado] if resultado else []
        if nombre is not None:
            return cliente_repo.by_nombre(nombre)
        return cliente_repo.activos()

    def deuda_cliente(self, clte: int) -> Decimal:
        """Delegado a CuentaCorrienteService (recalculado, no cacheado)."""
        return self.cuentas.saldo_cliente(clte)

    def puede_dar_de_baja(self, clte: int) -> tuple[bool, str]:
        """Decide si un cliente se puede eliminar.

        El legacy (Abmclte.frm Command7_Click -> Grabacion(), TipoMov3=2)
        hacía un DELETE físico sobre Clientes sin validar nada. Se agregó
        esta validación (confirmada con el usuario, no reemplaza el
        comportamiento legacy sin avisar): bloquear la baja si el cliente
        tiene saldo distinto de cero o cualquier movimiento histórico en
        el ledger de Ctasctes, para no dejar huérfanos esos registros.

        Devuelve (True, "") si se puede dar de baja, o (False, motivo) si no.
        """
        saldo = self.cuentas.saldo_cliente(clte)
        movimientos = self.repos.ctascte().by_cliente(clte)

        if saldo != 0 or movimientos:
            return (
                False,
                f"El cliente tiene saldo (${saldo}) o movimientos en Cuenta Corriente.",
            )
        return True, ""

    def estadisticas_facturacion(
        self, clte: int, desde: date, hasta: date
    ) -> EstadisticaFacturacion:
        """Totales facturados a un cliente en un rango de fechas.

        Usa FcivaVta (encabezados) y la misma fórmula de total confirmada en
        FacturaService (bruto + IVA Inscripto + percepción IIBB — sin
        IVANOINS, código legacy muerto de ClteCIVA=2, ver TotalFactura).
        """
        fciva_repo = self.repos.fciva_vta()
        comprobantes = [
            c
            for c in fciva_repo.by_cliente(clte)
            if c.FECHA is not None and desde <= c.FECHA <= hasta
        ]
        total = Decimal("0")
        for c in comprobantes:
            bruto = c.GRINS or Decimal("0")
            iva = c.IVAINS or Decimal("0")
            percepcion = c.TOTIB or Decimal("0")
            total += _round2(bruto) + _round2(iva) + _round2(percepcion)

        return EstadisticaFacturacion(
            cliente=clte,
            desde=desde,
            hasta=hasta,
            cantidad_comprobantes=len(comprobantes),
            total_facturado=total,
        )


# ---------------------------------------------------------------------------
# ArticuloService
# ---------------------------------------------------------------------------


class ArticuloService:
    """Búsquedas y validaciones de artículos (ABM, migración de AbmArt.frm).

    No expone nada de Stock acá a propósito: Articulo.STOCK/STMIN/STANT son
    campos manuales que ya se documentaron como no confiables (ver
    ArticuloRepository.bajo_stock) — el stock real vive en `Stock.STUNID`
    (StockService), una pantalla distinta.
    """

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def buscar(
        self,
        cod1: Optional[str] = None,
        cod2: Optional[str] = None,
        descri: Optional[str] = None,
    ):
        """Búsqueda combinada: clave compuesta exacta o descripción parcial."""
        articulo_repo = self.repos.articulo()
        if cod1 is not None and cod2 is not None:
            resultado = articulo_repo.by_cod1_cod2(cod1, cod2)
            return [resultado] if resultado else []
        if descri is not None:
            return articulo_repo.by_descri(descri)
        return []

    def puede_dar_de_baja(self, cod1: str, cod2: str) -> tuple[bool, str]:
        """Decide si un artículo se puede eliminar.

        El legacy (AbmArt.frm Command6_Click -> Grabacion(), TipoMov2=2)
        hacía un DELETE físico sobre Articulo sin validar nada. Se agregó
        esta validación (confirmada con el usuario, mismo criterio que
        ClienteService.puede_dar_de_baja): bloquear si el artículo tiene
        renglones históricos en Fcestad1 (facturado alguna vez).
        """
        movimientos = self.repos.fcestad1().by_cod1_cod2(cod1, cod2)
        if movimientos:
            return False, f"El artículo tiene {len(movimientos)} movimiento(s) facturado(s)."
        return True, ""

    # ------------------------------------------------------------------
    # Modificación de Precios por Porcentaje (ModPrec.frm)
    # ------------------------------------------------------------------

    @staticmethod
    def calcular_precio_con_porcentaje(precio: Decimal, porcentaje: Decimal) -> Decimal:
        """precio_nuevo = precio_actual × (1 + porcentaje/100), redondeado a 2 decimales.

        Réplica de ModPrec.frm CargaGrilla()/AGrabar():
        `PCos = RgART!prec * ((CCur(Text5.Text) / 100) + 1)` seguido de
        `RgART!prec = CCur(FG1.TextMatrix(I1, 3))` — pese a que la
        variable legacy se llama "PCos" (nombre engañoso, copiado de
        otra pantalla), el resultado se graba en `PREC`, nunca en
        `PCOS`. Acá se nombra `precio_nuevo` para no arrastrar la
        confusión. Porcentaje negativo = descuento (el propio tooltip
        del legacy lo indica: "colocar signo '-' adelante").
        """
        return _round2(precio * (porcentaje / Decimal("100") + 1))

    def candidatos_modificacion_precio(
        self,
        cod1: Optional[str] = None,
        desde: Optional[int] = None,
        hasta: Optional[int] = None,
    ) -> list[Articulo]:
        """Artículos candidatos a la modificación masiva de precio.

        `cod1=None` → todos los artículos (modo "Todos" de ModPrec.frm).
        `cod1` dado → artículos de esa Sección; si además se pasan
        `desde`/`hasta`, se filtran comparando `Val(COD2)` como un único
        número contra ese rango — réplica EXACTA de
        `CargaGrilla(): If Val(RgART!Cod2) < Val(Text2.Text) Or
        Val(RgART!Cod2) > Val(Text4.Text) Then GoTo Salta`.

        AVISO (verificado leyendo el código, no es una suposición): el
        legacy también muestra en pantalla un SEGUNDO par DESDE/HASTA
        (Text1/Text3) para Secciones con dos dimensiones de código (ej.
        GPN: mm + telas) — pero `CargaGrilla()` nunca los lee, sólo usa
        el primer par contra el COD2 completo. Es un campo de UI sin
        cablear, igual categoría que otros controles fantasma ya
        encontrados (Abmclte Remito/Flete/Lista) — no se migra ese
        segundo filtro porque nunca tuvo efecto real.
        """
        if cod1 is None:
            return self.repos.articulo().todos()

        articulos = self.repos.articulo().by_seccion(cod1)
        if desde is None and hasta is None:
            return articulos

        limite_inf = desde if desde is not None else 0
        limite_sup = hasta if hasta is not None else 99999999
        return [
            a for a in articulos if limite_inf <= _val_vb6(a.COD2) <= limite_sup
        ]

    def aplicar_modificacion_precio(self, articulos: list[Articulo], porcentaje: Decimal) -> int:
        """Aplica el nuevo precio a cada artículo de la lista y confirma.
        Devuelve la cantidad de artículos modificados."""
        articulo_repo = self.repos.articulo()
        for art in articulos:
            precio_actual = art.PREC or Decimal("0")
            nuevo = self.calcular_precio_con_porcentaje(precio_actual, porcentaje)
            articulo_repo.update(art.id, {"PREC": nuevo})
        return len(articulos)


# ---------------------------------------------------------------------------
# TablaService
# ---------------------------------------------------------------------------


class TablaService:
    """Validaciones del ABM de Tablas genéricas (Fctabla1), migración de
    ABMTablas.frm + Acttabla.frm.

    El legacy (Acttabla.frm Grabacion(), TIPOMov=2) hacía un DELETE físico
    sobre Fctabla1 sin validar nada, igual que Abmclte/AbmArt. Acá se
    agrega la misma protección, pero por categoría: cada CTAB tiene un
    "referenciador" distinto (quién más lo usa) confirmado leyendo el VB6:
      - SC (Sección): la usan artículos (Articulo.COD1).
      - UM (Unidad de Medida): la puede usar cualquier Sección en sus 7
        slots ALF1-ALF7 (Acttabla.frm bloques Especificación/Cálculo de
        Precio/Unid. Facturación — ver `resolver_segmentos` en
        `migration/ui/articulo_codigo.py` para el uso real de ALF1-2).
      - VD/ZN/CV (Vendedor/Zona/Cond.Venta): los usan clientes
        (Cliente.VEND/ZONA/CVTA).
      - PV/MT/VS: no se encontró ninguna pantalla ya migrada que las
        referencie por clave foránea — se permite la baja sin más chequeo
        (no hay nada confirmado que romper todavía).
    """

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def buscar(self, ctab: str):
        """Todas las entradas de una categoría (CTAB), orden por COD."""
        return self.repos.fctablas().by_ctab(ctab)

    def puede_dar_de_baja(self, ctab: str, cod: str) -> tuple[bool, str]:
        ctab = ctab.strip().upper()
        cod = cod.strip()

        if ctab == "SC":
            if self.repos.articulo().by_cod1(cod) is not None:
                return False, "Hay artículos que usan esta Sección."
            return True, ""

        if ctab == "UM":
            secciones = self.repos.fctablas().secciones_que_usan_unidad(cod)
            if secciones:
                nombres = ", ".join((s.COD or "").strip() for s in secciones)
                return False, f"La usan estas Secciones: {nombres}."
            return True, ""

        if ctab in ("VD", "ZN", "CV"):
            try:
                cod_num = int(cod)
            except ValueError:
                return True, ""
            clientes = {
                "VD": self.repos.cliente().by_vendedor,
                "ZN": self.repos.cliente().by_zona,
                "CV": self.repos.cliente().by_cond_venta,
            }[ctab](cod_num)
            if clientes:
                etiqueta = {"VD": "Vendedor", "ZN": "Zona", "CV": "Condición de Venta"}[ctab]
                return False, f"Hay clientes con esta {etiqueta} asignada."
            return True, ""

        return True, ""


# ---------------------------------------------------------------------------
# EmisionFacturaService
# ---------------------------------------------------------------------------


@dataclass
class RenglonEmision:
    """Un renglón de Detalle ya resuelto y calculado (Sección/Artículo,
    código, cantidades, precio) — lo produce el llamador combinando
    `factura_renglon.py` (armado de código y cálculo de Precio Unitario/
    Importe) y `DespachoSelectorDialog` (lote elegido, si corresponde).
    `EmisionFacturaService` no recalcula nada de esto, sólo lo persiste.

    Los 5 campos `pulg/mtr/milim/telas/cantidad_unidades` corresponden 1:1
    a las columnas dinámicas 1/2/3/4/6 de la grilla `FG1` del legacy
    (`Fctabla1(UM).NUMSD3`, ver `factura_renglon.SegmentoRenglon.posicion`)
    — sólo `mtr` (posición 2, "Mtr/Kg") y `cantidad_unidades` (posición 6,
    "Unidades") descuentan stock; `pulg`/`milim`/`telas` son dimensiones
    del código, se guardan para estadística pero no mueven stock (fiel a
    `EmiFact.frm Sub Graba`, líneas 1864-1882 y 1996-2005). **Sólo
    `cantidad_unidades` cuenta para `Fcestad1.CANT`/`FcivaVta.TOTCAN`/
    `Totales.UNIDA-B`** (línea 2044 y 2077: ambas leen exclusivamente la
    columna 6) — un renglón facturado por Kg/Mtr no suma "unidades", es
    el comportamiento real, no una omisión.
    """

    cod1: str  # Sección (o "**" para ítem libre, ver SECCION_ITEM_LIBRE)
    cod2: str  # código compuesto (RESTO) — "" para ítem libre
    descripcion: str
    precio_unitario: Decimal
    importe: Decimal  # bruto de este renglón, antes de bonificación (FG1 col 8)
    pulg: Decimal = Decimal("0")
    mtr: Decimal = Decimal("0")
    milim: int = 0
    telas: int = 0
    cantidad_unidades: Decimal = Decimal("0")
    nrodesp_elegido: Optional[str] = None  # lote elegido en DespachoSelectorDialog


@dataclass
class ResultadoEmision:
    fcivavta_id: int
    ctascte_id: int
    totales_id: int
    total: "TotalFactura"


class EmisionFacturaService:
    """Persistencia atómica de una Factura ya calculada — equivalente a
    `EmiFact.frm Sub Graba()`, acotado al alcance confirmado con el
    usuario: **sólo Factura, Letra A/B**. Nota de Crédito, Nota de
    Débito, Remito y Exportación quedan para una fase posterior — esas
    ramas de `Graba()` (TipoFac 2/3/5, Letra "C"/"R", `ImputaFact()`) no
    se migran todavía.

    No calcula precios/descuentos de renglón (ya resueltos por
    `factura_renglon.py` al cargar el Detalle) ni el total de la factura
    (`FacturaService.calcular_total()`, ya invocado por el llamador) —
    sólo escribe el resultado en Stock/MovStock/Fcestad1/Despachos/
    FcivaVta/Clientes.DEUDA/Ctasctes/Totales, en una única transacción.

    **Simplificación real sobre el legacy, no un cambio de negocio**: VB6
    usa 3 transacciones ADO separadas (`dbCCTE`/`dbFACT`/`dbSTOK`, eran 3
    archivos Access distintos) con un flag `NoAnduvo` para decidir
    Commit/Rollback de las tres a mano. Acá todo vive en una sola base
    Postgres: colapsa a un único `session.commit()` con
    `session.rollback()` ante cualquier excepción — más seguro que el
    original (una falla a mitad de camino no puede dejar Stock
    actualizado sin su Ctasctes correspondiente).

    **Bug real corregido, no replicado** (`Graba()` línea ~1876): para
    artículos cuya cantidad viaja en la columna "Mtr/Kg" (posición 2), el
    legacy hace `Stock.STUnid = Stock.Mtr - cantidad` (pisa con un campo
    legado no sincronizado) en vez de `Stock.STUnid = Stock.STUnid -
    cantidad`. Acá SIEMPRE se resta/suma sobre `STUnid`, sin importar qué
    columna trae la cantidad — afecta 12 Secciones reales en producción
    hoy (AC, BAR, BL, CT, GPN, GPS, HL, MHR1, MHR12, MHR2, PISO, PL), ver
    `pyqt6_ui_progress.md`.

    **Detalle menor encontrado y NO replicado**: `Fcestad1.USUARIO` en el
    legacy está hardcodeado a 3 espacios literales (`RgFACT!USUARIO =
    "   "`), con la línea correcta (`Mid(Red_Usuario, 1, 6)`) comentada
    arriba — un descuido, no una decisión (Stock/MovStock/Ctasctes sí
    graban el usuario real un par de líneas antes/después). Acá se graba
    el usuario real en los tres, por consistencia — el campo no afecta
    ningún cálculo de dinero/stock, sólo la trazabilidad de quién facturó.
    """

    # MovStock.TIPO por Letra — sólo Factura (TipoFac=1); NC (11/12) y
    # Nota de Crédito 'C' (17)/Remito 'R' (23) quedan fuera del alcance
    # actual (EmiFact.frm:1902-1928).
    CODMOVS_POR_LETRA = {"A": "21", "B": "22"}

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def emitir_factura(
        self,
        cliente: Cliente,
        letra: str,
        punto_venta: int,
        numero_comprobante: int,
        renglones: list[RenglonEmision],
        total: "TotalFactura",
        porcentaje_iibb: Decimal,
        usuario: str,
        fecha: Optional[date] = None,
        tipo_venta: str = "",
        motivo: int = 0,
    ) -> ResultadoEmision:
        """Réplica de `EmiFact.frm Sub Graba()`, rama Factura (TipoFac=1).

        `total` es el resultado ya calculado por
        `FacturaService.calcular_total()` sobre la suma de
        `renglon.importe` de todos los renglones menos la bonificación
        total — este método NO lo recalcula, sólo lo persiste (ver
        docstring de la clase).
        """
        if letra not in self.CODMOVS_POR_LETRA:
            raise ValueError(
                f"Letra '{letra}' no soportada por EmisionFacturaService — "
                "el alcance actual es sólo Factura A/B."
            )
        if not renglones:
            raise ValueError("La factura no tiene renglones.")

        fecha = fecha or date.today()
        ahora = datetime.now()
        usuario6 = (usuario or "")[:6]
        codmovs = self.CODMOVS_POR_LETRA[letra]

        try:
            total_cantidad_unidades = Decimal("0")
            for item, renglon in enumerate(renglones, start=1):
                self._grabar_renglon(renglon, item, codmovs, punto_venta, numero_comprobante, cliente, letra, fecha, ahora, usuario6)
                total_cantidad_unidades += renglon.cantidad_unidades

            totales = self._upsert_totales(fecha, letra, total, total_cantidad_unidades, usuario6)

            fcivavta = FcivaVta(
                FECHA=ahora.date(),
                PTOVTA=punto_venta,
                CPBTE=numero_comprobante,
                LETRA=letra,
                TIPO="1",
                CLTE=cliente.CODIGO,
                NOMB=cliente.NOMB,
                PCIA=cliente.PCIA,
                CVTA=str(cliente.CVTA) if cliente.CVTA is not None else None,
                MOTI=tipo_venta,
                CIVA=str(cliente.CIVA) if cliente.CIVA is not None else None,
                CUIT=(cliente.CUIT or "")[:14],
                VEND=cliente.VEND,
                ZONA=0,
                TOTCAN=int(total_cantidad_unidades),
                GRINS=total.neto_gravado,
                GRNOINS=Decimal("0"),
                IVAINS=total.iva,
                IVANOINS=Decimal("0"),
                PORCIB=Decimal(porcentaje_iibb),
                TOTIB=total.percepcion_iibb,
                EXENTO=Decimal("0"),
                TOTCOS=Decimal("0"),
                ITEMS=len(renglones),
                BON=total.descuento,
                NOIMPR="0",
                COMIS=Decimal("0"),
            )
            self.db.add(fcivavta)

            cliente.DEUDA = (cliente.DEUDA or Decimal("0")) + total.total
            cliente.USUARIO = usuario6
            cliente.FACTUAL = fecha

            dias_vto = self._dias_vencimiento(cliente.CVTA)
            ctascte = Ctascte(
                CLTE=cliente.CODIGO,
                FECHA=fecha,
                TIPO=1,
                PREFIJO=punto_venta,
                CPBTE=numero_comprobante,
                LETRA=letra,
                IMPUT1="0 ", IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
                IMPTE=total.total,
                DEBE=total.total,
                CVTA=str(cliente.CVTA) if cliente.CVTA is not None else None,
                BON=total.descuento,
                TIPO9="0",
                MOTI=str(motivo),
                FECVTO=fecha + timedelta(days=dias_vto),
                USUAR=usuario6,
            )
            self.db.add(ctascte)

            self.db.commit()
        except Exception:
            self.db.rollback()
            raise

        self.db.refresh(fcivavta)
        self.db.refresh(ctascte)
        self.db.refresh(totales)
        return ResultadoEmision(
            fcivavta_id=fcivavta.id, ctascte_id=ctascte.id, totales_id=totales.id, total=total
        )

    # ------------------------------------------------------------------
    def _dias_vencimiento(self, cvta: Optional[int]) -> int:
        """Días de la Cond. de Venta del cliente (`Fctabla1` CTAB='CV',
        `NUMSD3`) — réplica de `Graba()` líneas 2274-2285. 30 días si no
        se encuentra la Cond. de Venta (mismo default del legacy)."""
        if cvta is None:
            return 30
        entrada = self.repos.fctablas().by_ctab_cod("CV", str(cvta))
        if entrada is None or entrada.NUMSD3 is None:
            return 30
        return int(entrada.NUMSD3)

    def _grabar_renglon(
        self,
        renglon: RenglonEmision,
        item: int,
        codmovs: str,
        punto_venta: int,
        numero_comprobante: int,
        cliente: Cliente,
        letra: str,
        fecha: date,
        ahora: datetime,
        usuario6: str,
    ) -> None:
        cod1, cod2 = renglon.cod1, renglon.cod2

        # --- Stock (upsert) ---------------------------------------------
        stock = self.repos.stock().by_cod1_cod2(cod1, cod2)
        if stock is None:
            stock = Stock(
                COD1=cod1, COD2=cod2, STMIN=0, STMAX=0, STREP=0, DEP1=0, EST1=0, ESTT1=0,
                DEP2=0, EST2=0, ESTT2=0, STANT=0, PULG=Decimal("0"), MTR=Decimal("0"),
                STUNID=Decimal("0"), ENTMES=Decimal("0"), SALMES=Decimal("0"),
                AJEMES=Decimal("0"), AJSMES=Decimal("0"),
            )
            self.db.add(stock)

        if renglon.cantidad_unidades > 0:
            stock.STUNID = (stock.STUNID or Decimal("0")) - renglon.cantidad_unidades
            stock.SALMES = (stock.SALMES or Decimal("0")) + renglon.cantidad_unidades
        if renglon.mtr > 0:
            # Bug corregido (ver docstring de la clase): SIEMPRE STUnid, el
            # legacy pisa con Stock.Mtr acá para esta rama.
            stock.STUNID = (stock.STUNID or Decimal("0")) - renglon.mtr
            stock.SALMES = (stock.SALMES or Decimal("0")) + renglon.mtr
        stock.FACTUAL = fecha
        stock.USUARIO = usuario6

        # --- MovStock (siempre un insert nuevo, sin dedup — fiel al
        # legacy, que tiene el chequeo de duplicado comentado) -----------
        self.db.add(
            MovStock(
                COD1=cod1, COD2=cod2, TIPO=codmovs, PTOVTA=punto_venta, CPBTE=numero_comprobante,
                ITEM=item, FECHA=ahora, CANT=renglon.cantidad_unidades, PULG=renglon.pulg,
                MTR=renglon.mtr, MILIM=renglon.milim, TELAS=renglon.telas, USUARIO=usuario6,
            )
        )

        # --- Fcestad1 (estadística por artículo) -------------------------
        bon_renglon = self._bonificacion_renglon(cliente.CODIGO, cod1, renglon.importe)
        self.db.add(
            Fcestad1(
                COD1=cod1, COD2=cod2, TIPO=1, LETRA=letra, PTOVTA=punto_venta,
                CPBTE=numero_comprobante, ITEM=item, CLTE=cliente.CODIGO, FECHA=fecha,
                GPTIPO="", GPDSD="", GPHST="", PULG=renglon.pulg, MTR=renglon.mtr,
                MILIM=renglon.milim, TELAS=renglon.telas, CANT=renglon.cantidad_unidades,
                PCOS=Decimal("0"), PVTA=renglon.precio_unitario, PESP=renglon.precio_unitario,
                BON=bon_renglon, IVA=Decimal("0"), IMPTE=renglon.importe, TIPO9="0",
                USUARIO=usuario6,
            )
        )

        # --- Despacho/Lote (trazabilidad adicional, opcional) -------------
        if renglon.nrodesp_elegido:
            despacho = self.repos.despacho().by_nrodesp(cod1, cod2, renglon.nrodesp_elegido)
            if despacho is not None:
                if renglon.cantidad_unidades > 0:
                    despacho.STOCK = (despacho.STOCK or Decimal("0")) - renglon.cantidad_unidades
                    despacho.SALIDAS = (despacho.SALIDAS or Decimal("0")) + renglon.cantidad_unidades
                if renglon.mtr > 0:
                    despacho.STOCK = (despacho.STOCK or Decimal("0")) - renglon.mtr
                    despacho.SALIDAS = (despacho.SALIDAS or Decimal("0")) + renglon.mtr
                despacho.FECUSAL = fecha
            # Si no se encuentra (cambió entre la selección y la emisión),
            # se sigue sin tocar Despachos — fiel al `GoTo SinGrabar1` del
            # legacy, no es motivo para abortar la factura completa.

    def _bonificacion_renglon(self, clte: int, cod_seccion: str, importe: Decimal) -> Decimal:
        """Recalcula fresco el % de bonificación en cascada para
        `Fcestad1.BON` (réplica de `DxClte` + bucle, `Graba()` líneas
        1819-1827) — no reutiliza ningún valor ya mostrado en pantalla,
        mismo criterio de frescura que el resto del sistema."""
        fila = self.repos.dtoxclte().by_cliente_seccion(clte, cod_seccion)
        if fila is None:
            return Decimal("0")
        porcentajes = [fila.DTO1, fila.DTO2, fila.DTO3, fila.DTO4, fila.DTO5]
        porcentajes = [Decimal(p) for p in porcentajes if p is not None]
        return FacturaService(self.db).calcular_bonificacion_cascada(importe, porcentajes)

    def _upsert_totales(
        self, fecha: date, letra: str, total: "TotalFactura", total_cantidad_unidades: Decimal, usuario6: str
    ) -> Totales:
        totales = self.repos.totales().by_fecha(fecha)
        if totales is None:
            totales = Totales(FECHA=fecha)
            self.db.add(totales)

        def _sumar(campo: str, delta) -> None:
            actual = getattr(totales, campo) or 0
            setattr(totales, campo, actual + delta)

        # TotPcos nunca se asigna en Graba() (Dim local, se queda en 0 por
        # default de VB6) — confirmado dead, no se suma nada a PCOS.
        _sumar("PVTA", total.neto_gravado)
        _sumar("PESP", total.neto_gravado)
        if letra == "A":
            _sumar("FACA", 1)
            _sumar("UNIDA", int(total_cantidad_unidades))
            _sumar("PESPA", total.neto_gravado)
        else:  # "B" — único otro valor posible, ya validado en emitir_factura()
            _sumar("FACB", 1)
            _sumar("UNIDB", int(total_cantidad_unidades))
            _sumar("PESPB", total.neto_gravado)

        totales.FACTUAL = fecha
        totales.USUARIO = usuario6
        return totales


# ---------------------------------------------------------------------------
# EmisionNotaCreditoService — Nota de Crédito / Débito de "Concepto Libre"
# ---------------------------------------------------------------------------


def _dias_vencimiento_cond_venta(repos: RepositoryFactory, cvta: Optional[int]) -> int:
    """Días de la Cond. de Venta de un cliente (`Fctabla1` CTAB='CV',
    `NUMSD3`) — réplica de `EmiFact.frm Graba()` líneas 2274-2285. 30
    días si no se encuentra la Cond. de Venta (mismo default del
    legacy). Función compartida por `EmisionFacturaService`/
    `EmisionNotaCreditoService` (misma fórmula, cualquier tipo de
    comprobante que grabe en Ctasctes la necesita)."""
    if cvta is None:
        return 30
    entrada = repos.fctablas().by_ctab_cod("CV", str(cvta))
    if entrada is None or entrada.NUMSD3 is None:
        return 30
    return int(entrada.NUMSD3)


@dataclass
class ConceptoNotaCredito:
    """Un renglón de `DetNC.frm` (hasta 3 por comprobante, alcance
    "Concepto Libre" — Motivo distinto de 1/"DEV.MERC.", ver docstring
    de `EmisionNotaCreditoService`): descripción libre + importe + si
    ese importe está gravado por IVA (`Check1` de `DetNC.frm`, "CON IVA"/
    "SIN IVA")."""

    descripcion: str
    importe: Decimal
    con_iva: bool = True


@dataclass
class TotalNotaCreditoConcepto:
    """Desglose de una Nota de Crédito/Débito de "Concepto Libre".

    **Bug real encontrado y NO replicado** (`EmiFact.frm Sub
    ConectaAFIP()`, línea ~2644): el legacy arma el total que se PIDE a
    AFIP y se GRABA en Ctasctes/FCIVAVta como `TotBruto + TotIVAIns +
    ValIB` — sin sumar `TotSIVA` (el acumulado de las líneas "SIN IVA",
    `Check1.Value=0` en `DetNC.frm`) — pese a que el total que se
    MUESTRA e IMPRIME al operador sí lo incluye (`TOTNETO = TotBruto +
    TotIVAIns + TotIVANI + TotSIVA + ValIB`, líneas 1181/1669). Un
    importe "SIN IVA" cargado por el operador desaparecía
    silenciosamente del CAE pedido y de lo grabado. Decisión confirmada
    con el usuario (2026-08-19): acá SÍ se incluye, como importe no
    gravado — viaja a AFIP como `ImpTotConc` (`AfipWSFEv1Cliente.
    solicitar_cae`, hardcodeado en "0.00" en el legacy, línea 2635)."""

    base_gravada: Decimal  # suma de los conceptos "CON IVA" (antes de IVA)
    iva: Decimal
    no_gravado: Decimal  # suma de los conceptos "SIN IVA" (antes se perdía, ver arriba)
    total: Decimal


class NotaCreditoConceptoService:
    """Cálculo de totales de una Nota de Crédito/Débito de "Concepto
    Libre" (hasta 3 renglones de texto libre + importe, `DetNC.frm`) —
    equivalente conceptual a `FacturaService.calcular_total()` pero sin
    grilla de artículos ni bonificación en cascada (`DetNC.frm` no las
    tiene). Reusa la misma alícuota/gate de IVA que Factura (misma
    fuente `Parametro.IVAINS`/`Cliente.CIVA`, `FacturaService.
    CIVA_GRAVADO_MAX`)."""

    def __init__(self, db: Session):
        self.db = db
        self.factura_service = FacturaService(db)

    def calcular_total(self, conceptos: list[ConceptoNotaCredito], civa_cliente: int) -> TotalNotaCreditoConcepto:
        base_gravada = sum((c.importe for c in conceptos if c.con_iva), Decimal("0"))
        no_gravado = sum((c.importe for c in conceptos if not c.con_iva), Decimal("0"))

        gravado = civa_cliente < FacturaService.CIVA_GRAVADO_MAX
        iva = Decimal("0")
        if gravado and base_gravada > 0:
            iva = _round2(base_gravada * self.factura_service.alicuota_iva_inscripto())

        total = _round2(base_gravada) + iva + _round2(no_gravado)
        return TotalNotaCreditoConcepto(
            base_gravada=_round2(base_gravada), iva=iva, no_gravado=_round2(no_gravado), total=total
        )


@dataclass
class ResultadoEmisionNotaCredito:
    fcivavta_id: int
    ctascte_id: int
    totales_id: int
    imputacion_id: Optional[int]
    total: Decimal


class EmisionNotaCreditoService:
    """Persistencia atómica de una Nota de Crédito/Débito de "Concepto
    Libre" — equivalente a `EmiFact.frm Sub Graba()`, rama TipoFac 2/3
    con Motivo distinto de 1 ("DEV.MERC.", que usa la grilla de
    artículos y reversa Stock — pantalla aparte, "Nota de Crédito —
    Devolución de Mercadería", fuera del alcance de ESTE servicio).

    **Alcance de Motivo confirmado con el usuario (2026-08-19)**, fuente
    real `CabFact.frm Sub Combo5_Click` (líneas 726-746): `Motivo = 1`
    ("DEV.MERC.", único código real hoy con ese valor, confirmado contra
    `Fctabla1` real) muestra `DetFact.Show` (grilla de artículos);
    cualquier otro Motivo muestra `DetNC.Show` (lo que persiste ESTE
    servicio).

    Réplica de `Graba()` (líneas 2166-2318, rama común a TipoFac 2/3) +
    `Sub ImputaFact()` (líneas 2351-2437, sólo TipoFac=2): arma FcivaVta/
    Ctascte/Totales igual que Factura pero con TIPO=2 (NC) o 3 (ND), y
    si es NC (TIPO=2) imputa el importe contra un comprobante con deuda
    real elegido por el operador (`CuentaCorrienteService.imputar_pago`,
    `tipo_imputacion="2"`) — reusa el MISMO mecanismo de slots IMPUT1-6
    ya usado por Recibo, no una implementación paralela.

    **Nota de Débito NO imputa contra nada**: el legacy sólo llama
    `ImputaFact()` `If TipoFac = 2` (línea 2238) — la ND simplemente suma
    a la deuda del cliente igual que una Factura (`Graba()` línea 2244:
    `Else RgCLTE!DEUDA = RgCLTE!DEUDA + TOTNETO`). `comprobante_a_
    imputar` es obligatorio para NC y se ignora (debe venir `None`) para
    ND.

    **Sólo Letra A/B** — mismo alcance ya confirmado para Factura
    (`EmisionFacturaService`); Letra "C" (Consumidor Final papel) queda
    fuera, `FacturaService.letra_comprobante()` nunca la devuelve."""

    TIPO_NC = 2
    TIPO_ND = 3
    LETRAS_SOPORTADAS = ("A", "B")

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)
        self.cuentas = CuentaCorrienteService(db)

    def emitir_concepto(
        self,
        cliente: Cliente,
        tipo: int,
        letra: str,
        punto_venta: int,
        numero_comprobante: int,
        conceptos: list[ConceptoNotaCredito],
        total: TotalNotaCreditoConcepto,
        motivo: int,
        usuario: str,
        comprobante_a_imputar: Optional[Ctascte] = None,
        fecha: Optional[date] = None,
    ) -> ResultadoEmisionNotaCredito:
        if tipo not in (self.TIPO_NC, self.TIPO_ND):
            raise ValueError(f"Tipo {tipo!r} no soportado — sólo Nota de Crédito (2) o Débito (3).")
        if letra not in self.LETRAS_SOPORTADAS:
            raise ValueError(f"Letra '{letra}' no soportada — el alcance actual es sólo A/B.")
        if tipo == self.TIPO_NC and comprobante_a_imputar is None:
            raise ValueError("La Nota de Crédito necesita un comprobante contra el cual imputarse.")
        if not conceptos:
            raise ValueError("La Nota no tiene conceptos cargados.")

        fecha = fecha or date.today()
        usuario6 = (usuario or "")[:6]

        try:
            fcivavta = FcivaVta(
                FECHA=fecha,
                PTOVTA=punto_venta,
                CPBTE=numero_comprobante,
                LETRA=letra,
                TIPO=str(tipo),
                CLTE=cliente.CODIGO,
                NOMB=cliente.NOMB,
                PCIA=cliente.PCIA,
                CVTA=str(cliente.CVTA) if cliente.CVTA is not None else None,
                MOTI=str(motivo),
                CIVA=str(cliente.CIVA) if cliente.CIVA is not None else None,
                CUIT=(cliente.CUIT or "")[:14],
                VEND=cliente.VEND,
                ZONA=0,
                TOTCAN=0,
                GRINS=total.base_gravada,
                GRNOINS=Decimal("0"),
                IVAINS=total.iva,
                IVANOINS=Decimal("0"),
                PORCIB=Decimal("0"),
                TOTIB=Decimal("0"),
                # `EXENTO` es el campo real más cercano a "importe no
                # gravado" que tiene el esquema — mismo destino que el
                # `ImpTotConc` que se manda a AFIP (ver
                # `TotalNotaCreditoConcepto`).
                EXENTO=total.no_gravado,
                TOTCOS=Decimal("0"),
                ITEMS=len(conceptos),
                BON=Decimal("0"),
                NOIMPR="0",
                COMIS=Decimal("0"),
            )
            self.db.add(fcivavta)

            dias_vto = _dias_vencimiento_cond_venta(self.repos, cliente.CVTA)
            corr = self.repos.cliente().proximo_correlativo(cliente)
            ctascte = Ctascte(
                CLTE=cliente.CODIGO,
                FECHA=fecha,
                TIPO=tipo,
                PREFIJO=punto_venta,
                CPBTE=numero_comprobante,
                LETRA=letra,
                IMPUT1="0 ", IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
                IMPTE=total.total,
                DEBE=total.total,
                CVTA=str(cliente.CVTA) if cliente.CVTA is not None else None,
                BON=Decimal("0"),
                TIPO9="0",
                MOTI=str(motivo),
                FECVTO=fecha + timedelta(days=dias_vto),
                USUAR=usuario6,
            )
            self.db.add(ctascte)

            imputacion_id: Optional[int] = None
            if tipo == self.TIPO_NC:
                # Réplica de Graba() línea 2314 (`RgCCTE!imput1 = Corr`):
                # la propia fila recién creada se marca a sí misma con su
                # correlativo — mismo patrón de doble marca ya usado en
                # `EmisionReciboService` para la fila del propio Recibo.
                ctascte.IMPUT1 = str(corr)
                cliente.DEUDA = (cliente.DEUDA or Decimal("0")) - total.total
                resultado_imputacion = self.cuentas.imputar_pago(
                    comprobante=comprobante_a_imputar,
                    importe_aplicado=total.total,
                    corr=corr,
                    cpbte_recibo=numero_comprobante,
                    usuario=usuario6,
                    clte=cliente.CODIGO,
                    fecha=fecha,
                    commit=False,
                    tipo_imputacion=CuentaCorrienteService.TIPO_IMPUTACION_NOTA_CREDITO,
                )
                imputacion_id = resultado_imputacion.imputacion_id
            else:
                cliente.DEUDA = (cliente.DEUDA or Decimal("0")) + total.total

            cliente.USUARIO = usuario6
            cliente.FACTUAL = fecha

            totales = self._upsert_totales(fecha, tipo, letra, total)

            self.db.commit()
        except Exception:
            self.db.rollback()
            raise

        self.db.refresh(fcivavta)
        self.db.refresh(ctascte)
        self.db.refresh(totales)
        return ResultadoEmisionNotaCredito(
            fcivavta_id=fcivavta.id, ctascte_id=ctascte.id, totales_id=totales.id,
            imputacion_id=imputacion_id, total=total.total,
        )

    def _upsert_totales(self, fecha: date, tipo: int, letra: str, total: TotalNotaCreditoConcepto) -> Totales:
        """Réplica de `Graba()` Case 2 (NC, RESTA) / Case 3 (ND, SUMA),
        líneas 2126-2158. `TotBruto` ahí es sólo la base gravada ("CON
        IVA") — igual criterio acá (`total.base_gravada`): es un campo
        de reporte de ventas gravadas, no el importe realmente cobrado/
        acreditado (que sí incluye `no_gravado`, ver `TotalNotaCredito
        Concepto`). Sin `TotCan` real en el modo "Concepto Libre" (no
        hay cantidades, sólo importes) — `UNIDA/UNIDB` no se tocan,
        a diferencia de Factura."""
        totales = self.repos.totales().by_fecha(fecha)
        if totales is None:
            totales = Totales(FECHA=fecha)
            self.db.add(totales)

        def _sumar(campo: str, delta) -> None:
            actual = getattr(totales, campo) or 0
            setattr(totales, campo, actual + delta)

        signo = -1 if tipo == self.TIPO_NC else 1
        _sumar("PVTA", signo * total.base_gravada)
        _sumar("PESP", signo * total.base_gravada)
        campo_contador = {
            ("A", self.TIPO_NC): "NCA", ("B", self.TIPO_NC): "NCB",
            ("A", self.TIPO_ND): "NDA", ("B", self.TIPO_ND): "NDB",
        }.get((letra, tipo))
        if campo_contador:
            _sumar(campo_contador, 1)
        campo_pesp = "PESPA" if letra == "A" else "PESPB"
        _sumar(campo_pesp, signo * total.base_gravada)

        totales.FACTUAL = fecha
        return totales


# ---------------------------------------------------------------------------
# EmisionNotaCreditoMercaderiaService — Nota de Crédito por Devolución de
# Mercadería (Motivo=1/"DEV.MERC.")
# ---------------------------------------------------------------------------


class EmisionNotaCreditoMercaderiaService:
    """Persistencia atómica de una Nota de Crédito por Devolución de
    Mercadería (Motivo=1/"DEV.MERC.") — reusa la MISMA grilla/cálculo de
    Factura (`FacturaService.calcular_total()`, `RenglonEmision`) que el
    legacy reutiliza literalmente (`CabFact.frm Sub Combo5_Click`:
    `Motivo=1` muestra `DetFact.Show`, la grilla de artículos, en vez de
    `DetNC.Show`), pero graba TIPO=2 y REVERSA Stock en vez de
    descontarlo.

    Réplica de `EmiFact.frm Sub Graba()` (rama común TipoFac 2/3, ver
    `EmisionNotaCreditoService` para la parte compartida con "Concepto
    Libre") + el bloque de artículos (`ConArticulos:`, líneas 983-1076)
    + reversa de Stock (líneas 1864-1928, rama `Else` de `If TipoFac = 1
    Then ... Else STUnid = STUnid + cantidad ...`, con códigos `MovStock`
    dedicados `11`/`12` = N/Créd.'A'/'B') + `Sub ImputaFact()` (líneas
    2351-2437, mismo mecanismo de slots IMPUT1-6 ya usado por Recibo y
    por `EmisionNotaCreditoService`).

    **Precio de cada renglón = el precio REAL de la Factura original**
    (`Fcestad1.PVTA`/`.IMPTE` de ese renglón puntual), no el precio de
    catálogo actual — decisión del usuario (2026-08-19), mejora
    deliberada sobre el legacy (que reutilizaba la grilla en blanco,
    forzando a re-tipear el artículo a precio de catálogo). El llamador
    arma los `RenglonEmision` desde `Fcestad1Repository.by_comprobante()`
    de la Factura elegida — este servicio no vuelve a calcular precio ni
    bonificación, usa el importe ya neto tal cual viene.

    **Límite conocido, heredado del propio esquema, no de esta
    migración**: `Fcestad1` no guarda el lote/Despacho de cada renglón
    vendido (a diferencia de `RenglonEmision.nrodesp_elegido`, que sólo
    existe en memoria durante la Factura original) — la reversa de
    Stock ajusta el agregado (`Stock.STUnid`) pero NO un `Despacho`
    puntual. Misma limitación que ya tiene hoy la consulta de detalle de
    Factura Emitida."""

    TIPO_NC = 2
    LETRAS_SOPORTADAS = ("A", "B")
    # N/Créd.'A' / N/Créd.'B' — EmiFact.frm:1915-1924. Letra "C" (17)
    # queda fuera, mismo alcance que el resto del sistema.
    CODMOVS_POR_LETRA = {"A": "11", "B": "12"}

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)
        self.cuentas = CuentaCorrienteService(db)

    def emitir(
        self,
        cliente: Cliente,
        letra: str,
        punto_venta: int,
        numero_comprobante: int,
        renglones: list[RenglonEmision],
        total: "TotalFactura",
        motivo: int,
        usuario: str,
        comprobante_a_imputar: Ctascte,
        fecha: Optional[date] = None,
    ) -> ResultadoEmisionNotaCredito:
        if letra not in self.LETRAS_SOPORTADAS:
            raise ValueError(f"Letra '{letra}' no soportada — el alcance actual es sólo Factura A/B.")
        if not renglones:
            raise ValueError("La Nota de Crédito no tiene renglones para devolver.")
        if comprobante_a_imputar is None:
            raise ValueError("La Nota de Crédito necesita un comprobante contra el cual imputarse.")

        fecha = fecha or date.today()
        ahora = datetime.now()
        usuario6 = (usuario or "")[:6]
        codmovs = self.CODMOVS_POR_LETRA[letra]

        try:
            total_cantidad_unidades = Decimal("0")
            for item, renglon in enumerate(renglones, start=1):
                self._grabar_renglon(renglon, item, codmovs, punto_venta, numero_comprobante, cliente, letra, fecha, ahora, usuario6)
                total_cantidad_unidades += renglon.cantidad_unidades

            fcivavta = FcivaVta(
                FECHA=fecha,
                PTOVTA=punto_venta,
                CPBTE=numero_comprobante,
                LETRA=letra,
                TIPO=str(self.TIPO_NC),
                CLTE=cliente.CODIGO,
                NOMB=cliente.NOMB,
                PCIA=cliente.PCIA,
                CVTA=str(cliente.CVTA) if cliente.CVTA is not None else None,
                MOTI=str(motivo),
                CIVA=str(cliente.CIVA) if cliente.CIVA is not None else None,
                CUIT=(cliente.CUIT or "")[:14],
                VEND=cliente.VEND,
                ZONA=0,
                TOTCAN=int(total_cantidad_unidades),
                GRINS=total.neto_gravado,
                GRNOINS=Decimal("0"),
                IVAINS=total.iva,
                IVANOINS=Decimal("0"),
                PORCIB=Decimal("0"),
                TOTIB=Decimal("0"),
                EXENTO=Decimal("0"),
                TOTCOS=Decimal("0"),
                ITEMS=len(renglones),
                BON=total.descuento,
                NOIMPR="0",
                COMIS=Decimal("0"),
            )
            self.db.add(fcivavta)

            dias_vto = _dias_vencimiento_cond_venta(self.repos, cliente.CVTA)
            corr = self.repos.cliente().proximo_correlativo(cliente)
            ctascte = Ctascte(
                CLTE=cliente.CODIGO,
                FECHA=fecha,
                TIPO=self.TIPO_NC,
                PREFIJO=punto_venta,
                CPBTE=numero_comprobante,
                LETRA=letra,
                IMPUT1="0 ", IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
                IMPTE=total.total,
                DEBE=total.total,
                CVTA=str(cliente.CVTA) if cliente.CVTA is not None else None,
                BON=total.descuento,
                TIPO9="0",
                MOTI=str(motivo),
                FECVTO=fecha + timedelta(days=dias_vto),
                USUAR=usuario6,
            )
            self.db.add(ctascte)

            # Réplica de Graba() línea 2314 (`RgCCTE!imput1 = Corr`) — la
            # propia fila recién creada se marca a sí misma con su
            # correlativo, mismo patrón ya usado en `EmisionReciboService`/
            # `EmisionNotaCreditoService`.
            ctascte.IMPUT1 = str(corr)
            cliente.DEUDA = (cliente.DEUDA or Decimal("0")) - total.total
            cliente.USUARIO = usuario6
            cliente.FACTUAL = fecha

            resultado_imputacion = self.cuentas.imputar_pago(
                comprobante=comprobante_a_imputar,
                importe_aplicado=total.total,
                corr=corr,
                cpbte_recibo=numero_comprobante,
                usuario=usuario6,
                clte=cliente.CODIGO,
                fecha=fecha,
                commit=False,
                tipo_imputacion=CuentaCorrienteService.TIPO_IMPUTACION_NOTA_CREDITO,
            )

            totales = self._upsert_totales(fecha, letra, total, total_cantidad_unidades, usuario6)

            self.db.commit()
        except Exception:
            self.db.rollback()
            raise

        self.db.refresh(fcivavta)
        self.db.refresh(ctascte)
        self.db.refresh(totales)
        return ResultadoEmisionNotaCredito(
            fcivavta_id=fcivavta.id, ctascte_id=ctascte.id, totales_id=totales.id,
            imputacion_id=resultado_imputacion.imputacion_id, total=total.total,
        )

    def _grabar_renglon(
        self,
        renglon: RenglonEmision,
        item: int,
        codmovs: str,
        punto_venta: int,
        numero_comprobante: int,
        cliente: Cliente,
        letra: str,
        fecha: date,
        ahora: datetime,
        usuario6: str,
    ) -> None:
        cod1, cod2 = renglon.cod1, renglon.cod2

        # --- Stock (upsert, REVERSADO respecto de Factura) --------------
        stock = self.repos.stock().by_cod1_cod2(cod1, cod2)
        if stock is None:
            stock = Stock(
                COD1=cod1, COD2=cod2, STMIN=0, STMAX=0, STREP=0, DEP1=0, EST1=0, ESTT1=0,
                DEP2=0, EST2=0, ESTT2=0, STANT=0, PULG=Decimal("0"), MTR=Decimal("0"),
                STUNID=Decimal("0"), ENTMES=Decimal("0"), SALMES=Decimal("0"),
                AJEMES=Decimal("0"), AJSMES=Decimal("0"),
            )
            self.db.add(stock)

        # SUMA en vez de restar (EmiFact.frm:1864-1882, rama Else) —
        # devolución real, el artículo vuelve a estar disponible.
        if renglon.cantidad_unidades > 0:
            stock.STUNID = (stock.STUNID or Decimal("0")) + renglon.cantidad_unidades
            stock.ENTMES = (stock.ENTMES or Decimal("0")) + renglon.cantidad_unidades
        if renglon.mtr > 0:
            stock.STUNID = (stock.STUNID or Decimal("0")) + renglon.mtr
            stock.ENTMES = (stock.ENTMES or Decimal("0")) + renglon.mtr
        stock.FACTUAL = fecha
        stock.USUARIO = usuario6

        # --- MovStock -----------------------------------------------------
        self.db.add(
            MovStock(
                COD1=cod1, COD2=cod2, TIPO=codmovs, PTOVTA=punto_venta, CPBTE=numero_comprobante,
                ITEM=item, FECHA=ahora, CANT=renglon.cantidad_unidades, PULG=renglon.pulg,
                MTR=renglon.mtr, MILIM=renglon.milim, TELAS=renglon.telas, USUARIO=usuario6,
            )
        )

        # --- Fcestad1 (estadística por artículo, TIPO=2) ------------------
        self.db.add(
            Fcestad1(
                COD1=cod1, COD2=cod2, TIPO=self.TIPO_NC, LETRA=letra, PTOVTA=punto_venta,
                CPBTE=numero_comprobante, ITEM=item, CLTE=cliente.CODIGO, FECHA=fecha,
                GPTIPO="", GPDSD="", GPHST="", PULG=renglon.pulg, MTR=renglon.mtr,
                MILIM=renglon.milim, TELAS=renglon.telas, CANT=renglon.cantidad_unidades,
                PCOS=Decimal("0"), PVTA=renglon.precio_unitario, PESP=renglon.precio_unitario,
                BON=Decimal("0"), IVA=Decimal("0"), IMPTE=renglon.importe, TIPO9="0",
                USUARIO=usuario6,
            )
        )
        # Sin reversa de Despacho/lote puntual — ver docstring de la clase.

    def _upsert_totales(
        self, fecha: date, letra: str, total: "TotalFactura", total_cantidad_unidades: Decimal, usuario6: str
    ) -> Totales:
        """Réplica de `Graba()` Case 2 (NC, RESTA) — líneas 2126-2141,
        mismos campos que `EmisionFacturaService._upsert_totales` (Case
        1) pero restando en vez de sumando, y contando en `NCA`/`NCB` en
        vez de `FACA`/`FACB`."""
        totales = self.repos.totales().by_fecha(fecha)
        if totales is None:
            totales = Totales(FECHA=fecha)
            self.db.add(totales)

        def _sumar(campo: str, delta) -> None:
            actual = getattr(totales, campo) or 0
            setattr(totales, campo, actual + delta)

        _sumar("PVTA", -total.neto_gravado)
        _sumar("PESP", -total.neto_gravado)
        if letra == "A":
            _sumar("NCA", 1)
            _sumar("UNIDA", -int(total_cantidad_unidades))
            _sumar("PESPA", -total.neto_gravado)
        else:  # "B"
            _sumar("NCB", 1)
            _sumar("UNIDB", -int(total_cantidad_unidades))
            _sumar("PESPB", -total.neto_gravado)

        totales.FACTUAL = fecha
        totales.USUARIO = usuario6
        return totales


# ---------------------------------------------------------------------------
# EmisionReciboService
# ---------------------------------------------------------------------------


@dataclass
class AplicacionPago:
    """Una fila de la grilla "Pendiente" del Recibo (`DetRec.frm FG1`):
    cuánto de un comprobante con saldo (`CtascteRepository.
    pendientes_cobro()`) se está pagando, y si tiene descuento por
    pronto pago."""

    comprobante: Ctascte
    importe_aplicado: Decimal
    descuento: Decimal = Decimal("0")


@dataclass
class PagoCheque:
    """Un cheque de terceros recibido en pago — `DetPago.frm` Frame3
    "Detalle del Cheque", cuando el operador tipea un Nº de Cheque real
    (a diferencia de elegir una "Operación" del combo, ver
    `PagoRetencion`)."""

    nro_cheque: int
    banco: str
    fecha_emision: date
    fecha_vencimiento: date
    importe: Decimal
    a_la_orden: bool = False
    datos_adicionales: str = ""
    concepto: str = ""
    observaciones: str = ""
    tipo_cheque: int = 0  # 0=Común (Papel), 1=Electrónico — Combo2 de DetPago.frm


@dataclass
class PagoRetencion:
    """Anticipo/Retención/Tarjeta/Transferencia/Baja Incobrable —
    `Combo1` ("Operación") de `DetPago.frm`, `TIPREG` 1-9 en el mismo
    orden: 1-Retenc.Gan., 2-Retenc.IB, 3-Retenc.IVA, 4-Retenc.SUS,
    5-Dep./Transf., 6-Tarj.Déb/Cr., 7-Cheques Vs., 8-Cheq.Electr.,
    9-Baja Incobr."""

    tipreg: int
    importe: Decimal
    banco_o_dato: str = ""  # BcoSuc — sólo se usa realmente para Tarjeta (TIPREG=6)
    datos_adicionales: str = ""
    concepto: str = ""
    observaciones: str = ""


@dataclass
class ResultadoRecibo:
    numero: int
    ctascte_id: int
    total_pago: Decimal
    correlativo: int


class EmisionReciboService:
    """Persistencia atómica de un Recibo de Cuenta Corriente — equivalente
    a `EmiRec.frm Sub Graba()`.

    **Sin AFIP/CAE**: confirmado que un Recibo no es un comprobante
    fiscal electrónico — no hay una sola referencia a WSFE/CAE en los 5
    forms del circuito (`CabRec`/`DetRec`/`DetPago`/`PieRec`/`EmiRec`),
    sólo mueve Cuenta Corriente.

    **Decisiones confirmadas con el usuario (2026-08-15)**:
    - Orden Confirmar → Grabar → Imprimir: a diferencia del legacy
      (`EmiRec.frm Command1_Click` imprime en papel ANTES de pedir la
      confirmación real y grabar — si el operador contesta No/Cancelar,
      ya se imprimió un recibo que nunca quedó grabado), acá el PDF sólo
      se genera después de un `emitir_recibo()` exitoso — responsabilidad
      del llamador (`ReciboWindow`), no de este servicio.
    - El "Anticipo" (importe suelto sin comprobante puntual, `DetRec.frm`
      Command1/Text1) graba su propia fila `Ctascte` TIPO=5 ("Pago a
      Cuenta"), trazable aparte en la Consulta — el legacy lo mezclaba
      sin dejar rastro propio dentro del total del Recibo (TIPO=4). Esta
      fila nueva graba `DEBE=0` a propósito (no `DEBE=anticipo`): la
      queremos como ajuste de saldo puro, visible en el extracto, SIN
      sumarla a `CtascteRepository.pendientes_cobro()` de un Recibo
      futuro (que filtra `DEBE > 0`) — evita inventar una semántica de
      "aplicación futura" para TIPO=5 que el legacy nunca definió (el
      circuito real nunca crea esta fila hoy). Si en el uso real se
      necesita que el Anticipo sea aplicable a una factura futura vía
      `imputar_pago()`, avisar para revisar este valor.
    - El Nº de cheque es único: a diferencia del legacy (que silenciaba
      en silencio un cheque repetido, `ChequeRepository.by_nrocheq()`
      ya existía para esto), acá se avisa con `ValueError` en vez de
      perder el pago de un segundo cliente en silencio.

    **Réplica fiel del resto de `Graba()`**: numeración (`Parametro.
    NUME8`), correlativo `Cliente.CORR1` (ciclo 1-99, `ClienteRepository.
    proximo_correlativo()`), imputación por comprobante
    (`CuentaCorrienteService.imputar_pago()`, ya implementado — se llama
    una vez por fila con `commit=False` para que TODO el Recibo sea una
    única transacción), fila de descuento agregado (`Ctascte` TIPO=6,
    además del detalle por comprobante que ya deja `imputar_pago()` en
    `Imputacion`), `Cliente.DEUDA` actualizado por el total efectivamente
    cobrado, y persistencia de Cheques/MovimVS/Efectivo.
    """

    LETRA_RECIBO = "X"
    PREFIJO_RECIBO = 1
    TIPO_RECIBO = 4
    TIPO_DESCUENTO = 6
    TIPO_ANTICIPO = 5
    TIPO_EFECTIVO_MOVIMVS = 4  # Efectivo.TIPO / MovimVS.TIPO — siempre "Recibo"

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)
        self.cuentas = CuentaCorrienteService(db)

    def emitir_recibo(
        self,
        cliente: Cliente,
        numero: int,
        aplicaciones: list[AplicacionPago],
        anticipo: Decimal,
        importe_efectivo: Decimal,
        cheques: list[PagoCheque],
        retenciones: list[PagoRetencion],
        usuario: str,
        fecha: Optional[date] = None,
    ) -> ResultadoRecibo:
        """Réplica de `EmiRec.frm Sub Graba()`.

        Lanza `ValueError` si el Pago (efectivo+cheques+retenciones) no
        coincide EXACTO con lo que hay que cobrar (aplicado a
        comprobantes + anticipo − descuentos) — misma regla que gatea el
        botón "Grabar" en el legacy (`PieRec.frm Label304_Change`, sin
        tolerancia de redondeo), o si algún cheque ya está registrado.
        """
        fecha = fecha or date.today()
        usuario6 = (usuario or "")[:6]
        anticipo = Decimal(anticipo)
        importe_efectivo = Decimal(importe_efectivo)

        total_aplicado = sum((a.importe_aplicado for a in aplicaciones), Decimal("0"))
        total_descuento = sum((a.descuento for a in aplicaciones), Decimal("0"))
        a_pagar = total_aplicado + anticipo - total_descuento

        total_pago = (
            importe_efectivo
            + sum((c.importe for c in cheques), Decimal("0"))
            + sum((r.importe for r in retenciones), Decimal("0"))
        )

        if a_pagar <= 0:
            raise ValueError("El Recibo no tiene ningún importe a cobrar.")
        if total_pago != a_pagar:
            raise ValueError(
                f"El Pago (${format_decimal(total_pago)}) no coincide con lo que hay que cobrar "
                f"(${format_decimal(a_pagar)}) — la Diferencia debe ser $0 para poder grabar."
            )

        for pago_cheque in cheques:
            existente = self.repos.cheque().by_nrocheq(pago_cheque.nro_cheque)
            if existente is not None:
                raise ValueError(
                    f"El cheque Nº {pago_cheque.nro_cheque} ya está registrado "
                    f"(Cliente {existente.CLTE}, Recibo {existente.CPBING}) — no se puede reutilizar."
                )

        corr = self.repos.cliente().proximo_correlativo(cliente)

        try:
            parametro = self.repos.parametro().get_config()
            if parametro is not None:
                parametro.NUME8 = numero

            for aplicacion in aplicaciones:
                if aplicacion.importe_aplicado <= 0:
                    continue
                self.cuentas.imputar_pago(
                    comprobante=aplicacion.comprobante,
                    importe_aplicado=aplicacion.importe_aplicado,
                    corr=corr,
                    cpbte_recibo=numero,
                    usuario=usuario6,
                    clte=cliente.CODIGO,
                    fecha=fecha,
                    descuento=aplicacion.descuento,
                    commit=False,
                )

            ctascte_recibo = Ctascte(
                CLTE=cliente.CODIGO,
                FECHA=fecha,
                TIPO=self.TIPO_RECIBO,
                PREFIJO=self.PREFIJO_RECIBO,
                CPBTE=numero,
                LETRA=self.LETRA_RECIBO,
                IMPUT1=str(corr),
                IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
                DEBE=total_pago,
                IMPTE=total_pago,
                USUAR=usuario6,
                MOTI="0",
                TIPO9="0",
                CVTA="0",
                BON=Decimal("0"),
                FECVTO=fecha,
            )
            self.db.add(ctascte_recibo)

            if total_descuento > 0:
                self.db.add(
                    Ctascte(
                        CLTE=cliente.CODIGO,
                        FECHA=fecha,
                        TIPO=self.TIPO_DESCUENTO,
                        PREFIJO=self.PREFIJO_RECIBO,
                        CPBTE=numero,
                        LETRA=self.LETRA_RECIBO,
                        IMPUT1=str(corr),
                        IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
                        DEBE=total_descuento,
                        IMPTE=total_descuento,
                        USUAR=usuario6,
                        MOTI="0",
                        TIPO9="0",
                        CVTA="0",
                        BON=Decimal("0"),
                        FECVTO=fecha,
                    )
                )

            if anticipo > 0:
                self.db.add(
                    Ctascte(
                        CLTE=cliente.CODIGO,
                        FECHA=fecha,
                        TIPO=self.TIPO_ANTICIPO,
                        PREFIJO=self.PREFIJO_RECIBO,
                        CPBTE=numero,
                        LETRA=self.LETRA_RECIBO,
                        IMPUT1=str(corr),
                        IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
                        DEBE=Decimal("0"),  # ver docstring de la clase
                        IMPTE=anticipo,
                        USUAR=usuario6,
                        MOTI="0",
                        TIPO9="0",
                        CVTA="0",
                        BON=Decimal("0"),
                        FECVTO=fecha,
                    )
                )

            for pago_cheque in cheques:
                self.db.add(
                    Cheque(
                        NROCHEQ=pago_cheque.nro_cheque,
                        BCOSUC=pago_cheque.banco,
                        CPBCO="",
                        DATOSAD=pago_cheque.datos_adicionales,
                        TIPING=1,  # 1-Recibo CC
                        CPBING=numero,
                        CLTE=cliente.CODIGO,
                        CONCEP=pago_cheque.concepto,
                        IMPORTE=pago_cheque.importe,
                        PROPIO=0,
                        ORDEN=1 if pago_cheque.a_la_orden else 0,
                        TIPEGR=0,
                        CPBEGR=0,
                        DESTINO=None,
                        OBSERV=pago_cheque.observaciones,
                        ESTADO="1",  # 1-En Cartera
                        MONEDA="0",
                        COTIZ=Decimal("0"),
                        MOTRECH=0,
                        FECEMI=pago_cheque.fecha_emision,
                        FECVTO=pago_cheque.fecha_vencimiento,
                        FECEGR=None,
                        FECRECH=None,
                        NROOPE=0,
                        NROOPS=0,
                        USUARIO=usuario6,
                        FECING=fecha,
                    )
                )

            for retencion in retenciones:
                self.db.add(
                    MovimVS(
                        CLTE=cliente.CODIGO,
                        FECHA=fecha,
                        TIPO=str(self.TIPO_EFECTIVO_MOVIMVS),
                        CPBTE=numero,
                        TIPREG=str(retencion.tipreg),
                        IMPTE=retencion.importe,
                        BCOSUC=retencion.banco_o_dato,
                        DATOSAD=retencion.datos_adicionales,
                        CONCEP=retencion.concepto,
                        OBSERV=retencion.observaciones,
                    )
                )

            if importe_efectivo > 0:
                self.db.add(
                    Efectivo(
                        FECHA=fecha,
                        TIPO=self.TIPO_EFECTIVO_MOVIMVS,
                        CPBTE=numero,
                        IMPTE=importe_efectivo,
                    )
                )

            cliente.CORR1 = corr
            cliente.DEUDA = (cliente.DEUDA or Decimal("0")) - a_pagar
            cliente.USUARIO = usuario6
            cliente.FACTUAL = fecha

            self.db.commit()
        except Exception:
            self.db.rollback()
            raise

        self.db.refresh(ctascte_recibo)
        return ResultadoRecibo(
            numero=numero,
            ctascte_id=ctascte_recibo.id,
            total_pago=total_pago,
            correlativo=corr,
        )


# ---------------------------------------------------------------------------
# ChequeService
# ---------------------------------------------------------------------------


class ChequeService:
    """Consulta general de Cheques (todos los clientes) + acciones de
    Egreso/Borrado — migración de `VerCheq.frm` ("Consulta de Cheques").

    **Alcance ampliado sobre el legacy** (bajo criterio propio, sin
    impacto real: `Cheques` está confirmada vacía en `fcmenu_dev`, no
    hay datos reales en juego): el legacy sólo habilita filtrar
    "En Cartera"/"Egresados" (`Option1(3)`/`Option1(4)`, "Rechazados"/
    "Otros", están con `Enabled=False` en el propio `.frm`, sin ningún
    comentario que explique por qué) — acá se habilitan los 4 estados
    reales, ya modelados en `ChequeRepository`.

    **Borrado simplificado, no replicado tal cual**: el legacy pide
    retipear un código de 5 letras generado al azar (`BorraCheque`/
    `text1_Change`/`GrabaAnul`) antes de confirmar el DELETE físico —
    fricción real pero sin ningún valor de seguridad genuino (no es una
    contraseña, sólo un código visible en pantalla). Acá se usa el mismo
    patrón de confirmación Yes/No ya establecido en TODO el resto de la
    app para acciones destructivas (Supr en grillas, etc.) — decisión
    tomada por consistencia de UX, no hay datos reales que arriesgar.
    """

    # "Destino del Egreso" — lista fija real decodificada del recurso
    # binario `VerCheq.frx` (offset del `Combo3`, mismo mecanismo ya usado
    # para Condición IVA/Forma Ped. en fases anteriores). El índice
    # elegido es lo que se graba en `Cheque.TIPEGR` (`Combo3.ListIndex`
    # tal cual, EmiRec.frm/VerCheq.frm:1090) — OJO: la etiqueta "Destino"
    # del legacy está en el control equivocado: `Cheque.DESTINO` (el
    # campo de texto libre) en realidad guarda la "Descripción"
    # (`Text2(1)`, rotulada así en pantalla); acá se nombran ambos según
    # su significado real, no según la etiqueta engañosa del legacy.
    DESTINOS_EGRESO = [
        "Depósito", "Pago a Proveedores", "Cambio/Venta",
        "Cobro", "Pago Imptos.", "Caución", "Gastos Vs.",
    ]

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def buscar(self, estado: str, fecha_desde: date, limite: int = 50) -> list[Cheque]:
        """Réplica de `VerCheq.frm Sub DoVer2` — ver
        `ChequeRepository.by_estado_desde()`."""
        return self.repos.cheque().by_estado_desde(estado, fecha_desde, limite)

    def registrar_egreso(
        self,
        nrocheq: int,
        destino_index: int,
        cpbte_respaldo: int,
        descripcion: str,
        fecha_egreso: date,
    ) -> Cheque:
        """Marca un cheque en cartera como Egresado — réplica de
        `VerCheq.frm Sub Command4_Click` (líneas 1061-1107). Sólo válido
        para cheques `ESTADO='1'` (En Cartera); levanta `ValueError` si
        no existe o ya no está en cartera (mismo criterio de guarda que
        el resto de los servicios de emisión)."""
        cheque = self.repos.cheque().by_nrocheq(nrocheq)
        if cheque is None:
            raise ValueError(f"No existe el cheque Nº {nrocheq}")
        if cheque.ESTADO != self.repos.cheque().ESTADO_EN_CARTERA:
            raise ValueError(f"El cheque Nº {nrocheq} no está En Cartera.")
        if not (0 <= destino_index < len(self.DESTINOS_EGRESO)):
            raise ValueError("Destino de Egreso inválido.")

        cheque.TIPEGR = destino_index
        cheque.CPBEGR = cpbte_respaldo
        cheque.DESTINO = descripcion
        cheque.FECEGR = fecha_egreso
        cheque.ESTADO = self.repos.cheque().ESTADO_EGRESADO

        self.db.commit()
        self.db.refresh(cheque)
        return cheque

    def eliminar(self, nrocheq: int) -> None:
        """Borrado físico de un cheque — réplica de `VerCheq.frm Sub
        GrabaAnul` (líneas 1429-1465), sin el paso de re-tipear un código
        al azar (ver docstring de la clase)."""
        cheque = self.repos.cheque().by_nrocheq(nrocheq)
        if cheque is None:
            raise ValueError(f"No existe el cheque Nº {nrocheq}")
        self.db.delete(cheque)
        self.db.commit()


# ---------------------------------------------------------------------------
# EstadisticaVentasService
# ---------------------------------------------------------------------------


@dataclass
class MovimientoEstadistica:
    """Un renglón de `Fcestad1` con el Cliente ya resuelto — réplica de
    `ESTADIST.frm Sub CargaGrilla` ("Ventas de Artículos por Cliente")."""

    fecha: Optional[date]
    tipo_label: str
    letra: str
    comprobante: str  # "PtoVta-Cpbte" formateado, mismo padding del legacy
    cliente_codigo: Optional[int]
    cliente_nombre: str
    pulg: Decimal
    mtr: Decimal
    milim: int
    telas: int
    cantidad: Decimal
    precio_venta: Decimal
    importe: Decimal


@dataclass
class VentasSeccionRango:
    seccion: str
    fecha_desde: date
    fecha_hasta: date  # inclusive
    movimientos: list[MovimientoEstadistica]
    cantidad_movimientos: int
    total_importe: Decimal


@dataclass
class NodoAgrupado:
    """Un nivel de agrupamiento de `VentasArticuloAgrupadas` — réplica de
    una fila de subtotal del `Outline`/`SubTotal` real de `VTAXART.frm`.
    Sólo el último nivel de cada rama tiene `movimientos` (las filas
    reales); los niveles intermedios sólo tienen `hijos`."""

    etiqueta: str
    subtotales: dict[str, Decimal]
    hijos: list["NodoAgrupado"] = field(default_factory=list)
    movimientos: list[MovimientoEstadistica] = field(default_factory=list)


@dataclass(frozen=True)
class ConfigAgrupamiento:
    niveles: tuple[str, ...]
    columnas_subtotal: tuple[str, ...]


@dataclass
class VentasArticuloAgrupadas:
    seccion: str
    fecha_desde: date
    fecha_hasta_exclusiva: date
    cod2_desde: int
    cod2_hasta: int
    niveles_agrupamiento: tuple[str, ...]
    raiz: list[NodoAgrupado]
    cantidad_movimientos: int
    total_importe: Decimal


class EstadisticaVentasService:
    """Consultas de ventas por Sección/mes sobre `Fcestad1` — migración de
    `ESTADIST.frm` ("Ventas de Artículos por Cliente") y `VTAXART.frm`
    ("Ventas por Artículo", agrupada/subtotalizada por dimensión física).
    """

    # TIPO que suman al Total (vs. restan) — misma fórmula ya usada en
    # CtascteRepository.TIPOS_DEBE para el signo, acá aplicada sobre
    # Fcestad1.IMPTE en vez de Ctascte.IMPTE (ESTADIST.frm:450-454).
    TIPOS_SUMAN = {1, 3, 7}

    # Réplica exacta de ESTADIST.frm:457-468 / VTAXART.frm:602-613.
    ETIQUETAS_TIPO = {1: "Fact.", 2: "NCréd.", 3: "NDéb.", 9: "Anul."}

    # Agrupamiento de VTAXART.frm Sub CargaGrilla (líneas 683-739) —
    # hardcodeado por NOMBRE real de Sección en el legacy, confirmado con
    # el usuario (2026-08-16) replicar tal cual las 4 excepciones. Los
    # "niveles" son campos de `MovimientoEstadistica`, en el mismo orden
    # de agrupamiento externo→interno que `FG3.Col`/`ColSel` real; las
    # "columnas_subtotal" son los campos que se SUMAN en cada nivel — OJO:
    # son sumas CRUDAS de `Fcestad1.IMPTE` (lo que hace `FG3.SubTotal` de
    # verdad, sin el signo por TIPO), a diferencia del "Total Importe"
    # general de abajo, que SÍ es signado (mismo criterio que ESTADIST).
    CONFIG_AGRUPAMIENTO_POR_SECCION: dict[str, ConfigAgrupamiento] = {
        "GPS": ConfigAgrupamiento(niveles=("milim", "telas"), columnas_subtotal=("mtr", "importe")),
        "CT": ConfigAgrupamiento(niveles=("milim", "telas"), columnas_subtotal=("mtr", "importe")),
        "PL": ConfigAgrupamiento(niveles=("pulg", "milim", "telas"), columnas_subtotal=("mtr", "importe")),
        "SF": ConfigAgrupamiento(niveles=("mtr",), columnas_subtotal=("cantidad", "importe")),
    }
    CONFIG_AGRUPAMIENTO_DEFAULT = ConfigAgrupamiento(
        niveles=("pulg",), columnas_subtotal=("mtr", "cantidad", "importe")
    )

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    @staticmethod
    def fin_de_mes_exclusivo(fecha: date) -> date:
        """Primer día del mes siguiente al de `fecha` — límite superior
        EXCLUSIVO del rango de un mes.

        Réplica SIMPLIFICADA, con el MISMO resultado numérico, de
        `ESTADIST.frm Sub DoVer3` (líneas 344-361): el legacy calcula el
        último día del mes elegido (`ElDia`, con el caso especial de
        Febrero vía `DateAdd`) y compara `FECHA < FECHAHst`, PERO si la
        fecha elegida por el operador es justo ese último día
        (`FECHADsd = FECHAHst`), lo empuja al día 1 del mes siguiente
        para no excluirse a sí misma del rango. Usar directamente "día 1
        del mes siguiente al de `fecha`" como límite superior da
        exactamente el mismo resultado para CUALQUIER día de entrada
        (incluido ese caso especial) — confirmado por cálculo, no es una
        suposición de negocio distinta, sólo una implementación más simple.

        `VTAXART.frm` (líneas 505-508) tiene el MISMO cálculo pero SIN el
        caso especial de Febrero (`ElDia = 28` fijo, ignora años
        bisiestos — un 29/02 en un año bisiesto quedaría afuera de su
        propio mes) — acá se usa esta versión correcta para las dos
        pantallas, no se replica esa pérdida de datos real."""
        if fecha.month == 12:
            return date(fecha.year + 1, 1, 1)
        return date(fecha.year, fecha.month + 1, 1)

    def _armar_movimiento(self, r: Fcestad1) -> MovimientoEstadistica:
        cliente = self.repos.cliente().by_codigo(r.CLTE) if r.CLTE else None
        nombre_cliente = (cliente.NOMB or "").strip() if cliente else "*** Desconocido ***"
        return MovimientoEstadistica(
            fecha=r.FECHA,
            tipo_label=self.ETIQUETAS_TIPO.get(r.TIPO, "--"),
            letra=(r.LETRA or "").strip(),
            comprobante=f"{(r.PTOVTA or 0):04d}-{(r.CPBTE or 0):06d}",
            cliente_codigo=r.CLTE,
            cliente_nombre=nombre_cliente,
            pulg=r.PULG or Decimal("0"),
            mtr=r.MTR or Decimal("0"),
            milim=r.MILIM or 0,
            telas=r.TELAS or 0,
            cantidad=r.CANT or Decimal("0"),
            precio_venta=r.PVTA or Decimal("0"),
            importe=r.IMPTE or Decimal("0"),
        )

    def _total_signado(self, renglones: list[Fcestad1]) -> Decimal:
        total = Decimal("0")
        for r in renglones:
            importe = r.IMPTE or Decimal("0")
            total += importe if r.TIPO in self.TIPOS_SUMAN else -importe
        return total

    def ventas_seccion_por_rango(self, cod_seccion: str, fecha_desde: date, fecha_hasta: date) -> VentasSeccionRango:
        """Ventas de una Sección en un rango de fechas explícito
        (`fecha_hasta` INCLUSIVE) — base en `ESTADIST.frm Sub DoVer3`/
        `CargaGrilla`, con el rango de fechas AMPLIADO sobre el legacy
        (feedback del usuario, 2026-08-16): el original sólo dejaba
        elegir un día de inicio y mostraba automáticamente hasta fin de
        ESE mes (`Sub DoVer3`, líneas 344-361) — "ahora muestra sólo el
        mes seleccionado, es incorrecto". Acá el operador elige Desde Y
        Hasta libremente, sin la limitación de un solo mes calendario."""
        fecha_hasta_exclusiva = fecha_hasta + timedelta(days=1)
        renglones = self.repos.fcestad1().by_seccion_y_rango(cod_seccion, fecha_desde, fecha_hasta_exclusiva)

        return VentasSeccionRango(
            seccion=cod_seccion,
            fecha_desde=fecha_desde,
            fecha_hasta=fecha_hasta,
            movimientos=[self._armar_movimiento(r) for r in renglones],
            cantidad_movimientos=len(renglones),
            total_importe=self._total_signado(renglones),
        )

    # ------------------------------------------------------------------
    # VTAXART.frm — "Ventas por Artículo", agrupada por dimensión física
    # ------------------------------------------------------------------

    def config_agrupamiento(self, cod_seccion: str) -> ConfigAgrupamiento:
        return self.CONFIG_AGRUPAMIENTO_POR_SECCION.get(cod_seccion.strip(), self.CONFIG_AGRUPAMIENTO_DEFAULT)

    def ventas_articulo_agrupadas(
        self,
        cod_seccion: str,
        fecha_desde: date,
        cod2_desde: Optional[int] = None,
        cod2_hasta: Optional[int] = None,
    ) -> VentasArticuloAgrupadas:
        """Réplica de `VTAXART.frm Sub DoVer3`/`CargaGrilla` (líneas
        496-748), incluido el agrupamiento hardcodeado por Sección real
        (ver `CONFIG_AGRUPAMIENTO_POR_SECCION`)."""
        fecha_hasta = self.fin_de_mes_exclusivo(fecha_desde)
        renglones = self.repos.fcestad1().by_seccion_y_rango(cod_seccion, fecha_desde, fecha_hasta)

        limite_inf = cod2_desde if cod2_desde is not None else 0
        limite_sup = cod2_hasta if cod2_hasta is not None else 99999999
        renglones = [r for r in renglones if limite_inf <= _val_vb6(r.COD2) <= limite_sup]

        movimientos = [self._armar_movimiento(r) for r in renglones]
        config = self.config_agrupamiento(cod_seccion)
        raiz = self._agrupar(movimientos, config.niveles, config.columnas_subtotal)

        return VentasArticuloAgrupadas(
            seccion=cod_seccion,
            fecha_desde=fecha_desde,
            fecha_hasta_exclusiva=fecha_hasta,
            cod2_desde=limite_inf,
            cod2_hasta=limite_sup,
            niveles_agrupamiento=config.niveles,
            raiz=raiz,
            cantidad_movimientos=len(renglones),
            total_importe=self._total_signado(renglones),
        )

    @staticmethod
    def _subtotales(movimientos: list[MovimientoEstadistica], columnas: tuple[str, ...]) -> dict[str, Decimal]:
        return {campo: sum((getattr(m, campo) for m in movimientos), Decimal("0")) for campo in columnas}

    def _agrupar(
        self,
        movimientos: list[MovimientoEstadistica],
        niveles: tuple[str, ...],
        columnas_subtotal: tuple[str, ...],
    ) -> list[NodoAgrupado]:
        if not niveles:
            return []

        campo, resto = niveles[0], niveles[1:]
        grupos: dict[Decimal, list[MovimientoEstadistica]] = {}
        for m in movimientos:
            clave = getattr(m, campo)
            grupos.setdefault(clave, []).append(m)

        nodos = []
        for clave in sorted(grupos):  # flexSortNumericAscending
            miembros = grupos[clave]
            etiqueta = format_decimal(clave) if isinstance(clave, Decimal) else str(clave)
            if resto:
                nodos.append(
                    NodoAgrupado(
                        etiqueta=etiqueta,
                        subtotales=self._subtotales(miembros, columnas_subtotal),
                        hijos=self._agrupar(miembros, resto, columnas_subtotal),
                    )
                )
            else:
                nodos.append(
                    NodoAgrupado(
                        etiqueta=etiqueta,
                        subtotales=self._subtotales(miembros, columnas_subtotal),
                        movimientos=miembros,
                    )
                )
        return nodos


@dataclass
class FilaFacturaEmitida:
    """Una fila de la grilla principal de `VerFact.frm` ("Facturas
    Emitidas") — réplica de `Sub CargaGrilla` (líneas 1101-1199)."""

    fecha: Optional[date]
    tipo: int
    tipo_label: str
    letra: str
    ptovta: int
    cpbte: int
    comprobante: str  # "Letra Ptovta-Cpbte" formateado, mismo padding del legacy
    bruto: Decimal
    iva: Decimal
    neto: Decimal
    clte: Optional[int]
    cliente_nombre: str
    factura: FcivaVta  # fila real, para abrir el detalle sin releer


@dataclass
class ResultadoFacturasEmitidas:
    fecha_desde: date
    limite: int
    filas: list[FilaFacturaEmitida]
    total_bruto: Decimal
    total_iva: Decimal
    total_neto: Decimal


class FacturasEmitidasService:
    """Migración de `VerFact.frm` ("Facturas Emitidas") — consulta de
    comprobantes del subdiario de ventas (`FCIVAVTA`) por fecha, con
    detalle de Cliente e ítems reales (`Fcestad1`) por comprobante.

    **Sólo la consulta se migra, de sólo lectura.** La función "Anular"
    del legacy (`AnulaFactura`/`GrabaAnul`) NO se migra — decisión
    confirmada con el usuario (2026-08-16): no es compatible con un
    comprobante que ya tiene CAE real de AFIP (cancelar una Factura
    autorizada electrónicamente requiere emitir una Nota de Crédito
    contra AFIP, no poner los importes en cero localmente) y además el
    legacy nunca repone `Stock.STUnid`/`MovStock`/`Despachos.STOCK` al
    anular — sólo revierte `CtasCtes`/`Clientes.DEUDA`/`Fcestad1`. Queda
    pendiente para la fase de NC/ND (ya identificada como fuera del
    alcance actual, ver `EmisionFacturaService`)."""

    # TIPO que suman al Total (vs. restan) — Factura/ND/Factura Mostrador,
    # réplica exacta de `VerFact.frm:1139` (mismo criterio de signo que
    # `EstadisticaVentasService.TIPOS_SUMAN`, sobre `FcivaVta` en vez de
    # `Fcestad1`).
    TIPOS_SUMAN = {1, 3, 7}

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    @staticmethod
    def _tipo_int(tipo_raw: Optional[str]) -> Optional[int]:
        """`FcivaVta.TIPO` es texto (a diferencia de `Fcestad1.TIPO`,
        que es Integer) — tolera el mismo padding/blanco ya visto en
        otras columnas TIPO migradas desde Access."""
        if tipo_raw is None:
            return None
        texto = tipo_raw.strip()
        return int(texto) if texto else None

    @staticmethod
    def _formatear_comprobante(letra: Optional[str], ptovta: Optional[int], cpbte: Optional[int]) -> str:
        return f"{(letra or '').strip()} {(ptovta or 0):04d}-{(cpbte or 0):08d}"

    def listar(self, fecha_desde: date, limite: int) -> ResultadoFacturasEmitidas:
        """Comprobantes desde `fecha_desde` hasta fin de ESE mes (mismo
        cálculo que `EstadisticaVentasService.fin_de_mes_exclusivo`,
        reusado tal cual — `VerFact.frm Sub DoVer2` líneas 1027-1046
        hace el mismo cálculo de "fin de mes" con el mismo caso especial
        de Febrero), recortados a los primeros `limite`."""
        fecha_hasta_exclusiva = EstadisticaVentasService.fin_de_mes_exclusivo(fecha_desde)
        registros = self.repos.fciva_vta().emitidas_desde(fecha_desde, fecha_hasta_exclusiva, limite)

        filas: list[FilaFacturaEmitida] = []
        total_bruto = total_iva = total_neto = Decimal("0")
        for r in registros:
            tipo_int = self._tipo_int(r.TIPO)
            bruto = r.GRINS or Decimal("0")
            iva = r.IVAINS or Decimal("0")
            neto = bruto + iva
            signo = 1 if tipo_int in self.TIPOS_SUMAN else -1
            total_bruto += signo * bruto
            total_iva += signo * iva
            total_neto += signo * neto

            filas.append(
                FilaFacturaEmitida(
                    fecha=r.FECHA,
                    tipo=tipo_int or 0,
                    tipo_label=ETIQUETAS_TIPO_FCIVAVTA.get(tipo_int, "--"),
                    letra=(r.LETRA or "").strip(),
                    ptovta=r.PTOVTA or 0,
                    cpbte=r.CPBTE or 0,
                    comprobante=self._formatear_comprobante(r.LETRA, r.PTOVTA, r.CPBTE),
                    bruto=bruto,
                    iva=iva,
                    neto=neto,
                    clte=r.CLTE,
                    cliente_nombre=(r.NOMB or "").strip(),
                    factura=r,
                )
            )

        return ResultadoFacturasEmitidas(
            fecha_desde=fecha_desde,
            limite=limite,
            filas=filas,
            total_bruto=total_bruto,
            total_iva=total_iva,
            total_neto=total_neto,
        )


@dataclass
class FilaTotalDiario:
    """Una fila de la grilla principal de `TotFact.frm` ("Totales
    Diarios") — réplica de `Sub CargaGrilla` (líneas 1885-1949)."""

    fecha: date
    precio_venta: Decimal
    precio_costo: Decimal
    venta_real: Decimal  # PESP
    cantidad_facturas: int
    cantidad_notas_credito: int
    totales: Totales  # fila real, para abrir el detalle sin releer


@dataclass
class ResultadoTotalesDiarios:
    fecha_desde: date
    limite: int
    filas: list[FilaTotalDiario]
    total_precio_venta: Decimal
    total_precio_costo: Decimal
    total_venta_real: Decimal


@dataclass
class ResumenMensual:
    """Resultado de "Total Mensual" (`TotFact.frm Sub MuestraTotal`,
    botón/F4) — ver `TotalesDiariosService.resumen_mensual` para la
    corrección de rango confirmada con el usuario (2026-08-16)."""

    fecha_desde: date  # día 1 del mes
    fecha_hasta: date  # fecha elegida (inclusive)
    dias_habiles_mes: int
    dias_trabajados: int
    total_precio_venta: Decimal
    total_precio_costo: Decimal
    total_venta_real: Decimal
    total_unidades: int
    promedio_diario: Decimal
    proyeccion: Decimal
    unidad_promedio: Decimal


@dataclass
class DetalleTotalDia:
    """Detalle completo de un día (`TotFact.frm Sub MuestraDetalle`,
    Picture1) — expone la fila cruda de `Totales` (todos los importes/
    unidades/comprobantes reales, por canal Cta.Cte/Mostrador/
    Exportación) más los 6 promedios derivados (3 "Importes", 3
    "Unidades"), ya con la misma protección de división por cero del
    legacy (`If X <> 0 Then ... End If`, si no el promedio queda en 0).

    **Dos etiquetas del legacy corregidas, sin cambiar el cálculo**: los
    2 grupos "Unidades Promedio x Día" (uno en el panel Importes, otro
    en Unidades) NO calculan nada "por día" pese al nombre — uno es
    precio promedio POR UNIDAD vendida (`Impte/Cantidad`), el otro es
    cantidad promedio POR COMPROBANTE (`Unidades/Nº de facturas`). Se
    exponen acá como `precio_prom_*`/`unid_prom_*` con nombres que
    describen lo que realmente calculan."""

    totales: Totales
    precio_prom_cta_cte: Decimal
    precio_prom_mostrador: Decimal
    precio_prom_tipo_c: Decimal
    unid_prom_cta_cte: Decimal
    unid_prom_mostrador: Decimal
    unid_prom_tipo_c: Decimal


class TotalesDiariosService:
    """Migración de `TotFact.frm` ("Totales Diarios") sobre la tabla
    `Totales` (contador diario agregado, ya migrada en la fase del
    Facturador — ver `EmisionFacturaService`)."""

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    @staticmethod
    def _dias_habiles(anio: int, mes: int, hasta_dia: int) -> int:
        """Cuenta días Lunes a Sábado (excluye SÓLO Domingo, sin
        feriados) desde el día 1 hasta `hasta_dia` del mes/año dados —
        réplica de `TotFact.frm Function DiasHabiles` (líneas
        2109-2125), incluido el `If DiasHabiles = 0 Then DiasHabiles = 1`
        del legacy (evita división por cero, aunque en la práctica
        nunca da 0 para un mes real)."""
        dias = sum(1 for d in range(1, hasta_dia + 1) if date(anio, mes, d).weekday() != 6)  # 6 = domingo
        return dias or 1

    @staticmethod
    def _division_segura(numerador: Decimal, denominador: int) -> Decimal:
        """Réplica de `If X <> 0 Then Promed = .../X End If` — si el
        denominador es 0, el promedio queda en 0 (no `ZeroDivisionError`,
        mismo comportamiento silencioso del legacy)."""
        if not denominador:
            return Decimal("0")
        return numerador / Decimal(denominador)

    def detalle_dia(self, fecha: date) -> Optional[DetalleTotalDia]:
        """Detalle de un día puntual — réplica de `Sub MuestraDetalle`
        (líneas 1959-2057). Devuelve `None` si no hay fila de `Totales`
        para esa fecha (`If TOTL.Recordset.EOF Then Exit Sub`)."""
        r = self.repos.totales().by_fecha(fecha)
        if r is None:
            return None

        cero = Decimal("0")
        return DetalleTotalDia(
            totales=r,
            precio_prom_cta_cte=self._division_segura(
                (r.PESPA or cero) + (r.PESPB or cero), (r.UNIDA or 0) + (r.UNIDB or 0)
            ),
            precio_prom_mostrador=self._division_segura(
                (r.PESPMOSTA or cero) + (r.PESPMOSTB or cero), (r.UNMOSTA or 0) + (r.UNMOSTB or 0)
            ),
            precio_prom_tipo_c=self._division_segura(
                (r.PESPC or cero) + (r.PESPMOSTC or cero), (r.UNIDC or 0) + (r.UNMOSTC or 0)
            ),
            unid_prom_cta_cte=self._division_segura(
                Decimal((r.UNIDA or 0) + (r.UNIDB or 0)), (r.FACA or 0) + (r.FACB or 0)
            ),
            unid_prom_mostrador=self._division_segura(
                Decimal((r.UNMOSTA or 0) + (r.UNMOSTB or 0)), (r.MOSTA or 0) + (r.MOSTB or 0)
            ),
            unid_prom_tipo_c=self._division_segura(
                Decimal((r.UNIDC or 0) + (r.UNMOSTC or 0)), (r.FACC or 0) + (r.MOSTC or 0)
            ),
        )

    def listar(self, fecha_desde: date, limite: int) -> ResultadoTotalesDiarios:
        """Días desde `fecha_desde` hasta fin de ESE mes (inclusive,
        mismo cálculo de "fin de mes" que `FacturasEmitidasService`),
        recortados a los primeros `limite`."""
        fecha_hasta_inclusiva = EstadisticaVentasService.fin_de_mes_exclusivo(fecha_desde) - timedelta(days=1)
        registros = self.repos.totales().entre_fechas(fecha_desde, fecha_hasta_inclusiva, limite)

        filas: list[FilaTotalDiario] = []
        total_pvta = total_pcos = total_pesp = Decimal("0")
        for r in registros:
            pvta = r.PVTA or Decimal("0")
            pcos = r.PCOS or Decimal("0")
            pesp = r.PESP or Decimal("0")
            total_pvta += pvta
            total_pcos += pcos
            total_pesp += pesp
            filas.append(
                FilaTotalDiario(
                    fecha=r.FECHA,
                    precio_venta=pvta,
                    precio_costo=pcos,
                    venta_real=pesp,
                    # TotFact.frm:1917-1921 (TOTFAC/TOTNC) suman Tipo A/B +
                    # Mostrador + Exportación pero se olvidan del Tipo "C"
                    # (FACC/NCC) — el detalle del mismo día sí lo muestra
                    # aparte (Label115/116). Descuido objetivo del propio
                    # Sub, no una decisión de negocio — se completa acá.
                    cantidad_facturas=(r.FACA or 0) + (r.FACB or 0) + (r.FACC or 0)
                    + (r.EXPA or 0) + (r.MOSTA or 0) + (r.MOSTB or 0),
                    cantidad_notas_credito=(r.NCA or 0) + (r.NCB or 0) + (r.NCC or 0)
                    + (r.NCEXPA or 0) + (r.NCMOSTA or 0) + (r.NCMOSTB or 0),
                    totales=r,
                )
            )

        return ResultadoTotalesDiarios(
            fecha_desde=fecha_desde,
            limite=limite,
            filas=filas,
            total_precio_venta=total_pvta,
            total_precio_costo=total_pcos,
            total_venta_real=total_pesp,
        )

    def resumen_mensual(self, fecha_elegida: date) -> Optional[ResumenMensual]:
        """"Total Mensual" (`Sub MuestraTotal`, líneas 2059-2107).

        **Corrección de rango confirmada con el usuario (2026-08-16)**:
        el legacy suma `Totales` en `[fecha_elegida, fin_de_mes)` — el
        mismo rango "hacia adelante" que la grilla principal — pero
        divide ese total por "Días Trabajados" (días hábiles desde el
        día 1 del mes HASTA `fecha_elegida`, hacia atrás): dos rangos de
        fechas distintos mezclados en la misma división. Con la fecha
        por defecto (hoy), el numerador en la práctica sería casi
        siempre sólo el total de HOY (los días futuros del mes todavía
        no tienen fila en `Totales`), dividido por los días ya
        transcurridos — un promedio sin interpretación de negocio
        coherente. Acá se suma `[día 1 del mes, fecha_elegida]`
        (inclusive): "Totales a la Fecha" es literalmente eso, y
        "Promedio Diario"/"Proyección" quedan coherentes con "Días
        Trabajados"/"Días Hábiles" como divisor/multiplicador.

        **Segundo bug real corregido de paso, mismo hallazgo**: la
        fórmula de "Proyección" (`Label42`) tenía `Text1`/`Text2`
        invertidos (`(totesp / dias_habiles_mes) * dias_trabajados` en
        vez de `(totesp / dias_trabajados) * dias_habiles_mes`) — con
        esa fórmula la "proyección" siempre daba un número MENOR o
        igual al total acumulado a la fecha (nunca una extrapolación
        hacia adelante), que no tiene sentido como proyección de fin de
        mes. Acá: `Proyección = Promedio Diario × Días Hábiles del mes`.
        "Promedio Diario" (`totesp/dias_trabajados`) y "Unidad Prom."
        (`totuni/dias_trabajados`) ya estaban con la fórmula correcta en
        el legacy, sin cambios ahí.

        Devuelve `None` si no hay ninguna fila de `Totales` en el rango
        (mismo `If TOTL.Recordset.EOF Then Exit Sub` del legacy)."""
        inicio_mes = date(fecha_elegida.year, fecha_elegida.month, 1)
        registros = self.repos.totales().entre_fechas(inicio_mes, fecha_elegida)
        if not registros:
            return None

        total_pvta = sum((r.PVTA or Decimal("0") for r in registros), Decimal("0"))
        total_pcos = sum((r.PCOS or Decimal("0") for r in registros), Decimal("0"))
        total_pesp = sum((r.PESP or Decimal("0") for r in registros), Decimal("0"))
        total_unidades = sum(
            (
                (r.UNIDA or 0) + (r.UNIDB or 0) + (r.UNIDC or 0)
                + (r.UNMOSTA or 0) + (r.UNMOSTB or 0) + (r.UNMOSTC or 0)
                - (r.DEV or 0) - (r.MOSTDEV or 0)
            )
            for r in registros
        )

        ultimo_dia_mes = (EstadisticaVentasService.fin_de_mes_exclusivo(fecha_elegida) - timedelta(days=1)).day
        dias_habiles_mes = self._dias_habiles(fecha_elegida.year, fecha_elegida.month, ultimo_dia_mes)
        dias_trabajados = self._dias_habiles(fecha_elegida.year, fecha_elegida.month, fecha_elegida.day)

        promedio_diario = total_pesp / dias_trabajados
        proyeccion = promedio_diario * dias_habiles_mes
        unidad_promedio = Decimal(total_unidades) / dias_trabajados

        return ResumenMensual(
            fecha_desde=inicio_mes,
            fecha_hasta=fecha_elegida,
            dias_habiles_mes=dias_habiles_mes,
            dias_trabajados=dias_trabajados,
            total_precio_venta=total_pvta,
            total_precio_costo=total_pcos,
            total_venta_real=total_pesp,
            total_unidades=total_unidades,
            promedio_diario=promedio_diario,
            proyeccion=proyeccion,
            unidad_promedio=unidad_promedio,
        )


class ArregloCtaCteService:
    """Migración de `CargaCC.frm` ("Arreglos > Cuenta Corriente", menú
    oculto `Complementos`/`RECICLA1`/`ARRECCTE1`) — editor directo de
    UN renglón de `CtasCtes`, para corregir a mano un registro roto.

    **Sigue siendo una herramienta de "romper vidrio"**: no valida IVA,
    no revisa Stock, no exige coherencia con `FCIVAVta`/`Fcestad1` — el
    mismo alcance crudo del legacy (Alta/Cambio/Baja campo a campo,
    directo sobre la tabla), decisión confirmada con el usuario
    (2026-08-16), NO se le agregan las validaciones de negocio del
    Facturador/Recibo.

    **Corrección sobre el legacy, también confirmada**: `Clientes.DEUDA`
    se recalcula automáticamente desde `CtasCtes` (reusando
    `CtascteRepository.deuda_cliente()`, la misma fórmula real ya usada
    en toda la app) después de CADA Grabar/Borrar — el legacy nunca
    tocaba `DEUDA`, dejándola silenciosamente desincronizada.

    Réplica de `Sub Command2_Click`/`Command3_Click` (líneas 634-741):
    la clave real de búsqueda/guardado es CLTE+LETRA+TIPO+PREFIJO+CPBTE
    (la que usa `Grabacion`) — el legacy tenía una clave DISTINTA e
    incompleta (sin LETRA/PREFIJO) sólo en el lookup de
    `text104_lostfocus` (precarga de pantalla), inconsistencia menor sin
    impacto real (CPBTE ya es prácticamente único por CLTE+TIPO en la
    práctica) que no se replica — acá se usa siempre la clave completa."""

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def buscar(self, clte: int, letra: str, tipo: int, prefijo: int, cpbte: int) -> Optional[Ctascte]:
        for fila in self.repos.ctascte().by_comprobante(tipo, letra, prefijo, cpbte):
            if fila.CLTE == clte:
                return fila
        return None

    def grabar(
        self,
        *,
        clte: int,
        fecha: date,
        tipo: int,
        letra: str,
        prefijo: int,
        cpbte: int,
        imput1: str,
        imput2: str,
        imput3: str,
        debe: Decimal,
        impte: Decimal,
        fecvto: date,
    ) -> Ctascte:
        fila = self.buscar(clte, letra, tipo, prefijo, cpbte)
        if fila is None:
            # Defaults SÓLO para alta — réplica de CargaCC.frm:665-679:
            # en un Cambio (fila ya existe) estos campos NO se tocan.
            fila = Ctascte(
                CLTE=clte,
                FECHA=fecha,
                TIPO=tipo,
                LETRA=letra,
                PREFIJO=prefijo,
                CPBTE=cpbte,
                IMPUT4="0 ",  # slot libre real, ver nota de padding en EmisionFacturaService
                IMPUT5="0 ",
                IMPUT6="0 ",
                MOTI="0",
                TIPO9="0",
                CVTA="0",
                BON=Decimal("0"),
                USUAR="Arregl",  # marca real del legacy: identifica un registro parchado a mano
            )
            self.db.add(fila)
        fila.IMPUT1 = imput1
        fila.IMPUT2 = imput2
        fila.IMPUT3 = imput3
        fila.DEBE = debe
        fila.IMPTE = impte
        fila.FECVTO = fecvto
        self.db.flush()
        self._resincronizar_deuda(clte)
        self.db.commit()
        return fila

    def borrar(self, clte: int, letra: str, tipo: int, prefijo: int, cpbte: int) -> bool:
        fila = self.buscar(clte, letra, tipo, prefijo, cpbte)
        if fila is None:
            return False
        self.db.delete(fila)
        self.db.flush()
        self._resincronizar_deuda(clte)
        self.db.commit()
        return True

    def _resincronizar_deuda(self, clte: int) -> None:
        cliente = self.repos.cliente().by_codigo(clte)
        if cliente is None:
            return
        cliente.DEUDA = self.repos.ctascte().deuda_cliente(clte)["deuda"]


class ArregloSubdiarioService:
    """Migración de `CargaFC.frm` ("Arreglos > Subdiario Ventas", menú
    oculto) — editor directo de la CABECERA de un comprobante en
    `FcivaVta` (mismo alcance de "romper vidrio" que
    `ArregloCtaCteService`, ver su docstring — no valida IVA/Stock/
    coherencia con `Fcestad1`).

    **Nuance real sobre el resync de `Clientes.DEUDA`**: acá también se
    recalcula tras cada Grabar/Borrar por el mismo criterio confirmado
    con el usuario, pero `DEUDA` se deriva SIEMPRE de `CtasCtes` (nunca
    de `FCIVAVta` directamente) — así que este resync no corrige nada
    causado por ESTE editor puntual (que no toca `CtasCtes`), sólo dejy
    `DEUDA` consistente con el estado ACTUAL de `CtasCtes` por si venía
    desincronizada de antes. Documentado así para que quede claro que
    no es una protección real contra el uso de esta pantalla en
    particular — sólo un efecto secundario inofensivo.

    **Bug real encontrado y NO replicado** (`Sub Command2_Click`, línea
    1106): el legacy graba `FCIVAVta.CIVA = Combo3.ListIndex` (posición
    0-based del combo: 0-4), pero el propio `.frm` define
    `Combo3.ItemData` con los códigos reales 1-5 (Inscripto=1 ...
    Monotributo=5, idénticos a `Cliente.CIVA` en el resto de la app) sin
    usarlos nunca — evidencia clara de que el código quería el
    `ItemData` (1-5) y usó `ListIndex` (0-4) por error. Un comprobante
    reparado con este editor legacy quedaría con la Condición de IVA
    corrida en -1 (ej. "Cons.Final" grabado como si fuera "Resp.No
    Insc."), rompiendo silenciosamente la fórmula de Letra A/B
    (`FacturaService.letra_comprobante`, que compara `CIVA < 3`) para
    ESE comprobante puntual. Acá se usa directamente el código real
    (1-5)."""

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    def buscar(self, letra: str, tipo: int, ptovta: int, cpbte: int) -> Optional[FcivaVta]:
        return self.repos.fciva_vta().by_comprobante(str(tipo), letra, ptovta, cpbte)

    def grabar(
        self,
        *,
        tipo: int,
        letra: str,
        ptovta: int,
        cpbte: int,
        fecha: date,
        clte: int,
        nomb: str,
        pcia: str,
        cuit: str,
        grins: Decimal,
        ivains: Decimal,
        ivanoins: Decimal,
        exento: Decimal,
        bon: Decimal,
        porcib: Decimal,
        totib: Decimal,
        items: int,
        totcan: int,
        civa: int,
        vend: int,
        zona: int,
        cvta: int,
        moti: str,
    ) -> FcivaVta:
        fila = self.buscar(letra, tipo, ptovta, cpbte)
        if fila is None:
            # Defaults SÓLO para alta — réplica de CargaFC.frm:1077-1088.
            fila = FcivaVta(
                FECHA=fecha,
                TIPO=str(tipo),
                LETRA=letra,
                PTOVTA=ptovta,
                CPBTE=cpbte,
                TOTCOS=Decimal("0"),
                GRNOINS=Decimal("0"),
                COMIS=Decimal("0"),
                NOIMPR=0,
            )
            self.db.add(fila)
        fila.CLTE = clte
        fila.NOMB = nomb
        fila.PCIA = pcia
        fila.CUIT = cuit
        fila.GRINS = grins
        fila.IVAINS = ivains
        fila.IVANOINS = ivanoins
        fila.EXENTO = exento
        fila.BON = bon
        fila.PORCIB = porcib
        fila.TOTIB = totib
        fila.ITEMS = items
        fila.TOTCAN = totcan
        fila.CIVA = str(civa)  # código real 1-5, NO el ListIndex 0-4 del legacy (ver docstring)
        fila.VEND = vend
        fila.ZONA = zona
        fila.CVTA = cvta
        fila.MOTI = moti
        self.db.flush()
        if clte:
            self._resincronizar_deuda(clte)
        self.db.commit()
        return fila

    def borrar(self, letra: str, tipo: int, ptovta: int, cpbte: int) -> bool:
        fila = self.buscar(letra, tipo, ptovta, cpbte)
        if fila is None:
            return False
        clte = fila.CLTE
        self.db.delete(fila)
        self.db.flush()
        if clte:
            self._resincronizar_deuda(clte)
        self.db.commit()
        return True

    def _resincronizar_deuda(self, clte: int) -> None:
        cliente = self.repos.cliente().by_codigo(clte)
        if cliente is None:
            return
        cliente.DEUDA = self.repos.ctascte().deuda_cliente(clte)["deuda"]


@dataclass
class FilaIngresosBrutos:
    pcia: str
    pcia_nombre: str
    tipo: int
    gravado: Decimal
    iva: Decimal
    total: Decimal


@dataclass
class FilaSubdiarioComisiones:
    """Un renglón de `LisSubVtasCom` ("Subdiario Vtas. para Comisiones")
    — ya con el signo de NC aplicado y el caso Exento/Ushuaia separado
    en su propia columna, igual que el legacy."""

    fecha: Optional[date]
    letra: str
    comprobante: str
    clte: Optional[int]
    nombre: str
    cuit: str
    civa_label: str
    gravado: Decimal
    exento: Decimal
    iva: Decimal
    percepcion_ib: Decimal
    total: Decimal
    factura: FcivaVta


@dataclass
class FilaPercepcionArba:
    fecha: Optional[date]
    letra: str
    comprobante: str
    nombre: str
    cuit: str
    civa_label: str
    imponible: Decimal
    percepcion: Decimal
    linea_txt: str
    factura: FcivaVta


@dataclass
class ResultadoPercepcionesArba:
    filas: list[FilaPercepcionArba]
    total_imponible: Decimal
    total_percepcion: Decimal
    contenido_txt: str
    nombre_archivo_txt: str


@dataclass
class FilaComisionCobranza:
    clte: int
    nombre: str
    cpbte: int
    fecha: Optional[date]
    importe: Decimal
    es_nd_rechazada: bool


@dataclass
class ResultadoComisionesCobranza:
    vendedor: int
    filas: list[FilaComisionCobranza]
    total_cobrado: Decimal
    total_nd_rechazadas: int
    total_nd_rechazadas_importe: Decimal
    neto: Decimal  # (cobrado - nd_rechazadas) / (1 + alícuota IVA) — base de comisión


CIVA_LABELS_CORTOS = {1: "R.I.", 2: "RNI.", 3: "C.F.", 4: "Exto", 5: "Mon."}


class ListadosService:
    """Migración de `Listados.frm` ("Listados Varios", 5549 líneas, menú
    `LISTADO1` de `FCMENU.frm`) — 11 reportes reales migrados fielmente
    más 2 reconstruidos con el contenido real que su etiqueta promete
    (decisiones confirmadas con el usuario, 2026-08-16, ver
    `pyqt6_ui_progress.md`):

    - **"Deuda Pendiente"** (`SelLisFact`/`LisLisFact` original): el
      legacy NO filtraba deuda para nada — internamente se llamaba
      "Subdiario de Ventas" y listaba TODOS los comprobantes de un
      día/mes, duplicando la opción 3. Acá `deuda_pendiente()` es un
      reporte real de Facturas/ND impagas.
    - **"Lista de Precios"** (`SelPrecios`/`LisPrecios` original):
      consultaba una tabla `ArtiFalt` que no existe en ningún esquema
      de este proyecto, con SQL inválido (dos `WHERE` concatenados
      después de un `ORDER BY`) y campos (`NROFORD`, `ART.CALIF`,
      `STOK.STREP`) que no existen en ningún modelo — código muerto,
      probablemente copiado de otro sistema y nunca adaptado, no pudo
      haber funcionado nunca. Acá `lista_precios()` es lo que su
      nombre promete: Artículos con su Precio.

    **NO se migran** (confirmado, sin ambigüedad):
    - "Comisiones por Ventas" y "Deuda Vencida": `Sel.../Lis...`
      completamente vacíos en el legacy — nunca se implementaron.
    - "Comisiones por Cobr." (vieja, `SelComisCobra`/`LisComisCobra`):
      tiene `MsgBox "1"`/`"11"`/`"111"` de debug bloqueando el reporte
      — reemplazada por "Comis. x Cobr. (nvo.)", la única migrada acá
      (`comisiones_cobranzas()`).
    - `SelSubCobzas`/`LisSubCobzas` y `SelStkVal`/`LisStkVal`: código
      huérfano, sin ningún llamador en todo el archivo (confirmado con
      grep).

    `IMPRE.frm` (visor/impresor genérico compartido por todos los
    listados) tampoco se migra como pantalla — reemplazado por el
    patrón PDF ya usado en Factura/Recibo (`migration.pdf.
    generar_pdf_listado`)."""

    FILTRO_UNO = "uno"
    FILTRO_ZONA = "zona"
    FILTRO_TODOS = "todos"
    FILTRO_PROVINCIA = "provincia"

    def __init__(self, db: Session):
        self.db = db
        self.repos = RepositoryFactory(db)

    # ------------------------------------------------------------------
    def _clientes_por_filtro(self, filtro: str, valor=None) -> list[Cliente]:
        """Convención compartida "Uno Sólo / Por Zona / Todos" (a veces
        "Por Provincia") que repiten casi todos los listados de
        cliente — réplica de `Option2_Click` del legacy."""
        if filtro == self.FILTRO_UNO:
            cliente = self.repos.cliente().by_codigo(valor)
            return [cliente] if cliente else []
        if filtro == self.FILTRO_ZONA:
            return self.repos.cliente().by_zona(valor)
        if filtro == self.FILTRO_PROVINCIA:
            return self.repos.cliente().by_pcia(valor)
        return self.repos.cliente().todos_ordenados_por_nombre()

    # ------------------------------------------------------------------
    # 1 — Clientes (SelClientes/LisClientes)
    # ------------------------------------------------------------------
    def listado_clientes(self, filtro: str, valor=None) -> list[Cliente]:
        return self._clientes_por_filtro(filtro, valor)

    # ------------------------------------------------------------------
    # 2 — Lista de Precios — RECONSTRUIDO (ver docstring de la clase)
    # ------------------------------------------------------------------
    def lista_precios(self, cod_seccion: Optional[str] = None) -> list[Articulo]:
        if cod_seccion:
            return self.repos.articulo().by_seccion(cod_seccion)
        return self.repos.articulo().todos()

    # ------------------------------------------------------------------
    # 3 — Subdiario de Ventas (SelSubVtas/LisSubVtas)
    # ------------------------------------------------------------------
    def subdiario_ventas(self, fecha_desde: date, fecha_hasta_inclusiva: date) -> list[FcivaVta]:
        return self.repos.fciva_vta().by_fecha_rango(fecha_desde, fecha_hasta_inclusiva)

    # ------------------------------------------------------------------
    # 8 — Deuda Pendiente — RECONSTRUIDO (ver docstring de la clase);
    # también sirve de base a "Planilla de Cobranzas" (10, mismo
    # contenido real, otro layout impreso en el legacy).
    # ------------------------------------------------------------------
    def deuda_pendiente(self, filtro: str, valor=None) -> list[tuple[Cliente, list[Ctascte]]]:
        resultado = []
        for cliente in self._clientes_por_filtro(filtro, valor):
            pendientes = self.repos.ctascte().facturas_nd_impagas(cliente.CODIGO)
            if pendientes:
                resultado.append((cliente, pendientes))
        return resultado

    # ------------------------------------------------------------------
    # 11 — Estado de Cuenta (SelCCTE/LisCCTE)
    # ------------------------------------------------------------------
    def estado_cuenta(self, filtro: str, valor=None) -> list[tuple[Cliente, list[Ctascte]]]:
        resultado = []
        for cliente in self._clientes_por_filtro(filtro, valor):
            movimientos = self.repos.ctascte().by_cliente(cliente.CODIGO)
            if movimientos:
                resultado.append((cliente, movimientos))
        return resultado

    # ------------------------------------------------------------------
    # 12 — Saldos de Cta. Cte. (SelSaldos/LisSaldos)
    # ------------------------------------------------------------------
    def saldos_cta_cte(self, filtro: str, valor=None) -> list[dict]:
        """Reusa `CtascteRepository.saldos_todos_clientes()` (agregado
        en SQL) y filtra en Python por uno/zona/todos — los ~2.774
        clientes reales no ameritan una consulta agregada aparte por
        zona."""
        todos = self.repos.ctascte().saldos_todos_clientes()
        if filtro == self.FILTRO_TODOS:
            return todos
        if filtro == self.FILTRO_UNO:
            return [f for f in todos if f["codigo"] == valor]
        clientes_zona = {c.CODIGO for c in self.repos.cliente().by_zona(valor)}
        return [f for f in todos if f["codigo"] in clientes_zona]

    # ------------------------------------------------------------------
    # 4 — Subd. de Cobranzas AFIP (SelSubCobzasAFIP/LisSubCobzasAFIP)
    # ------------------------------------------------------------------
    def subdiario_cobranzas(self, fecha_desde: date, fecha_hasta_exclusiva: date) -> list[Ctascte]:
        return self.repos.ctascte().by_tipo_y_rango(4, fecha_desde, fecha_hasta_exclusiva)

    # ------------------------------------------------------------------
    # 5 — Ingresos Brutos (SelIngBrutos/LisIngBrutos)
    # ------------------------------------------------------------------
    def ingresos_brutos(self, fecha_desde: date, fecha_hasta_exclusiva: date) -> list[FilaIngresosBrutos]:
        """`GROUP BY Provincia, TIPO`, con el mismo signo de NC (TIPO=2
        resta) que el legacy — réplica de `LisIngBrutos:3054-3064`."""
        filas = []
        for r in self.repos.fciva_vta().resumen_por_provincia_y_tipo(fecha_desde, fecha_hasta_exclusiva):
            signo = -1 if r["tipo"] == 2 else 1
            gravado = r["grins"] * signo
            iva = r["ivains"] * signo
            filas.append(
                FilaIngresosBrutos(
                    pcia=(r["pcia"] or "").strip(),
                    pcia_nombre=nombre_provincia(r["pcia"]),
                    tipo=r["tipo"] or 0,
                    gravado=gravado,
                    iva=iva,
                    total=gravado + iva,
                )
            )
        return filas

    # ------------------------------------------------------------------
    # 15 — Subdiario Vtas. (Comisiones) (SelSubVtasCom/LisSubVtasCom)
    # ------------------------------------------------------------------
    def subdiario_ventas_comisiones(
        self, fecha_desde: date, fecha_hasta_exclusiva: date, vend: Optional[int] = None
    ) -> list[FilaSubdiarioComisiones]:
        """Réplica de `LisSubVtasCom:1801-1861`: NC (TIPO=2) resta
        (importes negativos), Exento (`CIVA=4`, caso real "Ushuaia" —
        Tierra del Fuego, mismo criterio de `FacturaService.
        letra_comprobante`) va a su propia columna en vez de "Gravado"."""
        filas = []
        for r in self.repos.fciva_vta().by_fecha_rango_vendedor(fecha_desde, fecha_hasta_exclusiva, vend):
            signo = -1 if (r.TIPO or "").strip() == "2" else 1
            civa = int(r.CIVA) if (r.CIVA or "").strip().isdigit() else None
            grins = (r.GRINS or Decimal("0")) * signo
            ivains = (r.IVAINS or Decimal("0")) * signo
            totib = (r.TOTIB or Decimal("0")) * signo

            if civa == 4:
                exento, gravado = grins, Decimal("0")
            else:
                exento, gravado = Decimal("0"), grins

            filas.append(
                FilaSubdiarioComisiones(
                    fecha=r.FECHA,
                    letra=(r.LETRA or "").strip(),
                    comprobante=f"{(r.PTOVTA or 0):04d}-{(r.CPBTE or 0):08d}",
                    clte=r.CLTE,
                    nombre=(r.NOMB or "").strip(),
                    cuit=(r.CUIT or "").strip(),
                    civa_label=CIVA_LABELS_CORTOS.get(civa, ""),
                    gravado=gravado,
                    exento=exento,
                    iva=ivains,
                    percepcion_ib=totib,
                    total=gravado + exento + ivains + totib,
                    factura=r,
                )
            )
        return filas

    # ------------------------------------------------------------------
    # 13 — Percepciones ARBA (SelPerARBA/LisPerARBA)
    # ------------------------------------------------------------------
    def percepciones_arba(self, fecha_desde: date, fecha_hasta_exclusiva: date) -> ResultadoPercepcionesArba:
        """Réplica de `LisPerARBA:2468-2558`: mismo signo de NC/Exento
        que `subdiario_ventas_comisiones`, más el caso especial real
        `TIPO=3 AND MOTI=' 2'` (Nota de Débito por Cheque Rechazado,
        código real confirmado contra `Fctabla1` MT='2') que
        recalcula el Gravado a partir del IVA (el legacy sólo graba el
        total+IVA de esa ND, sin un Gravado limpio propio).

        **El archivo `.txt` de presentación replica el formato fijo del
        legacy tal cual** (decisión confirmada con el usuario,
        2026-08-16) — la vigencia de ese formato ante ARBA hoy NO está
        verificada, hay que confirmarla con un contador antes de usarlo
        para una presentación real."""
        alicuota = FacturaService(self.db).alicuota_iva_inscripto()  # 0.21 real, no hardcodeado
        filas: list[FilaPercepcionArba] = []
        lineas_txt: list[str] = []
        total_imponible = Decimal("0")
        total_percepcion = Decimal("0")

        for r in self.repos.fciva_vta().con_percepcion_ib(fecha_desde, fecha_hasta_exclusiva):
            tipo_texto = (r.TIPO or "").strip()
            civa = int(r.CIVA) if (r.CIVA or "").strip().isdigit() else None
            signo = -1 if tipo_texto == "2" else 1

            grins = (r.GRINS or Decimal("0")) * signo
            ivains = (r.IVAINS or Decimal("0")) * signo
            totib = (r.TOTIB or Decimal("0")) * signo

            if civa == 4:
                imponible = Decimal("0")  # Exento/Ushuaia: no es base imponible de IIBB
            else:
                imponible = grins

            # ND por Cheque Rechazado (TIPO=3, MOTI=' 2'): el legacy no
            # graba un Gravado limpio para este tipo de ND — lo
            # recalcula a partir del IVA con la alícuota real.
            if tipo_texto == "3" and (r.MOTI or "").strip() == "2":
                imponible = (r.IVAINS or Decimal("0")) / alicuota

            cpbte_txt = f"{(r.PTOVTA or 0):04d}-{(r.CPBTE or 0):08d}"
            cuit_fmt = _formatear_cuit(r.CUIT)

            filas.append(
                FilaPercepcionArba(
                    fecha=r.FECHA,
                    letra=(r.LETRA or "").strip(),
                    comprobante=cpbte_txt,
                    nombre=(r.NOMB or "").strip(),
                    cuit=cuit_fmt,
                    civa_label=CIVA_LABELS_CORTOS.get(civa, ""),
                    imponible=imponible,
                    percepcion=totib,
                    linea_txt=self._linea_txt_arba(r, imponible, totib),
                    factura=r,
                )
            )
            total_imponible += imponible
            total_percepcion += totib
            lineas_txt.append(filas[-1].linea_txt)

        fecha_ref = fecha_desde
        nombre_archivo = f"AR-33703467909-{fecha_ref.year:04d}{fecha_ref.month:02d}0-7-LOTE1.txt"

        return ResultadoPercepcionesArba(
            filas=filas,
            total_imponible=total_imponible,
            total_percepcion=total_percepcion,
            contenido_txt="\n".join(lineas_txt),
            nombre_archivo_txt=nombre_archivo,
        )

    @staticmethod
    def _linea_txt_arba(r: FcivaVta, imponible: Decimal, percepcion: Decimal) -> str:
        """Réplica literal de `LisPerARBA:2557-2558` (ancho fijo real
        del archivo de presentación)."""
        tipo_texto = (r.TIPO or "").strip()
        letras_tipo = {"1": "F", "2": "C", "3": "D", "4": "R"}
        tcpbte = letras_tipo.get(tipo_texto, "")
        cuit_fmt = _formatear_cuit(r.CUIT)
        fecha_txt = r.FECHA.strftime("%d/%m/%Y") if r.FECHA else ""
        cpbte_txt = f"{(r.PTOVTA or 0):04d}{(r.CPBTE or 0):08d}"
        if tcpbte == "C":
            p_imponible = format(imponible, "013.2f").rjust(11)
            p_percep = format(percepcion, "011.2f").rjust(9)
        else:
            p_imponible = format(imponible, "013.2f").rjust(13)
            p_percep = format(percepcion, "011.2f").rjust(10)
        return f"{cuit_fmt}{fecha_txt}{tcpbte}{(r.LETRA or '').strip()}{cpbte_txt}{p_imponible}{p_percep}A"

    # ------------------------------------------------------------------
    # 14 — Comis. x Cobr. (nvo.) (SelComCobNVO/LisComCobNVO)
    # ------------------------------------------------------------------
    def comisiones_cobranzas(self, vend: int, fecha_desde: date, fecha_hasta_exclusiva: date) -> ResultadoComisionesCobranza:
        """Réplica de `LisComCobNVO:3408-3494`: Recibos/Pagos a Cta.
        (TIPO 4/5) de los clientes del Vendedor en el mes, más sus ND
        por Cheque Rechazado (TIPO=3, MOTI=' 2') listadas pero
        EXCLUIDAS del total cobrado — el "Neto" (base de comisión)
        resta esas ND y divide por `1 + alícuota IVA real` (no
        hardcodeado a 1.21 como el legacy)."""
        alicuota = FacturaService(self.db).alicuota_iva_inscripto()
        filas: list[FilaComisionCobranza] = []
        total_cobrado = Decimal("0")
        total_nd = Decimal("0")
        cant_nd = 0

        for cliente in self.repos.cliente().by_vendedor(vend):
            movimientos = self.repos.ctascte().cobranzas_y_nd_rechazadas(cliente.CODIGO, fecha_desde, fecha_hasta_exclusiva)
            for m in movimientos:
                es_nd = m.TIPO == 3
                importe = m.IMPTE or Decimal("0")
                filas.append(
                    FilaComisionCobranza(
                        clte=cliente.CODIGO, nombre=(cliente.NOMB or "").strip(),
                        cpbte=m.CPBTE or 0, fecha=m.FECHA, importe=importe, es_nd_rechazada=es_nd,
                    )
                )
                if es_nd:
                    total_nd += importe
                    cant_nd += 1
                else:
                    total_cobrado += importe

        neto = (total_cobrado - total_nd) / (Decimal("1") + alicuota)
        return ResultadoComisionesCobranza(
            vendedor=vend, filas=filas, total_cobrado=total_cobrado,
            total_nd_rechazadas=cant_nd, total_nd_rechazadas_importe=total_nd, neto=neto,
        )


def _formatear_cuit(cuit: Optional[str]) -> str:
    """"00-00000000-0" — réplica de `Format(Val(PCUIT), "00-00000000-0")`
    en varios listados."""
    digitos = "".join(ch for ch in (cuit or "") if ch.isdigit()).rjust(11, "0")
    return f"{digitos[0:2]}-{digitos[2:10]}-{digitos[10:11]}"


__all__ = [
    "TotalFactura",
    "FacturaService",
    "ImputacionResultado",
    "CuentaCorrienteService",
    "StockService",
    "EstadisticaFacturacion",
    "ClienteService",
    "ArticuloService",
    "TablaService",
    "RenglonEmision",
    "ResultadoEmision",
    "EmisionFacturaService",
    "AplicacionPago",
    "PagoCheque",
    "PagoRetencion",
    "ResultadoRecibo",
    "EmisionReciboService",
    "ChequeService",
    "MovimientoEstadistica",
    "VentasSeccionRango",
    "NodoAgrupado",
    "ConfigAgrupamiento",
    "VentasArticuloAgrupadas",
    "EstadisticaVentasService",
    "FilaFacturaEmitida",
    "ResultadoFacturasEmitidas",
    "FacturasEmitidasService",
    "FilaTotalDiario",
    "ResultadoTotalesDiarios",
    "ResumenMensual",
    "DetalleTotalDia",
    "TotalesDiariosService",
    "ArregloCtaCteService",
    "ArregloSubdiarioService",
    "FilaIngresosBrutos",
    "FilaSubdiarioComisiones",
    "FilaPercepcionArba",
    "ResultadoPercepcionesArba",
    "FilaComisionCobranza",
    "ResultadoComisionesCobranza",
    "ListadosService",
]
