"""`CotizacionVentaWindow` — Cotización, el comprobante SIN CAE del
Facturador (`CabFact.frm TipoFac=4`). NO confundir con "Cotización del
Dólar" (`CotizacionWindow`/`cotizacion_window.py`, ya migrada, tabla
`Cotizacion` — tipo de cambio, sin relación con esto).

**Análisis confirmado contra el legacy** (`CabFact.frm`+`EmiFact.frm`,
ver `pyqt6_ui_progress.md` para el detalle completo citado línea por
línea):

- Misma grilla de Detalle que Factura (`CabFact.frm:604`, "Case 1, 4"
  ambos muestran `DetFact.Show` — NO la de 3 líneas de concepto libre de
  NC), mismo cálculo de totales (`FacturaService.calcular_total`, misma
  fórmula).
- Letra fija "X" (no A/B derivada del Cliente) y punto de venta fijo
  "0001" (no el real de AFIP) — hardcodeados en el legacy
  (`EmiFact.frm:807`, `:845`).
- Numeración propia (`Parametro.NUME4`, ver `CotizacionVentaService`),
  independiente de Factura/NC/ND.
- **NUNCA persiste nada ni pide CAE** — confirmado en `EmiFact.frm:621-
  623`: `Graba()` (que es quien toca FCIVAVTA/Ctasctes/Stock/AFIP) se
  salta explícitamente `TipoFac = 4`. Es un documento puramente
  informativo: se genera el PDF y listo, no queda registro en el sistema
  más allá del número consumido.
- Por eso NO exige CUIT del Cliente (a diferencia de Facturador) ni
  verifica deuda — no es un acto de venta real todavía.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Optional

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.models import Cliente
from migration.pdf import DatosFacturaPDF, generar_pdf_factura
from migration.provincias import nombre_provincia
from migration.repository import RepositoryFactory
from migration.services import (
    CotizacionVentaService,
    FacturaService,
    RenglonEmision,
    TotalFactura,
    condicion_venta_texto,
)

from .cliente_busqueda_window import ClienteBusquedaWindow
from .decimals import format_decimal
from .detalle_grid import DetalleGrid
from .pdf_preview_dialog import PdfPreviewDialog
from .widgets import crear_recuadro_destacado, redimensionar_pct_pantalla, texto_contacto_cliente

LETRA_COTIZACION = "X"
PUNTO_VENTA_COTIZACION = 1  # "0001" fijo, EmiFact.frm:845 — no el real de AFIP

FORMAS_PEDIDO = ["1-Personal", "2-Telefónico", "3-Mail", "4-Fax", "5-Otro"]
FORMA_PEDIDO_DEFAULT_INDEX = 1


def _formatear_porcentaje(valor: Decimal) -> str:
    texto = format_decimal(valor)
    return texto[:-3] if texto.endswith(",00") else texto


class CotizacionVentaWindow(QMainWindow):
    """Ventana única (Cabecera/Detalle/Pie), calcada de `FacturadorWindow`
    pero sin AFIP/CAE ni persistencia — ver docstring del módulo."""

    TEXTO_BTN_ELEGIR_CLIENTE = "Elegir Cliente..."
    TEXTO_BTN_CAMBIAR_CLIENTE = "Cambiar Cliente"

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Cotización")
        redimensionar_pct_pantalla(self, 90, 85)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.factura_service = FacturaService(self.db)
        self.cotizacion_service = CotizacionVentaService(self.db)

        self.cliente_actual: Optional[Cliente] = None

        self._construir_ui()
        self._nueva_cotizacion()

    @property
    def renglones(self) -> list[RenglonEmision]:
        return self.grid_detalle.renglones

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_cabecera())
        layout.addWidget(self._armar_detalle(), stretch=1)
        layout.addWidget(self._armar_pie())

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QGroupBox:
        grupo = QGroupBox("Cabecera — Cotización (sin CAE, no es Factura)")
        layout = QVBoxLayout(grupo)

        fila_cliente = QHBoxLayout()
        self.lbl_cliente = QLabel("(sin cliente elegido)")
        fila_cliente.addWidget(self.lbl_cliente, stretch=1)
        self.btn_elegir_cliente = QPushButton(self.TEXTO_BTN_ELEGIR_CLIENTE)
        self.btn_elegir_cliente.clicked.connect(self._on_elegir_cliente)
        fila_cliente.addWidget(self.btn_elegir_cliente)
        self.btn_nueva = QPushButton("Nueva")
        self.btn_nueva.clicked.connect(self._nueva_cotizacion)
        fila_cliente.addWidget(self.btn_nueva)
        layout.addLayout(fila_cliente)

        self.lbl_cliente_contacto = QLabel("—")
        layout.addWidget(self.lbl_cliente_contacto)

        fila_datos = QHBoxLayout()
        fila_datos.addWidget(QLabel("Vendedor :"))
        self.combo_vendedor = QComboBox()
        fila_datos.addWidget(self.combo_vendedor)

        fila_datos.addWidget(QLabel("Forma Ped. :"))
        self.combo_forma_pedido = QComboBox()
        self.combo_forma_pedido.addItems(FORMAS_PEDIDO)
        self.combo_forma_pedido.setCurrentIndex(FORMA_PEDIDO_DEFAULT_INDEX)
        fila_datos.addWidget(self.combo_forma_pedido)

        self.chk_en_dolares = QCheckBox("En Dólares")
        self.chk_en_dolares.stateChanged.connect(self._on_en_dolares_cambiado)
        fila_datos.addWidget(self.chk_en_dolares)

        fila_datos.addStretch()
        self.lbl_proximo_numero = QLabel("Próx. Nº: —")
        fila_datos.addWidget(self.lbl_proximo_numero)
        layout.addLayout(fila_datos)

        self._cargar_vendedores()
        return grupo

    def _cargar_vendedores(self) -> None:
        self.combo_vendedor.clear()
        self.combo_vendedor.addItem("(sin vendedor)", None)
        for fila in self.repos.fctablas().by_ctab("VD"):
            cod = (fila.COD or "").strip()
            self.combo_vendedor.addItem(f"{cod} - {fila.DESCRI or ''}", int(cod) if cod.isdigit() else None)

    # ------------------------------------------------------------------
    # Detalle
    # ------------------------------------------------------------------
    def _armar_detalle(self) -> QGroupBox:
        grupo = QGroupBox("Detalle — Sección/**, código, cantidad (F2 = buscar Artículo, Supr = borrar renglón)")
        layout = QVBoxLayout(grupo)

        self.grid_detalle = DetalleGrid(
            self.repos,
            cliente_actual=lambda: self.cliente_actual,
            en_dolares=lambda: self.chk_en_dolares.isChecked(),
            cotizacion=self._cotizacion_actual,
            al_cambiar=self._recalcular_totales,
            parent=self,
        )
        layout.addWidget(self.grid_detalle, stretch=1)

        return grupo

    # ------------------------------------------------------------------
    # Pie
    # ------------------------------------------------------------------
    def _armar_pie(self) -> QGroupBox:
        grupo = QGroupBox("Totales")
        layout = QHBoxLayout(grupo)

        lbl_subtotal_titulo = QLabel("Subtotal:")
        lbl_subtotal_titulo.setStyleSheet("font-weight: bold;")
        self.lbl_subtotal = QLabel("$ 0,00")
        self.lbl_subtotal.setStyleSheet("font-weight: bold;")
        layout.addWidget(lbl_subtotal_titulo)
        layout.addWidget(self.lbl_subtotal)
        layout.addSpacing(16)

        layout.addWidget(QLabel("Desc.:"))
        self.lbl_descuento = QLabel("$ 0,00")
        layout.addWidget(self.lbl_descuento)

        layout.addWidget(QLabel("Neto Grav.:"))
        self.lbl_neto = QLabel("$ 0,00")
        layout.addWidget(self.lbl_neto)

        layout.addWidget(QLabel("IVA Insc.:"))
        self.lbl_iva = QLabel("$ 0,00")
        layout.addWidget(self.lbl_iva)
        layout.addSpacing(16)

        layout.addWidget(QLabel("Perc. IIBB:"))
        self.txt_porcentaje_iibb = QLineEdit("0")
        self.txt_porcentaje_iibb.setMaximumWidth(40)
        self.txt_porcentaje_iibb.editingFinished.connect(self._recalcular_totales)
        layout.addWidget(self.txt_porcentaje_iibb)
        layout.addWidget(QLabel("%"))
        self.lbl_percepcion_iibb = QLabel("$ 0,00")
        layout.addWidget(self.lbl_percepcion_iibb)
        layout.addStretch()

        recuadro_total, self.lbl_total = crear_recuadro_destacado("TOTAL:")
        layout.addWidget(recuadro_total)
        layout.addStretch()

        self.btn_generar = QPushButton("Generar Cotización")
        self.btn_generar.setStyleSheet("font-weight: bold; padding: 6px;")
        self.btn_generar.setEnabled(False)
        self.btn_generar.clicked.connect(self._on_generar)
        layout.addWidget(self.btn_generar)

        layout.addSpacing(12)
        btn_salir = QPushButton("Salir")
        btn_salir.clicked.connect(self.close)
        layout.addWidget(btn_salir)

        return grupo

    # ------------------------------------------------------------------
    # Cliente
    # ------------------------------------------------------------------
    def _on_elegir_cliente(self) -> None:
        dialogo = ClienteBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.cliente_elegido is None:
            return
        self.cliente_actual = self.repos.cliente().by_codigo(dialogo.cliente_elegido.CODIGO)
        self._refrescar_cliente()

    def _refrescar_cliente(self) -> None:
        cliente = self.cliente_actual
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.lbl_cliente_contacto.setText("—")
            self.btn_elegir_cliente.setText(self.TEXTO_BTN_ELEGIR_CLIENTE)
            return

        # A diferencia del Facturador, acá NO se chequea deuda ni
        # "Hab. Facturación" (`Cliente.CANAL`) — una Cotización no es un
        # acto de venta, el legacy tampoco lo hace (`BuscaUltima`/
        # `Form_Load` de `EmiFact.frm` no llaman a `ConDeuda.frm` para
        # `TipoFac=4`).
        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()}")
        self.lbl_cliente_contacto.setText(texto_contacto_cliente(cliente))
        self.btn_elegir_cliente.setText(self.TEXTO_BTN_CAMBIAR_CLIENTE)

        self._refrescar_proximo_numero()

        if cliente.VEND is not None:
            idx = self.combo_vendedor.findData(cliente.VEND)
            if idx >= 0:
                self.combo_vendedor.setCurrentIndex(idx)

        self._recalcular_totales()
        self.grid_detalle.foco_inicial()

    def _refrescar_proximo_numero(self) -> None:
        numero = self.cotizacion_service.proximo_numero()
        self.lbl_proximo_numero.setText(f"Próx. Nº: {PUNTO_VENTA_COTIZACION:04d}-{numero:08d} (Letra X)")

    # ------------------------------------------------------------------
    # Dólares / cotización de cambio
    # ------------------------------------------------------------------
    def _cotizacion_actual(self) -> Decimal:
        """Ver el docstring homónimo en `facturador_window.py` — mismo
        bug real corregido acá (2026-08-21): el checkbox "En Dólares" no
        debe decidir si se busca la cotización, sólo si se multiplica o
        se divide dentro de `resolver_precio_articulo`."""
        hoy = self.repos.cotizacion().by_fecha(date.today())
        if hoy is not None and hoy.DOLAR:
            return Decimal(hoy.DOLAR)
        ultima = self.repos.cotizacion().ultima()
        return Decimal(ultima.DOLAR) if ultima is not None and ultima.DOLAR else Decimal("1")

    def _on_en_dolares_cambiado(self) -> None:
        if self.renglones:
            QMessageBox.information(
                self,
                "Cotización",
                "Cambiar la moneda no reconvierte los renglones ya cargados — se recomienda \"Nueva\" antes de cambiar.",
            )
        self._recalcular_totales()

    # ------------------------------------------------------------------
    # Totales (Pie)
    # ------------------------------------------------------------------
    def _porcentajes_bonificacion(self, cod1: str) -> list[Decimal]:
        if self.cliente_actual is None:
            return []
        fila_dto = self.repos.dtoxclte().by_cliente_seccion(self.cliente_actual.CODIGO, cod1)
        if fila_dto is None:
            return []
        return [Decimal(p) for p in (fila_dto.DTO1, fila_dto.DTO2, fila_dto.DTO3, fila_dto.DTO4, fila_dto.DTO5) if p is not None]

    def _bonificacion_total(self) -> Decimal:
        total = Decimal("0")
        for renglon in self.renglones:
            porcentajes = self._porcentajes_bonificacion(renglon.cod1)
            if not porcentajes:
                continue
            total += self.factura_service.calcular_bonificacion_cascada(renglon.importe, porcentajes)
        return total

    def _descuentos_por_renglon(self) -> list[str]:
        resultado = []
        for renglon in self.renglones:
            porcentajes = self._porcentajes_bonificacion(renglon.cod1)
            resultado.append("+".join(_formatear_porcentaje(p) for p in porcentajes))
        return resultado

    def _porcentaje_iibb(self) -> Decimal:
        try:
            return Decimal(self.txt_porcentaje_iibb.text().strip() or "0")
        except InvalidOperation:
            return Decimal("0")

    def _calcular_total(self) -> TotalFactura:
        bruto = sum((r.importe for r in self.renglones), Decimal("0"))
        descuento = self._bonificacion_total()
        civa = self.cliente_actual.CIVA if self.cliente_actual else 3
        return self.factura_service.calcular_total(
            bruto=bruto, descuento=descuento, civa_cliente=civa or 3, porcentaje_iibb=self._porcentaje_iibb()
        )

    def _recalcular_totales(self) -> None:
        habilitada = self.factura_service.percepcion_iibb_habilitada()
        self.txt_porcentaje_iibb.setEnabled(habilitada)

        total = self._calcular_total()
        self.lbl_subtotal.setText(f"$ {format_decimal(total.bruto)}")
        self.lbl_descuento.setText(f"$ {format_decimal(total.descuento)}")
        self.lbl_neto.setText(f"$ {format_decimal(total.neto_gravado)}")
        self.lbl_iva.setText(f"$ {format_decimal(total.iva)}")
        self.lbl_percepcion_iibb.setText(f"$ {format_decimal(total.percepcion_iibb)}")
        self.lbl_total.setText(f"$ {format_decimal(total.total)}")
        self.btn_generar.setEnabled(total.total > 0 and self.cliente_actual is not None)

    # ------------------------------------------------------------------
    # Nueva cotización
    # ------------------------------------------------------------------
    def _nueva_cotizacion(self) -> None:
        self.cliente_actual = None
        self.grid_detalle.reiniciar()
        self.chk_en_dolares.setChecked(False)
        self.combo_forma_pedido.setCurrentIndex(FORMA_PEDIDO_DEFAULT_INDEX)
        self.txt_porcentaje_iibb.setText("0")
        self._refrescar_cliente()
        self._recalcular_totales()
        QTimer.singleShot(0, self._on_elegir_cliente)

    # ------------------------------------------------------------------
    # Generar (PDF, sin AFIP ni persistencia — ver docstring del módulo)
    # ------------------------------------------------------------------
    def _on_generar(self) -> None:
        if self.cliente_actual is None:
            QMessageBox.warning(self, "Cotización", "Elegí un Cliente.")
            return
        if not self.renglones:
            QMessageBox.warning(self, "Cotización", "Agregá al menos un renglón.")
            return

        total = self._calcular_total()
        fecha = date.today()
        descuentos_renglones = self._descuentos_por_renglon()

        # Consume el número acá, recién al confirmar generar — no en cada
        # vistazo (ver docstring de `CotizacionVentaService.
        # confirmar_numero`, diferencia deliberada con el legacy).
        try:
            numero = self.cotizacion_service.confirmar_numero()
        except ValueError as exc:
            QMessageBox.critical(self, "Cotización", str(exc))
            return

        try:
            datos_pdf = DatosFacturaPDF(
                letra=LETRA_COTIZACION,
                punto_venta=PUNTO_VENTA_COTIZACION,
                numero=numero,
                fecha=fecha,
                cliente_codigo=self.cliente_actual.CODIGO,
                cliente_nombre=(self.cliente_actual.NOMB or "").strip(),
                cliente_cuit=self.cliente_actual.CUIT or "",
                cliente_civa=self.cliente_actual.CIVA or 0,
                cliente_domicilio=(self.cliente_actual.DIR or "").strip(),
                cliente_localidad=(self.cliente_actual.LOC or "").strip(),
                cliente_cp=(self.cliente_actual.CP or "").strip(),
                cliente_provincia=nombre_provincia(self.cliente_actual.PCIA),
                condicion_venta=condicion_venta_texto(self.repos, self.cliente_actual.CVTA),
                renglones=list(self.renglones),
                descuentos_renglones=descuentos_renglones,
                total=total,
                en_dolares=self.chk_en_dolares.isChecked(),
                cotizacion=self._cotizacion_actual(),
                es_cotizacion=True,
            )
            ruta_pdf = generar_pdf_factura(datos_pdf)
        except Exception as exc:  # noqa: BLE001 — el número YA se consumió, no se revierte
            QMessageBox.critical(
                self,
                "Cotización",
                f"La Cotización Nº {PUNTO_VENTA_COTIZACION:04d}-{numero:08d} quedó numerada pero el PDF falló:\n{exc}",
            )
            return

        PdfPreviewDialog(
            ruta_pdf,
            titulo=f"Cotización {PUNTO_VENTA_COTIZACION:04d}-{numero:08d}",
            parent=self,
        ).exec()
        self._nueva_cotizacion()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        if self.renglones:
            respuesta = QMessageBox.question(
                self,
                "Cotización",
                "Hay una Cotización cargada — se va a perder el ingreso.\n¿Desea salir igual?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if respuesta != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.db.close()
        super().closeEvent(event)
