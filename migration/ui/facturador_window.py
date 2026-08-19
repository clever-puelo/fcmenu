"""`FacturadorWindow` — ventana única del Facturador (Cabecera/Detalle/
Pie), equivalente conceptual a `CabFact.frm`+`DetFact.frm`+`PieFact.frm`+
`EmiFact.frm` del legacy.

**Arquitectura confirmada con el usuario** (ver `pyqt6_ui_progress.md`):
una sola ventana cohesiva con tres secciones, NO los 4 formularios MDI
sin bordes que el legacy apila para simular una sola pantalla.

**Alcance confirmado**: sólo Factura, Letra A/B (NC/ND/Remito/Cotización/
Exportación quedan para una fase posterior). El circuito completo:

1. Cabecera: elegir Cliente (`ClienteBusquedaWindow` en modo selección,
   se abre solo al arrancar/"Nueva" — igual que el legacy, `FCMENU.frm`:
   `CabFact.Show` siempre seguido de `BusClte.Show`), Vendedor, Forma de
   Pedido, En Dólares. La Letra se deriva sola
   (`FacturaService.letra_comprobante`), no se elige a mano.
2. Detalle: **grilla interactiva** (`DetalleGrid`, estilo Excel/FlexGrid
   real — decisión del usuario 2026-08-16, ver docstring de
   `detalle_grid.py`), no diálogos de alta/edición/baja. Se tipea directo
   en la grilla: Sección → segmentos → cantidades, F2 o Enter en blanco
   abre el selector de Artículo, auto-popup de Lote/Despacho.
3. Pie: totales en vivo (`FacturaService.calcular_total`), % de
   percepción IIBB manual (gateado por `Parametro.MCAIB`, igual que el
   legacy).
4. Emitir: pide el CAE (`NumeracionYCAEProvider.solicitar_cae` — el
   stub por defecto hasta que haya credenciales reales de AFIP) y recién
   con el CAE aprobado persiste todo (`EmisionFacturaService.
   emitir_factura`). Si AFIP rechaza, NO se persiste nada — ni Stock ni
   Ctasctes se tocan sin CAE aprobado (a diferencia del circuito de
   papel del legacy, fuera de alcance).
"""

from __future__ import annotations

import os
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

from migration.afip import (
    NumeracionYCAEProvider,
    ResultadoCAE,
    codigo_afip,
    condicion_iva_receptor,
    crear_afip_provider,
    generar_qr_afip,
    punto_venta_por_tipo,
)
from migration.db import get_session
from migration.models import Cliente
from migration.pdf import DatosFacturaPDF, generar_pdf_factura
from migration.repository import RepositoryFactory
from migration.services import CuentaCorrienteService, EmisionFacturaService, FacturaService, RenglonEmision, TotalFactura

from .cliente_busqueda_window import ClienteBusquedaWindow
from .decimals import format_decimal
from .detalle_grid import DetalleGrid
from .nota_cliente_dialog import NotaClienteDialog
from .pdf_preview_dialog import PdfPreviewDialog
from .widgets import crear_recuadro_destacado, redimensionar_pct_pantalla

# CUIT del emisor — hardcodeado también en el legacy (`ConectaAFIP()`,
# `WSFE.Cuit = "33703467909"`, empresa "ALESTEL SRL" ya confirmada en la
# fase de Parámetros). Pendiente de mover a config real cuando se
# conecten credenciales — ver `migration/afip.py`.
CUIT_EMISOR_DEFAULT = "33703467909"

# Combo "Forma Ped." — 5 opciones fijas extraídas del recurso binario
# `CabFact.frx` (offset del `Combo4`, decodificado 2026-08-16): no vienen
# de una tabla, es una lista fija de diseño. `TipoVta = ListIndex + 1`
# se graba tal cual en `FcivaVta.MOTI` (CabFact.frm:721,690-692).
FORMAS_PEDIDO = ["1-Personal", "2-Telefónico", "3-Mail", "4-Fax", "5-Otro"]
FORMA_PEDIDO_DEFAULT_INDEX = 1  # "2-Telefónico" (CabFact.frm Form_Load: Combo4.ListIndex = 1)


class FacturadorWindow(QMainWindow):
    TEXTO_BTN_ELEGIR_CLIENTE = "Elegir Cliente..."
    TEXTO_BTN_CAMBIAR_CLIENTE = "Cambiar Cliente"

    def __init__(self, parent: QWidget | None = None, *, afip: Optional[NumeracionYCAEProvider] = None):
        super().__init__(parent)
        self.setWindowTitle("Facturador")
        # % de la pantalla real, no píxeles fijos (convención de sistema
        # #11, feedback del usuario, 2026-08-19) — valor sugerido, el
        # usuario va a probar y avisar cuáles ajustar.
        redimensionar_pct_pantalla(self, 90, 85)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.factura_service = FacturaService(self.db)
        self.emision_service = EmisionFacturaService(self.db)
        self.cuenta_corriente_service = CuentaCorrienteService(self.db)
        self.afip: NumeracionYCAEProvider = afip if afip is not None else crear_afip_provider()

        self.cliente_actual: Optional[Cliente] = None
        self.tiene_nota_cliente = False
        # Ya no se muestra en la cabecera (feedback del usuario,
        # 2026-08-18: "Quitar usuario de la cabecera") pero se sigue
        # necesitando para auditoría (`EmisionFacturaService.
        # emitir_factura(usuario=...)`).
        self._usuario = os.environ.get("USERNAME", "SISTEMA")[:6]

        self._construir_ui()
        self._nueva_factura()

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
        grupo = QGroupBox("Cabecera — Factura")
        layout = QVBoxLayout(grupo)

        # "Nueva" subido a esta fila (antes vivía abajo) y "Nota Clte."
        # corrido a su izquierda — quedan los dos juntos en la misma
        # línea, al extremo derecho (feedback del usuario, 2026-08-18,
        # tercera ronda: "Subir botón nueva y correr el botón nota clte
        # a la izquierda... El botón se superpone ahora" — el apilado
        # Próx.Nº/Letra + botón de la ronda anterior quedaba muy
        # apretado en `fila_datos`, junto con Vendedor/Forma Ped./
        # Dólares).
        fila_cliente = QHBoxLayout()
        self.btn_elegir_cliente = QPushButton(self.TEXTO_BTN_ELEGIR_CLIENTE)
        self.btn_elegir_cliente.clicked.connect(self._on_elegir_cliente)
        fila_cliente.addWidget(self.btn_elegir_cliente)
        self.lbl_cliente = QLabel("(sin cliente elegido)")
        fila_cliente.addWidget(self.lbl_cliente, stretch=1)
        self.btn_nota_cliente = QPushButton("Nota Clte.")
        self.btn_nota_cliente.setEnabled(False)
        self.btn_nota_cliente.clicked.connect(self._on_nota_cliente)
        fila_cliente.addWidget(self.btn_nota_cliente)
        self.btn_nueva = QPushButton("Nueva")
        self.btn_nueva.clicked.connect(self._nueva_factura)
        fila_cliente.addWidget(self.btn_nueva)
        layout.addLayout(fila_cliente)

        # Abajo sólo queda "Próx. Nº" (+ "Letra", en la misma línea) —
        # ya no compite por espacio con el botón "Nueva".
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
        # El label "Letra" aparte se sacó (feedback del usuario,
        # 2026-08-19: "Sacar la letra porque se repite, ya está en el
        # nro de factura") — la letra ya queda visible acá mismo, entre
        # paréntesis junto al número (ver `_refrescar_proximo_numero`).
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
        # Rediseño compacto en UNA sola línea horizontal (feedback del
        # usuario, 2026-08-19: "el panel inferior de totales hacerlo más
        # corto juntando líneas, para dar más espacio al detalle" — antes
        # eran 2 líneas; la reducción de las 2 a 1 le devuelve esa
        # segunda línea de alto al Detalle, que es donde el usuario
        # necesita ver más renglones a la vez).
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

        # TOTAL centrado en la línea, dentro de un recuadro (feedback
        # del usuario, 2026-08-18, tercera ronda: "El total colocarlo en
        # la misma línea pero en el centro y dentro de un panel/
        # recuadro") — el `addStretch()` de cada lado lo deja centrado
        # entre Percepción IIBB (izquierda) y Emitir/Salir (derecha).
        recuadro_total, self.lbl_total = crear_recuadro_destacado("TOTAL:")
        layout.addWidget(recuadro_total)
        layout.addStretch()

        self.btn_emitir = QPushButton("Emitir")
        self.btn_emitir.setStyleSheet("font-weight: bold; padding: 6px;")
        # Deshabilitado hasta que haya Importe (feedback del usuario,
        # 2026-08-15) — se habilita en `_recalcular_totales`.
        self.btn_emitir.setEnabled(False)
        self.btn_emitir.clicked.connect(self._on_emitir)
        layout.addWidget(self.btn_emitir)

        # Botón "Salir" (pedido del usuario, 2026-08-17) — la guarda de
        # "hay algo cargado, se va a perder" vive en `closeEvent` (no acá
        # directo) para cubrir también el cierre nativo de la subventana,
        # no sólo este botón.
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
        # `dialogo.cliente_elegido` pertenece a la sesión propia de
        # ClienteBusquedaWindow (`get_session()` propio) — NO a `self.db`.
        # Bug real encontrado probando el circuito completo: si se usa
        # ese objeto tal cual, `EmisionFacturaService` termina mutando
        # `Cliente.DEUDA` sobre un objeto ajeno a su sesión y el
        # `commit()` no persiste el cambio (se pierde en silencio). Se
        # vuelve a buscar por CODIGO en la sesión propia de esta ventana.
        self.cliente_actual = self.repos.cliente().by_codigo(dialogo.cliente_elegido.CODIGO)
        self._refrescar_cliente()

    def _refrescar_cliente(self) -> None:
        cliente = self.cliente_actual
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.btn_nota_cliente.setEnabled(False)
            self.tiene_nota_cliente = False
            self.btn_nota_cliente.setStyleSheet("")
            self.btn_elegir_cliente.setText(self.TEXTO_BTN_ELEGIR_CLIENTE)
            return

        # Aviso de Deuda (ConDeuda.frm, vía BusClte.frm Function
        # TieneDeuda()) — sólo acá, al elegir cliente para una Factura
        # nueva (no en Recibo/Cta.Cte, que reusan el mismo buscador de
        # Cliente pero no este chequeo). Si el cliente NO tiene
        # "Hab. Facturación" (Cliente.CANAL=9, checkbox ya migrado en
        # ClienteDetalleDialog), la selección se aborta del todo — igual
        # que el `Exit Sub` real del legacy, no es sólo un aviso.
        if self.cuenta_corriente_service.tiene_deuda_vieja(cliente.CODIGO):
            # El legacy (ConDeuda.frm) tiene el texto "90 días" hardcodeado
            # en el Label, sin importar el Parametro.NUME20 real configurado
            # — acá NO se replica ese hardcodeo (bug real: la base real
            # tiene NUME20=60, no 90), se arma el texto con el valor
            # configurado de verdad.
            dias = self.repos.parametro().get_config()
            dias_txt = str(dias.NUME20) if dias and dias.NUME20 is not None else "90"
            QMessageBox.warning(
                self, "Aviso de Deuda", f"Cliente con Deuda Mayor a {dias_txt} días\n\nConsulte"
            )
            if cliente.CANAL != 9:
                self.cliente_actual = None
                self._refrescar_cliente()
                return

        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()} — CUIT {cliente.CUIT or 's/d'}")
        self.btn_nota_cliente.setEnabled(True)
        self.btn_elegir_cliente.setText(self.TEXTO_BTN_CAMBIAR_CLIENTE)

        # Feedback del usuario (2026-08-15): antes de pasar al Detalle, se
        # verifica si el Cliente tiene Nota guardada y, si la tiene, se
        # muestra sola (mismo criterio que `ClienteDetalleDialog._cargar_
        # cliente` en el ABM standalone, acá aplicado al Facturador, que
        # no pasa por ese diálogo).
        self._actualizar_boton_nota_cliente()
        if self.tiene_nota_cliente:
            NotaClienteDialog(self.repos, cliente.CODIGO, parent=self).exec()
            self._actualizar_boton_nota_cliente()

        self._refrescar_proximo_numero()

        if cliente.VEND is not None:
            idx = self.combo_vendedor.findData(cliente.VEND)
            if idx >= 0:
                self.combo_vendedor.setCurrentIndex(idx)

        self._recalcular_totales()
        # Réplica del comportamiento pedido: apenas se resuelve el
        # Cliente, el cursor pasa directo a la primera celda (Sección)
        # de la primera fila de la grilla — no hace falta tocar nada más
        # a mano para empezar a cargar el Detalle.
        self.grid_detalle.foco_inicial()

    def _refrescar_proximo_numero(self) -> None:
        if self.cliente_actual is None:
            self.lbl_proximo_numero.setText("Próx. Nº: —")
            return
        letra = self.factura_service.letra_comprobante(self.cliente_actual.CIVA or 0, self.cliente_actual.PCIA)
        tipo_cbte = codigo_afip(letra, 1)
        punto_venta = punto_venta_por_tipo(1)
        try:
            ultimo = self.afip.ultimo_comprobante(punto_venta, tipo_cbte)
            self.lbl_proximo_numero.setText(f"Próx. Nº: {punto_venta:04d}-{ultimo + 1:08d} (Letra {letra})")
        except Exception:  # noqa: BLE001 — sólo informativo, no bloquea la carga
            self.lbl_proximo_numero.setText(f"Próx. Nº: (no disponible, Letra {letra})")

    def _actualizar_boton_nota_cliente(self) -> None:
        """Color del botón "Nota Clte." refleja si el Cliente actual tiene
        una Nota guardada (feedback del usuario, 2026-08-15) — alerta
        visual persistente, no sólo el popup que se ve una vez."""
        self.tiene_nota_cliente = self.cliente_actual is not None and (
            self.repos.notaclte().by_cliente(self.cliente_actual.CODIGO) is not None
        )
        self.btn_nota_cliente.setStyleSheet(
            "background-color: #ffd54f; font-weight: bold;" if self.tiene_nota_cliente else ""
        )

    def _on_nota_cliente(self) -> None:
        if self.cliente_actual is None:
            return
        dialogo = NotaClienteDialog(self.repos, self.cliente_actual.CODIGO, parent=self)
        dialogo.exec()
        self._actualizar_boton_nota_cliente()

    # ------------------------------------------------------------------
    # Dólares / cotización
    # ------------------------------------------------------------------
    def _cotizacion_actual(self) -> Decimal:
        if not self.chk_en_dolares.isChecked():
            return Decimal("1")
        ultima = self.repos.cotizacion().ultima()
        return Decimal(ultima.DOLAR) if ultima is not None and ultima.DOLAR else Decimal("1")

    def _on_en_dolares_cambiado(self) -> None:
        if self.renglones:
            QMessageBox.information(
                self,
                "Facturador",
                "Cambiar la moneda no reconvierte los renglones ya cargados — se recomienda \"Nueva\" antes de cambiar.",
            )
        self._recalcular_totales()

    # ------------------------------------------------------------------
    # Totales (Pie)
    # ------------------------------------------------------------------
    def _bonificacion_total(self) -> Decimal:
        if self.cliente_actual is None:
            return Decimal("0")
        total = Decimal("0")
        for renglon in self.renglones:
            fila_dto = self.repos.dtoxclte().by_cliente_seccion(self.cliente_actual.CODIGO, renglon.cod1)
            if fila_dto is None:
                continue
            porcentajes = [
                Decimal(p) for p in (fila_dto.DTO1, fila_dto.DTO2, fila_dto.DTO3, fila_dto.DTO4, fila_dto.DTO5) if p is not None
            ]
            total += self.factura_service.calcular_bonificacion_cascada(renglon.importe, porcentajes)
        return total

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
        # Feedback del usuario (2026-08-15): Emitir deshabilitado hasta
        # que haya Importe mayor a cero.
        self.btn_emitir.setEnabled(total.total > 0)

    # ------------------------------------------------------------------
    # Nueva factura
    # ------------------------------------------------------------------
    def _nueva_factura(self) -> None:
        self.cliente_actual = None
        self.grid_detalle.reiniciar()
        self.chk_en_dolares.setChecked(False)
        self.combo_forma_pedido.setCurrentIndex(FORMA_PEDIDO_DEFAULT_INDEX)
        self.txt_porcentaje_iibb.setText("0")
        self._refrescar_cliente()
        self._recalcular_totales()
        # Réplica de `Command4_Click` ("Nueva") y del arranque real del
        # Facturador (`FCMENU.frm`: `CabFact.Show` seguido siempre de
        # `BusClte.Show`) — arranca (y cada "Nueva") mostrando directo la
        # pantalla de selección de Cliente, con el cursor ya puesto en el
        # campo de búsqueda (`ClienteBusquedaWindow.spin_codigo`), no hay
        # que clickear "Elegir Cliente..." a mano primero.
        #
        # `QTimer.singleShot(0, ...)` en vez de la llamada directa
        # (feedback del usuario, 2026-08-17: "debe aparecer primero la
        # ventana correspondiente y recién luego la selección del
        # cliente") — llamado directo desde acá (dentro de `__init__`,
        # antes de que `MainMenuWindow._mostrar()` llegue a mostrar la
        # ventana) el buscador modal aparecía ANTES que el propio
        # Facturador. Encolado, corre recién en el próximo tick del loop
        # de eventos, después de que la ventana ya se mostró.
        QTimer.singleShot(0, self._on_elegir_cliente)

    # ------------------------------------------------------------------
    # Emitir
    # ------------------------------------------------------------------
    def _on_emitir(self) -> None:
        """**Boceto → Grabar/Cancelar → confirmar → AFIP** (pedido del
        usuario, 2026-08-18: "al apretar emitir, muestre un boceto en
        pdf de la factura con el boton grabar y cancelar abajo. Al
        apretar grabar, pregunte si se quiere grabar" — mismo espíritu
        que ya tiene `ReciboWindow._on_emitir`, pero acá el boceto NO
        puede tener CAE/QR reales todavía (sólo existen después de que
        AFIP aprueba) — se muestra ANTES de tocar AFIP para nada, con
        un sello "BORRADOR" (`DatosFacturaPDF.cae=None`, ver
        `migration/pdf.py`). La regla de oro de esta pantalla no cambia:
        nada se persiste ni se le pide nada a AFIP hasta el "Sí" de la
        confirmación final."""
        if self.cliente_actual is None:
            QMessageBox.warning(self, "Facturador", "Elegí un Cliente.")
            return
        if not self.renglones:
            QMessageBox.warning(self, "Facturador", "Agregá al menos un renglón.")
            return
        if not (self.cliente_actual.CUIT or "").strip():
            QMessageBox.warning(self, "Facturador", "El cliente no tiene CUIT cargado — no se puede solicitar CAE.")
            return

        letra = self.factura_service.letra_comprobante(self.cliente_actual.CIVA or 0, self.cliente_actual.PCIA)
        tipo_cbte = codigo_afip(letra, 1)
        punto_venta = punto_venta_por_tipo(1)
        total = self._calcular_total()
        fecha = date.today()

        try:
            numero = self.afip.ultimo_comprobante(punto_venta, tipo_cbte) + 1
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Facturador", f"Error al conectar con AFIP:\n{exc}")
            return

        try:
            datos_borrador = DatosFacturaPDF(
                letra=letra,
                punto_venta=punto_venta,
                numero=numero,
                fecha=fecha,
                cliente_codigo=self.cliente_actual.CODIGO,
                cliente_nombre=(self.cliente_actual.NOMB or "").strip(),
                cliente_cuit=self.cliente_actual.CUIT or "",
                cliente_civa=self.cliente_actual.CIVA or 0,
                cliente_domicilio=(self.cliente_actual.DIR or "").strip(),
                renglones=list(self.renglones),
                total=total,
                en_dolares=self.chk_en_dolares.isChecked(),
            )
            ruta_borrador = generar_pdf_factura(datos_borrador)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Facturador", f"No se pudo generar el boceto del PDF:\n{exc}")
            return

        vista_previa = PdfPreviewDialog(
            ruta_borrador,
            titulo=f"Boceto — Factura {letra} (sin CAE todavía)",
            mostrar_boton_grabar=True,
            parent=self,
        )
        vista_previa.exec()
        if not vista_previa.grabar:
            return

        respuesta = QMessageBox.question(
            self,
            "Facturador",
            f"¿Desea grabar la Factura {letra} por $ {format_decimal(total.total)}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        try:
            resultado_cae: ResultadoCAE = self.afip.solicitar_cae(
                tipo_cbte=tipo_cbte,
                punto_venta=punto_venta,
                cbte_nro=numero,
                cuit_receptor=self.cliente_actual.CUIT,
                importe_neto=total.neto_gravado,
                importe_iva=total.iva,
                importe_iibb=total.percepcion_iibb,
                importe_total=total.total,
                fecha_cbte=fecha,
                condicion_iva_receptor_id=condicion_iva_receptor(self.cliente_actual.CIVA),
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Facturador", f"Error al conectar con AFIP:\n{exc}")
            return

        if not resultado_cae.aprobado:
            QMessageBox.critical(self, "Facturador", f"AFIP no aprobó el comprobante.\nMotivo: {resultado_cae.motivo}")
            return

        try:
            resultado = self.emision_service.emitir_factura(
                cliente=self.cliente_actual,
                letra=letra,
                punto_venta=punto_venta,
                numero_comprobante=numero,
                renglones=self.renglones,
                total=total,
                porcentaje_iibb=self._porcentaje_iibb(),
                usuario=self._usuario,
                fecha=fecha,
                tipo_venta=str(self.combo_forma_pedido.currentIndex() + 1),
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self,
                "Facturador",
                f"AFIP aprobó el CAE {resultado_cae.cae} pero la Factura NO se pudo grabar en el sistema:\n{exc}\n"
                "Anotar el CAE para grabarla manualmente.",
            )
            return

        qr_url = generar_qr_afip(
            cuit_emisor=CUIT_EMISOR_DEFAULT,
            punto_venta=punto_venta,
            tipo_cbte=tipo_cbte,
            nro_cbte=numero,
            importe_total=total.total,
            tipo_doc_receptor=80,
            nro_doc_receptor=self.cliente_actual.CUIT,
            cae=resultado_cae.cae,
            fecha_cbte=fecha,
        )

        ruta_pdf = None
        try:
            datos_pdf = DatosFacturaPDF(
                letra=letra,
                punto_venta=punto_venta,
                numero=numero,
                fecha=fecha,
                cliente_codigo=self.cliente_actual.CODIGO,
                cliente_nombre=(self.cliente_actual.NOMB or "").strip(),
                cliente_cuit=self.cliente_actual.CUIT or "",
                cliente_civa=self.cliente_actual.CIVA or 0,
                cliente_domicilio=(self.cliente_actual.DIR or "").strip(),
                renglones=list(self.renglones),
                total=total,
                cae=resultado_cae.cae,
                cae_vencimiento=resultado_cae.vencimiento,
                qr_url=qr_url,
                en_dolares=self.chk_en_dolares.isChecked(),
            )
            ruta_pdf = generar_pdf_factura(datos_pdf)
        except Exception as exc:  # noqa: BLE001 — la factura YA se emitió/grabó, esto no debe revertirla
            QMessageBox.warning(
                self, "Facturador", f"La Factura se emitió y grabó correctamente, pero el PDF falló:\n{exc}"
            )

        self.ultimo_resultado = {
            "resultado": resultado,
            "letra": letra,
            "punto_venta": punto_venta,
            "numero": numero,
            "cae": resultado_cae,
            "total": total,
            "qr_url": qr_url,
            "cliente": self.cliente_actual,
            "renglones": list(self.renglones),
            "fecha": fecha,
            "pdf": ruta_pdf,
        }

        mensaje_pdf = f"\nPDF: {ruta_pdf}" if ruta_pdf else "\n(el PDF no se pudo generar, ver aviso anterior)"
        QMessageBox.information(
            self,
            "Facturador",
            f"Factura {letra} {punto_venta:04d}-{numero:08d} emitida.\n"
            f"CAE: {resultado_cae.cae}\nVencimiento: {resultado_cae.vencimiento}\nTotal: $ {format_decimal(total.total)}"
            f"{mensaje_pdf}",
        )
        self._nueva_factura()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        # Guarda de "Salir" (pedido del usuario, 2026-08-17): si hay algo
        # cargado, confirma antes de perderlo — cubre tanto el botón
        # "Salir" como el cierre nativo de la subventana (X del MDI).
        if self.renglones:
            respuesta = QMessageBox.question(
                self,
                "Facturador",
                "Hay una Factura cargada — se va a perder el ingreso.\n¿Desea salir igual?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if respuesta != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.db.close()
        super().closeEvent(event)
