"""`ReciboWindow` — ventana única del Recibo (Cabecera/Pendientes de
Cobro/Pagos/Pie), equivalente conceptual a `CabRec.frm`+`DetRec.frm`+
`DetPago.frm`+`PieRec.frm`+`EmiRec.frm` del legacy.

**Arquitectura confirmada con el usuario**: mismo criterio ya aplicado a
`FacturadorWindow` — una sola ventana cohesiva con varias secciones, no
los 5 formularios MDI que el legacy apila para simular una sola
pantalla.

**Circuito**:
1. Cabecera: elegir Cliente — mismo `ClienteBusquedaWindow` en modo
   selección que el Facturador (confirmado en el código: `FCMENU.frm
   INGREC_Click` usa el mismo patrón que `INGFAC_Click`, `CabRec.Show`
   + `BusClte` en modo selección). Verificación de Nota de Cliente
   (popup automático + color de botón) ANTES de mostrar los Pendientes
   — mismo criterio ya pedido para el Facturador (feedback del usuario,
   2026-08-15).
2. Pendientes de Cobro: grilla de comprobantes con saldo
   (`CtascteRepository.pendientes_cobro()`, réplica exacta de `DetRec.frm
   DoVer3()`). Un clic en cualquier columna que no sea Aplicado/Descuento
   alterna el Aplicado entre $0 y el Debe completo — igual que `DetRec.frm
   FG1_Click`. Editar Aplicado/Descuento a mano permite un pago parcial
   con descuento por pronto pago — OJO con la semántica real (`DetRec.frm
   FG1_AfterEdit` Case 4 + `EmiRec.frm Graba():1095-1097`, encontrada al
   escribir los tests de `EmisionReciboService`): el Descuento es un
   RECORTE del propio Aplicado, no algo que se suma aparte — el
   comprobante se salda por el Aplicado completo, y sólo `Aplicado −
   Descuento` es lo que el operador tiene que tender de verdad.
3. Anticipo: importe suelto sin comprobante puntual — graba su propia
   fila `Ctascte` TIPO=5 (decisión confirmada con el usuario, distinto
   del legacy que lo mezclaba sin dejar rastro).
4. Pagos: Efectivo directo, más Cheques/Retenciones vía `PagoDialog`.
5. Pie: Totales (Deuda Total/Selección/Descuentos/A Pagar/Pago/
   Diferencia). Emitir habilitado sólo cuando la Diferencia es
   EXACTAMENTE $0 y hay algo que cobrar — misma regla que gatea
   "Grabar" en el legacy (`PieRec.frm Label304_Change`, sin tolerancia
   de redondeo).
6. Emitir: **Vista Previa (PDF) → Grabar** (decisión del usuario,
   2026-08-15, ajustada dos veces el mismo día: primero de Confirmar→
   Grabar→Imprimir a Vista Previa→Confirmar→Grabar, y después sacando
   la ventana de confirmación aparte — el botón "Grabar" vive DENTRO de
   la vista previa). El PDF se genera con el número que TENDRÍA el
   comprobante (`Parametro.NUME8+1`, sin persistir nada todavía) y se
   muestra embebido (`PdfPreviewDialog`, `QPdfView`, `mostrar_boton_
   grabar=True`); sólo si se clickea "Grabar" ahí se llama a
   `EmisionReciboService.emitir_recibo()` — "Cancelar" no persiste nada.

**Fuera de alcance todavía**: "Ver C.C." (`CabRec.frm Command5`, abre
`CtaCte.frm`) — se agrega cuando se migre la Consulta de Cuenta
Corriente (próxima fase, orden confirmado por el usuario).
"""

from __future__ import annotations

import os
from datetime import date
from decimal import Decimal

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QBrush, QColor
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.fechas import formatear_fecha_larga
from migration.models import Cliente, Ctascte
from migration.pdf import DatosReciboPDF, generar_pdf_recibo
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import AplicacionPago, EmisionReciboService, PagoCheque, PagoRetencion

from .cliente_busqueda_window import ClienteBusquedaWindow
from .decimals import format_decimal, parse_decimal
from .nota_cliente_dialog import NotaClienteDialog
from .pago_dialog import TIPOS_RETENCION, PagoDialog
from .pdf_preview_dialog import PdfPreviewDialog
from .widgets import MontoLineEdit, compactar_alto_filas, redimensionar_pct_pantalla, texto_contacto_cliente

ETIQUETAS_TIPREG = dict(TIPOS_RETENCION)

COL_COMPROBANTE = 0
COL_TIPO = 1
COL_FECHA = 2
COL_FECVTO = 3
COL_DEBE = 4
COL_APLICADO = 5
COL_DESCUENTO = 6
COLUMNAS_PENDIENTES = ["Comprobante", "Tipo", "Fecha", "Fec.Vto.", "Debe", "Aplicado", "Descuento"]
COLUMNAS_SOLO_LECTURA = (COL_COMPROBANTE, COL_TIPO, COL_FECHA, COL_FECVTO, COL_DEBE)
# Ajustar a la derecha todos los campos numéricos (feedback del usuario,
# 2026-08-19) — antes sólo el Comprobante quedaba alineado a la derecha;
# Debe/Aplicado/Descuento (importes de verdad) quedaban a la izquierda.
COLUMNAS_ALINEADAS_DERECHA = (COL_COMPROBANTE, COL_DEBE, COL_APLICADO, COL_DESCUENTO)

# Colores "invertidos" para una fila con Aplicado > 0 (feedback del
# usuario, 2026-08-15) — mismo espíritu que el resaltado de selección
# que hacía DetRec.frm FG1_MouseMove/Click con Cell(...) sobre la fila
# activa, acá persistente mientras la fila tenga algo aplicado.
COLOR_FILA_APLICADA_FONDO = QColor("#1f3864")
COLOR_FILA_APLICADA_TEXTO = QColor("#ffffff")

COL_PAGO_TIPO = 0
COL_PAGO_REFERENCIA = 1
COL_PAGO_BANCO = 2
COL_PAGO_FECHA = 3
COL_PAGO_IMPORTE = 4
COLUMNAS_PAGOS = ["Tipo", "Nº/Concepto", "Banco", "Fecha", "Importe"]


class ReciboWindow(QMainWindow):
    TEXTO_BTN_ELEGIR_CLIENTE = "Elegir Cliente..."
    TEXTO_BTN_CAMBIAR_CLIENTE = "Cambiar Cliente"

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Recibo")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        redimensionar_pct_pantalla(self, 85, 55)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.emision_service = EmisionReciboService(self.db)

        self.cliente_actual: Cliente | None = None
        self.tiene_nota_cliente = False
        # Ya no se muestra en la cabecera (feedback del usuario,
        # 2026-08-18: "Quitar usuario de la cabecera") pero se sigue
        # necesitando para auditoría (`EmisionReciboService.
        # emitir_recibo(usuario=...)`).
        self._usuario = os.environ.get("USERNAME", "SISTEMA")[:6]
        self._pendientes: list[Ctascte] = []
        self._pagos: list[PagoCheque | PagoRetencion] = []
        # Guard contra reentrancia de `itemChanged` (mismo criterio que
        # `DetalleGrid`, aunque acá el modelo de edición es mucho más
        # simple — no hay editores dinámicos por Sección, sólo celdas
        # numéricas planas).
        self._actualizando = False

        self._construir_ui()
        self._nueva_recibo()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        # Líneas más juntas (feedback del usuario, 2026-08-17: "pasa lo
        # mismo que en abm clientes, se sale de la ventana... hay que
        # juntar mas las lineas") — antes usaba el spacing default de Qt
        # entre las 4 secciones apiladas.
        layout.setSpacing(6)

        layout.addWidget(self._armar_cabecera())

        # Pagos a la derecha de Pendientes (feedback del usuario,
        # 2026-08-17, segunda ronda: "Es muy largo, se sale del panel.
        # Colocar el panel de pagos a la derecha como en el legacy" —
        # `DetRec.frm`/`DetPago.frm` separados en el original en vez de
        # apilados). Antes eran 4 secciones en una sola columna vertical;
        # ahora Pendientes+Pagos comparten una fila, así se gana el alto
        # completo de esa fila para el pie.
        # Pendientes 20% más angosto / Pagos 20% más ancho (feedback del
        # usuario, 2026-08-19) — la proporción original era 2:1 (66,7%/
        # 33,3% del ancho); reducir un panel 20% y agrandar el otro 20%
        # sobre esa base da la nueva proporción 4:3 (53,3%/40%,
        # normalizado).
        fila_central = QHBoxLayout()
        fila_central.addWidget(self._armar_pendientes(), stretch=4)
        fila_central.addWidget(self._armar_pagos(), stretch=3)
        layout.addLayout(fila_central, stretch=1)

        layout.addWidget(self._armar_pie())

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QGroupBox:
        grupo = QGroupBox("Cabecera — Recibo")
        layout = QVBoxLayout(grupo)
        layout.setSpacing(4)

        # Código+nombre a la izquierda, botón cambiar-cliente a la
        # derecha (antes de "Notas") — pedido del usuario, 2026-08-20.
        fila_cliente = QHBoxLayout()
        self.lbl_cliente = QLabel("(sin cliente elegido)")
        fila_cliente.addWidget(self.lbl_cliente, stretch=1)
        self.btn_elegir_cliente = QPushButton(self.TEXTO_BTN_ELEGIR_CLIENTE)
        self.btn_elegir_cliente.clicked.connect(self._on_elegir_cliente)
        fila_cliente.addWidget(self.btn_elegir_cliente)
        self.btn_nota_cliente = QPushButton("Nota Clte.")
        self.btn_nota_cliente.setEnabled(False)
        self.btn_nota_cliente.clicked.connect(self._on_nota_cliente)
        fila_cliente.addWidget(self.btn_nota_cliente)
        layout.addLayout(fila_cliente)

        # Localidad/Teléfono/Email del cliente elegido (ídem).
        self.lbl_cliente_contacto = QLabel("—")
        layout.addWidget(self.lbl_cliente_contacto)

        fila_datos = QHBoxLayout()
        fila_datos.addWidget(QLabel("Crédito :"))
        self.lbl_credito = QLabel("$ 0,00")
        fila_datos.addWidget(self.lbl_credito)
        fila_datos.addWidget(QLabel("Deuda :"))
        self.lbl_deuda = QLabel("$ 0,00")
        fila_datos.addWidget(self.lbl_deuda)

        fila_datos.addStretch()
        self.lbl_proximo_numero = QLabel("Próx. Nº: —")
        fila_datos.addWidget(self.lbl_proximo_numero)
        self.btn_nueva = QPushButton("Nuevo")
        self.btn_nueva.clicked.connect(self._nueva_recibo)
        fila_datos.addWidget(self.btn_nueva)
        layout.addLayout(fila_datos)

        return grupo

    # ------------------------------------------------------------------
    # Pendientes de Cobro
    # ------------------------------------------------------------------
    def _armar_pendientes(self) -> QGroupBox:
        grupo = QGroupBox("Pendientes de Cobro — clic alterna Aplicado, doble clic edita Aplicado/Descuento")
        layout = QVBoxLayout(grupo)

        self.tabla_pendientes = QTableWidget(0, len(COLUMNAS_PENDIENTES))
        self.tabla_pendientes.setHorizontalHeaderLabels(COLUMNAS_PENDIENTES)
        self.tabla_pendientes.verticalHeader().setVisible(False)
        # Celdas más grandes (feedback del usuario, 2026-08-15: "esta
        # todo muy apretado") — fila más alta + tipografía un poco más
        # grande + padding lateral en cada celda. Achicada un 10% de
        # nuevo (feedback del usuario, 2026-08-19: "reducir la altura de
        # las filas 10% para que entren más renglones") sobre esa misma
        # base de 36px, no sobre el default de Qt.
        self.tabla_pendientes.verticalHeader().setDefaultSectionSize(round(36 * 0.9))
        fuente = self.tabla_pendientes.font()
        fuente.setPointSize(fuente.pointSize() + 1)
        self.tabla_pendientes.setFont(fuente)
        self.tabla_pendientes.setStyleSheet("QTableWidget::item { padding: 6px; }")
        self.tabla_pendientes.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabla_pendientes.itemClicked.connect(self._on_click_pendiente)
        self.tabla_pendientes.itemChanged.connect(self._on_item_pendiente_cambiado)
        layout.addWidget(self.tabla_pendientes, stretch=1)

        fila_anticipo = QHBoxLayout()
        fila_anticipo.addWidget(QLabel("Anticipo (sin comprobante) :"))
        self.txt_anticipo = MontoLineEdit("0")
        self.txt_anticipo.setMaximumWidth(120)
        self.txt_anticipo.editingFinished.connect(self._recalcular_totales)
        fila_anticipo.addWidget(self.txt_anticipo)
        fila_anticipo.addStretch()
        layout.addLayout(fila_anticipo)

        return grupo

    # ------------------------------------------------------------------
    # Pagos
    # ------------------------------------------------------------------
    def _armar_pagos(self) -> QGroupBox:
        # Panel angosto a la derecha ahora (ver `_construir_ui`), no una
        # franja horizontal completa — Efectivo y "Agregar..." pasan a
        # filas propias (antes compartían una sola fila ancha) y la
        # tabla ya no tiene techo de alto fijo, crece para ocupar todo
        # el panel, a la par de Pendientes.
        grupo = QGroupBox("Pagos")
        layout = QVBoxLayout(grupo)

        # Ancho acorde a la cantidad de dígitos que va a alojar (feedback
        # del usuario, 2026-08-19: antes tenía `stretch=1` y quedaba
        # estirado a todo el ancho del panel) — mismo ancho que el resto
        # de los `MontoLineEdit` de importe del sistema (120px, ver
        # `txt_anticipo` acá mismo). Centrado con `addStretch()` a los
        # dos lados en vez de pegado a la izquierda.
        fila_efectivo = QHBoxLayout()
        fila_efectivo.addStretch()
        fila_efectivo.addWidget(QLabel("Efectivo :"))
        self.txt_efectivo = MontoLineEdit("0")
        self.txt_efectivo.setMaximumWidth(120)
        self.txt_efectivo.editingFinished.connect(self._recalcular_totales)
        fila_efectivo.addWidget(self.txt_efectivo)
        fila_efectivo.addStretch()
        layout.addLayout(fila_efectivo)

        self.btn_agregar_pago = QPushButton("Agregar Cheque/Retención...")
        self.btn_agregar_pago.clicked.connect(self._on_agregar_pago)
        layout.addWidget(self.btn_agregar_pago)

        # Un clic sobre una forma de pago ya cargada la abre para
        # modificar o borrar (feedback del usuario, 2026-08-15) — ya no
        # hace falta un botón "Quitar" aparte, esa acción vive ahora
        # dentro de PagoDialog.
        self.tabla_pagos = QTableWidget(0, len(COLUMNAS_PAGOS))
        self.tabla_pagos.setHorizontalHeaderLabels(COLUMNAS_PAGOS)
        self.tabla_pagos.verticalHeader().setVisible(False)
        self.tabla_pagos.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabla_pagos.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        compactar_alto_filas(self.tabla_pagos)
        self.tabla_pagos.itemClicked.connect(self._on_click_pago)
        layout.addWidget(self.tabla_pagos, stretch=1)

        return grupo

    # ------------------------------------------------------------------
    # Pie
    # ------------------------------------------------------------------
    def _armar_pie(self) -> QGroupBox:
        grupo = QGroupBox("Totales")
        layout = QVBoxLayout(grupo)
        layout.setSpacing(4)

        fila1 = QHBoxLayout()
        fila1.addWidget(QLabel("Deuda Total:"))
        self.lbl_deuda_total = QLabel("$ 0,00")
        fila1.addWidget(self.lbl_deuda_total)
        fila1.addSpacing(20)
        fila1.addWidget(QLabel("Selección:"))
        self.lbl_seleccion = QLabel("$ 0,00")
        fila1.addWidget(self.lbl_seleccion)
        fila1.addWidget(QLabel("Descuentos:"))
        self.lbl_descuentos = QLabel("$ 0,00")
        fila1.addWidget(self.lbl_descuentos)
        fila1.addStretch()
        layout.addLayout(fila1)

        fila2 = QHBoxLayout()
        lbl_a_pagar_titulo = QLabel("A Pagar:")
        lbl_a_pagar_titulo.setStyleSheet("font-weight: bold; font-size: 13pt;")
        self.lbl_a_pagar = QLabel("$ 0,00")
        self.lbl_a_pagar.setStyleSheet("font-weight: bold; font-size: 13pt;")
        fila2.addWidget(lbl_a_pagar_titulo)
        fila2.addWidget(self.lbl_a_pagar)
        fila2.addSpacing(20)
        fila2.addWidget(QLabel("Pago:"))
        self.lbl_pago = QLabel("$ 0,00")
        fila2.addWidget(self.lbl_pago)
        fila2.addWidget(QLabel("Diferencia:"))
        self.lbl_diferencia = QLabel("$ 0,00")
        fila2.addWidget(self.lbl_diferencia)
        fila2.addStretch()

        self.btn_emitir = QPushButton("Emitir")
        self.btn_emitir.setStyleSheet("font-weight: bold; padding: 10px;")
        self.btn_emitir.setEnabled(False)
        self.btn_emitir.clicked.connect(self._on_emitir)
        fila2.addWidget(self.btn_emitir)

        # Botón "Salir" (pedido del usuario, 2026-08-17) — la guarda de
        # "hay algo cargado, se va a perder" vive en `closeEvent` (no acá
        # directo), mismo criterio que en `FacturadorWindow`.
        fila2.addSpacing(12)
        btn_salir = QPushButton("Salir")
        btn_salir.clicked.connect(self.close)
        fila2.addWidget(btn_salir)

        layout.addLayout(fila2)

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
        # ClienteBusquedaWindow — se re-busca en self.db, mismo motivo
        # (y mismo bug real ya encontrado) que en FacturadorWindow.
        self.cliente_actual = self.repos.cliente().by_codigo(dialogo.cliente_elegido.CODIGO)
        self._refrescar_cliente()

    def _refrescar_cliente(self) -> None:
        cliente = self.cliente_actual
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.lbl_cliente_contacto.setText("—")
            self.btn_nota_cliente.setEnabled(False)
            self.tiene_nota_cliente = False
            self.btn_nota_cliente.setStyleSheet("")
            self.btn_elegir_cliente.setText(self.TEXTO_BTN_ELEGIR_CLIENTE)
            self.lbl_credito.setText("$ 0,00")
            self.lbl_deuda.setText("$ 0,00")
            self._cargar_pendientes([])
            self._refrescar_proximo_numero()
            return

        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()}")
        self.lbl_cliente_contacto.setText(texto_contacto_cliente(cliente))
        self.btn_nota_cliente.setEnabled(True)
        self.btn_elegir_cliente.setText(self.TEXTO_BTN_CAMBIAR_CLIENTE)

        # Verificación de Nota ANTES de mostrar los Pendientes de Cobro
        # — mismo criterio ya pedido para el Facturador (feedback del
        # usuario, 2026-08-15).
        self._actualizar_boton_nota_cliente()
        if self.tiene_nota_cliente:
            NotaClienteDialog(self.repos, cliente.CODIGO, parent=self).exec()
            self._actualizar_boton_nota_cliente()

        self.lbl_credito.setText(f"$ {format_decimal(cliente.CREDIT or Decimal('0'))}")
        self.lbl_deuda.setText(f"$ {format_decimal(cliente.DEUDA or Decimal('0'))}")

        pendientes = self.repos.ctascte().pendientes_cobro(cliente.CODIGO)
        self._cargar_pendientes(pendientes)
        self._refrescar_proximo_numero()
        self._recalcular_totales()

    def _actualizar_boton_nota_cliente(self) -> None:
        self.tiene_nota_cliente = self.cliente_actual is not None and (
            self.repos.notaclte().by_cliente(self.cliente_actual.CODIGO) is not None
        )
        self.btn_nota_cliente.setStyleSheet(
            "background-color: #ffd54f; font-weight: bold;" if self.tiene_nota_cliente else ""
        )

    def _on_nota_cliente(self) -> None:
        if self.cliente_actual is None:
            return
        NotaClienteDialog(self.repos, self.cliente_actual.CODIGO, parent=self).exec()
        self._actualizar_boton_nota_cliente()

    def _refrescar_proximo_numero(self) -> None:
        parametro = self.repos.parametro().get_config()
        if parametro is None or parametro.NUME8 is None:
            self.lbl_proximo_numero.setText("Próx. Nº: —")
            return
        self.lbl_proximo_numero.setText(f"Próx. Nº: {parametro.NUME8 + 1:08d}")

    # ------------------------------------------------------------------
    # Pendientes de Cobro
    # ------------------------------------------------------------------
    def _cargar_pendientes(self, pendientes: list[Ctascte]) -> None:
        self._pendientes = pendientes
        self._actualizando = True
        try:
            self.tabla_pendientes.setRowCount(0)
            for fila_idx, comprobante in enumerate(pendientes):
                self.tabla_pendientes.insertRow(fila_idx)
                etiqueta = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
                debe = comprobante.DEBE or Decimal("0")
                valores = [
                    str(comprobante.CPBTE or ""),
                    etiqueta,
                    comprobante.FECHA.strftime("%d/%m/%Y") if comprobante.FECHA else "",
                    comprobante.FECVTO.strftime("%d/%m/%Y") if comprobante.FECVTO else "",
                    format_decimal(debe),
                    "0",
                    "0",
                ]
                for col, texto in enumerate(valores):
                    item = QTableWidgetItem(texto)
                    if col in COLUMNAS_SOLO_LECTURA:
                        item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                    if col in COLUMNAS_ALINEADAS_DERECHA:
                        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    self.tabla_pendientes.setItem(fila_idx, col, item)
            self.tabla_pendientes.resizeColumnsToContents()
            # "la celda aplicado es muy chica" (feedback del usuario,
            # 2026-08-15) — resizeColumnsToContents() la deja del ancho
            # de "0,00", muy angosta para escribir un importe real.
            for col in (COL_APLICADO, COL_DESCUENTO):
                ancho_actual = self.tabla_pendientes.columnWidth(col)
                self.tabla_pendientes.setColumnWidth(col, max(ancho_actual, 110))
        finally:
            self._actualizando = False

    def _on_click_pendiente(self, item: QTableWidgetItem) -> None:
        if self._actualizando or item.column() in (COL_APLICADO, COL_DESCUENTO):
            return
        # Clic en cualquier columna que no sea Aplicado/Descuento alterna
        # el Aplicado entre $0 y el Debe completo — mismo criterio que
        # DetRec.frm FG1_Click.
        self._alternar_aplicado(item.row())

    def _alternar_aplicado(self, fila: int) -> None:
        item_debe = self.tabla_pendientes.item(fila, COL_DEBE)
        item_aplicado = self.tabla_pendientes.item(fila, COL_APLICADO)
        debe = parse_decimal(item_debe.text())
        aplicado_actual = parse_decimal(item_aplicado.text())
        nuevo = Decimal("0") if aplicado_actual > 0 else debe
        self._actualizando = True
        try:
            item_aplicado.setText(format_decimal(nuevo))
            if nuevo == 0:
                self.tabla_pendientes.item(fila, COL_DESCUENTO).setText("0")
        finally:
            self._actualizando = False
        self._actualizar_color_fila(fila)
        self._recalcular_totales()

    def _on_item_pendiente_cambiado(self, item: QTableWidgetItem) -> None:
        if self._actualizando:
            return
        if item.column() == COL_APLICADO:
            aplicado = parse_decimal(item.text())
            debe = parse_decimal(self.tabla_pendientes.item(item.row(), COL_DEBE).text())
            if aplicado > debe:
                self._actualizando = True
                try:
                    item.setText(format_decimal(debe))
                finally:
                    self._actualizando = False
        elif item.column() == COL_DESCUENTO:
            descuento = parse_decimal(item.text())
            aplicado = parse_decimal(self.tabla_pendientes.item(item.row(), COL_APLICADO).text())
            if descuento > aplicado:
                self._actualizando = True
                try:
                    item.setText(format_decimal(aplicado))
                finally:
                    self._actualizando = False
        self._actualizar_color_fila(item.row())
        self._recalcular_totales()

    def _actualizar_color_fila(self, fila: int) -> None:
        """Colores "invertidos" mientras la fila tenga Aplicado > 0
        (feedback del usuario, 2026-08-15) — visualmente distingue de
        un vistazo qué comprobantes van a entrar en este Recibo."""
        aplicado = parse_decimal(self.tabla_pendientes.item(fila, COL_APLICADO).text())
        con_aplicado = aplicado > 0
        for col in range(self.tabla_pendientes.columnCount()):
            item = self.tabla_pendientes.item(fila, col)
            if item is None:
                continue
            if con_aplicado:
                item.setBackground(COLOR_FILA_APLICADA_FONDO)
                item.setForeground(COLOR_FILA_APLICADA_TEXTO)
            else:
                # Bug real encontrado con datos reales (2026-08-15): "al
                # dess-seleccionar queda la fila negra" — `setBackground
                # (QColor())` NO limpia el color, un `QColor` inválido se
                # pinta NEGRO sólido (Qt lo convierte en un QBrush con
                # RGB (0,0,0)). El que sí vuelve al color default de la
                # fila es `QBrush()` (sin pincel, `Qt.BrushStyle.NoBrush`).
                item.setBackground(QBrush())
                item.setForeground(QBrush())

    def _aplicaciones_actuales(self) -> list[AplicacionPago]:
        resultado = []
        for fila, comprobante in enumerate(self._pendientes):
            aplicado = parse_decimal(self.tabla_pendientes.item(fila, COL_APLICADO).text())
            if aplicado <= 0:
                continue
            descuento = parse_decimal(self.tabla_pendientes.item(fila, COL_DESCUENTO).text())
            resultado.append(AplicacionPago(comprobante=comprobante, importe_aplicado=aplicado, descuento=descuento))
        return resultado

    def _anticipo_actual(self) -> Decimal:
        return self.txt_anticipo.decimal()

    # ------------------------------------------------------------------
    # Pagos
    # ------------------------------------------------------------------
    def _on_agregar_pago(self) -> None:
        dialogo = PagoDialog(self.repos, parent=self)
        dialogo.exec()
        if dialogo.resultado is None:
            return
        self._pagos.append(dialogo.resultado)
        self._refrescar_tabla_pagos()
        self._recalcular_totales()

    def _on_click_pago(self, item: QTableWidgetItem) -> None:
        """Un clic sobre una forma de pago ya cargada la abre para
        modificar o borrar (feedback del usuario, 2026-08-15) — antes
        había un botón "Quitar" aparte, ahora esa acción vive dentro de
        `PagoDialog` (ver su botón "Borrar", sólo visible en modo edición)."""
        fila = item.row()
        if fila < 0 or fila >= len(self._pagos):
            return
        dialogo = PagoDialog(self.repos, pago_existente=self._pagos[fila], parent=self)
        dialogo.exec()
        if dialogo.eliminar:
            del self._pagos[fila]
        elif dialogo.resultado is not None:
            self._pagos[fila] = dialogo.resultado
        else:
            return  # cancelado, sin cambios
        self._refrescar_tabla_pagos()
        self._recalcular_totales()

    def _nombre_banco(self, codigo: str) -> str:
        codigo = (codigo or "").strip()
        if not codigo.isdigit():
            return codigo
        banco = self.repos.banco().by_cod(int(codigo))
        return (banco.NOMBRE or "").strip() if banco is not None else codigo

    def _refrescar_tabla_pagos(self) -> None:
        self.tabla_pagos.setRowCount(0)
        for fila, pago in enumerate(self._pagos):
            self.tabla_pagos.insertRow(fila)
            if isinstance(pago, PagoCheque):
                tipo = "Cheque"
                referencia = f"Nº {pago.nro_cheque}"
                banco = self._nombre_banco(pago.banco)
                fecha_txt = formatear_fecha_larga(pago.fecha_vencimiento)
            else:
                tipo = ETIQUETAS_TIPREG.get(pago.tipreg, str(pago.tipreg))
                referencia = pago.concepto or ""
                banco = self._nombre_banco(pago.banco_o_dato) if pago.banco_o_dato else ""
                fecha_txt = ""
            self.tabla_pagos.setItem(fila, COL_PAGO_TIPO, QTableWidgetItem(tipo))
            self.tabla_pagos.setItem(fila, COL_PAGO_REFERENCIA, QTableWidgetItem(referencia))
            self.tabla_pagos.setItem(fila, COL_PAGO_BANCO, QTableWidgetItem(banco))
            self.tabla_pagos.setItem(fila, COL_PAGO_FECHA, QTableWidgetItem(fecha_txt))
            item_importe = QTableWidgetItem(format_decimal(pago.importe))
            item_importe.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.tabla_pagos.setItem(fila, COL_PAGO_IMPORTE, item_importe)
        self.tabla_pagos.resizeColumnsToContents()

    def _importe_efectivo_actual(self) -> Decimal:
        return self.txt_efectivo.decimal()

    # ------------------------------------------------------------------
    # Totales (Pie)
    # ------------------------------------------------------------------
    def _recalcular_totales(self) -> None:
        deuda_total = Decimal("0")
        for fila in range(self.tabla_pendientes.rowCount()):
            deuda_total += parse_decimal(self.tabla_pendientes.item(fila, COL_DEBE).text())

        aplicaciones = self._aplicaciones_actuales()
        seleccion = sum((a.importe_aplicado for a in aplicaciones), Decimal("0"))
        descuentos = sum((a.descuento for a in aplicaciones), Decimal("0"))
        anticipo = self._anticipo_actual()
        a_pagar = seleccion + anticipo - descuentos

        pago = self._importe_efectivo_actual() + sum((p.importe for p in self._pagos), Decimal("0"))
        diferencia = a_pagar - pago

        self.lbl_deuda_total.setText(f"$ {format_decimal(deuda_total)}")
        self.lbl_seleccion.setText(f"$ {format_decimal(seleccion)}")
        self.lbl_descuentos.setText(f"$ {format_decimal(descuentos)}")
        self.lbl_a_pagar.setText(f"$ {format_decimal(a_pagar)}")
        self.lbl_pago.setText(f"$ {format_decimal(pago)}")
        self.lbl_diferencia.setText(f"$ {format_decimal(diferencia)}")

        # Emitir habilitado sólo con Diferencia exactamente $0 y algo
        # real para cobrar — misma regla que gatea "Grabar" en el
        # legacy (PieRec.frm Label304_Change), sin tolerancia de redondeo.
        self.btn_emitir.setEnabled(a_pagar > 0 and diferencia == 0)

    # ------------------------------------------------------------------
    # Nuevo Recibo
    # ------------------------------------------------------------------
    def _nueva_recibo(self) -> None:
        self.cliente_actual = None
        self._pendientes = []
        self._pagos = []
        self._actualizando = True
        try:
            self.tabla_pendientes.setRowCount(0)
        finally:
            self._actualizando = False
        self.tabla_pagos.setRowCount(0)
        self.txt_anticipo.set_decimal(Decimal("0"))
        self.txt_efectivo.set_decimal(Decimal("0"))
        self._refrescar_cliente()
        self._recalcular_totales()
        # Réplica del arranque real del Recibo (`FCMENU.frm INGREC_Click`/
        # `ReiniciaRec`): siempre muestra la búsqueda de Cliente sola, no
        # hace falta clickear "Elegir Cliente..." a mano primero — mismo
        # criterio ya aplicado al Facturador. `QTimer.singleShot(0, ...)`
        # (feedback del usuario, 2026-08-17): deja que la ventana del
        # Recibo se muestre primero, el buscador modal aparece recién en
        # el siguiente tick del loop de eventos — antes se veía primero
        # el buscador porque esto corría síncrono dentro de `__init__`.
        QTimer.singleShot(0, self._on_elegir_cliente)

    # ------------------------------------------------------------------
    # Emitir
    # ------------------------------------------------------------------
    def _on_emitir(self) -> None:
        """Vista Previa → Grabar (decisión del usuario, 2026-08-15,
        ajustada el mismo día): el PDF se genera y se muestra ANTES de
        persistir nada, con el botón "Grabar" DENTRO de la propia vista
        previa (`PdfPreviewDialog(mostrar_boton_grabar=True)`) — ya no
        hay una ventana de confirmación aparte, el usuario la sacó
        después de probar el flujo con dos pasos. `numero`/`correlativo`
        se calculan igual acá que dentro de `EmisionReciboService.
        emitir_recibo()` (mismas fuentes puras, `Parametro.NUME8+1`/
        `ClienteRepository.proximo_correlativo()`, ninguna las muta
        todavía) — el PDF de vista previa sigue siendo válido después de
        grabar, no hace falta generarlo dos veces."""
        if self.cliente_actual is None:
            QMessageBox.warning(self, "Recibo", "Elegí un Cliente.")
            return

        aplicaciones = self._aplicaciones_actuales()
        anticipo = self._anticipo_actual()
        importe_efectivo = self._importe_efectivo_actual()
        cheques = [p for p in self._pagos if isinstance(p, PagoCheque)]
        retenciones = [p for p in self._pagos if isinstance(p, PagoRetencion)]

        parametro = self.repos.parametro().get_config()
        if parametro is None or parametro.NUME8 is None:
            QMessageBox.critical(self, "Recibo", "No se encontró la configuración de numeración (Parametro).")
            return
        numero = parametro.NUME8 + 1
        correlativo_preview = self.repos.cliente().proximo_correlativo(self.cliente_actual)
        total_pago = importe_efectivo + sum((p.importe for p in self._pagos), Decimal("0"))

        try:
            datos_pdf = DatosReciboPDF(
                numero=numero,
                fecha=date.today(),
                correlativo=correlativo_preview,
                cliente_codigo=self.cliente_actual.CODIGO,
                cliente_nombre=(self.cliente_actual.NOMB or "").strip(),
                cliente_cuit=self.cliente_actual.CUIT or "",
                aplicaciones=aplicaciones,
                anticipo=anticipo,
                importe_efectivo=importe_efectivo,
                cheques=cheques,
                retenciones=retenciones,
                total_pago=total_pago,
            )
            ruta_pdf = generar_pdf_recibo(datos_pdf)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Recibo", f"No se pudo generar la vista previa del PDF:\n{exc}")
            return

        vista_previa = PdfPreviewDialog(
            ruta_pdf, titulo=f"Vista Previa — Recibo Nº {numero:08d}", mostrar_boton_grabar=True, parent=self
        )
        vista_previa.exec()
        if not vista_previa.grabar:
            return

        try:
            resultado = self.emision_service.emitir_recibo(
                cliente=self.cliente_actual,
                numero=numero,
                aplicaciones=aplicaciones,
                anticipo=anticipo,
                importe_efectivo=importe_efectivo,
                cheques=cheques,
                retenciones=retenciones,
                usuario=self._usuario,
                fecha=date.today(),
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Recibo", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Recibo", f"No se pudo grabar el Recibo:\n{exc}")
            return

        QMessageBox.information(
            self,
            "Recibo",
            f"Recibo Nº {resultado.numero:08d} grabado.\n"
            f"Total: $ {format_decimal(resultado.total_pago)}\nPDF: {ruta_pdf}",
        )
        self._nueva_recibo()

    # ------------------------------------------------------------------
    def _hay_algo_cargado(self) -> bool:
        """"Algo cargado" para la guarda de "Salir" (pedido del usuario,
        2026-08-17): cualquier comprobante aplicado, pago agregado,
        anticipo o efectivo tipeado — todo lo que se perdería al salir
        sin emitir."""
        if self._pagos:
            return True
        if self._anticipo_actual() > 0 or self._importe_efectivo_actual() > 0:
            return True
        return any(a.importe_aplicado > 0 for a in self._aplicaciones_actuales())

    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        # Guarda de "Salir" — cubre tanto el botón como el cierre nativo
        # de la subventana (X del MDI), mismo criterio que en
        # `FacturadorWindow`.
        if self._hay_algo_cargado():
            respuesta = QMessageBox.question(
                self,
                "Recibo",
                "Hay un Recibo cargado — se va a perder el ingreso.\n¿Desea salir igual?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if respuesta != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.db.close()
        super().closeEvent(event)
