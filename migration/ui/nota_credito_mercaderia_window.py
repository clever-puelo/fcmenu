"""`NotaCreditoMercaderiaWindow` — Nota de Crédito por Devolución de
Mercadería (Motivo=1/"DEV.MERC."), migración de `CabFact.frm`
(TipoFac=2, Motivo=1) + `DetFact.frm` (reusado tal cual por el legacy
para esta rama, ver `CabFact.frm Sub Combo5_Click`) + `PieFact.frm` +
`EmiFact.frm` (bloque `ConArticulos:`, reversa de Stock) +
`ImputFC.frm`.

**Alcance confirmado con el usuario (2026-08-19)** — ver también
`NotaCreditoConceptoWindow` (Motivo != 1, conceptos libres):
- El legacy reutiliza literalmente la grilla de artículos de Factura
  (`DetFact.FG1`) para esta rama, en blanco, a precio de catálogo
  actual. Acá se hace mejor: el operador busca la Factura ORIGINAL del
  Cliente y elige qué renglones reales devolver — el precio acreditado
  es el que realmente se facturó (`Fcestad1.PVTA`/`.IMPTE` de ese
  renglón puntual), no el de catálogo hoy.
- **Devolución PARCIAL de cantidad** (pedido explícito del usuario,
  2026-08-20 — reemplaza la primera versión, que sólo dejaba tildar el
  renglón completo): la columna "Cant. a Devolver" trae precargada la
  cantidad COMPLETA vendida en ese renglón (`Fcestad1.CANT`), editable
  hacia abajo (nunca por encima de lo vendido); en 0 el renglón no se
  devuelve (equivalente al checkbox "sin tildar" de la versión
  anterior). El importe/pulg/mtr/milim/telas de ese renglón se
  prorratean linealmente contra la cantidad editada — ver
  `renglon_devolucion_parcial()` (`services.py`) para la fórmula exacta
  y su límite conocido (renglones sin `CANT` real, ej. tela por metro,
  quedan todo-o-nada, sin poder partirse).
- Emitida la Nota de Crédito, el Stock del artículo se repone
  automáticamente (`EmisionNotaCreditoMercaderiaService`, reversa real
  de `EmiFact.frm:1864-1928`) y se imputa contra un comprobante con
  deuda elegido por el operador (réplica de `ImputFC.frm`, mismo
  mecanismo que `NotaCreditoConceptoWindow`) — no necesariamente la
  misma Factura que se está devolviendo, igual que en el legacy.

**Vista previa en PDF** (agregada 2026-08-20, pedido del usuario):
mismo patrón Boceto→Grabar que `FacturadorWindow`, reusando
`DatosFacturaPDF`/`generar_pdf_factura` directo (esta ventana ya arma
`RenglonEmision`/`TotalFactura`, no hace falta adaptador).

**Fuera de alcance**: sin percepción de IIBB (el legacy la permite
estructuralmente pero no es un caso real de negocio para una
devolución — avisar si hace falta).
"""

from __future__ import annotations

import os
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Optional

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
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
from migration.models import Cliente, Ctascte, FcivaVta
from migration.pdf import DatosFacturaPDF, generar_pdf_factura
from migration.provincias import nombre_provincia
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import (
    CuentaCorrienteService,
    EmisionNotaCreditoMercaderiaService,
    FacturaService,
    RenglonEmision,
    condicion_venta_texto,
    renglon_devolucion_parcial,
)

from .cliente_busqueda_window import ClienteBusquedaWindow
from .comprobante_aplicar_dialog import ComprobanteAplicarDialog
from .decimals import format_decimal
from .nota_cliente_dialog import NotaClienteDialog
from .pdf_preview_dialog import PdfPreviewDialog
from .widgets import (
    compactar_alto_filas,
    crear_recuadro_destacado,
    redimensionar_pct_pantalla,
    texto_contacto_cliente,
)

CUIT_EMISOR_DEFAULT = "33703467909"  # ídem FacturadorWindow
TIPO_NC = 2
MOTIVO_DEV_MERC = 1

COL_CODIGO = 0
COL_DESCRIPCION = 1
COL_CANT_VENDIDA = 2
COL_CANT_DEVOLVER = 3
COL_PRECIO = 4
COL_IMPORTE = 5
COLUMNAS_RENGLONES = ["Sección/Código", "Descripción", "Cant. Vendida", "Cant. a Devolver", "Precio Unit.", "Importe a Acreditar"]


class NotaCreditoMercaderiaWindow(QMainWindow):
    TEXTO_BTN_ELEGIR_CLIENTE = "Elegir Cliente..."
    TEXTO_BTN_CAMBIAR_CLIENTE = "Cambiar Cliente"

    def __init__(self, parent: QWidget | None = None, *, afip: Optional[NumeracionYCAEProvider] = None):
        super().__init__(parent)
        self.setWindowTitle("Nota de Crédito — Devolución de Mercadería")
        redimensionar_pct_pantalla(self, 85, 80)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.factura_service = FacturaService(self.db)
        self.cuenta_corriente_service = CuentaCorrienteService(self.db)
        self.emision_service = EmisionNotaCreditoMercaderiaService(self.db)
        self.afip: NumeracionYCAEProvider = afip if afip is not None else crear_afip_provider()

        self.cliente_actual: Optional[Cliente] = None
        self.tiene_nota_cliente = False
        self._facturas_cliente: list[FcivaVta] = []
        self._usuario = os.environ.get("USERNAME", "SISTEMA")[:6]
        # Comprobante elegido en la ventana flotante "Aplicar a" (ver
        # `ComprobanteAplicarDialog`) — reemplaza la tabla completa que
        # antes vivía siempre visible (feedback del usuario, 2026-08-19,
        # mismo patrón que `NotaCreditoConceptoWindow`).
        self._comprobante_seleccionado: Optional[Ctascte] = None

        self._construir_ui()
        self._nueva()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_cabecera())

        fila_central = QHBoxLayout()
        fila_central.addWidget(self._armar_renglones(), stretch=2)
        fila_central.addWidget(self._armar_comprobante_imputar(), stretch=1)
        layout.addLayout(fila_central, stretch=1)

        layout.addWidget(self._armar_pie())

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QGroupBox:
        grupo = QGroupBox("Cabecera — Nota de Crédito por Devolución de Mercadería")
        layout = QVBoxLayout(grupo)

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
        self.btn_nueva = QPushButton("Nueva")
        self.btn_nueva.clicked.connect(self._nueva)
        fila_cliente.addWidget(self.btn_nueva)
        layout.addLayout(fila_cliente)

        # Localidad/Teléfono/Email del cliente elegido (ídem).
        self.lbl_cliente_contacto = QLabel("—")
        layout.addWidget(self.lbl_cliente_contacto)

        fila_factura = QHBoxLayout()
        fila_factura.addWidget(QLabel("Factura Original :"))
        self.combo_factura = QComboBox()
        self.combo_factura.currentIndexChanged.connect(self._on_factura_elegida)
        fila_factura.addWidget(self.combo_factura, stretch=1)
        # "Aplicar a..." abre la ventana flotante de selección del
        # comprobante contra el que se imputa esta Nota de Crédito — NO
        # tiene por qué ser la misma Factura Original que se devuelve
        # (ver docstring del módulo) — mismo patrón/diálogo que
        # `NotaCreditoConceptoWindow` (feedback del usuario, 2026-08-19).
        self.btn_aplicar_a = QPushButton("Aplicar a...")
        self.btn_aplicar_a.setEnabled(False)
        self.btn_aplicar_a.clicked.connect(self._on_aplicar_a)
        fila_factura.addWidget(self.btn_aplicar_a)
        fila_factura.addStretch()
        self.lbl_proximo_numero = QLabel("Próx. Nº: —")
        fila_factura.addWidget(self.lbl_proximo_numero)
        layout.addLayout(fila_factura)

        return grupo

    # ------------------------------------------------------------------
    # Renglones de la Factura original (checkbox = se devuelve)
    # ------------------------------------------------------------------
    def _armar_renglones(self) -> QGroupBox:
        self.grupo_renglones = QGroupBox(
            'Renglones de la Factura Original — "Cant. a Devolver" editable (0 = no se devuelve ese renglón)'
        )
        grupo = self.grupo_renglones
        layout = QVBoxLayout(grupo)

        self.tabla_renglones = QTableWidget(0, len(COLUMNAS_RENGLONES))
        self.tabla_renglones.setHorizontalHeaderLabels(COLUMNAS_RENGLONES)
        # Edición habilitada a nivel tabla, pero sólo la columna "Cant. a
        # Devolver" trae el flag `ItemIsEditable` por celda (ver
        # `_agregar_fila_renglon`) — el resto queda de sólo lectura pese
        # a este trigger.
        self.tabla_renglones.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.tabla_renglones.verticalHeader().setVisible(False)
        compactar_alto_filas(self.tabla_renglones)
        self.tabla_renglones.itemChanged.connect(self._on_cantidad_renglon_cambiada)
        layout.addWidget(self.tabla_renglones)

        return grupo

    def _on_factura_elegida(self) -> None:
        factura = self.combo_factura.currentData()
        self.tabla_renglones.blockSignals(True)
        self.tabla_renglones.setRowCount(0)
        self._renglones_reales = []
        if factura is not None:
            renglones = self.repos.fcestad1().by_comprobante(
                int(factura.TIPO) if factura.TIPO is not None else 1,
                factura.LETRA or "", factura.PTOVTA or 0, factura.CPBTE or 0,
            )
            for renglon in renglones:
                self._agregar_fila_renglon(renglon)
        self.tabla_renglones.blockSignals(False)
        self._recalcular_totales()

    def _agregar_fila_renglon(self, renglon) -> None:
        cod1 = (renglon.COD1 or "").strip()
        cod2 = (renglon.COD2 or "").strip()
        articulo = self.repos.articulo().by_cod1_cod2(cod1, cod2)
        descripcion = (articulo.DESCRI or "").strip() if articulo is not None else f"{cod1}/{cod2}"
        cantidad = renglon.CANT or Decimal("0")
        precio = renglon.PVTA or Decimal("0")
        importe = renglon.IMPTE or Decimal("0")
        # Sin Cant. real (ej. tela vendida por Mtr) no hay base para
        # prorratear (ver `renglon_devolucion_parcial`) — la fila queda
        # todo-o-nada: "Cant. a Devolver" se prellena en "1" (= incluir
        # el renglón completo), editable a "0" para excluirlo, pero
        # nunca partible.
        tiene_cantidad_real = cantidad > 0
        valor_prellenado = cantidad if tiene_cantidad_real else Decimal("1")

        fila = self.tabla_renglones.rowCount()
        self.tabla_renglones.insertRow(fila)

        self.tabla_renglones.setItem(fila, COL_CODIGO, self._item_solo_lectura(f"{cod1}/{cod2}"))
        self.tabla_renglones.setItem(fila, COL_DESCRIPCION, self._item_solo_lectura(descripcion))
        self.tabla_renglones.setItem(
            fila,
            COL_CANT_VENDIDA,
            self._item_solo_lectura(format_decimal(cantidad) if tiene_cantidad_real else "— (por Mtr)"),
        )

        item_cant_devolver = QTableWidgetItem(format_decimal(valor_prellenado))
        item_cant_devolver.setFlags(
            Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        )
        item_cant_devolver.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.tabla_renglones.setItem(fila, COL_CANT_DEVOLVER, item_cant_devolver)

        self.tabla_renglones.setItem(fila, COL_PRECIO, self._item_solo_lectura(f"$ {format_decimal(precio)}"))
        self.tabla_renglones.setItem(fila, COL_IMPORTE, self._item_solo_lectura(f"$ {format_decimal(importe)}"))
        self._renglones_reales.append(renglon)

    @staticmethod
    def _item_solo_lectura(texto: str) -> QTableWidgetItem:
        item = QTableWidgetItem(texto)
        item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        return item

    def _on_cantidad_renglon_cambiada(self, item: QTableWidgetItem) -> None:
        if item.column() != COL_CANT_DEVOLVER:
            return
        fila = item.row()
        if fila >= len(self._renglones_reales):
            return
        original = self._renglones_reales[fila]
        cantidad_original = original.CANT or Decimal("0")

        try:
            cantidad_editada = Decimal(item.text().strip().replace(",", "."))
        except (InvalidOperation, ValueError):
            cantidad_editada = cantidad_original

        if cantidad_editada < 0:
            cantidad_editada = Decimal("0")
        elif cantidad_original > 0 and cantidad_editada > cantidad_original:
            # Nunca por encima de lo efectivamente vendido — recorte
            # silencioso, mismo criterio que `renglon_devolucion_
            # parcial()` (que igual lo vuelve a aplicar al emitir).
            cantidad_editada = cantidad_original
        elif cantidad_original <= 0 and cantidad_editada > 0:
            # Fila todo-o-nada (sin Cant. real, ver `_agregar_fila_
            # renglon`) — cualquier valor no-cero se normaliza a "1"
            # ("incluir completo"), no hay partes intermedias.
            cantidad_editada = Decimal("1")

        renglon_prorrateado = renglon_devolucion_parcial(original, cantidad_editada)
        importe_nuevo = renglon_prorrateado.importe if renglon_prorrateado is not None else Decimal("0")

        self.tabla_renglones.blockSignals(True)
        item.setText(format_decimal(cantidad_editada))
        self.tabla_renglones.item(fila, COL_IMPORTE).setText(f"$ {format_decimal(importe_nuevo)}")
        self.tabla_renglones.blockSignals(False)

        self._recalcular_totales()

    def _renglones_elegidos(self) -> list[RenglonEmision]:
        resultado = []
        for fila, original in enumerate(getattr(self, "_renglones_reales", [])):
            item_cant = self.tabla_renglones.item(fila, COL_CANT_DEVOLVER)
            if item_cant is None:
                continue
            try:
                cantidad_a_devolver = Decimal(item_cant.text().strip().replace(",", "."))
            except (InvalidOperation, ValueError):
                continue

            renglon = renglon_devolucion_parcial(original, cantidad_a_devolver)
            if renglon is None:
                continue
            renglon.descripcion = self.tabla_renglones.item(fila, COL_DESCRIPCION).text()
            resultado.append(renglon)
        return resultado

    # ------------------------------------------------------------------
    # Comprobante a Imputar — panel achicado (feedback del usuario,
    # 2026-08-19: "El panel comprobantes a imputar achicarlo y colocar
    # el cpbte. seleccionado"), la selección en sí vive ahora en la
    # ventana flotante `ComprobanteAplicarDialog` (botón "Aplicar a...",
    # junto a "Factura Original" — ver `_armar_cabecera`).
    # ------------------------------------------------------------------
    def _armar_comprobante_imputar(self) -> QGroupBox:
        grupo = QGroupBox("Comprobante a Imputar")
        layout = QHBoxLayout(grupo)
        self.lbl_comprobante_imputar = QLabel("(sin comprobante seleccionado — usá \"Aplicar a...\")")
        layout.addWidget(self.lbl_comprobante_imputar, stretch=1)
        return grupo

    def _refrescar_panel_comprobante(self) -> None:
        comprobante = self._comprobante_seleccionado
        if comprobante is None:
            self.lbl_comprobante_imputar.setText('(sin comprobante seleccionado — usá "Aplicar a...")')
            return
        etiqueta_tipo = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
        fecha_txt = comprobante.FECHA.strftime("%d/%m/%Y") if comprobante.FECHA else "—"
        self.lbl_comprobante_imputar.setText(
            f"{etiqueta_tipo} {(comprobante.PREFIJO or 0):04d}-{(comprobante.CPBTE or 0):08d} — {fecha_txt} — "
            f"Importe sin aplicar: $ {format_decimal(comprobante.DEBE or Decimal('0'))}"
        )

    def _on_aplicar_a(self) -> None:
        if self.cliente_actual is None:
            return
        elegido = ComprobanteAplicarDialog.elegir(
            self.repos, self.cuenta_corriente_service, self.cliente_actual.CODIGO, parent=self
        )
        if elegido is None:
            return
        self._comprobante_seleccionado = elegido
        self._refrescar_panel_comprobante()
        self._recalcular_totales()

    # ------------------------------------------------------------------
    # Pie
    # ------------------------------------------------------------------
    def _armar_pie(self) -> QGroupBox:
        grupo = QGroupBox("Totales")
        layout = QHBoxLayout(grupo)

        layout.addWidget(QLabel("Subtotal:"))
        self.lbl_subtotal = QLabel("$ 0,00")
        layout.addWidget(self.lbl_subtotal)
        layout.addWidget(QLabel("Neto Gravado:"))
        self.lbl_neto = QLabel("$ 0,00")
        layout.addWidget(self.lbl_neto)
        layout.addWidget(QLabel("IVA:"))
        self.lbl_iva = QLabel("$ 0,00")
        layout.addWidget(self.lbl_iva)
        layout.addStretch()

        recuadro_total, self.lbl_total = crear_recuadro_destacado("TOTAL:")
        layout.addWidget(recuadro_total)
        layout.addStretch()

        self.btn_emitir = QPushButton("Emitir")
        self.btn_emitir.setStyleSheet("font-weight: bold; padding: 8px;")
        self.btn_emitir.setEnabled(False)
        self.btn_emitir.clicked.connect(self._on_emitir)
        layout.addWidget(self.btn_emitir)

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
        self.combo_factura.blockSignals(True)
        self.combo_factura.clear()
        # Cambiar de Cliente invalida cualquier comprobante ya elegido
        # en "Aplicar a..." (era de OTRO Cliente).
        self._comprobante_seleccionado = None
        self._refrescar_panel_comprobante()
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.lbl_cliente_contacto.setText("—")
            self.btn_nota_cliente.setEnabled(False)
            self.btn_aplicar_a.setEnabled(False)
            self.tiene_nota_cliente = False
            self.btn_nota_cliente.setStyleSheet("")
            self.btn_elegir_cliente.setText(self.TEXTO_BTN_ELEGIR_CLIENTE)
            self.combo_factura.blockSignals(False)
            self.tabla_renglones.setRowCount(0)
            self._refrescar_proximo_numero()
            self._recalcular_totales()
            return

        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()} — CUIT {cliente.CUIT or 's/d'}")
        self.lbl_cliente_contacto.setText(texto_contacto_cliente(cliente))
        self.btn_nota_cliente.setEnabled(True)
        self.btn_aplicar_a.setEnabled(True)
        self.btn_elegir_cliente.setText(self.TEXTO_BTN_CAMBIAR_CLIENTE)

        self._actualizar_boton_nota_cliente()
        if self.tiene_nota_cliente:
            NotaClienteDialog(self.repos, cliente.CODIGO, parent=self).exec()
            self._actualizar_boton_nota_cliente()

        self._facturas_cliente = self.repos.fciva_vta().facturas_de_cliente(cliente.CODIGO)
        for factura in self._facturas_cliente:
            etiqueta = f"{(factura.LETRA or '').strip()} {(factura.PTOVTA or 0):04d}-{(factura.CPBTE or 0):08d} — {factura.FECHA.strftime('%d/%m/%Y') if factura.FECHA else ''}"
            self.combo_factura.addItem(etiqueta, factura)
        self.combo_factura.blockSignals(False)

        self._refrescar_proximo_numero()
        if self.combo_factura.count():
            self._on_factura_elegida()
        else:
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

    # ------------------------------------------------------------------
    # Numeración / totales
    # ------------------------------------------------------------------
    def _letra_actual(self) -> Optional[str]:
        if self.cliente_actual is None:
            return None
        return self.factura_service.letra_comprobante(self.cliente_actual.CIVA or 0, self.cliente_actual.PCIA)

    def _refrescar_proximo_numero(self) -> None:
        letra = self._letra_actual()
        if letra is None:
            self.lbl_proximo_numero.setText("Próx. Nº: —")
            return
        tipo_cbte = codigo_afip(letra, TIPO_NC)
        punto_venta = punto_venta_por_tipo(TIPO_NC)
        try:
            ultimo = self.afip.ultimo_comprobante(punto_venta, tipo_cbte)
            self.lbl_proximo_numero.setText(f"Próx. Nº: {punto_venta:04d}-{ultimo + 1:08d} (Letra {letra})")
        except Exception:  # noqa: BLE001 — sólo informativo
            self.lbl_proximo_numero.setText(f"Próx. Nº: (no disponible, Letra {letra})")

    def _recalcular_totales(self) -> None:
        # No habilita el Detalle (los renglones a devolver) hasta elegir
        # el comprobante a imputar (feedback del usuario, 2026-08-19:
        # "Si no selecciona la factura o ND no habilite detalle").
        self.grupo_renglones.setEnabled(self._comprobante_seleccionado is not None)

        renglones = self._renglones_elegidos()
        bruto = sum((r.importe for r in renglones), Decimal("0"))
        civa = self.cliente_actual.CIVA if self.cliente_actual else 3
        total = self.factura_service.calcular_total(bruto=bruto, descuento=Decimal("0"), civa_cliente=civa or 3)

        self.lbl_subtotal.setText(f"$ {format_decimal(total.bruto)}")
        self.lbl_neto.setText(f"$ {format_decimal(total.neto_gravado)}")
        self.lbl_iva.setText(f"$ {format_decimal(total.iva)}")
        self.lbl_total.setText(f"$ {format_decimal(total.total)}")

        habilitado = (
            self.cliente_actual is not None
            and total.total > 0
            and self._comprobante_seleccionado is not None
        )
        self.btn_emitir.setEnabled(habilitado)

    # ------------------------------------------------------------------
    # Nueva
    # ------------------------------------------------------------------
    def _nueva(self) -> None:
        self.cliente_actual = None
        self._comprobante_seleccionado = None
        self._renglones_reales = []
        self._refrescar_cliente()
        QTimer.singleShot(0, self._on_elegir_cliente)

    # ------------------------------------------------------------------
    # Emitir
    # ------------------------------------------------------------------
    def _on_emitir(self) -> None:
        if self.cliente_actual is None:
            QMessageBox.warning(self, "Nota de Crédito", "Elegí un Cliente.")
            return
        if not (self.cliente_actual.CUIT or "").strip():
            QMessageBox.warning(self, "Nota de Crédito", "El cliente no tiene CUIT cargado — no se puede solicitar CAE.")
            return

        renglones = self._renglones_elegidos()
        if not renglones:
            QMessageBox.warning(self, "Nota de Crédito", "Tildá al menos un renglón para devolver.")
            return

        comprobante_a_imputar = self._comprobante_seleccionado
        if comprobante_a_imputar is None:
            QMessageBox.warning(self, "Nota de Crédito", "Elegí el comprobante contra el cual se imputa esta Nota de Crédito.")
            return

        letra = self._letra_actual()
        tipo_cbte = codigo_afip(letra, TIPO_NC)
        punto_venta = punto_venta_por_tipo(TIPO_NC)
        bruto = sum((r.importe for r in renglones), Decimal("0"))
        total = self.factura_service.calcular_total(bruto=bruto, descuento=Decimal("0"), civa_cliente=self.cliente_actual.CIVA or 3)
        fecha = date.today()

        try:
            numero = self.afip.ultimo_comprobante(punto_venta, tipo_cbte) + 1
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Nota de Crédito", f"Error al conectar con AFIP:\n{exc}")
            return

        # Boceto → Grabar (mismo patrón que `FacturadorWindow._on_emitir`,
        # pedido del usuario 2026-08-20 para las 2 ventanas de NC) — se
        # muestra ANTES de pedir el CAE, con sello "BORRADOR" (sin
        # CAE/QR reales todavía).
        try:
            datos_borrador = self._datos_pdf(letra, punto_venta, numero, fecha, renglones, total)
            ruta_borrador = generar_pdf_factura(datos_borrador)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Nota de Crédito", f"No se pudo generar el boceto del PDF:\n{exc}")
            return

        vista_previa = PdfPreviewDialog(
            ruta_borrador, titulo=f"Boceto — Nota de Crédito {letra} (sin CAE todavía)",
            mostrar_boton_grabar=True, parent=self,
        )
        vista_previa.exec()
        if not vista_previa.grabar:
            return

        respuesta = QMessageBox.question(
            self,
            "Nota de Crédito",
            f"¿Desea grabar la Nota de Crédito {letra} por $ {format_decimal(total.total)}?",
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
                importe_iibb=Decimal("0"),
                importe_total=total.total,
                fecha_cbte=fecha,
                condicion_iva_receptor_id=condicion_iva_receptor(self.cliente_actual.CIVA),
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Nota de Crédito", f"Error al conectar con AFIP:\n{exc}")
            return

        if not resultado_cae.aprobado:
            QMessageBox.critical(self, "Nota de Crédito", f"AFIP no aprobó el comprobante.\nMotivo: {resultado_cae.motivo}")
            return

        try:
            self.emision_service.emitir(
                cliente=self.cliente_actual,
                letra=letra,
                punto_venta=punto_venta,
                numero_comprobante=numero,
                renglones=renglones,
                total=total,
                motivo=MOTIVO_DEV_MERC,
                usuario=self._usuario,
                comprobante_a_imputar=comprobante_a_imputar,
                fecha=fecha,
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self,
                "Nota de Crédito",
                f"AFIP aprobó el CAE {resultado_cae.cae} pero la Nota de Crédito NO se pudo grabar en el sistema:\n{exc}\n"
                "Anotar el CAE para grabarla manualmente.",
            )
            return

        qr_url = generar_qr_afip(
            cuit_emisor=CUIT_EMISOR_DEFAULT, punto_venta=punto_venta, tipo_cbte=tipo_cbte, nro_cbte=numero,
            importe_total=total.total, tipo_doc_receptor=80, nro_doc_receptor=self.cliente_actual.CUIT,
            cae=resultado_cae.cae, fecha_cbte=fecha,
        )
        ruta_pdf = None
        try:
            datos_pdf = self._datos_pdf(
                letra, punto_venta, numero, fecha, renglones, total,
                cae=resultado_cae.cae, cae_vencimiento=resultado_cae.vencimiento, qr_url=qr_url,
                observaciones_afip=resultado_cae.motivo,
            )
            ruta_pdf = generar_pdf_factura(datos_pdf)
        except Exception as exc:  # noqa: BLE001 — YA se grabó, esto no debe revertirlo
            QMessageBox.warning(self, "Nota de Crédito", f"Se emitió y grabó correctamente, pero el PDF falló:\n{exc}")

        QMessageBox.information(
            self,
            "Nota de Crédito",
            f"Nota de Crédito {letra} {punto_venta:04d}-{numero:08d} emitida.\n"
            f"CAE: {resultado_cae.cae}\nVencimiento: {resultado_cae.vencimiento}\nTotal: $ {format_decimal(total.total)}",
        )
        if ruta_pdf is not None:
            PdfPreviewDialog(
                ruta_pdf, titulo=f"Nota de Crédito {letra} {punto_venta:04d}-{numero:08d} — CAE {resultado_cae.cae}",
                parent=self,
            ).exec()
        self._nueva()

    def _datos_pdf(
        self, letra, punto_venta, numero, fecha, renglones, total,
        *, cae=None, cae_vencimiento=None, qr_url=None, observaciones_afip="",
    ) -> DatosFacturaPDF:
        """Réplica del layout ya construido para Factura — reusa
        `DatosFacturaPDF`/`generar_pdf_factura` (mismo criterio de
        renglones+precio real que ya usa esta ventana, ver docstring del
        módulo), con `titulo_comprobante` propio (réplica real,
        `EmiFact.frm:1282`: "NOTA DE CRÉDITO", sin el espaciado
        decorativo que sólo usa Factura) y los mismos datos de Cliente/
        "Aplicada a" agregados 2026-08-20 (hallazgo real comparando
        contra una Factura real de muestra)."""
        cliente = self.cliente_actual
        comprobante = self._comprobante_seleccionado
        return DatosFacturaPDF(
            letra=letra, punto_venta=punto_venta, numero=numero, fecha=fecha,
            cliente_codigo=cliente.CODIGO,
            cliente_nombre=(cliente.NOMB or "").strip(),
            cliente_cuit=cliente.CUIT or "",
            cliente_civa=cliente.CIVA or 0,
            cliente_domicilio=(cliente.DIR or "").strip(),
            cliente_localidad=(cliente.LOC or "").strip(),
            cliente_cp=(cliente.CP or "").strip(),
            cliente_provincia=nombre_provincia(cliente.PCIA),
            condicion_venta=condicion_venta_texto(self.repos, cliente.CVTA),
            renglones=list(renglones), total=total,
            cae=cae, cae_vencimiento=cae_vencimiento, qr_url=qr_url,
            observaciones_afip=observaciones_afip,
            titulo_comprobante="NOTA DE CRÉDITO",
            aplicada_a=(comprobante.PREFIJO or 0, comprobante.CPBTE or 0, comprobante.FECHA) if comprobante else None,
        )

    # ------------------------------------------------------------------
    def _hay_algo_cargado(self) -> bool:
        return bool(self._renglones_elegidos())

    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        if self._hay_algo_cargado():
            respuesta = QMessageBox.question(
                self,
                "Nota de Crédito",
                "Hay renglones tildados para devolver — se va a perder el ingreso.\n¿Desea salir igual?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if respuesta != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.db.close()
        super().closeEvent(event)
