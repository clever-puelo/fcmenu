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
- Selección de renglones por **renglón completo** (checkbox), no por
  cantidad parcial editable — simplificación deliberada de esta primera
  versión (avisar si hace falta devolución parcial de un renglón, se
  puede agregar después).
- Emitida la Nota de Crédito, el Stock del artículo se repone
  automáticamente (`EmisionNotaCreditoMercaderiaService`, reversa real
  de `EmiFact.frm:1864-1928`) y se imputa contra un comprobante con
  deuda elegido por el operador (réplica de `ImputFC.frm`, mismo
  mecanismo que `NotaCreditoConceptoWindow`) — no necesariamente la
  misma Factura que se está devolviendo, igual que en el legacy.

**Fuera de alcance de esta primera versión** (igual que "Concepto
Libre"): sin vista previa en PDF, sin percepción de IIBB (el legacy la
permite estructuralmente pero no es un caso real de negocio para una
devolución — avisar si hace falta).
"""

from __future__ import annotations

import os
from datetime import date
from decimal import Decimal
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
    punto_venta_por_tipo,
)
from migration.db import get_session
from migration.models import Cliente, Ctascte, FcivaVta
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import (
    CuentaCorrienteService,
    EmisionNotaCreditoMercaderiaService,
    FacturaService,
    RenglonEmision,
)

from .cliente_busqueda_window import ClienteBusquedaWindow
from .decimals import format_decimal
from .nota_cliente_dialog import NotaClienteDialog
from .widgets import TablaBusqueda, compactar_alto_filas, crear_recuadro_destacado, redimensionar_pct_pantalla

CUIT_EMISOR_DEFAULT = "33703467909"  # ídem FacturadorWindow
TIPO_NC = 2
MOTIVO_DEV_MERC = 1

COL_CHECK = 0
COL_CODIGO = 1
COL_DESCRIPCION = 2
COL_CANTIDAD = 3
COL_PRECIO = 4
COL_IMPORTE = 5
COLUMNAS_RENGLONES = ["", "Sección/Código", "Descripción", "Cantidad", "Precio Unit.", "Importe"]

COL_COMPROBANTE = 0
COL_TIPO = 1
COL_FECHA = 2
COL_FECVTO = 3
COL_DEBE = 4
COLUMNAS_PENDIENTES = ["Comprobante", "Tipo", "Fecha", "Fec.Vto.", "Debe"]


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
        self.btn_nueva.clicked.connect(self._nueva)
        fila_cliente.addWidget(self.btn_nueva)
        layout.addLayout(fila_cliente)

        fila_factura = QHBoxLayout()
        fila_factura.addWidget(QLabel("Factura Original :"))
        self.combo_factura = QComboBox()
        self.combo_factura.currentIndexChanged.connect(self._on_factura_elegida)
        fila_factura.addWidget(self.combo_factura, stretch=1)
        fila_factura.addStretch()
        self.lbl_proximo_numero = QLabel("Próx. Nº: —")
        fila_factura.addWidget(self.lbl_proximo_numero)
        layout.addLayout(fila_factura)

        return grupo

    # ------------------------------------------------------------------
    # Renglones de la Factura original (checkbox = se devuelve)
    # ------------------------------------------------------------------
    def _armar_renglones(self) -> QGroupBox:
        grupo = QGroupBox("Renglones de la Factura Original — tildá los que se devuelven")
        layout = QVBoxLayout(grupo)

        self.tabla_renglones = QTableWidget(0, len(COLUMNAS_RENGLONES))
        self.tabla_renglones.setHorizontalHeaderLabels(COLUMNAS_RENGLONES)
        self.tabla_renglones.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabla_renglones.verticalHeader().setVisible(False)
        compactar_alto_filas(self.tabla_renglones)
        self.tabla_renglones.itemChanged.connect(self._on_check_renglon_cambiado)
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

        fila = self.tabla_renglones.rowCount()
        self.tabla_renglones.insertRow(fila)
        item_check = QTableWidgetItem()
        item_check.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        item_check.setCheckState(Qt.CheckState.Unchecked)
        self.tabla_renglones.setItem(fila, COL_CHECK, item_check)
        self.tabla_renglones.setItem(fila, COL_CODIGO, QTableWidgetItem(f"{cod1}/{cod2}"))
        self.tabla_renglones.setItem(fila, COL_DESCRIPCION, QTableWidgetItem(descripcion))
        self.tabla_renglones.setItem(fila, COL_CANTIDAD, QTableWidgetItem(format_decimal(cantidad) if cantidad else "1"))
        self.tabla_renglones.setItem(fila, COL_PRECIO, QTableWidgetItem(f"$ {format_decimal(precio)}"))
        self.tabla_renglones.setItem(fila, COL_IMPORTE, QTableWidgetItem(f"$ {format_decimal(importe)}"))
        self._renglones_reales.append(renglon)

    def _on_check_renglon_cambiado(self, item: QTableWidgetItem) -> None:
        if item.column() == COL_CHECK:
            self._recalcular_totales()

    def _renglones_elegidos(self) -> list[RenglonEmision]:
        resultado = []
        for fila, renglon in enumerate(getattr(self, "_renglones_reales", [])):
            item_check = self.tabla_renglones.item(fila, COL_CHECK)
            if item_check is None or item_check.checkState() != Qt.CheckState.Checked:
                continue
            resultado.append(
                RenglonEmision(
                    cod1=(renglon.COD1 or "").strip(),
                    cod2=(renglon.COD2 or "").strip(),
                    descripcion=self.tabla_renglones.item(fila, COL_DESCRIPCION).text(),
                    precio_unitario=renglon.PVTA or Decimal("0"),
                    importe=renglon.IMPTE or Decimal("0"),
                    pulg=renglon.PULG or Decimal("0"),
                    mtr=renglon.MTR or Decimal("0"),
                    milim=renglon.MILIM or 0,
                    telas=renglon.TELAS or 0,
                    cantidad_unidades=renglon.CANT or Decimal("0"),
                )
            )
        return resultado

    # ------------------------------------------------------------------
    # Comprobante a Imputar
    # ------------------------------------------------------------------
    def _armar_comprobante_imputar(self) -> QGroupBox:
        grupo = QGroupBox("Comprobante a Imputar")
        layout = QVBoxLayout(grupo)
        self.tabla_pendientes = TablaBusqueda(COLUMNAS_PENDIENTES, columnas_derecha=(COL_COMPROBANTE, COL_DEBE))
        self.tabla_pendientes.itemSelectionChanged.connect(self._recalcular_totales)
        layout.addWidget(self.tabla_pendientes)
        return grupo

    def _cargar_pendientes(self) -> None:
        if self.cliente_actual is None:
            self.tabla_pendientes.cargar_filas([])
            return
        pendientes = self.cuenta_corriente_service.facturas_pendientes(self.cliente_actual.CODIGO)
        filas = [
            (
                [
                    f"{(p.PREFIJO or 0):04d}-{(p.CPBTE or 0):08d}",
                    ETIQUETAS_TIPO_CTASCTE.get(p.TIPO, "—"),
                    p.FECHA.strftime("%d/%m/%Y") if p.FECHA else "",
                    p.FECVTO.strftime("%d/%m/%Y") if p.FECVTO else "",
                    format_decimal(p.DEBE or Decimal("0")),
                ],
                p,
            )
            for p in pendientes
        ]
        self.tabla_pendientes.cargar_filas(filas)

    def _comprobante_a_imputar(self) -> Optional[Ctascte]:
        return self.tabla_pendientes.objeto_seleccionado()

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
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.btn_nota_cliente.setEnabled(False)
            self.tiene_nota_cliente = False
            self.btn_nota_cliente.setStyleSheet("")
            self.btn_elegir_cliente.setText(self.TEXTO_BTN_ELEGIR_CLIENTE)
            self.combo_factura.blockSignals(False)
            self.tabla_renglones.setRowCount(0)
            self._cargar_pendientes()
            self._refrescar_proximo_numero()
            self._recalcular_totales()
            return

        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()} — CUIT {cliente.CUIT or 's/d'}")
        self.btn_nota_cliente.setEnabled(True)
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

        self._cargar_pendientes()
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
            and self._comprobante_a_imputar() is not None
        )
        self.btn_emitir.setEnabled(habilitado)

    # ------------------------------------------------------------------
    # Nueva
    # ------------------------------------------------------------------
    def _nueva(self) -> None:
        self.cliente_actual = None
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

        comprobante_a_imputar = self._comprobante_a_imputar()
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

        QMessageBox.information(
            self,
            "Nota de Crédito",
            f"Nota de Crédito {letra} {punto_venta:04d}-{numero:08d} emitida.\n"
            f"CAE: {resultado_cae.cae}\nVencimiento: {resultado_cae.vencimiento}\nTotal: $ {format_decimal(total.total)}",
        )
        self._nueva()

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
