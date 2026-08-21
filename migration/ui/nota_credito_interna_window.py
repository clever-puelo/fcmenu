"""`NotaCreditoInternaWindow` — Nota de Crédito Interna, migración de
`NCInterna.frm` (847 líneas, analizado íntegro 2026-08-20).

**Distinta de las otras 2 pantallas de Nota de Crédito ya migradas**
(`NotaCreditoConceptoWindow`/`NotaCreditoMercaderiaWindow`, confirmado
leyendo el `.frm` completo, no asumido por el nombre parecido): NO es un
comprobante fiscal — Letra "X" fija, nunca pide CAE, nunca toca AFIP ni
`FcivaVta`/`Totales`. Es un ajuste puro de cuenta corriente: cancela de
una sola vez TODA la deuda pendiente (Facturas/ND con saldo) de un
Cliente, contra un único comprobante interno nuevo (`Ctascte.TIPO=8`,
"NCInt") — ver `EmisionNotaCreditoInternaService` para el detalle
completo del mecanismo y las decisiones confirmadas.

Flujo (réplica de `Form_Load`/`Command3_Click`/`BuscaDeuda`/`AGrabar`):
1. Elegir Cliente → se cargan automáticamente TODAS sus Facturas/ND con
   deuda (sin selección posible — es todo o nada, a diferencia de
   Recibo/NC que sí eligen qué imputar).
2. Motivo (misma tabla `Fctabla1 CTAB='MT'` que usan las otras 2
   pantallas de NC) + Nota (mínimo 7 caracteres, único gateo real del
   legacy — ver docstring de `EmisionNotaCreditoInternaService`).
3. "Cancelar Deuda" → confirma (mismo texto de advertencia del legacy:
   "Esta acción cancelará la deuda del cliente") → graba → muestra el
   PDF del documento interno (`generar_pdf_nota_credito_interna`).
"""

from __future__ import annotations

import os
from datetime import date
from typing import Optional

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QComboBox,
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
from migration.decimals import format_decimal
from migration.models import Cliente, Ctascte
from migration.pdf import DatosNotaCreditoInternaPDF, generar_pdf_nota_credito_interna
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import AplicacionPago, EmisionNotaCreditoInternaService

from .cliente_busqueda_window import ClienteBusquedaWindow
from .pdf_preview_dialog import PdfPreviewDialog
from .widgets import redimensionar_pct_pantalla, texto_contacto_cliente

COLUMNAS_DEUDA = ["Tipo", "Nro.", "Fecha", "Importe", "Deuda"]
COLUMNAS_ALINEADAS_DERECHA = {3, 4}


class NotaCreditoInternaWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Nota de Crédito Interna")
        redimensionar_pct_pantalla(self, 65, 65)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.servicio = EmisionNotaCreditoInternaService(self.db)

        self.cliente_actual: Optional[Cliente] = None
        self.deuda_pendiente: list[Ctascte] = []

        self._construir_ui()
        self._nueva()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_cabecera())
        layout.addWidget(self._armar_deuda(), stretch=1)
        layout.addWidget(self._armar_pie())

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QGroupBox:
        grupo = QGroupBox("Cliente — cancela TODA su deuda pendiente (sin CAE, documento interno)")
        layout = QVBoxLayout(grupo)

        fila_cliente = QHBoxLayout()
        self.lbl_cliente = QLabel("(sin cliente elegido)")
        fila_cliente.addWidget(self.lbl_cliente, stretch=1)
        self.btn_elegir_cliente = QPushButton("Elegir Cliente...")
        self.btn_elegir_cliente.clicked.connect(self._on_elegir_cliente)
        fila_cliente.addWidget(self.btn_elegir_cliente)
        layout.addLayout(fila_cliente)

        self.lbl_cliente_contacto = QLabel("—")
        layout.addWidget(self.lbl_cliente_contacto)

        # "Deuda Total" — mismo rojo/negrita destacado del Label1 real
        # (`NCInterna.frm`, ForeColor rojo, Tahoma 14.25 negrita).
        self.lbl_deuda_total = QLabel("Deuda Total: —")
        self.lbl_deuda_total.setStyleSheet("color: #c62828; font-weight: bold; font-size: 16pt;")
        layout.addWidget(self.lbl_deuda_total)

        return grupo

    # ------------------------------------------------------------------
    # Deuda a cancelar (sólo lectura — todo o nada, sin selección)
    # ------------------------------------------------------------------
    def _armar_deuda(self) -> QGroupBox:
        grupo = QGroupBox("Comprobantes que se van a cancelar")
        layout = QVBoxLayout(grupo)

        self.tabla_deuda = QTableWidget(0, len(COLUMNAS_DEUDA))
        self.tabla_deuda.setHorizontalHeaderLabels(COLUMNAS_DEUDA)
        self.tabla_deuda.verticalHeader().setVisible(False)
        self.tabla_deuda.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabla_deuda.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        layout.addWidget(self.tabla_deuda, stretch=1)

        return grupo

    # ------------------------------------------------------------------
    # Pie — Motivo, Nota, acciones
    # ------------------------------------------------------------------
    def _armar_pie(self) -> QGroupBox:
        grupo = QGroupBox("Motivo y Nota")
        layout = QVBoxLayout(grupo)

        fila_motivo = QHBoxLayout()
        fila_motivo.addWidget(QLabel("Motivo :"))
        self.combo_motivo = QComboBox()
        self._poblar_motivos()
        self.combo_motivo.currentIndexChanged.connect(self._actualizar_boton_grabar)
        fila_motivo.addWidget(self.combo_motivo, stretch=1)
        layout.addLayout(fila_motivo)

        fila_nota = QHBoxLayout()
        fila_nota.addWidget(QLabel("Nota :"))
        # Mismo límite de 100 caracteres del legacy (`Text1.MaxLength`);
        # una sola línea en vez de la caja multilínea de 2 renglones del
        # `.frm` — simplificación deliberada, el texto se sigue
        # imprimiendo completo en el PDF.
        self.txt_nota = QLineEdit()
        self.txt_nota.setMaxLength(100)
        self.txt_nota.textChanged.connect(self._actualizar_boton_grabar)
        fila_nota.addWidget(self.txt_nota, stretch=1)
        layout.addLayout(fila_nota)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        self.btn_grabar = QPushButton("Cancelar Deuda")
        self.btn_grabar.setStyleSheet("font-weight: bold; padding: 6px;")
        self.btn_grabar.setEnabled(False)
        self.btn_grabar.clicked.connect(self._on_grabar)
        fila_botones.addWidget(self.btn_grabar)

        fila_botones.addSpacing(12)
        btn_salir = QPushButton("Salir")
        btn_salir.clicked.connect(self.close)
        fila_botones.addWidget(btn_salir)
        layout.addLayout(fila_botones)

        return grupo

    def _poblar_motivos(self) -> None:
        self.combo_motivo.clear()
        self.combo_motivo.addItem("(elegir)", None)
        for fila in self.repos.fctablas().by_ctab("MT"):
            cod = (fila.COD or "").strip()
            if not cod.isdigit():
                continue
            self.combo_motivo.addItem(f"{cod} - {fila.DESCRI or ''}", int(cod))

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
            self.btn_elegir_cliente.setText("Elegir Cliente...")
            self._cargar_deuda()
            return

        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()}")
        self.lbl_cliente_contacto.setText(texto_contacto_cliente(cliente))
        self.btn_elegir_cliente.setText("Cambiar Cliente")
        self._cargar_deuda()

    def _cargar_deuda(self) -> None:
        self.deuda_pendiente = []
        self.tabla_deuda.setRowCount(0)
        if self.cliente_actual is None:
            self.lbl_deuda_total.setText("Deuda Total: —")
            self._actualizar_boton_grabar()
            return

        self.deuda_pendiente = self.servicio.deuda_pendiente(self.cliente_actual.CODIGO)
        saldo = sum((c.DEBE or 0 for c in self.deuda_pendiente), 0)

        for fila_idx, comprobante in enumerate(self.deuda_pendiente):
            self.tabla_deuda.insertRow(fila_idx)
            valores = [
                ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—"),
                str(comprobante.CPBTE or ""),
                comprobante.FECHA.strftime("%d/%m/%Y") if comprobante.FECHA else "",
                format_decimal(comprobante.IMPTE or 0),
                format_decimal(comprobante.DEBE or 0),
            ]
            for col, texto in enumerate(valores):
                item = QTableWidgetItem(texto)
                if col in COLUMNAS_ALINEADAS_DERECHA:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.tabla_deuda.setItem(fila_idx, col, item)
        self.tabla_deuda.resizeColumnsToContents()

        self.lbl_deuda_total.setText(
            f"Deuda Total: $ {format_decimal(saldo)}" if saldo > 0 else "Deuda Total: (sin deuda pendiente)"
        )
        self._actualizar_boton_grabar()

    # ------------------------------------------------------------------
    def _actualizar_boton_grabar(self) -> None:
        habilitado = (
            self.cliente_actual is not None
            and bool(self.deuda_pendiente)
            and self.combo_motivo.currentData() is not None
            and len(self.txt_nota.text()) >= EmisionNotaCreditoInternaService.LARGO_MINIMO_NOTA
        )
        self.btn_grabar.setEnabled(habilitado)

    # ------------------------------------------------------------------
    # Nueva
    # ------------------------------------------------------------------
    def _nueva(self) -> None:
        self.cliente_actual = None
        self.combo_motivo.setCurrentIndex(0)
        self.txt_nota.clear()
        self._refrescar_cliente()
        QTimer.singleShot(0, self._on_elegir_cliente)

    # ------------------------------------------------------------------
    # Cancelar Deuda
    # ------------------------------------------------------------------
    def _on_grabar(self) -> None:
        if self.cliente_actual is None or not self.deuda_pendiente:
            return

        respuesta = QMessageBox.question(
            self,
            "Nota de Crédito Interna",
            "¿Desea GRABAR la Nota de Crédito Interna?\n(Esta acción cancelará la deuda del cliente)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        motivo_id = self.combo_motivo.currentData()
        motivo_texto = self.combo_motivo.currentText()
        nota = self.txt_nota.text()
        fecha = date.today()

        # Snapshot ANTES de cancelar_deuda() — después, DEBE de cada
        # comprobante queda en 0 (ya cancelado).
        aplicaciones = [
            AplicacionPago(comprobante=c, importe_aplicado=c.DEBE or 0) for c in self.deuda_pendiente
        ]

        try:
            resultado = self.servicio.cancelar_deuda(
                cliente=self.cliente_actual,
                motivo=motivo_id,
                nota=nota,
                usuario=self._usuario(),
                fecha=fecha,
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Nota de Crédito Interna", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Nota de Crédito Interna", f"No se pudo grabar:\n{exc}")
            return

        try:
            cliente = self.cliente_actual
            datos_pdf = DatosNotaCreditoInternaPDF(
                numero=resultado.numero,
                fecha=fecha,
                cliente_codigo=cliente.CODIGO,
                cliente_nombre=(cliente.NOMB or "").strip(),
                cliente_cuit=cliente.CUIT or "",
                cliente_domicilio=(cliente.DIR or "").strip(),
                cliente_cp=(cliente.CP or "").strip(),
                cliente_localidad=(cliente.LOC or "").strip(),
                cliente_provincia_codigo=(cliente.PCIA or "").strip(),
                cliente_civa=cliente.CIVA or 0,
                motivo_texto=motivo_texto,
                nota=nota,
                comprobantes_cancelados=aplicaciones,
                total_cancelado=resultado.total_cancelado,
            )
            ruta_pdf = generar_pdf_nota_credito_interna(datos_pdf)
        except Exception as exc:  # noqa: BLE001 — YA se grabó, esto no debe revertirlo
            QMessageBox.warning(
                self, "Nota de Crédito Interna", f"Se grabó correctamente, pero el PDF falló:\n{exc}"
            )
            self._nueva()
            return

        PdfPreviewDialog(
            ruta_pdf,
            titulo=f"Nota de Crédito Interna X {resultado.numero:06d}",
            parent=self,
        ).exec()
        self._nueva()

    def _usuario(self) -> str:
        return os.environ.get("USERNAME", "SISTEMA")[:6]

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
