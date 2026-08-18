"""Diálogo de detalle de un Cheque — migración de `VerCheq.frm Sub
MuestraDetalle` (líneas 1291-1377), con el panel de Egreso (`Command3`/
`Command4`, líneas 1044-1107) y el borrado (`Command5`/`BorraCheque`/
`GrabaAnul`, líneas 1111-1465).

**Borrado simplificado** (ver docstring de `ChequeService`): confirmación
Yes/No estándar en vez de re-tipear un código de 5 letras al azar — sin
impacto real, `Cheques` está vacía en los datos reales.
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.decimals import format_decimal
from migration.models import Cheque
from migration.repository import RepositoryFactory
from migration.services import ChequeService

from .widgets import EnteroLineEdit, EnterAsTabFilter, UpperCaseLineEdit, crear_boton_hoy


class ChequeDetalleDialog(QDialog):
    """`self.borrado` queda en `True` si el operador eliminó el cheque —
    la ventana que abrió este diálogo debe refrescar su grilla."""

    def __init__(self, repos: RepositoryFactory, cheque: Cheque, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.cheque_service = ChequeService(repos.db)
        self.cheque = cheque
        self.borrado = False

        self.setWindowTitle(f"Cheque Nº {cheque.NROCHEQ}")
        self.resize(560, 520)

        self._construir_ui()
        self._cargar_datos()
        EnterAsTabFilter(self, boton_final=self.btn_grabar_egreso)

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        layout.addWidget(self._armar_cliente())
        layout.addWidget(self._armar_cheque())
        layout.addWidget(self._armar_egreso())

        fila_botones = QHBoxLayout()
        self.btn_borrar = QPushButton("Borrar")
        self.btn_borrar.clicked.connect(self._on_borrar)
        fila_botones.addWidget(self.btn_borrar)
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _armar_cliente(self) -> QGroupBox:
        grupo = QGroupBox("Cliente")
        form = QFormLayout(grupo)
        self.lbl_cliente_codigo = QLabel()
        self.lbl_cliente_nombre = QLabel()
        self.lbl_cliente_dir = QLabel()
        self.lbl_cliente_civa = QLabel()
        self.lbl_cliente_cuit = QLabel()
        form.addRow("Código :", self.lbl_cliente_codigo)
        form.addRow("Nombre :", self.lbl_cliente_nombre)
        form.addRow("Dirección :", self.lbl_cliente_dir)
        form.addRow("IVA :", self.lbl_cliente_civa)
        form.addRow("CUIT :", self.lbl_cliente_cuit)
        return grupo

    def _armar_cheque(self) -> QGroupBox:
        grupo = QGroupBox("Cheque")
        form = QFormLayout(grupo)
        self.lbl_importe = QLabel()
        self.lbl_orden = QLabel()
        self.lbl_recibo = QLabel()
        self.lbl_banco = QLabel()
        self.lbl_ingreso = QLabel()
        self.lbl_emision = QLabel()
        self.lbl_vencimiento = QLabel()
        self.lbl_concepto = QLabel()
        self.lbl_datosad = QLabel()
        form.addRow("Importe :", self.lbl_importe)
        form.addRow("", self.lbl_orden)
        form.addRow("Recibo :", self.lbl_recibo)
        form.addRow("Banco :", self.lbl_banco)
        form.addRow("Ingreso :", self.lbl_ingreso)
        form.addRow("Emisión :", self.lbl_emision)
        form.addRow("Vencimiento :", self.lbl_vencimiento)
        form.addRow("Concepto :", self.lbl_concepto)
        form.addRow("Datos Ad. :", self.lbl_datosad)
        return grupo

    def _armar_egreso(self) -> QGroupBox:
        self.grupo_egreso = QGroupBox("Egreso")
        form = QFormLayout(self.grupo_egreso)

        self.combo_destino = QComboBox()
        self.combo_destino.addItems(self.cheque_service.DESTINOS_EGRESO)
        form.addRow("Destino :", self.combo_destino)

        self.txt_cpbte = EnteroLineEdit()
        form.addRow("Nº Cpbte. Respaldo :", self.txt_cpbte)

        self.txt_descripcion = UpperCaseLineEdit()
        form.addRow("Descripción :", self.txt_descripcion)

        fila_fecha = QHBoxLayout()
        self.fecha_egreso = QDateEdit()
        self.fecha_egreso.setCalendarPopup(True)
        self.fecha_egreso.setDate(QDate.currentDate())
        fila_fecha.addWidget(self.fecha_egreso)
        self.btn_hoy_egreso = crear_boton_hoy(self.fecha_egreso)
        fila_fecha.addWidget(self.btn_hoy_egreso)
        form.addRow("Fecha :", fila_fecha)

        self.btn_grabar_egreso = QPushButton("Grabar Egreso")
        self.btn_grabar_egreso.clicked.connect(self._on_grabar_egreso)
        form.addRow("", self.btn_grabar_egreso)

        return self.grupo_egreso

    # ------------------------------------------------------------------
    def _cargar_datos(self) -> None:
        c = self.cheque
        self.lbl_importe.setText(f"$ {format_decimal(c.IMPORTE or Decimal('0'))}")
        self.lbl_orden.setText("N/Orden" if c.ORDEN else "")
        self.lbl_recibo.setText(str(c.CPBING or ""))
        self.lbl_ingreso.setText(c.FECING.strftime("%d/%m/%Y") if c.FECING else "")
        self.lbl_emision.setText(c.FECEMI.strftime("%d/%m/%Y") if c.FECEMI else "")
        self.lbl_vencimiento.setText(c.FECVTO.strftime("%d/%m/%Y") if c.FECVTO else "")
        self.lbl_concepto.setText(c.CONCEP or "")
        self.lbl_datosad.setText(c.DATOSAD or "")

        banco = self.repos.banco().by_cod(int(c.BCOSUC)) if c.BCOSUC else None
        self.lbl_banco.setText(f"{c.BCOSUC} - {(banco.NOMBRE or '').strip()}" if banco else "*** No Existe en Tabla ***")

        cliente = self.repos.cliente().by_codigo(c.CLTE) if c.CLTE else None
        if cliente is not None:
            self.lbl_cliente_codigo.setText(str(cliente.CODIGO))
            self.lbl_cliente_nombre.setText((cliente.NOMB or "").strip())
            self.lbl_cliente_dir.setText(f"{(cliente.DIR or '').strip()} ({cliente.CP or ''}) {(cliente.LOC or '').strip()}")
            self.lbl_cliente_civa.setText(str(cliente.CIVA or ""))
            self.lbl_cliente_cuit.setText(cliente.CUIT or "")

        repo_cheque = self.repos.cheque()
        if c.ESTADO == repo_cheque.ESTADO_EN_CARTERA:
            self.grupo_egreso.setTitle("Egreso — En Cartera, se puede registrar")
            self.combo_destino.setEnabled(True)
            self.txt_cpbte.setEnabled(True)
            self.txt_descripcion.setEnabled(True)
            self.fecha_egreso.setEnabled(True)
            self.btn_hoy_egreso.setEnabled(True)
            self.btn_grabar_egreso.setEnabled(True)
        else:
            self.grupo_egreso.setTitle("Egreso — ya registrado" if c.TIPEGR is not None else "Egreso — no corresponde")
            if c.TIPEGR is not None and 0 <= c.TIPEGR < len(self.cheque_service.DESTINOS_EGRESO):
                self.combo_destino.setCurrentIndex(c.TIPEGR)
            self.txt_cpbte.setText(str(c.CPBEGR) if c.CPBEGR is not None else "")
            self.txt_descripcion.setText(c.DESTINO or "")
            if c.FECEGR is not None:
                self.fecha_egreso.setDate(QDate(c.FECEGR.year, c.FECEGR.month, c.FECEGR.day))
            self.combo_destino.setEnabled(False)
            self.txt_cpbte.setEnabled(False)
            self.txt_descripcion.setEnabled(False)
            self.fecha_egreso.setEnabled(False)
            self.btn_hoy_egreso.setEnabled(False)
            self.btn_grabar_egreso.setEnabled(False)

    # ------------------------------------------------------------------
    def _on_grabar_egreso(self) -> None:
        cpbte = self.txt_cpbte.entero()
        if cpbte is None:
            QMessageBox.warning(self, "Egreso de Cheque", "Ingresá el Nº de Comprobante de respaldo.")
            return
        descripcion = self.txt_descripcion.text().strip()
        if not descripcion:
            QMessageBox.warning(self, "Egreso de Cheque", "Ingresá una Descripción.")
            return

        respuesta = QMessageBox.question(
            self, "Egreso de Cheque",
            "El siguiente proceso registrará el Egreso del Cheque.\n¿Desea continuar?",
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        try:
            self.cheque = self.cheque_service.registrar_egreso(
                self.cheque.NROCHEQ,
                destino_index=self.combo_destino.currentIndex(),
                cpbte_respaldo=cpbte,
                descripcion=descripcion,
                fecha_egreso=self.fecha_egreso.date().toPyDate(),
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Egreso de Cheque", str(exc))
            return

        self._cargar_datos()

    def _on_borrar(self) -> None:
        respuesta = QMessageBox.question(
            self, "Borrado de Cheque",
            f"¿Realmente desea borrar el Cheque Nº {self.cheque.NROCHEQ}?",
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return
        try:
            self.cheque_service.eliminar(self.cheque.NROCHEQ)
        except ValueError as exc:
            QMessageBox.warning(self, "Borrado de Cheque", str(exc))
            return
        self.borrado = True
        self.accept()
