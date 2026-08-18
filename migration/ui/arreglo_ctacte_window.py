"""`ArregloCtaCteWindow` — migración de `CargaCC.frm` ("Arreglos >
Cuenta Corriente", menú oculto `Complementos`/`RECICLA1`/`ARRECCTE1` de
`FCMENU.frm`).

Editor directo de UN renglón de `CtasCtes` — ver
`ArregloCtaCteService` para el alcance ("romper vidrio", sin las
validaciones de negocio del Facturador/Recibo) y la corrección
confirmada con el usuario (2026-08-16): `Clientes.DEUDA` se recalcula
solo tras cada Grabar/Borrar, el legacy nunca la tocaba.
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import ArregloCtaCteService

from .widgets import EnterAsTabFilter, EnteroLineEdit, MontoLineEdit, UpperCaseLineEdit, crear_boton_hoy

# Tipo Cpbte 0-6, réplica de CargaCC.frx (Combo1) — mismas etiquetas ya
# usadas en toda la app para Ctascte.TIPO.
TIPOS_CPBTE = [(t, ETIQUETAS_TIPO_CTASCTE[t]) for t in range(7)]

# Letra — réplica de CargaCC.frx (Combo2): "X" para movimientos sin
# letra fiscal real (Recibo/Pago a Cta./Descuento).
LETRAS = ["A", "B", "X"]


class ArregloCtaCteWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Arreglos — Cuenta Corriente")
        self.resize(520, 480)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.service = ArregloCtaCteService(self.db)
        self._fila_existente = False

        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        form = QFormLayout()

        self.spin_clte = QSpinBox()
        self.spin_clte.setRange(0, 999999)
        self.spin_clte.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.lbl_nombre = QLabel()
        fila_clte = QHBoxLayout()
        fila_clte.addWidget(self.spin_clte)
        fila_clte.addWidget(self.lbl_nombre, stretch=1)
        form.addRow("Cliente :", fila_clte)

        self.fecha = QDateEdit()
        self.fecha.setCalendarPopup(True)
        self.fecha.setDate(QDate.currentDate())
        fila_fecha = QHBoxLayout()
        fila_fecha.addWidget(self.fecha)
        fila_fecha.addWidget(crear_boton_hoy(self.fecha))
        form.addRow("Fecha :", fila_fecha)

        self.combo_tipo = QComboBox()
        for codigo, etiqueta in TIPOS_CPBTE:
            self.combo_tipo.addItem(f"{codigo}-{etiqueta}", codigo)
        form.addRow("Tipo Cpbte. :", self.combo_tipo)

        self.combo_letra = QComboBox()
        self.combo_letra.addItems(LETRAS)
        form.addRow("Letra :", self.combo_letra)

        self.txt_prefijo = EnteroLineEdit("0001")
        form.addRow("Prefijo (Pto.Vta.) :", self.txt_prefijo)

        self.txt_cpbte = EnteroLineEdit()
        form.addRow("Nro. Cpbte. :", self.txt_cpbte)

        self.lbl_estado = QLabel()
        self.lbl_estado.setStyleSheet("color: #b00020; font-weight: bold;")
        form.addRow("", self.lbl_estado)

        self.txt_imput1 = UpperCaseLineEdit()
        self.txt_imput1.setMaxLength(20)
        form.addRow("1ª Imputación :", self.txt_imput1)

        self.txt_imput2 = UpperCaseLineEdit()
        self.txt_imput2.setMaxLength(20)
        form.addRow("2ª Imputación :", self.txt_imput2)

        self.txt_imput3 = UpperCaseLineEdit()
        self.txt_imput3.setMaxLength(20)
        form.addRow("3ª Imputación :", self.txt_imput3)

        self.txt_debe = MontoLineEdit()
        form.addRow("Resto (Debe) :", self.txt_debe)

        self.txt_impte = MontoLineEdit()
        form.addRow("Importe :", self.txt_impte)

        self.fecvto = QDateEdit()
        self.fecvto.setCalendarPopup(True)
        self.fecvto.setDate(QDate.currentDate())
        fila_fecvto = QHBoxLayout()
        fila_fecvto.addWidget(self.fecvto)
        fila_fecvto.addWidget(crear_boton_hoy(self.fecvto))
        form.addRow("Fec. Vto. :", fila_fecvto)

        layout.addLayout(form)

        fila_botones = QHBoxLayout()
        self.btn_borrar = QPushButton("Borrar")
        self.btn_borrar.setVisible(False)
        self.btn_borrar.clicked.connect(self._on_borrar)
        fila_botones.addWidget(self.btn_borrar)
        fila_botones.addStretch()
        self.btn_grabar = QPushButton("Grabar")
        self.btn_grabar.setEnabled(False)
        self.btn_grabar.clicked.connect(self._on_grabar)
        fila_botones.addWidget(self.btn_grabar)
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

        self.spin_clte.editingFinished.connect(self._on_cliente_confirmado)
        self.txt_cpbte.editingFinished.connect(self._buscar_fila)
        self.combo_tipo.currentIndexChanged.connect(self._buscar_fila)
        self.combo_letra.currentIndexChanged.connect(self._buscar_fila)
        self.txt_prefijo.editingFinished.connect(self._buscar_fila)

        self._enter_as_tab = EnterAsTabFilter(self, boton_final=self.btn_grabar)

    # ------------------------------------------------------------------
    def _on_cliente_confirmado(self) -> None:
        clte = self.spin_clte.value()
        cliente = self.repos.cliente().by_codigo(clte) if clte else None
        self.lbl_nombre.setText((cliente.NOMB or "").strip() if cliente else "*** Cliente NO Existe ***")

    def _buscar_fila(self) -> None:
        clte = self.spin_clte.value()
        cpbte = self.txt_cpbte.entero()
        prefijo = self.txt_prefijo.entero()
        if not clte or cpbte is None or prefijo is None:
            return

        tipo = self.combo_tipo.currentData()
        letra = self.combo_letra.currentText()
        fila = self.service.buscar(clte, letra, tipo, prefijo, cpbte)

        if fila is not None:
            self._fila_existente = True
            self.lbl_estado.setText("")
            self.txt_imput1.setText((fila.IMPUT1 or "").strip())
            self.txt_imput2.setText((fila.IMPUT2 or "").strip())
            self.txt_imput3.setText((fila.IMPUT3 or "").strip())
            self.txt_debe.set_decimal(fila.DEBE or Decimal("0"))
            self.txt_impte.set_decimal(fila.IMPTE or Decimal("0"))
            if fila.FECVTO:
                self.fecvto.setDate(QDate(fila.FECVTO.year, fila.FECVTO.month, fila.FECVTO.day))
            self.btn_borrar.setVisible(True)
        else:
            self._fila_existente = False
            self.lbl_estado.setText("N U E V O")
            self.txt_imput1.setText("0")
            self.txt_imput2.setText("0")
            self.txt_imput3.setText("0")
            self.txt_debe.set_decimal(Decimal("0"))
            self.txt_impte.set_decimal(Decimal("0"))
            self.fecvto.setDate(QDate.currentDate())
            self.btn_borrar.setVisible(False)

        self.btn_grabar.setEnabled(True)

    # ------------------------------------------------------------------
    def _on_grabar(self) -> None:
        respuesta = QMessageBox.question(
            self, "Carga de Datos a la Tabla", "¿Desea continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.service.grabar(
            clte=self.spin_clte.value(),
            fecha=self.fecha.date().toPyDate(),
            tipo=self.combo_tipo.currentData(),
            letra=self.combo_letra.currentText(),
            prefijo=self.txt_prefijo.entero() or 0,
            cpbte=self.txt_cpbte.entero() or 0,
            imput1=self.txt_imput1.text().strip(),
            imput2=self.txt_imput2.text().strip(),
            imput3=self.txt_imput3.text().strip(),
            debe=self.txt_debe.decimal(),
            impte=self.txt_impte.decimal(),
            fecvto=self.fecvto.date().toPyDate(),
        )
        QMessageBox.information(self, "Arreglos", "Registro grabado.")
        self._buscar_fila()

    def _on_borrar(self) -> None:
        respuesta = QMessageBox.question(
            self, "Eliminar Datos de la Tabla", "¿Desea continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.service.borrar(
            self.spin_clte.value(), self.combo_letra.currentText(),
            self.combo_tipo.currentData(), self.txt_prefijo.entero() or 0, self.txt_cpbte.entero() or 0,
        )
        QMessageBox.information(self, "Arreglos", "Registro eliminado.")
        self._buscar_fila()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
