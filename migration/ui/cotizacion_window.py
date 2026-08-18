"""Cotización del Dólar — migración de `Cotizac.frm`.

Alcance confirmado con el usuario (2026-08-15, ver memoria de sesión
`pyqt6_ui_progress.md`):

- El legacy usaba una fila "ancla" con `FECHA='01/01/1900'` como
  plantilla de "última cotización" para precargar una fecha nueva, con
  doble escritura en cada Grabar. Su propia consulta de búsqueda por
  fecha comparaba contra el formato `yyyy/mm/dd`, que **no coincide**
  con el formato real de los datos (`dd/mm/yyyy`, confirmado en la fase
  de datos — ver `CotizacionRepository`). Se reemplaza por
  `CotizacionRepository.ultima()` (ya validado contra datos reales) para
  precargar, sin fila ancla ni doble escritura — ver
  `CotizacionRepository.guardar()`.
- `DOLARE` (6to campo del legacy) siempre se graba en 0 y no tiene
  control de UI — código muerto real, no se migra.
- No se replica el efecto del legacy donde clickear "Grabar" y luego
  "No" en el diálogo de confirmación igual cerraba la ventana sin
  grabar — acá "No"/"Cancelar" simplemente no hacen nada y la ventana
  queda abierta.
"""

from __future__ import annotations

from datetime import date

from PyQt6.QtWidgets import (
    QDateEdit,
    QFormLayout,
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
from migration.repository import RepositoryFactory

from .decimals import parse_decimal
from .widgets import EnterAsTabFilter


class CotizacionWindow(QMainWindow):
    """Carga de Cotización del Dólar (equivalente a Cotizac de Cotizac.frm)."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Cotización del Dólar")
        self.resize(420, 320)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)

        self._construir_ui()
        self._cargar_fecha()

        self._enter_as_tab = EnterAsTabFilter(self)

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        fila_fecha = QHBoxLayout()
        fila_fecha.addWidget(QLabel("Fecha :"))
        self.fecha = QDateEdit()
        self.fecha.setCalendarPopup(True)
        self.fecha.setDate(date.today())
        self.fecha.dateChanged.connect(self._cargar_fecha)
        fila_fecha.addWidget(self.fecha)
        fila_fecha.addStretch()
        self.lbl_nuevo = QLabel("N U E V O")
        self.lbl_nuevo.setStyleSheet("color: red; font-weight: bold;")
        self.lbl_nuevo.setVisible(False)
        fila_fecha.addWidget(self.lbl_nuevo)
        layout.addLayout(fila_fecha)

        grupo = QGroupBox("Cotización")
        form = QFormLayout(grupo)
        self.txt_dolar = QLineEdit("0")
        form.addRow("Dólar :", self.txt_dolar)
        self.txt_dolar_a = QLineEdit("0")
        form.addRow("Dólar 'A' :", self.txt_dolar_a)
        self.txt_dolar_b = QLineEdit("0")
        form.addRow("Dólar 'B' :", self.txt_dolar_b)
        self.txt_dolar_c = QLineEdit("0")
        form.addRow("Dólar 'C' :", self.txt_dolar_c)
        self.txt_dolar_d = QLineEdit("0")
        form.addRow("Dólar 'D' :", self.txt_dolar_d)
        layout.addWidget(grupo)

        layout.addStretch()
        layout.addLayout(self._armar_barra_botones())

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()
        fila.addStretch()

        self.btn_guardar = QPushButton("Grabar")
        self.btn_guardar.clicked.connect(self._on_guardar)
        fila.addWidget(self.btn_guardar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.close)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    # Réplica de Sub CargaFecha() (Cotizac.frm:399-438), sin la fila ancla
    # ------------------------------------------------------------------
    def _cargar_fecha(self) -> None:
        fecha = self.fecha.date().toPyDate()
        cotizacion = self.repos.cotizacion().by_fecha(fecha)

        if cotizacion is not None:
            self.lbl_nuevo.setVisible(False)
            self.txt_dolar.setText(str(cotizacion.DOLAR or 0))
            self.txt_dolar_a.setText(str(cotizacion.DOLARA or 0))
            self.txt_dolar_b.setText(str(cotizacion.DOLARB or 0))
            self.txt_dolar_c.setText(str(cotizacion.DOLARC or 0))
            self.txt_dolar_d.setText(str(cotizacion.DOLARD or 0))
            return

        self.lbl_nuevo.setVisible(True)
        anterior = self.repos.cotizacion().ultima()
        if anterior is None:
            self.txt_dolar.setText("0")
            self.txt_dolar_a.setText("0")
            self.txt_dolar_b.setText("0")
            self.txt_dolar_c.setText("0")
            self.txt_dolar_d.setText("0")
            return

        self.txt_dolar.setText(str(anterior.DOLAR or 0))
        self.txt_dolar_a.setText(str(anterior.DOLARA or 0))
        self.txt_dolar_b.setText(str(anterior.DOLARB or 0))
        self.txt_dolar_c.setText(str(anterior.DOLARC or 0))
        self.txt_dolar_d.setText(str(anterior.DOLARD or 0))

    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        respuesta = QMessageBox.question(
            self,
            "Cotización del Dólar",
            "Desea Grabar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        fecha = self.fecha.date().toPyDate()
        datos = dict(
            DOLAR=parse_decimal(self.txt_dolar.text()),
            DOLARA=parse_decimal(self.txt_dolar_a.text()),
            DOLARB=parse_decimal(self.txt_dolar_b.text()),
            DOLARC=parse_decimal(self.txt_dolar_c.text()),
            DOLARD=parse_decimal(self.txt_dolar_d.text()),
            DOLARE=0,
        )
        self.repos.cotizacion().guardar(fecha, datos)

        QMessageBox.information(self, "Cotización del Dólar", "Cotización grabada correctamente.")
        self.lbl_nuevo.setVisible(False)

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
