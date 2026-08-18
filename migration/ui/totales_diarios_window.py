"""`TotalesDiariosWindow` — migración de `TotFact.frm` ("Totales
Diarios"): grilla de un renglón por día (`Totales`) desde una fecha
elegida hasta fin de ese mes, recortada a los primeros N — mismo
patrón "Muestra desde ... Los Primeros ..." ya migrado en
`FacturasEmitidasWindow` (`VerFact.frm`; acá el legacy decía "Últimos",
etiqueta corregida — ver `TotalesDiariosService`).

Clic en un día abre el detalle (`TotalDiarioDetalleDialog`); botón
"Total Mensual" abre el resumen del mes con la corrección de rango
confirmada con el usuario."""

from __future__ import annotations

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.repository import RepositoryFactory
from migration.services import TotalesDiariosService

from .total_diario_detalle_dialog import TotalDiarioDetalleDialog
from .total_mensual_dialog import TotalMensualDialog
from .widgets import TablaBusqueda, crear_boton_hoy

COLUMNAS = ["Fecha", "Precio de Venta", "Precio de Costo", "Venta Real", "Facturas", "N. Crédito"]
COL_PVENTA, COL_PCOSTO, COL_VREAL, COL_FACTURAS, COL_NC = 1, 2, 3, 4, 5

OPCIONES_LIMITE = [5, 10, 25, 50, 75, 100, 150]


class TotalesDiariosWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Totales Diarios")
        # +20% ancho / +50% alto (feedback del usuario, 2026-08-18).
        self.resize(1140, 900)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.service = TotalesDiariosService(self.db)

        self._construir_ui()
        self._refrescar()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        fila_filtro = QHBoxLayout()
        fila_filtro.addWidget(QLabel("Muestra desde :"))
        self.fecha_desde = QDateEdit()
        self.fecha_desde.setCalendarPopup(True)
        self.fecha_desde.setDate(QDate.currentDate())
        fila_filtro.addWidget(self.fecha_desde)
        fila_filtro.addWidget(crear_boton_hoy(self.fecha_desde))

        fila_filtro.addWidget(QLabel("Los Primeros :"))
        self.combo_limite = QComboBox()
        for opcion in OPCIONES_LIMITE:
            self.combo_limite.addItem(str(opcion), opcion)
        self.combo_limite.setCurrentIndex(OPCIONES_LIMITE.index(50))
        fila_filtro.addWidget(self.combo_limite)

        fila_filtro.addStretch()
        btn_total_mensual = QPushButton("Total Mensual")
        btn_total_mensual.clicked.connect(self._on_total_mensual)
        fila_filtro.addWidget(btn_total_mensual)
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(COL_PVENTA, COL_PCOSTO, COL_VREAL, COL_FACTURAS, COL_NC))
        layout.addWidget(self.tabla, stretch=1)

        fila_totales = QHBoxLayout()
        fila_totales.addWidget(QLabel("Precio Venta :"))
        self.lbl_pvta = QLabel("$ 0,00")
        fila_totales.addWidget(self.lbl_pvta)
        fila_totales.addWidget(QLabel("Precio Costo :"))
        self.lbl_pcos = QLabel("$ 0,00")
        fila_totales.addWidget(self.lbl_pcos)
        fila_totales.addWidget(QLabel("Venta Real :"))
        self.lbl_pesp = QLabel("$ 0,00")
        self.lbl_pesp.setStyleSheet("font-weight: bold;")
        fila_totales.addWidget(self.lbl_pesp)
        fila_totales.addStretch()
        layout.addLayout(fila_totales)

        self.fecha_desde.dateChanged.connect(self._refrescar)
        self.combo_limite.currentIndexChanged.connect(self._refrescar)
        self.tabla.itemClicked.connect(self._on_click_fila)

    # ------------------------------------------------------------------
    def _refrescar(self) -> None:
        fecha_desde = self.fecha_desde.date().toPyDate()
        limite = self.combo_limite.currentData()
        resultado = self.service.listar(fecha_desde, limite)

        filas = [
            (
                [
                    f.fecha.strftime("%d/%m/%Y") if f.fecha else "",
                    format_decimal(f.precio_venta),
                    format_decimal(f.precio_costo),
                    format_decimal(f.venta_real),
                    str(f.cantidad_facturas),
                    str(f.cantidad_notas_credito),
                ],
                f.totales,
            )
            for f in resultado.filas
        ]
        self.tabla.cargar_filas(filas)
        self.lbl_pvta.setText(f"$ {format_decimal(resultado.total_precio_venta)}")
        self.lbl_pcos.setText(f"$ {format_decimal(resultado.total_precio_costo)}")
        self.lbl_pesp.setText(f"$ {format_decimal(resultado.total_venta_real)}")

    def _on_click_fila(self, item: QTableWidgetItem) -> None:
        totales = self.tabla.objeto_en_fila(item.row())
        if totales is None or totales.FECHA is None:
            return
        detalle = self.service.detalle_dia(totales.FECHA)
        if detalle is None:
            return
        TotalDiarioDetalleDialog(detalle, parent=self).exec()

    def _on_total_mensual(self) -> None:
        fecha_desde = self.fecha_desde.date().toPyDate()
        resumen = self.service.resumen_mensual(fecha_desde)
        if resumen is None:
            QMessageBox.information(self, "Total Mensual", "No hay datos para ese mes.")
            return
        TotalMensualDialog(resumen, parent=self).exec()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
