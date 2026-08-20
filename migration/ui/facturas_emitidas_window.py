"""`FacturasEmitidasWindow` — migración de `VerFact.frm` ("Facturas
Emitidas"): consulta de comprobantes del subdiario de ventas (`FCIVAVTA`)
desde una fecha elegida hasta fin de ese mes, recortada a los primeros N
(convención ya establecida: sólo la consulta, de sólo lectura — ver
`migration.services.FacturasEmitidasService` para la decisión confirmada
de NO migrar "Anular").

Clic en una fila abre el detalle (Cliente + Comprobante + Renglones
reales) — mismo patrón ya usado en `CtaCteWindow`/`ChequesConsultaWindow`
(`itemClicked`, un solo clic, sin diálogo modal de selección de por
medio)."""

from __future__ import annotations

from datetime import date

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.repository import RepositoryFactory
from migration.services import FacturasEmitidasService

from .factura_emitida_detalle_dialog import FacturaEmitidaDetalleDialog
from .widgets import TablaBusqueda, crear_boton_hoy, crear_recuadro_destacado, redimensionar_pct_pantalla

# Cód. de Cliente en columna separada de la Razón Social (feedback del
# usuario, 2026-08-18, segunda ronda: "para poder ordenar
# alfabéticamente") — antes venían concatenados en una sola celda de
# texto ("123 FULANO SA"), lo que impedía ordenar por Cliente de verdad
# (ordenaba por el código como prefijo del texto). `TablaBusqueda` ya
# ordena cada columna al clic sobre su título (ver `widgets.py`).
COLUMNAS = ["Fecha", "T. Cpbte.", "Número", "Imp. Bruto", "Imp. IVA", "Imp. Neto", "Cód.", "Cliente"]
COL_BRUTO, COL_IVA, COL_NETO, COL_COD_CLIENTE = 3, 4, 5, 6

OPCIONES_LIMITE = [5, 10, 25, 50, 75, 100, 150]


class FacturasEmitidasWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None, *, fecha_inicial: date | None = None):
        super().__init__(parent)
        self.setWindowTitle("Facturas Emitidas")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        redimensionar_pct_pantalla(self, 97, 90)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.service = FacturasEmitidasService(self.db)

        self._construir_ui()
        # "Ver Cpbtes." de Totales del Día (pedido del usuario,
        # 2026-08-19) abre este módulo ya filtrado desde esa fecha —
        # sin `fecha_inicial`, arranca en hoy como siempre.
        if fecha_inicial is not None:
            self.fecha_desde.setDate(QDate(fecha_inicial.year, fecha_inicial.month, fecha_inicial.day))
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
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(COL_BRUTO, COL_IVA, COL_NETO, COL_COD_CLIENTE))
        layout.addWidget(self.tabla, stretch=1)

        # Totales separados (no pegados a la izquierda) y "Neto" resaltado
        # en un recuadro (feedback del usuario, 2026-08-18, tercera ronda).
        fila_totales = QHBoxLayout()
        fila_totales.addWidget(QLabel("Bruto :"))
        self.lbl_bruto = QLabel("$ 0,00")
        fila_totales.addWidget(self.lbl_bruto)
        fila_totales.addSpacing(30)
        fila_totales.addWidget(QLabel("IVA :"))
        self.lbl_iva = QLabel("$ 0,00")
        fila_totales.addWidget(self.lbl_iva)
        fila_totales.addSpacing(30)
        recuadro_neto, self.lbl_neto = crear_recuadro_destacado("Neto :")
        fila_totales.addWidget(recuadro_neto)
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
                    f.tipo_label,
                    f.comprobante,
                    format_decimal(f.bruto),
                    format_decimal(f.iva),
                    format_decimal(f.neto),
                    str(f.clte or ""),
                    f.cliente_nombre,
                ],
                f.factura,
            )
            for f in resultado.filas
        ]
        self.tabla.cargar_filas(filas)
        self.lbl_bruto.setText(f"$ {format_decimal(resultado.total_bruto)}")
        self.lbl_iva.setText(f"$ {format_decimal(resultado.total_iva)}")
        self.lbl_neto.setText(f"$ {format_decimal(resultado.total_neto)}")

    def _on_click_fila(self, item: QTableWidgetItem) -> None:
        factura = self.tabla.objeto_en_fila(item.row())
        if factura is None:
            return
        FacturaEmitidaDetalleDialog(self.repos, factura, parent=self).exec()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
