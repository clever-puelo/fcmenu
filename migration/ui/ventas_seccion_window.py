"""`VentasSeccionWindow` — "Ventas de Artículos por Cliente", migración
de `ESTADIST.frm`: renglones de `Fcestad1` de una Sección en un rango de
fechas, con el Cliente de cada movimiento (no agrupado — a pesar del
título, el legacy muestra una lista plana, no una jerarquía por cliente).

**Ajustes pedidos por el usuario (2026-08-16) sobre el legacy**:
- Rango de fechas real (Desde/Hasta, ambos elegibles) en vez de un solo
  día + fin de ESE mes automático (`ESTADIST.frm` sólo dejaba elegir el
  "desde" — "ahora muestra sólo el mes seleccionado, es incorrecto").
- Código de Cliente en columna aparte (ajustado a la derecha, ordenable
  numéricamente) y Razón Social sola (ordenable alfabéticamente) — el
  legacy los concatenaba en una sola celda de texto.
- Sólo se muestran las columnas de Unidad de Medida (Pulg/MM/Telas/Mtr/
  Cant.) que la Sección elegida usa de verdad, según su configuración
  real en `Fctabla1` (`factura_renglon.posiciones_activas_seccion`) — el
  resto se oculta.
"""

from __future__ import annotations

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.repository import RepositoryFactory
from migration.services import EstadisticaVentasService

from .factura_renglon import posiciones_activas_seccion
from .widgets import TablaBusqueda, crear_boton_hoy

COLUMNAS = ["Fecha", "Comprobante", "Cód.", "Cliente", "Pulg", "Mtr", "MM", "Telas", "Cant.", "P.Venta", "Importe"]
COL_CODIGO, COL_CLIENTE, COL_PULG, COL_MTR, COL_MM, COL_TELAS, COL_CANT, COL_PVTA, COL_IMPORTE = range(2, 11)

# Posición del enum real (ver factura_renglon.POSICIONES_UM) que le
# corresponde a cada columna de Unidad de Medida de esta grilla.
POSICION_POR_COLUMNA = {COL_PULG: 1, COL_MTR: 2, COL_MM: 3, COL_TELAS: 4, COL_CANT: 6}


class VentasSeccionWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Ventas de Artículos por Cliente")
        self.resize(1050, 600)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.estadistica_service = EstadisticaVentasService(self.db)

        self._construir_ui()
        self._poblar_combo_seccion()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        fila_filtro = QHBoxLayout()
        fila_filtro.addWidget(QLabel("Sección :"))
        self.combo_seccion = QComboBox()
        fila_filtro.addWidget(self.combo_seccion)

        hoy = QDate.currentDate()
        fila_filtro.addWidget(QLabel("Desde :"))
        self.fecha_desde = QDateEdit()
        self.fecha_desde.setCalendarPopup(True)
        self.fecha_desde.setDate(QDate(hoy.year(), hoy.month(), 1))
        fila_filtro.addWidget(self.fecha_desde)
        fila_filtro.addWidget(crear_boton_hoy(self.fecha_desde))

        fila_filtro.addWidget(QLabel("Hasta :"))
        self.fecha_hasta = QDateEdit()
        self.fecha_hasta.setCalendarPopup(True)
        self.fecha_hasta.setDate(hoy)
        fila_filtro.addWidget(self.fecha_hasta)
        fila_filtro.addWidget(crear_boton_hoy(self.fecha_hasta))

        fila_filtro.addStretch()
        btn_cerrar = QPushButton("Salir")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.tabla = TablaBusqueda(
            COLUMNAS, columnas_derecha=(COL_CODIGO, COL_PULG, COL_MTR, COL_MM, COL_TELAS, COL_CANT, COL_PVTA, COL_IMPORTE)
        )
        layout.addWidget(self.tabla, stretch=1)

        fila_totales = QHBoxLayout()
        fila_totales.addWidget(QLabel("Tot. Movim. :"))
        self.lbl_cantidad = QLabel("0")
        fila_totales.addWidget(self.lbl_cantidad)
        fila_totales.addStretch()
        fila_totales.addWidget(QLabel("Total Importe :"))
        self.lbl_total = QLabel("$ 0,00")
        self.lbl_total.setStyleSheet("font-weight: bold;")
        fila_totales.addWidget(self.lbl_total)
        layout.addLayout(fila_totales)

        self.fecha_desde.dateChanged.connect(self._on_filtro_cambiado)
        self.fecha_hasta.dateChanged.connect(self._on_filtro_cambiado)
        self.combo_seccion.currentIndexChanged.connect(self._on_filtro_cambiado)

    def _poblar_combo_seccion(self) -> None:
        self.combo_seccion.blockSignals(True)
        for seccion in self.repos.fctablas().by_ctab("SC"):
            cod = (seccion.COD or "").strip()
            self.combo_seccion.addItem(f"{cod} - {(seccion.DESCRI or '').strip()}", cod)
        self.combo_seccion.blockSignals(False)
        if self.combo_seccion.count():
            self._on_filtro_cambiado()

    # ------------------------------------------------------------------
    def _on_filtro_cambiado(self) -> None:
        cod_seccion = self.combo_seccion.currentData()
        if not cod_seccion:
            return
        fecha_desde = self.fecha_desde.date().toPyDate()
        fecha_hasta = self.fecha_hasta.date().toPyDate()
        if fecha_hasta < fecha_desde:
            fecha_hasta = fecha_desde

        self._actualizar_columnas_visibles(cod_seccion)

        resultado = self.estadistica_service.ventas_seccion_por_rango(cod_seccion, fecha_desde, fecha_hasta)

        filas = [
            (
                [
                    m.fecha.strftime("%d/%m/%Y") if m.fecha else "",
                    f"{m.tipo_label} {m.letra} {m.comprobante}",
                    str(m.cliente_codigo or ""),
                    m.cliente_nombre,
                    format_decimal(m.pulg) if m.pulg else "",
                    format_decimal(m.mtr) if m.mtr else "",
                    str(m.milim) if m.milim else "",
                    str(m.telas) if m.telas else "",
                    format_decimal(m.cantidad) if m.cantidad else "",
                    format_decimal(m.precio_venta),
                    format_decimal(m.importe),
                ],
                m,
            )
            for m in resultado.movimientos
        ]
        self.tabla.cargar_filas(filas)
        self.lbl_cantidad.setText(str(resultado.cantidad_movimientos))
        self.lbl_total.setText(f"$ {format_decimal(resultado.total_importe)}")

    def _actualizar_columnas_visibles(self, cod_seccion: str) -> None:
        posiciones = posiciones_activas_seccion(self.repos.fctablas(), cod_seccion)
        for columna, posicion in POSICION_POR_COLUMNA.items():
            self.tabla.setColumnHidden(columna, posicion not in posiciones)

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
