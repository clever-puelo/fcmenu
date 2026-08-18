"""`StockConsultaWindow` — Consulta de Stock, migración de `Verstock.frm`.

Búsqueda por Sección (`Combo2` del legacy) o por Descripción — el legacy
sólo buscaba por Sección + hasta 2 segmentos de código armados a mano
(`Text2`/`Text3`, mismo mecanismo ya visto en `Stock.frm`); acá se
reusa el patrón ya establecido de Cliente/Artículo (código exacto O
descripción parcial), más simple y sin el hueco de los códigos
compuestos con un solo segmento tipeado (ver análisis de la fase).

Columnas agregadas sobre lo que ya lista `ArticuloRepository` (Precio
Venta/Costo): Stock Actual (`Stock.STUNID`, fuente de verdad) y Stock
Mínimo — doble clic en una fila abre el historial de movimientos
(`MovStockHistorialDialog`, migración de `Sub CargaMOVS`).
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.models import Articulo
from migration.repository import RepositoryFactory

from .movstock_historial_dialog import MovStockHistorialDialog
from .widgets import TablaBusqueda, UpperCaseLineEdit, redimensionar_pct_pantalla

COLUMNAS = ["Sección", "Código", "Descripción", "Precio Venta", "Precio Costo", "Stock Actual", "Stock Mínimo"]
COL_PVTA, COL_PCOS, COL_STOCK, COL_STMIN = 3, 4, 5, 6


class StockConsultaWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Consulta de Stock")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        redimensionar_pct_pantalla(self, 95, 90)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)

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
        self.combo_seccion.currentIndexChanged.connect(self._on_buscar)
        fila_filtro.addWidget(self.combo_seccion)

        fila_filtro.addWidget(QLabel("Descripción contiene :"))
        self.txt_descripcion = UpperCaseLineEdit()
        self.txt_descripcion.editingFinished.connect(self._on_buscar)
        fila_filtro.addWidget(self.txt_descripcion, stretch=1)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.clicked.connect(self._on_buscar)
        fila_filtro.addWidget(btn_buscar)

        # No tenía botón Cerrar (feedback del usuario, 2026-08-18).
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(COL_PVTA, COL_PCOS, COL_STOCK, COL_STMIN))
        self.tabla.itemActivated.connect(self._on_ver_historial)
        self.tabla.itemDoubleClicked.connect(self._on_ver_historial)
        layout.addWidget(self.tabla, stretch=1)

        self.lbl_cantidad = QLabel("")
        layout.addWidget(self.lbl_cantidad)

    def _poblar_combo_seccion(self) -> None:
        self.combo_seccion.blockSignals(True)
        self.combo_seccion.addItem("-- Todas --", None)
        for seccion in self.repos.fctablas().by_ctab("SC"):
            cod = (seccion.COD or "").strip()
            self.combo_seccion.addItem(f"{cod} - {(seccion.DESCRI or '').strip()}", cod)
        self.combo_seccion.blockSignals(False)
        self._on_buscar()

    # ------------------------------------------------------------------
    def _on_buscar(self) -> None:
        cod_seccion = self.combo_seccion.currentData()
        descripcion = self.txt_descripcion.text().strip()

        articulo_repo = self.repos.articulo()
        if cod_seccion:
            articulos = articulo_repo.by_seccion(cod_seccion)
            if descripcion:
                articulos = [a for a in articulos if descripcion in (a.DESCRI or "").upper()]
        elif descripcion:
            articulos = articulo_repo.by_descri(descripcion)
        else:
            articulos = articulo_repo.todos()

        self._cargar_tabla(articulos)

    def _cargar_tabla(self, articulos: list[Articulo]) -> None:
        stock_repo = self.repos.stock()
        filas = []
        for art in articulos:
            cod1 = (art.COD1 or "").strip()
            cod2 = (art.COD2 or "").strip()
            stock = stock_repo.by_cod1_cod2(cod1, cod2)
            stock_actual = stock.STUNID if stock and stock.STUNID is not None else Decimal("0")
            stock_min = stock.STMIN if stock and stock.STMIN is not None else Decimal("0")
            filas.append(
                (
                    [
                        cod1, cod2, (art.DESCRI or "").strip(),
                        format_decimal(art.PREC or Decimal("0")),
                        format_decimal(art.PCOS or Decimal("0")),
                        format_decimal(stock_actual),
                        format_decimal(stock_min),
                    ],
                    art,
                )
            )
        self.tabla.cargar_filas(filas)
        self.lbl_cantidad.setText(f"Renglones Encontrados : {len(articulos)}")

    def _on_ver_historial(self, item) -> None:
        articulo = self.tabla.objeto_en_fila(item.row())
        if articulo is None:
            return
        MovStockHistorialDialog(self.repos, articulo, parent=self).exec()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
