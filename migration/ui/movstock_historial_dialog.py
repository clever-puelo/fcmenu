"""Diálogo de historial de movimientos de un artículo — migración de
`Verstock.frm Sub CargaMOVS` (líneas 1056-1135), abierto al hacer doble
clic en un artículo de `StockConsultaWindow`.

La leyenda de `TIPO` es la misma tabla de decodificación real usada como
fuente de verdad para `StockMovimientoService.CODIGOS_ENTRADA/SALIDA`
(ver ese docstring) — acá se reusa tal cual, con las etiquetas completas
en vez de las abreviadas a 6 caracteres del FlexGrid original.
"""

from __future__ import annotations

from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.decimals import format_decimal
from migration.models import Articulo
from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda

# Réplica exacta de Verstock.frm:1098-1111.
ETIQUETAS_TIPO_MOVSTOCK = {
    "11": "N/Créd. A", "12": "N/Créd. B", "13": "Producción/Fabricación",
    "14": "Compra", "15": "Importación", "18": "Ajuste Inventario +",
    "21": "Factura A", "22": "Factura B", "23": "Remito",
    "24": "Promoción", "25": "Ajuste -", "28": "Ajuste Inventario -",
}

COLUMNAS = ["Tipo", "Comprobante", "Fecha", "Cantidad"]


class MovStockHistorialDialog(QDialog):
    def __init__(self, repos: RepositoryFactory, articulo: Articulo, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos

        cod1 = (articulo.COD1 or "").strip()
        cod2 = (articulo.COD2 or "").strip()
        self.setWindowTitle(f"Historial de Movimientos — {cod1} / {cod2}")
        self.resize(560, 460)

        self._construir_ui(cod1, cod2)

    def _construir_ui(self, cod1: str, cod2: str) -> None:
        layout = QVBoxLayout(self)

        movimientos = self.repos.movstock().by_cod1_cod2(cod1, cod2)
        layout.addWidget(QLabel(f"Renglones Encontrados : {len(movimientos)}"))

        tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(1, 3))
        filas = [
            (
                [
                    ETIQUETAS_TIPO_MOVSTOCK.get((m.TIPO or "").strip(), (m.TIPO or "").strip()),
                    str(m.CPBTE or ""),
                    m.FECHA.strftime("%d/%m/%Y") if m.FECHA else "",
                    format_decimal(m.CANT) if m.CANT is not None else "",
                ],
                m,
            )
            for m in movimientos
        ]
        tabla.cargar_filas(filas)
        layout.addWidget(tabla, stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)
