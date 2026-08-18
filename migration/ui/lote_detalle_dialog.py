"""Diálogo de detalle de un lote/despacho — migración de `VerDesp.frm
Sub DoVer` (líneas 474-515): artículos que componen un lote puntual,
con Precio de lista, Entrada real y Stock actual del lote."""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtWidgets import QDialog, QHBoxLayout, QPushButton, QVBoxLayout, QWidget

from migration.decimals import format_decimal
from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda

COLUMNAS = ["Sección", "Código", "Precio", "Entrada", "Stock del Lote"]
COLUMNAS_DERECHA = (2, 3, 4)


class LoteDetalleDialog(QDialog):
    def __init__(self, repos: RepositoryFactory, nrodesp: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.setWindowTitle(f"Lote {nrodesp}")
        self.resize(640, 440)

        self._construir_ui(nrodesp)

    def _construir_ui(self, nrodesp: str) -> None:
        layout = QVBoxLayout(self)

        tabla = TablaBusqueda(COLUMNAS, columnas_derecha=COLUMNAS_DERECHA)
        filas = []
        for despacho in self.repos.despacho().articulos_de_lote(nrodesp):
            cod1 = (despacho.COD1 or "").strip()
            cod2 = (despacho.COD2 or "").strip()
            articulo = self.repos.articulo().by_cod1_cod2(cod1, cod2)
            if articulo is None:
                continue
            filas.append(
                (
                    [
                        cod1, cod2,
                        format_decimal(articulo.PREC or Decimal("0")),
                        format_decimal(despacho.ENTRADA or Decimal("0")),
                        format_decimal(despacho.STOCK or Decimal("0")),
                    ],
                    despacho,
                )
            )
        tabla.cargar_filas(filas)
        layout.addWidget(tabla, stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)
