"""Diálogo de detalle de un lote/despacho — migración de `VerDesp.frm
Sub DoVer` (líneas 474-515): artículos que componen un lote puntual.

**Columnas y totales rediseñados** (pedido del usuario, 2026-08-19:
"abra una ventana flotante... con el artic(seccion y cod)/entrada/
salida/stock con totales abajo") — antes mostraba Sección/Código/
Precio/Entrada/"Stock del Lote" sin totales; ahora Artículo (Sección/
Código combinados)/Entrada/Salida/Stock, con una fila de totales al pie.
Se saca "Precio" (no lo pidió el usuario y no es un dato del lote en sí,
es del catálogo actual del Artículo) y se agrega "Salida"
(`Despachos.SALIDAS`, ya existía en el modelo, no se mostraba).

**Filtra también por `fecent`** (bug real encontrado por el usuario con
datos reales, mismo día: "22001IC05005447M dice que tiene 2 artículos y
al desplegar muestra más de 15 renglones") — ver docstring de
`DespachoRepository.articulos_de_lote()` para el detalle completo: un
mismo `NRODESP` puede repetirse en fechas de entrada distintas (2 lotes
de importación separados que reusan el número), y antes este diálogo
sólo filtraba por NRODESP, mezclando ambos lotes.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Optional

from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.decimals import format_decimal
from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda

COLUMNAS = ["Artículo", "Entrada", "Salida", "Stock"]
COL_ENTRADA, COL_SALIDA, COL_STOCK = 1, 2, 3
COLUMNAS_DERECHA = (COL_ENTRADA, COL_SALIDA, COL_STOCK)


class LoteDetalleDialog(QDialog):
    def __init__(
        self,
        repos: RepositoryFactory,
        nrodesp: str,
        fecent: Optional[date] = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.repos = repos
        titulo = f"Lote {nrodesp}"
        if fecent is not None:
            # La fecha en el título desambigua el caso real de un
            # NRODESP repetido en 2 fechas de entrada distintas (ver
            # docstring del módulo) — sin esto, 2 lotes con el mismo
            # número abrían una ventana con el mismo título.
            titulo = f"{titulo} — Entrada {fecent.strftime('%d/%m/%Y')}"
        self.setWindowTitle(titulo)
        self.resize(560, 440)

        self._construir_ui(nrodesp, fecent)

    def _construir_ui(self, nrodesp: str, fecent: Optional[date]) -> None:
        layout = QVBoxLayout(self)

        tabla = TablaBusqueda(COLUMNAS, columnas_derecha=COLUMNAS_DERECHA)
        filas = []
        total_entrada = Decimal("0")
        total_salida = Decimal("0")
        total_stock = Decimal("0")
        for despacho in self.repos.despacho().articulos_de_lote(nrodesp, fecent):
            cod1 = (despacho.COD1 or "").strip()
            cod2 = (despacho.COD2 or "").strip()
            entrada = despacho.ENTRADA or Decimal("0")
            salida = despacho.SALIDAS or Decimal("0")
            stock = despacho.STOCK or Decimal("0")
            total_entrada += entrada
            total_salida += salida
            total_stock += stock
            filas.append(
                (
                    [
                        f"{cod1}/{cod2}" if cod2 else cod1,
                        format_decimal(entrada),
                        format_decimal(salida),
                        format_decimal(stock),
                    ],
                    despacho,
                )
            )
        tabla.cargar_filas(filas)
        layout.addWidget(tabla, stretch=1)

        fila_totales = QHBoxLayout()
        lbl_titulo = QLabel("Totales :")
        lbl_titulo.setStyleSheet("font-weight: bold;")
        fila_totales.addWidget(lbl_titulo)
        for etiqueta, valor in (("Entrada :", total_entrada), ("Salida :", total_salida), ("Stock :", total_stock)):
            fila_totales.addWidget(QLabel(etiqueta))
            lbl_valor = QLabel(format_decimal(valor))
            lbl_valor.setStyleSheet("font-weight: bold;")
            fila_totales.addWidget(lbl_valor)
            fila_totales.addSpacing(16)
        fila_totales.addStretch()
        layout.addLayout(fila_totales)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)
