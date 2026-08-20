"""Diálogo "Ver Saldos" de la Consulta de Cuenta Corriente — migración de
`CtaCte.frm Sub DoVer8`/`CargaFG8` (Frame1/FG8, líneas 2205-2297): saldo
de TODOS los clientes, misma fórmula real de `deuda_cliente()`/
`TIPOS_DEBE` (ver `CtascteRepository.saldos_todos_clientes()`, agregado
en una sola consulta SQL en vez del loop cliente-por-cliente del legacy).

Un clic en una fila cierra este diálogo y cambia el cliente de la
ventana principal (`CtaCteWindow._on_ver_saldos`) — feedback del
usuario, 2026-08-19: "cuando se clickee en un cliente, cierre la
ventana y... cambie el cliente". Antes pedía doble clic (mismo criterio
que otros buscadores de esta app); un solo clic alcanza para un panel
de selección tan simple como éste, mismo criterio ya aplicado a
`ComprobanteAplicarDialog`/`DespachoSelectorDialog`.
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.decimals import format_decimal
from migration.services import CuentaCorrienteService

from .widgets import TablaBusqueda

COLUMNAS = ["Código", "Cliente", "Saldo"]


class SaldosClientesDialog(QDialog):
    codigo_elegido: int | None = None

    def __init__(self, cc_service: CuentaCorrienteService, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Saldos de Cuenta Corriente — Todos los Clientes")
        self.resize(600, 560)

        filas_saldo = cc_service.saldos_todos_clientes()
        self._construir_ui(filas_saldo)

    def _construir_ui(self, filas_saldo: list[dict]) -> None:
        layout = QVBoxLayout(self)

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(0, 2))  # Código, Saldo
        self.tabla.itemClicked.connect(self._on_elegir)
        filas = [
            ([str(f["codigo"]), f["nombre"] or "", format_decimal(f["saldo"])], f)
            for f in filas_saldo
        ]
        self.tabla.cargar_filas(filas)
        layout.addWidget(self.tabla, stretch=1)

        total = sum((f["saldo"] for f in filas_saldo), Decimal("0"))
        layout.addWidget(QLabel(f"SALDO : $ {format_decimal(total)}"))

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.reject)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _on_elegir(self, item) -> None:
        fila = self.tabla.objeto_en_fila(item.row())
        if fila is None:
            return
        self.codigo_elegido = fila["codigo"]
        self.accept()
