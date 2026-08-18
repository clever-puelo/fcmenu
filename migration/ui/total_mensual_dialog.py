"""Diálogo de "Total Mensual" (botón/F4) — migración de `TotFact.frm
Sub MuestraTotal` (líneas 2059-2107), con la corrección de rango y de
la fórmula de "Proyección" confirmadas con el usuario (2026-08-16) —
ver el docstring largo de `TotalesDiariosService.resumen_mensual`."""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.decimals import format_decimal
from migration.services import ResumenMensual


class TotalMensualDialog(QDialog):
    def __init__(self, resumen: ResumenMensual, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Total Mensual")
        self.resize(420, 380)

        self._construir_ui(resumen)

    def _construir_ui(self, r: ResumenMensual) -> None:
        layout = QVBoxLayout(self)

        grupo_dias = QGroupBox()
        form_dias = QFormLayout(grupo_dias)
        form_dias.addRow("Días Hábiles :", QLabel(str(r.dias_habiles_mes)))
        form_dias.addRow("Días Trabajados :", QLabel(str(r.dias_trabajados)))
        layout.addWidget(grupo_dias)

        rango = f"{r.fecha_desde.strftime('%d/%m/%Y')} al {r.fecha_hasta.strftime('%d/%m/%Y')}"
        grupo_tot = QGroupBox(f"Totales a la Fecha ({rango})")
        form_tot = QFormLayout(grupo_tot)
        form_tot.addRow("Precio Costo :", QLabel(f"$ {format_decimal(r.total_precio_costo)}"))
        form_tot.addRow("Precio Venta :", QLabel(f"$ {format_decimal(r.total_precio_venta)}"))
        form_tot.addRow("Venta Real :", QLabel(f"$ {format_decimal(r.total_venta_real)}"))
        form_tot.addRow("Unidades :", QLabel(str(r.total_unidades)))
        layout.addWidget(grupo_tot)

        grupo_proy = QGroupBox()
        form_proy = QFormLayout(grupo_proy)
        form_proy.addRow("Prom. Diario :", QLabel(f"$ {format_decimal(r.promedio_diario)}"))
        form_proy.addRow("Unidad Prom. :", QLabel(format_decimal(r.unidad_promedio)))
        lbl_proyeccion = QLabel(f"$ {format_decimal(r.proyeccion)}")
        lbl_proyeccion.setStyleSheet("font-weight: bold;")
        form_proy.addRow("Proyección :", lbl_proyeccion)
        layout.addWidget(grupo_proy)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)
