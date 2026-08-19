"""Modificación de Precios por Porcentaje — migración de `ModPrec.frm`.

Herramienta de acción masiva (no es un ABM por-registro, no usa el
patrón Búsqueda→Detalle): aplica un `%` a `Articulo.PREC`, en dos modos
— "Por Sección" (con rango opcional) o "Todos" el catálogo.

Hallazgos confirmados leyendo el código (no se preguntaron al usuario
por caer en el mismo criterio ya establecido para controles fantasma):

- El legacy muestra un SEGUNDO par DESDE/HASTA (Text1/Text3) para
  Secciones con dos dimensiones de código (ej. GPN: mm + telas), pero
  `CargaGrilla()` nunca los lee — sólo filtra con el primer par
  tratando el `COD2` completo como un único número (`Val()`). Ese
  segundo filtro nunca tuvo efecto real, no se migra — ver
  `ArticuloService.candidatos_modificacion_precio`.
- La variable legacy que calcula el precio nuevo se llama "PCos"
  (parece "Precio de Costo") pero el resultado se graba en `PREC` — es
  un nombre de variable engañoso, no una regla de negocio; acá se llama
  `precio_nuevo` sin ambigüedad.
- Simplificación de UX: el legacy resetea la grilla de vista previa
  después de "Modificar" sin importar si se confirmó, se rechazó o se
  canceló el diálogo de "¿Desea continuar?". Acá sólo se resetea tras
  una confirmación real (igual criterio ya aplicado en Cotización del
  Dólar).
"""

from __future__ import annotations

from decimal import Decimal
from typing import Optional

from PyQt6.QtWidgets import (
    QComboBox,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.models import Articulo
from migration.repository import RepositoryFactory
from migration.services import ArticuloService

from .decimals import format_decimal, parse_decimal
from .widgets import compactar_alto_filas


class ModPreciosWindow(QMainWindow):
    """Modificación de Precios por Porcentaje (equivalente a ModPrec de ModPrec.frm)."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Modificación de Precios por Porcentaje")
        self.resize(760, 560)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.articulo_service = ArticuloService(self.db)

        self.candidatos: list[Articulo] = []

        self._construir_ui()
        self._cargar_secciones()
        self._on_modo_cambiado()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addLayout(self._armar_fila_modo())
        layout.addWidget(self._armar_grupo_seccion())

        # El panel de Artículos sube y crece (feedback del usuario,
        # 2026-08-19: "usar todo el alto del panel central... subir el
        # panel de artículos, hacerlo más grande") — ya no compite con
        # el panel de Porcentaje aparte, que se sacó del todo (ver
        # `_armar_fila_modo`, ahora en un recuadro en esa misma línea).
        self.tabla = QTableWidget(0, 5)
        self.tabla.setHorizontalHeaderLabels(
            ["Sección", "Código", "Precio Actual", "Precio Nuevo", "Stock"]
        )
        self.tabla.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        # Filas 10% más cortas (feedback del usuario, 2026-08-19), mismo
        # criterio que el Detalle del Facturador y el Recibo: entran más
        # renglones a la vista.
        compactar_alto_filas(self.tabla)
        layout.addWidget(self.tabla, stretch=1)

        self.lbl_cantidad = QLabel("")
        layout.addWidget(self.lbl_cantidad)

        layout.addLayout(self._armar_barra_botones())

    def _armar_fila_modo(self) -> QHBoxLayout:
        fila = QHBoxLayout()
        self.radio_por_seccion = QRadioButton("Por Sección")
        self.radio_por_seccion.setChecked(True)
        self.radio_por_seccion.toggled.connect(self._on_modo_cambiado)
        self.radio_todos = QRadioButton("Todos")
        fila.addWidget(self.radio_por_seccion)
        fila.addWidget(self.radio_todos)
        fila.addStretch()

        # Porcentaje en un recuadro arriba a la derecha, en la misma
        # línea de "Por Sección/Todos" (feedback del usuario, 2026-08-19)
        # — antes vivía en su propio panel debajo de Sección (el
        # "segundo panel", ya sacado del todo).
        recuadro_porcentaje = QFrame()
        recuadro_porcentaje.setFrameShape(QFrame.Shape.Box)
        recuadro_porcentaje.setStyleSheet("QFrame { border: 1px solid #888; border-radius: 6px; }")
        fila_porcentaje = QHBoxLayout(recuadro_porcentaje)
        fila_porcentaje.setContentsMargins(10, 4, 10, 4)
        fila_porcentaje.addWidget(QLabel("Porcentaje :"))
        self.txt_porcentaje = QLineEdit("0")
        self.txt_porcentaje.setMaximumWidth(80)
        self.txt_porcentaje.setToolTip("En más o menos — colocar signo '-' adelante para descuento")
        self.txt_porcentaje.editingFinished.connect(self._actualizar_vista)
        fila_porcentaje.addWidget(self.txt_porcentaje)
        fila.addWidget(recuadro_porcentaje)
        return fila

    def _armar_grupo_seccion(self) -> QGroupBox:
        self.grupo_seccion = QGroupBox("Sección")
        fila = QHBoxLayout(self.grupo_seccion)

        fila.addWidget(QLabel("Sección :"))
        self.combo_seccion = QComboBox()
        self.combo_seccion.currentIndexChanged.connect(self._actualizar_vista)
        fila.addWidget(self.combo_seccion)

        fila.addWidget(QLabel("Código desde :"))
        self.txt_desde = QLineEdit()
        self.txt_desde.setMaximumWidth(90)
        self.txt_desde.setToolTip(
            "Filtra por el valor numérico completo del código (Val()), no por segmento — "
            "dejar vacío para no acotar."
        )
        self.txt_desde.editingFinished.connect(self._actualizar_vista)
        fila.addWidget(self.txt_desde)

        fila.addWidget(QLabel("hasta :"))
        self.txt_hasta = QLineEdit()
        self.txt_hasta.setMaximumWidth(90)
        self.txt_hasta.editingFinished.connect(self._actualizar_vista)
        fila.addWidget(self.txt_hasta)

        fila.addStretch()
        return self.grupo_seccion

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        self.btn_anular = QPushButton("Anular")
        self.btn_anular.clicked.connect(self._on_anular)
        fila.addWidget(self.btn_anular)

        fila.addStretch()

        self.btn_modificar = QPushButton("Modificar")
        self.btn_modificar.setEnabled(False)
        self.btn_modificar.clicked.connect(self._on_modificar)
        fila.addWidget(self.btn_modificar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.close)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    def _cargar_secciones(self) -> None:
        secciones = self.repos.fctablas().by_ctab("SC")
        self.combo_seccion.clear()
        for fila in secciones:
            cod = (fila.COD or "").strip()
            self.combo_seccion.addItem(f"{cod} - {fila.DESCRI or ''}", cod)

    def _on_modo_cambiado(self) -> None:
        self.grupo_seccion.setEnabled(self.radio_por_seccion.isChecked())
        self._actualizar_vista()

    # ------------------------------------------------------------------
    def _entero_opcional(self, campo: QLineEdit) -> Optional[int]:
        texto = campo.text().strip()
        if not texto:
            return None
        try:
            return int(texto)
        except ValueError:
            return None

    def _actualizar_vista(self) -> None:
        cod1 = self.combo_seccion.currentData() if self.radio_por_seccion.isChecked() else None
        desde = self._entero_opcional(self.txt_desde) if cod1 else None
        hasta = self._entero_opcional(self.txt_hasta) if cod1 else None
        porcentaje = parse_decimal(self.txt_porcentaje.text())

        self.candidatos = self.articulo_service.candidatos_modificacion_precio(
            cod1=cod1, desde=desde, hasta=hasta
        )

        self.tabla.setRowCount(0)
        for art in self.candidatos:
            precio_actual = art.PREC or Decimal("0")
            precio_nuevo = self.articulo_service.calcular_precio_con_porcentaje(precio_actual, porcentaje)
            fila = self.tabla.rowCount()
            self.tabla.insertRow(fila)
            self.tabla.setItem(fila, 0, QTableWidgetItem((art.COD1 or "").strip()))
            self.tabla.setItem(fila, 1, QTableWidgetItem((art.COD2 or "").strip()))
            self.tabla.setItem(fila, 2, QTableWidgetItem(f"$ {format_decimal(precio_actual)}"))
            self.tabla.setItem(fila, 3, QTableWidgetItem(f"$ {format_decimal(precio_nuevo)}"))
            self.tabla.setItem(fila, 4, QTableWidgetItem(str(art.STOCK if art.STOCK is not None else 0)))

        self.lbl_cantidad.setText(f"Artículos seleccionados : {len(self.candidatos)}")
        self.btn_modificar.setEnabled(len(self.candidatos) > 0)

    def _on_anular(self) -> None:
        self.txt_desde.clear()
        self.txt_hasta.clear()
        self.txt_porcentaje.setText("0")
        self._actualizar_vista()

    # ------------------------------------------------------------------
    def _on_modificar(self) -> None:
        porcentaje = parse_decimal(self.txt_porcentaje.text())

        respuesta = QMessageBox.question(
            self,
            "Actualización de Precios",
            "Desea Cargar los nuevos precios ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        cantidad = self.articulo_service.aplicar_modificacion_precio(self.candidatos, porcentaje)

        QMessageBox.information(
            self, "Actualización de Precios", f"{cantidad} artículo(s) modificado(s) correctamente."
        )
        self._on_anular()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
