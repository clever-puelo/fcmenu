"""Diálogo de Descuentos por Sección del Cliente — migración de `DtosxClte.frm`.

Se abre desde el botón "Cargar Descuentos..." de `ClienteDetalleDialog`.
Cada fila = una Sección con hasta 5 niveles de descuento en cascada
(`DTO1-5`) — es la fuente real que ya usa
`FacturaService.calcular_bonificacion_cascada()` al armar una factura
(ver `FCMENU.bas Sub DxClte`).

Hallazgos confirmados (no requirieron pregunta al usuario, mismo criterio
que otros controles fantasma ya encontrados esta sesión):

- Las columnas `RGO1-3` ("Rango") del legacy se cargan en memoria
  (`FCMENU.bas Sub DxClte`) pero **nunca se leen** en ningún cálculo
  (búsqueda exhaustiva en todo el código) — y en los datos reales
  migrados, las 12.035 filas tienen `RGO1=RGO2=RGO3=0` sin excepción.
  Quedan en el modelo (`DtoxClte`) por fidelidad de esquema pero no se
  muestran acá.
- Bug real en `DtosxClte.frm Sub Grabar()` (línea 856): si el operador
  repetía una Sección en la grilla, un `Exit Sub` sin `CommitTrans` ni
  `RollbackTrans` abandonaba TODA la transacción — no se guardaba nada,
  ni siquiera lo ya insertado esa vuelta. Acá se deduplica por Sección
  en vez de abortar (ver `DtoxClteRepository.reemplazar_para_cliente`).
- El buscador de Sección por prefijo de descripción (grilla flotante
  `FG2`/`Picture2`) se reemplaza por un combo simple — hay sólo 63
  Secciones reales, entran cómodas en un dropdown sin necesitar
  filtrado incremental.
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.repository import RepositoryFactory

from .decimals import parse_decimal

COLUMNAS = ["Sección", "Descripción", "Dto.1", "Dto.2", "Dto.3", "Dto.4", "Dto.5"]
COL_SECCION, COL_DESCRIPCION, COL_DTO1 = 0, 1, 2  # Dto.2-5 son COL_DTO1+1..+4


class DtosClienteDialog(QDialog):
    """Descuentos por Sección de un cliente (equivalente a DtosxClte de
    DtosxClte.frm)."""

    def __init__(
        self,
        repos: RepositoryFactory,
        codigo_cliente: int,
        nombre_cliente: str = "",
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.repos = repos
        self.codigo_cliente = codigo_cliente

        self.setWindowTitle(f"Descuentos del Cliente {codigo_cliente} - {nombre_cliente}")
        self.resize(680, 420)

        self._secciones = self.repos.fctablas().by_ctab("SC")

        self._construir_ui()
        self._cargar_datos()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Descuentos por Tipo de Artículo"))

        self.tabla = QTableWidget(0, len(COLUMNAS))
        self.tabla.setHorizontalHeaderLabels(COLUMNAS)
        layout.addWidget(self.tabla, stretch=1)

        layout.addLayout(self._armar_barra_botones())

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        btn_agregar = QPushButton("Agregar Fila")
        btn_agregar.clicked.connect(lambda: self._agregar_fila())
        fila.addWidget(btn_agregar)

        btn_eliminar = QPushButton("Eliminar Fila")
        btn_eliminar.clicked.connect(self._on_eliminar_fila)
        fila.addWidget(btn_eliminar)

        fila.addStretch()

        self.btn_guardar = QPushButton("Grabar")
        self.btn_guardar.clicked.connect(self._on_guardar)
        fila.addWidget(self.btn_guardar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.close)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    def _combo_secciones(self, seccion_actual: str = "") -> QComboBox:
        combo = QComboBox()
        for fila in self._secciones:
            cod = (fila.COD or "").strip()
            combo.addItem(f"{cod} - {fila.DESCRI or ''}", cod)
        idx = combo.findData(seccion_actual.strip()) if seccion_actual else -1
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.currentIndexChanged.connect(lambda: self._actualizar_descripcion(combo))
        return combo

    def _actualizar_descripcion(self, combo: QComboBox) -> None:
        fila = self._fila_de_combo(combo)
        if fila is None:
            return
        idx = combo.currentIndex()
        descri = combo.itemText(idx).split(" - ", 1)[-1] if idx >= 0 else ""
        self.tabla.setItem(fila, COL_DESCRIPCION, QTableWidgetItem(descri))

    def _fila_de_combo(self, combo: QComboBox) -> int | None:
        for fila in range(self.tabla.rowCount()):
            if self.tabla.cellWidget(fila, COL_SECCION) is combo:
                return fila
        return None

    def _agregar_fila(self, seccion: str = "", dtos: list[Decimal] | None = None) -> None:
        fila = self.tabla.rowCount()
        self.tabla.insertRow(fila)

        combo = self._combo_secciones(seccion)
        self.tabla.setCellWidget(fila, COL_SECCION, combo)

        descri = combo.currentText().split(" - ", 1)[-1] if combo.count() else ""
        self.tabla.setItem(fila, COL_DESCRIPCION, QTableWidgetItem(descri))

        dtos = dtos or [Decimal("0")] * 5
        for i, valor in enumerate(dtos):
            self.tabla.setItem(fila, COL_DTO1 + i, QTableWidgetItem(str(valor)))

    def _on_eliminar_fila(self) -> None:
        fila = self.tabla.currentRow()
        if fila < 0:
            return
        respuesta = QMessageBox.question(
            self,
            "Carga Descuentos",
            "Desea Eliminar el renglón ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta == QMessageBox.StandardButton.Yes:
            self.tabla.removeRow(fila)

    # ------------------------------------------------------------------
    def _cargar_datos(self) -> None:
        for descuento in self.repos.dtoxclte().by_cliente(self.codigo_cliente):
            dtos = [
                descuento.DTO1 or Decimal("0"),
                descuento.DTO2 or Decimal("0"),
                descuento.DTO3 or Decimal("0"),
                descuento.DTO4 or Decimal("0"),
                descuento.DTO5 or Decimal("0"),
            ]
            self._agregar_fila(descuento.SECCION or "", dtos)

    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        respuesta = QMessageBox.question(
            self,
            "Descuentos del Cliente",
            "Desea Grabar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        filas = []
        for fila in range(self.tabla.rowCount()):
            combo = self.tabla.cellWidget(fila, COL_SECCION)
            if combo is None or not combo.currentData():
                continue
            datos = {"SECCION": combo.currentData()}
            for i in range(5):
                item = self.tabla.item(fila, COL_DTO1 + i)
                datos[f"DTO{i + 1}"] = parse_decimal(item.text() if item else "0")
            filas.append(datos)

        cantidad = self.repos.dtoxclte().reemplazar_para_cliente(self.codigo_cliente, filas)

        QMessageBox.information(
            self, "Descuentos del Cliente", f"{cantidad} Sección(es) grabada(s) correctamente."
        )
        self.close()
