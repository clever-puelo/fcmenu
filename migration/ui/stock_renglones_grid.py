"""`StockRenglonesGrid` — grilla interactiva estilo Excel para los
renglones de `StockMovimientoWindow`, reemplazo del botón "Agregar
Renglón..."/`RenglonStockDialog` (feedback del usuario, 2026-08-16:
*"La grilla debe ser interactiva... Igual que en el legacy con el
vsflexgrid. Toma esto como norma general para todas las grillas del
sistema."*).

Mismo mecanismo ya validado en `DetalleGrid` (Facturador) — 5 bugs
reales de Qt encontrados ahí (RecursionError re-entrante, `setFlags()`
reseteando el editor activo, `setItem()` disparando `itemChanged`
espúreo, `currentRow()/currentColumn()` no confiables en `closeEditor()`,
cierres implícitos de editor disparando navegación) se replican acá tal
cual, sin volver a pisarlos.

**Diferencia de diseño con `DetalleGrid`**: acá NO hay segmentos de
código dinámicos por Sección — la columna "Sección / Código" nunca se
tipea a mano, sólo se resuelve vía F2/Enter (`ArticuloBusquedaWindow` en
modo selección), decisión ya confirmada al migrar `Stock.frm` (evita
reconstruir el armador de segmentos en una pantalla que no lo necesita:
el Artículo siempre tiene que existir de antemano). Sólo la columna
Cantidad se tipea directo.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Optional

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QAbstractItemDelegate,
    QAbstractItemView,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from migration.decimals import format_decimal, parse_decimal
from migration.models import Articulo
from migration.repository import RepositoryFactory
from migration.services import RenglonMovimientoStock, StockMovimientoService

from .articulo_busqueda_window import ArticuloBusquedaWindow

COL_CODIGO = 0
COL_DESCRIPCION = 1
COL_STOCK = 2
COL_CANTIDAD = 3
COLUMNAS = ["Sección / Código", "Descripción", "Stock Actual", "Cantidad"]


@dataclass
class _EstadoFila:
    articulo: Optional[Articulo] = None


class StockRenglonesGrid(QTableWidget):
    def __init__(self, repos: RepositoryFactory, parent: QWidget | None = None):
        super().__init__(0, len(COLUMNAS), parent)
        self.repos = repos
        self.movimiento_service = StockMovimientoService(repos.db)

        self._estado: dict[int, _EstadoFila] = {}
        self._actualizando = False
        # Misma razón que DetalleGrid: `currentRow()/currentColumn()` no
        # es confiable dentro de `closeEditor()` (editor "zombie" que
        # cierra tarde, ya con el foco en otra celda).
        self._celda_editando: Optional[tuple[int, int]] = None

        self.setHorizontalHeaderLabels(COLUMNAS)
        # Anchos explícitos (feedback del usuario, 2026-08-19, con
        # captura de pantalla: el título "Sección / Código" quedaba
        # cortado) — el ancho por defecto de Qt para una columna nueva
        # es de apenas 100px, no le alcanza a ese título; "Cantidad"
        # (última columna) se estira para ocupar lo que sobra.
        self.setColumnWidth(COL_CODIGO, 160)
        self.setColumnWidth(COL_DESCRIPCION, 260)
        self.setColumnWidth(COL_STOCK, 110)
        self.horizontalHeader().setStretchLastSection(True)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.verticalHeader().setVisible(False)
        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self.itemChanged.connect(self._on_item_changed)

        self._agregar_fila_vacia()

    # ------------------------------------------------------------------
    @property
    def renglones(self) -> list[RenglonMovimientoStock]:
        resultado = []
        for fila in range(self.rowCount()):
            estado = self._estado.get(fila)
            if estado is None or estado.articulo is None:
                continue
            cantidad = parse_decimal(self._texto(fila, COL_CANTIDAD))
            if cantidad <= 0:
                continue
            resultado.append(
                RenglonMovimientoStock(
                    cod1=(estado.articulo.COD1 or "").strip(),
                    cod2=(estado.articulo.COD2 or "").strip(),
                    cantidad=cantidad,
                )
            )
        return resultado

    # ------------------------------------------------------------------
    # Helpers de celda (idénticos a DetalleGrid)
    # ------------------------------------------------------------------
    def _texto(self, fila: int, columna: int) -> str:
        item = self.item(fila, columna)
        return item.text().strip() if item else ""

    def _set_texto(self, fila: int, columna: int, texto: str) -> None:
        item = self.item(fila, columna)
        if item is None:
            item = QTableWidgetItem()
            self.setItem(fila, columna, item)
        item.setText(texto)

    def _set_editable(self, fila: int, columna: int, editable: bool) -> None:
        item = self.item(fila, columna)
        if item is None:
            item = QTableWidgetItem()
            self.setItem(fila, columna, item)
            ya_editable = None
        else:
            ya_editable = bool(item.flags() & Qt.ItemFlag.ItemIsEditable)
        if ya_editable == editable:
            return  # ver DetalleGrid._set_editable: setFlags() redundante resetea el editor activo
        flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        if editable:
            flags |= Qt.ItemFlag.ItemIsEditable
        item.setFlags(flags)

    def _iniciar_edicion(self, fila: int, columna: int) -> None:
        self.setCurrentCell(fila, columna)
        self.editItem(self.item(fila, columna))

    def edit(self, index, trigger=None, event=None):  # noqa: N802 (Qt override, dos firmas en C++)
        self._celda_editando = (index.row(), index.column())
        if trigger is None:
            return super().edit(index)
        return super().edit(index, trigger, event)

    def _agregar_fila_vacia(self) -> None:
        fila = self.rowCount()
        self.insertRow(fila)
        self._actualizando = True
        try:
            for columna in range(len(COLUMNAS)):
                self.setItem(fila, columna, QTableWidgetItem(""))
        finally:
            self._actualizando = False
        self._estado[fila] = _EstadoFila()
        self._actualizar_editabilidad_fila(fila)

    def _asegurar_fila_vacia_al_final(self) -> None:
        ultima = self.rowCount() - 1
        if ultima < 0 or self._estado.get(ultima, _EstadoFila()).articulo is not None:
            self._agregar_fila_vacia()

    def _actualizar_editabilidad_fila(self, fila: int) -> None:
        estado = self._estado.get(fila, _EstadoFila())
        # COL_CODIGO nunca se tipea a mano — se resuelve por F2/Enter.
        self._set_editable(fila, COL_CODIGO, False)
        self._set_editable(fila, COL_DESCRIPCION, False)
        self._set_editable(fila, COL_STOCK, False)
        self._set_editable(fila, COL_CANTIDAD, estado.articulo is not None)

    # ------------------------------------------------------------------
    # Edición de Cantidad
    # ------------------------------------------------------------------
    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._actualizando:
            return
        fila, columna = item.row(), item.column()
        # Diferido, mismo motivo que DetalleGrid: itemChanged puede
        # dispararse todavía DENTRO de closeEditor().
        QTimer.singleShot(0, lambda f=fila, c=columna: self._procesar_dato_cambiado(f, c))

    def _procesar_dato_cambiado(self, fila: int, columna: int) -> None:
        if columna != COL_CANTIDAD:
            return
        texto = self._texto(fila, columna)
        if not texto:
            return
        cantidad = parse_decimal(texto)
        if cantidad <= 0:
            QMessageBox.warning(self, "Renglón", "La cantidad tiene que ser mayor a 0.")
            self._actualizando = True
            try:
                self._set_texto(fila, columna, "")
            finally:
                self._actualizando = False
            self._iniciar_edicion(fila, columna)
            return
        self._actualizando = True
        try:
            self._set_texto(fila, columna, format_decimal(cantidad))
        finally:
            self._actualizando = False

    # ------------------------------------------------------------------
    # Selector de Artículo (F2 / Enter en Código vacío)
    # ------------------------------------------------------------------
    def _abrir_selector_articulo(self, fila: int) -> None:
        dialogo = ArticuloBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.articulo_elegido is None:
            return
        # Objeto de la sesión propia de ArticuloBusquedaWindow -- se
        # re-busca en self.repos.db, mismo criterio que Facturador/Recibo.
        articulo = self.repos.articulo().by_cod1_cod2(
            dialogo.articulo_elegido.COD1.strip(), dialogo.articulo_elegido.COD2.strip()
        )
        if articulo is not None:
            self._aplicar_articulo(fila, articulo)

    def _aplicar_articulo(self, fila: int, articulo: Articulo) -> None:
        cod1 = (articulo.COD1 or "").strip()
        cod2 = (articulo.COD2 or "").strip()
        try:
            self.movimiento_service.validar_seccion(cod1)
        except ValueError as exc:
            QMessageBox.warning(self, "Renglón", str(exc))
            return

        estado = self._estado.setdefault(fila, _EstadoFila())
        estado.articulo = articulo

        stock = self.repos.stock().by_cod1_cod2(cod1, cod2)
        stock_actual = stock.STUNID if stock and stock.STUNID is not None else Decimal("0")

        self._actualizando = True
        try:
            self._set_texto(fila, COL_CODIGO, f"{cod1} / {cod2}")
            self._set_texto(fila, COL_DESCRIPCION, (articulo.DESCRI or "").strip())
            self._set_texto(fila, COL_STOCK, format_decimal(stock_actual))
        finally:
            self._actualizando = False

        self._actualizar_editabilidad_fila(fila)
        self._asegurar_fila_vacia_al_final()
        self._iniciar_edicion(fila, COL_CANTIDAD)

    # ------------------------------------------------------------------
    # Borrado de renglón (Supr, confirma)
    # ------------------------------------------------------------------
    def _eliminar_fila_actual(self) -> None:
        fila = self.currentRow()
        if fila < 0:
            return
        estado = self._estado.get(fila)
        if estado is None or estado.articulo is None:
            return  # fila vacía, nada que borrar

        respuesta = QMessageBox.question(
            self,
            "Renglón",
            "¿Desea eliminar el renglón?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.removeRow(fila)
        nuevo_estado = {}
        for indice, valor in self._estado.items():
            if indice == fila:
                continue
            nuevo_estado[indice - 1 if indice > fila else indice] = valor
        self._estado = nuevo_estado
        self._asegurar_fila_vacia_al_final()

    # ------------------------------------------------------------------
    # Teclado
    # ------------------------------------------------------------------
    def keyPressEvent(self, event) -> None:  # noqa: N802 (Qt override)
        if event.key() == Qt.Key.Key_F2:
            fila = self.currentRow()
            if fila >= 0:
                self._abrir_selector_articulo(fila)
            event.accept()
            return
        if event.key() == Qt.Key.Key_Delete and self.state() != QAbstractItemView.State.EditingState:
            self._eliminar_fila_actual()
            event.accept()
            return
        if (
            event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
            and self.currentColumn() == COL_CODIGO
            and self.state() != QAbstractItemView.State.EditingState
        ):
            fila = self.currentRow()
            if fila >= 0:
                self._abrir_selector_articulo(fila)
            event.accept()
            return
        super().keyPressEvent(event)

    def closeEditor(self, editor, hint) -> None:  # noqa: N802 (Qt override)
        celda = self._celda_editando
        self._celda_editando = None
        super().closeEditor(editor, hint)

        if celda is None or hint != QAbstractItemDelegate.EndEditHint.SubmitModelCache:
            return

        fila, columna = celda
        if columna == COL_CANTIDAD:
            QTimer.singleShot(0, lambda f=fila: self._despues_de_cantidad(f))

    def _despues_de_cantidad(self, fila: int) -> None:
        siguiente = fila + 1
        self._asegurar_fila_vacia_al_final()
        if siguiente < self.rowCount():
            self.setCurrentCell(siguiente, COL_CODIGO)

    # ------------------------------------------------------------------
    def foco_inicial(self) -> None:
        """Foco en la celda de Código de la primera fila lista para
        Enter/F2 — se llama al presionar Enter en Nº de Comprobante
        (feedback del usuario, 2026-08-16: "al dar enter en comprobante
        pase a la carga en la grilla")."""
        if self.rowCount() == 0:
            self._agregar_fila_vacia()
        fila = next((f for f in range(self.rowCount()) if self._estado.get(f, _EstadoFila()).articulo is None), 0)
        self.setCurrentCell(fila, COL_CODIGO)
        self.setFocus()

    def reiniciar(self) -> None:
        self._actualizando = True
        try:
            self.setRowCount(0)
        finally:
            self._actualizando = False
        self._estado = {}
        self._agregar_fila_vacia()
