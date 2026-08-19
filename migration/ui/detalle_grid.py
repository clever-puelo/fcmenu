"""`DetalleGrid` — grilla interactiva estilo Excel para el Detalle de
Factura, reemplazo directo de `RenglonDetalleDialog`/botones Agregar-
Editar-Quitar (decisión del usuario 2026-08-16, probando con datos
reales: *"Es muy estático... En el legacy es totalmente distinto. Hay
una grilla interactiva... Es como en Excel, se va cargando y navegando
en ella. No hay carga en una ventana y pasaje"*). Fiel al `FG1` FlexGrid
real de `DetFact.frm`, que es exactamente así: una grilla con columnas
fijas donde el operador tipea directo, Enter avanza campo, y `Posi()`
decide en cuál columna cae cada segmento dinámico de la Sección.

Interacción confirmada:
- Al llegar (después de elegir Cliente), el foco arranca en la primera
  celda (Sección) de la primera fila, editando.
- Enter en la celda de Sección vacía, o **F2 en cualquier celda**, abre
  el selector de Artículo (`ArticuloBusquedaWindow` en modo selección).
- Tipear una Sección real (o `**` para ítem libre) resuelve los
  segmentos de esa fila (`resolver_seccion_renglon`) — sólo quedan
  editables las columnas de segmento que esa Sección realmente usa
  (`SegmentoRenglon.posicion` → columna fija, igual que `Posi()` en el
  legacy).
- Al completar el segmento de cantidad (Unidad de Facturación), se
  disparan: cálculo de Precio Unitario/Importe, y el auto-popup del
  selector de Lote/Despacho (fiel al legacy, ya implementado en
  `DespachoSelectorDialog`).
- Para modificar un renglón ya cargado: subir con las flechas y
  sobreescribir cualquier celda — recalcula solo.
- Para borrar un renglón: `Supr`, confirma, y las filas siguientes se
  reordenan (no queda un hueco).
- Siempre hay UNA fila vacía al final, lista para seguir cargando.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Callable, Optional

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QAbstractItemDelegate,
    QAbstractItemView,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from migration.models import Articulo, Cliente
from migration.repository import RepositoryFactory
from migration.services import FacturaService, RenglonEmision

from .articulo_busqueda_window import ArticuloBusquedaWindow
from .articulo_codigo import armar_codigo, formatear_segmento, separar_codigo
from .decimals import format_decimal, parse_decimal
from .despacho_selector_dialog import DespachoSelectorDialog
from .factura_renglon import (
    SECCION_ITEM_LIBRE,
    SeccionInexistenteError,
    SeccionRenglon,
    SeccionSinUnidadFacturacionError,
    armar_codigo_renglon,
    calcular_precio_e_importe,
    calcular_precio_preview,
    descripcion_con_segmentos,
    resolver_precio_articulo,
    resolver_seccion_renglon,
)
from .nota_articulo_dialog import NotaArticuloDialog
from .widgets import compactar_alto_filas

# Columnas fijas — mismo criterio que FG1 del legacy: los segmentos
# dinámicos (ALF1-7) siempre caen en la misma columna física según su
# "Posición" (Fctabla1(UM).NUMSD3), independientemente de la Sección.
COL_SECCION = 0
COL_POS1 = 1  # Nro/Pulg
COL_POS2 = 2  # Mtr/Kg
COL_POS3 = 3  # MM
COL_POS4 = 4  # Telas
COL_DESCRIPCION = 5
COL_POS6 = 6  # Unidades — casi siempre la Unidad de Facturación
COL_PRECIO = 7
COL_IMPORTE = 8
COL_LOTE = 9

POSICION_A_COLUMNA = {1: COL_POS1, 2: COL_POS2, 3: COL_POS3, 4: COL_POS4, 6: COL_POS6}

# Encabezados con leyenda real en vez de los números de columna del
# FlexGrid legacy (feedback del usuario, 2026-08-15: "la grilla no tiene
# títulos claros, tiene números en la cabecera") — cada columna dinámica
# conserva el nombre de la Unidad de Medida que casi siempre cae en esa
# posición física (ver comentario de POSICION_A_COLUMNA más arriba).
COLUMNAS = ["Sección", "Nro/Pulg", "Mtr/Kg", "MM", "Telas", "Descripción", "Cantidad", "Precio Unit.", "Importe", "Lote"]

# Columnas que nunca se tipean directo: son resultado de un cálculo o de
# elegir un Artículo/Lote.
COLUMNAS_SIEMPRE_BLOQUEADAS = {COL_IMPORTE, COL_LOTE}


@dataclass
class _EstadoFila:
    seccion: Optional[SeccionRenglon] = None
    articulo: Optional[Articulo] = None
    nrodesp_elegido: Optional[str] = None
    precio_base: Decimal = field(default_factory=lambda: Decimal("0"))


class DetalleGrid(QTableWidget):
    def __init__(
        self,
        repos: RepositoryFactory,
        cliente_actual: Callable[[], Optional[Cliente]],
        en_dolares: Callable[[], bool],
        cotizacion: Callable[[], Decimal],
        al_cambiar: Callable[[], None],
        parent: QWidget | None = None,
    ):
        super().__init__(0, len(COLUMNAS), parent)
        self.repos = repos
        self._cliente_actual = cliente_actual
        self._en_dolares = en_dolares
        self._cotizacion = cotizacion
        self._al_cambiar = al_cambiar  # callback: recalcular totales del Pie

        self._estado: dict[int, _EstadoFila] = {}
        self._actualizando = False
        # Celda que ESTA VENTANA abrió a editar más recientemente (no la
        # que Qt cree que está activa: `self.currentRow()/currentColumn()`
        # no es confiable dentro de `closeEditor()` cuando se navega de
        # celda en celda muy rápido — encontrado con teclado real, ver
        # `_iniciar_edicion()`/`closeEditor()`).
        self._celda_editando: Optional[tuple[int, int]] = None

        self.setHorizontalHeaderLabels(COLUMNAS)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.verticalHeader().setVisible(False)
        # Filas 10% más cortas (feedback del usuario, 2026-08-19): entran
        # más renglones a la vista sin achicar la letra.
        compactar_alto_filas(self)
        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self.itemChanged.connect(self._on_item_changed)

        self._agregar_fila_vacia()

    # ------------------------------------------------------------------
    # Resultado — sólo las filas resueltas (con Sección + cantidad válida)
    # ------------------------------------------------------------------
    @property
    def renglones(self) -> list[RenglonEmision]:
        resultado = []
        for fila in range(self.rowCount()):
            renglon = self._armar_renglon(fila)
            if renglon is not None:
                resultado.append(renglon)
        return resultado

    def _armar_renglon(self, fila: int) -> Optional[RenglonEmision]:
        estado = self._estado.get(fila)
        if estado is None or estado.seccion is None:
            return None
        importe_texto = self._texto(fila, COL_IMPORTE)
        if not importe_texto:
            return None

        campos_cantidad = dict(pulg=Decimal("0"), mtr=Decimal("0"), milim=0, telas=0, cantidad_unidades=Decimal("0"))
        nombre_por_posicion = {1: "pulg", 2: "mtr", 3: "milim", 4: "telas", 6: "cantidad_unidades"}
        for segmento in estado.seccion.segmentos_cantidad:
            columna = POSICION_A_COLUMNA[segmento.posicion]
            valor = parse_decimal(self._texto(fila, columna))
            nombre_campo = nombre_por_posicion[segmento.posicion]
            campos_cantidad[nombre_campo] = int(valor) if nombre_campo in ("milim", "telas") else valor

        cod2 = ""
        if estado.seccion.segmentos_codigo:
            valores_codigo = [
                parse_decimal(self._texto(fila, POSICION_A_COLUMNA[s.posicion])) for s in estado.seccion.segmentos_codigo
            ]
            cod2 = armar_codigo_renglon(estado.seccion, valores_codigo)

        return RenglonEmision(
            cod1=estado.seccion.cod_seccion,
            cod2=cod2,
            descripcion=self._texto(fila, COL_DESCRIPCION),
            precio_unitario=parse_decimal(self._texto(fila, COL_PRECIO)),
            importe=parse_decimal(importe_texto),
            nrodesp_elegido=estado.nrodesp_elegido,
            **campos_cantidad,
        )

    # ------------------------------------------------------------------
    # Helpers de celda
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
            # Bug real encontrado probando con teclado de verdad:
            # `item.setFlags(...)` sobre la celda que se está editando EN
            # ESE MOMENTO le resetea el texto al editor activo (Qt
            # re-sincroniza el editor abierto cuando el modelo "cambia",
            # aunque el flag resultante sea idéntico al que ya tenía) —
            # típicamente pasa con la propia celda Sección, que dispara
            # `_actualizar_editabilidad_fila` para sí misma apenas se
            # resuelve. Evitar el `setFlags()` redundante lo soluciona.
            return
        flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        if editable:
            flags |= Qt.ItemFlag.ItemIsEditable
        item.setFlags(flags)

    def _iniciar_edicion(self, fila: int, columna: int) -> None:
        """Punto de conveniencia para pasar a editar una celda desde
        código propio (navegación tras Enter, F2, foco inicial). El
        registro real de `self._celda_editando` pasa por `edit()` (ver
        más abajo), así que esto también cubre ediciones que arranca el
        propio operador (clic, flechas + tipear) sin pasar por acá."""
        self.setCurrentCell(fila, columna)
        self.editItem(self.item(fila, columna))

    def edit(self, index, trigger=None, event=None):  # noqa: N802 (Qt override, dos firmas en C++)
        """Todo camino de edición (`editItem()`, F2, doble clic,
        `AnyKeyPressed`) pasa por acá — es el único lugar confiable para
        saber qué celda se está por editar. `closeEditor()` usaba
        `self.currentRow()/currentColumn()`, que no es confiable cuando
        se navega de celda en celda muy rápido (encontrado con teclado
        real: un editor "zombie" de una celda anterior podía cerrarse
        tarde, con `currentRow/currentColumn` ya apuntando a otra celda
        distinta, y el dato terminaba escribiéndose en el renglón
        equivocado)."""
        self._celda_editando = (index.row(), index.column())
        if trigger is None:
            return super().edit(index)
        return super().edit(index, trigger, event)

    def _agregar_fila_vacia(self) -> None:
        fila = self.rowCount()
        self.insertRow(fila)
        # `setItem()` en filas/columnas nuevas dispara `itemChanged` igual
        # que si el operador hubiera tipeado — sin este guard, crear una
        # fila en blanco encolaba (vía `QTimer.singleShot(0, ...)`, ver
        # `_on_item_changed`) un procesamiento espúreo por cada una de las
        # 10 celdas nuevas, que terminaba disparándose en un momento
        # impredecible (a veces mientras otra celda ya estaba en edición)
        # y le pisaba el editor activo — encontrado con teclado real.
        self._actualizando = True
        try:
            for columna in range(len(COLUMNAS)):
                item = QTableWidgetItem("")
                self.setItem(fila, columna, item)
        finally:
            self._actualizando = False
        self._estado[fila] = _EstadoFila()
        self._actualizar_editabilidad_fila(fila)

    def _asegurar_fila_vacia_al_final(self) -> None:
        ultima = self.rowCount() - 1
        if ultima < 0 or self._estado.get(ultima, _EstadoFila()).seccion is not None:
            self._agregar_fila_vacia()

    def _actualizar_editabilidad_fila(self, fila: int) -> None:
        estado = self._estado.get(fila, _EstadoFila())
        seccion = estado.seccion

        self._set_editable(fila, COL_SECCION, True)
        for columna in (COL_IMPORTE, COL_LOTE):
            self._set_editable(fila, columna, False)

        if seccion is None:
            for columna in (COL_POS1, COL_POS2, COL_POS3, COL_POS4, COL_DESCRIPCION, COL_POS6, COL_PRECIO):
                self._set_editable(fila, columna, False)
            return

        es_libre = seccion.cod_seccion == SECCION_ITEM_LIBRE
        posiciones_activas = {s.posicion for s in seccion.segmentos_codigo} | {s.posicion for s in seccion.segmentos_cantidad}
        for posicion, columna in POSICION_A_COLUMNA.items():
            self._set_editable(fila, columna, posicion in posiciones_activas)
        # Descripción/Precio: en ítem libre los tipea el operador; si hay
        # Artículo real, se autogeneran y no se tocan a mano.
        self._set_editable(fila, COL_DESCRIPCION, es_libre)
        self._set_editable(fila, COL_PRECIO, es_libre)

    # ------------------------------------------------------------------
    # Edición de celdas
    # ------------------------------------------------------------------
    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._actualizando:
            return
        fila, columna = item.row(), item.column()
        # Diferido a la próxima vuelta del loop de eventos: `itemChanged`
        # puede dispararse TODAVÍA DENTRO de `closeEditor()`/`super().
        # closeEditor()` (Qt confirma el dato antes de cerrar el editor
        # del todo) — procesar acá mismo (resolver Sección, quizás abrir
        # el selector de Artículo o cambiar flags de edición de la propia
        # celda que se está cerrando) generaba un `RecursionError` real
        # (encontrado probando la grilla con teclado de verdad). Con
        # `QTimer.singleShot(0, ...)` se corre recién cuando el cierre
        # del editor terminó del todo.
        QTimer.singleShot(0, lambda f=fila, c=columna: self._procesar_dato_cambiado(f, c))

    def _procesar_dato_cambiado(self, fila: int, columna: int) -> None:
        if columna == COL_SECCION:
            self._on_seccion_editada(fila, self._texto(fila, columna))
        elif columna in (COL_POS1, COL_POS2, COL_POS3, COL_POS4, COL_POS6, COL_PRECIO):
            self._reconsiderar_codigo_y_recalcular(fila)

        self._al_cambiar()

    def _on_seccion_editada(self, fila: int, texto: str) -> None:
        texto = texto.strip().upper()
        estado = self._estado.setdefault(fila, _EstadoFila())
        if not texto:
            estado.seccion = None
            estado.articulo = None
            self._actualizar_editabilidad_fila(fila)
            return

        try:
            seccion = resolver_seccion_renglon(self.repos.fctablas(), texto)
        except SeccionInexistenteError:
            QMessageBox.warning(self, "Detalle", f"La Sección '{texto}' no existe.")
            self._limpiar_celda(fila, COL_SECCION)
            return
        except SeccionSinUnidadFacturacionError:
            QMessageBox.warning(self, "Detalle", f"La Sección '{texto}' no tiene Unidad de Facturación cargada.")
            self._limpiar_celda(fila, COL_SECCION)
            return

        estado.seccion = seccion
        estado.articulo = None
        self._actualizando = True
        try:
            self._set_texto(fila, COL_SECCION, seccion.cod_seccion)
        finally:
            self._actualizando = False
        self._actualizar_editabilidad_fila(fila)
        self._asegurar_fila_vacia_al_final()

        if not seccion.segmentos_codigo and seccion.cod_seccion != SECCION_ITEM_LIBRE:
            # Sección sin segmentos de código (COD2 siempre "") — no hay
            # ninguna celda de código que "completar" para disparar la
            # búsqueda del Artículo (ver `_reconsiderar_codigo_y_
            # recalcular`), así que se intenta acá mismo, apenas se
            # resuelve la Sección.
            articulo = self.repos.articulo().by_cod1_cod2(seccion.cod_seccion, "")
            if articulo is not None:
                self._aplicar_articulo(fila, articulo)

    def _limpiar_celda(self, fila: int, columna: int) -> None:
        self._actualizando = True
        try:
            self._set_texto(fila, columna, "")
        finally:
            self._actualizando = False

    def _limpiar_fila(self, fila: int) -> None:
        """Vacía por completo una fila y reinicia su estado — usado cuando
        hay que "volver al primer campo de la Sección" (ej. Artículo sin
        Precio cargado): no tiene sentido dejar segmentos de código ya
        tipeados si la fila entera se descarta."""
        self._estado[fila] = _EstadoFila()
        self._actualizando = True
        try:
            for columna in range(len(COLUMNAS)):
                self._set_texto(fila, columna, "")
        finally:
            self._actualizando = False
        self._actualizar_editabilidad_fila(fila)

    def _reconsiderar_codigo_y_recalcular(self, fila: int) -> None:
        """Si se completaron todos los segmentos de código a mano (sin
        pasar por F2), intenta resolver el Artículo — mismo criterio que
        el botón "Buscar por Código" que tenía `RenglonDetalleDialog`,
        ahora disparado solo al completar la última celda de código."""
        estado = self._estado.get(fila)
        if estado is None or estado.seccion is None:
            return

        if estado.articulo is None and estado.seccion.segmentos_codigo:
            valores = []
            completos = True
            for segmento in estado.seccion.segmentos_codigo:
                columna = POSICION_A_COLUMNA[segmento.posicion]
                texto = self._texto(fila, columna)
                if not texto:
                    completos = False
                    break
                valores.append(parse_decimal(texto))
            if completos:
                cod2 = armar_codigo(estado.seccion.segmentos_codigo, valores)
                articulo = self.repos.articulo().by_cod1_cod2(estado.seccion.cod_seccion, cod2)
                if articulo is not None:
                    self._aplicar_articulo(fila, articulo)
                    return  # _aplicar_articulo ya recalcula
                QMessageBox.information(
                    self, "Detalle", "No se encontró un Artículo con ese código — probá F2 para buscarlo."
                )

        self._recalcular_fila(fila)

    def _recalcular_fila(self, fila: int) -> None:
        estado = self._estado.get(fila)
        if estado is None or estado.seccion is None:
            return

        if estado.seccion.cod_seccion == SECCION_ITEM_LIBRE:
            estado.precio_base = parse_decimal(self._texto(fila, COL_PRECIO))

        valores_cantidad = {}
        for segmento in estado.seccion.segmentos_cantidad:
            columna = POSICION_A_COLUMNA[segmento.posicion]
            valores_cantidad[segmento.alf_index] = parse_decimal(self._texto(fila, columna))

        # Precio "en construcción" (feedback del usuario, 2026-08-15): se
        # muestra apenas se conoce el precio base del Artículo, y se va
        # actualizando solo a medida que se completan los segmentos de
        # cálculo (ALF4-6) — no hace falta esperar a que la Cantidad
        # (Unidad de Facturación) esté cargada para ver el Precio Unitario.
        precio_preview = calcular_precio_preview(estado.precio_base, valores_cantidad, estado.seccion)
        self._actualizando = True
        try:
            self._set_texto(fila, COL_PRECIO, format_decimal(precio_preview))
        finally:
            self._actualizando = False

        try:
            precio_unitario, importe = calcular_precio_e_importe(estado.precio_base, valores_cantidad, estado.seccion)
        except (ValueError, SeccionSinUnidadFacturacionError):
            self._actualizando = True
            try:
                self._set_texto(fila, COL_IMPORTE, "")
            finally:
                self._actualizando = False
            return

        self._actualizando = True
        try:
            self._set_texto(fila, COL_PRECIO, format_decimal(precio_unitario))
            self._set_texto(fila, COL_IMPORTE, format_decimal(importe))
        finally:
            self._actualizando = False

    # ------------------------------------------------------------------
    # Selector de Artículo (F2 / Enter en Sección vacía / código completo)
    # ------------------------------------------------------------------
    def _abrir_selector_articulo(self, fila: int) -> None:
        dialogo = ArticuloBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.articulo_elegido is not None:
            self._aplicar_articulo(fila, dialogo.articulo_elegido)
        else:
            self._iniciar_edicion(fila, COL_SECCION)

    def _aplicar_articulo(self, fila: int, articulo: Articulo) -> None:
        cod1 = (articulo.COD1 or "").strip()
        cod2 = (articulo.COD2 or "").strip()
        try:
            seccion = resolver_seccion_renglon(self.repos.fctablas(), cod1)
        except (SeccionInexistenteError, SeccionSinUnidadFacturacionError):
            QMessageBox.warning(self, "Detalle", "La Sección del artículo elegido no es válida para facturar.")
            return

        estado = self._estado.setdefault(fila, _EstadoFila())
        estado.seccion = seccion
        estado.articulo = articulo

        valores_codigo = separar_codigo(cod2, seccion.segmentos_codigo) if seccion.segmentos_codigo else []
        if seccion.usar_descripcion_seccion:
            descripcion = descripcion_con_segmentos(seccion, seccion.descripcion, valores_codigo)
        else:
            descripcion = (articulo.DESCRI or "").strip()
        precio_base = resolver_precio_articulo(
            articulo.PREC or Decimal("0"), seccion.precio_en_dolares, self._en_dolares(), self._cotizacion()
        )
        if precio_base <= 0:
            # Réplica del pedido del usuario (2026-08-15): un Artículo sin
            # Precio cargado no se puede facturar — avisar y devolver el
            # foco al primer campo de la Sección en vez de dejar cargado
            # un renglón a medias sin precio.
            QMessageBox.warning(
                self, "Detalle", f"El Artículo {cod1}-{cod2} no tiene Precio cargado — no se puede facturar."
            )
            self._limpiar_fila(fila)
            self._iniciar_edicion(fila, COL_SECCION)
            return
        estado.precio_base = precio_base

        self._actualizando = True
        try:
            self._set_texto(fila, COL_SECCION, cod1)
            for segmento, valor in zip(seccion.segmentos_codigo, valores_codigo):
                self._set_texto(fila, POSICION_A_COLUMNA[segmento.posicion], formatear_segmento(valor, segmento))
            self._set_texto(fila, COL_DESCRIPCION, descripcion)
        finally:
            self._actualizando = False

        self._actualizar_editabilidad_fila(fila)
        self._asegurar_fila_vacia_al_final()

        # Réplica de `AbmArt.frm Form_Activate` aplicada al Detalle
        # (feedback del usuario, 2026-08-15): si el Artículo tiene una
        # Nota guardada, se muestra sola apenas se lo carga en la grilla
        # — no hace falta ir a buscarla a mano.
        nota_articulo = self.repos.notaarticulo().by_articulo(cod1, cod2)
        if nota_articulo is not None:
            NotaArticuloDialog(self.repos, cod1, cod2, parent=self).exec()

        lote = DespachoSelectorDialog.abrir_si_corresponde(
            self.repos, cod1, cod2, descripcion_articulo=descripcion, parent=self
        )
        estado.nrodesp_elegido = (lote.NRODESP or "").strip() if lote is not None else None
        self._actualizando = True
        try:
            self._set_texto(fila, COL_LOTE, estado.nrodesp_elegido or "")
        finally:
            self._actualizando = False

        self._recalcular_fila(fila)
        self._al_cambiar()

        primera_cantidad = self._primera_columna_cantidad(seccion)
        if primera_cantidad is not None:
            self._iniciar_edicion(fila, primera_cantidad)

    @staticmethod
    def _primera_columna_cantidad(seccion: SeccionRenglon) -> Optional[int]:
        if not seccion.segmentos_cantidad:
            return None
        return POSICION_A_COLUMNA[seccion.segmentos_cantidad[0].posicion]

    # ------------------------------------------------------------------
    # Borrado de renglón (Supr)
    # ------------------------------------------------------------------
    def _eliminar_fila_actual(self) -> None:
        fila = self.currentRow()
        if fila < 0:
            return
        estado = self._estado.get(fila)
        if estado is None or estado.seccion is None:
            return  # fila vacía, nada que borrar

        respuesta = QMessageBox.question(
            self,
            "Detalle",
            "¿Eliminar este renglón?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.removeRow(fila)
        # Reindexar el diccionario de estados (las filas de abajo bajan un lugar).
        nuevo_estado = {}
        for indice, valor in self._estado.items():
            if indice == fila:
                continue
            nuevo_estado[indice - 1 if indice > fila else indice] = valor
        self._estado = nuevo_estado
        self._asegurar_fila_vacia_al_final()
        self._al_cambiar()

    # ------------------------------------------------------------------
    # Teclado: F2 (selector), Supr (borrar), Enter (avanza a la próxima
    # celda editable, o abre el selector si la Sección está en blanco).
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
        super().keyPressEvent(event)

    def closeEditor(self, editor, hint) -> None:  # noqa: N802 (Qt override)
        # `self._celda_editando` (seteado en `edit()`) es la única fuente
        # confiable de qué celda es esta — `self.currentRow()/
        # currentColumn()` puede ya apuntar a otro lado si Qt tarda en
        # cerrar el editor viejo y para entonces ya navegamos a otra
        # celda (editor "zombie" cerrándose tarde) — encontrado con
        # teclado real: el dato terminaba escribiéndose en el renglón
        # equivocado. Si ya no coincide con la celda que nosotros
        # abrimos, es justamente uno de esos cierres tardíos: se deja
        # que Qt confirme el dato (ya escribió lo que corresponda a ESA
        # celda) pero no se dispara navegación de nuevo por ella.
        celda = self._celda_editando
        self._celda_editando = None
        super().closeEditor(editor, hint)

        # Sólo se navega/abre el selector ante un cierre EXPLÍCITO del
        # operador (Enter/Tab, `SubmitModelCache` en todas las pruebas
        # reales con teclado) — un cierre "implícito" (`NoHint`: foco
        # perdido por otra causa, o Escape/`RevertModelCache`) no debe
        # disparar nada; encontrado con teclado real: un cierre tardío
        # sin motivo claro abría el selector de Artículo solo, con la
        # celda de Sección todavía en blanco.
        if celda is None or hint != QAbstractItemDelegate.EndEditHint.SubmitModelCache:
            return

        fila, columna = celda
        texto_antes_de_cerrar = editor.text() if hasattr(editor, "text") else None
        # Diferido, mismo motivo que `_on_item_changed`: abrir el
        # selector de Artículo o volver a llamar `editItem()` TODAVÍA
        # dentro de `closeEditor()` (re-entrante) causaba un
        # `RecursionError` real — encontrado probando con teclado de
        # verdad, no en un test superficial.
        QTimer.singleShot(
            0, lambda f=fila, c=columna, t=texto_antes_de_cerrar: self._despues_de_cerrar_editor(f, c, t)
        )

    def _despues_de_cerrar_editor(self, fila: int, columna: int, texto_antes_de_cerrar: Optional[str]) -> None:
        if columna == COL_SECCION and not (texto_antes_de_cerrar or "").strip():
            # Enter en la Sección vacía -> selector de Artículo (fiel al
            # comportamiento pedido: "al dar enter en blanco... aparece
            # la ventana de selección de artículos").
            self._abrir_selector_articulo(fila)
            return

        siguiente = self._siguiente_celda_editable(fila, columna)
        if siguiente is not None:
            self._iniciar_edicion(*siguiente)

    def _siguiente_celda_editable(self, fila: int, columna: int) -> Optional[tuple[int, int]]:
        total_columnas = self.columnCount()
        col = columna + 1
        vueltas_de_seguridad = 0
        while vueltas_de_seguridad < total_columnas * 2 + 4:
            if col >= total_columnas:
                fila += 1
                col = 0
                self._asegurar_fila_vacia_al_final()
                if fila >= self.rowCount():
                    return None
            item = self.item(fila, col)
            editable = bool(item is not None and (item.flags() & Qt.ItemFlag.ItemIsEditable))
            if editable:
                return fila, col
            col += 1
            vueltas_de_seguridad += 1
        return None

    # ------------------------------------------------------------------
    def foco_inicial(self) -> None:
        """Foco en la primera celda de la primera fila, editando —
        réplica de que el legacy deja a `FG1` con el cursor ya puesto en
        el primer renglón apenas se resuelve el Cliente."""
        if self.rowCount() == 0:
            self._agregar_fila_vacia()
        self._iniciar_edicion(0, COL_SECCION)

    def reiniciar(self) -> None:
        """Vacía la grilla por completo (equivalente a "Nueva") — deja
        una única fila en blanco lista para cargar."""
        self._actualizando = True
        try:
            self.setRowCount(0)
        finally:
            self._actualizando = False
        self._estado = {}
        self._agregar_fila_vacia()
