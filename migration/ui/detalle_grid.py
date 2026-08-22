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

from dataclasses import dataclass
from decimal import Decimal
from typing import Callable, Optional

from PyQt6.QtCore import QRegularExpression, Qt, QTimer
from PyQt6.QtGui import QRegularExpressionValidator
from PyQt6.QtWidgets import (
    QAbstractItemDelegate,
    QAbstractItemView,
    QLineEdit,
    QMessageBox,
    QStyledItemDelegate,
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
    calcular_importe,
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

# Columnas numéricas — alineadas a la derecha (feedback del usuario,
# 2026-08-19: "Ajustar a la derecha todos los campos numéricos") — Lote
# y Descripción son texto libre, Sección es un código alfanumérico
# ("**" para ítem libre), el resto siempre es una cantidad o un importe.
COLUMNAS_NUMERICAS = {COL_POS1, COL_POS2, COL_POS3, COL_POS4, COL_POS6, COL_PRECIO, COL_IMPORTE}

# Desvío máximo tolerado sin aviso entre el precio de lista y el que tipea
# el operador a mano en Precio Unit. — pedido del usuario (2026-08-21):
# "avise si el operador supera en un 20% en más o menos el valor
# original". Ver `DetalleGrid._on_precio_editado`.
DESVIO_PRECIO_MAXIMO = Decimal("0.20")


@dataclass
class _EstadoFila:
    seccion: Optional[SeccionRenglon] = None
    articulo: Optional[Articulo] = None
    nrodesp_elegido: Optional[str] = None
    # Precio de lista del Artículo (ya convertido por cotización) — se
    # fija UNA sola vez al resolver el Artículo y no se vuelve a tocar;
    # es la referencia del aviso de desvío ±20% (`_on_precio_editado`,
    # 2026-08-21). `None` para ítem libre (sin Artículo, nada contra qué
    # comparar). El precio que de verdad se usa para calcular el
    # Importe es el que esté tipeado en la celda en ese momento (ver
    # `_recalcular_fila`), no este campo.
    precio_lista: Optional[Decimal] = None


class _DelegadoNumerico(QStyledItemDelegate):
    """Filtra el editor de las columnas numéricas (`COLUMNAS_NUMERICAS`
    — cantidades y Precio Unit.) para que sólo dejen tipear dígitos y
    UN separador decimal, mismo criterio que `MontoLineEdit`/
    `EnteroLineEdit` (`widgets.py`) para el resto de la app — pedido del
    usuario (2026-08-15): "no debe permitir ingresar letras en las
    cantidades". El editor default de `QTableWidget` es un `QLineEdit`
    sin ninguna restricción; acá se le agrega un validador recién al
    crearlo, sin tocar el resto del comportamiento de edición de la
    grilla.

    Acepta "," O "." como separador decimal — bug real reportado por el
    usuario (2026-08-21, Precio Unit.): "no permite decimales en
    precio". El validador sólo dejaba pasar coma; en varios teclados
    numéricos (numpad) la tecla de decimal manda un PUNTO, no una coma
    (depende de la configuración regional de Windows de cada PC, no del
    layout que "debería" tener) — la tecla quedaba bloqueada del todo, el
    operador no podía cargar ningún decimal. `parse_decimal()`
    (`migration/decimals.py`) ya interpreta un punto como separador
    decimal cuando el texto no tiene ninguna coma, así que ensanchar acá
    el validador no rompe nada río abajo.

    **Bug grave real, encontrado con datos reales (2026-08-22, "no sale
    del precio, se queda allí")**: el regex de arriba sólo dejaba pasar
    UN separador en total — pero `format_decimal()` (lo que precarga
    Precio Unit. apenas se resuelve un Artículo) usa la convención es-AR
    completa, PUNTO de miles + COMA decimal, para cualquier importe
    ≥ 1000 (ej. "37.405,35"). Ese texto precargado tiene DOS separadores
    y el regex viejo lo marcaba como inválido — el validador, al
    rechazar el contenido ACTUAL del campo (no sólo bloquear la próxima
    tecla), hacía que Qt directamente IGNORARA el Enter sobre esa celda
    (ni committeaba el dato ni disparaba `closeEditor()` — confirmado
    instrumentando el código real: cero rastro de `closeEditor()` en
    ese Enter). Con un Artículo de precio chico (< 1000, un solo
    separador) nunca se notaba — de ahí que costó tanto reproducirlo.
    Nuevo regex: dígitos, con CUALQUIER cantidad de puntos intercalados
    (miles, sin exigir agruparlos de a 3 — no hace falta ser tan
    estricto, `parse_decimal()`/`format_decimal()` ya son la fuente de
    verdad del formato real) y a lo sumo UNA coma al final (decimales)."""

    _VALIDADOR = QRegularExpressionValidator(QRegularExpression(r"[0-9.]*,?[0-9]*"))

    def createEditor(self, parent, option, index):  # noqa: N802 (Qt override)
        editor = super().createEditor(parent, option, index)
        if index.column() in COLUMNAS_NUMERICAS and isinstance(editor, QLineEdit):
            editor.setValidator(self._VALIDADOR)
        return editor


class DetalleGrid(QTableWidget):
    def __init__(
        self,
        repos: RepositoryFactory,
        cliente_actual: Callable[[], Optional[Cliente]],
        en_dolares: Callable[[], bool],
        cotizacion: Callable[[], Decimal],
        hay_cotizacion_hoy: Callable[[], bool],
        al_cambiar: Callable[[], None],
        parent: QWidget | None = None,
    ):
        super().__init__(0, len(COLUMNAS), parent)
        self.repos = repos
        self._cliente_actual = cliente_actual
        self._en_dolares = en_dolares
        self._cotizacion = cotizacion
        self._hay_cotizacion_hoy = hay_cotizacion_hoy
        self._al_cambiar = al_cambiar  # callback: recalcular totales del Pie

        self._estado: dict[int, _EstadoFila] = {}
        self._actualizando = False
        # Celda que ESTA VENTANA abrió a editar más recientemente (no la
        # que Qt cree que está activa: `self.currentRow()/currentColumn()`
        # no es confiable dentro de `closeEditor()` cuando se navega de
        # celda en celda muy rápido — encontrado con teclado real, ver
        # `_iniciar_edicion()`/`closeEditor()`).
        self._celda_editando: Optional[tuple[int, int]] = None
        # `True` mientras `_aplicar_articulo()` está corriendo — incluye
        # el tiempo que pasa DENTRO de los diálogos modales que abre
        # (`NotaArticuloDialog`/`DespachoSelectorDialog`, ambos
        # `.exec()`). Ver `_despues_de_cerrar_editor()`.
        self._resolviendo_articulo = False

        self.setItemDelegate(_DelegadoNumerico(self))
        self.setHorizontalHeaderLabels(COLUMNAS)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.verticalHeader().setVisible(False)
        # Filas más cortas (feedback del usuario, 2026-08-19, apretado
        # más el 2026-08-21 — "deben ser 10 [renglones] aprox", con 4
        # visibles no alcanzaba): entran más renglones a la vista sin
        # achicar la letra.
        compactar_alto_filas(self, pct=82)
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
            self._aplicar_alineacion(item, columna)
            self.setItem(fila, columna, item)
        item.setText(texto)
        if columna == COL_DESCRIPCION:
            # Tooltip con la descripción COMPLETA (pedido del usuario,
            # 2026-08-22: "cuando se posiciona el cursor sobre la celda
            # de descripción, muestre un toolbox con la descripción
            # completa") — la columna es angosta y una descripción real
            # queda cortada visualmente; el tooltip nativo de Qt aparece
            # solo al posicionar el mouse encima, sin agrandar la celda.
            item.setToolTip(texto)

    @staticmethod
    def _aplicar_alineacion(item: QTableWidgetItem, columna: int) -> None:
        alineacion = Qt.AlignmentFlag.AlignVCenter | (
            Qt.AlignmentFlag.AlignRight if columna in COLUMNAS_NUMERICAS else Qt.AlignmentFlag.AlignLeft
        )
        item.setTextAlignment(alineacion)

    def _set_editable(self, fila: int, columna: int, editable: bool) -> None:
        # Bug real reportado por el usuario (2026-08-22, reproducido con
        # datos reales — Sección "A"/código 101 — con un script que
        # simula teclado real contra la grilla): `item.setFlags(...)`
        # más abajo dispara `itemChanged` TAMBIÉN (no sólo `setText()` lo
        # hace) — Qt no distingue "cambió el texto" de "cambió si se
        # puede editar" a los fines de esa señal. `_actualizar_
        # editabilidad_fila()` (la única llamadora real) corre SIEMPRE
        # DESPUÉS del bloque `self._actualizando = True/False` que
        # protege el `_set_texto()` de Precio en `_aplicar_articulo` —
        # este `setFlags()` quedaba MÁS de ese bloque protegido, así que
        # su `itemChanged` SÍ se procesaba como si el operador hubiera
        # tipeado algo en Precio de verdad, disparando `_on_precio_
        # editado()` (con el aviso de "falta cargar X") antes de tiempo
        # — de ahí el error apenas se resolvía el Artículo por código,
        # bastante antes de llegar a Precio. Mismo criterio que ya usa
        # `_set_texto()`: todo cambio programático de la celda queda
        # bajo esta guarda.
        self._actualizando = True
        try:
            item = self.item(fila, columna)
            if item is None:
                item = QTableWidgetItem()
                self.setItem(fila, columna, item)
                ya_editable = None
            else:
                ya_editable = bool(item.flags() & Qt.ItemFlag.ItemIsEditable)
            if ya_editable == editable:
                # Bug real encontrado probando con teclado de verdad:
                # `item.setFlags(...)` sobre la celda que se está
                # editando EN ESE MOMENTO le resetea el texto al editor
                # activo (Qt re-sincroniza el editor abierto cuando el
                # modelo "cambia", aunque el flag resultante sea
                # idéntico al que ya tenía) — típicamente pasa con la
                # propia celda Sección, que dispara `_actualizar_
                # editabilidad_fila` para sí misma apenas se resuelve.
                # Evitar el `setFlags()` redundante lo soluciona.
                return
            flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
            if editable:
                flags |= Qt.ItemFlag.ItemIsEditable
            item.setFlags(flags)
        finally:
            self._actualizando = False

    def _iniciar_edicion(self, fila: int, columna: int) -> None:
        """Punto de conveniencia para pasar a editar una celda desde
        código propio (navegación tras Enter, F2, foco inicial). El
        registro real de `self._celda_editando` pasa por `edit()` (ver
        más abajo), así que esto también cubre ediciones que arranca el
        propio operador (clic, flechas + tipear) sin pasar por acá.

        **Bug real reportado por el usuario (2026-08-22), reproducido
        con datos reales — Sección "A", código con lotes de Despacho**:
        "cargo A, 100, enter, aparece selección de despacho, enter,
        cantidad 5, enter, enter (en el precio) y no pasa nada". Un
        script con `QTest` simulando el teclado real lo confirmó: apenas
        se cierra `DespachoSelectorDialog` (`.accept()`/`.reject()`, la
        forma real en que se cierra, no sólo `.close()`), el foco de
        teclado se pierde DEL TODO — ni la grilla ni el editor nuevo que
        `editItem()` acababa de abrir lo tienen (`QApplication.
        focusWidget()` daba `None` incluso DESPUÉS de este método, con
        `self.state()` igual reportando `EditingState`). Sin foco real,
        las teclas que el operador sigue tipeando (cantidad, luego
        Enter en Precio) no le llegan a nadie — de ahí el "no pasa
        nada". `self.setFocus()` ANTES de abrir el editor fuerza a la
        grilla (y por lo tanto al editor que abre `editItem()` a
        continuación) a recuperar el foco real, sin depender de que Qt
        se lo devuelva solo después de un diálogo modal hijo. Hacía
        falta además reactivar la VENTANA (`activateWindow()`, no sólo
        `setFocus()` del widget) — confirmado con el mismo script: sin
        esto, `editItem()` seguía sin poder darle foco real a ningún
        editor nuevo."""
        ventana = self.window()
        if ventana is not None:
            ventana.activateWindow()
        self.setFocus()
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
                self._aplicar_alineacion(item, columna)
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
        # Descripción: en ítem libre la tipea el operador; si hay
        # Artículo real, se autogenera y no se toca a mano. Precio: SÍ
        # editable para cualquier Artículo ya resuelto (no sólo ítem
        # libre) — pedido del usuario (2026-08-21): "debe permitir que
        # el operador modifique el precio". Antes de resolver el
        # Artículo (o la Sección en sí) todavía no hay nada que
        # editar/mostrar ahí.
        self._set_editable(fila, COL_DESCRIPCION, es_libre)
        self._set_editable(fila, COL_PRECIO, es_libre or estado.articulo is not None)

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
        elif columna == COL_PRECIO:
            self._on_precio_editado(fila)
        elif columna in (COL_POS1, COL_POS2, COL_POS3, COL_POS4, COL_POS6):
            self._reconsiderar_codigo_y_recalcular(fila, columna)

        self._al_cambiar()

    def _on_seccion_editada(self, fila: int, texto: str) -> None:
        texto = texto.strip().upper()
        estado = self._estado.setdefault(fila, _EstadoFila())
        if not texto:
            estado.seccion = None
            estado.articulo = None
            # Bug real reportado por el usuario (2026-08-19): "si se
            # cambia un ítem en el facturador no resetea el Despacho/
            # Lote" — vaciar la Sección de un renglón ya cargado dejaba
            # el Lote elegido para el Artículo ANTERIOR pegado en
            # pantalla (y en `estado`), como si siguiera vigente para lo
            # que se cargue después. Ampliado (2026-08-22, mismo tipo de
            # bug reportado de nuevo): no sólo el Lote quedaba pegado —
            # los segmentos/Descripción/Precio/Importe de la Sección
            # ANTERIOR también, ver `_limpiar_datos_fila`.
            estado.nrodesp_elegido = None
            estado.precio_lista = None
            self._limpiar_datos_fila(fila)
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
        # Mismo motivo que el `if not texto:` de arriba: cambiar la
        # Sección de un renglón ya resuelto invalida TODO lo que se
        # había cargado para el Artículo/Sección anterior (segmentos,
        # Descripción, Precio, Importe, Lote) — se limpia acá mismo,
        # antes de que se complete el código nuevo (que recién va a
        # volver a pedir Lote si corresponde, vía `_aplicar_articulo`).
        estado.nrodesp_elegido = None
        estado.precio_lista = None
        self._limpiar_datos_fila(fila)
        self._actualizando = True
        try:
            self._set_texto(fila, COL_SECCION, seccion.cod_seccion)
            # Descripción de la Sección como placeholder TEMPORAL (pedido
            # del usuario, 2026-08-22: "luego de cargar la sección, en la
            # celda descripción, coloque temporalmente la descripción de
            # la sección. Cuando complete los datos, coloque la del
            # artículo") — el operador todavía no terminó de tipear el
            # código (o la Sección resuelve directo a un Artículo, ver
            # más abajo), pero ya tiene algo real para leer en vez de la
            # celda en blanco mientras carga. `_aplicar_articulo_impl`
            # la pisa con la descripción real (del Artículo, o armada con
            # los segmentos) apenas resuelve — ver `_set_texto(fila,
            # COL_DESCRIPCION, descripcion)` ahí.
            if seccion.descripcion:
                self._set_texto(fila, COL_DESCRIPCION, seccion.descripcion)
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

    def _limpiar_datos_fila(self, fila: int) -> None:
        """Vacía todas las celdas de DATOS de la fila (todo menos
        Sección, que ya trae el valor nuevo puesto por el llamador) —
        usado cuando la Sección de un renglón YA CARGADO cambia (a otra
        Sección o a blanco): los segmentos/Descripción/Precio/Importe/
        Lote quedaban pegados de la Sección/Artículo ANTERIOR, ya
        inválidos para lo que se tipee a continuación (bug real
        reportado por el usuario, 2026-08-22 — "cuando se cambia la
        sección debe limpiar el renglón")."""
        self._actualizando = True
        try:
            for columna in range(1, len(COLUMNAS)):  # todo menos COL_SECCION (0)
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

    def _reconsiderar_codigo_y_recalcular(self, fila: int, columna_editada: int) -> None:
        """Si se completaron todos los segmentos de código a mano (sin
        pasar por F2), intenta resolver el Artículo — mismo criterio que
        el botón "Buscar por Código" que tenía `RenglonDetalleDialog`,
        ahora disparado solo al completar la última celda de código."""
        estado = self._estado.get(fila)
        if estado is None or estado.seccion is None:
            return

        columnas_codigo = {POSICION_A_COLUMNA[s.posicion] for s in estado.seccion.segmentos_codigo}
        if columna_editada in columnas_codigo and (estado.articulo is not None or estado.nrodesp_elegido is not None):
            # Bug real reportado por el usuario (2026-08-19): "si se
            # cambia un ítem en el facturador no resetea el Despacho/
            # Lote". El renglón ya tenía un Artículo (y quizás un Lote)
            # resueltos, y el operador acaba de tocar a mano una celda
            # de CÓDIGO (no de cantidad) — el Artículo/Lote viejos ya no
            # son válidos para lo que se termine tipeando, así que se
            # limpian ACÁ MISMO, no recién cuando (si) se resuelve un
            # Artículo nuevo más abajo.
            estado.articulo = None
            estado.nrodesp_elegido = None
            self._actualizando = True
            try:
                self._set_texto(fila, COL_LOTE, "")
            finally:
                self._actualizando = False

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

    def _recalcular_fila(self, fila: int, *, avisar_si_falta: bool = False) -> None:
        """Recalcula sólo el Importe — el Precio Unitario NUNCA se toca
        acá (pedido del usuario, 2026-08-21: "deje el precio unitario
        del archivo sin modificar... muestre siempre el de la lista o
        el que cargó el operador"). Siempre lee el precio TAL COMO ESTÁ
        en la celda en este momento (de lista, recién resuelto el
        Artículo, o editado a mano) — ver `calcular_importe`.

        `avisar_si_falta`: si no se puede calcular el Importe porque
        falta cargar alguna cantidad, avisa CUÁL falta en vez de dejar
        el Importe en blanco sin explicación — bug real reportado por el
        usuario (2026-08-21): "cuando carga el precio unitario ya debe
        tener los datos para calcular el precio, sino avise (ej. falta
        la cant. en sección A)". Sólo `True` desde `_on_precio_editado`
        (el momento real en que el operador espera que YA se pueda
        calcular) — en el resto de los llamadores (mientras se completan
        los segmentos de código/cantidad en su orden normal) el Importe
        todavía no tiene por qué estar completo, avisar ahí sería
        prematuro y molesto."""
        estado = self._estado.get(fila)
        if estado is None or estado.seccion is None:
            return

        precio_actual = parse_decimal(self._texto(fila, COL_PRECIO))

        valores_cantidad = {}
        for segmento in estado.seccion.segmentos_cantidad:
            columna = POSICION_A_COLUMNA[segmento.posicion]
            valores_cantidad[segmento.alf_index] = parse_decimal(self._texto(fila, columna))

        try:
            importe = calcular_importe(precio_actual, valores_cantidad, estado.seccion)
        except (ValueError, SeccionSinUnidadFacturacionError):
            if avisar_si_falta:
                faltante = next(
                    (s for s in estado.seccion.segmentos_cantidad if valores_cantidad.get(s.alf_index, Decimal(0)) == 0),
                    None,
                )
                etiqueta = faltante.etiqueta if faltante is not None else "una cantidad"
                # Diferido (2026-08-22) — mismo motivo exacto que el
                # aviso de desvío ±20% de `_on_precio_editado`: un
                # `QMessageBox.warning()` SÍNCRONO acá pumpea el loop de
                # eventos mientras todavía hay un callback de navegación
                # encolado (`_despues_de_cerrar_editor`, disparado por
                # `closeEditor()`) — corría DENTRO del modal, con la
                # celda de Precio recién cerrada, y competía con el
                # cierre real del modal por el operador (reportado como
                # "el error de falta de dato da antes de que se cargue
                # el dato mismo"). Diferido, el callback de navegación ya
                # encolado corre primero, limpio, y el aviso queda para
                # el final.
                QTimer.singleShot(
                    0,
                    lambda seccion=estado.seccion.cod_seccion, etiqueta=etiqueta: QMessageBox.warning(
                        self,
                        "Detalle",
                        f'Falta cargar "{etiqueta}" en la Sección {seccion} para poder calcular el Importe.',
                    ),
                )
            self._actualizando = True
            try:
                self._set_texto(fila, COL_IMPORTE, "")
            finally:
                self._actualizando = False
            return

        self._actualizando = True
        try:
            self._set_texto(fila, COL_IMPORTE, format_decimal(importe))
        finally:
            self._actualizando = False

    def _on_precio_editado(self, fila: int) -> None:
        """Recalcula el Importe con el precio nuevo (avisando si falta
        alguna cantidad, ver `_recalcular_fila`) y dispara el aviso de
        desvío ±20% (si corresponde — hay precio de lista contra qué
        comparar, no es ítem libre). "Al final del ingreso" (no en cada
        tecla): se llama recién cuando la celda de Precio se CIERRA
        (Enter/Tab), mismo punto que dispara cualquier otro recálculo de
        la grilla.

        **Bug real encontrado con teclado real (2026-08-21, "a veces
        pasa al renglón siguiente y a veces no" tras cargar el Precio)**:
        el aviso ±20% es un `QMessageBox.warning()` MODAL, mostrado
        SÍNCRONO desde acá — pero acá mismo corre DIFERIDO por
        `QTimer.singleShot(0, ...)` (ver `_on_item_changed`), en la MISMA
        vuelta de evento en que Qt todavía tiene encolado (también
        diferido, encolado justo DESPUÉS) el callback que avanza al
        próximo renglón (`_despues_de_cerrar_editor`, disparado por
        `closeEditor()`). Un diálogo modal abierto acá pumpea el loop de
        eventos — y ESE segundo callback, todavía pendiente, terminaba
        corriendo DENTRO del loop anidado del propio modal (antes de que
        el operador llegara a cerrarlo), abriendo el editor del renglón
        siguiente TAPADO por el modal; al cerrar el modal, ese editor
        quedaba en un estado inconsistente (a veces "pegado" sin foco
        real) — de ahí el "a veces sí, a veces no": dependía pura y
        exclusivamente de si ESTE aviso llegaba a mostrarse (desvío
        >20%) o no. Difiriendo el propio `QMessageBox.warning()` acá
        (en vez de mostrarlo síncrono) se deja que el callback de avance
        ya encolado corra PRIMERO, limpio, sin ningún modal de por
        medio — el aviso de desvío queda para el final, ya con el foco
        en el renglón siguiente (no cambia su utilidad, es sólo
        informativo)."""
        self._recalcular_fila(fila, avisar_si_falta=True)

        estado = self._estado.get(fila)
        if estado is not None and estado.precio_lista:
            precio_actual = parse_decimal(self._texto(fila, COL_PRECIO))
            diferencia = precio_actual - estado.precio_lista
            if abs(diferencia) > estado.precio_lista * DESVIO_PRECIO_MAXIMO:
                direccion = "por ENCIMA" if diferencia > 0 else "por DEBAJO"
                QTimer.singleShot(
                    0,
                    lambda: QMessageBox.warning(
                        self,
                        "Detalle",
                        f"El precio tipeado (${format_decimal(precio_actual)}) está {direccion} en más "
                        f"de un 20% del precio de lista (${format_decimal(estado.precio_lista)}).",
                    ),
                )

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
        # Bug real reportado por el usuario (2026-08-22, reproducido con
        # datos reales — Sección "A", código 100): "cargo sección 'A',
        # pide nro. y luego de dar enter da el error [de falta de
        # dato]... debe esperar a llegar al precio". Encontrado con un
        # script que simula el teclado real (`QTest`) contra la grilla
        # con la base real: `_aplicar_articulo_impl()` (más abajo) abre
        # diálogos MODALES en el medio (`NotaArticuloDialog`/
        # `DespachoSelectorDialog`, ambos `.exec()`) ANTES de recalcular
        # y de mover el foco a la primera cantidad — un modal `.exec()`
        # bombea el loop de eventos, y ahí es donde el callback de
        # navegación YA ENCOLADO por el cierre de la celda de código
        # (`_despues_de_cerrar_editor`, ver su guarda `state() ==
        # EditingState`) se cuela y corre ANTES de que `_aplicar_
        # articulo_impl` haya llegado a mover el foco — en ESE momento
        # todavía no hay ningún editor abierto (`state()` da
        # `NoState`), la guarda no lo detecta, y navega por su cuenta a
        # la próxima celda de cantidad — abriéndole un editor propio que
        # compite con el que `_aplicar_articulo_impl` va a abrir
        # después, cuando el modal cierra y sigue con lo suyo. La
        # colisión terminaba comprometiendo Precio antes de tiempo, que
        # es de donde salía el aviso de "falta cargar" prematuro. Acá se
        # marca explícitamente el tramo completo (incluido el tiempo
        # DENTRO de esos modales) para que `_despues_de_cerrar_editor`
        # pueda detectarlo con certeza, sin depender de `state()` (que
        # esta corrida ya demostró que no alcanza).
        self._resolviendo_articulo = True
        try:
            self._aplicar_articulo_impl(fila, articulo)
        finally:
            self._resolviendo_articulo = False

    def _aplicar_articulo_impl(self, fila: int, articulo: Articulo) -> None:
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

        # Bloqueo real (2026-08-22, pedido del usuario: "si no hay dólar
        # cargado de hoy avanza igual, no debe hacerlo — antes pedía el
        # dólar y no dejaba seguir") — sólo cuando ESTE Artículo/Factura
        # realmente necesita convertir $/USD (`resolver_precio_articulo`
        # sólo usa la cotización cuando la Sección y la Factura difieren
        # en moneda); si los dos están en la misma moneda, la cotización
        # ni se usa, así que no hay motivo para bloquear.
        necesita_cotizacion = seccion.precio_en_dolares != self._en_dolares()
        if necesita_cotizacion and not self._hay_cotizacion_hoy():
            QMessageBox.warning(
                self,
                "Detalle",
                "No hay Cotización del Dólar cargada para HOY — cargala en "
                '"Cotización del Dólar" antes de facturar este Artículo.',
            )
            self._limpiar_fila(fila)
            self._iniciar_edicion(fila, COL_SECCION)
            return

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
        # Precio de lista — referencia fija para el aviso ±20%, y valor
        # inicial de la celda (pedido del usuario, 2026-08-21: "muestre
        # siempre el de la lista o el que cargó el operador"). Se
        # escribe UNA sola vez acá; de ahí en más `_recalcular_fila`
        # nunca la vuelve a tocar sola.
        estado.precio_lista = precio_base

        self._actualizando = True
        try:
            self._set_texto(fila, COL_SECCION, cod1)
            for segmento, valor in zip(seccion.segmentos_codigo, valores_codigo):
                self._set_texto(fila, POSICION_A_COLUMNA[segmento.posicion], formatear_segmento(valor, segmento))
            self._set_texto(fila, COL_DESCRIPCION, descripcion)
            self._set_texto(fila, COL_PRECIO, format_decimal(precio_base))
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
        if (
            event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
            and self.state() != QAbstractItemView.State.EditingState
        ):
            # Malentendido real, aclarado por el usuario (2026-08-22,
            # tercera vuelta): la versión anterior hacía que Enter, con
            # la celda sólo SELECCIONADA (sin editar), saltara derecho a
            # la próxima — "con enter está pasando las celdas y no debe
            # ser así tan fácil... con enter debe tomar el foco y luego
            # de hacer el ingreso pasar a la celda siguiente". Es decir:
            # el PRIMER Enter sobre una celda sólo seleccionada tiene
            # que ABRIRLA para editar (foco real adentro, como F2 o
            # tipear cualquier tecla) — el salto a la celda siguiente es
            # cosa del PRÓXIMO Enter, una vez que ya hay un dato
            # cargado de verdad, y eso ya lo resuelve solo el camino
            # normal (`closeEditor()` -> `_despues_de_cerrar_editor()`
            # -> `_navegar_o_bloquear()`, sin tocar acá).
            fila, columna = self.currentRow(), self.currentColumn()
            if fila >= 0 and columna >= 0:
                self._iniciar_edicion(fila, columna)
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

        if self.state() == QAbstractItemView.State.EditingState:
            # Bug real encontrado con teclado real (2026-08-21, "PL":
            # tipear MM, Enter, Telas, Enter): completar a mano el
            # ÚLTIMO segmento de código resuelve el Artículo y YA
            # navega sola a la primera cantidad
            # (`_procesar_dato_cambiado` -> `_reconsiderar_codigo_y_
            # recalcular` -> `_aplicar_articulo` -> `_iniciar_edicion`,
            # encolada con `QTimer.singleShot(0, ...)` desde
            # `_on_item_changed` — corre ANTES que este callback en la
            # misma vuelta del loop de eventos, porque se encola antes
            # que el propio `closeEditor()` encola a éste). Sin esta
            # guarda, ACÁ TAMBIÉN se abría un segundo editor sobre la
            # celda "siguiente" (calculada con el orden viejo de ESTA
            # celda) mientras el primero ya estaba abierto — Qt no lo
            # deja y tira "QAbstractItemView::edit: editing failed" por
            # consola, sin ningún aviso visible en la app. Si ya hay una
            # edición en curso (la abrió otro camino mientras este
            # callback esperaba su turno), no hay nada para navegar acá.
            return

        if self._resolviendo_articulo:
            # Ampliado (2026-08-22, mismo tipo de bug de arriba pero con
            # un modal de por medio — ver el docstring de
            # `_aplicar_articulo`): `state()` sólo detecta un editor YA
            # ABIERTO, pero mientras `_aplicar_articulo` está pausado
            # DENTRO de un `.exec()` (`NotaArticuloDialog`/
            # `DespachoSelectorDialog`) todavía no abrió ninguno —
            # `state()` da `NoState` justo en la ventana donde este
            # callback, si no fuera por este chequeo, se colaba a
            # navegar por su cuenta. `_aplicar_articulo` ya se ocupa de
            # su propia navegación (a la primera cantidad, o al volver a
            # Sección si el Artículo no tiene Precio) apenas termine.
            return

        self._navegar_o_bloquear(fila, columna)

    def _navegar_o_bloquear(self, fila: int, columna: int) -> None:
        """Decide y ejecuta el próximo paso de navegación desde `(fila,
        columna)` — código compartido entre `_despues_de_cerrar_editor`
        (cierre real de un editor) y `keyPressEvent` (Enter sobre una
        celda sólo SELECCIONADA, sin editar — ver ahí). Antes cada
        camino tenía su propia lógica de navegación por separado
        (`moveCursor`/flecha derecha para el segundo caso) — bug real
        reportado por el usuario (2026-08-22, tras el fix del código
        vacío/inexistente): "una vez que el ingreso está correcto,
        Enter en Precio debe saltar de renglón — con la flecha abajo
        cambia de línea, pero no debe ser así [tiene que ser con
        Enter]". La sospecha es que en la PC real, para cuando el
        operador aprieta Enter en Precio, la celda a veces ya no está
        realmente "editando" para Qt (mismo tipo de desincronización de
        foco que ya se encontró y corrigió para el diálogo de Despacho)
        — con la flecha-derecha como imitación, ESE Enter caía a
        Importe (columna vecina, nunca editable) en vez de saltar de
        renglón. Unificando las dos rutas en esta única función, el
        resultado de Enter es SIEMPRE el mismo salto "inteligente"
        (según `_orden_columnas_fila`, salta filas cuando corresponde),
        sin importar si Qt todavía considera la celda "en edición" o
        ya la dejó "sólo seleccionada"."""
        siguiente = self._siguiente_celda_editable(fila, columna)

        # Bug grave real reportado por el usuario (2026-08-22): "si no se
        # carga el código, sigue sin decir nada y ahí sí pasa a la línea
        # siguiente. Artículo no existe" — reproducido con datos reales
        # en 2 variantes: (a) código vacío + Enter, sin ningún aviso,
        # saltaba derecho a Cantidad; (b) código completo pero SIN
        # Artículo que lo tenga, `_reconsiderar_codigo_y_recalcular` SÍ
        # avisa ("No se encontró un Artículo...") pero de cualquier
        # forma dejaba seguir avanzando igual — el aviso no bloqueaba
        # nada. Las dos dejaban seguir cargando cantidad/precio sobre un
        # renglón SIN Artículo real (código armado a mano después, sin
        # nada real detrás). Se bloquea acá, en el único lugar que sabe
        # con certeza hacia dónde se estaba por navegar: si el próximo
        # destino natural queda FUERA de las columnas de código (es
        # decir, el operador ya terminó — o abandonó — la carga del
        # código) y todavía no hay Artículo resuelto, no se avanza.
        # Mientras el destino siga siendo OTRA columna de código (recién
        # completando el segmento 2 de 3, por ejemplo) se deja seguir
        # normal, ahí no hay nada que bloquear.
        estado = self._estado.get(fila)
        if estado is not None and estado.seccion is not None and estado.seccion.segmentos_codigo:
            columnas_codigo = {POSICION_A_COLUMNA[s.posicion] for s in estado.seccion.segmentos_codigo}
            if columna in columnas_codigo and estado.articulo is None:
                si_avanza_fuera_del_codigo = siguiente is None or siguiente[1] not in columnas_codigo
                if si_avanza_fuera_del_codigo:
                    codigo_completo = all(
                        self._texto(fila, POSICION_A_COLUMNA[s.posicion]) for s in estado.seccion.segmentos_codigo
                    )
                    if not codigo_completo:
                        # Único caso sin aviso previo — el código completo
                        # pero sin Artículo YA avisó "No se encontró..."
                        # en `_reconsiderar_codigo_y_recalcular`, no hace
                        # falta un segundo cartel para lo mismo.
                        QMessageBox.warning(
                            self, "Detalle", "Tenés que cargar un Artículo válido (código completo) antes de continuar."
                        )
                    self._iniciar_edicion(fila, columna)
                    return

        if siguiente is not None:
            self._iniciar_edicion(*siguiente)

    def _orden_columnas_fila(self, seccion: Optional[SeccionRenglon]) -> list[int]:
        """Orden real de CARGA (Enter) de una fila — a diferencia del
        orden VISUAL fijo de columnas (Nro/Pulg, Mtr/Kg, MM, Telas: los
        mismos 4 slots físicos para cualquier Sección, ver
        `POSICION_A_COLUMNA`/`Posi()` del legacy), el orden en que el
        operador tiene que completarlas depende de la Sección real
        (`ALF1..ALF7` de `Fctabla1`, ver `resolver_seccion_renglon`).

        **Bug real reportado por el usuario (2026-08-21, probado con la
        Sección "PL")**: tipear a mano debía saltar "a MM y luego a
        Telas y luego a Pulgadas y luego a Metros" — que es exactamente
        `segmentos_codigo` (arma el código del Artículo, ALF1-3) seguido
        de `segmentos_cantidad` (multiplican el precio, ALF4-7), NO el
        orden fijo de columnas de la grilla (que para PL daría Pulgadas
        antes que MM). Con F2 el código ya saltaba bien directo a la
        primera cantidad (`_aplicar_articulo` -> `_iniciar_edicion`),
        pero el PRÓXIMO Enter volvía a caer en MM/Telas — codificadas
        como "editables" (correcto, siguen usándose) pero sin ningún
        lugar que supiera que ya estaban completas y venían ANTES en el
        orden real, no después."""
        orden = [COL_SECCION]
        if seccion is not None:
            orden += [POSICION_A_COLUMNA[s.posicion] for s in seccion.segmentos_codigo]
            orden.append(COL_DESCRIPCION)
            orden += [POSICION_A_COLUMNA[s.posicion] for s in seccion.segmentos_cantidad]
            orden.append(COL_PRECIO)
        else:
            orden += [COL_POS1, COL_POS2, COL_POS3, COL_POS4, COL_DESCRIPCION, COL_POS6, COL_PRECIO]
        orden += [COL_IMPORTE, COL_LOTE]
        return orden

    def _siguiente_celda_editable(self, fila: int, columna: int) -> Optional[tuple[int, int]]:
        vueltas_de_seguridad = 0
        while vueltas_de_seguridad < self.rowCount() * 2 + 4:
            seccion = self._estado.get(fila, _EstadoFila()).seccion
            orden = self._orden_columnas_fila(seccion)
            desde = orden.index(columna) + 1 if columna in orden else 0
            for col in orden[desde:]:
                item = self.item(fila, col)
                editable = bool(item is not None and (item.flags() & Qt.ItemFlag.ItemIsEditable))
                if editable:
                    return fila, col

            fila += 1
            columna = -1  # no pertenece a ningún orden -> `desde = 0`, arranca desde el principio de la fila nueva
            self._asegurar_fila_vacia_al_final()
            if fila >= self.rowCount():
                return None
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
