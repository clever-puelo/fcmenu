"""Widgets y helpers de navegación compartidos por los formularios PyQt6.

Replican las convenciones que el usuario marcó como obligatorias para
todo el sistema, salvo aclaración específica:

1. Todo texto libre se escribe en MAYÚSCULAS automáticamente (CLAUDE.md,
   sección "Navegación Teclado", ej. Abmclte.frm:1443-1451).
2. Enter avanza el foco como si fuera Tab (y Shift+Enter retrocede) —
   ídem, Abmclte.frm:1437-1438. Ver `EnterAsTabFilter`.
3. Toda grilla de búsqueda permite ordenar de menor a mayor y viceversa
   al clickear el título de cada columna (confirmado 2026-08-16) — ver
   `TablaBusqueda` más abajo.
4. Todo campo numérico selecciona su contenido completo al tomar el
   foco, con Tab O con clic de mouse (confirmado 2026-08-15, mismo
   patrón que `Text1_GotFocus` en todo el legacy: `SelStart=0`/
   `SelLength=Len(Text1.Text)`) — así tipear reemplaza el valor entero
   en vez de pegarse al que ya estaba. Ver `_NumericLineEditBase`.
5. Toda ventana emergente (diálogo) arranca con el foco puesto en su
   primer campo, salvo aclaración específica (confirmado 2026-08-15) —
   se difiere con `QTimer.singleShot(0, ...)` en el `__init__` de cada
   diálogo (armar el foco antes de que la ventana esté mostrada no
   siempre "pega" — mismo criterio ya usado para `showPopup()` en
   `ArticuloBusquedaWindow`).
6. Todo campo numérico, sea importe o no (ej. Nº de Cheque, código de
   Banco), sólo acepta caracteres numéricos al tipear (confirmado
   2026-08-15) — `MontoLineEdit` además permite una coma decimal;
   `EnteroLineEdit` es la versión sin coma, para enteros puros.
7. En un diálogo con un botón de acción principal (ej. "Agregar"/
   "Guardar"), Enter en el ÚLTIMO campo lleva el foco a ese botón en
   vez de dar la vuelta al primer campo (confirmado 2026-08-15) — pasar
   `boton_final` a `EnterAsTabFilter`.
8. Toda grilla donde se CARGAN datos (no sólo se buscan) tiene que ser
   interactiva estilo Excel — se tipea/navega directo en la celda, sin
   diálogos modales "Agregar/Editar" de por medio (confirmado
   2026-08-16, primero para el Detalle del Facturador — *"Es muy
   estático... En el legacy es totalmente distinto... como en Excel"* —
   y reafirmado como norma general para toda la app al migrar `Stock.frm`:
   *"La grilla debe ser interactiva... Toma esto como norma general para
   todas las grillas del sistema"*). Ver `DetalleGrid`
   (`detalle_grid.py`) y `StockRenglonesGrid` (`stock_renglones_grid.py`)
   como referencia del patrón (F2/Enter abren el selector de Artículo,
   Supr borra el renglón con confirmación, navegación automática a la
   próxima celda al confirmar un valor) — replicar ese mismo patrón en
   cualquier grilla de carga nueva, no el de diálogo modal + botón
   Agregar/Quitar (ese patrón sigue vigente sólo para grillas de
   BÚSQUEDA/selección, que no cargan datos).
9. Todo campo de fecha (`QDateEdit`) tiene un botón "Hoy" al lado que lo
   resetea a la fecha actual (ya usado en `PagoDialog` para Fecha
   Emisión/Vencimiento; reafirmado como norma general 2026-08-16 al
   pedirlo también en Consulta de Cheques y Ventas de Artículos por
   Cliente) — ver `crear_boton_hoy()` más abajo.
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from typing import Any, Optional

from PyQt6.QtCore import QDate, QEvent, QObject, Qt, QTimer
from PyQt6.QtGui import QKeyEvent
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from .decimals import format_decimal, parse_decimal


def crear_boton_hoy(campo: QDateEdit) -> QPushButton:
    """Botón "Hoy" que resetea `campo` a la fecha actual — convención de
    sistema #9 (ver docstring del módulo)."""
    boton = QPushButton("Hoy")
    boton.setMaximumWidth(50)
    boton.clicked.connect(lambda: campo.setDate(QDate.currentDate()))
    return boton


class UpperCaseLineEdit(QLineEdit):
    """QLineEdit que fuerza mayúsculas al tipear.

    Réplica de Abmclte.frm Form_KeyPress (líneas 1443-1451): las minúsculas
    se convierten a mayúsculas, el resto de las teclas pasa sin cambios.
    """

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.textEdited.connect(self._forzar_mayusculas)

    def _forzar_mayusculas(self, texto: str) -> None:
        mayus = texto.upper()
        if mayus != texto:
            pos = self.cursorPosition()
            self.blockSignals(True)
            self.setText(mayus)
            self.blockSignals(False)
            self.setCursorPosition(pos)


class LowerCaseLineEdit(QLineEdit):
    """QLineEdit que fuerza minúsculas al tipear — excepción a propósito
    a la regla general de "todo en mayúsculas" (feedback del usuario,
    2026-08-18): un email en mayúsculas queda raro y no es la convención
    real de un email, se usa sólo acá (`ClienteDetalleDialog.txt_email`)."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.textEdited.connect(self._forzar_minusculas)

    def _forzar_minusculas(self, texto: str) -> None:
        minus = texto.lower()
        if minus != texto:
            pos = self.cursorPosition()
            self.blockSignals(True)
            self.setText(minus)
            self.blockSignals(False)
            self.setCursorPosition(pos)


class _NumericLineEditBase(QLineEdit):
    """Base compartida por `MontoLineEdit`/`EnteroLineEdit` — convención
    de sistema para TODO campo numérico (confirmada por el usuario,
    2026-08-15): selecciona su contenido completo al tomar el foco (con
    Tab O con clic de mouse) y sólo deja pasar caracteres numéricos
    (además de las teclas de edición/navegación). Ver puntos 4 y 6 del
    docstring de módulo.

    Subclases implementan `_caracter_permitido(texto_de_una_tecla)`."""

    # Teclas de edición/navegación que SIEMPRE tienen que pasar, sin
    # importar qué traiga `event.text()` — bug real encontrado con
    # teclado de verdad (2026-08-15): `Key_Backspace` en Qt no siempre
    # da `event.text() == ""` como el resto de las teclas de control
    # (en Windows llega como `"\x08"`, el carácter de control ASCII
    # backspace) — un `if texto and not texto.isdigit(): return` de
    # buena fe terminaba bloqueando Backspace/Delete, y el campo sólo
    # se podía "vaciar" letra por letra desde el final, nunca corregir
    # en el medio. Se filtra por `event.key()` explícito, no por texto.
    _TECLAS_SIEMPRE_PERMITIDAS = {
        Qt.Key.Key_Backspace, Qt.Key.Key_Delete, Qt.Key.Key_Left, Qt.Key.Key_Right,
        Qt.Key.Key_Home, Qt.Key.Key_End, Qt.Key.Key_Tab, Qt.Key.Key_Backtab,
        Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Escape,
    }

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignRight)

    def focusInEvent(self, event) -> None:  # noqa: N802 (Qt override)
        # Selecciona todo el contenido al tomar el foco (feedback del
        # usuario, 2026-08-15): sin esto tipear se pegaba al valor
        # existente en vez de reemplazarlo — mismo patrón que
        # `Text1_GotFocus` en todo el legacy.
        #
        # Diferido con `QTimer.singleShot(0, ...)`: si el foco se ganó
        # con un CLIC DE MOUSE (feedback del usuario, 2026-08-15), Qt
        # todavía tiene pendiente su propio `mousePressEvent()` — que
        # posiciona el cursor en el punto del clic y DESHACE cualquier
        # `selectAll()` hecho acá mismo, dentro de `focusInEvent`. Correr
        # `selectAll()` recién en la próxima vuelta del loop de eventos
        # (después de que Qt ya procesó el clic) lo deja seleccionado de
        # verdad sin importar si el foco vino de Tab o de un clic.
        super().focusInEvent(event)
        QTimer.singleShot(0, self.selectAll)

    def keyPressEvent(self, event: QKeyEvent) -> None:  # noqa: N802 (Qt override)
        if event.key() in self._TECLAS_SIEMPRE_PERMITIDAS:
            super().keyPressEvent(event)
            return
        if event.modifiers() & ~Qt.KeyboardModifier.ShiftModifier:
            # Ctrl+C/V/X/A y combinaciones con Alt/Meta pasan tal cual
            # (copiar/pegar/seleccionar todo) — sólo se filtra el
            # carácter tipeado "a mano", sin modificador.
            super().keyPressEvent(event)
            return
        texto = event.text()
        if texto and not self._caracter_permitido(texto):
            return
        super().keyPressEvent(event)

    def _caracter_permitido(self, texto: str) -> bool:
        raise NotImplementedError


class MontoLineEdit(_NumericLineEditBase):
    """`QLineEdit` para importes — pedido del usuario (2026-08-15): "en la
    carga de importes solo permita caracteres numericos y coma decimal"
    y "la carga debe ser editada y ajustada y quedar editada".

    Mientras se tipea, sólo deja pasar dígitos y UNA coma decimal
    (mismo criterio que los `KeyPress` de importe del legacy, ej.
    `DetPago.frm text4_KeyPress`: filtra todo lo que no sea `0-9`/`,`).
    Al perder el foco, reformatea con `format_decimal()` — el valor
    queda siempre "ajustado" (ej. "1000" -> "1.000,00"), nunca a medio
    tipear."""

    def __init__(self, valor_inicial: Decimal | str = "0", parent: QWidget | None = None):
        super().__init__(parent)
        inicial = valor_inicial if isinstance(valor_inicial, Decimal) else parse_decimal(str(valor_inicial))
        self.set_decimal(inicial)
        self.editingFinished.connect(self._reformatear)

    def _caracter_permitido(self, texto: str) -> bool:
        if texto == ",":
            return "," not in self.text()  # ya hay una coma, no se permite una segunda
        return texto.isdigit()

    def _reformatear(self) -> None:
        self.setText(format_decimal(parse_decimal(self.text())))

    def decimal(self) -> Decimal:
        return parse_decimal(self.text())

    def set_decimal(self, valor: Decimal) -> None:
        self.setText(format_decimal(Decimal(valor)))


class EnteroLineEdit(_NumericLineEditBase):
    """`QLineEdit` para números enteros SIN parte decimal (ej. Nº de
    Cheque, código de Banco) — pedido del usuario (2026-08-15): "en los
    campos numericos, aunque no sean importes... solo numeros". Mismas
    reglas de foco/navegación que `MontoLineEdit` (ver clase base), sin
    coma ni reformateo de moneda — un entero no necesita "ajustarse"."""

    def __init__(self, valor_inicial: int | str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self.setText(str(valor_inicial) if valor_inicial != "" else "")

    def _caracter_permitido(self, texto: str) -> bool:
        return texto.isdigit()

    def entero(self) -> Optional[int]:
        texto = self.text().strip()
        return int(texto) if texto.isdigit() else None


class AlfanumericoLineEdit(_NumericLineEditBase):
    """`QLineEdit` para códigos alfanuméricos (ej. Nº de Despacho —
    "25001IC04066997X", letras Y números, sin separadores) — pedido del
    usuario (2026-08-16, buscador de `DespachosConsultaWindow`): "El nº
    de despacho tiene letras, asi que la busqueda debe permitir el
    ingreso de letras, no espacios ni caracteres distintos a letras o
    numeros". Mismas reglas de foco/selección que el resto de los
    campos numéricos (ver clase base) — acá filtra por `isalnum()` en
    vez de sólo dígitos, y fuerza mayúsculas (convención #1 del sistema,
    mismo criterio que `UpperCaseLineEdit`), ya que los códigos reales
    de despacho vienen en mayúsculas."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.textEdited.connect(self._forzar_mayusculas)

    def _caracter_permitido(self, texto: str) -> bool:
        return texto.isalnum()

    def _forzar_mayusculas(self, texto: str) -> None:
        mayus = texto.upper()
        if mayus != texto:
            pos = self.cursorPosition()
            self.blockSignals(True)
            self.setText(mayus)
            self.blockSignals(False)
            self.setCursorPosition(pos)


class EnterAsTabFilter(QObject):
    """Event filter: Enter/Flecha-Abajo/Derecha = Tab, Flecha-Arriba/Izq = Shift+Tab.

    Réplica de Abmclte.frm Form_KeyDown (líneas 1437-1438):
        If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
        If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"

    **Bug real encontrado con teclado de verdad (2026-08-15)**: los 5
    formularios que ya "usaban" esto (`ClienteDetalleDialog`,
    `ArticuloDetalleDialog`, `CotizacionWindow`, `ParametrosWindow`,
    `TablasWindow`) construían el filtro con `EnterAsTabFilter(self)`
    pero NINGUNO llamaba `self.installEventFilter(...)` después — Enter
    nunca avanzó el foco en NINGUNO de esos formularios, a pesar de
    parecer "instalado". Y aunque lo hubieran llamado, instalar el
    filtro sobre la VENTANA no alcanza: los eventos de teclado le llegan
    al widget con foco (el `QLineEdit`/`QComboBox` puntual), no a la
    ventana contenedora — confirmado con una prueba real de
    `QTest.keyClick()` antes de este fix.

    El constructor ahora se autoinstala en cada campo "de entrada"
    (`QLineEdit`/`QComboBox`/`QSpinBox`/`QDateEdit`) descendiente de
    `ventana` en el momento de construirse — por eso hay que crearlo
    DESPUÉS de armar todos los campos (mismo orden que ya usaban los 5
    formularios existentes: `self._construir_ui()` primero,
    `EnterAsTabFilter(self)` al final), sin tener que tocar esos call
    sites uno por uno.

    **`QDateEdit` sólo con Enter, NUNCA con flechas** (feedback del
    usuario, 2026-08-15: "al darle enter a la fecha pase de campo"):
    Izquierda/Derecha son la forma nativa de moverse entre los
    segmentos día/mes/año DENTRO del propio campo de fecha — tratarlas
    como Tab/Shift+Tab (como sí corresponde para `QLineEdit`/
    `QComboBox`/`QSpinBox`, réplica fiel del legacy) rompería esa
    navegación interna.

    **Bug real encontrado con teclado de verdad (2026-08-15, banco en
    blanco dejó de abrir la búsqueda)**: `widget.focusNextChild()` NO
    dispara `editingFinished` de un `QLineEdit` — confirmado con
    `QTest.keyClick()` real (el signal sólo se emite por el manejo
    NATIVO de Enter de Qt, que acá se intercepta y nunca llega a
    correr, o por un focus-out "real" que tampoco es lo que hace
    `focusNextChild()`). Por eso `_avanzar()` emite `editingFinished` a
    mano ANTES de mover el foco — así cualquier handler que dependa de
    "se terminó de tipear este campo" (ej. `PagoDialog._on_banco_
    confirmado`) sigue disparando igual que con Tab real.

    **Segundo bug real encontrado el mismo día ("al darle enter a la
    fecha pase de campo" no andaba)**: `QDateEdit.focusNextChild()`
    devuelve `False` y NO mueve el foco — confirmado con un repro
    mínimo de dos `QDateEdit` en un layout (`QTest` real). No se pudo
    determinar por qué Qt no lo resuelve para este control puntual, así
    que en vez de confiar en el mecanismo nativo de foco de Qt (que ya
    demostró ser poco confiable dos veces), el filtro arma y mantiene
    su PROPIA lista ordenada de campos (`self._campos`, capturada una
    vez al construirse, en el mismo orden en que se agregaron a la
    ventana) y navega manualmente sobre esa lista — funciona igual para
    los cuatro tipos de control sin depender de la implementación
    interna de cada uno.

    **`boton_final`** (feedback del usuario, 2026-08-15: "el ultimo
    campo, observ., al dar enter pase el control al boton agregar.
    Ahora vuelve a cheque"): sin esto, Enter en el ÚLTIMO campo de la
    lista le da la vuelta al primero (`% total`) — correcto para
    navegar entre campos, pero no cuando lo que sigue lógicamente es el
    botón de acción principal del diálogo (ej. "Agregar"/"Guardar").
    Pasando `boton_final`, Enter en el último campo lleva el foco ahí
    en vez de repetir el ciclo — y una vez que el botón tiene el foco,
    Enter/Espacio lo activan solos (comportamiento nativo de
    `QPushButton`, no hace falta manejarlo acá).
    """

    _AVANZA_ENTER = {Qt.Key.Key_Return, Qt.Key.Key_Enter}
    _AVANZA_FLECHAS = {Qt.Key.Key_Down, Qt.Key.Key_Right}
    _RETROCEDE = {Qt.Key.Key_Up, Qt.Key.Key_Left}
    _CAMPOS = (QLineEdit, QComboBox, QSpinBox, QDateEdit)
    # QDateEdit usa Izquierda/Derecha para moverse entre los segmentos
    # día/mes/año DENTRO del propio campo — sólo Enter lo hace avanzar
    # de campo, nunca las flechas (ver docstring de la clase).
    _CAMPOS_CON_FLECHAS = (QLineEdit, QComboBox, QSpinBox)

    def __init__(self, ventana: QWidget, boton_final: QWidget | None = None):
        super().__init__(ventana)
        self._boton_final = boton_final
        # `QDateEdit` (y otros `QAbstractSpinBox`) tienen su PROPIO
        # `QLineEdit` interno como editor de texto — `findChildren()` lo
        # trae también, "colado" en la lista justo después del
        # `QDateEdit` que lo contiene (bug real encontrado con teclado
        # de verdad, 2026-08-15: ese `QLineEdit` fantasma rompía el
        # orden y el Enter en una fecha terminaba "avanzando" al editor
        # interno de la fecha siguiente en vez de a un campo real). Se
        # descarta cualquier candidato que sea descendiente de OTRO
        # candidato ya encontrado — sólo quedan los controles "de
        # primer nivel".
        todos = ventana.findChildren(self._CAMPOS)
        self._campos = [w for w in todos if not any(o is not w and o.isAncestorOf(w) for o in todos)]
        for widget in self._campos:
            widget.installEventFilter(self)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802 (Qt override)
        if event.type() != QEvent.Type.KeyPress or not isinstance(watched, self._CAMPOS):
            return False
        key = event.key()  # type: ignore[attr-defined]
        if key in self._AVANZA_ENTER:
            self._avanzar(watched)
            return True
        if isinstance(watched, self._CAMPOS_CON_FLECHAS):
            if key in self._AVANZA_FLECHAS:
                self._avanzar(watched)
                return True
            if key in self._RETROCEDE:
                self._mover_foco(watched, -1)
                return True
        return False

    def _avanzar(self, widget: QWidget) -> None:
        if isinstance(widget, QLineEdit):
            widget.editingFinished.emit()
        if widget.hasFocus():
            # El handler de `editingFinished` puede haber movido el foco
            # por su cuenta (ej. saltar a otro campo puntual, o abrir un
            # diálogo que termina cambiando el foco) — no pisarlo si ya
            # no hace falta.
            self._mover_foco(widget, +1)

    def _mover_foco(self, widget: QWidget, direccion: int) -> None:
        if widget not in self._campos:
            return
        total = len(self._campos)
        indice = self._campos.index(widget)

        if (
            direccion > 0
            and indice == total - 1
            and self._boton_final is not None
            and self._boton_final.isVisible()
            and self._boton_final.isEnabled()
        ):
            self._boton_final.setFocus()
            return

        for _ in range(total):
            indice = (indice + direccion) % total
            candidato = self._campos[indice]
            if candidato.isVisible() and candidato.isEnabled():
                candidato.setFocus()
                return


class _ItemOrdenable(QTableWidgetItem):
    """`QTableWidgetItem` que ordena numéricamente cuando el texto lo
    permite, texto si no.

    Corrige un bug real de Qt: `QTableWidgetItem` por defecto compara
    SIEMPRE como string, así que con `setSortingEnabled(True)` una
    columna numérica ordena mal en cuanto hay números de distinta
    cantidad de dígitos (ej. "1000" queda antes que "999", porque '1' <
    '9' al comparar el primer carácter) — encontrado probando el
    ordenamiento real de la grilla de Clientes.

    **Segundo bug real encontrado probando "Ver Saldos" de la Consulta
    de Cta.Cte. con datos reales (2026-08-16)**: la implementación
    original asumía formato US ("$ 1,234.56", coma de miles) y sólo
    borraba la coma — cualquier importe formateado con `format_decimal()`
    (convención es-AR real de todo el sistema, "1.234,56": PUNTO de
    miles, COMA decimal) quedaba con dos puntos ("1.234.56"),
    `Decimal()` tiraba `InvalidOperation` y la columna caía silenciosa a
    comparar como texto — exactamente el síntoma reportado ("el importe
    ordena como alfabético"). Nunca se había manifestado antes porque
    ninguna otra grilla ordenable (`TablaBusqueda`) mostraba un importe
    con `format_decimal()` en una columna — sólo códigos/enteros sin
    coma decimal, que no disparan el bug. Ahora reusa `parse_decimal()`
    (la MISMA función que arma/relee cualquier importe en toda la app,
    convención es-AR) en vez de una lógica de parseo propia.

    **Tercer bug real encontrado probando "Consulta de Despachos" con
    datos reales (2026-08-16, "no ordena por nº de despacho ni por
    fecha")**: el chequeo "¿tiene algún dígito?" es DEMASIADO amplio —
    un Nº de Despacho alfanumérico ("25001IC04066997X") o una fecha
    "dd/mm/yyyy" TIENEN dígitos pero no son números puros; `parse_decimal()`
    sobre ese texto no matchea el patrón de número real, tira
    `InvalidOperation` y devuelve `Decimal("0")` — así que TODAS las
    filas de esas columnas comparaban como "0 == 0" (ningún cambio de
    orden visible, exactamente el síntoma reportado). Corregido con un
    regex que exige que el texto COMPLETO (no sólo una parte) sea un
    número. Para columnas que necesitan un orden real distinto del texto
    (ej. fechas: "dd/mm/yyyy" NO ordena cronológicamente como string —
    "20/01/2019" > "05/06/2020" al comparar como texto) usar
    `clave_orden` (ver `__init__`) con el valor real (`date`, `Decimal`,
    etc.) en vez de confiar en el texto mostrado."""

    def __init__(self, texto: str, clave_orden: Any = None):
        super().__init__(texto)
        self._clave_orden = clave_orden

    def __lt__(self, otro: object) -> bool:
        if (
            self._clave_orden is not None
            and isinstance(otro, _ItemOrdenable)
            and otro._clave_orden is not None
        ):
            try:
                return self._clave_orden < otro._clave_orden
            except TypeError:
                pass  # tipos no comparables entre sí -> cae a los métodos de abajo
        propio = self._valor_numerico()
        ajeno = otro._valor_numerico() if isinstance(otro, _ItemOrdenable) else None
        if propio is not None and ajeno is not None:
            return propio < ajeno
        return self.text() < otro.text()

    # Número entero o decimal en convención es-AR completo — con o sin
    # separador de miles ("1234", "1.234,56", "-50,00") — nunca una parte
    # del texto ("25001IC04066997X" o "10/04/2025" no matchean enteros).
    _PATRON_NUMERO = re.compile(r"^-?\d+(\.\d{3})*(,\d+)?$")

    def _valor_numerico(self) -> Optional[Decimal]:
        texto = self.text().strip().replace("$", "").strip()
        if not self._PATRON_NUMERO.match(texto):
            return None
        return parse_decimal(texto)


class TablaBusqueda(QTableWidget):
    """Grilla de resultados de búsqueda, de sólo lectura, con columnas
    ordenables al clic (convención de sistema confirmada por el usuario
    el 2026-08-16 — reemplaza los `QListWidget` de una sola línea que se
    habían usado antes en Cliente/Artículo/Tablas)."""

    def __init__(
        self,
        columnas: list[str],
        parent: QWidget | None = None,
        *,
        columnas_derecha: tuple[int, ...] = (),
        columnas_centro: tuple[int, ...] = (),
    ):
        super().__init__(0, len(columnas), parent)
        self.setHorizontalHeaderLabels(columnas)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.verticalHeader().setVisible(False)
        # Filas más apretadas (feedback del usuario, 2026-08-16: achicar
        # el alto de fila sin tocar el tamaño de letra) — el alto default
        # de Fusion (~30px) le sobra a una sola línea de texto de 10pt.
        self.verticalHeader().setDefaultSectionSize(24)
        self.setSortingEnabled(True)
        # Prolijidad pedida por el usuario (2026-08-16): columnas numéricas
        # (código, comprobante, importe) a la derecha; imputaciones al
        # centro — parametrizable acá para que cualquier `TablaBusqueda`
        # nueva lo pueda pedir sin repetir el loop de alineación.
        self._columnas_derecha = columnas_derecha
        self._columnas_centro = columnas_centro

    def cargar_filas(
        self,
        filas: list[tuple[list[str], Any]],
        claves_orden: dict[int, list[Any]] | None = None,
    ) -> None:
        """`filas`: lista de (textos_por_columna, objeto_asociado). El
        objeto viaja en `Qt.ItemDataRole.UserRole` de la primera celda.

        `claves_orden`: opcional, `{columna: [valor_real_por_fila, ...]}`
        para columnas donde el texto mostrado NO ordena bien por sí solo
        (ej. fechas "dd/mm/yyyy" — no son cronológicas como string, ver
        `_ItemOrdenable`) — se pasa el valor real (`date`, `Decimal`,
        etc.) en el mismo orden que `filas`."""
        ordenamiento_previo = self.isSortingEnabled()
        self.setSortingEnabled(False)
        self.setRowCount(0)
        for fila, (textos, objeto) in enumerate(filas):
            fila_grilla = self.rowCount()
            self.insertRow(fila_grilla)
            for col, texto in enumerate(textos):
                clave = claves_orden.get(col, [None] * len(filas))[fila] if claves_orden else None
                item = _ItemOrdenable(texto, clave_orden=clave)
                if col == 0:
                    item.setData(Qt.ItemDataRole.UserRole, objeto)
                if col in self._columnas_derecha:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                elif col in self._columnas_centro:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.setItem(fila_grilla, col, item)
        self.setSortingEnabled(ordenamiento_previo)
        self.resizeColumnsToContents()

    def objeto_en_fila(self, fila: int) -> Any:
        item = self.item(fila, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def objeto_seleccionado(self) -> Any:
        fila = self.currentRow()
        return self.objeto_en_fila(fila) if fila >= 0 else None
