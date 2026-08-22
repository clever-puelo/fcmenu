"""Tema visual compartido de toda la app — pedido del usuario
(2026-08-16, última fase del proyecto, "conformar a gente reticente al
cambio"): **lindo, llamativo, funcional y muy parecido al anterior**.

- **Colores**: gama de verde claro/pastel para fondos, verde oscuro
  para títulos — paleta única (`Verde`) para que cualquier pantalla
  nueva la reuse sin inventar tonos propios.
- **Íconos**: simples, prácticos, evidentes — ver `iconos.py`.
- **"Parecido al anterior"**: el `MainMenuWindow` (`main_menu_window.py`)
  replica la estructura real de `FCMENU.frm` (grupos ABM's/Consultas/
  Ingreso/Listados/Herramientas + "Arreglos" como sección aparte, sin
  ocultarla como el legacy) — la paleta y esta hoja de estilo son la
  "piel" que se aplica sobre esa estructura y sobre cada pantalla ya
  migrada.

Aplicar UNA vez por proceso, en cada `main_xxx.py` (`aplicar_tema(app)`
justo después de crear el `QApplication`) — no hace falta tocar cada
ventana individualmente, es una hoja de estilo Qt (QSS) global.
"""

from __future__ import annotations

from PyQt6.QtCore import QCoreApplication, Qt, QTimer
from PyQt6.QtGui import QColor, QIcon, QPixmap
from PyQt6.QtWidgets import QApplication, QComboBox, QGraphicsDropShadowEffect, QLineEdit, QPushButton

# El navegador embebido de la ficha de Cliente ("Google Earth",
# `navegador_dialog.py`, `QWebEngineView`) necesita que TODAS las
# ventanas de la app compartan el mismo contexto OpenGL — Qt exige que
# esto se fije ANTES de crear la instancia de `QApplication`, no alcanza
# con importar `QtWebEngineWidgets` recién al abrir esa ventana puntual
# (que es justamente lo que hace `navegador_dialog.py` a propósito, para
# no cargar el paquete pesado en cada arranque de la app) — bug real
# reportado por el usuario (2026-08-21): abrir "Google Earth" tiraba
# `ImportError: QtWebEngineWidgets must be imported or
# Qt.AA_ShareOpenGLContexts must be set before a QCoreApplication
# instance is created`. Se fija acá, a nivel de MÓDULO (no dentro de
# `aplicar_tema()`, que ya corre demasiado tarde: todo `main_xxx.py`
# la llama recién DESPUÉS de construir el `QApplication`) — el `from
# .theme import aplicar_tema` de cada entry point ya se ejecuta antes de
# esa línea, así que alcanza con que este archivo se importe.
QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts, True)

# Cantidad de renglones visibles en el desplegable de CUALQUIER combo de
# la app antes de scrollear — pedido del usuario (2026-08-17): "el combo
# de selección de sección explota en toda la pantalla" (Artículos),
# "el combo vendedor también explota" (Factura). En vez de tocar cada
# combo uno por uno, se parchea `QComboBox.__init__` una sola vez acá
# (ver `aplicar_tema`) para que TODO combo nuevo de la app, sin
# excepción, salga ya limitado — "Registra que en todos los combos
# muestren 10 renglones (si hay espacio)... y se pueda scrollar hacia
# arriba y abajo" (el scroll lo da gratis Qt en cuanto hay más ítems que
# `maxVisibleItems`, no hace falta código extra para eso).
MAX_RENGLONES_VISIBLES_COMBO = 10


def _instalar_limite_combos() -> None:
    if getattr(QComboBox, "_fcmenu_limite_instalado", False):
        return  # ya parcheado (ej. tests que llaman aplicar_tema más de una vez)
    init_original = QComboBox.__init__

    def _init_con_limite(self, *args, **kwargs):  # noqa: ANN001, ANN002, ANN003
        init_original(self, *args, **kwargs)
        self.setMaxVisibleItems(MAX_RENGLONES_VISIBLES_COMBO)

    QComboBox.__init__ = _init_con_limite
    QComboBox._fcmenu_limite_instalado = True


def _instalar_seleccion_total_al_foco() -> None:
    """Parchea `QLineEdit.focusInEvent` UNA vez por proceso para que
    TODO campo de texto de la app, sin excepción, seleccione su
    contenido completo al tomar el foco (Tab o clic de mouse) — pedido
    del usuario (2026-08-21): "eso debería hacer siempre y en todos los
    campos de todo el sistema", detectado de nuevo con el Teléfono
    (contiene espacios — "011 4652-1040" — así que no tomaba el camino
    de `_NumericLineEditBase`, sólo numérico, y quedaba afuera).

    Antes esto sólo vivía en `_NumericLineEditBase.focusInEvent()`
    (campos de importe/entero/alfanumérico) — mismo criterio ya
    confirmado por el usuario el 2026-08-15 (ver docstring de esa
    clase), ahora extendido a CUALQUIER `QLineEdit` de la app (nombre,
    dirección, teléfono, email, búsquedas de texto libre, etc.), en vez
    de tener que acordarse de heredar de una clase especial en cada
    pantalla nueva. Las subclases que ya hacían `selectAll()` a mano
    (`_NumericLineEditBase`, que sí llama a `super().focusInEvent()`)
    quedan con un `selectAll()` de más, inofensivo — no hay ningún
    `focusInEvent` en el resto del código que NO llame a `super()`."""
    if getattr(QLineEdit, "_fcmenu_seleccion_instalada", False):
        return  # ya parcheado (ej. tests que llaman aplicar_tema más de una vez)
    focus_in_original = QLineEdit.focusInEvent

    def _focus_in_con_seleccion(self, event):  # noqa: ANN001
        focus_in_original(self, event)
        # Diferido a la próxima vuelta del loop de eventos — mismo
        # motivo ya documentado en `_NumericLineEditBase`: si el foco
        # se ganó con un clic de mouse, el propio `mousePressEvent()`
        # de Qt todavía está pendiente y deshace un `selectAll()` hecho
        # acá mismo, síncrono.
        #
        # Bug real encontrado con datos reales (2026-08-21, ABM de
        # Clientes): "Razón Social" aparecía resaltada apenas se elegía
        # un cliente en la búsqueda, aunque la ficha quedara en modo
        # Sólo-Ver (deshabilitada). Causa: la búsqueda es un diálogo
        # MODAL abierto ENCIMA de la ficha — el foco real de Qt vuelve a
        # `txt_nomb` (que lo tenía antes de abrirse, en modo Alta) recién
        # al cerrarse, disparando ESTE `focusInEvent` de nuevo, justo
        # cuando `_cargar_cliente()` está por deshabilitar el campo
        # (`_set_modo("ver")`). El `singleShot` quedaba pendiente y
        # ejecutaba `selectAll()` un instante después, ya con el nombre
        # real cargado — se ve el campo "seleccionado" pese a estar
        # deshabilitado. Se re-chequea `isEnabled()` recién al disparar
        # el timer (no alcanza con chequearlo acá arriba, síncrono:
        # todavía no se deshabilitó en este mismo instante).
        QTimer.singleShot(0, lambda: self.selectAll() if self.isEnabled() else None)

    QLineEdit.focusInEvent = _focus_in_con_seleccion
    QLineEdit._fcmenu_seleccion_instalada = True


def _instalar_sombra_botones() -> None:
    """Parchea `QPushButton.__init__` UNA vez por proceso para que TODO
    botón de la app, sin excepción, tenga una sombra real proyectada
    debajo — pedido del usuario (2026-08-21): "que los botones tengan
    apariencia tridimensional, con iluminación y sombras". El relieve/
    iluminación (gradiente + bisel) lo resuelve el QSS de
    `estilo_boton_3d()` — pero Qt Style Sheets NO soporta `box-shadow`
    (no es CSS real, un subconjunto), así que la sombra de verdad se
    agrega acá aparte, con `QGraphicsDropShadowEffect` (efecto gráfico
    real, no dibujable por QSS). Mismo patrón ya usado para
    `QComboBox.__init__` (`_instalar_limite_combos`) — parchear el
    constructor una sola vez alcanza para CUALQUIER botón nuevo de
    cualquier pantalla, sin tener que acordarse de aplicarlo a mano en
    cada una.

    Sólo `QPushButton` (no `QToolButton` acá) — `QToolButton` lo usa Qt
    también para controles internos chicos (ej. la "x" de cerrar de una
    pestaña) donde una sombra dramática se ve mal; `BotonIconoTexto`
    (`widgets.py`, el único `QToolButton` "de verdad" que arma la app)
    se aplica la sombra a mano en su propio `__init__`, con el mismo
    `aplicar_sombra_boton()` de acá abajo."""
    if getattr(QPushButton, "_fcmenu_sombra_instalada", False):
        return  # ya parcheado (ej. tests que llaman aplicar_tema más de una vez)
    init_original = QPushButton.__init__

    def _init_con_sombra(self, *args, **kwargs):  # noqa: ANN001, ANN002, ANN003
        init_original(self, *args, **kwargs)
        aplicar_sombra_boton(self)

    QPushButton.__init__ = _init_con_sombra
    QPushButton._fcmenu_sombra_instalada = True


def aplicar_sombra_boton(boton) -> None:  # noqa: ANN001 (cualquier QWidget-botón)
    """Sombra real proyectada (`QGraphicsDropShadowEffect`) — suave y
    chica a propósito (blur 8px, offset 2px, gris al 45% de opacidad):
    tiene que leerse como "el botón flota un poco", no como una sombra
    dura de post-it. Reusado por `_instalar_sombra_botones()` (todo
    `QPushButton`) y por `BotonIconoTexto.__init__` (`widgets.py`)."""
    sombra = QGraphicsDropShadowEffect(boton)
    sombra.setBlurRadius(8)
    sombra.setOffset(0, 2)
    sombra.setColor(QColor(0, 0, 0, 115))
    boton.setGraphicsEffect(sombra)


def _tono(color_hex: str, factor: float) -> str:
    """Aclara (`factor` > 1) u oscurece (`factor` < 1) `color_hex`
    multiplicando cada canal RGB — usado por `estilo_boton_3d()` para
    derivar el gradiente/bisel de CUALQUIER color base sin tener que
    calibrar una paleta de 3-4 tonos a mano por cada botón/pantalla."""
    color = QColor(color_hex)
    canal = lambda v: min(255, max(0, round(v * factor)))  # noqa: E731
    return QColor(canal(color.red()), canal(color.green()), canal(color.blue())).name()


def estilo_boton_3d(
    selector: str,
    color_base: str,
    color_texto: str = "",
    *,
    radio: int = 5,
    padding_v: int = 6,
    padding_h: int = 16,
    font_size: str = "",
) -> str:
    """Bloque QSS con apariencia tridimensional para `selector` (ej.
    `"QPushButton"`, `"QToolButton"`) — pedido del usuario (2026-08-21):
    "que los botones tengan apariencia tridimensional, con iluminación y
    sombras. Simular movimiento a pasar y al oprimir".

    - **Iluminación/relieve**: gradiente vertical (más claro arriba, más
      oscuro abajo) + bisel de 4 bordes (más claro arriba/izquierda, más
      oscuro abajo/derecha en reposo — el bisel clásico de un botón
      "elevado"; se invierte al presionar, simulando que se hunde).
    - **Al pasar el mouse**: gradiente más claro/brillante (`:hover`).
    - **Al oprimir**: gradiente y bisel invertidos + el contenido se
      corre 1px hacia abajo (`padding-top`/`padding-bottom` asimétricos
      en `:pressed`) — el botón "se hunde" de verdad, no sólo cambia de
      color.
    - **Sombra**: Qt Style Sheets NO soporta `box-shadow` — la sombra
      real la agrega aparte `aplicar_sombra_boton()` (`QGraphicsDropShadowEffect`),
      no este QSS.

    Los tonos se derivan automáticamente de `color_base` (`_tono()`) —
    sirve para cualquier color de cualquier pantalla, sin necesidad de
    definir una paleta de relieve aparte para cada una."""
    claro = _tono(color_base, 1.22)
    muy_claro = _tono(color_base, 1.45)
    oscuro = _tono(color_base, 0.8)
    muy_oscuro = _tono(color_base, 0.62)
    hover_claro = _tono(color_base, 1.55)
    hover_oscuro = _tono(color_base, 1.12)
    disabled = _tono(color_base, 1.35)
    color_prop = f"color: {color_texto};" if color_texto else ""
    fuente_prop = f"font-size: {font_size};" if font_size else ""

    return f"""
    {selector} {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {claro}, stop:1 {oscuro});
        {color_prop}
        {fuente_prop}
        border-top: 1px solid {muy_claro};
        border-left: 1px solid {muy_claro};
        border-bottom: 2px solid {muy_oscuro};
        border-right: 2px solid {muy_oscuro};
        border-radius: {radio}px;
        padding: {padding_v}px {padding_h}px;
        font-weight: bold;
    }}
    {selector}:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {hover_claro}, stop:1 {hover_oscuro});
    }}
    {selector}:pressed {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {oscuro}, stop:1 {claro});
        border-top: 2px solid {muy_oscuro};
        border-left: 2px solid {muy_oscuro};
        border-bottom: 1px solid {muy_claro};
        border-right: 1px solid {muy_claro};
        padding-top: {padding_v + 1}px;
        padding-bottom: {max(padding_v - 1, 0)}px;
    }}
    {selector}:disabled {{
        background: {disabled};
        border-color: {disabled};
    }}
    """


class Verde:
    """Escala de verde única de la app — de más oscuro (texto/títulos)
    a más claro (fondos). No usar códigos de color sueltos en ninguna
    pantalla nueva: importar de acá."""

    OSCURO = "#1B4332"       # Títulos, texto de énfasis, cabeceras
    MEDIO_OSCURO = "#2D6A4F"  # Hover/pressed de botones primarios
    MEDIO = "#40916C"         # Botones primarios, acentos activos
    MEDIO_CLARO = "#74C69D"   # Bordes de foco, chips, badges
    CLARO = "#B7E4C7"         # Bordes suaves, separadores
    PASTEL = "#D8F3DC"        # Tarjetas, filas alternadas, hover suave
    MUY_CLARO = "#F1FAF3"     # Fondo general de ventana
    BLANCO = "#FFFFFF"        # Fondo de inputs/tablas

    TEXTO = "#1B2B22"          # Texto normal (buen contraste sobre MUY_CLARO/BLANCO)
    TEXTO_SUAVE = "#4A5D52"    # Texto secundario/ayuda
    ALERTA = "#B00020"         # Errores/advertencias (sin cambios, ya usado en varias pantallas)
    DORADO = "#D4A017"         # Resaltados puntuales (ej. "tiene nota") — ya usado como #ffd54f en el Facturador, se mantiene aparte de la escala de verde a propósito
    AMARILLO = "#FFD600"       # Cartel de estado AFIP en la cabecera de MainMenuWindow — más saturado que DORADO a propósito, tiene que "gritar" sobre el verde oscuro


def hoja_de_estilo() -> str:
    """QSS global. Cubre los controles estándar de PyQt6 que ya usa
    toda la app (`QMainWindow`, `QDialog`, `QPushButton`, `QGroupBox`,
    `QTableWidget`, `QLineEdit`/`QComboBox`/`QDateEdit`/`QSpinBox`,
    `QTabWidget`) — no requiere tocar el código de cada pantalla."""
    v = Verde
    return f"""
    QMainWindow, QDialog {{
        background-color: {v.MUY_CLARO};
        color: {v.TEXTO};
    }}
    QWidget {{
        color: {v.TEXTO};
        font-family: "Segoe UI", "Verdana", sans-serif;
        font-size: 10pt;
    }}
    QLabel {{
        background: transparent;
    }}
    QGroupBox {{
        background-color: {v.BLANCO};
        border: 1px solid {v.CLARO};
        border-radius: 6px;
        margin-top: 12px;
        padding: 10px 8px 8px 8px;
        font-weight: bold;
        color: {v.OSCURO};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 6px;
        color: {v.OSCURO};
    }}
    {estilo_boton_3d("QPushButton", v.MEDIO, v.BLANCO)}
    QPushButton:default {{
        border: 2px solid {v.OSCURO};
    }}
    QComboBox {{
        /* CLAVE para que `setMaxVisibleItems()` (ver `_instalar_limite_
        combos`) funcione de verdad: con una hoja de estilo QSS global
        aplicada (como esta), Qt cambia el popup del combo al modo
        "menú nativo" por defecto, que crece con la cantidad de ítems
        SIN respetar `maxVisibleItems` — es el bug real detrás de "los
        combos explotan en toda la pantalla" (feedback del usuario,
        2026-08-18, confirmado con un script de prueba: sin esta línea
        un combo de 60 ítems con maxVisibleItems=10 mostraba ~39 filas
        igual). `combobox-popup: 0` fuerza el popup tipo lista, que sí
        lo respeta. */
        combobox-popup: 0;
    }}
    QLineEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox, QTextEdit, QPlainTextEdit {{
        background-color: {v.BLANCO};
        border: 1px solid {v.MEDIO_CLARO};
        border-radius: 4px;
        padding: 3px 5px;
        selection-background-color: {v.MEDIO_CLARO};
        selection-color: {v.OSCURO};
    }}
    QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
        border: 1.5px solid {v.MEDIO};
    }}
    QLineEdit:disabled, QComboBox:disabled {{
        background-color: {v.MUY_CLARO};
        color: {v.TEXTO_SUAVE};
    }}
    QTableWidget, QTreeWidget {{
        background-color: {v.BLANCO};
        alternate-background-color: {v.PASTEL};
        gridline-color: {v.CLARO};
        border: 1px solid {v.CLARO};
        border-radius: 4px;
    }}
    QHeaderView::section {{
        background-color: {v.OSCURO};
        color: {v.BLANCO};
        padding: 5px;
        border: none;
        font-weight: bold;
    }}
    QTableWidget::item:selected, QTreeWidget::item:selected {{
        background-color: {v.MEDIO_CLARO};
        color: {v.OSCURO};
    }}
    QTabWidget::pane {{
        border: 1px solid {v.CLARO};
        border-radius: 4px;
        background: {v.BLANCO};
    }}
    QTabBar::tab {{
        background: {v.PASTEL};
        color: {v.OSCURO};
        padding: 6px 14px;
        border: 1px solid {v.CLARO};
        border-bottom: none;
        border-top-left-radius: 4px;
        border-top-right-radius: 4px;
    }}
    QTabBar::tab:selected {{
        background: {v.BLANCO};
        font-weight: bold;
    }}
    QMenuBar {{
        background-color: {v.OSCURO};
        color: {v.BLANCO};
    }}
    QMenuBar::item:selected {{
        background-color: {v.MEDIO};
    }}
    QMenu {{
        background-color: {v.BLANCO};
        border: 1px solid {v.CLARO};
    }}
    QMenu::item:selected {{
        background-color: {v.PASTEL};
        color: {v.OSCURO};
    }}
    QStatusBar {{
        background-color: {v.PASTEL};
        color: {v.OSCURO};
    }}
    QToolTip {{
        background-color: {v.OSCURO};
        color: {v.BLANCO};
        border: none;
        padding: 4px 6px;
    }}
    QScrollBar:vertical {{
        background: {v.MUY_CLARO};
        width: 12px;
    }}
    QScrollBar::handle:vertical {{
        background: {v.MEDIO_CLARO};
        border-radius: 5px;
        min-height: 20px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {v.MEDIO};
    }}
    """


def aplicar_tema(app: QApplication) -> None:
    """Aplica la hoja de estilo global — llamar una sola vez, apenas
    creado el `QApplication`, en cada entry point `main_xxx.py`."""
    app.setStyleSheet(hoja_de_estilo())
    app.setStyle("Fusion")
    _instalar_limite_combos()
    _instalar_seleccion_total_al_foco()
    _instalar_sombra_botones()


def icono_app() -> QIcon:
    """Ícono de la aplicación (barra de tareas/título de ventana) —
    usa `assets/Icon-Alestel.png` (provisto por el usuario, 2026-08-20).
    Si por algún motivo faltara, cae a `ICON1.ICO` del proyecto legacy
    (comportamiento previo) y, en última instancia, a un `QIcon` vacío."""
    from pathlib import Path

    raiz = Path(__file__).resolve().parent.parent.parent
    ruta_png = raiz / "assets" / "Icon-Alestel.png"
    if ruta_png.exists():
        return QIcon(str(ruta_png))
    ruta_ico = raiz / "ICON1.ICO"
    return QIcon(str(ruta_ico)) if ruta_ico.exists() else QIcon()


def logo_empresa(alto: int = 52) -> QPixmap:
    """Logo real (`assets/Logo-Alestel.png`, provisto por el usuario
    2026-08-16) escalado a una altura fija, para la cabecera de
    `MainMenuWindow`. `QPixmap` vacío si el archivo no está (mismo
    criterio de `icono_app`, no revienta si falta el asset)."""
    from pathlib import Path

    ruta = Path(__file__).resolve().parent.parent.parent / "assets" / "Logo-Alestel.png"
    if not ruta.exists():
        return QPixmap()
    return QPixmap(str(ruta)).scaledToHeight(alto, Qt.TransformationMode.SmoothTransformation)
