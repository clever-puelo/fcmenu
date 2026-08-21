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

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtWidgets import QApplication, QComboBox, QLineEdit

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
        QTimer.singleShot(0, self.selectAll)

    QLineEdit.focusInEvent = _focus_in_con_seleccion
    QLineEdit._fcmenu_seleccion_instalada = True


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
    QPushButton {{
        background-color: {v.MEDIO};
        color: {v.BLANCO};
        border: none;
        border-radius: 5px;
        padding: 6px 16px;
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {v.MEDIO_OSCURO};
    }}
    QPushButton:pressed {{
        background-color: {v.OSCURO};
    }}
    QPushButton:disabled {{
        background-color: {v.CLARO};
        color: {v.TEXTO_SUAVE};
    }}
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
