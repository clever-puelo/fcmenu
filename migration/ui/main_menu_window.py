"""`MainMenuWindow` — ventana principal de la app, reemplaza al
`MDIForm FCMENU` del legacy (`FCMENU.frm`, 1472 líneas).

Rediseño pedido por el usuario (2026-08-16, sesión dedicada a la
estética, "conformar el gusto de personas reticentes al cambio")
emulando la estructura real de un MDI clásico:

- **Cabecera** (verde oscuro, fija arriba): logo (`assets/Logo-Alestel.png`)
  + "FCMENU" + subtítulo + versión a la izquierda; cartel de estado de
  facturación electrónica (`migration.afip.etiqueta_entorno_afip`) al
  centro-derecha; fecha/hora en vivo a la derecha (`QTimer` de 1s).
- **Panel lateral** (15-20% del ancho): 5 botones de sección — mismos
  nombres/orden que el menú "Tareas" real del legacy (`FCMENU.frm`
  líneas 503-651: A-B-M's/Consultas/Ingreso/Listados) más "Varios"
  (agrupa Configuración + Arreglos + lo que no encaja en las otras 4).
- **Panel principal** (resto del ancho): título = sección activa, barra
  fina de tareas específicas de esa sección, y debajo el espacio de
  trabajo real — un `QMdiArea` en modo pestañas donde se embeben las
  pantallas ya migradas (antes abrían como ventanas flotantes
  independientes; ver `_mostrar`).
- **Pie** (10% del alto, todo el ancho): panel de mensajes del sistema a
  la izquierda, estado de conexión a Internet en vivo
  (`estado_conexion.MonitorInternet`) al centro-derecha, botón Salir a
  la derecha.

Los diálogos modales puntuales (`.exec()`, ej. `NotaClienteDialog`,
`DtosClienteDialog`) NO se tocan — siguen flotando por encima de todo
como siempre, sólo las ~20 pantallas principales (antes top-level, ahora
subventanas del `QMdiArea`) cambiaron de contenedor. Al estar embebidas
como hijas de Qt dentro del `QMdiArea`, ya no hace falta una lista
Python aparte para que no las recolecte el garbage collector (al
contrario que en la versión anterior de este archivo)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import QSize, Qt, QTimer, QUrl
from PyQt6.QtGui import QDesktopServices, QFont, QFontMetrics
from PyQt6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMdiArea,
    QMdiSubWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from migration.afip import etiqueta_entorno_afip
from migration.db import get_session
from migration.fechas import formatear_fecha_corta
from migration.repository import RepositoryFactory

from .estado_conexion import MonitorInternet
from .iconos import icono
from .theme import Verde, icono_app, logo_empresa

VERSION_APP = "2.0.0"
INTERVALO_RELOJ_MS = 1_000
INTERVALO_CHEQUEO_INTERNET_MS = 15_000
ALTO_BARRA_TAREAS = 64  # botones de 52px + margen — deja lugar para 2 líneas de texto
# Alto cuando la barra necesita 2 líneas de BOTONES, no de texto (ver
# `_reconstruir_barra_tareas`) — 2 botones de 52px apilados + el mismo
# margen/espaciado que ya usaba la versión de 1 línea.
ALTO_BARRA_TAREAS_2_FILAS = 2 * (ALTO_BARRA_TAREAS - 12) + 10
# Si el ancho que le tocaría a cada botón en una sola fila cae por
# debajo de esto, la barra pasa a 2 líneas (feedback del usuario,
# 2026-08-18, segunda ronda: "en listados, como hay muchos botones, la
# ventana principal se estira y sale del monitor" — con muchos botones en
# una sola fila, el ancho mínimo de cada uno no entraba en pantalla y
# forzaba a la ventana principal a crecer más allá del monitor).
# Reemplaza al umbral viejo por CANTIDAD fija de botones (`MAX_BOTONES_
# UNA_FILA = 9`) — bug real, 2026-08-20: Consultas tiene justo 9 botones
# (el límite viejo la dejaba en 1 fila) y, al agrandar los íconos del
# menú (ver Fase A), el ancho mínimo real de cada botón pasó a superar
# lo que le tocaba por `stretch`, y Qt agrandaba la ventana SOLA para
# cumplir el nuevo mínimo del layout. Calcular por ancho real, no por
# cantidad, evita que esto se repita cada vez que cambie el tamaño de
# ícono/fuente.
ANCHO_MINIMO_BOTON_TAREA = 150
# Ícono real de los botones de la barra de tareas (`_crear_boton_tarea`)
# — reusado acá para que `_envolver_en_dos_lineas()` descuente el ancho
# correcto (bug real, 2026-08-20: quedó un "34" hardcodeado, calibrado
# para el ícono viejo de 22px, que ya no alcanzaba tras agrandarlo).
ANCHO_ICONO_TAREA = 40

# Posicionamiento de subventanas del MDI (feedback del usuario,
# 2026-08-18, tercera ronda: "quedaría más agradable si las ventanas/
# módulos aparecen en el centro del panel principal y no en el extremo
# superior izquierdo" — antes `QMdiArea.cascadeSubWindows()`). Listados
# pidió un margen puntual en mm en vez de centrado simple ("a 30mm por
# lado lateral y a 50mm de arriba") — aproximado a 96 DPI, suficiente
# para un margen visual dentro de una subventana, no para impresión.
PX_POR_MM = 96 / 25.4
MARGEN_LISTADOS_LATERAL_MM = 30
MARGEN_LISTADOS_SUPERIOR_MM = 50


def _mm_a_px(mm: float) -> int:
    return round(mm * PX_POR_MM)


# % del panel MDI real (NO de la pantalla completa) para las pantallas
# de trabajo más grandes — feedback del usuario, 2026-08-19: "ahora son
# gigantes". `redimensionar_pct_pantalla()` (convención #11,
# `widgets.py`) calcula contra la pantalla física completa — tiene
# sentido para uso standalone (`main_facturador.py` y afines, sin MDI
# de por medio), pero la ventana principal NO arranca maximizada
# (pedido explícito del usuario, no se toca) así que el panel MDI real
# siempre es más chico que la pantalla completa: un 90%/95% de la
# pantalla física terminaba siendo MÁS GRANDE que el panel MDI donde
# tiene que entrar, y se veía "gigante"/desbordado. Acá se recalcula
# centralizado contra el tamaño REAL del panel MDI en el momento de
# mostrarla (mismo criterio que ya usaba Listados con su margen en mm
# — por eso fue la única que el usuario no reportó como desajustada).
PCT_MDI_POR_CLASE: dict[str, tuple[float, float]] = {
    "TablasWindow": (80, 75),
    # Alto subido a 95% (feedback del usuario, 2026-08-19: "usar todo el
    # alto del panel central") — antes 70%, quedaba con espacio vacío
    # abajo mientras el panel de Artículos se veía chico.
    "ModPreciosWindow": (75, 95),
    # Alto subido a 95% (feedback del usuario, 2026-08-19: "usar todo el
    # alto del panel central") — antes 80%.
    "StockMovimientoWindow": (80, 95),
    "CobranzasZonaWindow": (90, 85),
    "NotaCreditoConceptoWindow": (80, 85),
    "NotaCreditoMercaderiaWindow": (90, 88),
    "TotalesDiariosWindow": (85, 88),
    "ChequesConsultaWindow": (75, 88),
    "VentasSeccionWindow": (80, 88),
    "VentasArticuloWindow": (85, 78),
    "StockConsultaWindow": (92, 88),
    "DespachosConsultaWindow": (60, 88),
    "ArregloCtaCteWindow": (65, 55),
    "ArregloSubdiarioWindow": (60, 60),
    "ParametrosWindow": (55, 45),
}

# Pantallas que el usuario pidió ver "en todo el largo y ancho
# disponible, o maximizada" (feedback 2026-08-19: Facturador, Recibo,
# Cta.Cte., Facturas Emitidas — "queda poco detalle" con el % que tenían
# antes) — en vez de subir el % a mano (ya se probó y no convergía, ver
# docstring de `PCT_MDI_POR_CLASE`), se maximizan de verdad dentro del
# panel MDI (`QMdiSubWindow.showMaximized()`): ocupan el 100% del panel
# real sin importar la resolución, y quedan afuera de `PCT_MDI_POR_CLASE`
# porque ese cálculo por % ya no aplica.
MAXIMIZAR_AL_ABRIR_MDI: frozenset[str] = frozenset(
    {"FacturadorWindow", "ReciboWindow", "CtaCteWindow", "FacturasEmitidasWindow"}
)


def _envolver_en_dos_lineas(texto: str, fuente: QFont, ancho_disponible: int) -> str:
    """Si `texto` no entra en una línea dentro de `ancho_disponible`
    (según `fuente`), lo parte en 2 líneas por la palabra más cercana a
    la mitad — pedido del usuario (2026-08-16): botones de tamaño
    uniforme en la barra de tareas, "achicar el texto o colocarlo en 2
    líneas" en vez de agrandar el botón o mostrar scroll."""
    metrica = QFontMetrics(fuente)
    # Descuenta ícono (40px, `ANCHO_ICONO_TAREA`) + padding/borde/espaciado
    # ícono-texto real del QSS del botón (`padding: 2px 8px` + `border:
    # 1px` + separación nativa de Qt entre ícono y texto, ~6px) — antes
    # era un "34" fijo calibrado para el ícono viejo de 22px (bug real,
    # 2026-08-20: quedó corto al agrandarlo, dejaba de envolver texto
    # largo que ya no entraba en una línea).
    ancho_libre = max(ancho_disponible - (ANCHO_ICONO_TAREA + 20), 40)
    if metrica.horizontalAdvance(texto) <= ancho_libre:
        return texto

    palabras = texto.split(" ")
    if len(palabras) < 2:
        return texto

    mejor_corte = 1
    mejor_diferencia = None
    for i in range(1, len(palabras)):
        linea1 = " ".join(palabras[:i])
        linea2 = " ".join(palabras[i:])
        diferencia = abs(metrica.horizontalAdvance(linea1) - metrica.horizontalAdvance(linea2))
        if mejor_diferencia is None or diferencia < mejor_diferencia:
            mejor_diferencia = diferencia
            mejor_corte = i
    return " ".join(palabras[:mejor_corte]) + "\n" + " ".join(palabras[mejor_corte:])


class MainMenuWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        # "FCMENU" -> "FcMenu II" en todo lo visible (feedback del
        # usuario, 2026-08-18, tercera ronda).
        self.setWindowTitle("FcMenu II  —  Facturación y Cuenta Corriente (Electrónica)")
        self.setWindowIcon(icono_app())
        self.resize(1280, 820)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self._seccion_actual: str | None = None
        # Subventanas abiertas por (sección, etiqueta) — evita disparar
        # el mismo módulo 2 veces (ver `_ejecutar_tarea`).
        self._ventanas_por_clave: dict[tuple[str, str], QMdiSubWindow] = {}

        self._construir_ui()
        self._ajustar_proporciones()
        self._seleccionar_seccion("A-B-M's")

        self._timer_reloj = QTimer(self)
        self._timer_reloj.timeout.connect(self._actualizar_fecha_hora)
        self._timer_reloj.start(INTERVALO_RELOJ_MS)
        self._actualizar_fecha_hora()

        self._monitor_internet = MonitorInternet(self)
        self._monitor_internet.resultado.connect(self._actualizar_estado_internet)
        self._timer_internet = QTimer(self)
        self._timer_internet.timeout.connect(self._monitor_internet.verificar)
        self._timer_internet.start(INTERVALO_CHEQUEO_INTERNET_MS)
        self._monitor_internet.verificar()

    # ------------------------------------------------------------------
    # Estructura general
    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout_raiz = QVBoxLayout(central)
        layout_raiz.setContentsMargins(0, 0, 0, 0)
        layout_raiz.setSpacing(0)

        layout_raiz.addWidget(self._armar_cabecera())

        cuerpo = QWidget()
        layout_cuerpo = QHBoxLayout(cuerpo)
        layout_cuerpo.setContentsMargins(0, 0, 0, 0)
        layout_cuerpo.setSpacing(0)
        layout_cuerpo.addWidget(self._armar_sidebar())
        layout_cuerpo.addWidget(self._armar_panel_principal(), stretch=1)
        layout_raiz.addWidget(cuerpo, stretch=1)

        layout_raiz.addWidget(self._armar_pie())

    def resizeEvent(self, event) -> None:  # noqa: N802 (override de Qt)
        super().resizeEvent(event)
        self._ajustar_proporciones()

    def _ajustar_proporciones(self) -> None:
        """Lateral 15-20% del ancho, pie 10% del alto — recalculado en
        cada resize para que la proporción se mantenga (con un piso/techo
        en píxeles para que no quede inusable en ventanas muy chicas o
        gigantes)."""
        self.sidebar.setFixedWidth(max(170, min(260, round(self.width() * 0.17))))
        self.pie.setFixedHeight(max(56, min(90, round(self.height() * 0.10))))
        # El ancho de los botones de la barra de tareas depende del ancho
        # real disponible — recalcular también en cada resize (no sólo al
        # cambiar de sección) para que sigan sin scroll ni texto cortado.
        if getattr(self, "_seccion_actual", None) and hasattr(self, "layout_barra_tareas"):
            self._reconstruir_barra_tareas()

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QWidget:
        cabecera = QWidget()
        cabecera.setStyleSheet(f"background-color: {Verde.OSCURO};")
        cabecera.setFixedHeight(92)
        fila = QHBoxLayout(cabecera)
        fila.setContentsMargins(20, 8, 20, 8)
        fila.setSpacing(16)

        lbl_logo = QLabel()
        pixmap_logo = logo_empresa(56)
        if not pixmap_logo.isNull():
            lbl_logo.setPixmap(pixmap_logo)
        else:
            lbl_logo.setPixmap(icono("app", 56).pixmap(56, 56))
        fila.addWidget(lbl_logo)

        bloque_titulo = QVBoxLayout()
        bloque_titulo.setSpacing(0)
        lbl_titulo = QLabel("FcMenu II")
        lbl_titulo.setStyleSheet(f"color: {Verde.BLANCO}; font-size: 22pt; font-weight: bold;")
        lbl_subtitulo = QLabel("Facturación y Cuenta Corriente (Electrónica)")
        lbl_subtitulo.setStyleSheet(f"color: {Verde.PASTEL}; font-size: 9.5pt;")
        # "FcMenu II - Versión 2.0" (feedback del usuario, 2026-08-18,
        # tercera ronda: "donde dice Versión 2.0 diga FCMENU - Versión
        # 2.0" — combinado con el rename de "FCMENU" a "FcMenu II").
        lbl_version = QLabel(f"FcMenu II - Versión {VERSION_APP}")
        lbl_version.setStyleSheet(f"color: {Verde.MEDIO_CLARO}; font-size: 8.5pt;")
        bloque_titulo.addWidget(lbl_titulo)
        bloque_titulo.addWidget(lbl_subtitulo)
        bloque_titulo.addWidget(lbl_version)
        fila.addLayout(bloque_titulo)

        fila.addStretch()

        self.lbl_estado_afip = QLabel(etiqueta_entorno_afip())
        self.lbl_estado_afip.setStyleSheet(
            f"""
            color: {Verde.AMARILLO};
            background-color: {Verde.MEDIO_OSCURO};
            border: 1px solid {Verde.AMARILLO};
            border-radius: 8px;
            font-size: 13pt;
            font-weight: bold;
            padding: 6px 18px;
            """
        )
        fila.addWidget(self.lbl_estado_afip)

        fila.addSpacing(24)

        bloque_reloj = QVBoxLayout()
        bloque_reloj.setSpacing(0)
        self.lbl_fecha = QLabel()
        self.lbl_fecha.setStyleSheet(f"color: {Verde.BLANCO}; font-size: 10.5pt;")
        self.lbl_fecha.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.lbl_hora = QLabel()
        self.lbl_hora.setStyleSheet(f"color: {Verde.BLANCO}; font-size: 15pt; font-weight: bold;")
        self.lbl_hora.setAlignment(Qt.AlignmentFlag.AlignRight)
        bloque_reloj.addWidget(self.lbl_fecha)
        bloque_reloj.addWidget(self.lbl_hora)
        fila.addLayout(bloque_reloj)

        return cabecera

    def _actualizar_fecha_hora(self) -> None:
        ahora = datetime.now()
        self.lbl_fecha.setText(formatear_fecha_corta(ahora.date()))
        self.lbl_hora.setText(ahora.strftime("%H:%M"))

    # ------------------------------------------------------------------
    # Panel lateral
    # ------------------------------------------------------------------
    # (clave_icono, etiqueta, nombre_seccion)
    _BOTONES_SIDEBAR = [
        ("nav_abms", "A-B-M's", "A-B-M's"),
        ("nav_ingreso", "Ingreso", "Ingreso"),
        ("nav_consultas", "Consultas", "Consultas"),
        ("nav_listados", "Listados", "Listados"),
        ("nav_varios", "Varios", "Varios"),
        ("nav_ayuda", "Ayuda", "Ayuda"),
    ]

    def _armar_sidebar(self) -> QWidget:
        self.sidebar = QWidget()
        self.sidebar.setStyleSheet(
            f"background-color: {Verde.CLARO}; border-right: 1px solid {Verde.MEDIO_CLARO};"
        )
        layout = QVBoxLayout(self.sidebar)
        layout.setContentsMargins(12, 0, 12, 0)
        layout.setSpacing(12)
        layout.addStretch()

        self._grupo_sidebar = QButtonGroup(self)
        self._grupo_sidebar.setExclusive(True)
        self._botones_sidebar: dict[str, QToolButton] = {}

        for clave_icono, etiqueta, nombre_seccion in self._BOTONES_SIDEBAR:
            boton = QToolButton()
            boton.setText(etiqueta)
            # Íconos al doble (pedido del usuario, 2026-08-20): 26->48px —
            # el mínimo de 68px del botón (línea de abajo) ya absorbe el
            # crecimiento sin que el botón cambie de tamaño.
            boton.setIcon(icono(clave_icono, 48))
            boton.setIconSize(QSize(48, 48))
            # Ícono a la izquierda, texto a la derecha, en una sola línea
            # (feedback del usuario, 2026-08-16 — antes el ícono quedaba
            # arriba del texto).
            boton.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
            boton.setCheckable(True)
            # 30% más altos que el original (52px) — pedido del usuario,
            # 2026-08-17.
            boton.setMinimumHeight(68)
            boton.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            boton.setStyleSheet(self._estilo_boton_sidebar())
            boton.clicked.connect(lambda _=False, s=nombre_seccion: self._seleccionar_seccion(s))
            self._grupo_sidebar.addButton(boton)
            self._botones_sidebar[nombre_seccion] = boton
            layout.addWidget(boton)

        layout.addStretch()
        return self.sidebar

    @staticmethod
    def _estilo_boton_sidebar() -> str:
        v = Verde
        return f"""
        QToolButton {{
            background-color: {v.PASTEL};
            color: {v.OSCURO};
            border: 1px solid {v.MEDIO_CLARO};
            border-radius: 10px;
            font-weight: bold;
            padding: 4px 4px 4px 14px;
            text-align: left;
        }}
        QToolButton:hover {{
            background-color: {v.MEDIO_CLARO};
        }}
        QToolButton:checked {{
            background-color: {v.MEDIO};
            color: {v.BLANCO};
            border: 1px solid {v.OSCURO};
        }}
        """

    # ------------------------------------------------------------------
    # Panel principal (título + barra de tareas + espacio de trabajo)
    # ------------------------------------------------------------------
    def _armar_panel_principal(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.lbl_titulo_seccion = QLabel()
        self.lbl_titulo_seccion.setFixedHeight(42)
        self.lbl_titulo_seccion.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.lbl_titulo_seccion.setStyleSheet(
            f"""
            background-color: {Verde.MEDIO_OSCURO};
            color: {Verde.BLANCO};
            font-size: 13pt;
            font-weight: bold;
            padding-left: 16px;
            """
        )
        layout.addWidget(self.lbl_titulo_seccion)

        # 2 líneas de botones cuando no entran cómodos en 1
        # (`ANCHO_MINIMO_BOTON_TAREA`, ver `_reconstruir_barra_tareas`) —
        # las 2 filas se arman siempre acá (la segunda arranca vacía/
        # oculta) y `_reconstruir_barra_tareas` decide cuántos botones va
        # en cada una según la sección activa.
        self.contenedor_tareas = QWidget()
        self.contenedor_tareas.setFixedHeight(ALTO_BARRA_TAREAS)
        self.contenedor_tareas.setStyleSheet(f"background-color: {Verde.PASTEL};")
        layout_tareas = QVBoxLayout(self.contenedor_tareas)
        layout_tareas.setContentsMargins(10, 6, 10, 6)
        layout_tareas.setSpacing(2)

        self.layout_barra_tareas = QHBoxLayout()
        self.layout_barra_tareas.setSpacing(6)
        layout_tareas.addLayout(self.layout_barra_tareas)

        self.layout_barra_tareas_2 = QHBoxLayout()
        self.layout_barra_tareas_2.setSpacing(6)
        layout_tareas.addLayout(self.layout_barra_tareas_2)

        layout.addWidget(self.contenedor_tareas)

        # Ventanas libres/cascada (probado 2026-08-16: se descartaron las
        # pestañas) — cada pantalla es una subventana que se puede mover,
        # redimensionar y superponer dentro del espacio de trabajo, más
        # parecido al MDI clásico del legacy.
        self.mdi = QMdiArea()
        self.mdi.setViewMode(QMdiArea.ViewMode.SubWindowView)
        self.mdi.setStyleSheet(f"QMdiArea {{ background: {Verde.CLARO}; }}")
        layout.addWidget(self.mdi, stretch=1)

        return panel

    def _seleccionar_seccion(self, nombre: str) -> None:
        self._seccion_actual = nombre
        self.lbl_titulo_seccion.setText(nombre.upper())
        boton = self._botones_sidebar.get(nombre)
        if boton is not None:
            boton.setChecked(True)
        self._reconstruir_barra_tareas()

    @staticmethod
    def _limpiar_fila_tareas(fila: QHBoxLayout) -> None:
        while fila.count():
            item = fila.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _reconstruir_barra_tareas(self) -> None:
        if self._seccion_actual is None:
            return
        self._limpiar_fila_tareas(self.layout_barra_tareas)
        self._limpiar_fila_tareas(self.layout_barra_tareas_2)

        items = self._secciones()[self._seccion_actual]

        # Se calcula a partir de `self.width()`/`self.sidebar.width()` (no
        # de `contenedor_tareas.width()`): recién después de que Qt activa
        # el layout completo un widget anidado refleja su ancho real, y
        # esto puede correr antes de eso (primera sección seleccionada en
        # `__init__`, todavía sin mostrar la ventana).
        ancho_disponible = max(self.width() - self.sidebar.width() - 40, 200)

        # 2 líneas de botones si NO ENTRAN cómodos en una sola (bug real,
        # 2026-08-20: con el umbral viejo por CANTIDAD de botones
        # (`MAX_BOTONES_UNA_FILA`), Consultas (9 botones, justo en el
        # límite) se quedaba en 1 fila con íconos ya agrandados —
        # `ANCHO_MINIMO_BOTON_TAREA` no entraba, el `QMainWindow` subía su
        # tamaño MÍNIMO y Qt terminaba agrandando la ventana sola al
        # entrar a esa sección. Ahora se decide por ancho real disponible
        # — mismo criterio que ya usaba Listados ("2 filas si no entran"),
        # pero calculado en vez de una cantidad fija que queda corta en
        # cuanto cambia el tamaño de ícono/fuente.
        dos_filas = ancho_disponible // len(items) < ANCHO_MINIMO_BOTON_TAREA
        if dos_filas:
            mitad = -(-len(items) // 2)  # ceil(len/2)
            items_fila1, items_fila2 = items[:mitad], items[mitad:]
        else:
            items_fila1, items_fila2 = items, []
        self.contenedor_tareas.setFixedHeight(
            ALTO_BARRA_TAREAS_2_FILAS if dos_filas else ALTO_BARRA_TAREAS
        )

        # "Calcular espacio para 8 botones y que sean proporcionales para
        # que no aparezca el scroll horizontal" (feedback del usuario,
        # 2026-08-16): el ancho de referencia (para decidir tipografía y
        # si el texto necesita 2 líneas) se calcula sobre el caso más
        # angosto REAL (8 o más si hay más botones que eso en la fila más
        # larga) — así el tamaño de letra queda consistente entre
        # secciones aunque una tenga menos botones (esos quedan más
        # anchos, "proporcionales" al espacio real disponible).
        ancho_referencia = ancho_disponible // max(len(items_fila1), len(items_fila2), 8)
        for clave_icono, etiqueta, callback in items_fila1:
            boton = self._crear_boton_tarea(clave_icono, etiqueta, callback, ancho_referencia)
            self.layout_barra_tareas.addWidget(boton, stretch=1)
        for clave_icono, etiqueta, callback in items_fila2:
            boton = self._crear_boton_tarea(clave_icono, etiqueta, callback, ancho_referencia)
            self.layout_barra_tareas_2.addWidget(boton, stretch=1)

    def _crear_boton_tarea(self, clave_icono: str, etiqueta: str, callback, ancho_referencia: int) -> QToolButton:
        boton = QToolButton()
        boton.setText(_envolver_en_dos_lineas(etiqueta, boton.font(), ancho_referencia))
        # Íconos casi al doble (pedido del usuario, 2026-08-20): 22->40px —
        # el alto fijo del botón (52px, ver `ALTO_BARRA_TAREAS`) no cambia,
        # sigue entrando con margen (padding+borde ~6px por lado).
        boton.setIcon(icono(clave_icono, ANCHO_ICONO_TAREA))
        boton.setIconSize(QSize(ANCHO_ICONO_TAREA, ANCHO_ICONO_TAREA))
        boton.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        boton.setFixedHeight(ALTO_BARRA_TAREAS - 12)
        boton.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        boton.setStyleSheet(
            f"""
            QToolButton {{
                background-color: {Verde.BLANCO};
                color: {Verde.OSCURO};
                border: 1px solid {Verde.MEDIO_CLARO};
                border-radius: 6px;
                font-weight: bold;
                font-size: 8.3pt;
                padding: 2px 8px;
            }}
            QToolButton:hover {{
                background-color: {Verde.MEDIO_CLARO};
            }}
            QToolButton:pressed {{
                background-color: {Verde.MEDIO};
                color: {Verde.BLANCO};
            }}
            QToolButton:disabled {{
                background-color: {Verde.MUY_CLARO};
                color: {Verde.TEXTO_SUAVE};
            }}
            """
        )
        if callback is None:
            boton.setEnabled(False)
            boton.setToolTip("Próximamente")
        else:
            boton.clicked.connect(
                lambda _=False, cb=callback, t=etiqueta, ic=clave_icono: self._ejecutar_tarea(cb, t, ic)
            )
        return boton

    def _ejecutar_tarea(self, callback, etiqueta: str, clave_icono: str) -> None:
        """Guarda la etiqueta del botón recién clickeado para que
        `_mostrar()` la use como título de la subventana que se abra
        (si el callback abre una) — "el título tiene que decir la
        opción elegida (Facturador, ABM Cliente, etc.), no el nombre
        interno de la pantalla" (feedback del usuario, 2026-08-16).

        **No permite disparar el mismo módulo 2 veces** (feedback del
        usuario, 2026-08-18, tercera ronda: "si un módulo se está
        ejecutando no se pueda disparar nuevamente... avise y no
        permita") — se identifica cada módulo por (sección, etiqueta),
        estable entre reconstrucciones de `_secciones()` (a diferencia
        de `id(callback)`: los callbacks de Listados son lambdas
        nuevas cada vez que se reconstruye la barra de tareas). Si ya
        hay una subventana abierta con esa clave, se la trae al frente
        en vez de abrir una segunda."""
        clave = (self._seccion_actual, etiqueta)
        existente = self._ventanas_por_clave.get(clave)
        if existente is not None:
            self.mdi.setActiveSubWindow(existente)
            self.mostrar_mensaje(f"'{etiqueta}' ya está abierto — no se abre de nuevo.", nivel="info")
            return
        self._titulo_tarea_actual = etiqueta
        self._clave_tarea_actual = clave
        self._clave_icono_actual = clave_icono
        callback()

    # ------------------------------------------------------------------
    # Pie
    # ------------------------------------------------------------------
    def _armar_pie(self) -> QWidget:
        self.pie = QWidget()
        self.pie.setStyleSheet(f"background-color: {Verde.PASTEL}; border-top: 1px solid {Verde.MEDIO_CLARO};")
        fila = QHBoxLayout(self.pie)
        fila.setContentsMargins(16, 0, 16, 0)
        fila.setSpacing(16)

        self.lbl_mensajes = QLabel("Listo")
        self.lbl_mensajes.setStyleSheet(f"color: {Verde.TEXTO};")
        self.lbl_mensajes.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        fila.addWidget(self.lbl_mensajes, stretch=1)

        self.lbl_internet = QLabel("Verificando conexión…")
        self.lbl_internet.setStyleSheet(f"color: {Verde.TEXTO_SUAVE}; font-weight: bold;")
        fila.addWidget(self.lbl_internet)

        # Corrido un poco a la izquierda del botón Salir (feedback del
        # usuario, 2026-08-16) — espacio extra en vez de quedar pegado.
        fila.addSpacing(40)

        btn_salir = QPushButton("Salir")
        # Ícono casi al doble (pedido del usuario, 2026-08-20).
        btn_salir.setIcon(icono("salir", 40))
        btn_salir.setIconSize(QSize(40, 40))
        # Más grande y con letra en verde oscuro (feedback del usuario,
        # 2026-08-16) — se distingue a propósito del verde lleno que usa
        # el resto de los botones de la app, es la acción de salir.
        btn_salir.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {Verde.PASTEL};
                color: {Verde.OSCURO};
                border: 1px solid {Verde.MEDIO_CLARO};
                border-radius: 8px;
                font-size: 11pt;
                font-weight: bold;
                padding: 8px 24px;
            }}
            QPushButton:hover {{
                background-color: {Verde.MEDIO_CLARO};
            }}
            QPushButton:pressed {{
                background-color: {Verde.MEDIO};
                color: {Verde.BLANCO};
            }}
            """
        )
        btn_salir.clicked.connect(self.close)
        fila.addWidget(btn_salir)

        return self.pie

    def mostrar_mensaje(self, texto: str, nivel: str = "info") -> None:
        """Panel de mensajes del pie — pantallas/servicios pueden llamar
        a esto (vía la instancia de `MainMenuWindow`) para reportar
        avisos e errores, además del uso interno al abrir/cerrar
        pantallas."""
        colores = {"info": Verde.TEXTO, "ok": Verde.MEDIO_OSCURO, "error": Verde.ALERTA}
        self.lbl_mensajes.setStyleSheet(f"color: {colores.get(nivel, Verde.TEXTO)};")
        self.lbl_mensajes.setText(texto)

    # ------------------------------------------------------------------
    # Estado de conexión a Internet
    # ------------------------------------------------------------------
    def _actualizar_estado_internet(self, hay_internet: bool) -> None:
        if hay_internet:
            self.lbl_internet.setText("●  Hay Internet")
            self.lbl_internet.setStyleSheet(f"color: {Verde.OSCURO}; font-weight: bold;")
        else:
            self.lbl_internet.setText("●  Sin Internet")
            self.lbl_internet.setStyleSheet(f"color: {Verde.ALERTA}; font-weight: bold;")

    # ------------------------------------------------------------------
    # Mapeo de pantallas por sección (reparto confirmado con el usuario,
    # 2026-08-16 — sigue los nombres reales de "Tareas" del legacy
    # FCMENU.frm y manda al último botón, "Varios", lo que ahí no encaja).
    # ------------------------------------------------------------------
    def _secciones(self) -> dict[str, list[tuple[str, str, object]]]:
        return {
            # Descuentos del Cliente/Notas del Cliente/Notas del Artículo/
            # Modif. de Precios salieron de acá (feedback del usuario,
            # 2026-08-17): los dos primeros ya son accesibles desde la
            # propia ficha (botones "Cargar Descuentos..."/"Notas" de
            # `ClienteDetalleDialog`/`ArticuloDetalleDialog`, sin pasar por
            # el menú), Modif. de Precios se mudó a Ingreso (ver abajo,
            # ahora "Precios").
            "A-B-M's": [
                ("clientes", "Clientes", self._abrir_clientes),
                ("articulos", "Artículos", self._abrir_articulos),
                ("tablas", "Tablas Varias", self._abrir_tablas),
            ],
            # N.Credito/N.Debito habilitados (2026-08-19) — alcance
            # "Concepto Libre" (`NotaCreditoConceptoWindow`, Motivo != 1/
            # "DEV.MERC."); la variante "Devolución de Mercadería" (grilla
            # de artículos + reversa de Stock) todavía no está migrada,
            # queda para una ronda posterior. Un solo botón para las dos
            # (feedback del usuario, 2026-08-19: "sacar el botón nota de
            # débito... colocar la leyenda 'Nota de Crédito y Débito'")
            # — la elección real entre Crédito/Débito vive DENTRO de la
            # ventana (radio Tipo), no hace falta un botón por tipo.
            # Cotización/Remito siguen deshabilitados ("Próximamente",
            # `callback=None`, mismo patrón que "Constancia AFIP" en
            # Varios) — Remito no existe en el legacy (confirmado, sin
            # `.frm` propio); Cotización es un tipo de comprobante del
            # Facturador (TipoFac=4, sin CAE), no la Cotización del Dólar
            # ya migrada.
            "Ingreso": [
                ("factura", "Factura", self._abrir_factura),
                ("ncredito", "Nota de Crédito y Débito", self._abrir_nota_credito),
                ("ncredito_merc", "N.Créd. x Mercadería", self._abrir_nota_credito_mercaderia),
                ("cotizacion_venta", "Cotización", None),
                ("remito", "Remito", None),
                ("recibo", "Recibos", self._abrir_recibo),
                ("precios", "Precios", self._abrir_mod_precios),
                ("ingreso_stock", "Stock", self._abrir_stock_movimiento),
            ],
            "Consultas": [
                ("ctacte", "Cuenta Corriente", self._abrir_ctacte),
                ("cobranzas", "Cobranzas por Zona", self._abrir_cobranzas_zona),
                ("facturas_emitidas", "Facturas Emitidas", self._abrir_facturas_emitidas),
                ("totales", "Totales de Facturación", self._abrir_totales_diarios),
                ("cheques", "Cheques", self._abrir_cheques),
                ("ventas", "Ventas de Art. por Cliente", self._abrir_ventas_seccion),
                ("ventas", "Ventas por Artículo", self._abrir_ventas_articulo),
                ("stock", "Stock", self._abrir_stock_consulta),
                ("despachos", "Despachos", self._abrir_despachos),
            ],
            # Los 12 listados ya estaban implementados adentro de
            # ListadosWindow (un combo interno) pero sólo se veía "el de
            # Clientes" al no tener botón propio (feedback del usuario,
            # 2026-08-16: "faltan todos los listados") — un botón por
            # reporte, cada uno abre la misma ventana con ese listado ya
            # elegido (el combo sigue ahí para cambiar sin volver al menú).
            # Orden temático de izquierda a derecha (feedback del usuario,
            # 2026-08-18, segunda ronda): Clientes, Cta.Cte., Cobranza,
            # Ventas, Impuestos — con 12 botones, la barra se parte en 2
            # líneas (ver `_reconstruir_barra_tareas`) y este orden sigue
            # siendo "izquierda a derecha, arriba y después abajo".
            "Listados": [
                ("clientes", "Clientes", lambda r="clientes": self._abrir_listado(r)),
                ("ctacte", "Deuda Pendiente", lambda r="deuda_pendiente": self._abrir_listado(r)),
                ("ctacte", "Estado de Cuenta", lambda r="estado_cuenta": self._abrir_listado(r)),
                ("ctacte", "Saldos de Cta. Cte.", lambda r="saldos": self._abrir_listado(r)),
                ("cobranzas", "Subd. de Cobranzas AFIP", lambda r="subdiario_cobranzas": self._abrir_listado(r)),
                ("cobranzas", "Planilla de Cobranzas", lambda r="planilla_cobranzas": self._abrir_listado(r)),
                ("cobranzas", "Comis. x Cobranzas", lambda r="comisiones_cobranzas": self._abrir_listado(r)),
                ("precios", "Lista de Precios", lambda r="precios": self._abrir_listado(r)),
                ("ventas", "Subdiario de Ventas", lambda r="subdiario_ventas": self._abrir_listado(r)),
                ("ventas", "Subd. Vtas. (Comisiones)", lambda r="subdiario_comisiones": self._abrir_listado(r)),
                ("listados", "Ingresos Brutos", lambda r="ingresos_brutos": self._abrir_listado(r)),
                ("constancia_afip", "Percepciones ARBA", lambda r="percepciones_arba": self._abrir_listado(r)),
            ],
            "Varios": [
                ("cotizacion", "Cotización del Dólar", self._abrir_cotizacion),
                ("parametros", "Parámetros", self._abrir_parametros),
                ("constancia_afip", "Constancia AFIP", None),  # diferido, ver docstring histórico
                ("arreglos", "Arreglo Cta.Cte.", self._abrir_arreglo_ctacte),
                ("arreglos", "Arreglo Subdiario de Ventas", self._abrir_arreglo_subdiario),
            ],
            # Sección nueva (feedback del usuario, 2026-08-17).
            "Ayuda": [
                ("acerca_de", "Acerca de FcMenu II", self._abrir_acerca_de),
                ("referencias", "Referencias", self._abrir_referencias),
                ("manual", "Manual", self._abrir_manual),
                ("sitio_web", "Sitio Alestel", self._abrir_sitio_web),
                ("lista_precios_pdf", "Lista de Precios (PDF)", self._abrir_lista_precios_pdf),
            ],
        }

    # ------------------------------------------------------------------
    def _mostrar(self, ventana: QWidget) -> None:
        # El título de la subventana tiene que decir la OPCIÓN elegida
        # (ej. "Facturador", "Clientes"), no el título interno de la
        # pantalla (ej. "Búsqueda de Clientes") — feedback del usuario,
        # 2026-08-16. `_ejecutar_tarea` deja la etiqueta del botón
        # clickeado en `_titulo_tarea_actual` justo antes de invocar el
        # callback que termina llamando acá.
        #
        # Prefijo con el título principal de la sección (feedback del
        # usuario, 2026-08-18, tercera ronda: "Consultas - Cuenta
        # Corriente" / "Ingreso - Recibos") — `self._seccion_actual` es
        # la misma sección activa que armó la barra de tareas que llevó
        # a este callback.
        titulo_tarea = getattr(self, "_titulo_tarea_actual", None) or ventana.windowTitle() or "Ventana"
        titulo = f"{self._seccion_actual} - {titulo_tarea}" if self._seccion_actual else titulo_tarea
        ventana.setWindowTitle(titulo)
        # Cada ventana muestra el mismo glifo que tenía el botón del menú
        # que la abrió (pedido del usuario, 2026-08-20) — `clave_icono`
        # viaja desde `_crear_boton_tarea` vía `_ejecutar_tarea`.
        icono_ventana = icono(getattr(self, "_clave_icono_actual", "app"), 32)
        ventana.setWindowIcon(icono_ventana)
        # Sin esto, QMdiSubWindow sólo se oculta al cerrar (no se destruye
        # de verdad) — cada pantalla abre su propia sesión de BD
        # (`get_session()`) y no tiene sentido dejarla viva-pero-oculta
        # después de que el usuario la cerró.
        ventana.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        # **Bug real encontrado con capturas de pantalla reales**
        # (2026-08-19: "sigue mal" — las ventanas quedaban chicas pese a
        # `redimensionar_pct_pantalla()` ya aplicado en el `__init__` de
        # cada pantalla): `QMdiArea.addSubWindow()` NO hereda el tamaño
        # que ya tenía el widget envuelto — el `QMdiSubWindow` que lo
        # envuelve arranca con su propio tamaño "natural" (bastante más
        # chico, confirmado con un script aislado: un widget de
        # 1500×900 quedaba envuelto en un `QMdiSubWindow` de apenas
        # 120×40 antes de mostrarlo). Por eso NINGÚN tamaño que se le
        # haya dado a `ventana` — ni los px fijos de rondas anteriores
        # ni el % de pantalla actual — se veía reflejado nunca: el
        # tamaño se perdía en el momento de envolverla. Se guarda ANTES
        # de envolver y se reaplica a mano sobre el WRAPPER, no sólo
        # sobre `ventana`.
        tamano_pedido = ventana.size()
        subventana = self.mdi.addSubWindow(ventana)
        subventana.resize(tamano_pedido)
        # `QMdiSubWindow` no siempre hereda el ícono del widget interno
        # por sí solo — se setea explícito para que el glifo se vea
        # también en la barra de título de la subventana dentro del MDI.
        subventana.setWindowIcon(icono_ventana)

        # Bug real de Qt (confirmado con un script de prueba aislado,
        # `QMdiArea.addSubWindow` + cerrar el widget interno): cerrar el
        # widget embebido NO cierra su `QMdiSubWindow` — sólo funciona al
        # revés. Sin este parche, el botón "Cerrar" de cada pantalla
        # oculta el contenido pero el marco de la subventana queda
        # flotando vacío en el MDI ("el botón cerrar limpia la ventana
        # pero no la elimina", feedback del usuario, 2026-08-17). Se
        # engancha el closeEvent real de `ventana` para, si quedó
        # aceptado (una pantalla puede seguir cancelando su cierre, ej.
        # "Salir" de Factura/Recibo con datos sin guardar), cerrar
        # también `subventana` — que Qt ya crea con WA_DeleteOnClose, así
        # que cerrarla la destruye de una.
        close_event_original = ventana.closeEvent

        def _cerrar_tambien_subventana(event, _original=close_event_original, _sub=subventana) -> None:
            _original(event)
            if event.isAccepted():
                _sub.close()

        ventana.closeEvent = _cerrar_tambien_subventana

        # Registro para no permitir 2 instancias del mismo módulo a la
        # vez (ver `_ejecutar_tarea`) — se limpia solo al cerrarse de
        # verdad. `_clave_tarea_actual` sólo existe cuando `_mostrar()`
        # se llamó desde un callback de la barra de tareas (todos los
        # casos reales de la app) — si algún día se llama desde otro
        # lado, simplemente no participa de la deduplicación.
        clave = getattr(self, "_clave_tarea_actual", None)
        if clave is not None:
            self._ventanas_por_clave[clave] = subventana
            subventana.destroyed.connect(lambda: self._ventanas_por_clave.pop(clave, None))

        subventana.destroyed.connect(lambda: self.mostrar_mensaje(f"Se cerró: {titulo}"))
        subventana.show()
        self.mdi.setActiveSubWindow(subventana)
        self._posicionar_subventana(subventana, ventana)
        self.mostrar_mensaje(f"Se abrió: {titulo}", nivel="ok")

    def _posicionar_subventana(self, subventana: QMdiSubWindow, ventana: QWidget) -> None:
        """Centrada en el panel principal en vez de en cascada desde el
        extremo superior izquierdo (feedback del usuario, 2026-08-18,
        tercera ronda: "quedaría más agradable si las ventanas/módulos
        aparecen en el centro del panel principal"). Listados pidió un
        margen puntual en mm en vez de centrado simple por tamaño ("a
        30mm por lado lateral y a 50mm de arriba") — acá se interpreta
        "en el centro... a 30mm/50mm" como simétrico (30mm también a
        los costados, 50mm también abajo), no sólo un ancho fijo pegado
        arriba.

        Para las pantallas de `PCT_MDI_POR_CLASE`, el tamaño se
        recalcula acá contra el panel MDI real (ver docstring de esa
        constante) — pisa el tamaño que la propia ventana ya traía de
        su `__init__` (útil sólo para cuando se corre standalone, sin
        este MDI de por medio).

        **Piso de seguridad = `minimumSizeHint()` real del contenido**
        (feedback del usuario, 2026-08-19, con capturas de pantalla:
        "algunas muy poco, otras son irreconocibles... que entren todos
        los campos completos en el área") — bug real encontrado:
        a diferencia de una ventana normal (que el propio SO no deja
        angostar más allá de lo que su layout necesita), un
        `QMdiSubWindow` SÍ se puede forzar por `resize()` a un tamaño
        MENOR al que su contenido necesita — ahí los campos quedan
        superpuestos/ilegibles (confirmado con un script: Arreglo
        Subdiario necesita mínimo 630×578, Recibo 1384×365 — el % que
        tenían calculado antes les daba menos que eso en un panel MDI
        chico, y se rompían). Nunca se resigna por debajo de ese
        mínimo real, aunque eso implique que la ventana quede más ancha
        que el panel visible (con scroll del propio `QMdiArea`, mejor
        que campos rotos)."""
        area = self.mdi.viewport().size()
        nombre_clase = type(ventana).__name__
        minimo = subventana.minimumSizeHint()
        if nombre_clase in MAXIMIZAR_AL_ABRIR_MDI:
            # Antes de maximizar, se deja una geometría "restaurada"
            # acotada al panel MDI real (feedback del usuario, 2026-08-19:
            # "cuando se normalizan no se salen de cuadro... ahora son
            # muy grandes") — bug real: sin este `resize()`/`move()` acá,
            # Qt recuerda como geometría "restaurada" (a la que vuelve al
            # des-maximizar) la que la propia `ventana` ya traía de su
            # `__init__` (`redimensionar_pct_pantalla()`, calculada contra
            # la PANTALLA física completa, mucho más grande que este panel
            # MDI) — al des-maximizar, el tamaño restaurado se salía del
            # cuadro. 85%×80% del panel real (mismo criterio de piso de
            # `minimumSizeHint()` que el resto de esta función) deja
            # margen para ver que sigue habiendo un panel MDI alrededor.
            ancho_restaurado = max(round(area.width() * 0.85), minimo.width(), 300)
            alto_restaurado = max(round(area.height() * 0.80), minimo.height(), 300)
            subventana.resize(ancho_restaurado, alto_restaurado)
            x = max((area.width() - subventana.width()) // 2, 0)
            y = max((area.height() - subventana.height()) // 2, 0)
            subventana.move(x, y)
            subventana.showMaximized()
            return
        if nombre_clase == "ListadosWindow":
            margen_lateral = _mm_a_px(MARGEN_LISTADOS_LATERAL_MM)
            margen_superior = _mm_a_px(MARGEN_LISTADOS_SUPERIOR_MM)
            ancho = max(area.width() - 2 * margen_lateral, minimo.width(), 300)
            alto = max(area.height() - 2 * margen_superior, minimo.height(), 300)
            subventana.resize(ancho, alto)
            subventana.move(margen_lateral, margen_superior)
            return
        porcentajes = PCT_MDI_POR_CLASE.get(nombre_clase)
        if porcentajes is not None:
            ancho_pct, alto_pct = porcentajes
            ancho = max(round(area.width() * ancho_pct / 100), minimo.width(), 300)
            alto = max(round(area.height() * alto_pct / 100), minimo.height(), 300)
            subventana.resize(ancho, alto)
        x = max((area.width() - subventana.width()) // 2, 0)
        y = max((area.height() - subventana.height()) // 2, 0)
        subventana.move(x, y)

    # -- A-B-M's --------------------------------------------------------
    def _abrir_clientes(self) -> None:
        """Orden pedido por el usuario (2026-08-16): la ficha (vacía)
        aparece primero como pantalla principal, la búsqueda se abre
        encima recién después — al revés del flujo viejo (búsqueda
        primero, detalle recién al elegir), para que el operador no
        pierda de vista qué opción eligió."""
        from migration.services import ClienteService

        from .cliente_detalle_dialog import ClienteDetalleDialog

        db = get_session()
        repos = RepositoryFactory(db)
        servicio = ClienteService(db)
        # `on_buscar` necesita una referencia a `detalle`, que todavía no
        # existe al construirlo — se resuelve con este casillero que se
        # completa apenas termina de construirse (el callback recién se
        # ejecuta más tarde, con un clic real en "Buscar...").
        casillero: dict = {}
        detalle = ClienteDetalleDialog(
            repos, servicio, cliente=None, parent=self,
            on_buscar=lambda: self._buscar_y_cargar_cliente(casillero["detalle"]),
        )
        casillero["detalle"] = detalle
        detalle.destroyed.connect(db.close)
        self._mostrar(detalle)
        self._buscar_y_cargar_cliente(detalle)

    def _buscar_y_cargar_cliente(self, detalle) -> None:
        from .cliente_busqueda_window import ClienteBusquedaWindow

        buscador = ClienteBusquedaWindow(parent=detalle, modo_seleccion=True)
        if buscador.exec() == QDialog.DialogCode.Accepted and buscador.cliente_elegido is not None:
            detalle.cargar_cliente(buscador.cliente_elegido)

    def _abrir_articulos(self) -> None:
        """Mismo orden ficha-primero-búsqueda-después que Clientes."""
        from migration.services import ArticuloService

        from .articulo_detalle_dialog import ArticuloDetalleDialog

        db = get_session()
        repos = RepositoryFactory(db)
        servicio = ArticuloService(db)
        casillero: dict = {}
        detalle = ArticuloDetalleDialog(
            repos, servicio, articulo=None, parent=self,
            on_buscar=lambda: self._buscar_y_cargar_articulo(casillero["detalle"]),
        )
        casillero["detalle"] = detalle
        detalle.destroyed.connect(db.close)
        self._mostrar(detalle)
        self._buscar_y_cargar_articulo(detalle)

    def _buscar_y_cargar_articulo(self, detalle) -> None:
        from .articulo_busqueda_window import ArticuloBusquedaWindow

        buscador = ArticuloBusquedaWindow(parent=detalle, modo_seleccion=True)
        if buscador.exec() == QDialog.DialogCode.Accepted and buscador.articulo_elegido is not None:
            detalle.cargar_articulo(buscador.articulo_elegido)

    # `_abrir_descuentos_cliente`/`_abrir_notas_cliente`/`_abrir_notas_
    # articulo` se sacaron de acá (feedback del usuario, 2026-08-17):
    # eran caminos duplicados hacia funciones que YA están accesibles
    # directo desde la propia ficha (`ClienteDetalleDialog.btn_
    # descuentos`/`btn_notas`, `ArticuloDetalleDialog.btn_notas`), sin
    # tener que buscar el Cliente/Artículo dos veces.

    def _abrir_mod_precios(self) -> None:
        """Se invoca ahora desde Ingreso ("Precios") — se sacó del
        sidebar de A-B-M's (feedback del usuario, 2026-08-17), este
        método no cambió."""
        from .mod_precios_window import ModPreciosWindow

        self._mostrar(ModPreciosWindow())

    def _abrir_tablas(self) -> None:
        from .tablas_window import TablasWindow

        self._mostrar(TablasWindow())

    # -- Ingreso --------------------------------------------------------
    def _abrir_factura(self) -> None:
        from .facturador_window import FacturadorWindow

        self._mostrar(FacturadorWindow())

    def _abrir_recibo(self) -> None:
        from .recibo_window import ReciboWindow

        self._mostrar(ReciboWindow())

    def _abrir_nota_credito(self) -> None:
        # Un solo botón para Crédito/Débito (feedback del usuario,
        # 2026-08-19) — la ventana arranca en modo Nota de Crédito, el
        # operador cambia el radio Tipo adentro si necesita Débito.
        from .nota_credito_concepto_window import NotaCreditoConceptoWindow

        self._mostrar(NotaCreditoConceptoWindow())

    def _abrir_nota_credito_mercaderia(self) -> None:
        from .nota_credito_mercaderia_window import NotaCreditoMercaderiaWindow

        self._mostrar(NotaCreditoMercaderiaWindow())

    def _abrir_stock_movimiento(self) -> None:
        from .stock_movimiento_window import StockMovimientoWindow

        self._mostrar(StockMovimientoWindow())

    # -- Consultas --------------------------------------------------------
    def _abrir_ctacte(self) -> None:
        from .ctacte_window import CtaCteWindow

        self._mostrar(CtaCteWindow())

    def _abrir_cobranzas_zona(self) -> None:
        from .cobranzas_zona_window import CobranzasZonaWindow

        self._mostrar(CobranzasZonaWindow())

    def _abrir_facturas_emitidas(self) -> None:
        from .facturas_emitidas_window import FacturasEmitidasWindow

        self._mostrar(FacturasEmitidasWindow())

    def _abrir_totales_diarios(self) -> None:
        from .totales_diarios_window import TotalesDiariosWindow

        self._mostrar(TotalesDiariosWindow())

    def _abrir_cheques(self) -> None:
        from .cheques_consulta_window import ChequesConsultaWindow

        self._mostrar(ChequesConsultaWindow())

    def _abrir_ventas_seccion(self) -> None:
        from .ventas_seccion_window import VentasSeccionWindow

        self._mostrar(VentasSeccionWindow())

    def _abrir_ventas_articulo(self) -> None:
        from .ventas_articulo_window import VentasArticuloWindow

        self._mostrar(VentasArticuloWindow())

    def _abrir_stock_consulta(self) -> None:
        from .stock_consulta_window import StockConsultaWindow

        self._mostrar(StockConsultaWindow())

    def _abrir_despachos(self) -> None:
        from .despachos_consulta_window import DespachosConsultaWindow

        self._mostrar(DespachosConsultaWindow())

    # -- Listados ---------------------------------------------------------
    def _abrir_listado(self, reporte: str) -> None:
        from .listados_window import ListadosWindow

        self._mostrar(ListadosWindow(reporte_inicial=reporte))

    # -- Varios -----------------------------------------------------------
    def _abrir_cotizacion(self) -> None:
        from .cotizacion_window import CotizacionWindow

        self._mostrar(CotizacionWindow())

    def _abrir_parametros(self) -> None:
        from .parametros_window import ParametrosWindow

        self._mostrar(ParametrosWindow())

    def _abrir_arreglo_ctacte(self) -> None:
        from .arreglo_ctacte_window import ArregloCtaCteWindow

        self._mostrar(ArregloCtaCteWindow())

    def _abrir_arreglo_subdiario(self) -> None:
        from .arreglo_subdiario_window import ArregloSubdiarioWindow

        self._mostrar(ArregloSubdiarioWindow())

    # -- Ayuda --------------------------------------------------------------
    def _abrir_acerca_de(self) -> None:
        """"Apunta a la ventana de presentación" (pedido del usuario) —
        la única pantalla de presentación real del legacy es
        `Inicio.frm`, migrada acá como diálogo modal informativo."""
        from .acerca_de_dialog import AcercaDeDialog

        AcercaDeDialog(parent=self).exec()

    def _abrir_referencias(self) -> None:
        """PDF de una página con tips/ayudamemorias — se genera una
        sola vez (`migration.referencias.generar_pdf_referencias`,
        cachea en `assets/docs/`) y se muestra con el mismo visor de
        PDF ya usado en Recibo (`PdfPreviewDialog`)."""
        from migration.referencias import generar_pdf_referencias

        from .pdf_preview_dialog import PdfPreviewDialog

        try:
            ruta_pdf = generar_pdf_referencias()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Referencias", f"No se pudo generar el PDF de Referencias:\n{exc}")
            return

        PdfPreviewDialog(ruta_pdf, titulo="Referencias — Tips y Ayudamemorias", parent=self).exec()

    # Manual real ya generado por el usuario en `assets/docs/` (pedido
    # del usuario, 2026-08-18, tercera ronda: "el manual del usuario
    # está en assets\docs... ya se puede conectar") — mismo patrón que
    # `RUTA_LISTA_PRECIOS_PDF`, pero ESTA sí vive dentro del proyecto
    # (`Path(__file__)` en vez de una ruta fija fuera de él).
    RUTA_MANUAL_USUARIO_PDF = Path(__file__).resolve().parent.parent.parent / "assets" / "docs" / "Manual_Usuario_FCMENU.pdf"

    def _abrir_manual(self) -> None:
        from .pdf_preview_dialog import PdfPreviewDialog

        if not self.RUTA_MANUAL_USUARIO_PDF.exists():
            QMessageBox.warning(
                self,
                "Manual",
                f"No se encontró el archivo:\n{self.RUTA_MANUAL_USUARIO_PDF}",
            )
            return
        PdfPreviewDialog(self.RUTA_MANUAL_USUARIO_PDF, titulo="Manual del Usuario", parent=self).exec()

    def _abrir_sitio_web(self) -> None:
        QDesktopServices.openUrl(QUrl("https://www.alestel.com.ar"))

    # Ruta fija fuera del proyecto (pedido del usuario, 2026-08-17,
    # segunda ronda) — el PDF lo mantiene el usuario, no lo genera FCMENU.
    RUTA_LISTA_PRECIOS_PDF = Path(r"c:\ListadePrecios\ListaPrecios.pdf")

    def _abrir_lista_precios_pdf(self) -> None:
        from .pdf_preview_dialog import PdfPreviewDialog

        if not self.RUTA_LISTA_PRECIOS_PDF.exists():
            QMessageBox.warning(
                self,
                "Lista de Precios",
                f"No se encontró el archivo:\n{self.RUTA_LISTA_PRECIOS_PDF}",
            )
            return
        PdfPreviewDialog(self.RUTA_LISTA_PRECIOS_PDF, titulo="Lista de Precios", parent=self).exec()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (override de Qt)
        self.mdi.closeAllSubWindows()
        if self.mdi.subWindowList():
            # Alguna pantalla hija canceló su cierre (p.ej. cambios sin
            # guardar) — no cerrar FCMENU por debajo suyo.
            event.ignore()
            return
        self.db.close()
        super().closeEvent(event)
