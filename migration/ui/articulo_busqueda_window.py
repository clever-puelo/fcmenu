"""Ventana de Búsqueda de Artículos — punto de entrada del ABM, y también
selector reutilizable para el Facturador.

Convención de navegación confirmada con el usuario (2026-08-15), aplicada
a todo el sistema salvo aclaración específica: primero se pide un
criterio de búsqueda, se muestra la lista de resultados abajo, y recién
al elegir uno de la lista se abre el diálogo de Detalle
(`ArticuloDetalleDialog`) para Ver/Cambio/Baja. El botón Alta, abajo a la
izquierda, abre el mismo diálogo de Detalle pero en blanco.

Para Artículos el criterio principal es la Sección (`Fctabla1` CTAB='SC')
— elegirla lista abajo todos los artículos de esa Sección. Además hay un
filtro secundario por Descripción, para cuando no se sabe la Sección.

La grilla (`TablaBusqueda`) ordena por columna al clic en el título
(convención de sistema, 2026-08-16); elegir una fila es con un clic o
doble clic (feedback del usuario, 2026-08-15 — reemplaza la decisión
previa de exigir doble clic, ver nota en `cliente_busqueda_window.py`).

**Modo selección** (2026-08-16, para `DetFact.frm`: tecla F2 o Sección
no encontrada al tipear un renglón, que en el legacy abrían `Busqueda.frm`
en modo sólo-búsqueda): con `modo_seleccion=True` esta misma ventana se
usa como `QDialog` modal (`.exec()`), sin botón Alta (acá no se da de
alta catálogo, sólo se elige uno existente para el renglón), y elegir una
fila deja el artículo en `self.articulo_elegido` y cierra con
`Accepted` en vez de abrir `ArticuloDetalleDialog`. El uso normal
(standalone, `main_articulo.py`) no cambia: `QDialog` admite `.show()`
no-modal igual que antes se usaba `QMainWindow`.
"""

from __future__ import annotations

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.repository import RepositoryFactory
from migration.services import ArticuloService

from .articulo_detalle_dialog import ArticuloDetalleDialog
from .decimals import format_decimal
from .widgets import TablaBusqueda


class ArticuloBusquedaWindow(QDialog):
    """Punto de entrada del ABM de Artículos (equivalente conceptual a
    Busqueda.frm + AbmArt.frm del legacy), y selector modal reutilizable
    cuando `modo_seleccion=True` (ver docstring del módulo)."""

    def __init__(self, parent: QWidget | None = None, *, modo_seleccion: bool = False):
        super().__init__(parent)
        self.modo_seleccion = modo_seleccion
        self.articulo_elegido = None
        self.setWindowTitle("Selección de Artículo" if modo_seleccion else "Búsqueda de Artículos")
        self.resize(700, 520)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.articulo_service = ArticuloService(self.db)

        self._construir_ui()
        self._cargar_secciones()

        if self.modo_seleccion:
            # Feedback del usuario (2026-08-15): al abrir con F2 desde el
            # Detalle, el combo de Sección debe aparecer ya desplegado —
            # es el único dato que hace falta para arrancar, no debería
            # hacer falta un clic más para verlo. `showPopup()` necesita
            # que el combo ya esté mostrado en pantalla, así que se
            # difiere a la próxima vuelta del loop de eventos (arranca
            # recién cuando `exec()` corre el loop modal).
            self.combo_seccion.setFocus()
            QTimer.singleShot(0, self.combo_seccion.showPopup)

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        layout.addLayout(self._armar_fila_filtros())

        grupo_lista = QGroupBox("Artículos")
        layout_lista = QVBoxLayout(grupo_lista)
        self.tabla = TablaBusqueda(["Sección", "Código", "Descripción", "Precio", "Stock"])
        # Selección con un clic o doble clic (feedback del usuario,
        # 2026-08-15) — antes sólo abría con doble clic/Enter.
        self.tabla.itemClicked.connect(lambda _: self._on_elegir_de_lista())
        self.tabla.itemDoubleClicked.connect(lambda _: self._on_elegir_de_lista())
        self.tabla.itemActivated.connect(lambda _: self._on_elegir_de_lista())
        layout_lista.addWidget(self.tabla)
        layout.addWidget(grupo_lista, stretch=1)

        layout.addLayout(self._armar_barra_botones())

    def _armar_fila_filtros(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        fila.addWidget(QLabel("Sección:"))
        self.combo_seccion = QComboBox()
        self.combo_seccion.currentIndexChanged.connect(self._refrescar_lista)
        fila.addWidget(self.combo_seccion)

        fila.addStretch()
        return fila

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        if not self.modo_seleccion:
            self.btn_alta = QPushButton("Alta")
            self.btn_alta.setToolTip("Dar de alta un artículo nuevo")
            self.btn_alta.clicked.connect(self._on_alta)
            fila.addWidget(self.btn_alta)

        fila.addStretch()

        self.btn_cerrar = QPushButton("Cancelar" if self.modo_seleccion else "Cerrar")
        self.btn_cerrar.clicked.connect(self.reject)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    # Carga de Secciones y lista de artículos
    # ------------------------------------------------------------------
    def _cargar_secciones(self) -> None:
        secciones = self.repos.fctablas().by_ctab("SC")
        self.combo_seccion.clear()
        self.combo_seccion.addItem("(elegir Sección)", "")
        for fila in secciones:
            cod = (fila.COD or "").strip()
            self.combo_seccion.addItem(f"{cod} - {fila.DESCRI or ''}", cod)

    def _refrescar_lista(self) -> None:
        cod_seccion = self.combo_seccion.currentData()
        if not cod_seccion:
            self.tabla.cargar_filas([])
            return

        articulos = self.repos.articulo().by_seccion(cod_seccion)
        filas = [
            (
                [
                    (a.COD1 or "").strip(),
                    (a.COD2 or "").strip(),
                    a.DESCRI or "",
                    f"$ {format_decimal(a.PREC or 0)}",
                    str(a.STOCK if a.STOCK is not None else 0),
                ],
                a,
            )
            for a in articulos
        ]
        self.tabla.cargar_filas(filas)

    def _on_elegir_de_lista(self) -> None:
        articulo = self.tabla.objeto_seleccionado()
        if articulo is None:
            return

        if self.modo_seleccion:
            self.articulo_elegido = articulo
            self.accept()
            return

        dialogo = ArticuloDetalleDialog(self.repos, self.articulo_service, articulo=articulo, parent=self)
        dialogo.exec()
        if dialogo.guardado_o_eliminado:
            self._refrescar_lista()

    def _on_alta(self) -> None:
        seccion_sugerida = self.combo_seccion.currentData() or None
        dialogo = ArticuloDetalleDialog(
            self.repos, self.articulo_service, articulo=None, seccion_sugerida=seccion_sugerida, parent=self
        )
        dialogo.exec()
        if dialogo.guardado_o_eliminado:
            self._refrescar_lista()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)

    def done(self, result: int) -> None:  # noqa: N802 (Qt override)
        # `accept()`/`reject()` (modo_seleccion) llaman `done()` pero NO
        # disparan `closeEvent` — sin este override, la sesión propia del
        # diálogo quedaba abierta (fuga de conexión) cada vez que se usa
        # como selector. `Session.close()` es idempotente, no rompe nada
        # si además se llega a disparar `closeEvent`.
        self.db.close()
        super().done(result)
