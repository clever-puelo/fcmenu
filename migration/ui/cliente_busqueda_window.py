"""Ventana de Búsqueda de Clientes — punto de entrada del ABM, y también
selector reutilizable para el Facturador.

Convención de navegación confirmada con el usuario (2026-08-15), aplicada
a todo el sistema salvo aclaración específica: primero se pide un
criterio de búsqueda, se muestra la lista de resultados abajo, y recién
al elegir uno de la lista se abre el diálogo de Detalle
(`ClienteDetalleDialog`) para Ver/Cambio/Baja. El botón Alta, abajo a la
izquierda, abre el mismo diálogo de Detalle pero en blanco.

Para Clientes no hay un único campo "categoría" como la Sección de
Artículos, así que el criterio de búsqueda es Código exacto o Nombre
parcial (igual que ofrecía `ClienteService.buscar`); sin filtro se listan
los clientes activos.

La grilla (`TablaBusqueda`) permite ordenar por cualquier columna
clickeando su título (confirmado 2026-08-16, convención de sistema para
toda búsqueda). Elegir una fila es con un clic o doble clic (feedback del
usuario, 2026-08-15 — reemplaza la decisión previa de exigir doble clic
para no disparar el diálogo "por accidente" al ordenar; ordenar sigue
siendo un clic en el TÍTULO de columna, no en una fila, así que no hay
conflicto real).

**Búsqueda** (feedback del usuario, 2026-08-15): el foco arranca en
"Razón Social" (antes "Nombre contiene") y filtra en vivo a medida que
se tipea (`LIKE '%dato%'`), sin botón "Buscar". El campo "Código" es
aparte: al confirmarlo (Enter/Tab) se busca ese código exacto — si
existe, se lo elige directo (en modo selección, cierra el diálogo con
ese Cliente, igual que un doble clic en la lista); si no existe, se
avisa y se lista igual, ordenada por Código de menor a mayor arrancando
desde el número tipeado (en vez de una lista vacía, ayuda a ubicar el
código real más cercano).

**Modo selección** (2026-08-16, para `CabFact.frm`: variable global
`CodCLTEf`, que en el legacy se llenaba abriendo `BusClte.frm`): con
`modo_seleccion=True` esta misma ventana se usa como `QDialog` modal
(`.exec()`), sin botón Alta (elegir cliente para una factura no es dar de
alta catálogo), y elegir una fila deja el cliente en
`self.cliente_elegido` y cierra con `Accepted` en vez de abrir
`ClienteDetalleDialog`. El uso normal (standalone, `main.py`) no cambia:
`QDialog` admite `.show()` no-modal igual que antes se usaba
`QMainWindow`.
"""

from __future__ import annotations

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import (
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.repository import RepositoryFactory
from migration.services import ClienteService

from .cliente_detalle_dialog import ClienteDetalleDialog
from .widgets import TablaBusqueda, UpperCaseLineEdit


class ClienteBusquedaWindow(QDialog):
    """Punto de entrada del ABM de Clientes (equivalente conceptual a
    BusClte.frm + Abmclte.frm del legacy), y selector modal reutilizable
    cuando `modo_seleccion=True` (ver docstring del módulo)."""

    def __init__(self, parent: QWidget | None = None, *, modo_seleccion: bool = False):
        super().__init__(parent)
        self.modo_seleccion = modo_seleccion
        self.cliente_elegido = None
        self.setWindowTitle("Selección de Cliente" if modo_seleccion else "Búsqueda de Clientes")
        self.resize(700, 520)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.cliente_service = ClienteService(self.db)

        self._construir_ui()
        self._refrescar_lista()
        self.txt_nombre.setFocus()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        layout.addLayout(self._armar_fila_filtros())

        grupo_lista = QGroupBox("Clientes")
        layout_lista = QVBoxLayout(grupo_lista)
        self.tabla = TablaBusqueda(["Código", "Razón Social", "Localidad", "Teléfono"])
        # Selección con un clic o doble clic (feedback del usuario, 2026-08-15).
        self.tabla.itemClicked.connect(lambda _: self._on_elegir_de_lista())
        self.tabla.itemDoubleClicked.connect(lambda _: self._on_elegir_de_lista())
        self.tabla.itemActivated.connect(lambda _: self._on_elegir_de_lista())
        layout_lista.addWidget(self.tabla)
        layout.addWidget(grupo_lista, stretch=1)

        layout.addLayout(self._armar_barra_botones())

    def _armar_fila_filtros(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        fila.addWidget(QLabel("Código:"))
        self.spin_codigo = QSpinBox()
        self.spin_codigo.setRange(0, 999999)
        # Código exacto: se confirma con Enter/Tab (no busca a cada
        # dígito, sería ruidoso con un spinner) — ver `_on_codigo_confirmado`.
        self.spin_codigo.editingFinished.connect(self._on_codigo_confirmado)
        fila.addWidget(self.spin_codigo)

        fila.addWidget(QLabel("Razón Social:"))
        self.txt_nombre = UpperCaseLineEdit()
        # Filtro en vivo ('%dato%') a medida que se tipea — sin botón
        # "Buscar" (feedback del usuario, 2026-08-15). Debounce de 250ms
        # para no pegarle a la base de datos en cada tecla.
        self._timer_filtro_nombre = QTimer(self)
        self._timer_filtro_nombre.setSingleShot(True)
        self._timer_filtro_nombre.timeout.connect(self._refrescar_lista)
        self.txt_nombre.textChanged.connect(lambda _: self._timer_filtro_nombre.start(250))
        fila.addWidget(self.txt_nombre)

        return fila

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        if not self.modo_seleccion:
            self.btn_alta = QPushButton("Alta")
            self.btn_alta.setToolTip("Dar de alta un cliente nuevo")
            self.btn_alta.clicked.connect(self._on_alta)
            fila.addWidget(self.btn_alta)

        fila.addStretch()

        self.btn_cerrar = QPushButton("Cancelar" if self.modo_seleccion else "Cerrar")
        self.btn_cerrar.clicked.connect(self.reject)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    def _cargar_resultados(self, resultados) -> None:
        filas = [
            (
                [str(c.CODIGO), c.NOMB or "", c.LOC or "", c.TEL1 or ""],
                c,
            )
            for c in resultados
        ]
        self.tabla.cargar_filas(filas)

    def _refrescar_lista(self) -> None:
        codigo = self.spin_codigo.value() or None
        nombre = self.txt_nombre.text().strip() or None

        resultados = self.cliente_service.buscar(codigo=codigo, nombre=nombre)
        self._cargar_resultados(resultados)

    def _on_codigo_confirmado(self) -> None:
        """Código exacto tipeado en el campo "Código" y confirmado
        (Enter/Tab): si existe, lo elige directo y sigue con el
        Facturador (o abre su Detalle en el ABM standalone) — igual que
        elegirlo de la lista. Si no existe, avisa y muestra la lista de
        Clientes ordenada por Código de menor a mayor arrancando desde
        el número tipeado (feedback del usuario, 2026-08-15)."""
        codigo = self.spin_codigo.value()
        if codigo <= 0:
            return
        cliente = self.repos.cliente().by_codigo(codigo)
        self.spin_codigo.setValue(0)  # no debe quedar "tapando" el filtro por Razón Social
        if cliente is not None:
            self._elegir_cliente(cliente)
            return
        QMessageBox.information(self, "Búsqueda de Clientes", f"No existe el cliente {codigo}.")
        self._cargar_resultados(self.repos.cliente().desde_codigo(codigo))

    def _elegir_cliente(self, cliente) -> None:
        if self.modo_seleccion:
            self.cliente_elegido = cliente
            self.accept()
            return

        dialogo = ClienteDetalleDialog(self.repos, self.cliente_service, cliente=cliente, parent=self)
        dialogo.exec()
        if dialogo.guardado_o_eliminado:
            self._refrescar_lista()

    def _on_elegir_de_lista(self) -> None:
        cliente = self.tabla.objeto_seleccionado()
        if cliente is None:
            return
        self._elegir_cliente(cliente)

    def _on_alta(self) -> None:
        dialogo = ClienteDetalleDialog(self.repos, self.cliente_service, cliente=None, parent=self)
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
