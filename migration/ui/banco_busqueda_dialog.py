"""Búsqueda de Banco por nombre — se abre desde `PagoDialog` cuando el
operador deja el campo "Banco" (código) en blanco al cargar un Cheque o
una Tarjeta (feedback del usuario, 2026-08-15: antes, dejar el código en
blanco disparaba un intento de búsqueda por código que daba un error
inconsistente — acá se ofrece directamente buscar por nombre).

No hay un ABM de Bancos migrado todavía (`ABMBcos.frm` queda pendiente)
— esto es sólo un selector liviano de sólo lectura, mismo criterio de
navegación que el resto de la migración: un clic o doble clic elige y
cierra (confirmado 2026-08-15 para toda grilla de selección).
"""

from __future__ import annotations

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QDialog, QGroupBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda, UpperCaseLineEdit


class BancoBusquedaDialog(QDialog):
    def __init__(self, repos: RepositoryFactory, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.banco_elegido = None

        self.setWindowTitle("Búsqueda de Banco")
        self.resize(480, 420)

        self._construir_ui()
        self._refrescar_lista()
        # Foco en el primer campo al abrir (convención de sistema,
        # 2026-08-15) — diferido porque el diálogo todavía no está
        # mostrado en este punto del __init__.
        QTimer.singleShot(0, self.txt_nombre.setFocus)

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        fila = QHBoxLayout()
        fila.addWidget(QLabel("Nombre contiene:"))
        self.txt_nombre = UpperCaseLineEdit()
        self._timer_filtro = QTimer(self)
        self._timer_filtro.setSingleShot(True)
        self._timer_filtro.timeout.connect(self._refrescar_lista)
        self.txt_nombre.textChanged.connect(lambda _: self._timer_filtro.start(250))
        fila.addWidget(self.txt_nombre)
        layout.addLayout(fila)

        grupo = QGroupBox("Bancos")
        layout_grupo = QVBoxLayout(grupo)
        self.tabla = TablaBusqueda(["Código", "Nombre", "Localidad"])
        self.tabla.itemClicked.connect(lambda _: self._on_elegir_de_lista())
        self.tabla.itemDoubleClicked.connect(lambda _: self._on_elegir_de_lista())
        self.tabla.itemActivated.connect(lambda _: self._on_elegir_de_lista())
        layout_grupo.addWidget(self.tabla)
        layout.addWidget(grupo, stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        self.btn_cancelar = QPushButton("Cancelar")
        self.btn_cancelar.clicked.connect(self.reject)
        fila_botones.addWidget(self.btn_cancelar)
        layout.addLayout(fila_botones)

    # ------------------------------------------------------------------
    def _refrescar_lista(self) -> None:
        texto = self.txt_nombre.text().strip()
        bancos = self.repos.banco().buscar_por_nombre(texto) if texto else self.repos.banco().todos()
        filas = [
            ([str(b.COD), (b.NOMBRE or "").strip(), (b.LOC or "").strip()], b)
            for b in bancos
        ]
        self.tabla.cargar_filas(filas)

    def _on_elegir_de_lista(self) -> None:
        banco = self.tabla.objeto_seleccionado()
        if banco is None:
            return
        self.banco_elegido = banco
        self.accept()
