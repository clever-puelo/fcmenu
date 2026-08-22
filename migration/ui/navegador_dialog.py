"""Navegador embebido (Chromium vía `QWebEngineView`) — pedido del
usuario (2026-08-21): el botón "Ver en el Mapa" de la ficha de Cliente
tiene que abrir "una ventana y dentro la web... (un navegador básico
dentro del programa)", **no** ir al navegador externo del sistema (antes
usaba `QDesktopServices.openUrl`, que sí lo hacía).

Básico a propósito (barra de Atrás/Adelante/Recargar + URL de sólo
lectura) — no es un navegador de uso general, sólo la ventana mínima
para que el operador vea el mapa sin salir de la app.

**Ventana chica a propósito** (pedido del usuario, 2026-08-21: "la
ventana es muy grande, reducila a un cuarto de ese tamaño") — esto es un
preview puntual de un mapa, no un navegador de uso general que necesite
ocupar media pantalla."""

from __future__ import annotations

from typing import Optional

from PyQt6.QtCore import QUrl
from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLineEdit, QPushButton, QVBoxLayout, QWidget

from .widgets import redimensionar_pct_pantalla


class NavegadorDialog(QDialog):
    """Ventana con un `QWebEngineView` adentro — navegador mínimo para
    contenido embebido (mapas, etc.), sin salir de la app. Requiere
    `PyQt6-WebEngine` instalado (ver `requirements.txt`).

    `html`, si se pasa, se carga con `setHtml()` en vez de navegar a
    `url` directo (`load()`) — necesario para el embed de Google Maps
    (`output=embed`): la propia página JS de Google chequea que corra
    DENTRO de un `<iframe>` (`window.self !== window.top`) y si no,
    rechaza con "The google maps embed API must be used in an iframe" en
    vez de mostrar el mapa — bug real reportado por el usuario
    (2026-08-21), reproducido al navegar `output=embed` directo con
    `load()` (eso es SIEMPRE el documento de nivel superior, nunca un
    iframe). `_abrir_mapa` (`cliente_detalle_dialog.py`) arma un HTML
    envoltorio mínimo con el mapa DENTRO de un iframe de verdad y lo pasa
    acá."""

    def __init__(
        self,
        url: str = "",
        titulo: str = "Navegador",
        *,
        html: Optional[str] = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        # Import diferido: `PyQt6.QtWebEngineWidgets` es un paquete aparte
        # (`PyQt6-WebEngine`) bastante más pesado que el resto de PyQt6 —
        # se importa recién acá, al abrir esta ventana puntual, en vez de
        # cargarlo en cada arranque de la app.
        from PyQt6.QtWebEngineWidgets import QWebEngineView

        self.setWindowTitle(titulo)
        # Un cuarto del tamaño anterior (70%x80% de la pantalla, área
        # 0,56) — mitad de cada lado da un cuarto del área (35%x40% = 0,14).
        redimensionar_pct_pantalla(self, 35, 40)

        layout = QVBoxLayout(self)

        barra = QHBoxLayout()
        self.btn_atras = QPushButton("◀")
        self.btn_atras.setMaximumWidth(36)
        barra.addWidget(self.btn_atras)
        self.btn_adelante = QPushButton("▶")
        self.btn_adelante.setMaximumWidth(36)
        barra.addWidget(self.btn_adelante)
        self.btn_recargar = QPushButton("⟳")
        self.btn_recargar.setMaximumWidth(36)
        barra.addWidget(self.btn_recargar)

        self.txt_url = QLineEdit()
        self.txt_url.setReadOnly(True)
        barra.addWidget(self.txt_url, stretch=1)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        barra.addWidget(btn_cerrar)

        layout.addLayout(barra)

        self.vista = QWebEngineView()
        layout.addWidget(self.vista, stretch=1)

        self.btn_atras.clicked.connect(self.vista.back)
        self.btn_adelante.clicked.connect(self.vista.forward)
        self.btn_recargar.clicked.connect(self.vista.reload)
        if html is not None:
            # El wrapper propio (`about:blank`/vacío) dispara `urlChanged`
            # con una URL sin sentido para mostrar — la barra de URL
            # queda mostrando la dirección REAL (la del iframe adentro),
            # fija, no la del documento contenedor.
            self.txt_url.setText(url)
            self.vista.setHtml(html)
        else:
            self.vista.urlChanged.connect(lambda qurl: self.txt_url.setText(qurl.toString()))
            self.vista.load(QUrl(url))
