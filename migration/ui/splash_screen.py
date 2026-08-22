"""`SplashScreen` — pantalla de bienvenida mostrada al arrancar
`FcMenuII.exe`/`migration.ui.main_menu`, mientras se conecta a la base y
se arma la ventana principal.

Pedido del usuario (2026-08-21): "Pantalla negra cambiarla por una de
bienvenida similar a la de Acerca de con una barra de progreso para
distraer la demora de la carga". La "pantalla negra" es el placeholder
que pinta Windows para cualquier ventana que tarda en aparecer/responder
— acá el arranque real (`get_session()` conectando a Postgres +
`MainMenuWindow.__init__` armando toda la UI) es TODO síncrono y puede
tardar unos segundos, tiempo en el que no hay ninguna ventana real
todavía pintada. Mostrar esta pantalla PRIMERO (liviana, sin tocar la
base) le da al usuario algo real para mirar de entrada, y evita que
Windows dibuje su placeholder negro.

Misma identidad visual que `AcercaDeDialog` (panel superior verde
oscuro con logo+título+subtítulo) — "similar a la de Acerca de", pedido
explícito del usuario — con una barra de progreso INDETERMINADA (estilo
"actividad en curso", no un % real: no hay forma barata de medir cuánto
falta de la conexión a la base o del armado de la UI sin instrumentar
cada paso por dentro) + una línea de estado que se puede ir actualizando
con `actualizar_estado()` en los puntos conocidos del arranque."""

from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QLabel, QProgressBar, QVBoxLayout, QWidget

from .theme import Verde, logo_empresa

VERSION_APP = "2.0.0"


class SplashScreen(QWidget):
    def __init__(self) -> None:
        # `Qt.WindowType.SplashScreen` es el tipo de ventana nativo para
        # esto (sin bordes/título, siempre arriba, no aparece en la
        # barra de tareas como una ventana normal más) — mismo tipo que
        # usa `QSplashScreen` por dentro, pero acá se arma a mano sobre
        # `QWidget` para poder meterle una `QProgressBar` real adentro
        # (`QSplashScreen` sólo soporta un mensaje de texto superpuesto,
        # sin espacio para una barra de progreso de verdad).
        super().__init__(None, Qt.WindowType.SplashScreen | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedSize(420, 280)
        self.setStyleSheet(f"background-color: {Verde.BLANCO};")

        layout = QVBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 24)

        layout.addWidget(self._armar_panel_superior())
        layout.addStretch()

        self.lbl_estado = QLabel("Iniciando…")
        self.lbl_estado.setStyleSheet(f"color: {Verde.TEXTO_SUAVE}; font-size: 9.5pt;")
        self.lbl_estado.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_estado)

        layout.addSpacing(8)

        # Indeterminada (min=max=0, "modo ocupado" nativo de Qt: franja
        # animada en vez de un % real) — no hay dato confiable de cuánto
        # falta del arranque real, sólo el objetivo de "distraer la
        # demora" (pedido del usuario).
        self.barra = QProgressBar()
        self.barra.setRange(0, 0)
        self.barra.setTextVisible(False)
        self.barra.setFixedHeight(10)
        self.barra.setStyleSheet(
            f"""
            QProgressBar {{
                background-color: {Verde.CLARO};
                border: none;
                border-radius: 5px;
            }}
            QProgressBar::chunk {{
                background-color: {Verde.MEDIO};
                border-radius: 5px;
            }}
            """
        )
        contenedor_barra = QWidget()
        contenedor_layout = QVBoxLayout(contenedor_barra)
        contenedor_layout.setContentsMargins(50, 0, 50, 0)
        contenedor_layout.addWidget(self.barra)
        layout.addWidget(contenedor_barra)

        self._centrar_en_pantalla()

    def _armar_panel_superior(self) -> QWidget:
        # Réplica del panel superior de `AcercaDeDialog` (logo+título+
        # subtítulo sobre franja verde oscuro) — "similar a la de Acerca
        # de", pedido explícito del usuario.
        panel = QWidget()
        panel.setStyleSheet(f"background-color: {Verde.OSCURO};")
        layout = QVBoxLayout(panel)
        layout.setSpacing(4)
        layout.setContentsMargins(30, 28, 30, 24)
        layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        lbl_logo = QLabel()
        pixmap_logo = logo_empresa(72)
        if not pixmap_logo.isNull():
            lbl_logo.setPixmap(pixmap_logo)
            lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(lbl_logo)

        lbl_titulo = QLabel("FcMenu II")
        lbl_titulo.setStyleSheet(
            f"color: {Verde.BLANCO}; font-family: 'Arial Black'; font-size: 28pt; font-weight: bold;"
        )
        lbl_titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_titulo)

        lbl_subtitulo = QLabel("Facturación y Cuentas Corrientes")
        lbl_subtitulo.setStyleSheet(f"color: {Verde.PASTEL}; font-size: 11pt; font-weight: bold;")
        lbl_subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_subtitulo)

        lbl_version = QLabel(f"Versión {VERSION_APP}")
        lbl_version.setStyleSheet(f"color: {Verde.MEDIO_CLARO}; font-size: 9pt;")
        lbl_version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_version)

        return panel

    def _centrar_en_pantalla(self) -> None:
        pantalla = QApplication.primaryScreen()
        if pantalla is None:
            return
        geometria = pantalla.availableGeometry()
        x = geometria.x() + (geometria.width() - self.width()) // 2
        y = geometria.y() + (geometria.height() - self.height()) // 2
        self.move(x, y)

    def actualizar_estado(self, texto: str) -> None:
        """Cambia la línea de estado y fuerza un repintado inmediato —
        `main()` la llama en cada paso conocido del arranque (conectando
        a la base, armando la ventana principal, etc.) para que el
        usuario vea que algo está pasando, no sólo la barra indeterminada
        girando sola."""
        self.lbl_estado.setText(texto)
        QApplication.processEvents()
