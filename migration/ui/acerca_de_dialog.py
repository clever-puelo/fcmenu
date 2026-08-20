"""`AcercaDeDialog` — "Acerca de FCMENU II", réplica en Python de la
ventana de presentación real del legacy (`Inicio.frm`: título "FCMENU",
subtítulo "Facturación y Cuentas Corrientes", autor "Clever G.
Colalillo", logo, cartel de entorno).

Pedido del usuario (2026-08-17): la sección "Ayuda" del sidebar tiene
un botón "Acerca de FCMENU II" que "apunte a la ventana de
presentación" — la única pantalla de presentación real que existe en
el proyecto legacy es `Inicio.frm` (el `MDIForm.Show` inicial de
`FCMENU.frm`, nunca migrado como tal porque no tenía lógica propia,
sólo texto). Acá se migra su contenido como un diálogo modal
informativo, no como splash de arranque.

**Diferencia a propósito con el legacy**: `Inicio.frm` tenía el texto
"PRODUCCIÓN" (`Label6`) hardcodeado sin importar el entorno real. Acá
se usa `migration.afip.etiqueta_entorno_afip()` — el mismo dato dinámico
que ya muestra el cartel de la cabecera de `MainMenuWindow` — para que
"Acerca de" nunca mienta sobre en qué entorno de AFIP está corriendo.

**Panel superior verde oscuro** (feedback del usuario, 2026-08-18): el
logo real (`assets/Logo-Alestel.png`) es muy claro, se perdía contra el
fondo blanco del diálogo — ahora logo+título+subtítulo viven sobre una
franja `Verde.OSCURO` (mismo tono que la cabecera de `MainMenuWindow`),
el resto del diálogo (entorno/versión/autor/Cerrar) sigue con el fondo
claro normal.
"""

from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QDialog, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.afip import etiqueta_entorno_afip

from .theme import Verde, logo_empresa

VERSION_APP = "2.0.0"
AUTOR = "Clever G. Colalillo"


class AcercaDeDialog(QDialog):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Acerca de FcMenu II")
        self.setFixedSize(420, 420)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(0, 0, 0, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        layout.addWidget(self._armar_panel_superior())

        lbl_entorno = QLabel(etiqueta_entorno_afip())
        lbl_entorno.setStyleSheet(
            f"""
            color: {Verde.OSCURO};
            background-color: {Verde.PASTEL};
            border: 1px solid {Verde.MEDIO_CLARO};
            border-radius: 8px;
            font-size: 12pt;
            font-weight: bold;
            padding: 6px 18px;
            """
        )
        lbl_entorno.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(14)
        layout.addWidget(lbl_entorno, alignment=Qt.AlignmentFlag.AlignHCenter)

        layout.addSpacing(14)

        lbl_version = QLabel(f"Versión {VERSION_APP}")
        lbl_version.setStyleSheet(f"color: {Verde.TEXTO_SUAVE}; font-size: 9pt;")
        lbl_version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_version)

        lbl_autor = QLabel(AUTOR)
        lbl_autor.setStyleSheet(f"color: {Verde.TEXTO_SUAVE}; font-size: 9pt;")
        lbl_autor.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_autor)

        layout.addStretch()

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar)

    def _armar_panel_superior(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet(f"background-color: {Verde.OSCURO};")
        layout = QVBoxLayout(panel)
        layout.setSpacing(4)
        layout.setContentsMargins(30, 24, 30, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        lbl_logo = QLabel()
        pixmap_logo = logo_empresa(72)
        if not pixmap_logo.isNull():
            lbl_logo.setPixmap(pixmap_logo)
            lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(lbl_logo)

        lbl_titulo = QLabel("FcMenu II")
        lbl_titulo.setStyleSheet(
            f"color: {Verde.BLANCO}; font-family: 'Arial Black'; font-size: 30pt; font-weight: bold;"
        )
        lbl_titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_titulo)

        lbl_subtitulo = QLabel("Facturación y Cuentas Corrientes")
        lbl_subtitulo.setStyleSheet(f"color: {Verde.PASTEL}; font-size: 11pt; font-weight: bold;")
        lbl_subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_subtitulo)

        lbl_electronica = QLabel("(Facturación Electrónica — AFIP)")
        lbl_electronica.setStyleSheet(f"color: {Verde.MEDIO_CLARO}; font-size: 9.5pt;")
        lbl_electronica.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_electronica)

        return panel
