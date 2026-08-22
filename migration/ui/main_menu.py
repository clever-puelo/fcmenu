"""Entry point de la aplicación completa — reemplaza a `FCMENU.frm`
(el `MDIForm` principal del legacy).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_menu
"""

from __future__ import annotations

import os
import sys

from PyQt6.QtWidgets import QApplication

from .main_menu_window import MainMenuWindow
from .theme import aplicar_tema


def _avisar_launcher_listo() -> None:
    """Le avisa a `launcher.py` (si fue el que arrancó este proceso,
    ver `VAR_ENTORNO_PID_LAUNCHER`) que la ventana real ya se mostró —
    a partir de acá el launcher cierra SU pantalla de bienvenida
    (pedido del usuario, 2026-08-22: "que elimina al llamador"). Si
    este proceso corre solo (dev, sin pasar por el launcher) la
    variable de entorno no está seteada y no hace nada."""
    # Import diferido a propósito: `launcher.py` es el módulo LIVIANO
    # (sólo PyQt6, ver su docstring) — importarlo acá arriba, a nivel
    # de módulo, no rompería nada funcionalmente, pero mezclaría sus
    # responsabilidades sin necesidad; este proceso (el programa
    # PESADO) sólo necesita 2 nombres sueltos de ahí, y sólo cuando
    # arrancó vía el launcher.
    from .launcher import VAR_ENTORNO_PID_LAUNCHER, sentinel_path

    pid_launcher = os.environ.get(VAR_ENTORNO_PID_LAUNCHER)
    if not pid_launcher:
        return
    try:
        sentinel_path(pid_launcher).touch()
    except OSError:
        pass  # sólo cosmético (cierra el splash del launcher un poco tarde) — no bloquea el arranque real


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)

    # Sin bienvenida propia acá — "fantasma" real encontrado a pedido
    # del usuario (2026-08-22, "me parece que quedó en el principal una
    # ventana de bienvenida que ahora es inútil"): este proceso ya
    # arranca vía `launcher.py` en el uso normal del .exe empaquetado, y
    # ESE es quien muestra la bienvenida real desde antes — una segunda
    # acá (aunque quedara tapada por la del lanzador, que es
    # `WindowStaysOnTopHint`) no aportaba nada, sólo una ventana de más
    # abriéndose y cerrándose de puro adorno. Corriendo standalone (dev,
    # `-m migration.ui.main_menu` directo, sin lanzador) tampoco hace
    # falta — quien quiera ver la bienvenida en dev puede correr `-m
    # migration.ui.launcher`, que la muestra igual (ver
    # `_comando_programa_principal()`, funciona también sin estar
    # congelado con PyInstaller).
    ventana = MainMenuWindow()
    ventana.show()
    _avisar_launcher_listo()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
