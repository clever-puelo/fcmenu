"""Entry point de la aplicación completa — reemplaza a `FCMENU.frm`
(el `MDIForm` principal del legacy).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_menu
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .main_menu_window import MainMenuWindow
from .splash_screen import SplashScreen
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)

    # Pantalla de bienvenida con barra de progreso (pedido del usuario,
    # 2026-08-21: "Pantalla negra cambiarla por una de bienvenida... para
    # distraer la demora de la carga") — se muestra ANTES de conectar a
    # la base/armar la ventana principal (`MainMenuWindow.__init__` es
    # síncrono y puede tardar unos segundos: conexión a Postgres +
    # `RepositoryFactory` + toda la UI), que es justo el hueco donde
    # antes no había ninguna ventana real pintada todavía.
    splash = SplashScreen()
    splash.show()
    splash.actualizar_estado("Conectando con la base de datos…")

    ventana = MainMenuWindow()

    splash.actualizar_estado("Listo")
    ventana.show()
    splash.close()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
