"""Entry point de la aplicación completa — reemplaza a `FCMENU.frm`
(el `MDIForm` principal del legacy).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_menu
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .main_menu_window import MainMenuWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = MainMenuWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
