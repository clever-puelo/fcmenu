"""Entry point standalone para probar `ParametrosWindow` (Paramet.frm).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_parametros
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .parametros_window import ParametrosWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ParametrosWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
