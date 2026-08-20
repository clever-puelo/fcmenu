"""Entry point standalone para probar `TablasWindow` (ABMTablas.frm + Acttabla.frm).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_tablas
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .iconos import icono
from .tablas_window import TablasWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = TablasWindow()
    ventana.setWindowIcon(icono("tablas", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
