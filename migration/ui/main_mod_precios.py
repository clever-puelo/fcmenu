"""Entry point standalone para probar `ModPreciosWindow` (ModPrec.frm).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_mod_precios
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .iconos import icono
from .mod_precios_window import ModPreciosWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ModPreciosWindow()
    ventana.setWindowIcon(icono("precios", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
