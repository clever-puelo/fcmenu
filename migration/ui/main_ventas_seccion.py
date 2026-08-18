"""Entry point standalone para "Ventas de Artículos por Cliente",
migración de `ESTADIST.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_ventas_seccion
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .ventas_seccion_window import VentasSeccionWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = VentasSeccionWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
