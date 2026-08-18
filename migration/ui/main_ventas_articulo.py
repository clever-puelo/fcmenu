"""Entry point standalone para "Ventas por Artículo", migración de
`VTAXART.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_ventas_articulo
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .ventas_articulo_window import VentasArticuloWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = VentasArticuloWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
