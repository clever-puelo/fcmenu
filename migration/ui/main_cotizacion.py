"""Entry point standalone para probar `CotizacionWindow` (Cotizac.frm).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_cotizacion
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .cotizacion_window import CotizacionWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = CotizacionWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
