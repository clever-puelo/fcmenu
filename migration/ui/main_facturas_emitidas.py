"""Entry point standalone para "Facturas Emitidas", migración de
`VerFact.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_facturas_emitidas
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .facturas_emitidas_window import FacturasEmitidasWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = FacturasEmitidasWindow()
    ventana.setWindowIcon(icono("facturas_emitidas", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
