"""Entry point standalone para probar la Consulta de Cuenta Corriente,
migración de `CtaCte.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_ctacte
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .ctacte_window import CtaCteWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = CtaCteWindow()
    ventana.setWindowIcon(icono("ctacte", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
