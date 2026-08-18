"""Entry point standalone para "Arreglos > Cuenta Corriente", migración
de `CargaCC.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_arreglo_ctacte
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .arreglo_ctacte_window import ArregloCtaCteWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ArregloCtaCteWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
