"""Entry point standalone para "Totales Diarios", migración de
`TotFact.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_totales_diarios
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .totales_diarios_window import TotalesDiariosWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = TotalesDiariosWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
