"""Entry point standalone para "Listados Varios", migración de
`Listados.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_listados
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .listados_window import ListadosWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ListadosWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
