"""Entry point standalone para la Consulta de Stock, migración de
`Verstock.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_stock_consulta
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .stock_consulta_window import StockConsultaWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = StockConsultaWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
