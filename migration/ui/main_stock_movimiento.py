"""Entry point standalone para el Ingreso de Movimientos al Stock,
migración de `Stock.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_stock_movimiento
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .iconos import icono
from .stock_movimiento_window import StockMovimientoWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = StockMovimientoWindow()
    ventana.setWindowIcon(icono("ingreso_stock", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
