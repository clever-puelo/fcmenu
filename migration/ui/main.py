"""Entry point standalone para probar el ABM de Clientes completo
(Búsqueda -> Detalle), migración de BusClte.frm + Abmclte.frm.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .cliente_busqueda_window import ClienteBusquedaWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ClienteBusquedaWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
