"""Entry point standalone para "Arreglos > Subdiario Ventas", migración
de `CargaFC.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_arreglo_subdiario
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .arreglo_subdiario_window import ArregloSubdiarioWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ArregloSubdiarioWindow()
    ventana.setWindowIcon(icono("arreglos", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
