"""Entry point standalone para probar el ABM de Artículos completo
(Búsqueda -> Detalle), migración de Busqueda.frm + AbmArt.frm.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_articulo
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .articulo_busqueda_window import ArticuloBusquedaWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ArticuloBusquedaWindow()
    ventana.setWindowIcon(icono("articulos", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
