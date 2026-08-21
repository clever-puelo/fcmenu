"""Entry point standalone para probar la Nota de Crédito Interna
(`NCInterna.frm`).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_nota_credito_interna
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .iconos import icono
from .nota_credito_interna_window import NotaCreditoInternaWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = NotaCreditoInternaWindow()
    ventana.setWindowIcon(icono("nci", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
