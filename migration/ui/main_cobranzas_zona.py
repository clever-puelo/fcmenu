"""Entry point standalone para "Cobranzas por Zona", migración de
`VerCobra.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_cobranzas_zona
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .cobranzas_zona_window import CobranzasZonaWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = CobranzasZonaWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
