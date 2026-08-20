"""Entry point standalone para la Consulta de Despachos, migración de
`VerDesp.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_despachos_consulta
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .despachos_consulta_window import DespachosConsultaWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = DespachosConsultaWindow()
    ventana.setWindowIcon(icono("despachos", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
