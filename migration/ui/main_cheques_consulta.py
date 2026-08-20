"""Entry point standalone para la Consulta de Cheques, migración de
`VerCheq.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_cheques_consulta
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .cheques_consulta_window import ChequesConsultaWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ChequesConsultaWindow()
    ventana.setWindowIcon(icono("cheques", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
