"""Entry point standalone para probar `NotaArticuloDialog` (Notartic.frm)
sin pasar por el ABM de Artículos completo.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_notas_articulo <COD1> <COD2>
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from migration.db import get_session
from migration.repository import RepositoryFactory

from .nota_articulo_dialog import NotaArticuloDialog
from .theme import aplicar_tema


def main() -> int:
    cod1 = sys.argv[1] if len(sys.argv) > 1 else "GPN"
    cod2 = sys.argv[2] if len(sys.argv) > 2 else "0300"
    app = QApplication(sys.argv)
    aplicar_tema(app)
    db = get_session()
    dialogo = NotaArticuloDialog(RepositoryFactory(db), cod1, cod2)
    dialogo.show()
    resultado = app.exec()
    db.close()
    return resultado


if __name__ == "__main__":
    sys.exit(main())
