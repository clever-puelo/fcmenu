"""Entry point standalone para probar `DtosClienteDialog` (DtosxClte.frm)
sin pasar por el ABM de Clientes completo.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_dtos_cliente <codigo_cliente>
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from migration.db import get_session
from migration.repository import RepositoryFactory

from .dtos_cliente_dialog import DtosClienteDialog
from .theme import aplicar_tema


def main() -> int:
    codigo = int(sys.argv[1]) if len(sys.argv) > 1 else 26
    app = QApplication(sys.argv)
    aplicar_tema(app)
    db = get_session()
    repos = RepositoryFactory(db)
    cliente = repos.cliente().by_codigo(codigo)
    nombre = cliente.NOMB if cliente else ""
    dialogo = DtosClienteDialog(repos, codigo, nombre or "")
    dialogo.show()
    resultado = app.exec()
    db.close()
    return resultado


if __name__ == "__main__":
    sys.exit(main())
