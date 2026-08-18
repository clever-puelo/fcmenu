"""Entry point standalone para "Listados Varios", migración de
`Listados.frm`.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_listados
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .listados_window import ListadosWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    # `reporte_inicial` ahora es obligatorio (ya no hay combo adentro de
    # la ventana para elegirlo/cambiarlo, ver docstring de
    # ListadosWindow) — "Clientes" como default razonable para probar
    # standalone.
    ventana = ListadosWindow(reporte_inicial="clientes")
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
