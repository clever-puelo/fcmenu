"""Entry point standalone para probar el Facturador completo
(Cabecera/Detalle/Pie/Emitir), migración de CabFact.frm+DetFact.frm+
PieFact.frm+EmiFact.frm.

Usa `AfipWSFEv1Stub` por defecto (sin credenciales de Homologación
todavía, ver `migration/afip.py`) — las facturas emitidas acá NO tienen
CAE real.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_facturador
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .facturador_window import FacturadorWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = FacturadorWindow()
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
