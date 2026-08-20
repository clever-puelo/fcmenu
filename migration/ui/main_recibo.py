"""Entry point standalone para probar el Recibo completo
(Cabecera/Pendientes de Cobro/Pagos/Pie/Emitir), migración de
`CabRec.frm`+`DetRec.frm`+`DetPago.frm`+`PieRec.frm`+`EmiRec.frm`.

Sin AFIP: un Recibo no es comprobante fiscal electrónico (ver docstring
de `EmisionReciboService`), no hace falta ningún stub/credencial acá.

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_recibo
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .iconos import icono
from .recibo_window import ReciboWindow
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = ReciboWindow()
    ventana.setWindowIcon(icono("recibo", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
