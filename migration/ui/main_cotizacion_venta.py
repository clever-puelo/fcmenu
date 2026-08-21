"""Entry point standalone para probar la Cotización (comprobante sin CAE
del Facturador, `CabFact.frm TipoFac=4`) — NO confundir con
`main_cotizacion.py` (Cotización del Dólar).

Uso:
    .venv/Scripts/python.exe -m migration.ui.main_cotizacion_venta
"""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from .cotizacion_venta_window import CotizacionVentaWindow
from .iconos import icono
from .theme import aplicar_tema


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)
    ventana = CotizacionVentaWindow()
    ventana.setWindowIcon(icono("cotizacion_venta", 32))
    ventana.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
