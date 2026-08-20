"""`ReciboDetalleDialog` — ventana flotante de sólo lectura para ver el
detalle de un Recibo, abierta desde el drill-down del extracto de Cuenta
Corriente (pedido del usuario, 2026-08-20: "similar a la de facturas
para ver el detalle de recibos en la cuenta corriente").

Mismo patrón simple de `cheque_detalle_dialog.py` — varios `QGroupBox`
apilados, sólo lectura, un botón "Cerrar". Toda la composición del
detalle (cabecera, comprobantes aplicados, formas de pago) vive en
`CuentaCorrienteService.detalle_recibo()`, esta ventana sólo la muestra.

**Sin botón "Imprimir"** en esta primera versión, a diferencia del
detalle de factura (`FacturaEmitidaDetalleDialog`) — el pedido del
usuario fue "ver el detalle", no reimprimir; se puede agregar después
reusando `migration.pdf.generar_pdf_recibo` si hace falta.

**Comprobantes Aplicados de sólo lectura, sin drill-down**: el schema de
`Imputacion` no guarda LETRA/PTOVTA del comprobante original (sólo
`CPBTEI`, ver docstring de `Imputacion` en `models.py`), así que no
alcanza para volver a abrir el detalle de esa factura desde acá sin
ambigüedad.
"""

from __future__ import annotations

from PyQt6.QtWidgets import QDialog, QGroupBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.decimals import format_decimal
from migration.services import CuentaCorrienteService

from .widgets import TablaBusqueda, texto_contacto_cliente

COLUMNAS_APLICACIONES = ["Tipo", "Número", "Fecha", "Importe Aplicado"]
COLUMNAS_FORMAS_PAGO = ["Forma", "Detalle", "Importe"]


class ReciboDetalleDialog(QDialog):
    def __init__(self, cc_service: CuentaCorrienteService, clte: int, numero: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle(f"Recibo Nº {numero}")
        self.resize(640, 620)

        self.detalle = cc_service.detalle_recibo(clte, numero)
        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        if self.detalle is None:
            layout.addWidget(QLabel("No se encontró el detalle de este Recibo."))
            fila_botones = QHBoxLayout()
            fila_botones.addStretch()
            btn_cerrar = QPushButton("Cerrar")
            btn_cerrar.clicked.connect(self.accept)
            fila_botones.addWidget(btn_cerrar)
            layout.addLayout(fila_botones)
            return

        layout.addWidget(self._armar_cliente())
        layout.addWidget(self._armar_recibo())
        layout.addWidget(self._armar_aplicaciones(), stretch=1)
        layout.addWidget(self._armar_formas_pago(), stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _armar_cliente(self) -> QGroupBox:
        cliente = self.detalle.cliente
        grupo = QGroupBox("Cliente")
        layout = QVBoxLayout(grupo)
        lbl_nombre = QLabel(f"{cliente.CODIGO}  {(cliente.NOMB or '').strip()}")
        lbl_nombre.setStyleSheet("font-weight: bold;")
        layout.addWidget(lbl_nombre)
        layout.addWidget(QLabel(texto_contacto_cliente(cliente)))
        return grupo

    def _armar_recibo(self) -> QGroupBox:
        d = self.detalle
        grupo = QGroupBox("Recibo")
        layout = QHBoxLayout(grupo)

        for etiqueta, valor in (
            ("Número :", str(d.numero)),
            ("Fecha :", d.fecha.strftime("%d/%m/%Y") if d.fecha else ""),
            ("Total :", f"$ {format_decimal(d.total)}"),
        ):
            lbl_etiqueta = QLabel(etiqueta)
            lbl_etiqueta.setStyleSheet("font-weight: bold;")
            layout.addWidget(lbl_etiqueta)
            layout.addWidget(QLabel(valor))
            layout.addSpacing(16)
        if d.descuento:
            layout.addWidget(QLabel("Descuento :"))
            layout.addWidget(QLabel(f"$ {format_decimal(d.descuento)}"))
            layout.addSpacing(16)
        if d.anticipo:
            layout.addWidget(QLabel("Anticipo :"))
            layout.addWidget(QLabel(f"$ {format_decimal(d.anticipo)}"))
        layout.addStretch()
        return grupo

    def _armar_aplicaciones(self) -> QGroupBox:
        grupo = QGroupBox("Comprobantes Aplicados")
        layout = QVBoxLayout(grupo)
        tabla = TablaBusqueda(COLUMNAS_APLICACIONES, columnas_derecha=(3,))
        filas = [
            (
                [a.tipo_label, a.numero, a.fecha.strftime("%d/%m/%Y") if a.fecha else "", f"$ {format_decimal(a.importe)}"],
                a,
            )
            for a in self.detalle.aplicaciones
        ]
        tabla.cargar_filas(filas)
        layout.addWidget(tabla)
        return grupo

    def _armar_formas_pago(self) -> QGroupBox:
        grupo = QGroupBox("Formas de Pago")
        layout = QVBoxLayout(grupo)
        tabla = TablaBusqueda(COLUMNAS_FORMAS_PAGO, columnas_derecha=(2,))
        filas = [
            ([f.forma, f.detalle, f"$ {format_decimal(f.importe)}"], f)
            for f in self.detalle.formas_pago
        ]
        tabla.cargar_filas(filas)
        layout.addWidget(tabla)
        return grupo
