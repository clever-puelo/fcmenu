"""Vista previa embebida de un PDF ya generado en disco — pedido del
usuario (2026-08-15) para el Recibo: "al oprimir emitir, muestre el
recibo (Pdf) en una ventana emergente y luego pregunte si se quiere
grabar" — el orden real queda Vista Previa → Confirmar → Grabar (el PDF
de vista previa se genera con el número que TENDRÍA el comprobante,
antes de persistir nada; ver `ReciboWindow._on_emitir`).

**Ajuste del mismo día**: el usuario probó ese flujo (Vista Previa + un
`QMessageBox.question` aparte al cerrarla) y pidió simplificarlo — el
botón "Grabar" vive directo acá (`mostrar_boton_grabar=True`), sin una
ventana de confirmación aparte. `self.grabar` queda en `True` sólo si
se clickeó "Grabar"; "Cerrar"/"Cancelar" no confirman nada.

Usa `QPdfView`/`QPdfDocument` (`PyQt6.QtPdfWidgets`/`PyQt6.QtPdf`, ya
incluidos en la instalación de PyQt6 del proyecto) — sin depender de un
visor externo del sistema operativo.
"""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView
from PyQt6.QtWidgets import QDialog, QHBoxLayout, QPushButton, QVBoxLayout, QWidget


class PdfPreviewDialog(QDialog):
    def __init__(
        self,
        ruta_pdf: Path,
        *,
        titulo: str = "Vista Previa",
        mostrar_boton_grabar: bool = False,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.resize(720, 900)
        self.grabar = False

        layout = QVBoxLayout(self)

        self._documento = QPdfDocument(self)
        self._documento.load(str(ruta_pdf))

        self._vista = QPdfView(self)
        self._vista.setDocument(self._documento)
        self._vista.setPageMode(QPdfView.PageMode.MultiPage)
        self._vista.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        layout.addWidget(self._vista, stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        if mostrar_boton_grabar:
            self.btn_grabar = QPushButton("Grabar")
            self.btn_grabar.setStyleSheet("font-weight: bold; padding: 8px;")
            self.btn_grabar.clicked.connect(self._on_grabar)
            fila_botones.addWidget(self.btn_grabar)
            self.btn_cerrar = QPushButton("Cancelar")
        else:
            self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.reject)
        fila_botones.addWidget(self.btn_cerrar)
        layout.addLayout(fila_botones)

    def _on_grabar(self) -> None:
        self.grabar = True
        self.accept()
