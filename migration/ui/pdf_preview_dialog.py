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

**Zoom + Imprimir** (feedback del usuario, 2026-08-18, tercera ronda:
"Los pdf tienen que tener alguna manera de aumentar o disminuir zoom.
Poder imprimir 1, varias o todas las páginas") — barra propia con
Zoom −/+ / Ajustar Ancho e "Imprimir..." (abre el diálogo nativo de
impresión de Windows, que ya deja elegir Todas/Selección/Rango de
páginas — no hace falta reinventar ese selector).

**Guardar como...** (feedback del usuario, 2026-08-19, pedido para el
"Imprimir" de un comprobante ya emitido: "permita al operador guardar o
imprimir") — copia el PDF ya generado a la carpeta/nombre que el
operador elija (`QFileDialog.getSaveFileName` + copia de archivo, el
PDF de origen siempre queda intacto en `migration/pdf_output/`).

Usa `QPdfView`/`QPdfDocument` (`PyQt6.QtPdfWidgets`/`PyQt6.QtPdf`, ya
incluidos en la instalación de PyQt6 del proyecto) — sin depender de un
visor externo del sistema operativo.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from PyQt6.QtCore import QEvent, QObject, QPoint, QSize, Qt, QTimer
from PyQt6.QtGui import QPainter
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView
from PyQt6.QtPrintSupport import QPrintDialog, QPrinter
from PyQt6.QtWidgets import QDialog, QFileDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton, QVBoxLayout, QWidget

# Pasos de zoom fijos (feedback del usuario: "aumentar o disminuir
# zoom") — más simple y predecible que un cálculo continuo por
# porcentaje libre; alcanza para leer cómodo un comprobante en
# cualquier tamaño de pantalla.
PASOS_ZOOM = [0.5, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0]


class _VistaPdfConManito(QPdfView):
    """`QPdfView` con pan-por-arrastre estilo "manito" — pedido del
    usuario (2026-08-21): "aplicar la manito para mover el pdf ampliado
    dentro de la ventana" (mismo gesto de cualquier visor de PDF real:
    con el documento con zoom, arrastrar con el mouse para moverlo en
    vez de depender sólo de las scrollbars).

    `QPdfView` es un `QAbstractScrollArea` — no trae ningún modo "mano"
    propio, así que se arma acá con un filtro de eventos sobre su
    `viewport()` (los eventos de mouse de un `QAbstractScrollArea` le
    llegan al viewport, no al widget contenedor). Cursor mano-abierta en
    reposo, mano-cerrada mientras se arrastra — igual que Acrobat/
    cualquier visor real."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
        self._arrastrando = False
        self._pos_inicial = QPoint()
        self._scroll_inicial = QPoint()
        self.viewport().installEventFilter(self)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802 (Qt override)
        if watched is not self.viewport():
            return False

        tipo = event.type()
        if tipo == QEvent.Type.MouseButtonPress and event.button() == Qt.MouseButton.LeftButton:  # type: ignore[attr-defined]
            self._arrastrando = True
            self._pos_inicial = event.globalPosition().toPoint()  # type: ignore[attr-defined]
            self._scroll_inicial = QPoint(self.horizontalScrollBar().value(), self.verticalScrollBar().value())
            self.viewport().setCursor(Qt.CursorShape.ClosedHandCursor)
            return True
        if tipo == QEvent.Type.MouseMove and self._arrastrando:
            delta = event.globalPosition().toPoint() - self._pos_inicial  # type: ignore[attr-defined]
            self.horizontalScrollBar().setValue(self._scroll_inicial.x() - delta.x())
            self.verticalScrollBar().setValue(self._scroll_inicial.y() - delta.y())
            return True
        if tipo == QEvent.Type.MouseButtonRelease and self._arrastrando:
            self._arrastrando = False
            self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
            return True
        return False


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
        self._ruta_pdf = Path(ruta_pdf)

        layout = QVBoxLayout(self)

        self._documento = QPdfDocument(self)
        self._documento.load(str(ruta_pdf))

        layout.addLayout(self._armar_barra_zoom())

        self._vista = _VistaPdfConManito(self)
        self._vista.setDocument(self._documento)
        self._vista.setPageMode(QPdfView.PageMode.MultiPage)
        self._vista.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        layout.addWidget(self._vista, stretch=1)
        # `zoomFactor()` con `FitToWidth` recién queda calculado después
        # de que Qt activa el layout — se difiere un tick, mismo criterio
        # ya usado en otras pantallas para timings de Qt.
        QTimer.singleShot(0, self._actualizar_etiqueta_zoom)

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

    # ------------------------------------------------------------------
    # Zoom
    # ------------------------------------------------------------------
    def _armar_barra_zoom(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        # Botones angostos con su propio padding chico (feedback del
        # usuario, 2026-08-19: "falta el signo '-' y '+' en los botones
        # de zoom") — bug real de estilo: la hoja de estilo QSS global
        # de la app (`theme.py`) le da a TODO `QPushButton` un padding
        # de 16px por lado; con `setMaximumWidth(32)` ese padding solo
        # ya ocupaba el ancho completo del botón, sin dejarle lugar al
        # signo — quedaba dibujado pero recortado a 0px, invisible.
        # También se cambia el signo "−" (MINUS SIGN, U+2212) por un
        # guion ASCII simple "-", más seguro entre fuentes.
        estilo_boton_zoom = "padding: 2px 4px;"
        btn_zoom_menos = QPushButton("-")
        btn_zoom_menos.setStyleSheet(estilo_boton_zoom)
        btn_zoom_menos.setFixedWidth(32)
        btn_zoom_menos.setToolTip("Disminuir zoom")
        btn_zoom_menos.clicked.connect(self._on_zoom_menos)
        fila.addWidget(btn_zoom_menos)

        self.lbl_zoom = QLabel("100%")
        self.lbl_zoom.setMinimumWidth(50)
        self.lbl_zoom.setAlignment(Qt.AlignmentFlag.AlignCenter)
        fila.addWidget(self.lbl_zoom)

        btn_zoom_mas = QPushButton("+")
        btn_zoom_mas.setStyleSheet(estilo_boton_zoom)
        btn_zoom_mas.setFixedWidth(32)
        btn_zoom_mas.setToolTip("Aumentar zoom")
        btn_zoom_mas.clicked.connect(self._on_zoom_mas)
        fila.addWidget(btn_zoom_mas)

        btn_ajustar = QPushButton("Ajustar Ancho")
        btn_ajustar.clicked.connect(self._on_zoom_ajustar)
        fila.addWidget(btn_ajustar)

        fila.addStretch()

        # "Guardar como..." (pedido del usuario, 2026-08-19: "permita al
        # operador guardar o imprimir") — el PDF ya queda escrito en
        # disco (`migration/pdf_output/`) apenas se genera, esto es sólo
        # para dejarle al operador elegir OTRA carpeta/nombre a mano
        # (ej. guardarlo en el escritorio para mandarlo por mail).
        btn_guardar = QPushButton("Guardar como...")
        btn_guardar.clicked.connect(self._on_guardar_como)
        fila.addWidget(btn_guardar)

        # "Poder imprimir 1, varias o todas las páginas" — el diálogo
        # nativo de impresión (`QPrintDialog`, con `PrintPageRange`
        # habilitado) ya deja elegir Todas / Página actual / Rango, no
        # hace falta un selector propio.
        btn_imprimir = QPushButton("Imprimir...")
        btn_imprimir.clicked.connect(self._on_imprimir)
        fila.addWidget(btn_imprimir)

        return fila

    def _on_guardar_como(self) -> None:
        destino, _filtro = QFileDialog.getSaveFileName(
            self, "Guardar PDF como...", self._ruta_pdf.name, "PDF (*.pdf)"
        )
        if not destino:
            return
        try:
            shutil.copyfile(self._ruta_pdf, destino)
        except OSError as exc:
            QMessageBox.critical(self, "Guardar como...", f"No se pudo guardar el PDF:\n{exc}")

    def _on_zoom_mas(self) -> None:
        self._vista.setZoomMode(QPdfView.ZoomMode.Custom)
        actual = self._vista.zoomFactor()
        siguiente = next((z for z in PASOS_ZOOM if z > actual + 0.01), PASOS_ZOOM[-1])
        self._vista.setZoomFactor(siguiente)
        self._actualizar_etiqueta_zoom()

    def _on_zoom_menos(self) -> None:
        self._vista.setZoomMode(QPdfView.ZoomMode.Custom)
        actual = self._vista.zoomFactor()
        anterior = next((z for z in reversed(PASOS_ZOOM) if z < actual - 0.01), PASOS_ZOOM[0])
        self._vista.setZoomFactor(anterior)
        self._actualizar_etiqueta_zoom()

    def _on_zoom_ajustar(self) -> None:
        self._vista.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        QTimer.singleShot(0, self._actualizar_etiqueta_zoom)

    def _actualizar_etiqueta_zoom(self) -> None:
        self.lbl_zoom.setText(f"{round(self._vista.zoomFactor() * 100)}%")

    # ------------------------------------------------------------------
    # Imprimir
    # ------------------------------------------------------------------
    def _on_imprimir(self) -> None:
        """Imprime 1, varias o todas las páginas — el diálogo nativo de
        Windows ya resuelve la elección (`PrintRange.PageRange` con
        Desde/Hasta, o `AllPages`). Cada página del PDF se renderiza a
        imagen a la resolución real de la impresora (`QPdfDocument.
        render`, no hay forma de pasarle un `QPdfDocument` directo a un
        `QPrinter`)."""
        total_paginas = self._documento.pageCount()
        if total_paginas <= 0:
            return

        impresora = QPrinter(QPrinter.PrinterMode.HighResolution)
        impresora.setFromTo(1, total_paginas)
        dialogo = QPrintDialog(impresora, self)
        dialogo.setOption(QPrintDialog.PrintDialogOption.PrintPageRange, True)
        if dialogo.exec() != QDialog.DialogCode.Accepted:
            return

        if impresora.printRange() == QPrinter.PrintRange.PageRange:
            desde = max(impresora.fromPage(), 1)
            hasta = min(impresora.toPage(), total_paginas)
        else:
            desde, hasta = 1, total_paginas

        pintor = QPainter(impresora)
        try:
            resolucion = impresora.resolution()
            primera = True
            for pagina in range(desde - 1, hasta):
                if not primera:
                    impresora.newPage()
                primera = False
                tamano_pto = self._documento.pagePointSize(pagina)
                ancho_px = max(round(tamano_pto.width() / 72 * resolucion), 1)
                alto_px = max(round(tamano_pto.height() / 72 * resolucion), 1)
                imagen = self._documento.render(pagina, QSize(ancho_px, alto_px))
                pintor.drawImage(pintor.viewport(), imagen)
        finally:
            pintor.end()

    def _on_grabar(self) -> None:
        self.grabar = True
        self.accept()
