"""Selector de Lote/Despacho — réplica del panel flotante `Frame2`/`FG2`
de `DetFact.frm Sub VerSiBusca`, para el renglón de Detalle del futuro
Facturador.

Comportamiento confirmado con el usuario (2026-08-16, fiel al legacy):
al cargar en un renglón un artículo que tiene lotes (`Despachos`) con
`STOCK > 0`, este diálogo se abre AUTOMÁTICAMENTE (no hace falta que el
operador lo pida) — mismo criterio que el panel flotante original, que
aparecía solo debajo del renglón apenas se resolvía el artículo. Un solo
clic en un lote lo elige y cierra (ídem `FG2_Click`, un solo clic — no
doble como en los buscadores modales de Cliente/Artículo: acá es un
panel transitorio, no una grilla de búsqueda de varias columnas para
mirar con calma). El botón "Sin lote específico" (equivalente a no
clickear nada y seguir tipeando en el legacy) cierra sin elegir — es la
salida normal cuando el lote no importa para ese renglón.

**El lote elegido es sólo trazabilidad, no reserva de stock**: en
`EmiFact.frm Sub Graba()` el descuento de `Despachos.STOCK` es ADICIONAL
al descuento general de `Stock.STUnid` (que ocurre siempre, corrigiendo
el bug de `Mtr` documentado en `pyqt6_ui_progress.md`) — no elegir lote
no cambia el resultado del stock general del artículo.
"""

from __future__ import annotations

from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.models import Despacho
from migration.repository import RepositoryFactory

from .decimals import format_decimal
from .widgets import TablaBusqueda


class DespachoSelectorDialog(QDialog):
    """Panel de selección de lote. Usar `abrir_si_corresponde()` desde el
    renglón de Detalle en vez de instanciar directo, para no abrir el
    diálogo cuando el artículo no tiene lotes con stock (mismo criterio
    del legacy: sin filas, no se muestra `Frame2`)."""

    def __init__(
        self,
        lotes: list[Despacho],
        *,
        descripcion_articulo: str = "",
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.despacho_elegido: Despacho | None = None

        titulo = "Lotes Disponibles"
        if descripcion_articulo:
            titulo = f"{titulo} — {descripcion_articulo}"
        self.setWindowTitle(titulo)
        self.resize(480, 300)

        self._construir_ui(lotes)

    # ------------------------------------------------------------------
    def _construir_ui(self, lotes: list[Despacho]) -> None:
        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel("Este artículo tiene lotes con stock disponible. Elegí uno o continuá sin lote específico:")
        )

        self.tabla = TablaBusqueda(["Lote", "Comprobante", "Fecha Entrada", "Stock"])
        self.tabla.itemClicked.connect(self._on_elegir)
        self._cargar_lotes(lotes)
        layout.addWidget(self.tabla, stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        self.btn_sin_lote = QPushButton("Sin lote específico")
        self.btn_sin_lote.clicked.connect(self.reject)
        fila_botones.addWidget(self.btn_sin_lote)
        layout.addLayout(fila_botones)

    def _cargar_lotes(self, lotes: list[Despacho]) -> None:
        filas = [
            (
                [
                    (d.NRODESP or "").strip(),
                    str(d.CPBTE) if d.CPBTE is not None else "",
                    d.FECENT.strftime("%d/%m/%Y") if d.FECENT else "",
                    format_decimal(d.STOCK) if d.STOCK is not None else "0",
                ],
                d,
            )
            for d in lotes
        ]
        self.tabla.cargar_filas(filas)

    def _on_elegir(self, item) -> None:
        despacho = self.tabla.objeto_en_fila(item.row())
        if despacho is None:
            return
        self.despacho_elegido = despacho
        self.accept()

    # ------------------------------------------------------------------
    @staticmethod
    def abrir_si_corresponde(
        repos: RepositoryFactory,
        cod1: str,
        cod2: str,
        *,
        descripcion_articulo: str = "",
        parent: QWidget | None = None,
    ) -> Despacho | None:
        """Conveniencia para el renglón de Detalle: consulta
        `DespachoRepository.con_stock()` y sólo abre el diálogo si hay
        lotes — si no hay, devuelve `None` sin interrumpir la carga
        (mismo criterio que el legacy, que tampoco muestra el panel
        flotante cuando `Despachos` no tiene filas con `STOCK>0` para ese
        artículo). Devuelve el lote elegido, o `None` si no hay lotes o
        el operador cerró con "Sin lote específico"."""
        lotes = repos.despacho().con_stock(cod1, cod2)
        if not lotes:
            return None

        dialogo = DespachoSelectorDialog(lotes, descripcion_articulo=descripcion_articulo, parent=parent)
        dialogo.exec()
        return dialogo.despacho_elegido
