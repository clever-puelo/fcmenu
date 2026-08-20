"""`ComprobanteAplicarDialog` — selector flotante de "Comprobante a
Imputar" para Nota de Crédito (Concepto Libre y Devolución de
Mercadería), reemplaza el panel de tabla completa que antes vivía
siempre visible en la ventana (pedido del usuario, 2026-08-19: "subir
los botones nota clte y nueva a la linea de NCredito y nDeb., en ese
lugar colocar un boto que diga 'Aplicar a'... Cuando se oprima, aparezca
una ventana flotante para seleccionar. 6 renglones con scroll. Facturas
o NDebito sin aplicar/pagar total o parcialmente (resto > 0)").

`CuentaCorrienteService.facturas_pendientes()` ya filtra exactamente eso
(TIPO Factura/ND con `DEBE > 0` — `DEBE` es el resto real sin aplicar de
ESE comprobante puntual, no el saldo general del Cliente), así que acá
sólo se muestra tal cual.

6 renglones visibles con scroll para el resto del pedido explícito — la
tabla tiene alto fijo, no crece con la cantidad de filas. Un clic elige
y cierra, mismo criterio que `DespachoSelectorDialog`/
`ClienteBusquedaWindow` (modo_seleccion): un clic alcanza para un panel
de selección simple, no hace falta doble clic.

Columnas pedidas: Fecha / Número / Cant. Items / Importe sin Aplicar.
"Número" lleva el Tipo abreviado adelante (ej. "Fact. 0003-00000012")
para poder distinguir Factura de Nota de Débito sin agregar una columna
aparte."""

from __future__ import annotations

from decimal import Decimal
from typing import Optional

from PyQt6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from migration.models import Ctascte
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import CuentaCorrienteService

from .decimals import format_decimal
from .widgets import TablaBusqueda

COL_FECHA = 0
COL_NUMERO = 1
COL_CANT_ITEMS = 2
COL_IMPORTE = 3
COLUMNAS = ["Fecha", "Número", "Cant. Items", "Importe sin Aplicar"]

FILAS_VISIBLES = 6
ALTO_FILA = 24
ALTO_HEADER = 24


class ComprobanteAplicarDialog(QDialog):
    def __init__(
        self,
        repos: RepositoryFactory,
        pendientes: list[Ctascte],
        *,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.repos = repos
        self.comprobante_elegido: Optional[Ctascte] = None
        self.setWindowTitle("Aplicar a")
        self.resize(560, FILAS_VISIBLES * ALTO_FILA + ALTO_HEADER + 100)

        self._construir_ui(pendientes)

    # ------------------------------------------------------------------
    def _construir_ui(self, pendientes: list[Ctascte]) -> None:
        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel("Elegí la Factura o Nota de Débito que esta Nota cancela (total o parcialmente):")
        )

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(COL_CANT_ITEMS, COL_IMPORTE))
        # Alto fijo a 6 filas — el resto se ve haciendo scroll (pedido
        # explícito del usuario), no una tabla que crece sin límite.
        self.tabla.setFixedHeight(FILAS_VISIBLES * ALTO_FILA + ALTO_HEADER + 4)
        self.tabla.itemClicked.connect(self._on_elegir)
        self._cargar(pendientes)
        layout.addWidget(self.tabla)

        if not pendientes:
            layout.addWidget(QLabel("No hay Facturas ni Notas de Débito pendientes para este Cliente."))

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.clicked.connect(self.reject)
        fila_botones.addWidget(btn_cancelar)
        layout.addLayout(fila_botones)

    def _cargar(self, pendientes: list[Ctascte]) -> None:
        filas = []
        fechas_orden = []
        for comprobante in pendientes:
            tipo_abrev = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
            numero = f"{tipo_abrev} {(comprobante.PREFIJO or 0):04d}-{(comprobante.CPBTE or 0):08d}"
            cant_items = len(
                self.repos.fcestad1().by_comprobante(
                    comprobante.TIPO or 0,
                    comprobante.LETRA or "",
                    comprobante.PREFIJO or 0,
                    comprobante.CPBTE or 0,
                )
            )
            filas.append(
                (
                    [
                        comprobante.FECHA.strftime("%d/%m/%Y") if comprobante.FECHA else "",
                        numero,
                        str(cant_items),
                        format_decimal(comprobante.DEBE or Decimal("0")),
                    ],
                    comprobante,
                )
            )
            fechas_orden.append(comprobante.FECHA)
        self.tabla.cargar_filas(filas, claves_orden={COL_FECHA: fechas_orden})

    def _on_elegir(self, item) -> None:
        comprobante = self.tabla.objeto_en_fila(item.row())
        if comprobante is None:
            return
        self.comprobante_elegido = comprobante
        self.accept()

    # ------------------------------------------------------------------
    @staticmethod
    def elegir(
        repos: RepositoryFactory,
        cuenta_corriente_service: CuentaCorrienteService,
        clte: int,
        *,
        parent: QWidget | None = None,
    ) -> Optional[Ctascte]:
        """Conveniencia: arma la lista de pendientes y abre el diálogo.
        Devuelve el comprobante elegido, o `None` si se canceló."""
        pendientes = cuenta_corriente_service.facturas_pendientes(clte)
        dialogo = ComprobanteAplicarDialog(repos, pendientes, parent=parent)
        dialogo.exec()
        return dialogo.comprobante_elegido
