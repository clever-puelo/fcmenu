"""`DespachosConsultaWindow` — Consulta de Despachos, migración de
`VerDesp.frm`: lista de lotes reales (`Sub CargaDespa`, agrupados por
NRODESP+FECENT con cantidad de artículos) — un clic en uno abre el
detalle de artículos de ese lote (`LoteDetalleDialog`), centrado sobre
esta ventana (Qt, sin posicionamiento manual — ver `_on_ver_lote`). El
campo de búsqueda por Nº de Despacho exacto (`Sub DoVer`) filtra en vivo.

**Columna "Comprobante"** (pedido del usuario, 2026-08-19: "agregar
comprobante despues del despacho") — novedad respecto del legacy
(`VerDesp.frm` no la mostraba), confirmada explícitamente por el
usuario como mejora deliberada, no un hallazgo de paridad.
"""

from __future__ import annotations

from PyQt6.QtWidgets import QHBoxLayout, QLabel, QMainWindow, QPushButton, QVBoxLayout, QWidget

from migration.db import get_session
from migration.repository import RepositoryFactory

from .lote_detalle_dialog import LoteDetalleDialog
from .widgets import AlfanumericoLineEdit, TablaBusqueda, redimensionar_pct_pantalla

COLUMNAS = ["Nº Despacho", "Comprobante", "Fecha Entrada", "Artículos"]
COL_NRODESP = 0
COL_CPBTE = 1
COL_FECHA = 2
COL_ARTICULOS = 3


class DespachosConsultaWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Consulta de Despachos")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        redimensionar_pct_pantalla(self, 65, 90)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)

        self._construir_ui()
        self._cargar_resumen()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        fila_filtro = QHBoxLayout()
        fila_filtro.addWidget(QLabel("Nº Despacho :"))
        self.txt_filtro = AlfanumericoLineEdit()
        self.txt_filtro.textChanged.connect(self._on_filtro_cambiado)
        fila_filtro.addWidget(self.txt_filtro)
        btn_limpiar = QPushButton("Ver Todos")
        btn_limpiar.clicked.connect(self._cargar_resumen)
        fila_filtro.addWidget(btn_limpiar)
        fila_filtro.addStretch()

        # No tenía botón Cerrar (feedback del usuario, 2026-08-18).
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(COL_NRODESP, COL_CPBTE, COL_ARTICULOS))
        # Un clic (no doble) abre el detalle — pedido del usuario,
        # 2026-08-19: "cuando se clickee en un renglon, abra una ventana
        # flotante" — mismo criterio ya aplicado a otros paneles de
        # selección simples esta misma sesión (ComprobanteAplicarDialog,
        # SaldosClientesDialog).
        self.tabla.itemClicked.connect(self._on_ver_lote)
        layout.addWidget(self.tabla, stretch=1)

    # ------------------------------------------------------------------
    def _cargar_resumen(self) -> None:
        self.txt_filtro.blockSignals(True)
        self.txt_filtro.setText("")
        self.txt_filtro.blockSignals(False)
        filas = self.repos.despacho().resumen_lotes()
        self._cargar_tabla(filas)

    def _on_filtro_cambiado(self, texto: str) -> None:
        # .upper() acá mismo (no confiar en que AlfanumericoLineEdit ya
        # dejó el widget en mayúsculas): fuerza mayúsculas vía textEdited
        # con blockSignals(True), así que este textChanged puede llegar
        # con el valor todavía en minúscula un instante antes.
        texto = texto.strip().upper()
        if not texto:
            self._cargar_resumen()
            return
        filas = [
            f for f in self.repos.despacho().resumen_lotes() if texto in (f["nrodesp"] or "").upper()
        ]
        self._cargar_tabla(filas)

    def _cargar_tabla(self, filas: list[dict]) -> None:
        datos = [
            (
                [
                    f["nrodesp"] or "",
                    str(f["cpbte"]) if f["cpbte"] else "",
                    f["fecent"].strftime("%d/%m/%Y") if f["fecent"] else "",
                    str(f["cantidad"]),
                ],
                f,
            )
            for f in filas
        ]
        # Fecha real como clave de orden explícita: el texto "dd/mm/yyyy"
        # no ordena cronológicamente como string (ver _ItemOrdenable).
        self.tabla.cargar_filas(datos, claves_orden={COL_FECHA: [f["fecent"] for f in filas]})

    def _on_ver_lote(self, item) -> None:
        fila = self.tabla.objeto_en_fila(item.row())
        if fila is None or not fila["nrodesp"]:
            return
        # `fecent` desambigua el caso real de un NRODESP repetido en 2
        # fechas de entrada distintas (ver docstring de
        # `LoteDetalleDialog`/`DespachoRepository.articulos_de_lote`).
        #
        # Sin posición explícita (feedback del usuario, 2026-08-19: "la
        # ventana debe comenzar en el centro del panel central porque
        # sino se sale de pantalla") — la ronda anterior la movía a la
        # derecha de esta ventana con `move()`, pero eso no respeta el
        # borde real de la pantalla y podía salirse por completo si esta
        # ventana ya estaba corrida hacia la derecha. Sin `move()`, Qt
        # centra el diálogo modal sobre su padre solo (mismo
        # comportamiento que ya usan el resto de los diálogos de la app,
        # ninguno hace posicionamiento manual).
        LoteDetalleDialog(self.repos, fila["nrodesp"], fila["fecent"], parent=self).exec()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
