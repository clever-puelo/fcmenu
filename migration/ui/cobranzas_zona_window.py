"""`CobranzasZonaWindow` — Consulta "Cobranzas por Zona", migración de
`VerCobra.frm`: Facturas/ND impagas de todos los clientes de una Zona,
agrupadas por cliente con subtotales.

**Jerarquía con `QTreeWidget`** en vez de replicar el mecanismo real de
`VSFlexGrid` (`MergeCells`+`Outline`+`SubTotal`, una grilla plana con
filas fusionadas que simulan agrupación): `QTreeWidget` ya resuelve
expandir/contraer de forma nativa (los botones "Expandir"/"Contraer" del
legacy pasan a ser `expandAll()`/`collapseAll()`) y el subtotal por
cliente se muestra directo en la fila padre (Importe/Debe agregados),
sin necesitar una fila "Total" aparte — mismo criterio ya aplicado en
otras pantallas de este proyecto (preferir el widget Qt idiomático a
replicar el mecanismo interno del FlexGrid).
"""

from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.repository import RepositoryFactory
from migration.services import CuentaCorrienteService

from .procesando_dialog import ejecutar_con_progreso

COLUMNAS = ["Cliente / Comprobante", "Tipo", "Fecha", "Fec.Vto.", "Importe", "Debe", "Días Venc.", "Estado"]
COL_TIPO, COL_FECHA, COL_FECVTO, COL_IMPORTE, COL_DEBE, COL_DIAS, COL_ESTADO = range(1, 8)
COLUMNAS_DERECHA = {COL_IMPORTE, COL_DEBE, COL_DIAS}

COLOR_VENCIDO = Qt.GlobalColor.red


class CobranzasZonaWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Cobranzas por Zona")
        # +30% ancho / +60% alto más (feedback del usuario, 2026-08-18,
        # tercera ronda — el usuario repitió el mismo pedido, confirmó
        # sumarlo de nuevo).
        self.resize(1690, 1587)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.cc_service = CuentaCorrienteService(self.db)

        self._construir_ui()
        self._poblar_combo_zona()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        fila_filtro = QHBoxLayout()
        fila_filtro.addWidget(QLabel("Zona :"))
        self.combo_zona = QComboBox()
        self.combo_zona.currentIndexChanged.connect(self._on_zona_cambiada)
        fila_filtro.addWidget(self.combo_zona)

        btn_expandir = QPushButton("Expandir Todo")
        btn_expandir.clicked.connect(lambda: self.arbol.expandAll())
        fila_filtro.addWidget(btn_expandir)
        btn_contraer = QPushButton("Contraer Todo")
        btn_contraer.clicked.connect(lambda: self.arbol.collapseAll())
        fila_filtro.addWidget(btn_contraer)

        fila_filtro.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.arbol = QTreeWidget()
        self.arbol.setColumnCount(len(COLUMNAS))
        self.arbol.setHeaderLabels(COLUMNAS)
        self.arbol.setSortingEnabled(True)
        layout.addWidget(self.arbol, stretch=1)

        fila_totales = QHBoxLayout()
        fila_totales.addWidget(QLabel("Deuda Total :"))
        self.lbl_total_deuda = QLabel("0,00")
        self.lbl_total_deuda.setStyleSheet("font-weight: bold;")
        fila_totales.addWidget(self.lbl_total_deuda)
        fila_totales.addSpacing(30)
        fila_totales.addWidget(QLabel("Deuda Venc. :"))
        self.lbl_total_vencido = QLabel("0,00")
        self.lbl_total_vencido.setStyleSheet("font-weight: bold; color: #c00000;")
        fila_totales.addWidget(self.lbl_total_vencido)
        fila_totales.addStretch()
        layout.addLayout(fila_totales)

    def _poblar_combo_zona(self) -> None:
        self.combo_zona.blockSignals(True)
        for zona in self.repos.fctablas().by_ctab("ZN"):
            cod = (zona.COD or "").strip()
            if not cod.isdigit():
                continue
            self.combo_zona.addItem(f"{cod}-{(zona.DESCRI or '').strip()}", int(cod))
        self.combo_zona.blockSignals(False)
        if self.combo_zona.count():
            self._on_zona_cambiada()

    # ------------------------------------------------------------------
    def _on_zona_cambiada(self) -> None:
        zona = self.combo_zona.currentData()
        if zona is None:
            return

        # "Procesando..." con progreso + Cancelar (pedido del usuario,
        # 2026-08-18) — la consulta corre en un hilo aparte con su
        # PROPIA sesión de base (`get_session()`, no `self.db`): las
        # sesiones de SQLAlchemy no son seguras para compartir entre
        # threads, ver docstring de `ejecutar_con_progreso`.
        def _consultar():
            db_hilo = get_session()
            try:
                return CuentaCorrienteService(db_hilo).cobranzas_por_zona(zona)
            finally:
                db_hilo.close()

        resultado = ejecutar_con_progreso(self, "Buscando cobranzas de la zona...", _consultar)
        if resultado is None:
            return  # cancelado por el usuario

        self.arbol.setSortingEnabled(False)
        self.arbol.clear()
        for cliente in resultado.clientes:
            item_cliente = QTreeWidgetItem(
                [
                    f"{cliente.cliente_codigo} — {cliente.cliente_nombre}",
                    "", "", "",
                    format_decimal(cliente.subtotal_importe),
                    format_decimal(cliente.subtotal_debe),
                    "", "",
                ]
            )
            fuente = item_cliente.font(0)
            fuente.setBold(True)
            for col in range(len(COLUMNAS)):
                item_cliente.setFont(col, fuente)
                if col in COLUMNAS_DERECHA:
                    item_cliente.setTextAlignment(col, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            for comp in cliente.comprobantes:
                item_comp = QTreeWidgetItem(
                    [
                        f"    {comp.tipo_label} {comp.letra} {comp.comprobante}",
                        comp.tipo_label,
                        comp.fecha.strftime("%d/%m/%Y") if comp.fecha else "",
                        comp.fecvto.strftime("%d/%m/%Y") if comp.fecvto else "",
                        format_decimal(comp.importe),
                        format_decimal(comp.debe),
                        str(comp.dias_vencido) if comp.vencido else "",
                        "Venc." if comp.vencido else "",
                    ]
                )
                for col in COLUMNAS_DERECHA:
                    item_comp.setTextAlignment(col, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if comp.vencido:
                    for col in range(len(COLUMNAS)):
                        item_comp.setForeground(col, COLOR_VENCIDO)
                item_cliente.addChild(item_comp)

            self.arbol.addTopLevelItem(item_cliente)

        self.arbol.expandAll()
        self.arbol.setSortingEnabled(True)
        for col in range(len(COLUMNAS)):
            self.arbol.resizeColumnToContents(col)

        self.lbl_total_deuda.setText(format_decimal(resultado.total_deuda))
        self.lbl_total_vencido.setText(format_decimal(resultado.total_vencido))

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
