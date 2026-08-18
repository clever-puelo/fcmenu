"""`VentasArticuloWindow` — "Ventas por Artículo", migración de
`VTAXART.frm`: ventas de `Fcestad1` de una Sección/mes, agrupadas y
subtotalizadas por dimensión física según la Sección (ver
`EstadisticaVentasService.CONFIG_AGRUPAMIENTO_POR_SECCION` — 4
excepciones reales hardcodeadas por nombre de Sección, confirmado con
el usuario 2026-08-16 replicar tal cual).

**`QTreeWidget`** en vez del mecanismo real de `VSFlexGrid`
(`MergeCells`+`Outline`+`SubTotal`, con niveles de agrupamiento
variables según la Sección) — mismo criterio ya aplicado en
`CobranzasZonaWindow`: expandir/contraer nativo del árbol, subtotal en
la fila del grupo.
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtCore import QDate, Qt
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
from migration.services import EstadisticaVentasService, NodoAgrupado

from .factura_renglon import posiciones_activas_seccion
from .widgets import EnteroLineEdit, redimensionar_pct_pantalla

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

ETIQUETAS_NIVEL = {"pulg": "Pulg", "milim": "MM", "telas": "Telas", "mtr": "Mtr"}
ETIQUETAS_COLUMNA = {"mtr": "Mtr", "cantidad": "Cant.", "importe": "Importe"}

COLUMNAS = ["Grupo / Comprobante", "Fecha", "Cliente", "Pulg", "MM", "Telas", "Mtr", "Cant.", "P.Venta", "Importe"]
COL_FECHA, COL_CLIENTE, COL_PULG, COL_MM, COL_TELAS, COL_MTR, COL_CANT, COL_PVTA, COL_IMPORTE = range(1, 10)
COLUMNAS_DERECHA = {COL_PULG, COL_MM, COL_TELAS, COL_MTR, COL_CANT, COL_PVTA, COL_IMPORTE}

# Posición del enum real (ver factura_renglon.POSICIONES_UM) que le
# corresponde a cada columna de Unidad de Medida — mismo criterio que
# VentasSeccionWindow (feedback del usuario, 2026-08-16: "similar al
# anterior").
POSICION_POR_COLUMNA = {COL_PULG: 1, COL_MTR: 2, COL_MM: 3, COL_TELAS: 4, COL_CANT: 6}


class VentasArticuloWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Ventas por Artículo")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        redimensionar_pct_pantalla(self, 88, 80)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.estadistica_service = EstadisticaVentasService(self.db)

        self._construir_ui()
        self._poblar_combos()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # 2 líneas de opciones (feedback del usuario, 2026-08-18: "hacer
        # 2 lineas de opciones, colocar mes y año en una segunda linea
        # para contraer la ventana") — Mes/Año bajan a una fila propia,
        # el resto queda en la primera.
        fila_filtro = QHBoxLayout()
        fila_filtro.addWidget(QLabel("Sección :"))
        self.combo_seccion = QComboBox()
        fila_filtro.addWidget(self.combo_seccion)

        fila_filtro.addWidget(QLabel("Desde :"))
        self.txt_desde = EnteroLineEdit(0)
        self.txt_desde.setMaximumWidth(80)
        fila_filtro.addWidget(self.txt_desde)

        fila_filtro.addWidget(QLabel("Hasta :"))
        self.txt_hasta = EnteroLineEdit(99999999)
        self.txt_hasta.setMaximumWidth(80)
        fila_filtro.addWidget(self.txt_hasta)

        btn_expandir = QPushButton("Expandir")
        btn_expandir.clicked.connect(lambda: self.arbol.expandAll())
        fila_filtro.addWidget(btn_expandir)
        btn_contraer = QPushButton("Contraer")
        btn_contraer.clicked.connect(lambda: self.arbol.collapseAll())
        fila_filtro.addWidget(btn_contraer)

        fila_filtro.addStretch()
        btn_cerrar = QPushButton("Salir")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        fila_fecha = QHBoxLayout()
        fila_fecha.addWidget(QLabel("Mes :"))
        self.combo_mes = QComboBox()
        self.combo_mes.addItems(MESES)
        fila_fecha.addWidget(self.combo_mes)

        fila_fecha.addWidget(QLabel("Año :"))
        self.combo_anio = QComboBox()
        fila_fecha.addWidget(self.combo_anio)
        fila_fecha.addStretch()
        layout.addLayout(fila_fecha)

        self.arbol = QTreeWidget()
        self.arbol.setColumnCount(len(COLUMNAS))
        self.arbol.setHeaderLabels(COLUMNAS)
        layout.addWidget(self.arbol, stretch=1)

        fila_totales = QHBoxLayout()
        fila_totales.addWidget(QLabel("Tot. Movim. :"))
        self.lbl_cantidad = QLabel("0")
        fila_totales.addWidget(self.lbl_cantidad)
        fila_totales.addStretch()
        fila_totales.addWidget(QLabel("Total Importe :"))
        self.lbl_total = QLabel("$ 0,00")
        self.lbl_total.setStyleSheet("font-weight: bold;")
        fila_totales.addWidget(self.lbl_total)
        layout.addLayout(fila_totales)

    def _poblar_combos(self) -> None:
        self.combo_seccion.blockSignals(True)
        for seccion in self.repos.fctablas().by_ctab("SC"):
            cod = (seccion.COD or "").strip()
            self.combo_seccion.addItem(f"{cod} - {(seccion.DESCRI or '').strip()}", cod)
        self.combo_seccion.blockSignals(False)

        hoy = QDate.currentDate()
        self.combo_anio.blockSignals(True)
        for anio in range(hoy.year() - 6, hoy.year() + 1):
            self.combo_anio.addItem(str(anio), anio)
        self.combo_anio.setCurrentIndex(self.combo_anio.count() - 1)
        self.combo_anio.blockSignals(False)

        self.combo_mes.blockSignals(True)
        self.combo_mes.setCurrentIndex(hoy.month() - 1)
        self.combo_mes.blockSignals(False)

        self.combo_seccion.currentIndexChanged.connect(self._on_filtro_cambiado)
        self.combo_mes.currentIndexChanged.connect(self._on_filtro_cambiado)
        self.combo_anio.currentIndexChanged.connect(self._on_filtro_cambiado)
        self.txt_desde.editingFinished.connect(self._on_filtro_cambiado)
        self.txt_hasta.editingFinished.connect(self._on_filtro_cambiado)

        if self.combo_seccion.count():
            self._on_filtro_cambiado()

    # ------------------------------------------------------------------
    def _on_filtro_cambiado(self) -> None:
        from datetime import date

        cod_seccion = self.combo_seccion.currentData()
        if not cod_seccion:
            return
        anio = self.combo_anio.currentData()
        mes = self.combo_mes.currentIndex() + 1
        fecha_desde = date(anio, mes, 1)

        self._actualizar_columnas_visibles(cod_seccion)

        resultado = self.estadistica_service.ventas_articulo_agrupadas(
            cod_seccion, fecha_desde, cod2_desde=self.txt_desde.entero(), cod2_hasta=self.txt_hasta.entero()
        )

        self.arbol.clear()
        columnas_subtotal = self.estadistica_service.config_agrupamiento(cod_seccion).columnas_subtotal
        for nodo in resultado.raiz:
            self.arbol.addTopLevelItem(self._armar_item(nodo, resultado.niveles_agrupamiento, 0, columnas_subtotal))
        self.arbol.expandAll()
        for col in range(len(COLUMNAS)):
            self.arbol.resizeColumnToContents(col)

        self.lbl_cantidad.setText(str(resultado.cantidad_movimientos))
        self.lbl_total.setText(f"$ {format_decimal(resultado.total_importe)}")

    def _actualizar_columnas_visibles(self, cod_seccion: str) -> None:
        posiciones = posiciones_activas_seccion(self.repos.fctablas(), cod_seccion)
        for columna, posicion in POSICION_POR_COLUMNA.items():
            self.arbol.setColumnHidden(columna, posicion not in posiciones)

    def _armar_item(
        self, nodo: NodoAgrupado, niveles: tuple[str, ...], profundidad: int, columnas_subtotal: tuple[str, ...]
    ) -> QTreeWidgetItem:
        campo_nivel = niveles[profundidad]
        etiqueta_campo = ETIQUETAS_NIVEL.get(campo_nivel, campo_nivel)
        valores = [f"{etiqueta_campo}: {nodo.etiqueta}", "", "", "", "", "", "", "", "", ""]

        col_por_campo = {"pulg": COL_PULG, "milim": COL_MM, "telas": COL_TELAS, "mtr": COL_MTR, "cantidad": COL_CANT, "importe": COL_IMPORTE}
        for campo_sub, valor in nodo.subtotales.items():
            valores[col_por_campo[campo_sub]] = format_decimal(valor)

        item = QTreeWidgetItem(valores)
        fuente = item.font(0)
        fuente.setBold(True)
        for col in range(len(COLUMNAS)):
            item.setFont(col, fuente)
            if col in COLUMNAS_DERECHA:
                item.setTextAlignment(col, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        if nodo.hijos:
            for hijo in nodo.hijos:
                item.addChild(self._armar_item(hijo, niveles, profundidad + 1, columnas_subtotal))
        else:
            for m in nodo.movimientos:
                item.addChild(self._armar_item_movimiento(m))
        return item

    @staticmethod
    def _armar_item_movimiento(m) -> QTreeWidgetItem:
        valores = [
            f"{m.tipo_label} {m.letra} {m.comprobante}",
            m.fecha.strftime("%d/%m/%Y") if m.fecha else "",
            f"{m.cliente_codigo or ''} {m.cliente_nombre}",
            format_decimal(m.pulg) if m.pulg else "",
            str(m.milim) if m.milim else "",
            str(m.telas) if m.telas else "",
            format_decimal(m.mtr) if m.mtr else "",
            format_decimal(m.cantidad) if m.cantidad else "",
            format_decimal(m.precio_venta),
            format_decimal(m.importe),
        ]
        item = QTreeWidgetItem(valores)
        for col in COLUMNAS_DERECHA:
            item.setTextAlignment(col, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        return item

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
