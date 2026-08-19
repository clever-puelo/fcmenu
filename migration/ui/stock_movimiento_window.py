"""`StockMovimientoWindow` — "Ingreso de Movimientos al Stock", migración
de `Stock.frm`.

**Arquitectura**: una sola ventana (cabecera + grilla interactiva de
renglones + Pie), mismo criterio que Facturador/Recibo/Consulta de
Cta.Cte. Los renglones se cargan directo en la grilla (`StockRenglonesGrid`,
estilo Excel — feedback del usuario, 2026-08-16: *"La grilla debe ser
interactiva. Toma esto como norma general para todas las grillas del
sistema. Igual que en el legacy con el vsflexgrid"*), no vía diálogo
modal — el botón "Agregar Renglón..." (y `RenglonStockDialog`, que lo
respaldaba) se sacaron.

Enter en "Nº Cpbante." pasa el foco directo a la grilla, lista para F2/
Enter (mismo pedido: *"al dar enter en comprobante pase a la carga en la
grilla"*).

**Tipo de Movimiento → Forma**: reemplaza los 3 `Frame`/`OptionButton`
del legacy (`Frame1` Tipo, `Frame2` Forma de Entrada, `Frame3` Forma de
Salida) por 2 grupos de radios: Entrada/Salida, y la Forma correspondiente
según cuál esté elegida (mismos 4+4 valores reales: Fábrica/Importación/
Compra Local/Ajuste "+" Inv. — Remito/Ajuste "-"/Promoción/Ajuste "-" Inv.).

**Nº de Despacho**: sólo obligatorio y visible para Entrada+Importación —
igual que el legacy. Ver `StockMovimientoService` para el porqué NO se
replica el resto de la lógica de `Despachos` de `Graba()` (creaba lotes
fantasma con `NRODESP=''` para cualquier otro movimiento).
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QButtonGroup,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.repository import RepositoryFactory
from migration.services import StockMovimientoService

from .stock_renglones_grid import StockRenglonesGrid
from .widgets import EnteroLineEdit, UpperCaseLineEdit

FORMAS_ENTRADA = [
    ("fabrica", "Fábrica"),
    ("importacion", "Importación"),
    ("compra_local", "Compra Local"),
    ("ajuste_mas", 'Ajuste "+" Inv.'),
]
FORMAS_SALIDA = [
    ("remito", "Remito"),
    ("ajuste_menos", "Ajuste -"),
    ("promocion", "Promoción"),
    ("ajuste_inv_menos", 'Ajuste "-" Inv.'),
]


class StockMovimientoWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Ingreso de Movimientos al Stock")
        self.resize(720, 560)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.movimiento_service = StockMovimientoService(self.db)

        self._construir_ui()
        self._on_tipo_cambiado()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        # Más apretado (feedback del usuario, 2026-08-17, segunda ronda:
        # "Achicar. sale del panel") — antes usaba el spacing default de
        # Qt entre las 4 secciones apiladas.
        layout.setSpacing(6)

        layout.addWidget(self._armar_tipo_y_forma())
        layout.addWidget(self._armar_comprobante())
        layout.addWidget(self._armar_renglones(), stretch=1)
        layout.addLayout(self._armar_botones())

    # ------------------------------------------------------------------
    # Tipo de Movimiento / Forma
    # ------------------------------------------------------------------
    def _armar_tipo_y_forma(self) -> QWidget:
        contenedor = QWidget()
        fila = QHBoxLayout(contenedor)
        fila.setContentsMargins(0, 0, 0, 0)

        # Entradas/Salidas uno al lado del otro, no apilados — mismo
        # motivo (2026-08-17, segunda ronda): esto y "Forma" en fila
        # (ver abajo) eran las 2 secciones más altas de la ventana.
        # Márgenes internos apretados (feedback del usuario, 2026-08-19:
        # "subir los paneles de Tipo de mov. y forma y reducir su
        # altura") — sólo tienen una fila de radios adentro, el margen
        # default de Qt les sobraba y empujaba todo lo de abajo (Nº de
        # Cpbte., Renglones) más lejos de lo necesario.
        grupo_tipo = QGroupBox("Tipo de Movimiento")
        layout_tipo = QHBoxLayout(grupo_tipo)
        layout_tipo.setContentsMargins(8, 4, 8, 4)
        self.radio_entrada = QRadioButton("Entradas")
        self.radio_salida = QRadioButton("Salidas")
        self.radio_entrada.setChecked(True)
        self.grupo_botones_tipo = QButtonGroup(self)
        self.grupo_botones_tipo.addButton(self.radio_entrada)
        self.grupo_botones_tipo.addButton(self.radio_salida)
        self.radio_entrada.toggled.connect(self._on_tipo_cambiado)
        layout_tipo.addWidget(self.radio_entrada)
        layout_tipo.addWidget(self.radio_salida)
        fila.addWidget(grupo_tipo)

        # Los 4 radios de "Forma" en fila (antes apilados) — cabían los
        # 4 "Ajuste "-" Inv." horizontalmente sin problema, la fila
        # ya tiene stretch=1 (ancha).
        self.grupo_forma = QGroupBox("Forma")
        self.layout_forma = QHBoxLayout(self.grupo_forma)
        self.layout_forma.setContentsMargins(8, 4, 8, 4)
        self.grupo_botones_forma = QButtonGroup(self)
        self.radios_forma: dict[str, QRadioButton] = {}
        fila.addWidget(self.grupo_forma, stretch=1)

        return contenedor

    def _on_tipo_cambiado(self) -> None:
        es_entrada = self.radio_entrada.isChecked()
        formas = FORMAS_ENTRADA if es_entrada else FORMAS_SALIDA

        for boton in list(self.radios_forma.values()):
            self.grupo_botones_forma.removeButton(boton)
            self.layout_forma.removeWidget(boton)
            boton.deleteLater()
        self.radios_forma.clear()

        for clave, etiqueta in formas:
            boton = QRadioButton(etiqueta)
            self.grupo_botones_forma.addButton(boton)
            self.layout_forma.addWidget(boton)
            boton.toggled.connect(self._on_forma_cambiada)
            self.radios_forma[clave] = boton
        self.radios_forma[formas[0][0]].setChecked(True)

        if hasattr(self, "grilla"):
            self.grilla.reiniciar()
        self._on_forma_cambiada()

    def _forma_elegida(self) -> str:
        for clave, boton in self.radios_forma.items():
            if boton.isChecked():
                return clave
        return ""

    def _on_forma_cambiada(self) -> None:
        es_importacion = self.radio_entrada.isChecked() and self._forma_elegida() == "importacion"
        self.lbl_nro_despacho.setVisible(es_importacion)
        self.txt_nro_despacho.setVisible(es_importacion)

    # ------------------------------------------------------------------
    # Comprobante
    # ------------------------------------------------------------------
    def _armar_comprobante(self) -> QWidget:
        contenedor = QWidget()
        fila = QHBoxLayout(contenedor)

        fila.addWidget(QLabel("Nº Cpbante. :"))
        self.txt_nro_comprobante = EnteroLineEdit()
        self.txt_nro_comprobante.setMaximumWidth(120)
        # Enter en Nº Cpbante. pasa directo a la grilla (feedback del
        # usuario, 2026-08-16) — no se usa EnterAsTabFilter acá porque el
        # salto es de la cabecera A la grilla, no un ciclo entre campos.
        self.txt_nro_comprobante.returnPressed.connect(lambda: self.grilla.foco_inicial())
        fila.addWidget(self.txt_nro_comprobante)

        self.lbl_nro_despacho = QLabel("Nº Despacho :")
        self.txt_nro_despacho = UpperCaseLineEdit()
        self.txt_nro_despacho.setMaximumWidth(160)
        self.txt_nro_despacho.returnPressed.connect(lambda: self.grilla.foco_inicial())
        fila.addWidget(self.lbl_nro_despacho)
        fila.addWidget(self.txt_nro_despacho)

        fila.addStretch()
        return contenedor

    # ------------------------------------------------------------------
    # Renglones
    # ------------------------------------------------------------------
    def _armar_renglones(self) -> QGroupBox:
        grupo = QGroupBox("Renglones — F2/Enter para elegir Artículo, Supr para quitar el renglón")
        layout = QVBoxLayout(grupo)

        self.grilla = StockRenglonesGrid(self.repos)
        layout.addWidget(self.grilla, stretch=1)

        return grupo

    # ------------------------------------------------------------------
    # Botones de acción
    # ------------------------------------------------------------------
    def _armar_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()
        self.btn_grabar = QPushButton("Grabar")
        self.btn_grabar.setStyleSheet("font-weight: bold; padding: 8px;")
        self.btn_grabar.clicked.connect(self._on_grabar)
        fila.addWidget(self.btn_grabar)

        btn_limpiar = QPushButton("Limpiar")
        btn_limpiar.clicked.connect(self._on_limpiar)
        fila.addWidget(btn_limpiar)

        fila.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila.addWidget(btn_cerrar)
        return fila

    def _on_limpiar(self) -> None:
        self.txt_nro_comprobante.clear()
        self.txt_nro_despacho.clear()
        self.grilla.reiniciar()

    def _on_grabar(self) -> None:
        nro_comprobante = self.txt_nro_comprobante.entero()
        if nro_comprobante is None:
            QMessageBox.warning(self, "Movimiento de Stock", "Ingresá un Nº de Comprobante válido.")
            return

        es_entrada = self.radio_entrada.isChecked()
        forma = self._forma_elegida()
        nro_despacho = self.txt_nro_despacho.text().strip() or None

        renglones = self.grilla.renglones
        if not renglones:
            QMessageBox.warning(self, "Movimiento de Stock", "No hay renglones para grabar.")
            return

        vbmsgbox = QMessageBox.question(
            self,
            "Grabación del Movimiento",
            "El siguiente proceso registrará un Movimiento en el Stock.\n¿Desea continuar?",
        )
        if vbmsgbox != QMessageBox.StandardButton.Yes:
            return

        try:
            resultado = self.movimiento_service.emitir_movimiento(
                es_entrada=es_entrada,
                forma=forma,
                nro_comprobante=nro_comprobante,
                renglones=renglones,
                usuario=self._usuario(),
                nro_despacho=nro_despacho,
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Movimiento de Stock", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Movimiento de Stock", f"No se pudo grabar el movimiento:\n{exc}")
            return

        QMessageBox.information(
            self,
            "Movimiento de Stock",
            f"Movimiento grabado (tipo {resultado.tipo_codigo}, "
            f"{resultado.cantidad_renglones} renglón/es).",
        )
        self._on_limpiar()

    @staticmethod
    def _usuario() -> str:
        import os

        return os.environ.get("USERNAME", "SISTEMA")[:6]

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
