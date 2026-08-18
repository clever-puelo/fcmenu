"""`ChequesConsultaWindow` — "Consulta de Cheques" general (todos los
clientes), migración de `VerCheq.frm`.

**Alcance de Estados ampliado sobre el legacy** (ver docstring de
`ChequeService`): se habilitan los 4 estados reales (En Cartera/
Egresados/Rechazados/Otros) — el legacy tenía "Rechazados"/"Otros"
permanentemente deshabilitados (`Option1(3)/(4) Enabled=False`), sin
ningún comentario que explique el motivo, y sin datos reales en juego
para arriesgar (`Cheques` está vacía en `fcmenu_dev`).
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QDateEdit,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.repository import RepositoryFactory
from migration.services import ChequeService

from .cheque_detalle_dialog import ChequeDetalleDialog
from .widgets import TablaBusqueda, crear_boton_hoy

CANTIDADES = [5, 10, 25, 50, 75, 100, 150]
ESTADOS = [
    ("1", "En Cartera"),
    ("2", "Egresados"),
    ("3", "Rechazados"),
    ("4", "Otros"),
]

COLUMNAS = ["Nº Cheque", "Fec.Vto.", "Banco", "Recibo", "Cliente", "Importe", "N/Orden"]
COL_NROCHEQ, COL_FECVTO, COL_BANCO, COL_RECIBO, COL_CLIENTE, COL_IMPORTE, COL_ORDEN = range(7)


class ChequesConsultaWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Consulta de Cheques")
        # +30% alto más (feedback del usuario, 2026-08-18, tercera ronda
        # — repitió el mismo pedido, confirmó sumarlo de nuevo).
        self.resize(1000, 1362)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.cheque_service = ChequeService(self.db)

        self._construir_ui()
        self._on_filtro_cambiado()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        fila_filtro = QHBoxLayout()

        grupo_estado = QGroupBox("Estado")
        layout_estado = QHBoxLayout(grupo_estado)
        self.grupo_botones_estado = QButtonGroup(self)
        self.radios_estado: dict[str, QRadioButton] = {}
        for codigo, etiqueta in ESTADOS:
            boton = QRadioButton(etiqueta)
            self.grupo_botones_estado.addButton(boton)
            layout_estado.addWidget(boton)
            self.radios_estado[codigo] = boton
        fila_filtro.addWidget(grupo_estado)

        fila_filtro.addWidget(QLabel("Desde :"))
        self.fecha_desde = QDateEdit()
        self.fecha_desde.setCalendarPopup(True)
        self.fecha_desde.setDate(QDate.currentDate())
        self.fecha_desde.dateChanged.connect(self._on_filtro_cambiado)
        fila_filtro.addWidget(self.fecha_desde)
        fila_filtro.addWidget(crear_boton_hoy(self.fecha_desde))

        fila_filtro.addWidget(QLabel("Cantidad :"))
        self.combo_cantidad = QComboBox()
        for cantidad in CANTIDADES:
            self.combo_cantidad.addItem(str(cantidad), cantidad)
        self.combo_cantidad.setCurrentIndex(CANTIDADES.index(50))
        self.combo_cantidad.currentIndexChanged.connect(self._on_filtro_cambiado)
        fila_filtro.addWidget(self.combo_cantidad)

        fila_filtro.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_filtro.addWidget(btn_cerrar)
        layout.addLayout(fila_filtro)

        self.tabla = TablaBusqueda(COLUMNAS, columnas_derecha=(COL_NROCHEQ, COL_RECIBO, COL_IMPORTE))
        self.tabla.itemClicked.connect(self._on_ver_detalle)
        layout.addWidget(self.tabla, stretch=1)

        self.lbl_total = QLabel("Total: $ 0,00")
        self.lbl_total.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.lbl_total)

        # El estado por default se fija y los signals de filtro recién se
        # conectan ACÁ, después de que `self.tabla`/`self.fecha_desde`/
        # `self.combo_cantidad` ya existen — bug real encontrado con
        # teclado/ejecución real: conectar `toggled` y hacer
        # `setChecked(True)` DENTRO del loop que arma los radios disparaba
        # `_on_filtro_cambiado()` de inmediato (síncrono, el `toggled` de
        # Qt se emite en el momento) contra atributos que todavía no
        # existían — una excepción dentro de un slot conectado a una señal
        # Qt no siempre deja un traceback Python visible, sólo aborta el
        # proceso en silencio (mismo tipo de bug de orden ya visto en
        # otras grillas de este proyecto, acá en la cabecera del filtro).
        self.radios_estado["1"].setChecked(True)
        for boton in self.radios_estado.values():
            boton.toggled.connect(self._on_filtro_cambiado)

    # ------------------------------------------------------------------
    def _estado_elegido(self) -> str:
        for codigo, boton in self.radios_estado.items():
            if boton.isChecked():
                return codigo
        return "1"

    def _on_filtro_cambiado(self) -> None:
        estado = self._estado_elegido()
        fecha_desde = self.fecha_desde.date().toPyDate()
        cantidad = self.combo_cantidad.currentData() or 50

        cheques = self.cheque_service.buscar(estado, fecha_desde, cantidad)
        self._cargar_tabla(cheques)

    def _cargar_tabla(self, cheques) -> None:
        total = Decimal("0")
        filas = []
        for c in cheques:
            # Siempre resuelve el NOMBRE real del banco (no sólo el código)
            # — reforzado a pedido del usuario (2026-08-16), mismo criterio
            # que el detalle (ChequeDetalleDialog) para el caso "no existe
            # en la tabla".
            banco = self.repos.banco().by_cod(int(c.BCOSUC)) if c.BCOSUC else None
            if banco is not None:
                nombre_banco = f"{c.BCOSUC} - {(banco.NOMBRE or '').strip()}"
            elif c.BCOSUC:
                nombre_banco = f"{c.BCOSUC} - *** No Existe en Tabla ***"
            else:
                nombre_banco = ""
            cliente = self.repos.cliente().by_codigo(c.CLTE) if c.CLTE else None
            nombre_cliente = f"{c.CLTE} {(cliente.NOMB or '').strip()}" if cliente else f"{c.CLTE} *** Desconocido ***"
            importe = c.IMPORTE or Decimal("0")
            total += importe
            filas.append(
                (
                    [
                        str(c.NROCHEQ),
                        c.FECVTO.strftime("%d/%m/%Y") if c.FECVTO else "",
                        nombre_banco,
                        str(c.CPBING or ""),
                        nombre_cliente,
                        format_decimal(importe),
                        "N/Orden" if c.ORDEN else "",
                    ],
                    c,
                )
            )
        self.tabla.cargar_filas(filas)
        self.lbl_total.setText(f"Total: $ {format_decimal(total)}")

    def _on_ver_detalle(self, item) -> None:
        cheque = self.tabla.objeto_en_fila(item.row())
        if cheque is None:
            return
        dialogo = ChequeDetalleDialog(self.repos, cheque, parent=self)
        dialogo.exec()
        # Siempre se refresca al cerrar: un Egreso o un Borrado pueden
        # sacar al cheque del filtro actual (cambia de Estado o desaparece).
        self._on_filtro_cambiado()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
