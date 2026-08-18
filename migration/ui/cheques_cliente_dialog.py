"""Diálogo "Ver Cheques" de la Consulta de Cuenta Corriente — migración
de `CtaCte.frm Sub DoVer9`/`CargaCheq` (Picture4/GRILLA2, líneas
2009-2126) + `Sub MuestraCheq` (Picture5, líneas 2135-2200).

Alcance: cheques EN CARTERA (`ESTADO='1'`) de un cliente puntual — no es
la Consulta de Cheques general (`VerCheq.frm`, con las 4 categorías
Cartera/Egresados/Rechazados/Otros), que queda para una fase futura de
"consultas varias" (orden ya confirmado por el usuario).
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtWidgets import (
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.decimals import format_decimal
from migration.models import Cheque
from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda

COLUMNAS = ["Nº Cheque", "Fec. Venc.", "Banco", "Recibo", "Importe", "N/Orden"]


class ChequesClienteDialog(QDialog):
    def __init__(self, repos: RepositoryFactory, codigo_cliente: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.setWindowTitle(f"Cheques en Cartera — Cliente {codigo_cliente}")
        self.resize(720, 480)

        self._cheques = repos.cheque().en_cartera_de_cliente(codigo_cliente)
        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        total = sum((c.IMPORTE or Decimal("0") for c in self._cheques), Decimal("0"))
        layout.addWidget(QLabel(f"Hay {len(self._cheques)} cheques, por $ {format_decimal(total)}"))

        self.tabla = TablaBusqueda(COLUMNAS)
        self.tabla.itemClicked.connect(self._on_click)
        self._cargar_tabla()
        layout.addWidget(self.tabla, stretch=1)

        self.grupo_detalle = self._armar_detalle()
        self.grupo_detalle.setVisible(False)
        layout.addWidget(self.grupo_detalle)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _cargar_tabla(self) -> None:
        filas = []
        for c in self._cheques:
            banco = self.repos.banco().by_cod(int(c.BCOSUC)) if c.BCOSUC else None
            nombre_banco = f"{c.BCOSUC} - {(banco.NOMBRE or '').strip()}" if banco else f"{c.BCOSUC or ''}"
            filas.append(
                (
                    [
                        str(c.NROCHEQ),
                        c.FECVTO.strftime("%d/%m/%Y") if c.FECVTO else "",
                        nombre_banco,
                        str(c.CPBING or ""),
                        format_decimal(c.IMPORTE or Decimal("0")),
                        "N/Orden" if c.ORDEN else "",
                    ],
                    c,
                )
            )
        self.tabla.cargar_filas(filas)

    def _armar_detalle(self) -> QGroupBox:
        grupo = QGroupBox("Detalle del Cheque")
        form = QFormLayout(grupo)
        self.lbl_banco = QLabel()
        self.lbl_emision = QLabel()
        self.lbl_vencimiento = QLabel()
        self.lbl_ingreso = QLabel()
        self.lbl_concepto = QLabel()
        self.lbl_datosad = QLabel()
        form.addRow("Banco :", self.lbl_banco)
        form.addRow("Emisión :", self.lbl_emision)
        form.addRow("Vencimiento :", self.lbl_vencimiento)
        form.addRow("Ingreso :", self.lbl_ingreso)
        form.addRow("Concepto :", self.lbl_concepto)
        form.addRow("Datos Ad. :", self.lbl_datosad)
        return grupo

    # ------------------------------------------------------------------
    def _on_click(self, item) -> None:
        cheque: Cheque | None = self.tabla.objeto_en_fila(item.row())
        if cheque is None:
            return
        banco = self.repos.banco().by_cod(int(cheque.BCOSUC)) if cheque.BCOSUC else None
        self.lbl_banco.setText(f"{cheque.BCOSUC} - {(banco.NOMBRE or '').strip()}" if banco else "*** No Existe en Tabla ***")
        self.lbl_emision.setText(cheque.FECEMI.strftime("%d/%m/%Y") if cheque.FECEMI else "")
        self.lbl_vencimiento.setText(cheque.FECVTO.strftime("%d/%m/%Y") if cheque.FECVTO else "")
        self.lbl_ingreso.setText(cheque.FECING.strftime("%d/%m/%Y") if cheque.FECING else "")
        self.lbl_concepto.setText(cheque.CONCEP or "")
        self.lbl_datosad.setText(cheque.DATOSAD or "")
        self.grupo_detalle.setVisible(True)
