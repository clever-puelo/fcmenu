"""Diálogo de drill-down a una factura/NC/ND desde el extracto de Cuenta
Corriente — migración de `CtaCte.frm Sub MuestraDetalle` (Picture3) +
`Sub Command6_Click` ("Ver Items", Picture2), líneas 2318-2382 y
1742-1792.

**Simplificación de UX confirmada por el mismo criterio ya aplicado en
el resto de la migración** (ej. sacar el paso de confirmación extra del
Recibo): el legacy mostraba primero sólo los totales y recién al
clickear "Ver Items" cargaba el detalle en un panel aparte — acá se
cargan los renglones directo, sin un segundo clic, porque la consulta es
liviana (una factura real tiene pocos renglones) y no hay ningún motivo
de negocio para ocultarlos.
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
from migration.models import FcivaVta
from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda

COLUMNAS_ITEMS = [
    "Sección", "Pulg", "Mtr", "Milím.", "Telas", "Cant.", "P.Esp.", "P.Costo", "P.Venta", "Importe",
]


def construir_grupo_renglones(repos: RepositoryFactory, factura: FcivaVta, titulo: str = "Renglones") -> QGroupBox:
    """Renglones reales (`Fcestad1`) de una factura/NC/ND puntual, en un
    `QGroupBox` listo para insertar en cualquier diálogo — extraído de
    `FacturaDetalleDialog._armar_items` para reusarlo también en
    `FacturaEmitidaDetalleDialog` (`VerFact.frm`, botón "Ver Items"),
    mismo mecanismo (`Fcestad1Repository.by_comprobante`)."""
    grupo = QGroupBox(titulo)
    layout = QVBoxLayout(grupo)
    tabla = TablaBusqueda(COLUMNAS_ITEMS)

    renglones = repos.fcestad1().by_comprobante(
        int(factura.TIPO) if factura.TIPO is not None else 1,
        factura.LETRA or "",
        factura.PTOVTA or 0,
        factura.CPBTE or 0,
    )
    filas = [
        (
            [
                (r.COD1 or "").strip(),
                format_decimal(r.PULG) if r.PULG else "",
                format_decimal(r.MTR) if r.MTR else "",
                str(r.MILIM) if r.MILIM else "",
                str(r.TELAS) if r.TELAS else "",
                format_decimal(r.CANT) if r.CANT else "",
                format_decimal(r.PESP) if r.PESP else "",
                format_decimal(r.PCOS) if r.PCOS else "",
                format_decimal(r.PVTA) if r.PVTA else "",
                format_decimal(r.IMPTE) if r.IMPTE else "",
            ],
            r,
        )
        for r in renglones
    ]
    tabla.cargar_filas(filas)
    layout.addWidget(tabla)
    return grupo


class FacturaDetalleDialog(QDialog):
    """Encabezado (Neto/IVA Ins./IVA No Ins./Total/Ítems/Cantidad) +
    renglones reales (`Fcestad1`) de una factura/NC/ND puntual."""

    def __init__(self, repos: RepositoryFactory, factura: FcivaVta, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.factura = factura

        ptovta_txt = str(factura.PTOVTA or 0).rjust(4, "0")
        cpbte_txt = str(factura.CPBTE or 0).rjust(8, "0")
        self.setWindowTitle(f"Comprobante {factura.LETRA or ''}  {ptovta_txt}-{cpbte_txt}")
        self.resize(700, 480)

        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.addWidget(self._armar_cabecera())
        layout.addWidget(self._armar_items(), stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _armar_cabecera(self) -> QGroupBox:
        f = self.factura
        grupo = QGroupBox(f"{(f.NOMB or '').strip()}  —  {f.CLTE}")
        form = QFormLayout(grupo)

        neto = f.GRINS or Decimal("0")
        iva_ins = f.IVAINS or Decimal("0")
        iva_noins = f.IVANOINS or Decimal("0")
        total = neto + iva_ins + iva_noins

        form.addRow("Tot. Neto :", QLabel(f"$ {format_decimal(neto)}"))
        form.addRow("IVA Ins. :", QLabel(f"$ {format_decimal(iva_ins)}"))
        form.addRow("IVA N/Ins. :", QLabel(f"$ {format_decimal(iva_noins)}"))
        lbl_total = QLabel(f"$ {format_decimal(total)}")
        lbl_total.setStyleSheet("font-weight: bold;")
        form.addRow("Tot. Bruto :", lbl_total)
        form.addRow("Ítems :", QLabel(str(f.ITEMS or 0)))
        form.addRow("Unids. :", QLabel(str(f.TOTCAN or 0)))

        return grupo

    def _armar_items(self) -> QGroupBox:
        return construir_grupo_renglones(self.repos, self.factura)
