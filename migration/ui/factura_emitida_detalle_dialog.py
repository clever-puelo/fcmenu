"""Diálogo de detalle de un comprobante emitido — migración de
`VerFact.frm Sub MuestraDetalle` (Picture1, líneas 1209-1281) + "Ver
Items" (Picture2, ya cubierto por `construir_grupo_renglones` de
`factura_detalle_dialog.py`).

Al abrir el detalle se re-busca el Cliente REAL (`ClienteRepository.
by_codigo`) para Dirección/Localidad/CP/Provincia — más autorizado que
el snapshot guardado en `FCIVAVTA` al momento de facturar (mismo
criterio que el legacy: `Label1/2/10/11/30` arrancan con el snapshot y
se sobrescriben si el Cliente todavía existe, `GoTo SALIR` si no). Sin
Cliente real, el panel se queda con el snapshot para Nombre/Provincia/
Cond.IVA/CUIT, y en blanco para Dirección/Localidad/CP — el legacy
tampoco tenía de dónde sacarlos en ese caso.
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
from migration.provincias import nombre_provincia
from migration.repository import RepositoryFactory

from .cliente_detalle_dialog import CIVA_OPCIONES
from .factura_detalle_dialog import construir_grupo_renglones

CIVA_LABELS = dict(CIVA_OPCIONES)


class FacturaEmitidaDetalleDialog(QDialog):
    def __init__(self, repos: RepositoryFactory, factura: FcivaVta, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.factura = factura
        self.cliente = repos.cliente().by_codigo(factura.CLTE) if factura.CLTE else None

        ptovta_txt = str(factura.PTOVTA or 0).rjust(4, "0")
        cpbte_txt = str(factura.CPBTE or 0).rjust(8, "0")
        self.setWindowTitle(f"Comprobante {(factura.LETRA or '').strip()} {ptovta_txt}-{cpbte_txt}")
        self.resize(760, 560)

        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.addWidget(self._armar_cliente())
        layout.addWidget(self._armar_comprobante())
        layout.addWidget(construir_grupo_renglones(self.repos, self.factura), stretch=1)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _armar_cliente(self) -> QGroupBox:
        f = self.factura
        c = self.cliente
        grupo = QGroupBox("Cliente")
        form = QFormLayout(grupo)

        codigo = c.CODIGO if c else f.CLTE
        nombre = ((c.NOMB if c else f.NOMB) or "").strip()
        form.addRow("Código :", QLabel(str(codigo or "")))
        form.addRow("Nombre :", QLabel(nombre))

        if c is not None:
            form.addRow("Dirección :", QLabel((c.DIR or "").strip()))
            form.addRow("Localidad :", QLabel(f"({(c.CP or '').strip()}) {(c.LOC or '').strip()}"))
            pcia_codigo = c.PCIA
            civa_codigo = c.CIVA
            cuit = c.CUIT
        else:
            civa_texto = (f.CIVA or "").strip()
            pcia_codigo = f.PCIA
            civa_codigo = int(civa_texto) if civa_texto.isdigit() else None
            cuit = f.CUIT

        form.addRow("Provincia :", QLabel(nombre_provincia(pcia_codigo)))
        form.addRow("Cond. IVA :", QLabel(CIVA_LABELS.get(civa_codigo, "")))
        form.addRow("CUIT :", QLabel((cuit or "").strip()))
        return grupo

    def _armar_comprobante(self) -> QGroupBox:
        f = self.factura
        grupo = QGroupBox("Comprobante")
        form = QFormLayout(grupo)

        ptovta_txt = str(f.PTOVTA or 0).rjust(4, "0")
        cpbte_txt = str(f.CPBTE or 0).rjust(8, "0")
        form.addRow("Número :", QLabel(f"{(f.LETRA or '').strip()} {ptovta_txt}-{cpbte_txt}"))
        form.addRow("Fecha :", QLabel(f.FECHA.strftime("%d/%m/%Y") if f.FECHA else ""))
        form.addRow("Ítems :", QLabel(str(f.ITEMS or 0)))
        form.addRow("Unids. :", QLabel(str(f.TOTCAN or 0)))

        bruto = f.GRINS or Decimal("0")
        iva = f.IVAINS or Decimal("0")
        ib = f.TOTIB or Decimal("0")
        # VerFact.frm:1234-1238 — "Tot. Neto" (ElTotal): la condición
        # `If Not RgFCIVA!TotIB Then` es un `Not` bitwise de VB6 sobre un
        # número, no una negación booleana de "¿hay IIBB?" — para
        # cualquier TotIB real (0 o positivo) da un valor no-cero, así
        # que la rama "True" (GRINS+IVAIns+TotIB) es la que efectivamente
        # se ejecuta siempre; la rama Else (que excluiría TotIB del
        # total) es alcanzable sólo con TotIB=-1 exacto, que no ocurre
        # con dinero real. Se implementa el resultado real: total
        # SIEMPRE incluye la percepción de IIBB.
        neto = bruto + iva + ib
        form.addRow("Tot. Bruto :", QLabel(f"$ {format_decimal(bruto)}"))
        form.addRow("IVA Ins. :", QLabel(f"$ {format_decimal(iva)}"))
        form.addRow("I.B. Bs.As. :", QLabel(f"$ {format_decimal(ib)}"))
        lbl_total = QLabel(f"$ {format_decimal(neto)}")
        lbl_total.setStyleSheet("font-weight: bold;")
        form.addRow("Tot. Neto :", lbl_total)
        return grupo
