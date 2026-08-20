"""Grilla reusable de renglones reales (`Fcestad1`) de una factura/NC/ND
puntual — migración de `CtaCte.frm Sub Command6_Click` ("Ver Items",
Picture2, líneas 1742-1792) / `VerFact.frm` ("Ver Items").

Usada tanto por el drill-down de Cuenta Corriente
(`CtaCteWindow._on_click_fila_extracto`, que ahora reusa directo
`FacturaEmitidaDetalleDialog` — antes tenía su propio `FacturaDetalleDialog`
más simple, unificado en una sola pantalla para no mantener 2 vistas de
lo mismo) como por `FacturaEmitidaDetalleDialog` (`VerFact.frm`)."""

from __future__ import annotations

from typing import Optional

from PyQt6.QtWidgets import QGroupBox, QVBoxLayout

from migration.decimals import format_decimal
from migration.models import FcivaVta
from migration.repository import RepositoryFactory

from .widgets import TablaBusqueda

COLUMNAS_ITEMS = [
    "Sección", "Pulg", "Mtr", "Milím.", "Telas", "Cant.", "P.Esp.", "P.Costo", "P.Venta", "Importe",
]

# Alto de fila/encabezado real de `TablaBusqueda` (`widgets.py`) — usado
# para calcular cuántas filas entran visibles sin scroll cuando se pide
# `filas_visibles` (ver más abajo).
_ALTO_FILA_TABLA_BUSQUEDA = 24
_ALTO_HEADER_TABLA_BUSQUEDA = 24


def construir_grupo_renglones(
    repos: RepositoryFactory,
    factura: FcivaVta,
    titulo: str = "Renglones",
    *,
    filas_visibles: Optional[int] = None,
) -> QGroupBox:
    """Renglones reales (`Fcestad1`) de una factura/NC/ND puntual, en un
    `QGroupBox` listo para insertar en cualquier diálogo.

    `filas_visibles`: si se pasa, fija un alto mínimo para que ese
    número de renglones se vea sin necesidad de scrollear (el resto
    scrollea normal, `TablaBusqueda` ya es una `QTableWidget` con su
    propio scroll) — pedido del usuario para `FacturaEmitidaDetalleDialog`
    (2026-08-19: "Agrandar los mas posible el panel de detalle como para
    poder ver 10 renglones y poder scrolear")."""
    grupo = QGroupBox(titulo)
    layout = QVBoxLayout(grupo)
    tabla = TablaBusqueda(COLUMNAS_ITEMS)
    if filas_visibles is not None:
        tabla.setMinimumHeight(filas_visibles * _ALTO_FILA_TABLA_BUSQUEDA + _ALTO_HEADER_TABLA_BUSQUEDA + 4)

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
    layout.addWidget(tabla, stretch=1)
    return grupo
