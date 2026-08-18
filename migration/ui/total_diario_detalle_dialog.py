"""Diálogo de detalle de un día — migración de `TotFact.frm Sub
MuestraDetalle` (Picture1, líneas 1959-2057).

El legacy alterna 3 vistas (Importes/Unidades/Comprobantes) con un mismo
panel mostrando/ocultando `Frame7`/`Frame12`/`Frame8` según el radio-
button `Option1` elegido (o F1/F2/F3) — acá se usan pestañas (`QTabWidget`),
más natural en una UI moderna, mismo contenido.

**Bug real encontrado y corregido** (verificado sistemáticamente
comparando la posición `Top` de cada `Label` de caption contra el
`Label` de valor que el código realmente escribe en esa posición —
único caso de todo el formulario donde no coinciden): en el grupo
"Totales" del panel Importes, `MuestraDetalle` escribe
`Label24.Caption = PVta` y `Label25.Caption = PCos`, pero por posición
en pantalla `Label24` está bajo el caption "Precio Costo" y `Label25`
bajo "Precio Venta" — Precio Venta y Precio Costo aparecen
INTERCAMBIADOS en el legacy. Acá van cada uno bajo su etiqueta correcta.

**Etiquetas renombradas, sin cambiar el cálculo** (ver
`DetalleTotalDia` en services.py): los 2 grupos "Unidades Promedio x
Día" no calculan nada "por día" — uno es precio promedio por unidad
vendida, el otro cantidad promedio por comprobante.
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
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from migration.decimals import format_decimal
from migration.services import DetalleTotalDia


def _fila_moneda(form: QFormLayout, etiqueta: str, valor: Decimal | None) -> None:
    form.addRow(etiqueta, QLabel(f"$ {format_decimal(valor or Decimal('0'))}"))


def _fila_entero(form: QFormLayout, etiqueta: str, valor: int | None) -> None:
    form.addRow(etiqueta, QLabel(str(valor or 0)))


def _fila_decimal(form: QFormLayout, etiqueta: str, valor: Decimal | None) -> None:
    """Para promedios que NO son dinero (ej. unidades por comprobante)
    — sin el prefijo "$" de `_fila_moneda`."""
    form.addRow(etiqueta, QLabel(format_decimal(valor or Decimal("0"))))


class TotalDiarioDetalleDialog(QDialog):
    def __init__(self, detalle: DetalleTotalDia, parent: QWidget | None = None):
        super().__init__(parent)
        self.detalle = detalle
        fecha = detalle.totales.FECHA
        self.setWindowTitle(f"Totales del día {fecha.strftime('%d/%m/%Y') if fecha else ''}")
        self.resize(620, 480)

        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        tabs.addTab(self._tab_importes(), "Importes")
        tabs.addTab(self._tab_unidades(), "Unidades")
        tabs.addTab(self._tab_comprobantes(), "Comprobantes")
        layout.addWidget(tabs)

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    # ------------------------------------------------------------------
    def _tab_importes(self) -> QWidget:
        t = self.detalle.totales
        contenedor = QWidget()
        layout = QVBoxLayout(contenedor)

        grupo_cc = QGroupBox("Cta. Cte.")
        form_cc = QFormLayout(grupo_cc)
        _fila_moneda(form_cc, 'Facturas "A" :', t.PESPA)
        _fila_moneda(form_cc, 'Facturas "B" :', t.PESPB)
        _fila_moneda(form_cc, 'Facturas "C" :', t.PESPC)
        layout.addWidget(grupo_cc)

        grupo_mo = QGroupBox("Mostrador")
        form_mo = QFormLayout(grupo_mo)
        _fila_moneda(form_mo, 'Facturas "A" :', t.PESPMOSTA)
        _fila_moneda(form_mo, 'Facturas "B" :', t.PESPMOSTB)
        _fila_moneda(form_mo, 'Facturas "C" :', t.PESPMOSTC)
        layout.addWidget(grupo_mo)

        grupo_exp = QGroupBox("Exportación")
        form_exp = QFormLayout(grupo_exp)
        _fila_moneda(form_exp, 'Factura "E" :', t.PESPEXPA)
        layout.addWidget(grupo_exp)

        grupo_prom = QGroupBox("Precio Promedio x Unidad")
        form_prom = QFormLayout(grupo_prom)
        _fila_moneda(form_prom, "Por Cta. Cte. :", self.detalle.precio_prom_cta_cte)
        _fila_moneda(form_prom, "Por Mostrador :", self.detalle.precio_prom_mostrador)
        _fila_moneda(form_prom, 'Por Tipo "C" :', self.detalle.precio_prom_tipo_c)
        layout.addWidget(grupo_prom)

        grupo_tot = QGroupBox("Totales")
        form_tot = QFormLayout(grupo_tot)
        _fila_moneda(form_tot, "Precio Venta :", t.PVTA)
        _fila_moneda(form_tot, "Precio Costo :", t.PCOS)
        lbl_venta_real = QLabel(f"$ {format_decimal(t.PESP or Decimal('0'))}")
        lbl_venta_real.setStyleSheet("font-weight: bold;")
        form_tot.addRow("Venta Real :", lbl_venta_real)
        layout.addWidget(grupo_tot)

        layout.addStretch()
        return contenedor

    def _tab_unidades(self) -> QWidget:
        t = self.detalle.totales
        contenedor = QWidget()
        layout = QVBoxLayout(contenedor)

        grupo_cc = QGroupBox("Cta. Cte.")
        form_cc = QFormLayout(grupo_cc)
        _fila_entero(form_cc, 'Facturas "A" :', t.UNIDA)
        _fila_entero(form_cc, 'Facturas "B" :', t.UNIDB)
        _fila_entero(form_cc, 'Facturas "C" :', t.UNIDC)
        _fila_entero(form_cc, "Devolución :", t.DEV)
        layout.addWidget(grupo_cc)

        grupo_mo = QGroupBox("Mostrador")
        form_mo = QFormLayout(grupo_mo)
        _fila_entero(form_mo, 'Facturas "A" :', t.UNMOSTA)
        _fila_entero(form_mo, 'Facturas "B" :', t.UNMOSTB)
        _fila_entero(form_mo, 'Facturas "C" :', t.UNMOSTC)
        _fila_entero(form_mo, "Devolución :", t.MOSTDEV)
        layout.addWidget(grupo_mo)

        grupo_exp = QGroupBox("Exportación")
        form_exp = QFormLayout(grupo_exp)
        _fila_entero(form_exp, 'Factura "E" :', t.UNEXPA)
        _fila_entero(form_exp, "Devolución :", t.EXPDEV)
        layout.addWidget(grupo_exp)

        grupo_prom = QGroupBox("Unidades Promedio x Comprobante")
        form_prom = QFormLayout(grupo_prom)
        _fila_decimal(form_prom, "Por Cta. Cte. :", self.detalle.unid_prom_cta_cte)
        _fila_decimal(form_prom, "Por Mostrador :", self.detalle.unid_prom_mostrador)
        _fila_decimal(form_prom, 'Por Tipo "C" :', self.detalle.unid_prom_tipo_c)
        layout.addWidget(grupo_prom)

        layout.addStretch()
        return contenedor

    def _tab_comprobantes(self) -> QWidget:
        t = self.detalle.totales
        contenedor = QWidget()
        layout = QVBoxLayout(contenedor)

        grupo_cc = QGroupBox("Cta. Cte.")
        form_cc = QFormLayout(grupo_cc)
        _fila_entero(form_cc, 'Facturas "A" :', t.FACA)
        _fila_entero(form_cc, 'Facturas "B" :', t.FACB)
        _fila_entero(form_cc, 'Facturas "C" :', t.FACC)
        _fila_entero(form_cc, 'N/Créd. "A" :', t.NCA)
        _fila_entero(form_cc, 'N/Créd. "B" :', t.NCB)
        _fila_entero(form_cc, 'N/Créd. "C" :', t.NCC)
        _fila_entero(form_cc, 'N/Déb. "A" :', t.NDA)
        _fila_entero(form_cc, 'N/Déb. "B" :', t.NDB)
        layout.addWidget(grupo_cc)

        grupo_mo = QGroupBox("Mostrador")
        form_mo = QFormLayout(grupo_mo)
        _fila_entero(form_mo, 'Facturas "A" :', t.MOSTA)
        _fila_entero(form_mo, 'Facturas "B" :', t.MOSTB)
        _fila_entero(form_mo, 'Facturas "C" :', t.MOSTC)
        _fila_entero(form_mo, 'Créditos "A" :', t.NCMOSTA)
        _fila_entero(form_mo, 'Créditos "B" :', t.NCMOSTB)
        _fila_entero(form_mo, 'Créditos "C" :', t.NCMOSTC)
        layout.addWidget(grupo_mo)

        grupo_exp = QGroupBox("Exportación")
        form_exp = QFormLayout(grupo_exp)
        _fila_entero(form_exp, 'Facturas "E" :', t.EXPA)
        _fila_entero(form_exp, 'Créditos "E" :', t.EXPDEV)
        layout.addWidget(grupo_exp)

        layout.addStretch()
        return contenedor
