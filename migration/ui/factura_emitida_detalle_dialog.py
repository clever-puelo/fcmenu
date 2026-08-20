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

**Cabecera achicada + Detalle agrandado** (feedback del usuario,
2026-08-19: "hay que achicar los datos de cabecera para poder ver el
detalle... agrandar lo mas posible el panel de detalle como para poder
ver 10 renglones y poder scrolear") — Dirección/Localidad/Provincia en
una sola línea, Cond.IVA/CUIT en otra; en "Comprobante" Número/Fecha
juntos, Ítems/Unids. juntos, y los 4 importes en una sola línea — cada
grupo pasó de una lista vertical (`QFormLayout`, una fila por dato) a
filas horizontales armadas a mano, liberando el alto que gana la grilla
de renglones (`construir_grupo_renglones(..., filas_visibles=10)`).

**Botón "Imprimir"** (mismo pedido): genera un PDF con el mismo formato
del boceto de Factura (`pdf.generar_pdf_factura`) reconstruido a partir
de los renglones reales (`Fcestad1`) y los totales guardados en
`FCIVAVTA` — y lo muestra en `PdfPreviewDialog` para guardar/imprimir.
**Limitación real del esquema, no un olvido**: el CAE/QR de AFIP nunca
se persiste en la base (sólo vive en el PDF generado al momento de
emitir, ver `migration/afip.py`/`FacturadorWindow`) — así que esta
reimpresión SIEMPRE sale sin CAE/QR (con el mismo sello "BORRADOR" que
usa el boceto), no es un comprobante fiscal reimprimible 1:1. Igual de
esperable: no hay Despacho por renglón guardado en `Fcestad1` (ver
`NotaCreditoMercaderiaWindow`, mismo límite) ni el "%Descuento" por
línea (no se guarda desglosado), así que esas columnas salen vacías.
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtWidgets import (
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.decimals import format_decimal
from migration.models import FcivaVta
from migration.pdf import DatosFacturaPDF, generar_pdf_factura
from migration.provincias import nombre_provincia
from migration.repository import RepositoryFactory
from migration.services import RenglonEmision, TotalFactura

from .cliente_detalle_dialog import CIVA_OPCIONES
from .factura_detalle_dialog import construir_grupo_renglones
from .widgets import crear_recuadro_destacado
from .pdf_preview_dialog import PdfPreviewDialog

CIVA_LABELS = dict(CIVA_OPCIONES)

# 10 renglones visibles sin scroll (pedido explícito del usuario).
FILAS_VISIBLES_DETALLE = 10


def _fila(*pares: tuple[str, str]) -> QHBoxLayout:
    """Una línea horizontal con varios "Etiqueta: valor" — reemplaza el
    `QFormLayout` de una fila por dato (demasiado alto para lo que pide
    esta pantalla, ver docstring del módulo)."""
    fila = QHBoxLayout()
    for etiqueta, valor in pares:
        lbl_etiqueta = QLabel(etiqueta)
        lbl_etiqueta.setStyleSheet("font-weight: bold;")
        fila.addWidget(lbl_etiqueta)
        fila.addWidget(QLabel(valor))
        fila.addSpacing(16)
    fila.addStretch()
    return fila


class FacturaEmitidaDetalleDialog(QDialog):
    def __init__(self, repos: RepositoryFactory, factura: FcivaVta, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.factura = factura
        self.cliente = repos.cliente().by_codigo(factura.CLTE) if factura.CLTE else None

        ptovta_txt = str(factura.PTOVTA or 0).rjust(4, "0")
        cpbte_txt = str(factura.CPBTE or 0).rjust(8, "0")
        self.setWindowTitle(f"Comprobante {(factura.LETRA or '').strip()} {ptovta_txt}-{cpbte_txt}")
        # Más alto que antes (760→900) — el espacio ganado por achicar
        # la cabecera va directo al Detalle, no a agrandar la ventana
        # entera de más (pedido: "agrandar lo mas posible el panel de
        # detalle", no la ventana en sí).
        self.resize(820, 720)

        self._construir_ui()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.addWidget(self._armar_cliente())
        layout.addWidget(self._armar_comprobante())
        layout.addWidget(
            construir_grupo_renglones(self.repos, self.factura, filas_visibles=FILAS_VISIBLES_DETALLE),
            stretch=1,
        )
        layout.addWidget(self._armar_totales())

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        btn_imprimir = QPushButton("Imprimir...")
        btn_imprimir.clicked.connect(self._on_imprimir)
        fila_botones.addWidget(btn_imprimir)
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.accept)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

    def _armar_cliente(self) -> QGroupBox:
        f = self.factura
        c = self.cliente
        grupo = QGroupBox("Cliente")
        layout = QVBoxLayout(grupo)

        codigo = c.CODIGO if c else f.CLTE
        nombre = ((c.NOMB if c else f.NOMB) or "").strip()
        # Tel./Email en la línea del nombre (pedido del usuario,
        # 2026-08-20) — igual que Dirección/Localidad más abajo, sólo
        # están disponibles con Cliente real (el snapshot de FCIVAVTA no
        # los guarda).
        telefono = (c.TEL1 or "").strip() if c is not None else ""
        email = (c.EMAIL or "").strip() if c is not None else ""
        layout.addLayout(
            _fila(
                ("Código :", str(codigo or "")),
                ("Nombre :", nombre),
                ("Tel. :", telefono),
                ("Email :", email),
            )
        )

        if c is not None:
            direccion = (c.DIR or "").strip()
            localidad = f"({(c.CP or '').strip()}) {(c.LOC or '').strip()}"
            pcia_codigo = c.PCIA
            civa_codigo = c.CIVA
            cuit = c.CUIT
        else:
            direccion = ""
            localidad = ""
            civa_texto = (f.CIVA or "").strip()
            pcia_codigo = f.PCIA
            civa_codigo = int(civa_texto) if civa_texto.isdigit() else None
            cuit = f.CUIT

        # Dirección/Localidad/Provincia en la misma línea (pedido del
        # usuario, 2026-08-19).
        layout.addLayout(
            _fila(
                ("Dirección :", direccion),
                ("Localidad :", localidad),
                ("Provincia :", nombre_provincia(pcia_codigo)),
            )
        )
        # Cond.IVA/CUIT en la misma línea (ídem).
        layout.addLayout(_fila(("Cond. IVA :", CIVA_LABELS.get(civa_codigo, "")), ("CUIT :", (cuit or "").strip())))
        return grupo

    def _armar_comprobante(self) -> QGroupBox:
        f = self.factura
        grupo = QGroupBox("Comprobante")
        layout = QVBoxLayout(grupo)

        ptovta_txt = str(f.PTOVTA or 0).rjust(4, "0")
        cpbte_txt = str(f.CPBTE or 0).rjust(8, "0")
        # Número/Fecha en la misma línea (pedido del usuario, 2026-08-19).
        layout.addLayout(
            _fila(
                ("Número :", f"{(f.LETRA or '').strip()} {ptovta_txt}-{cpbte_txt}"),
                ("Fecha :", f.FECHA.strftime("%d/%m/%Y") if f.FECHA else ""),
            )
        )
        # Ítems/Unids. en la misma línea, abajo (ídem).
        layout.addLayout(_fila(("Ítems :", str(f.ITEMS or 0)), ("Unids. :", str(f.TOTCAN or 0))))
        return grupo

    def _armar_totales(self) -> QGroupBox:
        """Panel "Totales" nuevo, entre el Detalle y los botones (pedido
        del usuario, 2026-08-20: sacar los importes del panel
        "Comprobante" — que queda achicado a Número/Fecha/Ítems/Unids. —
        y ponerlos acá, con el Tot. Neto resaltado en un recuadro a la
        derecha, mismo patrón que el TOTAL de `FacturadorWindow`)."""
        total = self._total_factura()
        grupo = QGroupBox("Totales")
        layout = QHBoxLayout(grupo)

        for etiqueta, valor in (
            ("Tot. Bruto :", total.neto_gravado),
            ("IVA Ins. :", total.iva),
            ("I.B. Bs.As. :", total.percepcion_iibb),
        ):
            lbl_etiqueta = QLabel(etiqueta)
            lbl_etiqueta.setStyleSheet("font-weight: bold;")
            layout.addWidget(lbl_etiqueta)
            layout.addWidget(QLabel(f"$ {format_decimal(valor)}"))
            layout.addSpacing(16)
        layout.addStretch()

        recuadro_total, lbl_total = crear_recuadro_destacado("Tot. Neto:")
        lbl_total.setText(f"$ {format_decimal(total.total)}")
        layout.addWidget(recuadro_total)
        return grupo

    # ------------------------------------------------------------------
    # Totales / reconstrucción para el PDF
    # ------------------------------------------------------------------
    def _total_factura(self) -> TotalFactura:
        """Reconstruye el desglose a partir de los agregados YA guardados
        en `FCIVAVTA` al momento de emitir — no se recalcula nada (sería
        recomputar sobre datos que pueden haber cambiado desde entonces,
        ej. alícuotas), sólo se etiquetan los mismos campos que ya usaba
        el panel "Comprobante" (`GRINS`/`IVAINS`/`TOTIB`), más `BON`
        (bonificación total) para completar `bruto`/`descuento`."""
        f = self.factura
        neto_gravado = f.GRINS or Decimal("0")
        iva = f.IVAINS or Decimal("0")
        ib = f.TOTIB or Decimal("0")
        descuento = f.BON or Decimal("0")
        # VerFact.frm:1234-1238 ("ElTotal") — ver docstring histórico:
        # el total real SIEMPRE incluye la percepción de IIBB.
        total = neto_gravado + iva + ib
        return TotalFactura(
            bruto=neto_gravado + descuento,
            descuento=descuento,
            neto_gravado=neto_gravado,
            iva=iva,
            percepcion_iibb=ib,
            total=total,
        )

    def _renglones_reales(self) -> list[RenglonEmision]:
        f = self.factura
        renglones_fcestad1 = self.repos.fcestad1().by_comprobante(
            int(f.TIPO) if f.TIPO is not None else 1, f.LETRA or "", f.PTOVTA or 0, f.CPBTE or 0
        )
        resultado = []
        for r in renglones_fcestad1:
            cod1 = (r.COD1 or "").strip()
            cod2 = (r.COD2 or "").strip()
            articulo = self.repos.articulo().by_cod1_cod2(cod1, cod2)
            descripcion = (articulo.DESCRI or "").strip() if articulo is not None else f"{cod1}/{cod2}"
            resultado.append(
                RenglonEmision(
                    cod1=cod1,
                    cod2=cod2,
                    descripcion=descripcion,
                    precio_unitario=r.PVTA or Decimal("0"),
                    importe=r.IMPTE or Decimal("0"),
                    pulg=r.PULG or Decimal("0"),
                    mtr=r.MTR or Decimal("0"),
                    milim=r.MILIM or 0,
                    telas=r.TELAS or 0,
                    cantidad_unidades=r.CANT or Decimal("0"),
                    # Sin Despacho por renglón — no se persiste en
                    # Fcestad1 (ver docstring del módulo).
                )
            )
        return resultado

    # ------------------------------------------------------------------
    # Imprimir
    # ------------------------------------------------------------------
    def _on_imprimir(self) -> None:
        f = self.factura
        c = self.cliente
        codigo = c.CODIGO if c else f.CLTE
        nombre = ((c.NOMB if c else f.NOMB) or "").strip()
        cuit = (c.CUIT if c else f.CUIT) or ""
        domicilio = (c.DIR or "").strip() if c is not None else ""
        if c is not None:
            civa_codigo = c.CIVA or 0
        else:
            civa_texto = (f.CIVA or "").strip()
            civa_codigo = int(civa_texto) if civa_texto.isdigit() else 0

        try:
            datos = DatosFacturaPDF(
                letra=(f.LETRA or "").strip(),
                punto_venta=f.PTOVTA or 0,
                numero=f.CPBTE or 0,
                fecha=f.FECHA,
                cliente_codigo=codigo or 0,
                cliente_nombre=nombre,
                cliente_cuit=cuit,
                cliente_civa=civa_codigo,
                cliente_domicilio=domicilio,
                renglones=self._renglones_reales(),
                total=self._total_factura(),
                # Sin CAE/QR: no se persisten en la base, ver docstring
                # del módulo — reimpresión sale con el mismo sello
                # "BORRADOR" que el boceto de Facturador.
            )
            ruta_pdf = generar_pdf_factura(datos)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Comprobante", f"No se pudo generar el PDF:\n{exc}")
            return

        ptovta_txt = str(f.PTOVTA or 0).rjust(4, "0")
        cpbte_txt = str(f.CPBTE or 0).rjust(8, "0")
        PdfPreviewDialog(
            ruta_pdf,
            titulo=f"Comprobante {(f.LETRA or '').strip()} {ptovta_txt}-{cpbte_txt}",
            parent=self,
        ).exec()
