"""Generación de PDF de Factura Electrónica — reemplaza el circuito de
impresión de `EmiFact.frm` (VSPrinter/formulario preimpreso, fuera de
alcance) y el QR externo (`Shell("qrcode.exe ...")`, bug real ya
corregido en `migration/afip.py`).

Usa **ReportLab puro** (`reportlab.pdfgen.canvas` + `reportlab.graphics.
barcode.qr.QrCodeWidget`) — el QR se dibuja nativamente dentro del PDF,
sin depender de un generador de imágenes externo ni de una librería de
QR aparte (CLAUDE.md: "Generación de PDF automáticos... mediante
ReportLab").

Layout de una sola página (hasta 25 renglones — mismo límite real que
`DetFact.frm`, ver `factura_renglon.py`): encabezado (empresa/Factura/
Letra/Número/Fecha), datos del cliente, tabla de renglones, totales, y
el bloque de CAE + Vencimiento + QR que exige AFIP en todo comprobante
electrónico.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .decimals import format_decimal
from .repository import ETIQUETAS_TIPO_CTASCTE
from .services import AplicacionPago, PagoCheque, PagoRetencion, RenglonEmision, TotalFactura

# Réplica de CIVA_OPCIONES (`cliente_detalle_dialog.py`) — texto para
# imprimir la Condición de IVA del receptor.
CONDICIONES_IVA = {
    1: "IVA Responsable Inscripto",
    2: "IVA Responsable No Inscripto",
    3: "Consumidor Final",
    4: "IVA Exento",
    5: "Monotributo",
}

DIR_SALIDA_DEFAULT = Path("migration") / "pdf_output"


@dataclass
class DatosFacturaPDF:
    """Todo lo que necesita el PDF — ya resuelto por el llamador
    (`FacturadorWindow`), este módulo sólo dibuja.

    `cae`/`cae_vencimiento`/`qr_url` son `None` para el **boceto**
    (pedido del usuario, 2026-08-18: "al apretar emitir, muestre un
    boceto en pdf de la factura... Al apretar grabar, pregunte si se
    quiere grabar" — recién ahí, si confirma, se pide el CAE a AFIP).
    Como el CAE/QR sólo existen DESPUÉS de que AFIP aprueba, el boceto
    se genera y se muestra ANTES de tocar AFIP para nada — se dibuja un
    sello "BORRADOR" en su lugar, ver `generar_pdf_factura`."""

    letra: str
    punto_venta: int
    numero: int
    fecha: date
    cliente_codigo: int
    cliente_nombre: str
    cliente_cuit: str
    cliente_civa: int
    cliente_domicilio: str
    renglones: list[RenglonEmision]
    total: TotalFactura
    cae: Optional[str] = None
    cae_vencimiento: Optional[date] = None
    qr_url: Optional[str] = None
    empresa_nombre: str = "ALESTEL SRL"
    empresa_cuit: str = "33703467909"
    en_dolares: bool = False


def _nombre_archivo(datos: DatosFacturaPDF) -> str:
    sufijo = "_BORRADOR" if datos.cae is None else ""
    return f"Factura_{datos.letra}_{datos.punto_venta:04d}-{datos.numero:08d}{sufijo}.pdf"


def generar_pdf_factura(datos: DatosFacturaPDF, directorio_salida: Optional[Path] = None) -> Path:
    """Genera el PDF y lo guarda en disco (`directorio_salida`, por
    defecto `migration/pdf_output/`) — réplica de "Resguardo y Gestión"
    del CLAUDE.md. Devuelve la ruta del archivo creado."""
    directorio = Path(directorio_salida) if directorio_salida is not None else DIR_SALIDA_DEFAULT
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / _nombre_archivo(datos)

    c = canvas.Canvas(str(ruta), pagesize=A4)
    ancho, alto = A4
    margen = 15 * mm
    y = alto - margen

    # --- Encabezado ---------------------------------------------------
    c.setFont("Helvetica-Bold", 14)
    c.drawString(margen, y, datos.empresa_nombre)
    c.setFont("Helvetica", 9)
    c.drawString(margen, y - 14, f"CUIT: {datos.empresa_cuit}")

    letra_x = ancho / 2
    c.setFont("Helvetica-Bold", 22)
    c.rect(letra_x - 12, y - 24, 24, 24)
    c.drawCentredString(letra_x, y - 18, datos.letra)

    c.setFont("Helvetica-Bold", 13)
    c.drawRightString(ancho - margen, y, f"FACTURA {datos.letra}")
    c.setFont("Helvetica", 10)
    c.drawRightString(ancho - margen, y - 14, f"Nº {datos.punto_venta:04d}-{datos.numero:08d}")
    c.drawRightString(ancho - margen, y - 26, f"Fecha: {datos.fecha.strftime('%d/%m/%Y')}")
    if datos.en_dolares:
        c.drawRightString(ancho - margen, y - 38, "Comprobante en DÓLARES")

    y -= 45
    c.line(margen, y, ancho - margen, y)

    # --- Cliente --------------------------------------------------------
    y -= 16
    c.setFont("Helvetica-Bold", 10)
    c.drawString(margen, y, "Cliente")
    y -= 13
    c.setFont("Helvetica", 9)
    c.drawString(margen, y, f"Código: {datos.cliente_codigo}  —  {datos.cliente_nombre}")
    y -= 12
    c.drawString(margen, y, f"CUIT: {datos.cliente_cuit or 's/d'}")
    c.drawString(margen + 90 * mm, y, CONDICIONES_IVA.get(datos.cliente_civa, "—"))
    if datos.cliente_domicilio:
        y -= 12
        c.drawString(margen, y, f"Domicilio: {datos.cliente_domicilio}")

    y -= 16
    c.line(margen, y, ancho - margen, y)

    # --- Tabla de renglones ---------------------------------------------
    y -= 14
    columnas = [
        (margen, "Código"),
        (margen + 45 * mm, "Descripción"),
        (margen + 115 * mm, "Cantidad"),
        (margen + 140 * mm, "Precio Unit."),
        (margen + 168 * mm, "Importe"),
    ]
    c.setFont("Helvetica-Bold", 8)
    for x, titulo in columnas:
        c.drawString(x, y, titulo)
    y -= 4
    c.line(margen, y, ancho - margen, y)

    c.setFont("Helvetica", 8)
    for renglon in datos.renglones:
        y -= 12
        if y < 70 * mm:  # deja lugar para totales + CAE/QR — 25 renglones reales entran cómodos
            break
        cantidad = renglon.cantidad_unidades if renglon.cantidad_unidades > 0 else renglon.mtr
        cod = f"{renglon.cod1}/{renglon.cod2}" if renglon.cod2 else renglon.cod1
        c.drawString(columnas[0][0], y, cod[:18])
        c.drawString(columnas[1][0], y, renglon.descripcion[:38])
        c.drawRightString(columnas[2][0] + 18 * mm, y, format_decimal(cantidad))
        c.drawRightString(columnas[3][0] + 22 * mm, y, format_decimal(renglon.precio_unitario))
        c.drawRightString(columnas[4][0] + 22 * mm, y, format_decimal(renglon.importe))

    y -= 10
    c.line(margen, y, ancho - margen, y)

    # --- Totales ----------------------------------------------------------
    y -= 16
    c.setFont("Helvetica", 9)
    etiquetas_totales = [
        ("Subtotal", datos.total.bruto),
        ("Descuento", datos.total.descuento),
        ("Neto Gravado", datos.total.neto_gravado),
        ("IVA Insc. 21%", datos.total.iva),
        ("Percepción IIBB", datos.total.percepcion_iibb),
    ]
    for etiqueta, valor in etiquetas_totales:
        c.drawString(ancho - margen - 70 * mm, y, etiqueta)
        c.drawRightString(ancho - margen, y, f"$ {format_decimal(valor)}")
        y -= 12

    c.setFont("Helvetica-Bold", 12)
    c.drawString(ancho - margen - 70 * mm, y, "TOTAL")
    c.drawRightString(ancho - margen, y, f"$ {format_decimal(datos.total.total)}")

    # --- CAE + Vencimiento + QR (obligatorio AFIP) -------------------------
    # Si todavía no hay CAE (boceto — pedido del usuario, 2026-08-18: se
    # muestra ANTES de pedirle nada a AFIP, ver docstring de
    # `DatosFacturaPDF`) no hay CAE/QR reales que dibujar: un sello
    # "BORRADOR" en su lugar, nada de datos legales inventados.
    pie_y = 30 * mm
    c.setFont("Helvetica", 9)
    if datos.cae is not None and datos.cae_vencimiento is not None and datos.qr_url is not None:
        c.drawString(margen, pie_y + 14, f"CAE: {datos.cae}")
        c.drawString(margen, pie_y, f"Vencimiento CAE: {datos.cae_vencimiento.strftime('%d/%m/%Y')}")

        qr_widget = qr.QrCodeWidget(datos.qr_url)
        bounds = qr_widget.getBounds()
        ancho_qr = bounds[2] - bounds[0]
        alto_qr = bounds[3] - bounds[1]
        tamano = 25 * mm
        dibujo = Drawing(tamano, tamano, transform=[tamano / ancho_qr, 0, 0, tamano / alto_qr, 0, 0])
        dibujo.add(qr_widget)
        renderPDF.draw(dibujo, c, ancho - margen - tamano, pie_y - 5 * mm)
    else:
        c.drawString(margen, pie_y + 14, "CAE: (pendiente — todavía no se solicitó a AFIP)")
        _dibujar_sello_borrador(c, ancho, alto)

    c.showPage()
    c.save()
    return ruta


def _dibujar_sello_borrador(c: canvas.Canvas, ancho: float, alto: float) -> None:
    """Sello diagonal "BORRADOR" — para que el boceto que se muestra
    antes de pedir el CAE (`FacturadorWindow._on_emitir`) nunca se
    confunda con el comprobante fiscal real."""
    c.saveState()
    c.setFillColorRGB(0.85, 0.2, 0.2, alpha=0.35)
    c.setFont("Helvetica-Bold", 60)
    c.translate(ancho / 2, alto / 2)
    c.rotate(35)
    c.drawCentredString(0, 0, "BORRADOR")
    c.restoreState()


# ---------------------------------------------------------------------------
# PDF de Recibo — EmisionReciboService, sin CAE/QR (Recibo no es
# comprobante fiscal electrónico, ver docstring de EmisionReciboService)
# ---------------------------------------------------------------------------


@dataclass
class DatosReciboPDF:
    """Todo lo que necesita el PDF del Recibo — ya resuelto por el
    llamador (`ReciboWindow`), este módulo sólo dibuja. Mucho más simple
    que `DatosFacturaPDF`: sin Letra/CAE/QR, y con el desglose de medios
    de pago en vez de un detalle de renglones."""

    numero: int
    fecha: date
    correlativo: int
    cliente_codigo: int
    cliente_nombre: str
    cliente_cuit: str
    aplicaciones: list[AplicacionPago]
    anticipo: Decimal
    importe_efectivo: Decimal
    cheques: list[PagoCheque]
    retenciones: list[PagoRetencion]
    total_pago: Decimal
    empresa_nombre: str = "ALESTEL SRL"
    empresa_cuit: str = "33703467909"


# Media hoja A4 (feedback del usuario, 2026-08-15: "el recibo impreso es
# medio A4") — un Recibo es un documento corto (unos pocos renglones
# aplicados + medios de pago), no tiene sentido imprimirlo en una hoja
# A4 completa. El bug real que además hacía parecer el PDF "vacío" en
# el medio (a full A4 o a media hoja): "Correlativo" se dibujaba en una
# posición ABSOLUTA fija (`20mm` desde el borde inferior) en vez de
# relativa a dónde terminó el contenido — con pocos renglones, dejaba un
# hueco enorme entre el Total y el Correlativo. Corregido: todo el
# layout es relativo, línea tras línea, sin posiciones absolutas.
MEDIA_A4 = (A4[0], A4[1] / 2)


def _nombre_archivo_recibo(datos: DatosReciboPDF) -> str:
    return f"Recibo_{datos.numero:08d}.pdf"


def generar_pdf_recibo(datos: DatosReciboPDF, directorio_salida: Optional[Path] = None) -> Path:
    """Genera el PDF del Recibo (media hoja A4) y lo guarda en disco.
    Devuelve la ruta.

    **Orden Vista Previa → Grabar** (decisión del usuario, 2026-08-15):
    este generador se llama ANTES de `EmisionReciboService.
    emitir_recibo()` (para la vista previa) — el número que lleva el
    PDF es el que TENDRÍA el comprobante, no necesariamente uno ya
    persistido. Ver `ReciboWindow._on_emitir`."""
    directorio = Path(directorio_salida) if directorio_salida is not None else DIR_SALIDA_DEFAULT
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / _nombre_archivo_recibo(datos)

    c = canvas.Canvas(str(ruta), pagesize=MEDIA_A4)
    ancho, alto = MEDIA_A4
    margen = 10 * mm
    y = alto - margen

    # --- Encabezado ---------------------------------------------------
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margen, y, datos.empresa_nombre)
    c.setFont("Helvetica", 8)
    c.drawString(margen, y - 11, f"CUIT: {datos.empresa_cuit}")

    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(ancho - margen, y, "RECIBO")
    c.setFont("Helvetica", 9)
    c.drawRightString(ancho - margen, y - 11, f"Nº {datos.numero:08d}")
    c.drawRightString(ancho - margen, y - 21, f"Fecha: {datos.fecha.strftime('%d/%m/%Y')}")

    y -= 30
    c.line(margen, y, ancho - margen, y)

    # --- Cliente --------------------------------------------------------
    y -= 11
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margen, y, "Recibí de")
    y -= 10
    c.setFont("Helvetica", 8)
    c.drawString(margen, y, f"Código: {datos.cliente_codigo}  —  {datos.cliente_nombre}")
    y -= 9
    c.drawString(margen, y, f"CUIT: {datos.cliente_cuit or 's/d'}")

    y -= 11
    c.line(margen, y, ancho - margen, y)

    # --- Aplicado a (comprobantes pagados) -------------------------------
    y -= 10
    c.setFont("Helvetica-Bold", 8)
    c.drawString(margen, y, "Comprobante")
    c.drawString(margen + 50 * mm, y, "Fecha")
    c.drawRightString(margen + 125 * mm, y, "Aplicado")
    c.drawRightString(margen + 160 * mm, y, "Descuento")
    y -= 3
    c.line(margen, y, ancho - margen, y)

    # Reserva de espacio para lo que viene DESPUÉS de esta lista (medios
    # de pago + total + correlativo) — con la página ya acotada a media
    # A4, hay que dejar de agregar renglones antes de quedarse sin lugar
    # real, no sólo antes del borde de la hoja.
    lineas_pie_estimadas = 3 + len(datos.cheques) + len(datos.retenciones) + (1 if datos.importe_efectivo > 0 else 0)
    limite_inferior = margen + lineas_pie_estimadas * 9.5 + 20

    c.setFont("Helvetica", 7.5)
    for aplicacion in datos.aplicaciones:
        y -= 9.5
        if y < limite_inferior:
            break
        comprobante = aplicacion.comprobante
        etiqueta = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
        fecha_txt = comprobante.FECHA.strftime("%d/%m/%Y") if comprobante.FECHA else "—"
        c.drawString(margen, y, f"{etiqueta} {comprobante.CPBTE}")
        c.drawString(margen + 50 * mm, y, fecha_txt)
        c.drawRightString(margen + 125 * mm, y, format_decimal(aplicacion.importe_aplicado))
        c.drawRightString(margen + 160 * mm, y, format_decimal(aplicacion.descuento) if aplicacion.descuento else "")

    if datos.anticipo > 0:
        y -= 9.5
        c.drawString(margen, y, "Pago a Cuenta (Anticipo)")
        c.drawRightString(margen + 125 * mm, y, format_decimal(datos.anticipo))

    y -= 8
    c.line(margen, y, ancho - margen, y)

    # --- Medios de pago ---------------------------------------------------
    y -= 12
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margen, y, "Medios de pago")
    y -= 10
    c.setFont("Helvetica", 8)
    if datos.importe_efectivo > 0:
        c.drawString(margen, y, "Efectivo")
        c.drawRightString(ancho - margen, y, f"$ {format_decimal(datos.importe_efectivo)}")
        y -= 9.5
    for pago_cheque in datos.cheques:
        c.drawString(
            margen,
            y,
            f"Cheque Nº {pago_cheque.nro_cheque} — Bco. {pago_cheque.banco} "
            f"— Vto. {pago_cheque.fecha_vencimiento.strftime('%d/%m/%Y')}",
        )
        c.drawRightString(ancho - margen, y, f"$ {format_decimal(pago_cheque.importe)}")
        y -= 9.5
    for retencion in datos.retenciones:
        c.drawString(margen, y, retencion.concepto or f"Retención/Anticipo tipo {retencion.tipreg}")
        c.drawRightString(ancho - margen, y, f"$ {format_decimal(retencion.importe)}")
        y -= 9.5

    y -= 4
    c.line(margen, y, ancho - margen, y)

    # --- Total + Correlativo ---------------------------------------------
    # Ambos relativos a `y` (no a una posición absoluta desde el borde
    # de la hoja) — ver comentario de MEDIA_A4 más arriba.
    y -= 15
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margen, y, "TOTAL RECIBIDO")
    c.drawRightString(ancho - margen, y, f"$ {format_decimal(datos.total_pago)}")

    y -= 12
    c.setFont("Helvetica", 7)
    c.drawString(margen, y, f"Correlativo: {datos.correlativo}")

    c.showPage()
    c.save()
    return ruta


# ---------------------------------------------------------------------
# Listados varios (migración de Listados.frm) — un único generador
# genérico de tabla paginada, reusado por los ~13 reportes en vez de
# repetir el dibujo manual con `canvas` de arriba: acá la mayoría de las
# columnas del legacy no importan (fuentes/anchos del RTF de VSPrinter),
# lo único que hay que replicar fielmente son los datos/filtros — así
# que se arma con Platypus (`SimpleDocTemplate`/`Table`), que pagina
# solo, en vez de reimplementar el salto de página a mano como hacía
# cada `Sub LisXxx` del legacy (`If Cont > NNNN Then ... NewPage`).
# ---------------------------------------------------------------------

_ESTILOS = getSampleStyleSheet()


def generar_pdf_listado(
    *,
    titulo: str,
    subtitulo: str,
    columnas: list[str],
    filas: list[list[str]],
    columnas_derecha: tuple[int, ...] = (),
    pie: Optional[list[tuple[str, str]]] = None,
    apaisado: bool = True,
    nombre_archivo: str = "listado.pdf",
    directorio_salida: Optional[Path] = None,
) -> Path:
    """Genera un PDF tabular genérico (título + subtítulo + tabla +
    totales opcionales al pie) — reusado por todos los listados de
    `migration.services.ListadosService`.

    `columnas_derecha`: índices de columna a alinear a la derecha
    (importes/cantidades), mismo criterio que `TablaBusqueda` en la UI.
    `pie`: lista de (etiqueta, valor) para la fila de totales al final.
    """
    directorio = Path(directorio_salida) if directorio_salida is not None else DIR_SALIDA_DEFAULT
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / nombre_archivo

    tamano_pagina = landscape(A4) if apaisado else A4
    doc = SimpleDocTemplate(
        str(ruta), pagesize=tamano_pagina,
        leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
        title=titulo,
    )
    ancho_util = tamano_pagina[0] - 24 * mm

    # Cabecera de impresión (pedido del usuario, 2026-08-18): "ALESTEL
    # SRL" arriba a la izquierda, fecha arriba a la derecha, el título
    # del listado centrado, y debajo de éste (también centrado) la
    # selección del operador (Cliente/Zona/Todos/etc. — lo que hasta
    # ahora era `subtitulo`, mismo texto, sólo cambia dónde y cómo se
    # dibuja).
    estilo_empresa = ParagraphStyle("EmpresaListado", parent=_ESTILOS["Normal"], fontSize=11, fontName="Helvetica-Bold", alignment=TA_LEFT)
    estilo_fecha = ParagraphStyle("FechaListado", parent=_ESTILOS["Normal"], fontSize=9, alignment=TA_RIGHT)
    estilo_titulo = ParagraphStyle(
        "TituloListado", parent=_ESTILOS["Heading1"], fontSize=14, alignment=TA_CENTER, spaceBefore=6, spaceAfter=2
    )
    estilo_subtitulo = ParagraphStyle("SubtituloListado", parent=_ESTILOS["Normal"], fontSize=9, alignment=TA_CENTER, spaceAfter=8)

    fila_encabezado = Table(
        [[Paragraph("ALESTEL SRL", estilo_empresa), Paragraph(date.today().strftime("%d/%m/%Y"), estilo_fecha)]],
        colWidths=[ancho_util / 2, ancho_util / 2],
    )
    fila_encabezado.setStyle(
        TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)])
    )

    elementos = [
        fila_encabezado,
        Paragraph(titulo, estilo_titulo),
        Paragraph(subtitulo, estilo_subtitulo),
        Spacer(1, 4),
    ]

    datos_tabla = [columnas] + filas
    tabla = Table(datos_tabla, repeatRows=1)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#bbbbbb")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    for col in columnas_derecha:
        estilo.append(("ALIGN", (col, 0), (col, -1), "RIGHT"))
    tabla.setStyle(TableStyle(estilo))
    elementos.append(tabla)

    if pie:
        elementos.append(Spacer(1, 10))
        filas_pie = [[etiqueta, valor] for etiqueta, valor in pie]
        tabla_pie = Table(filas_pie, colWidths=[120 * mm, 40 * mm])
        tabla_pie.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ]
            )
        )
        elementos.append(tabla_pie)

    doc.build(elementos)
    return ruta
