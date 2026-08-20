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

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, SimpleDocTemplate, Table, TableStyle

from .decimals import format_decimal
from .fechas import formatear_fecha_corta
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
    # Texto "%Descuento" por renglón (mismo orden que `renglones`, ver
    # `FacturadorWindow._descuentos_por_renglon`) — réplica de la
    # columna "% Descuento" que arma `EmiFact.frm` (líneas 1044-1052)
    # concatenando `DtoXClte.DTO1..DTO5` con "+" (ej. "10+5"). Lista
    # vacía (default) para quien todavía no la calcula.
    descuentos_renglones: list[str] = field(default_factory=list)
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
    # Réplica fiel de `EmiFact.frm` (líneas 887-892, "Ttulo" de FG1 en el
    # impreso): Cant. | Detalle | % Descuento | Precio Unit. | Importe —
    # SIN columna de Código propia (el legacy nunca la imprime aparte,
    # el Código sólo se usa en pantalla) y con el Despacho pegado al
    # final del Detalle (líneas 1054-1061), no en columna aparte
    # (pedido del usuario, 2026-08-19: "El detalle lleva el despacho").
    # Encabezados alineados al mismo punto que sus valores (pedido del
    # usuario: "acomodar columna de valores" — antes los títulos se
    # dibujaban a la izquierda mientras los importes de abajo quedaban
    # alineados a la derecha, un desfasaje visual real).
    y -= 14
    x_cant_d = margen + 20 * mm  # ancla derecha
    x_detalle_i = margen + 22 * mm  # ancla izquierda
    x_desc_d = margen + 126 * mm
    x_precio_d = margen + 150 * mm
    x_importe_d = margen + 180 * mm

    c.setFont("Helvetica-Bold", 8)
    c.drawRightString(x_cant_d, y, "Cant.")
    c.drawString(x_detalle_i, y, "Detalle")
    c.drawRightString(x_desc_d, y, "%Descuento")
    c.drawRightString(x_precio_d, y, "Precio Unit.")
    c.drawRightString(x_importe_d, y, "Importe")
    y -= 4
    c.line(margen, y, ancho - margen, y)

    c.setFont("Helvetica", 8)
    descuentos_renglones = datos.descuentos_renglones or []
    for indice, renglon in enumerate(datos.renglones):
        y -= 12
        if y < 70 * mm:  # deja lugar para totales + CAE/QR — 25 renglones reales entran cómodos
            break
        cantidad = renglon.cantidad_unidades if renglon.cantidad_unidades > 0 else renglon.mtr
        detalle = renglon.descripcion.strip()
        if renglon.nrodesp_elegido:
            detalle = f"{detalle} — Desp. {renglon.nrodesp_elegido}"
        desc_texto = descuentos_renglones[indice] if indice < len(descuentos_renglones) else ""

        c.drawRightString(x_cant_d, y, format_decimal(cantidad))
        c.drawString(x_detalle_i, y, detalle[:46])
        c.drawRightString(x_desc_d, y, desc_texto)
        c.drawRightString(x_precio_d, y, format_decimal(renglon.precio_unitario))
        c.drawRightString(x_importe_d, y, format_decimal(renglon.importe))

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
# lo único que hay que replicar fielmente son los datos/filtros.
#
# **Paginado manual** (pedido del usuario, 2026-08-20 — título+fecha en
# cada hoja, pie con nro. de hoja, "Viene/Transporte" entre hojas,
# totales en recuadro bajo su columna, cantidad de renglones impresos):
# antes se armaba una única `Table` gigante y se dejaba que Platypus
# paginara solo (`repeatRows=1`), lo que alcanzaba para repetir el
# encabezado de columnas pero no para dibujar título/fecha/pie en cada
# hoja (Platypus sólo dibuja un flowable una vez) ni para insertar
# "Viene"/"Transporte" exactamente en el corte de cada hoja (no hay forma
# de saber de antemano dónde va a cortar Platypus). Acá se arma una
# `Table` por cada bloque de filas que entra en una página (separadas por
# `PageBreak()`), y el encabezado/pie de página los dibuja `_NumberedCanvas`
# directo con `canvas`, en cada hoja.
# ---------------------------------------------------------------------


def generar_pdf_listado(
    *,
    titulo: str,
    subtitulo: str,
    columnas: list[str],
    filas: list[list[str]],
    columnas_derecha: tuple[int, ...] = (),
    pie: Optional[list[tuple[str, str, int]]] = None,
    columna_transporte: Optional[int] = None,
    valores_transporte: Optional[list[Decimal]] = None,
    apaisado: bool = True,
    nombre_archivo: str = "listado.pdf",
    directorio_salida: Optional[Path] = None,
) -> Path:
    """Genera un PDF tabular genérico — reusado por todos los listados de
    `migration.services.ListadosService`. En cada hoja: "ALESTEL SRL" +
    fecha (`formatear_fecha_corta`) arriba, título+subtítulo centrados,
    "Hoja N de M" centrado abajo de todo.

    `columnas_derecha`: índices de columna a alinear a la derecha
    (importes/cantidades), mismo criterio que `TablaBusqueda` en la UI.

    `pie`: lista de `(etiqueta, valor, columna)` para los totales finales
    — `valor` ya formateado tal cual se quiere mostrar (con "$" si es un
    importe, igual que antes; no todos los totales son un importe puro,
    ej. "ND Rechazadas" de Comisiones combina cantidad + importe en un
    solo texto). Cada uno se dibuja en un recuadro bajo SU columna real,
    al pie de la tabla de la ÚLTIMA hoja (antes era una tablita aparte,
    sin relación con la columna de importe).

    `columna_transporte`/`valores_transporte` (opcionales — "cuando
    corresponda": reportes sin una columna de importe corrida con
    sentido de negocio, ej. listado de Clientes, no los pasan):
    `valores_transporte` es paralelo a `filas` (un `Decimal` por fila).
    Si se pasan ambos, cada hoja (salvo la primera) arranca con
    "Viene: $X" (acumulado de hojas previas) y cada hoja (salvo la
    última) termina con "Transporte: $Y" (acumulado hasta ahí) — el
    total de la última hoja coincide con el total general.

    Al final de la última hoja se agrega también, sola, la cantidad de
    renglones impresos.
    """
    directorio = Path(directorio_salida) if directorio_salida is not None else DIR_SALIDA_DEFAULT
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / nombre_archivo

    tamano_pagina = landscape(A4) if apaisado else A4
    ancho_pagina, alto_pagina = tamano_pagina
    margen = 12 * mm

    # Alto reservado arriba/abajo para el encabezado/pie que dibuja
    # `_NumberedCanvas` en cada hoja — pedido del usuario (2026-08-20,
    # "agregar más renglones... que empiece 2mm después de los títulos y
    # hasta 2mm antes del nro de hoja"): antes había un colchón mucho más
    # generoso ahí (60pt/24pt) sin pedido explícito, que le sacaba
    # renglones de más a cada hoja — ahora es sólo lo que ocupa el texto
    # (empresa/fecha + título + subtítulo arriba; "Hoja N de M" abajo)
    # más 2mm de aire antes de que arranque/termine la tabla de datos.
    GAP_ENCABEZADO_PIE = 2 * mm
    alto_texto_encabezado = 40  # empresa/fecha + título + subtítulo (ver `_dibujar_marco`)
    alto_texto_pie = 10  # una línea de 8pt ("Hoja N de M")
    alto_encabezado = alto_texto_encabezado + GAP_ENCABEZADO_PIE
    alto_pie = alto_texto_pie + GAP_ENCABEZADO_PIE
    top_margin = margen + alto_encabezado
    bottom_margin = margen + alto_pie

    num_columnas = len(columnas)
    ancho_util = ancho_pagina - 2 * margen

    def _estilo_base() -> list[tuple]:
        # Ahorro de tonner (pedido del usuario, 2026-08-20): sin relleno
        # de color en los títulos (antes una barra sólida oscura en toda
        # la fila de encabezado) — en cambio, títulos en negrita más
        # grande que el cuerpo, centrados en cada columna, con una sola
        # línea debajo para separarlos visualmente. Padding de celda más
        # angosto (2pt en vez del default de Platypus, 3pt) para que
        # entren más renglones por hoja ("lo máximo posible", ídem).
        estilo = [
            ("FONTSIZE", (0, 1), (-1, -1), 7.5),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#bbbbbb")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        for col in columnas_derecha:
            estilo.append(("ALIGN", (col, 0), (col, -1), "RIGHT"))
        # Encabezado: va DESPUÉS del `ALIGN` de `columnas_derecha` de
        # arriba a propósito, para pisar el "RIGHT" de esa fila con
        # "CENTER" (los comandos de `TableStyle` se aplican en orden, el
        # último gana sobre el mismo rango).
        estilo.append(("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"))
        estilo.append(("FONTSIZE", (0, 0), (-1, 0), 9))
        estilo.append(("ALIGN", (0, 0), (-1, 0), "CENTER"))
        estilo.append(("LINEBELOW", (0, 0), (-1, 0), 1, colors.black))
        return estilo

    # --- Cuántas filas de datos entran por página -----------------------
    # Medido de verdad con `Table.wrap()` (no un alto de fila "a ojo") —
    # bug real encontrado probando con 79 filas (2026-08-20): un valor
    # aproximado dejaba cada bloque un pelo más grande de lo que entraba
    # de verdad, y Platypus terminaba partiendo CADA bloque en 2 páginas
    # reales (una con casi todo el bloque, otra con 1 sola fila + la fila
    # de Transporte) — funcionaba (no se perdía nada) pero desperdiciaba
    # una hoja extra en cada corte. Se arman 2 tablas de prueba con el
    # mismo estilo real (sólo encabezado, y encabezado + 1 fila) y se
    # resta para saber el alto real de una fila de datos.
    _tabla_solo_header = Table([columnas], repeatRows=1)
    _tabla_solo_header.setStyle(TableStyle(_estilo_base()))
    _, _alto_header = _tabla_solo_header.wrap(ancho_util, alto_pagina)

    _tabla_una_fila = Table([columnas, [""] * num_columnas], repeatRows=1)
    _tabla_una_fila.setStyle(TableStyle(_estilo_base()))
    _, _alto_header_mas_fila = _tabla_una_fila.wrap(ancho_util, alto_pagina)

    alto_fila = max(_alto_header_mas_fila - _alto_header, 1)
    alto_util_datos = alto_pagina - top_margin - bottom_margin
    # Colchón generoso (Viene + Transporte + hasta 3 líneas de `pie` + el
    # renglón de conteo) — mejor una página con algo de aire de más que
    # arriesgar el mismo desborde; si igual se queda corta, Platypus
    # parte la tabla puntual sola (red de seguridad, no se pierde nada).
    FILAS_RESERVADAS = 6
    filas_por_pagina = max(1, int((alto_util_datos - _alto_header) / alto_fila) - FILAS_RESERVADAS)

    def _fila_total(etiqueta: str, valor: str, columna: int) -> list[str]:
        # `valor` viaja ya formateado tal cual se quiere mostrar (con "$"
        # si es un importe, como ya hacían los call sites existentes) —
        # no se le antepone nada acá: algunos totales no son un importe
        # puro (ej. "ND Rechazadas" de Comisiones combina cantidad +
        # importe en un solo texto).
        fila = [""] * num_columnas
        if columna <= 0:
            fila[0] = f"{etiqueta} {valor}"
        else:
            fila[0] = etiqueta
            fila[columna] = valor
        return fila

    def _estilizar_fila_total(estilo: list[tuple], fila_idx: int, columna: int) -> None:
        estilo.append(("FONTNAME", (0, fila_idx), (-1, fila_idx), "Helvetica-Bold"))
        if columna > 0:
            estilo.append(("SPAN", (0, fila_idx), (columna - 1, fila_idx)))
            estilo.append(("BOX", (columna, fila_idx), (columna, fila_idx), 0.8, colors.black))
            estilo.append(("ALIGN", (columna, fila_idx), (columna, fila_idx), "RIGHT"))
        else:
            estilo.append(("SPAN", (0, fila_idx), (-1, fila_idx)))
            estilo.append(("BOX", (0, fila_idx), (0, fila_idx), 0.8, colors.black))

    # --- Armado de las páginas ------------------------------------------
    total_filas = len(filas)
    num_paginas = -(-total_filas // filas_por_pagina) if total_filas else 1
    elementos: list = []
    acumulado_transporte = Decimal("0")

    for num_pagina in range(num_paginas):
        es_primera = num_pagina == 0
        es_ultima = num_pagina == num_paginas - 1
        desde = num_pagina * filas_por_pagina
        hasta = min(desde + filas_por_pagina, total_filas)
        bloque = filas[desde:hasta]

        datos_tabla = [columnas]
        estilo = _estilo_base()
        filas_zebra: list[int] = []  # índices dentro de ESTA tabla, sólo filas de datos reales

        if not es_primera and columna_transporte is not None:
            fila_idx = len(datos_tabla)
            datos_tabla.append(_fila_total("Viene:", f"$ {format_decimal(acumulado_transporte)}", columna_transporte))
            _estilizar_fila_total(estilo, fila_idx, columna_transporte)

        for fila in bloque:
            filas_zebra.append(len(datos_tabla))
            datos_tabla.append(fila)

        if valores_transporte is not None and columna_transporte is not None:
            acumulado_transporte += sum(valores_transporte[desde:hasta], Decimal("0"))

        if not es_ultima and columna_transporte is not None:
            fila_idx = len(datos_tabla)
            datos_tabla.append(_fila_total("Transporte:", f"$ {format_decimal(acumulado_transporte)}", columna_transporte))
            _estilizar_fila_total(estilo, fila_idx, columna_transporte)

        if es_ultima:
            for etiqueta, valor, columna in pie or []:
                fila_idx = len(datos_tabla)
                datos_tabla.append(_fila_total(f"{etiqueta}:", valor, columna))
                _estilizar_fila_total(estilo, fila_idx, columna)

            fila_idx = len(datos_tabla)
            # Bug real encontrado probando (2026-08-20): "renglón"/
            # "impreso" no pluralizan con el mismo sufijo ("renglones",
            # no "renglónes"; "impresos", no "impresoes") — un solo
            # sufijo compartido quedaba mal en los dos.
            texto_conteo = "1 renglón impreso" if total_filas == 1 else f"{total_filas} renglones impresos"
            fila_conteo = [""] * num_columnas
            fila_conteo[0] = texto_conteo
            datos_tabla.append(fila_conteo)
            estilo.append(("FONTNAME", (0, fila_idx), (-1, fila_idx), "Helvetica-Oblique"))
            estilo.append(("SPAN", (0, fila_idx), (-1, fila_idx)))

        # Zebra striping sólo sobre filas de datos reales — no sobre el
        # encabezado ni las de total (que ya tienen su propio estilo).
        for i, idx in enumerate(filas_zebra):
            if i % 2 == 1:
                estilo.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#f2f2f2")))

        tabla = Table(datos_tabla, repeatRows=1)
        tabla.setStyle(TableStyle(estilo))
        elementos.append(tabla)
        if not es_ultima:
            elementos.append(PageBreak())

    class _NumberedCanvas(canvas.Canvas):
        """Difiere el dibujo del pie de página (necesita el total de
        páginas, que sólo se sabe al terminar) hasta `save()` — patrón
        estándar de ReportLab para numerar "Hoja N de M"."""

        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self._paginas_guardadas: list[dict] = []

        def showPage(self) -> None:  # noqa: N802 (override de reportlab)
            self._paginas_guardadas.append(dict(self.__dict__))
            self._startPage()

        def save(self) -> None:  # noqa: N802 (override de reportlab)
            total_paginas = len(self._paginas_guardadas)
            for estado in self._paginas_guardadas:
                self.__dict__.update(estado)
                self._dibujar_marco(total_paginas)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def _dibujar_marco(self, total_paginas: int) -> None:
            self.saveState()
            y_tope = alto_pagina - margen
            self.setFont("Helvetica-Bold", 11)
            self.drawString(margen, y_tope - 8, "ALESTEL SRL")
            self.setFont("Helvetica", 9)
            self.drawRightString(ancho_pagina - margen, y_tope - 8, formatear_fecha_corta(date.today()))
            self.setFont("Helvetica-Bold", 13)
            self.drawCentredString(ancho_pagina / 2, y_tope - 26, titulo)
            self.setFont("Helvetica", 8.5)
            self.drawCentredString(ancho_pagina / 2, y_tope - 38, subtitulo)
            self.setFont("Helvetica", 8)
            # Alineado con `alto_texto_pie` de más arriba — la tabla
            # termina 2mm por encima de este texto (ver `alto_pie`).
            self.drawCentredString(ancho_pagina / 2, margen + 4, f"Hoja {self.getPageNumber()} de {total_paginas}")
            self.restoreState()

    doc = SimpleDocTemplate(
        str(ruta), pagesize=tamano_pagina,
        leftMargin=margen, rightMargin=margen, topMargin=top_margin, bottomMargin=bottom_margin,
        title=titulo,
    )
    doc.build(elementos, canvasmaker=_NumberedCanvas)
    return ruta
    return ruta
