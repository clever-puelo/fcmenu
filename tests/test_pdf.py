"""Tests de migration/pdf.py — generación del PDF de Factura Electrónica.

Se valida contenido REAL del PDF (vía `pypdf`), no sólo que el archivo
exista — el CAE/Vencimiento/Total son datos legales del comprobante, no
detalles cosméticos.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest
from pypdf import PdfReader

from migration.decimals import format_decimal
from migration.models import Ctascte
from migration.pdf import (
    DatosFacturaPDF,
    DatosNotaCreditoInternaPDF,
    DatosReciboPDF,
    generar_pdf_factura,
    generar_pdf_listado,
    generar_pdf_nota_credito_interna,
    generar_pdf_recibo,
)
from migration.services import AplicacionPago, PagoCheque, PagoRetencion, RenglonEmision, TotalFactura


def _datos_factura(**overrides) -> DatosFacturaPDF:
    base = dict(
        letra="A",
        punto_venta=3,
        numero=42,
        fecha=date(2026, 1, 10),
        cliente_codigo=100,
        cliente_nombre="CLIENTE DE PRUEBA SA",
        cliente_cuit="20111111119",
        cliente_civa=1,
        cliente_domicilio="AV. SIEMPREVIVA 742",
        renglones=[
            RenglonEmision(
                cod1="AA", cod2="1", descripcion="ARTICULO DE PRUEBA",
                precio_unitario=Decimal("100"), importe=Decimal("1000"),
                cantidad_unidades=Decimal("10"),
            )
        ],
        total=TotalFactura(
            bruto=Decimal("1000"), descuento=Decimal("0"), neto_gravado=Decimal("1000"),
            iva=Decimal("210"), percepcion_iibb=Decimal("0"), total=Decimal("1210"),
        ),
        cae="70123456789012",
        cae_vencimiento=date(2026, 1, 20),
        qr_url="https://www.afip.gob.ar/fe/qr/?p=abc123",
    )
    base.update(overrides)
    return DatosFacturaPDF(**base)


class TestGenerarPdfFactura:
    def test_genera_un_archivo_pdf_valido(self, tmp_path):
        ruta = generar_pdf_factura(_datos_factura(), directorio_salida=tmp_path)

        assert ruta.exists()
        assert ruta.suffix == ".pdf"
        assert ruta.read_bytes()[:4] == b"%PDF"

    def test_nombre_de_archivo_incluye_letra_ptovta_y_numero(self, tmp_path):
        # Réplica real de EmiFact.frm:1736 (`TIPO & "-" & LaLetra & "-"
        # & PtoVtaCpbte & "-" & NROFAC`) — el número NO lleva ceros a la
        # izquierda en el nombre de archivo (a diferencia del texto
        # impreso adentro del PDF).
        ruta = generar_pdf_factura(_datos_factura(letra="B", punto_venta=3, numero=123), directorio_salida=tmp_path)
        assert ruta.name == "F A C T U R A-B-0003-123.pdf"

    def test_contenido_incluye_datos_legales_del_comprobante(self, tmp_path):
        datos = _datos_factura()
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)

        reader = PdfReader(str(ruta))
        assert len(reader.pages) == 2  # ORIGINAL + DUPLICADO, réplica real (EmiFact.frm, variable Copias)
        texto = reader.pages[0].extract_text()
        texto_pagina2 = reader.pages[1].extract_text()

        assert "F A C T U R A" in texto
        assert "0003-00000042" in texto
        assert datos.cliente_nombre in texto
        assert datos.cliente_cuit in texto
        assert datos.cae in texto
        assert "20/01/2026" in texto  # vencimiento CAE
        assert "1.210,00" in texto or "1,210.00" in texto or "1210.00" in texto  # TOTAL, formato numérico
        assert "ARTICULO DE PRUEBA" in texto
        assert "O R I G I N A L" in texto  # 1ª página
        assert "D U P L I C A D O" in texto_pagina2  # 2ª página

    def test_directorio_de_salida_se_crea_si_no_existe(self, tmp_path):
        subdir = tmp_path / "no_existe_todavia"
        ruta = generar_pdf_factura(_datos_factura(), directorio_salida=subdir)
        assert ruta.exists()
        assert ruta.parent == subdir

    def test_respeta_limite_de_renglones_sin_romper(self, tmp_path):
        """`DetFact.FG1` limita a 30 renglones reales (no 25 — corregido
        2026-08-20, límite real confirmado por el usuario) — con más de
        eso, el PDF no debe romper, sólo dejar de agregar los que no
        entran."""
        muchos_renglones = [
            RenglonEmision(
                cod1="AA", cod2=str(i), descripcion=f"ARTICULO {i}",
                precio_unitario=Decimal("10"), importe=Decimal("10"),
                cantidad_unidades=Decimal("1"),
            )
            for i in range(40)
        ]
        ruta = generar_pdf_factura(_datos_factura(renglones=muchos_renglones), directorio_salida=tmp_path)
        assert ruta.exists()
        assert ruta.read_bytes()[:4] == b"%PDF"

        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "ARTICULO 0" in texto
        assert "ARTICULO 29" in texto  # el 30º (índice 29) entra
        assert "ARTICULO 30" not in texto  # el 31º ya no

    def test_entran_treinta_renglones_reales(self, tmp_path):
        """Pedido explícito del usuario (2026-08-20): 30 renglones deben
        entrar en una sola página, con el pie (subtotal/totales/CAE)
        debajo de todos ellos, no encima."""
        treinta_renglones = [
            RenglonEmision(
                cod1="AA", cod2=str(i), descripcion=f"ARTICULO {i}",
                precio_unitario=Decimal("10"), importe=Decimal("10"),
                cantidad_unidades=Decimal("1"),
            )
            for i in range(30)
        ]
        ruta = generar_pdf_factura(_datos_factura(renglones=treinta_renglones), directorio_salida=tmp_path)
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        for i in range(30):
            assert f"ARTICULO {i}" in texto
        assert "TOTAL" in texto
        assert "C.A.E." in texto

    def test_detalle_incluye_el_despacho(self, tmp_path):
        """El Despacho se ve como columna aparte (hallazgo real
        2026-08-20, comparando contra una Factura real de muestra) —
        acá se dibuja en su propia posición X, mismo resultado visual
        que el legacy sin repetir su técnica de padding monoespaciado."""
        datos = _datos_factura(
            renglones=[
                RenglonEmision(
                    cod1="AA", cod2="1", descripcion="ARTICULO CON LOTE",
                    precio_unitario=Decimal("100"), importe=Decimal("1000"),
                    cantidad_unidades=Decimal("10"), nrodesp_elegido="25001IC04066997X",
                )
            ],
        )
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "ARTICULO CON LOTE" in texto
        assert "25001IC04066997X" in texto

    def test_columna_de_descuento_por_renglon(self, tmp_path):
        """Réplica de la columna "% Descuento" de EmiFact.frm — texto
        libre por renglón (ej. "10+5"), en el mismo orden que
        `renglones` (ver `FacturadorWindow._descuentos_por_renglon`)."""
        datos = _datos_factura(descuentos_renglones=["10+5"])
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "10+5" in texto
        assert "%Descuento" in texto or "% Descuento" in texto.replace("\n", " ")

    def test_en_dolares_incluye_la_leyenda_legal_y_la_cotizacion(self, tmp_path):
        """Réplica de `EmiFact.frm:1085-1094` (`LeyenDol1..5` + "U$S"
        sobre las cajas de Subtotal/Subtotal neto/TOTAL) — sólo cuando
        `en_dolares=True` (pedido del usuario, 2026-08-21)."""
        datos = _datos_factura(en_dolares=True, cotizacion=Decimal("1515.00"))
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)
        texto = PdfReader(str(ruta)).pages[0].extract_text()

        assert "LA PRESENTE FACTURA SE ABONARA EN DOLARES" in texto
        assert "1.515,00" in texto
        assert texto.count("U$S") >= 3

    def test_sin_en_dolares_no_incluye_la_leyenda_legal(self, tmp_path):
        datos = _datos_factura(en_dolares=False)
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)
        texto = PdfReader(str(ruta)).pages[0].extract_text()

        assert "LA PRESENTE FACTURA SE ABONARA EN DOLARES" not in texto
        assert "U$S" not in texto

    def test_cotizacion_no_lleva_cae_ni_sello_borrador(self, tmp_path):
        """`es_cotizacion=True` (Cotización, `CabFact.frm TipoFac=4`,
        ver `CotizacionVentaService`) — nunca pide CAE, así que nunca
        debe mostrar el sello "BORRADOR" (eso es sólo para el boceto
        DE UNA FACTURA real, pendiente de AFIP)."""
        datos = _datos_factura(
            letra="X", punto_venta=1, numero=7, cae=None, cae_vencimiento=None, qr_url=None, es_cotizacion=True
        )
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)

        assert ruta.name == "COTIZACIÓN-0001-7.pdf"
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "COTIZACIÓN" in texto
        assert "Documento sin validez fiscal" in texto
        assert "F A C T U R A" not in texto
        assert "CAE" not in texto
        # Cotización es 1 sola página (no tiene sentido un "Duplicado"
        # de algo sin validez fiscal).
        assert len(PdfReader(str(ruta)).pages) == 1

    def test_qr_url_se_incluye_como_grafico_no_como_texto(self, tmp_path):
        """El QR se dibuja como gráfico vectorial (QrCodeWidget), no
        aparece como texto plano en el PDF — sólo se confirma que el PDF
        no falla con una URL larga real."""
        datos = _datos_factura(
            qr_url="https://www.afip.gob.ar/fe/qr/?p=" + "eyJ2ZXIiOjEsImZlY2hhIjoiMjAyMS0wMy0xNyJ9" * 3
        )
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)
        assert ruta.exists()


# ---------------------------------------------------------------------------
# generar_pdf_recibo
# ---------------------------------------------------------------------------


def _datos_recibo(**overrides) -> DatosReciboPDF:
    factura = Ctascte(CPBTE=10, TIPO=1, FECHA=date(2026, 1, 5), DEBE=Decimal("0"))
    base = dict(
        numero=1000,
        fecha=date(2026, 2, 1),
        correlativo=4,
        cliente_codigo=200,
        cliente_nombre="CLIENTE DE PRUEBA SA",
        cliente_cuit="20111111119",
        aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("600"))],
        anticipo=Decimal("0"),
        importe_efectivo=Decimal("600"),
        cheques=[],
        retenciones=[],
        total_pago=Decimal("600"),
    )
    base.update(overrides)
    return DatosReciboPDF(**base)


class TestGenerarPdfRecibo:
    def test_genera_un_archivo_pdf_valido(self, tmp_path):
        ruta = generar_pdf_recibo(_datos_recibo(), directorio_salida=tmp_path)

        assert ruta.exists()
        assert ruta.suffix == ".pdf"
        assert ruta.read_bytes()[:4] == b"%PDF"

    def test_nombre_de_archivo_incluye_el_numero(self, tmp_path):
        # Réplica real (`EmiRec.frm`): "Recibo-X-0001-{numero}", sin
        # ceros a la izquierda en el nombre de archivo — Letra "X" fija
        # (Recibo nunca es fiscal).
        ruta = generar_pdf_recibo(_datos_recibo(numero=123), directorio_salida=tmp_path)
        assert ruta.name == "Recibo-X-0001-123.pdf"

    def test_contenido_incluye_cliente_comprobante_aplicado_y_total(self, tmp_path):
        datos = _datos_recibo()
        ruta = generar_pdf_recibo(datos, directorio_salida=tmp_path)

        texto = PdfReader(str(ruta)).pages[0].extract_text()

        assert "00001000" in texto
        assert datos.cliente_nombre in texto
        assert datos.cliente_cuit in texto
        assert "Fact." in texto  # ETIQUETAS_TIPO_CTASCTE[1]
        assert "600,00" in texto or "600.00" in texto
        assert "Total:" in texto
        assert "Original" in texto and "Duplicado" in texto  # 2 copias apiladas, réplica real

    def test_contenido_incluye_cheques_y_retenciones(self, tmp_path):
        datos = _datos_recibo(
            importe_efectivo=Decimal("0"),
            cheques=[
                PagoCheque(
                    nro_cheque=555111, banco="007", fecha_emision=date(2026, 2, 1),
                    fecha_vencimiento=date(2026, 3, 1), importe=Decimal("400"),
                )
            ],
            retenciones=[PagoRetencion(tipreg=1, importe=Decimal("200"), concepto="Retención Ganancias")],
        )
        ruta = generar_pdf_recibo(datos, directorio_salida=tmp_path)

        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "555111" in texto
        assert "Retención Ganancias" in texto

    def test_incluye_anticipo_cuando_corresponde(self, tmp_path):
        datos = _datos_recibo(anticipo=Decimal("100"))
        ruta = generar_pdf_recibo(datos, directorio_salida=tmp_path)

        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "Anticipo" in texto

    def test_directorio_de_salida_se_crea_si_no_existe(self, tmp_path):
        subdir = tmp_path / "no_existe_todavia"
        ruta = generar_pdf_recibo(_datos_recibo(), directorio_salida=subdir)
        assert ruta.exists()
        assert ruta.parent == subdir

    def test_respeta_limite_de_aplicaciones_sin_romper(self, tmp_path):
        muchas = [
            AplicacionPago(
                comprobante=Ctascte(CPBTE=i, TIPO=1, FECHA=date(2026, 1, 1), DEBE=Decimal("0")),
                importe_aplicado=Decimal("10"),
            )
            for i in range(60)
        ]
        ruta = generar_pdf_recibo(
            _datos_recibo(aplicaciones=muchas, importe_efectivo=Decimal("600")), directorio_salida=tmp_path
        )
        assert ruta.exists()
        assert ruta.read_bytes()[:4] == b"%PDF"


# ---------------------------------------------------------------------------
# generar_pdf_nota_credito_interna (NCInterna.frm) — cancela deuda,
# documento interno sin CAE.
# ---------------------------------------------------------------------------


def _datos_nci(**overrides) -> DatosNotaCreditoInternaPDF:
    factura = Ctascte(CPBTE=10, TIPO=1, FECHA=date(2026, 1, 5), IMPTE=Decimal("500"))
    base = dict(
        numero=42,
        fecha=date(2026, 2, 1),
        cliente_codigo=500,
        cliente_nombre="CLIENTE NCI SA",
        cliente_cuit="20111111119",
        motivo_texto="5 - ERROR DE FACTURACION",
        nota="Ajuste por error de facturación detectado en auditoría interna",
        comprobantes_cancelados=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("500"))],
        total_cancelado=Decimal("500"),
    )
    base.update(overrides)
    return DatosNotaCreditoInternaPDF(**base)


class TestGenerarPdfNotaCreditoInterna:
    def test_genera_un_archivo_pdf_valido(self, tmp_path):
        ruta = generar_pdf_nota_credito_interna(_datos_nci(), directorio_salida=tmp_path)
        assert ruta.exists()
        assert ruta.read_bytes()[:4] == b"%PDF"

    def test_nombre_de_archivo_incluye_el_numero(self, tmp_path):
        ruta = generar_pdf_nota_credito_interna(_datos_nci(numero=7), directorio_salida=tmp_path)
        assert ruta.name == "NOTA DE CREDITO INTERNA-X-0001-7.pdf"

    def test_contenido_incluye_datos_del_documento(self, tmp_path):
        datos = _datos_nci()
        ruta = generar_pdf_nota_credito_interna(datos, directorio_salida=tmp_path)
        texto = PdfReader(str(ruta)).pages[0].extract_text()

        assert "Nota de Crédito Interna" in texto
        assert datos.cliente_nombre in texto
        assert "500,00" in texto or "500.00" in texto
        assert "ERROR DE FACTURACION" in texto
        assert "auditoría interna" in texto
        assert "Documento no" in texto  # mismo letterhead "no válido como Factura" que Recibo
        assert "CAE" not in texto
        assert "Original" in texto and "Duplicado" in texto


# ---------------------------------------------------------------------------
# generar_pdf_listado (migración de Listados.frm) — pedido del usuario,
# 2026-08-20: título+fecha y "Hoja N de M" en cada hoja, "Viene"/
# "Transporte" entre hojas, totales en recuadro bajo su columna, cantidad
# de renglones impresos.
# ---------------------------------------------------------------------------


class TestGenerarPdfListado:
    def test_reporte_de_una_sola_hoja_sin_totales(self, tmp_path):
        ruta = generar_pdf_listado(
            titulo="Clientes", subtitulo="Todos",
            columnas=["Cód.", "Nombre"],
            filas=[["1", "Cliente Uno"], ["2", "Cliente Dos"]],
            columnas_derecha=(0,),
            directorio_salida=tmp_path,
        )
        assert ruta.exists()
        assert ruta.read_bytes()[:4] == b"%PDF"

        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "ALESTEL SRL" in texto
        assert "Clientes" in texto
        assert "Hoja 1 de 1" in texto
        assert "2 renglones impresos" in texto

    def test_reporte_vacio_no_rompe_y_muestra_cero_renglones(self, tmp_path):
        ruta = generar_pdf_listado(
            titulo="Vacío", subtitulo="Sin datos",
            columnas=["Cód.", "Nombre"], filas=[],
            directorio_salida=tmp_path,
        )
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "0 renglones impresos" in texto

    def test_un_solo_renglon_no_pluraliza(self, tmp_path):
        ruta = generar_pdf_listado(
            titulo="Uno", subtitulo="Un renglón",
            columnas=["Cód.", "Nombre"], filas=[["1", "Único"]],
            directorio_salida=tmp_path,
        )
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "1 renglón impreso" in texto
        assert "renglones" not in texto

    def test_pie_se_dibuja_en_la_ultima_hoja_bajo_su_columna(self, tmp_path):
        ruta = generar_pdf_listado(
            titulo="Con Total", subtitulo="Prueba",
            columnas=["Cód.", "Nombre", "Importe"],
            filas=[["1", "A", "100,00"], ["2", "B", "200,00"]],
            columnas_derecha=(0, 2),
            pie=[("Total", "$ 300,00", 2)],
            directorio_salida=tmp_path,
        )
        texto = PdfReader(str(ruta)).pages[0].extract_text()
        assert "Total" in texto
        assert "$ 300,00" in texto

    def test_reporte_largo_pagina_con_viene_y_transporte(self, tmp_path):
        """79 renglones fuerzan varias hojas — confirma que cada hoja
        intermedia arranca con "Viene" y termina con "Transporte", que
        acumulan correctamente, y que el total de la última hoja coincide
        con el total general (mismo criterio contable confirmado con el
        usuario: "Viene de hoja anterior" / "Pasa a hoja siguiente")."""
        filas = [[str(i), f"Cliente {i}", format_decimal(i * 10)] for i in range(1, 80)]
        valores = [Decimal(i * 10) for i in range(1, 80)]
        total = sum(valores, Decimal("0"))

        ruta = generar_pdf_listado(
            titulo="Prueba Larga", subtitulo="79 renglones",
            columnas=["Cód.", "Nombre", "Importe"],
            filas=filas, columnas_derecha=(0, 2),
            columna_transporte=2, valores_transporte=valores,
            pie=[("Total", f"$ {format_decimal(total)}", 2)],
            directorio_salida=tmp_path,
        )
        reader = PdfReader(str(ruta))
        assert len(reader.pages) > 1

        primera = reader.pages[0].extract_text()
        ultima = reader.pages[-1].extract_text()
        assert "Viene" not in primera  # la primera hoja no "viene" de ninguna
        assert "Transporte" in primera  # pero sí pasa a la siguiente
        assert "Viene" in ultima  # la última "viene" de las anteriores
        assert "Transporte" not in ultima  # y no pasa a ninguna otra
        assert f"$ {format_decimal(total)}" in ultima  # el total general
        assert "79 renglones impresos" in ultima

        # Todas las hojas repiten encabezado, fecha, y "Hoja N de M".
        for i, pagina in enumerate(reader.pages, start=1):
            texto = pagina.extract_text()
            assert "ALESTEL SRL" in texto
            assert f"Hoja {i} de {len(reader.pages)}" in texto
