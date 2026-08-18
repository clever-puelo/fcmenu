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

from migration.models import Ctascte
from migration.pdf import DatosFacturaPDF, DatosReciboPDF, generar_pdf_factura, generar_pdf_recibo
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
        ruta = generar_pdf_factura(_datos_factura(letra="B", punto_venta=3, numero=123), directorio_salida=tmp_path)
        assert ruta.name == "Factura_B_0003-00000123.pdf"

    def test_contenido_incluye_datos_legales_del_comprobante(self, tmp_path):
        datos = _datos_factura()
        ruta = generar_pdf_factura(datos, directorio_salida=tmp_path)

        texto = PdfReader(str(ruta)).pages[0].extract_text()

        assert "FACTURA A" in texto
        assert "0003-00000042" in texto
        assert datos.cliente_nombre in texto
        assert datos.cliente_cuit in texto
        assert datos.cae in texto
        assert "20/01/2026" in texto  # vencimiento CAE
        assert "1.210,00" in texto or "1,210.00" in texto or "1210.00" in texto  # TOTAL, formato numérico
        assert "ARTICULO DE PRUEBA" in texto

    def test_directorio_de_salida_se_crea_si_no_existe(self, tmp_path):
        subdir = tmp_path / "no_existe_todavia"
        ruta = generar_pdf_factura(_datos_factura(), directorio_salida=subdir)
        assert ruta.exists()
        assert ruta.parent == subdir

    def test_respeta_limite_de_renglones_sin_romper(self, tmp_path):
        """DetFact.frm limita a 25 renglones reales — con más de eso, el
        PDF no debe romper, sólo dejar de agregar renglones que no
        entren en la página."""
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
        ruta = generar_pdf_recibo(_datos_recibo(numero=123), directorio_salida=tmp_path)
        assert ruta.name == "Recibo_00000123.pdf"

    def test_contenido_incluye_cliente_comprobante_aplicado_y_total(self, tmp_path):
        datos = _datos_recibo()
        ruta = generar_pdf_recibo(datos, directorio_salida=tmp_path)

        texto = PdfReader(str(ruta)).pages[0].extract_text()

        assert "00001000" in texto
        assert datos.cliente_nombre in texto
        assert datos.cliente_cuit in texto
        assert "Fact. 10" in texto  # ETIQUETAS_TIPO_CTASCTE[1] + Cpbte
        assert "600,00" in texto or "600.00" in texto
        assert "TOTAL RECIBIDO" in texto

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
