"""Tests de migration/afip.py.

`generar_qr_afip` se valida contra el ejemplo real de AFIP que el legacy
tenía hardcodeado en `EmiFact.frm:2785` (Obtener_QR, ver docstring del
bug del QR estático) — decodificando el QR generado y comparando los
valores, no sólo que "no rompa".
"""

from __future__ import annotations

import base64
import json
from datetime import date
from decimal import Decimal

import pytest

from migration.afip import AfipWSFEv1Stub, codigo_afip, generar_qr_afip, punto_venta_por_tipo


class TestCodigosAfip:
    def test_factura_a_es_codigo_1(self):
        assert codigo_afip("A", 1) == 1

    def test_factura_b_es_codigo_6(self):
        assert codigo_afip("B", 1) == 6

    def test_nc_a_es_codigo_3_y_nd_a_es_codigo_2(self):
        # EmiFact.frm:723-725 — no es el orden intuitivo (NC=3, ND=2).
        assert codigo_afip("A", 2) == 3  # NC A
        assert codigo_afip("A", 3) == 2  # ND A

    def test_combinacion_no_soportada_rechaza(self):
        with pytest.raises(ValueError):
            codigo_afip("C", 1)


class TestPuntoVentaPorTipo:
    def test_factura_usa_punto_venta_3(self):
        """Hallazgo que corrige la nota previa de memoria: el Punto de
        Venta depende del TIPO de comprobante (Factura vs NC/ND), no de
        la Letra — EmiFact.frm:2617 agrupa CodFactAFIP 1 y 6 (Factura A
        y B) bajo el mismo Punto de Venta 3."""
        assert punto_venta_por_tipo(1) == 3

    def test_nc_nd_usa_punto_venta_4(self):
        assert punto_venta_por_tipo(2) == 4
        assert punto_venta_por_tipo(3) == 4


class TestAfipWSFEv1Stub:
    def test_ultimo_comprobante_arranca_en_cero(self):
        stub = AfipWSFEv1Stub()
        assert stub.ultimo_comprobante(3, 1) == 0

    def test_solicitar_cae_aprueba_y_actualiza_ultimo_comprobante(self):
        stub = AfipWSFEv1Stub()
        resultado = stub.solicitar_cae(
            tipo_cbte=1, punto_venta=3, cbte_nro=1,
            cuit_receptor="20111111119", importe_neto=Decimal("100"),
            importe_iva=Decimal("21"), importe_iibb=Decimal("0"),
            importe_total=Decimal("121"), fecha_cbte=date(2026, 1, 10),
        )
        assert resultado.aprobado is True
        assert resultado.vencimiento == date(2026, 1, 20)  # +10 días corridos
        assert stub.ultimo_comprobante(3, 1) == 1

    def test_solicitar_cae_devuelve_un_cae_numerico_compatible_con_el_qr(self):
        """Bug real encontrado probando el circuito completo de
        `FacturadorWindow`: un CAE con letras (ej. prefijo "STUB") rompe
        `generar_qr_afip()`, que asume un CAE numérico como el real de
        AFIP (`int(cae)` para el campo `codAut`). El stub debe ser
        intercambiable en forma con un CAE real, aunque sea obviamente
        falso en valor."""
        stub = AfipWSFEv1Stub()
        resultado = stub.solicitar_cae(
            tipo_cbte=1, punto_venta=3, cbte_nro=1,
            cuit_receptor="20111111119", importe_neto=Decimal("100"),
            importe_iva=Decimal("21"), importe_iibb=Decimal("0"),
            importe_total=Decimal("121"), fecha_cbte=date(2026, 1, 10),
        )
        assert resultado.cae.isdigit()

        url = generar_qr_afip(
            cuit_emisor="33703467909", punto_venta=3, tipo_cbte=1, nro_cbte=1,
            importe_total=Decimal("121"), tipo_doc_receptor=80,
            nro_doc_receptor="20111111119", cae=resultado.cae,
        )
        assert url.startswith("https://www.afip.gob.ar/fe/qr/?p=")

    def test_solicitar_cae_sin_cuit_rechaza(self):
        stub = AfipWSFEv1Stub()
        with pytest.raises(ValueError):
            stub.solicitar_cae(
                tipo_cbte=1, punto_venta=3, cbte_nro=1, cuit_receptor="",
                importe_neto=Decimal("100"), importe_iva=Decimal("21"),
                importe_iibb=Decimal("0"), importe_total=Decimal("121"),
                fecha_cbte=date(2026, 1, 10),
            )

    def test_dos_puntos_de_venta_no_se_pisan(self):
        stub = AfipWSFEv1Stub()
        stub.solicitar_cae(
            tipo_cbte=1, punto_venta=3, cbte_nro=5, cuit_receptor="20111111119",
            importe_neto=Decimal("1"), importe_iva=Decimal("0"), importe_iibb=Decimal("0"),
            importe_total=Decimal("1"), fecha_cbte=date(2026, 1, 1),
        )
        assert stub.ultimo_comprobante(3, 1) == 5
        assert stub.ultimo_comprobante(4, 2) == 0  # NC en otro punto de venta, sin tocar


class TestGenerarQrAfip:
    def _decodificar(self, url: str) -> dict:
        assert url.startswith("https://www.afip.gob.ar/fe/qr/?p=")
        b64 = url.split("?p=", 1)[1]
        return json.loads(base64.b64decode(b64).decode("utf-8"))

    def test_qr_replica_el_formato_real_de_afip(self):
        """Contra el ejemplo real embebido en el legacy (EmiFact.frm:2785,
        QR real de AFIP, CUIT 33708755309, decodificado en la
        investigación de este módulo)."""
        url = generar_qr_afip(
            cuit_emisor="33708755309",
            punto_venta=5,
            tipo_cbte=1,
            nro_cbte=1511,
            importe_total=Decimal("38478.00"),
            tipo_doc_receptor=80,
            nro_doc_receptor="305227806",
            cae="7111931826965 6".replace(" ", ""),
            fecha_cbte=date(2021, 3, 17),
        )
        payload = self._decodificar(url)
        assert payload["ver"] == 1
        assert payload["fecha"] == "2021-03-17"
        assert payload["cuit"] == 33708755309
        assert payload["ptoVta"] == 5
        assert payload["tipoCmp"] == 1
        assert payload["nroCmp"] == 1511
        assert payload["importe"] == 38478.00
        assert payload["moneda"] == "PES"
        assert payload["ctz"] == 1.0
        assert payload["tipoDocRec"] == 80
        assert payload["nroDocRec"] == 305227806
        assert payload["tipoCodAut"] == "E"
        assert payload["codAut"] == 71119318269656

    def test_qr_cada_factura_lleva_su_propio_cae(self):
        """Réplica del bug real corregido: acá SÍ cambia con cada
        factura (el legacy imprimía siempre el mismo QR hardcodeado)."""
        url_1 = generar_qr_afip(
            cuit_emisor="33703467909", punto_venta=3, tipo_cbte=1, nro_cbte=100,
            importe_total=Decimal("1000"), tipo_doc_receptor=80,
            nro_doc_receptor="20111111119", cae="70000000000001",
        )
        url_2 = generar_qr_afip(
            cuit_emisor="33703467909", punto_venta=3, tipo_cbte=1, nro_cbte=101,
            importe_total=Decimal("2000"), tipo_doc_receptor=80,
            nro_doc_receptor="20111111119", cae="70000000000002",
        )
        assert url_1 != url_2
        assert self._decodificar(url_1)["codAut"] == 70000000000001
        assert self._decodificar(url_2)["codAut"] == 70000000000002


class TestCondicionIvaReceptor:
    """`CondicionIVAReceptorId` (RG 5616/2024) — mapeo confirmado
    probando contra Homologación real (2026-08-19): AFIP rechazó un CAE
    de prueba sin este campo (`[10246] Campo Condicion Frente al IVA
    del receptor es obligatorio...`)."""

    def test_mapeo_por_civa(self):
        from migration.afip import condicion_iva_receptor

        assert condicion_iva_receptor(1) == 1  # Inscripto -> Responsable Inscripto
        assert condicion_iva_receptor(2) == 1  # Resp.No Insc. (categoría muerta) -> mismo criterio que CIVA=1
        assert condicion_iva_receptor(3) == 5  # Cons.Final -> Consumidor Final
        assert condicion_iva_receptor(4) == 4  # Exento -> Sujeto Exento
        assert condicion_iva_receptor(5) == 6  # Monotributo -> Responsable Monotributo

    def test_sin_civa_no_inventa_default(self):
        from migration.afip import condicion_iva_receptor

        assert condicion_iva_receptor(None) is None


class TestCacheDeTokenWSAA:
    """Bug real encontrado probando contra Homologación por primera vez
    (2026-08-19): WSAA rechaza un `loginCms` nuevo mientras el TA
    anterior siga vigente (`Fault: "El CEE ya posee un TA valido..."`)
    — ver docstring de `AfipWSFEv1Cliente._archivo_cache_ta()`."""

    def _cliente(self, tmp_path, monkeypatch, archivo=None):
        from migration.afip import AfipWSFEv1Cliente

        cliente = AfipWSFEv1Cliente(
            cuit_emisor="20111111112", certificado_path="x.crt", clave_privada_path="x.key"
        )
        monkeypatch.setattr(cliente, "_archivo_cache_ta", lambda: archivo or (tmp_path / "ta.json"))
        return cliente

    def test_token_cacheado_se_reusa_si_todavia_vale(self, tmp_path, monkeypatch):
        from datetime import datetime, timedelta, timezone

        archivo = tmp_path / "ta.json"
        original = self._cliente(tmp_path, monkeypatch, archivo)
        original._token, original._sign = "TOKEN-REAL", "SIGN-REAL"
        original._token_vence = datetime.now(timezone(timedelta(hours=-3))) + timedelta(hours=10)
        original._guardar_token_cacheado()

        nuevo_proceso = self._cliente(tmp_path, monkeypatch, archivo)
        nuevo_proceso._cargar_token_cacheado()

        assert nuevo_proceso._token == "TOKEN-REAL"
        assert nuevo_proceso._sign == "SIGN-REAL"
        assert nuevo_proceso._token_vigente() is True

    def test_token_vencido_no_se_reusa(self, tmp_path, monkeypatch):
        from datetime import datetime, timedelta, timezone

        archivo = tmp_path / "ta.json"
        original = self._cliente(tmp_path, monkeypatch, archivo)
        original._token, original._sign = "VIEJO", "VIEJO"
        original._token_vence = datetime.now(timezone(timedelta(hours=-3))) - timedelta(hours=1)
        original._guardar_token_cacheado()

        nuevo_proceso = self._cliente(tmp_path, monkeypatch, archivo)
        nuevo_proceso._cargar_token_cacheado()

        assert nuevo_proceso._token_vigente() is False

    def test_sin_cache_no_rompe(self, tmp_path, monkeypatch):
        cliente = self._cliente(tmp_path, monkeypatch, tmp_path / "no_existe.json")
        cliente._cargar_token_cacheado()  # no debe tirar excepción
        assert cliente._token_vigente() is False
