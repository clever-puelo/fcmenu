"""Tests de migration/services.py.

Cada cálculo se contrasta contra las fórmulas confirmadas con el usuario
(ver data_layer_progress.md) y contra las líneas exactas del VB6 citadas
en los docstrings de services.py.
"""

from __future__ import annotations

from datetime import date, timedelta
from decimal import Decimal
from unittest.mock import patch

import pytest

from migration.models import (
    Articulo,
    Cheque,
    Cliente,
    Cotizacion,
    Ctascte,
    Despacho,
    DtoxClte,
    Efectivo,
    Fcestad1,
    Fctabla1,
    FcivaVta,
    Imputacion,
    MovimVS,
    MovStock,
    Parametro,
    Stock,
    Totales,
)
from migration.repository import RepositoryFactory
from migration.services import (
    AplicacionPago,
    ArticuloService,
    ChequeService,
    ClienteService,
    CuentaCorrienteService,
    EmisionFacturaService,
    EmisionReciboService,
    EstadisticaVentasService,
    FacturaService,
    FacturasEmitidasService,
    PagoCheque,
    PagoRetencion,
    RenglonEmision,
    RenglonMovimientoStock,
    ArregloCtaCteService,
    ArregloSubdiarioService,
    ListadosService,
    StockMovimientoService,
    StockService,
    TablaService,
    TotalesDiariosService,
)


def _set_parametro(db, ivains="21", ivani="0", mcaib=1):
    db.add(Parametro(CLAVE="1", IVAINS=Decimal(ivains), IVANI=Decimal(ivani), MCAIB=mcaib))
    db.commit()


# ---------------------------------------------------------------------------
# FacturaService
# ---------------------------------------------------------------------------


class TestFacturaService:
    def test_bonificacion_cascada_no_es_una_suma_simple(self, db):
        """EmiFact.frm:1821-1827: cada % se aplica sobre lo que queda del
        importe tras los descuentos previos, no son porcentajes que se
        suman entre sí (10% + 5% != 15%)."""
        _set_parametro(db)
        fs = FacturaService(db)

        resultado = fs.calcular_bonificacion_cascada(Decimal("1000"), [Decimal("10"), Decimal("5")])

        # 10% de 1000 = 100; 5% de (1000-100)=900 -> 45; total 145 (no 150).
        assert resultado == Decimal("145")

    def test_bonificacion_cascada_ignora_porcentajes_no_positivos(self, db):
        _set_parametro(db)
        fs = FacturaService(db)
        resultado = fs.calcular_bonificacion_cascada(Decimal("1000"), [Decimal("0"), Decimal("-5")])
        assert resultado == Decimal("0")

    def test_bonificacion_cascada_ignora_mas_de_5_niveles(self, db):
        _set_parametro(db)
        fs = FacturaService(db)
        seis_dieces = [Decimal("10")] * 6
        con_seis = fs.calcular_bonificacion_cascada(Decimal("1000"), seis_dieces)
        con_cinco = fs.calcular_bonificacion_cascada(Decimal("1000"), seis_dieces[:5])
        assert con_seis == con_cinco

    def test_alicuota_iva_inscripto_viene_de_parametro_no_hardcodeada(self, db):
        _set_parametro(db, ivains="10.5")
        fs = FacturaService(db)
        assert fs.alicuota_iva_inscripto() == Decimal("10.5") / Decimal(100)

    def test_alicuota_iva_inscripto_sin_config_lanza_error(self, db):
        fs = FacturaService(db)
        with pytest.raises(ValueError):
            fs.alicuota_iva_inscripto()

    def test_letra_a_para_inscripto_y_no_inscripto(self, db):
        """CabFact.frm:709: ClteCIVA < 3 -> Letra A."""
        fs = FacturaService(db)
        assert fs.letra_comprobante(civa_cliente=1, provincia="B ") == "A"
        assert fs.letra_comprobante(civa_cliente=2, provincia="B ") == "A"

    def test_letra_b_para_consumidor_final_y_monotributo(self, db):
        fs = FacturaService(db)
        assert fs.letra_comprobante(civa_cliente=3, provincia="B ") == "B"
        assert fs.letra_comprobante(civa_cliente=5, provincia="B ") == "B"

    def test_letra_a_para_exento_en_tierra_del_fuego(self, db):
        """Caso especial real: ClteCIVA=4 (Exento) sólo da Letra A si la
        Provincia es 'V' (Tierra del Fuego) — cualquier otra provincia
        con CIVA=4 da Letra B."""
        fs = FacturaService(db)
        assert fs.letra_comprobante(civa_cliente=4, provincia="V ") == "A"
        assert fs.letra_comprobante(civa_cliente=4, provincia="B ") == "B"

    def test_letra_tolera_provincia_sin_padding(self, db):
        fs = FacturaService(db)
        assert fs.letra_comprobante(civa_cliente=4, provincia="V") == "A"
        assert fs.letra_comprobante(civa_cliente=4, provincia=None) == "B"

    def test_total_civa_1_aplica_iva_y_percepcion(self, db):
        _set_parametro(db, ivains="21", mcaib=1)
        fs = FacturaService(db)

        total = fs.calcular_total(
            bruto=Decimal("1000"),
            descuento=Decimal("100"),
            civa_cliente=1,
            porcentaje_iibb=Decimal("3"),
        )

        assert total.neto_gravado == Decimal("900.00")
        assert total.iva == Decimal("189.00")  # 900 * 0.21
        assert total.percepcion_iibb == Decimal("27.00")  # 900 * 0.03
        assert total.total == Decimal("1116.00")

    def test_total_percepcion_deshabilitada_por_parametro(self, db):
        """EmiFact.frm:1177 — ConPercep <> 1 pone ValIB en 0 (Parametro.MCAIB)."""
        _set_parametro(db, ivains="21", mcaib=2)
        fs = FacturaService(db)

        total = fs.calcular_total(
            bruto=Decimal("1000"),
            descuento=Decimal("0"),
            civa_cliente=1,
            porcentaje_iibb=Decimal("3"),
        )

        assert total.percepcion_iibb == Decimal("0")
        assert total.iva == Decimal("210.00")

    def test_total_civa_3_consumidor_final_sin_iva_desglosado(self, db):
        """ClteCIVA >= 3 (ej. Consumidor Final) no dispara el cálculo de IVA
        separado en el legacy (CIVA_GRAVADO_MAX)."""
        _set_parametro(db, ivains="21", mcaib=1)
        fs = FacturaService(db)

        total = fs.calcular_total(
            bruto=Decimal("1000"),
            descuento=Decimal("0"),
            civa_cliente=3,
            porcentaje_iibb=Decimal("3"),
        )

        assert total.iva == Decimal("0")
        assert total.percepcion_iibb == Decimal("0")
        assert total.total == Decimal("1000.00")

    def test_total_civa_2_no_inscripto_es_codigo_muerto_documentado(self, db):
        """Decisión confirmada con el usuario: ClteCIVA=2 (categoría AFIP
        eliminada en 2003) NO dispara IVA en este service — a diferencia
        del legacy, que sí calculaba TotIVANI para esa rama."""
        _set_parametro(db, ivains="21", ivani="10", mcaib=1)
        fs = FacturaService(db)

        total = fs.calcular_total(
            bruto=Decimal("1000"), descuento=Decimal("0"), civa_cliente=2
        )

        # ClteCIVA=2 sigue siendo < CIVA_GRAVADO_MAX (3), o sea el legacy SÍ
        # le aplicaba IVA inscripto además del (ya descartado) IVANI —
        # acá replicamos solo la parte de IVA inscripto, que es la vigente.
        assert total.iva == Decimal("210.00")


# ---------------------------------------------------------------------------
# CuentaCorrienteService
# ---------------------------------------------------------------------------


class TestCuentaCorrienteService:
    def test_saldo_cliente_delegado_al_repositorio_corregido(self, db):
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("1000")),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=4, IMPTE=Decimal("400")),
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        assert cc.saldo_cliente(1) == Decimal("600")

    def test_facturas_pendientes_solo_tipo_1_y_3_con_debe_positivo(self, db):
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("1000"), DEBE=Decimal("1000")),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=1, IMPTE=Decimal("500"), DEBE=Decimal("0")),  # saldada
                Ctascte(CLTE=1, FECHA=date(2026, 1, 3), TIPO=3, IMPTE=Decimal("200"), DEBE=Decimal("200")),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 4), TIPO=4, IMPTE=Decimal("300")),  # recibo, no aplica
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        pendientes = cc.facturas_pendientes(1)
        deuda_por_comprobante = sorted(p.DEBE for p in pendientes)
        assert deuda_por_comprobante == [Decimal("200"), Decimal("1000")]

    # ------------------------------------------------------------------
    # tiene_deuda_vieja (ConDeuda.frm / BusClte.frm TieneDeuda())
    # ------------------------------------------------------------------

    def test_tiene_deuda_vieja_true_cuando_supera_nume20_dias(self, db):
        db.add(Parametro(CLAVE="1", NUME20=90))
        db.add(Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("500")))
        db.commit()

        cc = CuentaCorrienteService(db)
        assert cc.tiene_deuda_vieja(1, hoy=date(2026, 6, 1)) is True  # 151 días

    def test_tiene_deuda_vieja_false_si_no_llega_a_nume20_dias(self, db):
        db.add(Parametro(CLAVE="1", NUME20=90))
        db.add(Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("500")))
        db.commit()

        cc = CuentaCorrienteService(db)
        assert cc.tiene_deuda_vieja(1, hoy=date(2026, 2, 1)) is False  # 31 días

    def test_tiene_deuda_vieja_ignora_facturas_saldadas_y_otros_tipos(self, db):
        db.add(Parametro(CLAVE="1", NUME20=90))
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2020, 1, 1), TIPO=1, DEBE=Decimal("0")),  # saldada, no cuenta
                Ctascte(CLTE=1, FECHA=date(2020, 1, 1), TIPO=1, DEBE=Decimal("0.50")),  # DEBE<=1, no cuenta
                Ctascte(CLTE=1, FECHA=date(2020, 1, 1), TIPO=4, DEBE=Decimal("500")),  # recibo, no es TIPO=1
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        assert cc.tiene_deuda_vieja(1, hoy=date(2026, 1, 1)) is False

    def test_tiene_deuda_vieja_toma_la_factura_impaga_mas_vieja(self, db):
        db.add(Parametro(CLAVE="1", NUME20=90))
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2026, 5, 1), TIPO=1, DEBE=Decimal("500")),  # reciente
                Ctascte(CLTE=1, FECHA=date(2025, 1, 1), TIPO=1, DEBE=Decimal("500")),  # la más vieja
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        # Desde la más vieja (2025-01-01) hasta 2026-06-01 hay > 90 días.
        assert cc.tiene_deuda_vieja(1, hoy=date(2026, 6, 1)) is True

    def test_tiene_deuda_vieja_sin_config_devuelve_false(self, db):
        db.add(Ctascte(CLTE=1, FECHA=date(2020, 1, 1), TIPO=1, DEBE=Decimal("500")))
        db.commit()

        cc = CuentaCorrienteService(db)
        assert cc.tiene_deuda_vieja(1, hoy=date(2026, 1, 1)) is False

    def test_tiene_deuda_vieja_cliente_sin_movimientos_devuelve_false(self, db):
        db.add(Parametro(CLAVE="1", NUME20=90))
        db.commit()
        cc = CuentaCorrienteService(db)
        assert cc.tiene_deuda_vieja(999) is False

    # ------------------------------------------------------------------
    # cobranzas_por_zona (VerCobra.frm)
    # ------------------------------------------------------------------

    def test_cobranzas_por_zona_agrupa_por_cliente_con_subtotales(self, db):
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Beto", ZONA=5),
                Cliente(CODIGO=2, NOMB="Ana", ZONA=5),
                Cliente(CODIGO=3, NOMB="Otra Zona", ZONA=9),  # zona distinta, no debe aparecer
                Ctascte(
                    CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, LETRA="A", PREFIJO=1, CPBTE=100,
                    IMPTE=Decimal("1000"), DEBE=Decimal("1000"), FECVTO=date(2026, 2, 1),
                ),
                Ctascte(
                    CLTE=2, FECHA=date(2026, 1, 5), TIPO=3, LETRA="A", PREFIJO=1, CPBTE=200,
                    IMPTE=Decimal("500"), DEBE=Decimal("500"), FECVTO=date(2026, 3, 1),
                ),
                Ctascte(CLTE=3, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("999")),  # otra zona
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.cobranzas_por_zona(5, hoy=date(2026, 6, 1))

        # Orden por NOMB ("Ana" antes que "Beto"), mismo criterio del legacy.
        assert [c.cliente_nombre for c in resultado.clientes] == ["Ana", "Beto"]
        assert resultado.total_deuda == Decimal("1500")

        ana = resultado.clientes[0]
        assert ana.subtotal_debe == Decimal("500")
        assert len(ana.comprobantes) == 1
        assert ana.comprobantes[0].tipo_label == "N/D."
        assert ana.comprobantes[0].comprobante == "01-000200"

    def test_cobranzas_por_zona_vencido_usa_fecvto_no_fecha_emision(self, db):
        """A diferencia de tiene_deuda_vieja() (que mide FECHA de
        emisión), acá "vencido" se calcula con FECVTO real — dos
        pantallas legacy, dos criterios distintos, ambos fieles."""
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Cliente Uno", ZONA=5),
                Ctascte(
                    CLTE=1, FECHA=date(2020, 1, 1), TIPO=1, PREFIJO=1, CPBTE=1,
                    IMPTE=Decimal("100"), DEBE=Decimal("100"), FECVTO=date(2099, 1, 1),  # vto futuro
                ),
                Ctascte(
                    CLTE=1, FECHA=date(2026, 5, 1), TIPO=1, PREFIJO=1, CPBTE=2,
                    IMPTE=Decimal("200"), DEBE=Decimal("200"), FECVTO=date(2026, 1, 1),  # vto pasado
                ),
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.cobranzas_por_zona(5, hoy=date(2026, 6, 1))

        comprobantes = {c.comprobante: c for c in resultado.clientes[0].comprobantes}
        assert comprobantes["01-000001"].vencido is False  # emitida hace años pero vto futuro
        assert comprobantes["01-000002"].vencido is True
        assert comprobantes["01-000002"].dias_vencido == (date(2026, 6, 1) - date(2026, 1, 1)).days
        assert resultado.total_vencido == Decimal("200")
        assert resultado.total_deuda == Decimal("300")

    def test_cobranzas_por_zona_ignora_debe_menor_o_igual_a_1(self, db):
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Cliente Uno", ZONA=5),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("1.00")),  # umbral, no cuenta
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=4, DEBE=Decimal("500")),  # recibo, TIPO no aplica
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.cobranzas_por_zona(5)
        assert resultado.clientes == []
        assert resultado.total_deuda == Decimal("0")

    def test_cobranzas_por_zona_sin_clientes_devuelve_vacio(self, db):
        cc = CuentaCorrienteService(db)
        resultado = cc.cobranzas_por_zona(999)
        assert resultado.clientes == []
        assert resultado.total_deuda == Decimal("0")
        assert resultado.total_vencido == Decimal("0")

    def test_imputar_pago_resta_debe_y_registra_historico(self, db):
        factura = Ctascte(
            CLTE=1,
            FECHA=date(2026, 1, 5),
            TIPO=1,
            CPBTE=100,
            IMPTE=Decimal("1000"),
            DEBE=Decimal("1000"),
            FECVTO=date(2026, 2, 5),
        )
        db.add(factura)
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.imputar_pago(
            factura,
            importe_aplicado=Decimal("300"),
            corr=7,
            cpbte_recibo=555,
            usuario="clever",
        )

        assert resultado.slot_usado == "IMPUT1"
        assert resultado.debe_anterior == Decimal("1000")
        assert resultado.debe_nuevo == Decimal("700")
        # IMPUT1 es TEXT(2) real en Access (confirmado vía DAO) -> string.
        assert factura.IMPUT1 == "7"

        historico = RepositoryFactory(db).imputacion().by_comprobante_original(1, "1", 100)
        assert len(historico) == 1
        assert historico[0].IMPTE == Decimal("300")
        assert historico[0].TIPO == "4"
        assert historico[0].CPBTE == 555

    def test_imputar_pago_con_descuento_crea_segundo_registro(self, db):
        factura = Ctascte(
            CLTE=1, FECHA=date(2026, 1, 5), TIPO=1, CPBTE=100,
            IMPTE=Decimal("1000"), DEBE=Decimal("1000"),
        )
        db.add(factura)
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.imputar_pago(
            factura, Decimal("300"), corr=1, cpbte_recibo=555,
            usuario="clever", descuento=Decimal("15"),
        )

        assert resultado.descuento_imputacion_id is not None
        historico = RepositoryFactory(db).imputacion().by_comprobante_original(1, "1", 100)
        tipos = {h.TIPO for h in historico}
        assert tipos == {"4", "6"}

    def test_imputar_pago_tipo_nc_suma_al_debe_en_vez_de_restar(self, db):
        """EmiRec.frm:1094-1098: TIPO en {2,5,8} suma al DEBE."""
        nota_credito = Ctascte(
            CLTE=1, FECHA=date(2026, 1, 5), TIPO=2, CPBTE=200,
            IMPTE=Decimal("500"), DEBE=Decimal("100"),
        )
        db.add(nota_credito)
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.imputar_pago(
            nota_credito, Decimal("50"), corr=1, cpbte_recibo=555, usuario="clever",
        )
        assert resultado.debe_nuevo == Decimal("150")

    def test_imputar_pago_usa_siguiente_slot_libre(self, db):
        factura = Ctascte(
            CLTE=1, FECHA=date(2026, 1, 5), TIPO=1, CPBTE=100,
            IMPTE=Decimal("1000"), DEBE=Decimal("1000"), IMPUT1="9",
        )
        db.add(factura)
        db.commit()

        cc = CuentaCorrienteService(db)
        resultado = cc.imputar_pago(factura, Decimal("100"), corr=2, cpbte_recibo=556, usuario="clever")
        assert resultado.slot_usado == "IMPUT2"

    def test_imputar_pago_sin_slots_libres_lanza_error(self, db):
        factura = Ctascte(
            CLTE=1, FECHA=date(2026, 1, 5), TIPO=1, CPBTE=100,
            IMPTE=Decimal("1000"), DEBE=Decimal("1000"),
            IMPUT1="1", IMPUT2="2", IMPUT3="3", IMPUT4="4", IMPUT5="5", IMPUT6="6",
        )
        db.add(factura)
        db.commit()

        cc = CuentaCorrienteService(db)
        with pytest.raises(ValueError):
            cc.imputar_pago(factura, Decimal("100"), corr=7, cpbte_recibo=557, usuario="clever")

    # ------------------------------------------------------------------
    # Consulta (CtaCte.frm)
    # ------------------------------------------------------------------

    def test_extracto_anual_saldo_inicial_usa_tipos_debe_completo(self, db):
        """Corregido 2026-08-16 (cuenta real 407): el Saldo Inicial de
        años previos usa `TIPOS_DEBE` (0/1/3/7 débito, el resto crédito)
        — TODOS los TIPO cuentan, no sólo 1/3/2/4 como hacía la fórmula
        vieja del legacy (ver `CtascteRepository.saldo_inicial`)."""
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2024, 1, 1), TIPO=1, IMPTE=Decimal("1000")),  # débito
                Ctascte(CLTE=1, FECHA=date(2024, 6, 1), TIPO=4, IMPTE=Decimal("300")),  # crédito
                Ctascte(CLTE=1, FECHA=date(2024, 7, 1), TIPO=5, IMPTE=Decimal("9999")),  # crédito (antes se ignoraba)
                Ctascte(CLTE=1, FECHA=date(2025, 3, 1), TIPO=3, IMPTE=Decimal("200")),  # débito
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        extracto = cc.extracto_anual(1, 2026)

        assert extracto.saldo_inicial_debe == Decimal("1200")  # 1000 + 200
        assert extracto.saldo_inicial_credito == Decimal("10299")  # 300 + 9999
        assert extracto.filas[0].tipo_label == "Sdo.Ant."
        assert extracto.filas[0].saldo == Decimal("-9099")  # 1200 - 10299

    def test_extracto_anual_acumula_saldo_corriente_fila_a_fila(self, db):
        db.add_all(
            [
                Ctascte(
                    CLTE=1, FECHA=date(2026, 1, 10), TIPO=1, CPBTE=100, LETRA="B",
                    IMPTE=Decimal("1000"), DEBE=Decimal("1000"),
                ),
                Ctascte(CLTE=1, FECHA=date(2026, 2, 5), TIPO=4, CPBTE=555, IMPTE=Decimal("400")),
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        extracto = cc.extracto_anual(1, 2026)

        assert len(extracto.filas) == 3  # Sdo.Ant. + Factura + Recibo
        fila_factura = extracto.filas[1]
        assert fila_factura.tipo_label == "Fact."
        assert fila_factura.debe == Decimal("1000")
        assert fila_factura.saldo == Decimal("1000")
        assert fila_factura.resto == Decimal("1000")  # RESTO sólo en TIPO 1/3

        fila_recibo = extracto.filas[2]
        assert fila_recibo.tipo_label == "Recibo"
        assert fila_recibo.haber == Decimal("400")
        assert fila_recibo.saldo == Decimal("600")
        assert fila_recibo.resto is None  # RESTO vacío fuera de TIPO 1/3

        assert extracto.total_debe == Decimal("1000")
        assert extracto.total_haber == Decimal("400")
        assert extracto.saldo_final == Decimal("600")

    def test_resumen_cliente_junta_los_5_campos_confirmados(self, db):
        cliente = Cliente(
            CODIGO=50, NOMB="Cliente Resumen", TEL1="011-4444", FALTA=date(2020, 5, 1),
        )
        db.add(cliente)
        db.add_all(
            [
                Cheque(NROCHEQ=1, CLTE=50, ESTADO="1", IMPORTE=Decimal("500")),  # en cartera, cuenta
                Cheque(NROCHEQ=2, CLTE=50, ESTADO="2", IMPORTE=Decimal("9999")),  # egresado, NO cuenta
                FcivaVta(FECHA=date(2026, 3, 1), CLTE=50, TIPO="1"),
                FcivaVta(FECHA=date(2026, 5, 1), CLTE=50, TIPO="1"),  # la más reciente
                # Vencida: FECVTO en el pasado, DEBE > 0.
                Ctascte(
                    CLTE=50, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("700"),
                    DEBE=Decimal("700"), FECVTO=date(2026, 2, 1),
                ),
                # No vencida todavía.
                Ctascte(
                    CLTE=50, FECHA=date(2026, 8, 1), TIPO=1, IMPTE=Decimal("300"),
                    DEBE=Decimal("300"), FECVTO=date(2099, 1, 1),
                ),
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        with patch("migration.repository.date") as mock_date:
            mock_date.today.return_value = date(2026, 8, 16)
            resumen = cc.resumen_cliente(cliente)

        assert resumen.telefono == "011-4444"
        assert resumen.fecha_alta == date(2020, 5, 1)
        assert resumen.valores_pendientes == Decimal("500")
        assert resumen.ultima_venta == date(2026, 5, 1)
        assert resumen.saldo_vencido == Decimal("700")

    def test_saldos_todos_clientes_misma_formula_que_deuda_cliente(self, db):
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Ana"),
                Cliente(CODIGO=2, NOMB="Beto"),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("1000")),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=4, IMPTE=Decimal("400")),
                Ctascte(CLTE=2, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("50")),
            ]
        )
        db.commit()

        cc = CuentaCorrienteService(db)
        saldos = {f["codigo"]: f["saldo"] for f in cc.saldos_todos_clientes()}
        assert saldos[1] == Decimal("600")  # igual que saldo_cliente(1) ya testeado arriba
        assert saldos[2] == Decimal("50")


# ---------------------------------------------------------------------------
# StockService
# ---------------------------------------------------------------------------


class TestStockService:
    def test_articulos_criticos_usa_stunid(self, db):
        db.add(Stock(COD1="A", COD2="", STMIN=Decimal("10"), STUNID=Decimal("5")))
        db.commit()
        ss = StockService(db)
        assert [c.COD1 for c in ss.articulos_criticos()] == ["A"]

    def test_a_reponer_usa_strep(self, db):
        db.add_all(
            [
                Stock(COD1="A", COD2="", STREP=Decimal("20"), STUNID=Decimal("5")),
                Stock(COD1="B", COD2="", STREP=Decimal("20"), STUNID=Decimal("50")),
            ]
        )
        db.commit()
        ss = StockService(db)
        assert [i.COD1 for i in ss.a_reponer()] == ["A"]

    def test_registrar_movimiento_salida_descuenta_stunid_y_suma_salmes(self, db):
        db.add(Stock(COD1="A", COD2="", STUNID=Decimal("50"), SALMES=Decimal("0")))
        db.commit()

        ss = StockService(db)
        item = ss.registrar_movimiento("A", "", Decimal("20"), es_salida=True)

        assert item.STUNID == Decimal("30")
        assert item.SALMES == Decimal("20")

    def test_registrar_movimiento_entrada_suma_stunid_y_entmes(self, db):
        """Corregido 2026-08-16 (ver docstring de `registrar_movimiento`):
        una Entrada suma a ENTMES, nunca resta de SALMES."""
        db.add(Stock(COD1="A", COD2="", STUNID=Decimal("50"), ENTMES=Decimal("0"), SALMES=Decimal("0")))
        db.commit()

        ss = StockService(db)
        item = ss.registrar_movimiento("A", "", Decimal("20"), es_salida=False)

        assert item.STUNID == Decimal("70")
        assert item.ENTMES == Decimal("20")
        assert item.SALMES == Decimal("0")

    def test_registrar_movimiento_item_inexistente_lanza_error(self, db):
        ss = StockService(db)
        with pytest.raises(ValueError):
            ss.registrar_movimiento("NOEXISTE", "", Decimal("1"), es_salida=True)


class TestStockMovimientoService:
    """Tests de `StockMovimientoService`, migración de `Stock.frm`
    ("Ingreso de Movimientos al Stock")."""

    @staticmethod
    def _cargar_seccion(db, cod="A    ", alf7="UNID "):
        db.add_all(
            [
                Fctabla1(CTAB="SC   ", COD=cod, DESCRI="SECCION", ALF7=alf7),
                Fctabla1(CTAB="UM   ", COD=alf7, DESCRI="UNIDADES", NUMSD1=5, NUMSD2=0, NUMSD3=6),
            ]
        )
        db.commit()

    def test_validar_seccion_inexistente_lanza_error(self, db):
        sm = StockMovimientoService(db)
        with pytest.raises(ValueError):
            sm.validar_seccion("ZZZ")

    def test_validar_seccion_con_unidad_invalida_lanza_error(self, db):
        self._cargar_seccion(db, cod="B    ", alf7="PULG ")
        db.add(Fctabla1(CTAB="UM   ", COD="PULG ", DESCRI="PULGADAS", NUMSD1=2, NUMSD2=0, NUMSD3=1))
        db.commit()
        sm = StockMovimientoService(db)
        with pytest.raises(ValueError):
            sm.validar_seccion("B")

    def test_validar_seccion_unid_mtrs_kg_pasa(self, db):
        self._cargar_seccion(db)
        StockMovimientoService(db).validar_seccion("A")  # no debe lanzar

    def test_emitir_movimiento_entrada_fabrica_suma_stock_y_entmes(self, db):
        """13 = Fabricación (confirmado contra Graba() y VerStock.frm)."""
        db.add(Stock(COD1="A", COD2="1", STUNID=Decimal("10"), ENTMES=Decimal("0"), SALMES=Decimal("0")))
        db.commit()

        sm = StockMovimientoService(db)
        resultado = sm.emitir_movimiento(
            es_entrada=True, forma="fabrica", nro_comprobante=100,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("5"))],
            usuario="clever",
        )

        assert resultado.tipo_codigo == "13"
        stock = RepositoryFactory(db).stock().by_cod1_cod2("A", "1")
        assert stock.STUNID == Decimal("15")
        assert stock.ENTMES == Decimal("5")
        assert stock.SALMES == Decimal("0")

        mov = RepositoryFactory(db).movstock().by_cod1_cod2("A", "1")
        assert len(mov) == 1
        assert mov[0].TIPO == "13"
        assert mov[0].CPBTE == 100

    def test_emitir_movimiento_entrada_importacion_y_compra_local_no_estan_invertidos(self, db):
        """Regresión directa del bug encontrado en BuscaCPBTE() (NO
        replicado): Importación -> 15, Compra Local -> 14."""
        sm = StockMovimientoService(db)
        r_importacion = sm.emitir_movimiento(
            es_entrada=True, forma="importacion", nro_comprobante=1,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
            usuario="clever", nro_despacho="LOTE1",
        )
        r_compra = sm.emitir_movimiento(
            es_entrada=True, forma="compra_local", nro_comprobante=2,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
            usuario="clever",
        )
        assert r_importacion.tipo_codigo == "15"
        assert r_compra.tipo_codigo == "14"

    def test_emitir_movimiento_salida_remito_resta_stock_y_suma_salmes(self, db):
        db.add(Stock(COD1="A", COD2="1", STUNID=Decimal("10")))
        db.commit()

        sm = StockMovimientoService(db)
        resultado = sm.emitir_movimiento(
            es_entrada=False, forma="remito", nro_comprobante=200,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("3"))],
            usuario="clever",
        )

        assert resultado.tipo_codigo == "23"
        stock = RepositoryFactory(db).stock().by_cod1_cod2("A", "1")
        assert stock.STUNID == Decimal("7")
        assert stock.SALMES == Decimal("3")

    def test_emitir_movimiento_crea_stock_si_no_existia(self, db):
        sm = StockMovimientoService(db)
        sm.emitir_movimiento(
            es_entrada=True, forma="compra_local", nro_comprobante=300,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="2", cantidad=Decimal("8"))],
            usuario="clever",
        )
        stock = RepositoryFactory(db).stock().by_cod1_cod2("A", "2")
        assert stock is not None
        assert stock.STUNID == Decimal("8")

    def test_emitir_movimiento_forma_no_corresponde_al_sentido_lanza_error(self, db):
        """"remito" es una forma de Salida — pedirla para una Entrada
        tiene que rechazarse, no aceptarse con el código equivocado."""
        sm = StockMovimientoService(db)
        with pytest.raises(ValueError):
            sm.emitir_movimiento(
                es_entrada=True, forma="remito", nro_comprobante=1,
                renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
                usuario="clever",
            )

    def test_emitir_movimiento_importacion_sin_despacho_lanza_error(self, db):
        sm = StockMovimientoService(db)
        with pytest.raises(ValueError):
            sm.emitir_movimiento(
                es_entrada=True, forma="importacion", nro_comprobante=1,
                renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
                usuario="clever",
            )

    def test_emitir_movimiento_sin_renglones_lanza_error(self, db):
        sm = StockMovimientoService(db)
        with pytest.raises(ValueError):
            sm.emitir_movimiento(
                es_entrada=True, forma="fabrica", nro_comprobante=1, renglones=[], usuario="clever",
            )

    def test_emitir_movimiento_importacion_actualiza_despachos(self, db):
        sm = StockMovimientoService(db)
        sm.emitir_movimiento(
            es_entrada=True, forma="importacion", nro_comprobante=400,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("6"))],
            usuario="clever", nro_despacho="LOTE001",
        )
        despacho = RepositoryFactory(db).despacho().by_nrodesp("A", "1", "LOTE001")
        assert despacho is not None
        assert despacho.STOCK == Decimal("6")
        assert despacho.ENTRADA == Decimal("6")

    def test_emitir_movimiento_sin_nro_despacho_no_crea_lote_fantasma(self, db):
        """Regresión del bug real encontrado y NO replicado: `Graba()`
        tocaba `Despachos` con NRODESP='' para CUALQUIER movimiento sin
        Nº de Despacho real — acá sólo se toca la tabla si el llamador
        pasa un `nro_despacho`."""
        sm = StockMovimientoService(db)
        sm.emitir_movimiento(
            es_entrada=True, forma="fabrica", nro_comprobante=500,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("2"))],
            usuario="clever",
        )
        assert db.query(Despacho).count() == 0

    def test_emitir_movimiento_duplicado_lanza_error(self, db):
        sm = StockMovimientoService(db)
        sm.emitir_movimiento(
            es_entrada=True, forma="fabrica", nro_comprobante=600,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
            usuario="clever",
        )
        with pytest.raises(ValueError):
            sm.emitir_movimiento(
                es_entrada=True, forma="fabrica", nro_comprobante=600,
                renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
                usuario="clever",
            )

    def test_emitir_movimiento_mismo_numero_pero_tipo_distinto_no_es_duplicado(self, db):
        """El chequeo de duplicado es por TIPO+CPBTE, no sólo CPBTE — un
        Remito Nº 700 y una Fábrica Nº 700 son comprobantes distintos."""
        sm = StockMovimientoService(db)
        sm.emitir_movimiento(
            es_entrada=True, forma="fabrica", nro_comprobante=700,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
            usuario="clever",
        )
        sm.emitir_movimiento(  # no debe lanzar
            es_entrada=False, forma="remito", nro_comprobante=700,
            renglones=[RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("1"))],
            usuario="clever",
        )

    def test_emitir_movimiento_atomico_revierte_todo_si_falla_a_mitad(self, db):
        db.add(Stock(COD1="A", COD2="1", STUNID=Decimal("10")))
        db.commit()

        sm = StockMovimientoService(db)
        original = sm._grabar_renglon
        llamadas = {"n": 0}

        def _falla_en_el_segundo(*args, **kwargs):
            llamadas["n"] += 1
            if llamadas["n"] == 2:
                raise RuntimeError("falla simulada")
            return original(*args, **kwargs)

        with patch.object(sm, "_grabar_renglon", side_effect=_falla_en_el_segundo):
            with pytest.raises(RuntimeError):
                sm.emitir_movimiento(
                    es_entrada=True, forma="fabrica", nro_comprobante=800,
                    renglones=[
                        RenglonMovimientoStock(cod1="A", cod2="1", cantidad=Decimal("5")),
                        RenglonMovimientoStock(cod1="A", cod2="2", cantidad=Decimal("1")),
                    ],
                    usuario="clever",
                )

        assert db.query(MovStock).filter(MovStock.CPBTE == 800).count() == 0
        stock = RepositoryFactory(db).stock().by_cod1_cod2("A", "1")
        assert stock.STUNID == Decimal("10")  # el ajuste del primer renglón se revirtió


# ---------------------------------------------------------------------------
# ClienteService
# ---------------------------------------------------------------------------


class TestClienteService:
    def test_buscar_por_codigo(self, db):
        db.add(Cliente(CODIGO=1, NOMB="Ana"))
        db.commit()
        cs = ClienteService(db)
        assert [c.NOMB for c in cs.buscar(codigo=1)] == ["Ana"]
        assert cs.buscar(codigo=999) == []

    def test_deuda_cliente_ignora_campo_deuda_cacheado(self, db):
        """Decisión confirmada: recalcular desde Ctasctes, no leer
        Clientes.DEUDA (puede desincronizarse)."""
        db.add(Cliente(CODIGO=1, NOMB="Ana", DEUDA=Decimal("999999")))
        db.add(Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("300")))
        db.commit()

        cs = ClienteService(db)
        assert cs.deuda_cliente(1) == Decimal("300")

    def test_estadisticas_facturacion_filtra_por_rango_y_suma_total(self, db):
        db.add_all(
            [
                FcivaVta(
                    FECHA=date(2026, 1, 15), CLTE=1, TIPO="1", LETRA="A", PTOVTA=3, CPBTE=1,
                    GRINS=Decimal("1000"), IVAINS=Decimal("210"), TOTIB=Decimal("30"),
                ),
                FcivaVta(
                    FECHA=date(2026, 1, 20), CLTE=1, TIPO="1", LETRA="A", PTOVTA=3, CPBTE=2,
                    GRINS=Decimal("500"), IVAINS=Decimal("105"), TOTIB=Decimal("15"),
                ),
                FcivaVta(
                    FECHA=date(2026, 3, 1), CLTE=1, TIPO="1", LETRA="A", PTOVTA=3, CPBTE=3,
                    GRINS=Decimal("9999"), IVAINS=Decimal("0"), TOTIB=Decimal("0"),
                ),  # fuera de rango
            ]
        )
        db.commit()

        cs = ClienteService(db)
        stats = cs.estadisticas_facturacion(1, date(2026, 1, 1), date(2026, 1, 31))

        assert stats.cantidad_comprobantes == 2
        assert stats.total_facturado == Decimal("1860.00")  # (1000+210+30)+(500+105+15)

    def test_puede_dar_de_baja_cliente_sin_saldo_ni_movimientos(self, db):
        db.add(Cliente(CODIGO=1, NOMB="Ana"))
        db.commit()

        puede, motivo = ClienteService(db).puede_dar_de_baja(1)
        assert puede is True
        assert motivo == ""

    def test_puede_dar_de_baja_bloquea_con_saldo_pendiente(self, db):
        db.add(Cliente(CODIGO=1, NOMB="Ana"))
        db.add(Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("300")))
        db.commit()

        puede, motivo = ClienteService(db).puede_dar_de_baja(1)
        assert puede is False
        assert "saldo" in motivo.lower()

    def test_puede_dar_de_baja_bloquea_con_movimientos_aunque_salden_en_cero(self, db):
        """No alcanza con saldo=0: si hay historial en Ctasctes también se
        bloquea (evita dejar huérfano el ledger de cuentas corrientes)."""
        db.add(Cliente(CODIGO=1, NOMB="Ana"))
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("300")),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=4, IMPTE=Decimal("300")),
            ]
        )
        db.commit()

        puede, motivo = ClienteService(db).puede_dar_de_baja(1)
        assert puede is False
        assert "movimientos" in motivo.lower()


# ---------------------------------------------------------------------------
# ArticuloService
# ---------------------------------------------------------------------------


class TestArticuloService:
    def test_buscar_por_clave_compuesta(self, db):
        db.add(Articulo(COD1="GPN", COD2="0300", DESCRI="GOMA 3MM"))
        db.commit()

        svc = ArticuloService(db)
        assert [a.DESCRI for a in svc.buscar(cod1="GPN", cod2="0300")] == ["GOMA 3MM"]
        assert svc.buscar(cod1="GPN", cod2="9999") == []

    def test_buscar_por_descripcion_parcial(self, db):
        db.add(Articulo(COD1="GPN", COD2="0300", DESCRI="GOMA PLANCHA NAT. 3MM"))
        db.commit()

        svc = ArticuloService(db)
        assert len(svc.buscar(descri="plancha")) == 1

    def test_puede_dar_de_baja_articulo_sin_movimientos(self, db):
        db.add(Articulo(COD1="GPN", COD2="0300", DESCRI="GOMA 3MM"))
        db.commit()

        puede, motivo = ArticuloService(db).puede_dar_de_baja("GPN", "0300")
        assert puede is True
        assert motivo == ""

    def test_puede_dar_de_baja_bloquea_si_fue_facturado(self, db):
        db.add(Articulo(COD1="GPN", COD2="0300", DESCRI="GOMA 3MM"))
        db.add(
            Fcestad1(
                COD1="GPN", COD2="0300", TIPO=1, LETRA="A", PTOVTA=3, CPBTE=1,
                ITEM=1, CLTE=1, FECHA=date(2026, 1, 1),
            )
        )
        db.commit()

        puede, motivo = ArticuloService(db).puede_dar_de_baja("GPN", "0300")
        assert puede is False
        assert "movimiento" in motivo.lower()

    def test_puede_dar_de_baja_no_se_confunde_con_otro_articulo_de_la_misma_seccion(self, db):
        """COD1 (Sección) por sí solo no es la clave — un movimiento de OTRO
        artículo de la misma Sección no debe bloquear esta baja."""
        db.add(Articulo(COD1="GPN", COD2="0300", DESCRI="GOMA 3MM"))
        db.add(
            Fcestad1(
                COD1="GPN", COD2="0400", TIPO=1, LETRA="A", PTOVTA=3, CPBTE=1,
                ITEM=1, CLTE=1, FECHA=date(2026, 1, 1),
            )
        )
        db.commit()

        puede, motivo = ArticuloService(db).puede_dar_de_baja("GPN", "0300")
        assert puede is True

    # ------------------------------------------------------------------
    # Modificación de Precios por Porcentaje (ModPrec.frm)
    # ------------------------------------------------------------------

    def test_calcular_precio_con_porcentaje_positivo(self):
        nuevo = ArticuloService.calcular_precio_con_porcentaje(Decimal("100"), Decimal("10"))
        assert nuevo == Decimal("110.00")

    def test_calcular_precio_con_porcentaje_negativo_es_descuento(self):
        nuevo = ArticuloService.calcular_precio_con_porcentaje(Decimal("100"), Decimal("-15"))
        assert nuevo == Decimal("85.00")

    def test_calcular_precio_con_porcentaje_cero_no_cambia(self):
        nuevo = ArticuloService.calcular_precio_con_porcentaje(Decimal("532"), Decimal("0"))
        assert nuevo == Decimal("532.00")

    def test_candidatos_modo_todos_ignora_seccion(self, db):
        db.add_all(
            [
                Articulo(COD1="GPN", COD2="0300", PREC=Decimal("100")),
                Articulo(COD1="A", COD2="100", PREC=Decimal("50")),
            ]
        )
        db.commit()

        candidatos = ArticuloService(db).candidatos_modificacion_precio()
        assert len(candidatos) == 2

    def test_candidatos_por_seccion_sin_rango(self, db):
        db.add_all(
            [
                Articulo(COD1="GPN", COD2="0300", PREC=Decimal("100")),
                Articulo(COD1="A", COD2="100", PREC=Decimal("50")),
            ]
        )
        db.commit()

        candidatos = ArticuloService(db).candidatos_modificacion_precio(cod1="GPN")
        assert [a.COD2 for a in candidatos] == ["0300"]

    def test_candidatos_por_seccion_con_rango_filtra_por_val_del_cod2_completo(self, db):
        """Réplica fiel del filtro real: Val(COD2) tratado como un solo
        número, no por segmento — confirmado en el código de
        CargaGrilla(). Con la sección NRO 'A' (código = número libre)
        el rango se comporta como uno esperaría."""
        db.add_all(
            [
                Articulo(COD1="A", COD2="100", PREC=Decimal("10")),
                Articulo(COD1="A", COD2="150", PREC=Decimal("20")),
                Articulo(COD1="A", COD2="200", PREC=Decimal("30")),
            ]
        )
        db.commit()

        candidatos = ArticuloService(db).candidatos_modificacion_precio(cod1="A", desde=100, hasta=150)
        assert sorted(a.COD2 for a in candidatos) == ["100", "150"]

    def test_candidatos_por_seccion_con_rango_sobre_codigo_compuesto_no_es_por_segmento(self, db):
        """Documenta el comportamiento real (no ideal) para Secciones de
        2 dimensiones: el rango compara el COD2 completo (ej. '0300'
        -> 300), no el segmento de milímetros por separado."""
        db.add_all(
            [
                Articulo(COD1="GPN", COD2="0300", PREC=Decimal("100")),  # Val=300
                Articulo(COD1="GPN", COD2="0501", PREC=Decimal("120")),  # Val=501
            ]
        )
        db.commit()

        candidatos = ArticuloService(db).candidatos_modificacion_precio(cod1="GPN", desde=0, hasta=400)
        assert [a.COD2 for a in candidatos] == ["0300"]

    def test_aplicar_modificacion_precio_actualiza_y_devuelve_cantidad(self, db):
        a1 = Articulo(COD1="A", COD2="100", PREC=Decimal("100"))
        a2 = Articulo(COD1="A", COD2="101", PREC=Decimal("200"))
        db.add_all([a1, a2])
        db.commit()

        cantidad = ArticuloService(db).aplicar_modificacion_precio([a1, a2], Decimal("10"))

        assert cantidad == 2
        assert db.query(Articulo).filter(Articulo.COD2 == "100").first().PREC == Decimal("110.00")
        assert db.query(Articulo).filter(Articulo.COD2 == "101").first().PREC == Decimal("220.00")


# ---------------------------------------------------------------------------
# TablaService
# ---------------------------------------------------------------------------


class TestTablaService:
    def test_buscar_devuelve_solo_la_categoria_pedida(self, db):
        db.add_all(
            [
                Fctabla1(CTAB="SC   ", COD="GPN  ", DESCRI="GOMA PLANCHA NAT."),
                Fctabla1(CTAB="VD   ", COD="1    ", DESCRI="PEREZ"),
            ]
        )
        db.commit()

        resultado = TablaService(db).buscar("SC")
        assert [r.DESCRI for r in resultado] == ["GOMA PLANCHA NAT."]

    def test_puede_dar_de_baja_seccion_sin_articulos(self, db):
        db.add(Fctabla1(CTAB="SC   ", COD="GPN  ", DESCRI="GOMA"))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("SC", "GPN")
        assert puede is True

    def test_puede_dar_de_baja_bloquea_seccion_con_articulos(self, db):
        db.add(Fctabla1(CTAB="SC   ", COD="GPN  ", DESCRI="GOMA"))
        db.add(Articulo(COD1="GPN", COD2="0300", DESCRI="GOMA 3MM"))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("SC", "GPN")
        assert puede is False
        assert "artículos" in motivo.lower()

    def test_puede_dar_de_baja_bloquea_unidad_usada_por_seccion(self, db):
        """MM1 está en ALF1 de GPN — no se puede borrar sin romper el
        armado de código de esa Sección."""
        db.add(Fctabla1(CTAB="SC   ", COD="GPN  ", DESCRI="GOMA", ALF1="MM1  ", ALF2="TELAS"))
        db.add(Fctabla1(CTAB="UM   ", COD="MM1  ", DESCRI="MM", NUMSD1=2, NUMSD2=1))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("UM", "MM1")
        assert puede is False
        assert "GPN" in motivo

    def test_puede_dar_de_baja_unidad_usada_solo_como_alf7(self, db):
        """La Unidad de Facturación (ALF7) también cuenta como uso — no
        sólo los slots de código/precio."""
        db.add(Fctabla1(CTAB="SC   ", COD="AC   ", DESCRI="ACARREADOR", ALF1="MTRS ", ALF7="MTRS "))
        db.add(Fctabla1(CTAB="UM   ", COD="MTRS ", DESCRI="MTRS", NUMSD1=3, NUMSD2=0))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("UM", "MTRS")
        assert puede is False

    def test_puede_dar_de_baja_unidad_sin_uso(self, db):
        db.add(Fctabla1(CTAB="UM   ", COD="KG   ", DESCRI="KILOS", NUMSD1=2, NUMSD2=0))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("UM", "KG")
        assert puede is True

    def test_puede_dar_de_baja_bloquea_vendedor_con_clientes(self, db):
        db.add(Fctabla1(CTAB="VD   ", COD="7    ", DESCRI="PEREZ"))
        db.add(Cliente(CODIGO=1, NOMB="Ana", VEND=7))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("VD", "7")
        assert puede is False
        assert "vendedor" in motivo.lower()

    def test_puede_dar_de_baja_bloquea_zona_con_clientes(self, db):
        db.add(Fctabla1(CTAB="ZN   ", COD="26   ", DESCRI="CENTRAL"))
        db.add(Cliente(CODIGO=1, NOMB="Ana", ZONA=26))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("ZN", "26")
        assert puede is False

    def test_puede_dar_de_baja_bloquea_cond_venta_con_clientes(self, db):
        db.add(Fctabla1(CTAB="CV   ", COD="4    ", DESCRI="CONTADO"))
        db.add(Cliente(CODIGO=1, NOMB="Ana", CVTA=4))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("CV", "4")
        assert puede is False

    def test_puede_dar_de_baja_categorias_sin_referenciador_conocido(self, db):
        """PV/MT/VS: nada las referencia todavía en lo migrado — se
        permite la baja."""
        db.add(Fctabla1(CTAB="PV   ", COD="1    ", DESCRI="BUENOS AIRES"))
        db.commit()

        puede, motivo = TablaService(db).puede_dar_de_baja("PV", "1")
        assert puede is True


# ---------------------------------------------------------------------------
# EmisionFacturaService
# ---------------------------------------------------------------------------


class TestEmisionFacturaService:
    """Réplica de `EmiFact.frm Sub Graba()`, acotada a Factura Letra A/B
    (alcance confirmado). Cada test cita la línea de `Graba()` que
    verifica — ver docstrings en `services.py`."""

    def _cliente(self, db, civa=1, cvta=1, deuda="0"):
        cliente = Cliente(
            CODIGO=100, NOMB="Cliente Prueba", PCIA="B", CVTA=cvta, CIVA=civa,
            CUIT="20111111119", VEND=5, DEUDA=Decimal(deuda),
        )
        db.add(cliente)
        # Fctabla1 CV -> NUMSD3 = días de vencimiento (Graba():2274-2285)
        db.add(Fctabla1(CTAB="CV   ", COD=str(cvta).ljust(5), DESCRI="CONTADO", NUMSD3=15))
        db.commit()
        return cliente

    def _renglon(self, cod1="AA", cod2="1", cantidad="10", precio="100", **extra):
        importe = extra.pop("importe", Decimal(cantidad) * Decimal(precio))
        return RenglonEmision(
            cod1=cod1, cod2=cod2, descripcion="ARTICULO PRUEBA",
            precio_unitario=Decimal(precio), importe=Decimal(importe),
            cantidad_unidades=Decimal(cantidad), **extra,
        )

    def test_emitir_factura_letra_a_actualiza_stock_fcivavta_y_ctasctes(self, db):
        _set_parametro(db)  # IVAINS=21%, MCAIB=1 (percepción habilitada)
        cliente = self._cliente(db, civa=1)
        db.add(Stock(COD1="AA", COD2="1", STUNID=Decimal("50")))
        db.commit()

        renglon = self._renglon(cantidad="10", precio="100")  # importe=1000
        total = FacturaService(db).calcular_total(
            bruto=renglon.importe, descuento=Decimal("0"), civa_cliente=cliente.CIVA,
            porcentaje_iibb=Decimal("3"),
        )

        resultado = EmisionFacturaService(db).emitir_factura(
            cliente=cliente, letra="A", punto_venta=1, numero_comprobante=555,
            renglones=[renglon], total=total, porcentaje_iibb=Decimal("3"),
            usuario="ana", fecha=date(2026, 1, 10),
        )

        stock = db.query(Stock).filter(Stock.COD1 == "AA", Stock.COD2 == "1").one()
        assert stock.STUNID == Decimal("40")  # 50 - 10 (Graba():1864-1867)

        mov = db.query(MovStock).filter(MovStock.CPBTE == 555).one()
        assert mov.TIPO == "21"  # Factura 'A' (Graba():1902-1912)
        assert mov.CANT == Decimal("10")

        fcestad = db.query(Fcestad1).filter(Fcestad1.CPBTE == 555).one()
        assert fcestad.IMPTE == Decimal("1000")
        assert fcestad.CANT == Decimal("10")

        cliente_db = db.query(Cliente).filter(Cliente.CODIGO == 100).one()
        assert cliente_db.DEUDA == total.total  # Graba():2244

        ctascte = db.query(Ctascte).filter(Ctascte.id == resultado.ctascte_id).one()
        assert ctascte.IMPTE == total.total
        # "0 " es el marcador real de slot libre (confirmado contra
        # fcmenu_dev: 204 filas reales con IMPUT1='0 '), el mismo que
        # espera CuentaCorrienteService.imputar_pago().
        assert ctascte.IMPUT1 == "0 "
        assert ctascte.FECVTO == date(2026, 1, 10) + timedelta(days=15)

        fcivavta = db.query(FcivaVta).filter(FcivaVta.id == resultado.fcivavta_id).one()
        assert fcivavta.GRINS == total.neto_gravado
        assert fcivavta.IVAINS == total.iva
        assert fcivavta.TOTCAN == 10

        totales = db.query(Totales).filter(Totales.FECHA == date(2026, 1, 10)).one()
        assert totales.FACA == 1
        assert totales.UNIDA == 10
        assert totales.PVTA == total.neto_gravado

    def test_emitir_factura_bug_stunid_mtr_corregido(self, db):
        """EmiFact.frm:1876 pisaba con `Stock.Mtr` (campo legado no
        sincronizado) en vez de restar de `Stock.STUnid` para renglones
        en la columna Mtr/Kg (posición 2) — acá siempre opera sobre
        `STUnid`, sin importar qué columna trae la cantidad."""
        _set_parametro(db)
        cliente = self._cliente(db)
        db.add(Stock(COD1="GPN", COD2="0300", STUNID=Decimal("100"), MTR=Decimal("9999")))
        db.commit()

        renglon = RenglonEmision(
            cod1="GPN", cod2="0300", descripcion="GOMA", precio_unitario=Decimal("50"),
            importe=Decimal("500"), mtr=Decimal("10"),
        )
        total = FacturaService(db).calcular_total(Decimal("500"), Decimal("0"), cliente.CIVA)

        EmisionFacturaService(db).emitir_factura(
            cliente=cliente, letra="A", punto_venta=1, numero_comprobante=1,
            renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
        )

        stock = db.query(Stock).filter(Stock.COD1 == "GPN", Stock.COD2 == "0300").one()
        assert stock.STUNID == Decimal("90")  # 100 - 10, NO pisado por Mtr=9999
        assert stock.MTR == Decimal("9999")  # campo legado intacto, no es la fuente

    def test_emitir_factura_calcula_bonificacion_fresca_para_fcestad1(self, db):
        """Graba():1819-1827: `Fcestad1.BON` se recalcula fresco vía
        `DxClte`, no reutiliza ningún % ya mostrado en pantalla."""
        _set_parametro(db)
        cliente = self._cliente(db)
        db.add(DtoxClte(CLTE=100, SECCION="AA   ", DTO1=Decimal("10"), DTO2=Decimal("5")))
        db.commit()

        renglon = self._renglon(cantidad="1", precio="1000")  # importe=1000
        total = FacturaService(db).calcular_total(renglon.importe, Decimal("0"), cliente.CIVA)

        EmisionFacturaService(db).emitir_factura(
            cliente=cliente, letra="A", punto_venta=1, numero_comprobante=2,
            renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
        )

        fcestad = db.query(Fcestad1).filter(Fcestad1.CPBTE == 2).one()
        # cascada: 1000*10%=100 -> resto 900*5%=45 -> total 145
        assert fcestad.BON == Decimal("145")

    def test_emitir_factura_descuenta_despacho_elegido(self, db):
        """Graba():1974-2009: el lote elegido descuenta ADEMÁS del stock
        general (trazabilidad, no reemplazo)."""
        _set_parametro(db)
        cliente = self._cliente(db)
        db.add(Despacho(COD1="AA", COD2="1", NRODESP="LOTE1", STOCK=Decimal("30"), SALIDAS=Decimal("0")))
        db.commit()

        renglon = self._renglon(cantidad="5", precio="10", nrodesp_elegido="LOTE1")
        total = FacturaService(db).calcular_total(renglon.importe, Decimal("0"), cliente.CIVA)

        EmisionFacturaService(db).emitir_factura(
            cliente=cliente, letra="A", punto_venta=1, numero_comprobante=3,
            renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
        )

        despacho = db.query(Despacho).filter(Despacho.NRODESP == "LOTE1").one()
        assert despacho.STOCK == Decimal("25")
        assert despacho.SALIDAS == Decimal("5")

    def test_emitir_factura_lote_no_encontrado_no_rompe(self, db):
        """Fiel al `GoTo SinGrabar1` (Graba():1982-1984): si el lote ya no
        existe al emitir, se sigue sin tocar Despachos, no aborta la
        factura completa."""
        _set_parametro(db)
        cliente = self._cliente(db)
        renglon = self._renglon(cantidad="1", precio="10", nrodesp_elegido="NOEXISTE")
        total = FacturaService(db).calcular_total(renglon.importe, Decimal("0"), cliente.CIVA)

        resultado = EmisionFacturaService(db).emitir_factura(
            cliente=cliente, letra="A", punto_venta=1, numero_comprobante=4,
            renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
        )
        assert resultado is not None

    def test_emitir_factura_letra_b_incrementa_contadores_b(self, db):
        _set_parametro(db)
        cliente = self._cliente(db)
        renglon = self._renglon(cantidad="2", precio="10")
        total = FacturaService(db).calcular_total(renglon.importe, Decimal("0"), cliente.CIVA)

        EmisionFacturaService(db).emitir_factura(
            cliente=cliente, letra="B", punto_venta=1, numero_comprobante=5,
            renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
            fecha=date(2026, 2, 1),
        )

        totales = db.query(Totales).filter(Totales.FECHA == date(2026, 2, 1)).one()
        assert totales.FACB == 1
        assert not totales.FACA

        mov = db.query(MovStock).filter(MovStock.CPBTE == 5).one()
        assert mov.TIPO == "22"  # Factura 'B'

    def test_emitir_factura_totales_se_acumulan_en_el_mismo_dia(self, db):
        """Totales tiene PK real por FECHA (upsert diario, no un insert
        por factura) — dos facturas el mismo día suman en la misma fila."""
        _set_parametro(db)
        cliente = self._cliente(db)
        fecha = date(2026, 3, 1)

        for cpbte in (10, 11):
            renglon = self._renglon(cantidad="1", precio="10")
            total = FacturaService(db).calcular_total(renglon.importe, Decimal("0"), cliente.CIVA)
            EmisionFacturaService(db).emitir_factura(
                cliente=cliente, letra="A", punto_venta=1, numero_comprobante=cpbte,
                renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
                fecha=fecha,
            )

        filas = db.query(Totales).filter(Totales.FECHA == fecha).all()
        assert len(filas) == 1
        assert filas[0].FACA == 2
        assert filas[0].UNIDA == 2

    def test_emitir_factura_rechaza_letra_fuera_de_alcance(self, db):
        """Alcance confirmado: sólo Factura A/B — NC/ND/Remito quedan
        para una fase posterior."""
        _set_parametro(db)
        cliente = self._cliente(db)
        renglon = self._renglon()
        total = FacturaService(db).calcular_total(renglon.importe, Decimal("0"), cliente.CIVA)

        with pytest.raises(ValueError):
            EmisionFacturaService(db).emitir_factura(
                cliente=cliente, letra="C", punto_venta=1, numero_comprobante=6,
                renglones=[renglon], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
            )

    def test_emitir_factura_sin_renglones_rechaza(self, db):
        _set_parametro(db)
        cliente = self._cliente(db)
        total = FacturaService(db).calcular_total(Decimal("0"), Decimal("0"), cliente.CIVA)
        with pytest.raises(ValueError):
            EmisionFacturaService(db).emitir_factura(
                cliente=cliente, letra="A", punto_venta=1, numero_comprobante=7,
                renglones=[], total=total, porcentaje_iibb=Decimal("0"), usuario="ana",
            )

    def test_emitir_factura_falla_a_mitad_de_camino_no_deja_nada_a_medias(self, db):
        """Simplificación real sobre las 3 transacciones ADO separadas del
        legacy: acá es una única transacción — una falla en el segundo
        renglón revierte también lo que ya se había escrito del primero."""
        _set_parametro(db)
        cliente = self._cliente(db)
        db.add(Stock(COD1="AA", COD2="1", STUNID=Decimal("50")))
        db.commit()

        renglon_ok = self._renglon(cod1="AA", cod2="1", cantidad="5", precio="10")
        renglon_malo = self._renglon(cod1="BB", cod2="2", cantidad="1", precio="10")
        total = FacturaService(db).calcular_total(
            renglon_ok.importe + renglon_malo.importe, Decimal("0"), cliente.CIVA
        )

        servicio = EmisionFacturaService(db)
        original = servicio._grabar_renglon
        llamadas = {"n": 0}

        def _falla_en_el_segundo(*args, **kwargs):
            llamadas["n"] += 1
            if llamadas["n"] == 2:
                raise RuntimeError("falla simulada")
            return original(*args, **kwargs)

        with patch.object(servicio, "_grabar_renglon", side_effect=_falla_en_el_segundo):
            with pytest.raises(RuntimeError):
                servicio.emitir_factura(
                    cliente=cliente, letra="A", punto_venta=1, numero_comprobante=99,
                    renglones=[renglon_ok, renglon_malo], total=total,
                    porcentaje_iibb=Decimal("0"), usuario="ana",
                )

        assert db.query(MovStock).filter(MovStock.CPBTE == 99).count() == 0
        assert db.query(Ctascte).filter(Ctascte.CPBTE == 99).count() == 0
        stock = db.query(Stock).filter(Stock.COD1 == "AA", Stock.COD2 == "1").one()
        assert stock.STUNID == Decimal("50")  # el descuento del primer renglón se revirtió


# ---------------------------------------------------------------------------
# CtascteRepository.pendientes_cobro / ClienteRepository.proximo_correlativo
# ---------------------------------------------------------------------------


class TestPendientesCobroYCorrelativo:
    def test_pendientes_cobro_incluye_factura_nd_pagoacta_y_recibo_anulado(self, db):
        """Réplica de DetRec.frm Sub DoVer3() (líneas 510-512): TIPO IN
        (1,3,5,7) con DEBE>0 — NC(2)/Recibo(4)/Descuento(6)/NCInt(8)/
        Anulación(9) NUNCA aparecen como "pendiente de cobro" de un
        Recibo nuevo."""
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("100")),  # Factura
                Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=3, DEBE=Decimal("50")),  # ND
                Ctascte(CLTE=1, FECHA=date(2026, 1, 3), TIPO=5, DEBE=Decimal("30")),  # Pago a Cta
                Ctascte(CLTE=1, FECHA=date(2026, 1, 4), TIPO=7, DEBE=Decimal("20")),  # Recibo Anulado
                Ctascte(CLTE=1, FECHA=date(2026, 1, 5), TIPO=2, DEBE=Decimal("10")),  # NC -> no
                Ctascte(CLTE=1, FECHA=date(2026, 1, 6), TIPO=4, DEBE=Decimal("10")),  # Recibo -> no
                Ctascte(CLTE=1, FECHA=date(2026, 1, 7), TIPO=1, DEBE=Decimal("0")),  # saldada -> no
            ]
        )
        db.commit()

        pendientes = RepositoryFactory(db).ctascte().pendientes_cobro(1)

        assert {p.TIPO for p in pendientes} == {1, 3, 5, 7}
        assert len(pendientes) == 4

    def test_pendientes_cobro_excluye_tipo9_cotizacion(self, db):
        db.add_all(
            [
                Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("100"), TIPO9="1"),
                Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=1, DEBE=Decimal("100"), TIPO9="0"),
            ]
        )
        db.commit()

        pendientes = RepositoryFactory(db).ctascte().pendientes_cobro(1)
        assert len(pendientes) == 1
        assert pendientes[0].TIPO9 == "0"

    def test_proximo_correlativo_incrementa(self, db):
        cliente = Cliente(CODIGO=1, NOMB="X", CORR1=5)
        assert RepositoryFactory(db).cliente().proximo_correlativo(cliente) == 6

    def test_proximo_correlativo_sin_valor_arranca_en_1(self, db):
        cliente = Cliente(CODIGO=1, NOMB="X", CORR1=None)
        assert RepositoryFactory(db).cliente().proximo_correlativo(cliente) == 1

    def test_proximo_correlativo_da_la_vuelta_pasado_99(self, db):
        cliente = Cliente(CODIGO=1, NOMB="X", CORR1=99)
        assert RepositoryFactory(db).cliente().proximo_correlativo(cliente) == 1


# ---------------------------------------------------------------------------
# EmisionReciboService
# ---------------------------------------------------------------------------


class TestEmisionReciboService:
    """Réplica de `EmiRec.frm Sub Graba()` — ver decisiones confirmadas
    con el usuario (2026-08-15) en el docstring de `EmisionReciboService`."""

    def _cliente(self, db, corr1=0, deuda="1000"):
        cliente = Cliente(CODIGO=200, NOMB="Cliente Recibo", CORR1=corr1, DEUDA=Decimal(deuda))
        db.add(cliente)
        db.add(Parametro(CLAVE="1", NUME8=999))
        db.commit()
        return cliente

    def _factura_pendiente(self, db, clte=200, cpbte=10, debe="600", tipo=1, fecha=date(2026, 1, 5)):
        factura = Ctascte(
            CLTE=clte, FECHA=fecha, TIPO=tipo, PREFIJO=1, CPBTE=cpbte, LETRA="A",
            IMPUT1="0 ", IMPUT2="0 ", IMPUT3="0 ", IMPUT4="0 ", IMPUT5="0 ", IMPUT6="0 ",
            DEBE=Decimal(debe), IMPTE=Decimal(debe), FECVTO=fecha,
        )
        db.add(factura)
        db.commit()
        return factura

    def test_emitir_recibo_aplica_pago_completo_a_una_factura(self, db):
        cliente = self._cliente(db, corr1=3, deuda="600")
        factura = self._factura_pendiente(db, debe="600")

        resultado = EmisionReciboService(db).emitir_recibo(
            cliente=cliente, numero=1000,
            aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("600"))],
            anticipo=Decimal("0"), importe_efectivo=Decimal("600"),
            cheques=[], retenciones=[], usuario="ana", fecha=date(2026, 2, 1),
        )

        db.refresh(factura)
        assert factura.DEBE == Decimal("0")
        assert factura.IMPUT1 == "4"  # corr = 3 + 1 (ClienteRepository.proximo_correlativo)

        ctascte_recibo = db.query(Ctascte).filter(Ctascte.id == resultado.ctascte_id).one()
        assert ctascte_recibo.TIPO == 4
        assert ctascte_recibo.IMPTE == Decimal("600")
        assert ctascte_recibo.LETRA == "X"
        assert ctascte_recibo.DEBE == Decimal("600")

        cliente_db = db.query(Cliente).filter(Cliente.CODIGO == 200).one()
        assert cliente_db.DEUDA == Decimal("0")
        assert cliente_db.CORR1 == 4

        parametro = db.query(Parametro).filter(Parametro.CLAVE == "1").one()
        assert parametro.NUME8 == 1000

        imputacion = db.query(Imputacion).filter(Imputacion.CPBTE == 1000).one()
        assert imputacion.IMPTE == Decimal("600")
        assert imputacion.CPBTEI == factura.CPBTE

        assert resultado.total_pago == Decimal("600")
        assert resultado.correlativo == 4

    def test_emitir_recibo_con_descuento_crea_fila_agregada_tipo_6(self, db):
        """Además del `Imputacion` TIPO="6" por comprobante que ya deja
        `imputar_pago()`, el Recibo deja su propia fila `Ctascte` TIPO=6
        con el total de descuentos — EmiRec.frm:1207-1238.

        OJO con la semántica real de `aplicado`/`descuento` (`DetRec.frm
        FG1_AfterEdit` Case 4 + `Graba()`:1095-1097): `descuento` es un
        RECORTE del propio `aplicado` (no algo aparte que se suma) — el
        DEBE del comprobante baja por el `aplicado` COMPLETO (acá 1000,
        salda la factura entera), y sólo lo que el operador tiene que
        tender de verdad es `aplicado − descuento` (950)."""
        cliente = self._cliente(db, deuda="1000")
        factura = self._factura_pendiente(db, debe="1000")

        EmisionReciboService(db).emitir_recibo(
            cliente=cliente, numero=1001,
            aplicaciones=[
                AplicacionPago(comprobante=factura, importe_aplicado=Decimal("1000"), descuento=Decimal("50"))
            ],
            anticipo=Decimal("0"), importe_efectivo=Decimal("950"),
            cheques=[], retenciones=[], usuario="ana",
        )

        db.refresh(factura)
        assert factura.DEBE == Decimal("0")  # saldada por completo (1000 - 1000)

        descuento = db.query(Ctascte).filter(Ctascte.TIPO == 6, Ctascte.CPBTE == 1001).one()
        assert descuento.IMPTE == Decimal("50")

    def test_emitir_recibo_anticipo_crea_fila_propia_tipo_5_sin_mezclarse(self, db):
        """Decisión confirmada con el usuario: a diferencia del legacy
        (que no dejaba rastro propio del Anticipo), acá queda trazable
        aparte."""
        cliente = self._cliente(db, deuda="1000")
        factura = self._factura_pendiente(db, debe="700")

        resultado = EmisionReciboService(db).emitir_recibo(
            cliente=cliente, numero=1002,
            aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("700"))],
            anticipo=Decimal("300"), importe_efectivo=Decimal("1000"),
            cheques=[], retenciones=[], usuario="ana",
        )

        anticipo = db.query(Ctascte).filter(Ctascte.TIPO == 5, Ctascte.CPBTE == 1002).one()
        assert anticipo.IMPTE == Decimal("300")

        recibo = db.query(Ctascte).filter(Ctascte.id == resultado.ctascte_id).one()
        assert recibo.IMPTE == Decimal("1000")  # el total pagado sigue completo en el Recibo

    def test_emitir_recibo_diferencia_no_cero_lanza_error_y_no_graba_nada(self, db):
        """Misma regla que gatea 'Grabar' en el legacy (PieRec.frm
        Label304_Change): Pago debe calzar EXACTO con lo que hay que
        cobrar — y nada debe quedar a medio grabar si no calza."""
        cliente = self._cliente(db, deuda="600")
        factura = self._factura_pendiente(db, debe="600")

        with pytest.raises(ValueError):
            EmisionReciboService(db).emitir_recibo(
                cliente=cliente, numero=1003,
                aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("600"))],
                anticipo=Decimal("0"), importe_efectivo=Decimal("500"),  # falta 100
                cheques=[], retenciones=[], usuario="ana",
            )

        assert db.query(Ctascte).filter(Ctascte.CPBTE == 1003).count() == 0
        db.refresh(factura)
        assert factura.DEBE == Decimal("600")  # sin tocar

    def test_emitir_recibo_sin_nada_para_cobrar_lanza_error(self, db):
        cliente = self._cliente(db)
        with pytest.raises(ValueError):
            EmisionReciboService(db).emitir_recibo(
                cliente=cliente, numero=1004, aplicaciones=[], anticipo=Decimal("0"),
                importe_efectivo=Decimal("0"), cheques=[], retenciones=[], usuario="ana",
            )

    def test_emitir_recibo_con_cheque_lo_persiste(self, db):
        cliente = self._cliente(db, deuda="500")
        factura = self._factura_pendiente(db, debe="500")

        EmisionReciboService(db).emitir_recibo(
            cliente=cliente, numero=1005,
            aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("500"))],
            anticipo=Decimal("0"), importe_efectivo=Decimal("0"),
            cheques=[
                PagoCheque(
                    nro_cheque=555111, banco="007", fecha_emision=date(2026, 2, 1),
                    fecha_vencimiento=date(2026, 3, 1), importe=Decimal("500"), a_la_orden=True,
                )
            ],
            retenciones=[], usuario="ana",
        )

        cheque = db.query(Cheque).filter(Cheque.NROCHEQ == 555111).one()
        assert cheque.CLTE == 200
        assert cheque.CPBING == 1005
        assert cheque.IMPORTE == Decimal("500")
        assert cheque.ORDEN == 1
        assert cheque.ESTADO == "1"

    def test_emitir_recibo_cheque_repetido_lanza_error_en_vez_de_perderlo(self, db):
        """Bug real del legacy NO replicado (EmiRec.frm:1267-1273): el
        legacy deduplicaba por NROCHEQ en silencio, perdiendo el pago del
        segundo cliente. Acá avisa con ValueError."""
        db.add(Cheque(NROCHEQ=999, CLTE=1, CPBING=1, IMPORTE=Decimal("10")))
        db.commit()
        cliente = self._cliente(db, deuda="500")
        factura = self._factura_pendiente(db, debe="500")

        with pytest.raises(ValueError):
            EmisionReciboService(db).emitir_recibo(
                cliente=cliente, numero=1006,
                aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("500"))],
                anticipo=Decimal("0"), importe_efectivo=Decimal("0"),
                cheques=[
                    PagoCheque(
                        nro_cheque=999, banco="007", fecha_emision=date(2026, 2, 1),
                        fecha_vencimiento=date(2026, 3, 1), importe=Decimal("500"),
                    )
                ],
                retenciones=[], usuario="ana",
            )

    def test_emitir_recibo_con_retencion_la_persiste_en_movimvs(self, db):
        cliente = self._cliente(db, deuda="500")
        factura = self._factura_pendiente(db, debe="500")

        EmisionReciboService(db).emitir_recibo(
            cliente=cliente, numero=1007,
            aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("500"))],
            anticipo=Decimal("0"), importe_efectivo=Decimal("0"),
            cheques=[],
            retenciones=[PagoRetencion(tipreg=1, importe=Decimal("500"), concepto="Retencion Ganancias")],
            usuario="ana",
        )

        mov = db.query(MovimVS).filter(MovimVS.CPBTE == 1007).one()
        assert mov.TIPREG == "1"
        assert mov.IMPTE == Decimal("500")
        assert mov.TIPO == "4"

    def test_emitir_recibo_efectivo_lo_persiste(self, db):
        cliente = self._cliente(db, deuda="200")
        factura = self._factura_pendiente(db, debe="200")

        EmisionReciboService(db).emitir_recibo(
            cliente=cliente, numero=1008,
            aplicaciones=[AplicacionPago(comprobante=factura, importe_aplicado=Decimal("200"))],
            anticipo=Decimal("0"), importe_efectivo=Decimal("200"),
            cheques=[], retenciones=[], usuario="ana",
        )

        efectivo = db.query(Efectivo).filter(Efectivo.CPBTE == 1008).one()
        assert efectivo.IMPTE == Decimal("200")
        assert efectivo.TIPO == 4


# ---------------------------------------------------------------------------
# ChequeService (VerCheq.frm)
# ---------------------------------------------------------------------------


class TestChequeService:
    def test_buscar_filtra_por_estado_y_fecha_desde_y_limita(self, db):
        db.add_all(
            [
                Cheque(NROCHEQ=1, ESTADO="1", FECING=date(2026, 1, 1), FECVTO=date(2026, 2, 1)),
                Cheque(NROCHEQ=2, ESTADO="1", FECING=date(2026, 1, 5), FECVTO=date(2026, 1, 15)),
                Cheque(NROCHEQ=3, ESTADO="1", FECING=date(2025, 1, 1), FECVTO=date(2025, 2, 1)),  # antes del filtro
                Cheque(NROCHEQ=4, ESTADO="2", FECING=date(2026, 1, 1), FECVTO=date(2026, 2, 1)),  # otro estado
            ]
        )
        db.commit()

        cs = ChequeService(db)
        resultado = cs.buscar("1", fecha_desde=date(2026, 1, 1), limite=50)

        # Orden por FECVTO asc: el 2 (15/1) antes que el 1 (1/2).
        assert [c.NROCHEQ for c in resultado] == [2, 1]

    def test_buscar_respeta_el_limite(self, db):
        for i in range(1, 6):
            db.add(Cheque(NROCHEQ=i, ESTADO="1", FECING=date(2026, 1, 1), FECVTO=date(2026, 1, i)))
        db.commit()

        cs = ChequeService(db)
        resultado = cs.buscar("1", fecha_desde=date(2026, 1, 1), limite=3)
        assert len(resultado) == 3

    def test_registrar_egreso_marca_egresado_y_completa_datos(self, db):
        db.add(Cheque(NROCHEQ=100, ESTADO="1"))
        db.commit()

        cs = ChequeService(db)
        cheque = cs.registrar_egreso(
            100, destino_index=1, cpbte_respaldo=555, descripcion="Pago a Proveedor X", fecha_egreso=date(2026, 3, 1)
        )

        assert cheque.ESTADO == "2"
        assert cheque.TIPEGR == 1
        assert cheque.CPBEGR == 555
        assert cheque.DESTINO == "Pago a Proveedor X"
        assert cheque.FECEGR == date(2026, 3, 1)

    def test_registrar_egreso_cheque_inexistente_lanza_error(self, db):
        cs = ChequeService(db)
        with pytest.raises(ValueError):
            cs.registrar_egreso(999, 0, 1, "x", date(2026, 1, 1))

    def test_registrar_egreso_cheque_no_en_cartera_lanza_error(self, db):
        db.add(Cheque(NROCHEQ=101, ESTADO="2"))  # ya egresado
        db.commit()

        cs = ChequeService(db)
        with pytest.raises(ValueError):
            cs.registrar_egreso(101, 0, 1, "x", date(2026, 1, 1))

    def test_registrar_egreso_destino_invalido_lanza_error(self, db):
        db.add(Cheque(NROCHEQ=102, ESTADO="1"))
        db.commit()

        cs = ChequeService(db)
        with pytest.raises(ValueError):
            cs.registrar_egreso(102, destino_index=99, cpbte_respaldo=1, descripcion="x", fecha_egreso=date(2026, 1, 1))

    def test_eliminar_borra_el_registro(self, db):
        db.add(Cheque(NROCHEQ=200, ESTADO="1"))
        db.commit()

        cs = ChequeService(db)
        cs.eliminar(200)

        assert db.query(Cheque).filter(Cheque.NROCHEQ == 200).first() is None

    def test_eliminar_cheque_inexistente_lanza_error(self, db):
        cs = ChequeService(db)
        with pytest.raises(ValueError):
            cs.eliminar(999)


# ---------------------------------------------------------------------------
# EstadisticaVentasService (ESTADIST.frm)
# ---------------------------------------------------------------------------


class TestEstadisticaVentasService:
    def test_fin_de_mes_exclusivo_da_el_primer_dia_del_mes_siguiente(self, db):
        es = EstadisticaVentasService(db)
        assert es.fin_de_mes_exclusivo(date(2026, 8, 1)) == date(2026, 9, 1)
        assert es.fin_de_mes_exclusivo(date(2026, 8, 15)) == date(2026, 9, 1)
        assert es.fin_de_mes_exclusivo(date(2026, 8, 31)) == date(2026, 9, 1)  # caso especial legacy, mismo resultado
        assert es.fin_de_mes_exclusivo(date(2026, 12, 5)) == date(2027, 1, 1)  # cruce de año
        assert es.fin_de_mes_exclusivo(date(2024, 2, 10)) == date(2024, 3, 1)  # bisiesto, sin importar

    def test_ventas_seccion_por_rango_filtra_por_seccion_y_rango(self, db):
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Cliente Uno"),
                Fcestad1(COD1="A", TIPO=1, LETRA="A", PTOVTA=1, CPBTE=100, CLTE=1, FECHA=date(2026, 8, 5), IMPTE=Decimal("1000")),
                Fcestad1(COD1="A", TIPO=1, LETRA="A", PTOVTA=1, CPBTE=101, CLTE=1, FECHA=date(2026, 7, 31), IMPTE=Decimal("999")),  # antes del rango
                Fcestad1(COD1="A", TIPO=1, LETRA="A", PTOVTA=1, CPBTE=102, CLTE=1, FECHA=date(2026, 9, 1), IMPTE=Decimal("999")),  # después del rango
                Fcestad1(COD1="B", TIPO=1, LETRA="A", PTOVTA=1, CPBTE=103, CLTE=1, FECHA=date(2026, 8, 5), IMPTE=Decimal("999")),  # otra sección
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_seccion_por_rango("A", date(2026, 8, 1), date(2026, 8, 31))

        assert resultado.cantidad_movimientos == 1
        assert resultado.movimientos[0].comprobante == "0001-000100"
        assert resultado.movimientos[0].cliente_nombre == "Cliente Uno"
        assert resultado.total_importe == Decimal("1000")

    def test_ventas_seccion_por_rango_hasta_es_inclusive(self, db):
        db.add_all(
            [
                Fcestad1(COD1="A", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 20), IMPTE=Decimal("100")),
                Fcestad1(COD1="A", TIPO=1, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 21), IMPTE=Decimal("200")),  # justo al día siguiente del "hasta"
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_seccion_por_rango("A", date(2026, 8, 15), date(2026, 8, 20))
        assert resultado.cantidad_movimientos == 1
        assert resultado.total_importe == Decimal("100")

    def test_ventas_seccion_por_rango_no_esta_limitado_a_un_solo_mes(self, db):
        """Feedback del usuario (2026-08-16): a diferencia del legacy
        (un día + fin de ESE mes automático), acá el rango puede cruzar
        meses libremente."""
        db.add_all(
            [
                Fcestad1(COD1="A", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 7, 15), IMPTE=Decimal("100")),
                Fcestad1(COD1="A", TIPO=1, PTOVTA=1, CPBTE=2, FECHA=date(2026, 9, 15), IMPTE=Decimal("200")),
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_seccion_por_rango("A", date(2026, 7, 1), date(2026, 9, 30))
        assert resultado.cantidad_movimientos == 2
        assert resultado.total_importe == Decimal("300")

    def test_ventas_seccion_por_rango_total_resta_segun_tipo(self, db):
        db.add_all(
            [
                Fcestad1(COD1="A", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), IMPTE=Decimal("1000")),  # suma
                Fcestad1(COD1="A", TIPO=2, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 6), IMPTE=Decimal("300")),  # NC, resta
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_seccion_por_rango("A", date(2026, 8, 1), date(2026, 8, 31))
        assert resultado.total_importe == Decimal("700")
        etiquetas = {m.comprobante: m.tipo_label for m in resultado.movimientos}
        assert etiquetas["0001-000001"] == "Fact."
        assert etiquetas["0001-000002"] == "NCréd."

    def test_ventas_seccion_por_rango_cliente_desconocido(self, db):
        db.add(Fcestad1(COD1="A", TIPO=1, PTOVTA=1, CPBTE=1, CLTE=99999, FECHA=date(2026, 8, 5), IMPTE=Decimal("100")))
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_seccion_por_rango("A", date(2026, 8, 1), date(2026, 8, 31))
        assert resultado.movimientos[0].cliente_nombre == "*** Desconocido ***"

    def test_ventas_seccion_por_rango_sin_movimientos(self, db):
        es = EstadisticaVentasService(db)
        resultado = es.ventas_seccion_por_rango("ZZ", date(2026, 8, 1), date(2026, 8, 31))
        assert resultado.movimientos == []
        assert resultado.total_importe == Decimal("0")

    # ------------------------------------------------------------------
    # ventas_articulo_agrupadas (VTAXART.frm)
    # ------------------------------------------------------------------

    def test_ventas_articulo_agrupadas_seccion_default_agrupa_por_pulg(self, db):
        db.add_all(
            [
                Fcestad1(COD1="XX", COD2="0001", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), PULG=Decimal("3"), IMPTE=Decimal("100")),
                Fcestad1(COD1="XX", COD2="0002", TIPO=1, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 6), PULG=Decimal("3"), IMPTE=Decimal("200")),
                Fcestad1(COD1="XX", COD2="0003", TIPO=1, PTOVTA=1, CPBTE=3, FECHA=date(2026, 8, 7), PULG=Decimal("5"), IMPTE=Decimal("50")),
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_articulo_agrupadas("XX", date(2026, 8, 1))

        assert resultado.niveles_agrupamiento == ("pulg",)
        assert len(resultado.raiz) == 2  # dos valores distintos de Pulg: 3 y 5
        nodo_3 = next(n for n in resultado.raiz if n.etiqueta == "3,00")
        assert len(nodo_3.movimientos) == 2
        assert nodo_3.subtotales["importe"] == Decimal("300")
        assert nodo_3.hijos == []

    def test_ventas_articulo_agrupadas_gps_agrupa_por_mm_y_telas(self, db):
        db.add_all(
            [
                Fcestad1(COD1="GPS  ", COD2="0001", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), MILIM=3, TELAS=0, MTR=Decimal("10"), IMPTE=Decimal("100")),
                Fcestad1(COD1="GPS  ", COD2="0002", TIPO=1, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 6), MILIM=3, TELAS=1, MTR=Decimal("20"), IMPTE=Decimal("200")),
                Fcestad1(COD1="GPS  ", COD2="0003", TIPO=1, PTOVTA=1, CPBTE=3, FECHA=date(2026, 8, 7), MILIM=4, TELAS=0, MTR=Decimal("5"), IMPTE=Decimal("50")),
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_articulo_agrupadas("GPS", date(2026, 8, 1))

        assert resultado.niveles_agrupamiento == ("milim", "telas")
        assert len(resultado.raiz) == 2  # MM=3 y MM=4
        nodo_mm3 = next(n for n in resultado.raiz if n.etiqueta == "3")
        assert nodo_mm3.hijos != []  # nivel intermedio: sin movimientos propios
        assert nodo_mm3.movimientos == []
        assert len(nodo_mm3.hijos) == 2  # Telas=0 y Telas=1
        assert nodo_mm3.subtotales["mtr"] == Decimal("30")  # 10+20, subtotal del nivel MM agregando ambos Telas
        nodo_telas0 = next(h for h in nodo_mm3.hijos if h.etiqueta == "0")
        assert len(nodo_telas0.movimientos) == 1
        assert nodo_telas0.subtotales["mtr"] == Decimal("10")

    def test_ventas_articulo_agrupadas_pl_tiene_3_niveles(self, db):
        db.add(Fcestad1(COD1="PL   ", COD2="0001", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), PULG=Decimal("2"), MILIM=3, TELAS=1, MTR=Decimal("10"), IMPTE=Decimal("100")))
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_articulo_agrupadas("PL", date(2026, 8, 1))
        assert resultado.niveles_agrupamiento == ("pulg", "milim", "telas")
        assert resultado.raiz[0].etiqueta == "2,00"
        assert resultado.raiz[0].hijos[0].etiqueta == "3"
        assert resultado.raiz[0].hijos[0].hijos[0].etiqueta == "1"
        assert len(resultado.raiz[0].hijos[0].hijos[0].movimientos) == 1

    def test_ventas_articulo_agrupadas_sf_agrupa_por_mtr_suma_cantidad_e_importe(self, db):
        db.add_all(
            [
                Fcestad1(COD1="SF   ", COD2="0001", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), MTR=Decimal("10"), CANT=Decimal("2"), IMPTE=Decimal("100")),
                Fcestad1(COD1="SF   ", COD2="0002", TIPO=1, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 6), MTR=Decimal("10"), CANT=Decimal("3"), IMPTE=Decimal("150")),
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_articulo_agrupadas("SF", date(2026, 8, 1))
        assert resultado.niveles_agrupamiento == ("mtr",)
        assert len(resultado.raiz) == 1
        nodo = resultado.raiz[0]
        assert nodo.subtotales == {"cantidad": Decimal("5"), "importe": Decimal("250")}
        assert "mtr" not in nodo.subtotales  # SF no subtotaliza Mtr, sólo Cant/Importe

    def test_ventas_articulo_agrupadas_filtra_por_rango_de_codigo(self, db):
        db.add_all(
            [
                Fcestad1(COD1="XX", COD2="0050", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), PULG=Decimal("1"), IMPTE=Decimal("100")),
                Fcestad1(COD1="XX", COD2="0150", TIPO=1, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 6), PULG=Decimal("2"), IMPTE=Decimal("200")),
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_articulo_agrupadas("XX", date(2026, 8, 1), cod2_desde=0, cod2_hasta=100)
        assert resultado.cantidad_movimientos == 1
        assert resultado.total_importe == Decimal("100")

    def test_ventas_articulo_agrupadas_subtotal_es_suma_cruda_sin_signo_por_tipo(self, db):
        """El subtotal de grupo (FG3.SubTotal real) suma Impte crudo, a
        diferencia del total_importe general que sí es signado por TIPO."""
        db.add_all(
            [
                Fcestad1(COD1="XX", COD2="0001", TIPO=1, PTOVTA=1, CPBTE=1, FECHA=date(2026, 8, 5), PULG=Decimal("1"), IMPTE=Decimal("1000")),  # factura, suma
                Fcestad1(COD1="XX", COD2="0002", TIPO=2, PTOVTA=1, CPBTE=2, FECHA=date(2026, 8, 6), PULG=Decimal("1"), IMPTE=Decimal("300")),  # NC, resta en el total general
            ]
        )
        db.commit()

        es = EstadisticaVentasService(db)
        resultado = es.ventas_articulo_agrupadas("XX", date(2026, 8, 1))
        assert resultado.raiz[0].subtotales["importe"] == Decimal("1300")  # crudo, 1000+300
        assert resultado.total_importe == Decimal("700")  # signado, 1000-300


# ---------------------------------------------------------------------------
# FacturasEmitidasService (VerFact.frm)
# ---------------------------------------------------------------------------


class TestFacturasEmitidasService:
    def test_listar_filtra_desde_fecha_hasta_fin_de_mes(self, db):
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 7, 31), TIPO="1", LETRA="A", PTOVTA=1, CPBTE=1, CLTE=1, NOMB="Antes del rango", GRINS=Decimal("100"), IVAINS=Decimal("21")),
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", LETRA="A", PTOVTA=1, CPBTE=2, CLTE=1, NOMB="Cliente Uno", GRINS=Decimal("1000"), IVAINS=Decimal("210")),
                FcivaVta(FECHA=date(2026, 8, 31), TIPO="1", LETRA="A", PTOVTA=1, CPBTE=3, CLTE=1, NOMB="Cliente Uno", GRINS=Decimal("500"), IVAINS=Decimal("105")),
                FcivaVta(FECHA=date(2026, 9, 1), TIPO="1", LETRA="A", PTOVTA=1, CPBTE=4, CLTE=1, NOMB="Después del rango", GRINS=Decimal("999"), IVAINS=Decimal("0")),
            ]
        )
        db.commit()

        servicio = FacturasEmitidasService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=50)

        assert [f.cpbte for f in resultado.filas] == [2, 3]
        assert resultado.filas[0].comprobante == "A 0001-00000002"
        assert resultado.filas[0].tipo_label == "Fact."
        assert resultado.total_bruto == Decimal("1500")
        assert resultado.total_iva == Decimal("315")
        assert resultado.total_neto == Decimal("1815")

    def test_listar_respeta_el_limite_los_primeros_n(self, db):
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, d), TIPO="1", LETRA="A", PTOVTA=1, CPBTE=d, CLTE=1, GRINS=Decimal("100"), IVAINS=Decimal("0"))
                for d in range(1, 6)
            ]
        )
        db.commit()

        servicio = FacturasEmitidasService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=3)
        assert len(resultado.filas) == 3
        assert [f.cpbte for f in resultado.filas] == [1, 2, 3]  # ordenado por fecha/cpbte ascendente

    def test_listar_resta_tipos_que_no_suman(self, db):
        """VerFact.frm:1139: sólo TIPO 1/3/7 (Factura/ND/Fact.Mostrador)
        suman al total — el resto (ej. NC=2) resta."""
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", LETRA="A", PTOVTA=1, CPBTE=1, CLTE=1, GRINS=Decimal("1000"), IVAINS=Decimal("0")),
                FcivaVta(FECHA=date(2026, 8, 6), TIPO="2", LETRA="A", PTOVTA=1, CPBTE=2, CLTE=1, GRINS=Decimal("300"), IVAINS=Decimal("0")),
            ]
        )
        db.commit()

        servicio = FacturasEmitidasService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=50)
        assert resultado.filas[1].tipo_label == "NCréd."
        assert resultado.total_bruto == Decimal("700")  # 1000 - 300

    def test_listar_etiquetas_fcivavta_distintas_de_ctascte_en_tipo_7_y_8(self, db):
        """`FcivaVta.TIPO` 7/8 son "Factura Mostrador"/"Devolución
        Mostrador" (no "RecAn."/"NCInt" como en `Ctascte.TIPO`) — dos
        enumeraciones distintas, ver ETIQUETAS_TIPO_FCIVAVTA."""
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="7", LETRA="A", PTOVTA=1, CPBTE=1, CLTE=1, GRINS=Decimal("100"), IVAINS=Decimal("0")),
                FcivaVta(FECHA=date(2026, 8, 6), TIPO="8", LETRA="A", PTOVTA=1, CPBTE=2, CLTE=1, GRINS=Decimal("50"), IVAINS=Decimal("0")),
            ]
        )
        db.commit()

        servicio = FacturasEmitidasService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=50)
        assert resultado.filas[0].tipo_label == "F.Most."
        assert resultado.filas[1].tipo_label == "Dev.M."


# ---------------------------------------------------------------------------
# TotalesDiariosService (TotFact.frm)
# ---------------------------------------------------------------------------


class TestTotalesDiariosService:
    def test_listar_filtra_desde_fecha_hasta_fin_de_mes_inclusive(self, db):
        """A diferencia de FacturasEmitidasService (VerFact.frm), acá el
        fin de mes es INCLUSIVE (`FECHA <= FECHAHst` en el legacy, no
        exclusivo con empuje al mes siguiente)."""
        db.add_all(
            [
                Totales(FECHA=date(2026, 7, 31), PVTA=Decimal("999")),  # antes del rango
                Totales(FECHA=date(2026, 8, 5), PVTA=Decimal("100")),
                Totales(FECHA=date(2026, 8, 31), PVTA=Decimal("200")),  # último día, INCLUSIVE
                Totales(FECHA=date(2026, 9, 1), PVTA=Decimal("999")),  # después del rango
            ]
        )
        db.commit()

        servicio = TotalesDiariosService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=50)
        assert [f.fecha for f in resultado.filas] == [date(2026, 8, 5), date(2026, 8, 31)]
        assert resultado.total_precio_venta == Decimal("300")

    def test_listar_respeta_el_limite(self, db):
        db.add_all([Totales(FECHA=date(2026, 8, d), PVTA=Decimal("1")) for d in range(1, 6)])
        db.commit()

        servicio = TotalesDiariosService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=3)
        assert len(resultado.filas) == 3
        assert [f.fecha for f in resultado.filas] == [date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)]

    def test_listar_completa_facturas_y_notas_credito_tipo_c(self, db):
        """TotFact.frm:1917-1921 (TOTFAC/TOTNC) se olvida de FACC/NCC —
        acá se completa (ver docstring de `listar`)."""
        db.add_all(
            [
                Totales(
                    FECHA=date(2026, 8, 5),
                    FACA=1, FACB=2, FACC=3, EXPA=4, MOSTA=5, MOSTB=6,
                    NCA=1, NCB=1, NCC=1, NCEXPA=1, NCMOSTA=1, NCMOSTB=1,
                )
            ]
        )
        db.commit()

        servicio = TotalesDiariosService(db)
        resultado = servicio.listar(date(2026, 8, 1), limite=50)
        assert resultado.filas[0].cantidad_facturas == 1 + 2 + 3 + 4 + 5 + 6
        assert resultado.filas[0].cantidad_notas_credito == 6

    def test_resumen_mensual_suma_desde_el_dia_1_del_mes_hasta_la_fecha_elegida(self, db):
        """Corrección de rango confirmada con el usuario (2026-08-16):
        [día 1 del mes, fecha elegida], NO [fecha elegida, fin de mes)
        como hace el legacy — ver docstring de `resumen_mensual`."""
        db.add_all(
            [
                Totales(FECHA=date(2026, 7, 31), PESP=Decimal("8888"), UNIDA=999),  # antes del mes, excluido
                Totales(FECHA=date(2026, 8, 1), PESP=Decimal("100"), PVTA=Decimal("1000"), PCOS=Decimal("800"), UNIDA=10),
                Totales(FECHA=date(2026, 8, 5), PESP=Decimal("200"), UNIDA=20),
                Totales(FECHA=date(2026, 8, 10), PESP=Decimal("300"), UNIDA=30),  # fecha elegida
                Totales(FECHA=date(2026, 8, 15), PESP=Decimal("9999"), UNIDA=999),  # después de la elegida, excluido
            ]
        )
        db.commit()

        servicio = TotalesDiariosService(db)
        resumen = servicio.resumen_mensual(date(2026, 8, 10))  # lunes

        assert resumen.fecha_desde == date(2026, 8, 1)
        assert resumen.fecha_hasta == date(2026, 8, 10)
        assert resumen.total_venta_real == Decimal("600")  # 100+200+300, sin las filas excluidas
        assert resumen.total_precio_venta == Decimal("1000")
        assert resumen.total_precio_costo == Decimal("800")
        assert resumen.total_unidades == 60  # 10+20+30

        # Agosto 2026: domingos 2/9/16/23/30 -> 26 días hábiles en el mes.
        assert resumen.dias_habiles_mes == 26
        # Días 1..10 de agosto 2026: domingos 2 y 9 -> 8 días hábiles.
        assert resumen.dias_trabajados == 8

        assert resumen.promedio_diario == Decimal("600") / 8
        # Fórmula corregida: Promedio Diario x Días Hábiles del mes (no al
        # revés, ver el bug de Text1/Text2 invertidos documentado arriba).
        assert resumen.proyeccion == resumen.promedio_diario * 26
        assert resumen.unidad_promedio == Decimal("60") / 8

    def test_resumen_mensual_sin_datos_da_none(self, db):
        servicio = TotalesDiariosService(db)
        assert servicio.resumen_mensual(date(2026, 8, 10)) is None

    def test_dias_habiles_cuenta_lunes_a_sabado_excluye_domingo(self):
        # Agosto 2026: día 1 = sábado, día 2 = domingo.
        assert TotalesDiariosService._dias_habiles(2026, 8, 1) == 1  # sólo el sábado
        assert TotalesDiariosService._dias_habiles(2026, 8, 2) == 1  # + domingo excluido, sigue en 1
        assert TotalesDiariosService._dias_habiles(2026, 8, 31) == 26  # mes completo, 5 domingos

    def test_detalle_dia_calcula_los_6_promedios_derivados(self, db):
        db.add(
            Totales(
                FECHA=date(2026, 8, 5),
                PESPA=Decimal("1000"), PESPB=Decimal("500"), UNIDA=10, UNIDB=5,
                PESPMOSTA=Decimal("300"), PESPMOSTB=Decimal("100"), UNMOSTA=4, UNMOSTB=1,
                PESPC=Decimal("200"), PESPMOSTC=Decimal("50"), UNIDC=5, UNMOSTC=0,
                FACA=2, FACB=1, MOSTA=1, MOSTB=1, FACC=1, MOSTC=0,
            )
        )
        db.commit()

        servicio = TotalesDiariosService(db)
        detalle = servicio.detalle_dia(date(2026, 8, 5))

        assert detalle is not None
        assert detalle.precio_prom_cta_cte == Decimal("1500") / 15  # (1000+500)/(10+5)
        assert detalle.precio_prom_mostrador == Decimal("400") / 5  # (300+100)/(4+1)
        assert detalle.precio_prom_tipo_c == Decimal("250") / 5  # (200+50)/(5+0)
        assert detalle.unid_prom_cta_cte == Decimal("15") / 3  # (10+5)/(2+1)
        assert detalle.unid_prom_mostrador == Decimal("5") / 2  # (4+1)/(1+1)
        assert detalle.unid_prom_tipo_c == Decimal("5") / 1  # (5+0)/(1+0)

    def test_detalle_dia_division_por_cero_da_cero_no_excepcion(self, db):
        db.add(Totales(FECHA=date(2026, 8, 5)))  # todo en blanco/None
        db.commit()

        servicio = TotalesDiariosService(db)
        detalle = servicio.detalle_dia(date(2026, 8, 5))
        assert detalle.precio_prom_cta_cte == Decimal("0")
        assert detalle.unid_prom_tipo_c == Decimal("0")

    def test_detalle_dia_sin_fila_da_none(self, db):
        servicio = TotalesDiariosService(db)
        assert servicio.detalle_dia(date(2026, 8, 5)) is None


# ---------------------------------------------------------------------------
# ArregloCtaCteService (CargaCC.frm, menú oculto "Arreglos")
# ---------------------------------------------------------------------------


class TestArregloCtaCteService:
    def test_grabar_crea_fila_nueva_con_defaults_y_resincroniza_deuda(self, db):
        db.add(Cliente(CODIGO=999999, NOMB="Cliente Ficticio", DEUDA=Decimal("0")))
        db.commit()

        servicio = ArregloCtaCteService(db)
        fila = servicio.grabar(
            clte=999999, fecha=date(2026, 8, 1), tipo=1, letra="A", prefijo=1, cpbte=555,
            imput1="0", imput2="0", imput3="0", debe=Decimal("1000"), impte=Decimal("1210"),
            fecvto=date(2026, 10, 1),
        )

        assert fila.IMPUT4 == "0 " and fila.IMPUT5 == "0 " and fila.IMPUT6 == "0 "
        assert fila.USUAR == "Arregl"
        assert fila.IMPTE == Decimal("1210")

        cliente = db.query(Cliente).filter(Cliente.CODIGO == 999999).first()
        assert cliente.DEUDA == Decimal("1210")  # TIPO=1 (Factura) suma al saldo deudor

    def test_grabar_sobre_fila_existente_no_toca_imput456(self, db):
        db.add(Cliente(CODIGO=999999, NOMB="Cliente Ficticio", DEUDA=Decimal("0")))
        db.add(
            Ctascte(
                CLTE=999999, FECHA=date(2026, 8, 1), TIPO=1, LETRA="A", PREFIJO=1, CPBTE=555,
                IMPUT1="0", IMPUT2="0", IMPUT3="0", IMPUT4="777", IMPUT5="0 ", IMPUT6="0 ",
                DEBE=Decimal("1000"), IMPTE=Decimal("1000"),
            )
        )
        db.commit()

        servicio = ArregloCtaCteService(db)
        fila = servicio.grabar(
            clte=999999, fecha=date(2026, 8, 1), tipo=1, letra="A", prefijo=1, cpbte=555,
            imput1="123", imput2="0", imput3="0", debe=Decimal("500"), impte=Decimal("1000"),
            fecvto=date(2026, 10, 1),
        )
        assert fila.IMPUT1 == "123"
        assert fila.IMPUT4 == "777"  # NO se toca en un Cambio, sólo en Alta

    def test_borrar_elimina_fila_y_resincroniza_deuda(self, db):
        db.add(Cliente(CODIGO=999999, NOMB="Cliente Ficticio", DEUDA=Decimal("1210")))
        db.add(
            Ctascte(
                CLTE=999999, FECHA=date(2026, 8, 1), TIPO=1, LETRA="A", PREFIJO=1, CPBTE=555,
                IMPTE=Decimal("1210"),
            )
        )
        db.commit()

        servicio = ArregloCtaCteService(db)
        assert servicio.borrar(999999, "A", 1, 1, 555) is True
        assert servicio.buscar(999999, "A", 1, 1, 555) is None

        cliente = db.query(Cliente).filter(Cliente.CODIGO == 999999).first()
        assert cliente.DEUDA == Decimal("0")

    def test_borrar_sin_fila_da_false(self, db):
        servicio = ArregloCtaCteService(db)
        assert servicio.borrar(999999, "A", 1, 1, 555) is False


# ---------------------------------------------------------------------------
# ArregloSubdiarioService (CargaFC.frm, menú oculto "Arreglos")
# ---------------------------------------------------------------------------


class TestArregloSubdiarioService:
    def _grabar_ejemplo(self, servicio, **overrides):
        datos = dict(
            tipo=1, letra="A", ptovta=1, cpbte=777, fecha=date(2026, 8, 1),
            clte=999999, nomb="Cliente Ficticio", pcia="B ", cuit="20999999990",
            grins=Decimal("1000"), ivains=Decimal("210"), ivanoins=Decimal("0"),
            exento=Decimal("0"), bon=Decimal("0"), porcib=Decimal("0"), totib=Decimal("0"),
            items=1, totcan=1, civa=3, vend=1, zona=1, cvta=1, moti="1",
        )
        datos.update(overrides)
        return servicio.grabar(**datos)

    def test_grabar_crea_fila_nueva_con_defaults(self, db):
        servicio = ArregloSubdiarioService(db)
        fila = self._grabar_ejemplo(servicio)

        assert fila.TIPO == "1"
        assert fila.TOTCOS == Decimal("0")
        assert fila.GRNOINS == Decimal("0")

    def test_grabar_usa_el_codigo_real_de_civa_no_el_listindex(self, db):
        """Bug real del legacy (Combo3.ListIndex 0-4 en vez de
        Combo3.ItemData 1-5) NO replicado — ver docstring del service."""
        servicio = ArregloSubdiarioService(db)
        fila = self._grabar_ejemplo(servicio, civa=4)  # 4 = Exento, código real
        assert fila.CIVA == "4"

    def test_grabar_sobre_fila_existente_actualiza_sin_resetear_defaults(self, db):
        servicio = ArregloSubdiarioService(db)
        self._grabar_ejemplo(servicio, totib=Decimal("50"))
        fila2 = self._grabar_ejemplo(servicio, grins=Decimal("2000"), totib=Decimal("999"))
        assert fila2.GRINS == Decimal("2000")
        assert fila2.TOTIB == Decimal("999")
        assert fila2.TOTCOS == Decimal("0")  # default de alta, no se vuelve a tocar en el cambio

    def test_borrar_elimina_fila(self, db):
        servicio = ArregloSubdiarioService(db)
        self._grabar_ejemplo(servicio)
        assert servicio.borrar("A", 1, 1, 777) is True
        assert servicio.buscar("A", 1, 1, 777) is None

    def test_borrar_sin_fila_da_false(self, db):
        servicio = ArregloSubdiarioService(db)
        assert servicio.borrar("A", 1, 1, 777) is False

    def test_grabar_sin_cliente_real_no_falla(self, db):
        """CLTE tipeado a mano puede no existir en `Clientes` (mismo
        editor sin validaciones) — el resync de DEUDA debe tolerarlo."""
        servicio = ArregloSubdiarioService(db)
        fila = self._grabar_ejemplo(servicio, clte=123456)  # cliente inexistente
        assert fila.CLTE == 123456


# ---------------------------------------------------------------------------
# ListadosService (Listados.frm)
# ---------------------------------------------------------------------------


class TestListadosService:
    def _clientes_de_prueba(self, db):
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Beta S.A.", PCIA="B ", ZONA=1, VEND=10),
                Cliente(CODIGO=2, NOMB="Alfa S.A.", PCIA="X", ZONA=1, VEND=10),
                Cliente(CODIGO=3, NOMB="Otra Zona", PCIA="B ", ZONA=2, VEND=20),
            ]
        )
        db.commit()

    def test_listado_clientes_filtro_uno(self, db):
        self._clientes_de_prueba(db)
        servicio = ListadosService(db)
        filas = servicio.listado_clientes(ListadosService.FILTRO_UNO, 1)
        assert [c.CODIGO for c in filas] == [1]

    def test_listado_clientes_filtro_zona(self, db):
        self._clientes_de_prueba(db)
        servicio = ListadosService(db)
        filas = servicio.listado_clientes(ListadosService.FILTRO_ZONA, 1)
        assert {c.CODIGO for c in filas} == {1, 2}

    def test_listado_clientes_filtro_todos_ordena_por_nombre(self, db):
        self._clientes_de_prueba(db)
        servicio = ListadosService(db)
        filas = servicio.listado_clientes(ListadosService.FILTRO_TODOS)
        assert [c.NOMB for c in filas] == ["Alfa S.A.", "Beta S.A.", "Otra Zona"]

    def test_listado_clientes_filtro_provincia(self, db):
        self._clientes_de_prueba(db)
        servicio = ListadosService(db)
        filas = servicio.listado_clientes(ListadosService.FILTRO_PROVINCIA, "B")
        assert {c.CODIGO for c in filas} == {1, 3}

    def test_lista_precios_por_seccion_y_todos(self, db):
        db.add_all(
            [
                Articulo(COD1="A", COD2="1", DESCRI="Uno", PREC=Decimal("100")),
                Articulo(COD1="B", COD2="2", DESCRI="Dos", PREC=Decimal("200")),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        assert len(servicio.lista_precios("A")) == 1
        assert len(servicio.lista_precios()) == 2

    def test_subdiario_ventas_filtra_rango_inclusive(self, db):
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 1), TIPO="1", CPBTE=1, PTOVTA=1),
                FcivaVta(FECHA=date(2026, 8, 31), TIPO="1", CPBTE=2, PTOVTA=1),
                FcivaVta(FECHA=date(2026, 9, 1), TIPO="1", CPBTE=3, PTOVTA=1),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        filas = servicio.subdiario_ventas(date(2026, 8, 1), date(2026, 8, 31))
        assert [f.CPBTE for f in filas] == [1, 2]

    def test_deuda_pendiente_agrupa_por_cliente_y_omite_sin_deuda(self, db):
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Con Deuda", ZONA=1),
                Cliente(CODIGO=2, NOMB="Sin Deuda", ZONA=1),
                Ctascte(CLTE=1, TIPO=1, DEBE=Decimal("500"), FECHA=date(2026, 8, 1), CPBTE=1),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        resultado = servicio.deuda_pendiente(ListadosService.FILTRO_ZONA, 1)
        assert len(resultado) == 1
        cliente, pendientes = resultado[0]
        assert cliente.CODIGO == 1
        assert len(pendientes) == 1

    def test_saldos_cta_cte_filtra_por_zona(self, db):
        self._clientes_de_prueba(db)
        db.add_all(
            [
                Ctascte(CLTE=1, TIPO=1, IMPTE=Decimal("100"), FECHA=date(2026, 8, 1), CPBTE=1),
                Ctascte(CLTE=3, TIPO=1, IMPTE=Decimal("999"), FECHA=date(2026, 8, 1), CPBTE=2),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        filas = servicio.saldos_cta_cte(ListadosService.FILTRO_ZONA, 1)
        assert {f["codigo"] for f in filas} == {1, 2}  # zona 1 = clientes 1 y 2 (aunque el 2 tenga saldo 0)
        saldo_cliente_1 = next(f for f in filas if f["codigo"] == 1)
        assert saldo_cliente_1["saldo"] == Decimal("100")

    def test_subdiario_cobranzas_filtra_tipo_4(self, db):
        db.add_all(
            [
                Ctascte(CLTE=1, TIPO=4, FECHA=date(2026, 8, 5), CPBTE=1, IMPTE=Decimal("500")),
                Ctascte(CLTE=1, TIPO=1, FECHA=date(2026, 8, 5), CPBTE=2, IMPTE=Decimal("500")),  # no es Recibo
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        filas = servicio.subdiario_cobranzas(date(2026, 8, 1), date(2026, 9, 1))
        assert len(filas) == 1
        assert filas[0].TIPO == 4

    def test_ingresos_brutos_resta_notas_de_credito(self, db):
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", PCIA="B ", GRINS=Decimal("1000"), IVAINS=Decimal("210"), CPBTE=1, PTOVTA=1),
                FcivaVta(FECHA=date(2026, 8, 6), TIPO="2", PCIA="B ", GRINS=Decimal("100"), IVAINS=Decimal("21"), CPBTE=2, PTOVTA=1),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        filas = servicio.ingresos_brutos(date(2026, 8, 1), date(2026, 9, 1))
        total_gravado = sum((f.gravado for f in filas), Decimal("0"))
        assert total_gravado == Decimal("900")  # 1000 - 100
        # Nombre de provincia decodificado
        assert any(f.pcia_nombre == "Bs. As." for f in filas)

    def test_subdiario_ventas_comisiones_nc_resta_y_exento_va_aparte(self, db):
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", CIVA="4", GRINS=Decimal("1000"), CPBTE=1, PTOVTA=1, LETRA="A"),  # Exento
                FcivaVta(FECHA=date(2026, 8, 6), TIPO="2", CIVA="1", GRINS=Decimal("100"), CPBTE=2, PTOVTA=1, LETRA="A"),  # NC
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        filas = servicio.subdiario_ventas_comisiones(date(2026, 8, 1), date(2026, 9, 1))
        fila_exenta = next(f for f in filas if f.civa_label == "Exto")
        assert fila_exenta.exento == Decimal("1000")
        assert fila_exenta.gravado == Decimal("0")
        fila_nc = next(f for f in filas if f.comprobante.endswith("00000002"))
        assert fila_nc.gravado == Decimal("-100")  # NC resta

    def test_subdiario_ventas_comisiones_filtra_por_vendedor(self, db):
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", VEND=10, GRINS=Decimal("100"), CPBTE=1, PTOVTA=1),
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", VEND=20, GRINS=Decimal("200"), CPBTE=2, PTOVTA=1),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        filas = servicio.subdiario_ventas_comisiones(date(2026, 8, 1), date(2026, 9, 1), vend=10)
        assert len(filas) == 1
        assert filas[0].gravado == Decimal("100")

    def test_percepciones_arba_filtra_solo_con_percepcion(self, db):
        _set_parametro(db, ivains="21")
        db.add_all(
            [
                FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", CIVA="1", GRINS=Decimal("1000"), TOTIB=Decimal("35"), CPBTE=1, PTOVTA=1, CUIT="20123456780"),
                FcivaVta(FECHA=date(2026, 8, 6), TIPO="1", CIVA="1", GRINS=Decimal("500"), TOTIB=Decimal("0"), CPBTE=2, PTOVTA=1),  # sin percepción
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        resultado = servicio.percepciones_arba(date(2026, 8, 1), date(2026, 9, 1))
        assert len(resultado.filas) == 1
        assert resultado.total_percepcion == Decimal("35")
        assert resultado.filas[0].cuit == "20-12345678-0"

    def test_percepciones_arba_nd_cheque_rechazado_recalcula_imponible(self, db):
        _set_parametro(db, ivains="21")
        db.add(
            FcivaVta(
                FECHA=date(2026, 8, 5), TIPO="3", MOTI=" 2", CIVA="1",
                GRINS=Decimal("0"), IVAINS=Decimal("210"), TOTIB=Decimal("10"), CPBTE=1, PTOVTA=1,
            )
        )
        db.commit()
        servicio = ListadosService(db)
        resultado = servicio.percepciones_arba(date(2026, 8, 1), date(2026, 9, 1))
        assert resultado.filas[0].imponible == Decimal("210") / Decimal("0.21")  # 1000

    def test_percepciones_arba_exento_no_es_base_imponible(self, db):
        _set_parametro(db, ivains="21")
        db.add(
            FcivaVta(FECHA=date(2026, 8, 5), TIPO="1", CIVA="4", GRINS=Decimal("1000"), TOTIB=Decimal("5"), CPBTE=1, PTOVTA=1)
        )
        db.commit()
        servicio = ListadosService(db)
        resultado = servicio.percepciones_arba(date(2026, 8, 1), date(2026, 9, 1))
        assert resultado.filas[0].imponible == Decimal("0")

    def test_comisiones_cobranzas_excluye_nd_rechazadas_del_total_pero_las_lista(self, db):
        _set_parametro(db, ivains="21")
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="Cliente Uno", VEND=10),
                Ctascte(CLTE=1, TIPO=4, FECHA=date(2026, 8, 5), CPBTE=1, IMPTE=Decimal("1210")),  # Recibo
                Ctascte(CLTE=1, TIPO=3, MOTI=" 2", FECHA=date(2026, 8, 6), CPBTE=2, IMPTE=Decimal("300")),  # ND Cheque Rechazado
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        resultado = servicio.comisiones_cobranzas(10, date(2026, 8, 1), date(2026, 9, 1))
        assert len(resultado.filas) == 2  # se listan ambas
        assert resultado.total_cobrado == Decimal("1210")  # la ND no suma
        assert resultado.total_nd_rechazadas == 1
        assert resultado.total_nd_rechazadas_importe == Decimal("300")
        assert resultado.neto == (Decimal("1210") - Decimal("300")) / Decimal("1.21")

    def test_comisiones_cobranzas_solo_clientes_del_vendedor(self, db):
        _set_parametro(db, ivains="21")
        db.add_all(
            [
                Cliente(CODIGO=1, NOMB="De Otro Vendedor", VEND=99),
                Ctascte(CLTE=1, TIPO=4, FECHA=date(2026, 8, 5), CPBTE=1, IMPTE=Decimal("500")),
            ]
        )
        db.commit()
        servicio = ListadosService(db)
        resultado = servicio.comisiones_cobranzas(10, date(2026, 8, 1), date(2026, 9, 1))
        assert resultado.filas == []
        assert resultado.total_cobrado == Decimal("0")
