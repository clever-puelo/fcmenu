"""Tests de migration/decimals.py — parseo/formato de montos es-AR.

`migration/ui/decimals.py` es sólo un re-export (ver su docstring) —
se testea acá, el módulo real.
"""

from __future__ import annotations

from decimal import Decimal

import pytest

from migration.decimals import format_decimal, parse_decimal


class TestParseDecimal:
    def test_texto_vacio_es_cero(self):
        assert parse_decimal("") == Decimal("0")
        assert parse_decimal("   ") == Decimal("0")

    def test_punto_decimal_simple(self):
        assert parse_decimal("1234.56") == Decimal("1234.56")

    def test_coma_decimal_es_ar(self):
        assert parse_decimal("1234,56") == Decimal("1234.56")

    def test_punto_de_miles_y_coma_decimal_es_ar(self):
        assert parse_decimal("1.234,56") == Decimal("1234.56")

    def test_texto_invalido_es_cero(self):
        assert parse_decimal("abc") == Decimal("0")


class TestFormatDecimal:
    def test_formatea_con_punto_de_miles_y_coma_decimal(self):
        assert format_decimal(Decimal("681240.23")) == "681.240,23"

    def test_sin_miles(self):
        assert format_decimal(Decimal("50")) == "50,00"

    def test_cero(self):
        assert format_decimal(Decimal("0")) == "0,00"

    def test_negativo(self):
        assert format_decimal(Decimal("-500.25")) == "-500,25"

    def test_varios_miles(self):
        assert format_decimal(Decimal("1000000.5")) == "1.000.000,50"


class TestRoundTrip:
    """Regresión del bug real encontrado probando el Recibo con datos
    reales (2026-08-15, cliente con Debe $681.240,23): escribir con
    `f"{valor:,.2f}"` (formato US) y releer con `parse_decimal()`
    corrompía cualquier importe >= 1000 en silencio."""

    @pytest.mark.parametrize(
        "valor",
        [
            Decimal("0"),
            Decimal("0.5"),
            Decimal("999.99"),
            Decimal("1000"),
            Decimal("1234.56"),
            Decimal("681240.23"),
            Decimal("1000000.50"),
            Decimal("-500.25"),
        ],
    )
    def test_format_y_parse_son_simetricos(self, valor):
        assert parse_decimal(format_decimal(valor)) == valor

    def test_formato_us_directo_NO_es_lo_que_hay_que_usar(self):
        """Documenta el bug: `f"{valor:,.2f}"` (nunca usar esto para un
        texto que se vuelve a parsear) rompe con parse_decimal a partir
        de 1000."""
        valor = Decimal("681240.23")
        texto_us_incorrecto = f"{valor:,.2f}"
        assert texto_us_incorrecto == "681,240.23"
        assert parse_decimal(texto_us_incorrecto) != valor  # corrompido
        assert parse_decimal(texto_us_incorrecto) == Decimal("681.24023")  # lo que pasaba en la práctica
