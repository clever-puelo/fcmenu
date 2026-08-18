"""Tests de migration/fechas.py."""

from __future__ import annotations

from datetime import date

from migration.fechas import formatear_fecha_larga


def test_formatea_dia_mes_anio_en_espanol():
    # 2026-08-15 es un sábado.
    assert formatear_fecha_larga(date(2026, 8, 15)) == "sábado 15 ago 2026"


def test_formatea_lunes():
    # 2026-08-17 es un lunes.
    assert formatear_fecha_larga(date(2026, 8, 17)) == "lunes 17 ago 2026"


def test_meses_abreviados_correctos():
    assert formatear_fecha_larga(date(2026, 1, 1)) == "jueves 01 ene 2026"
    assert formatear_fecha_larga(date(2026, 12, 31)) == "jueves 31 dic 2026"
