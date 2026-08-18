"""Tests de migration/provincias.py."""

from __future__ import annotations

from migration.provincias import nombre_provincia


def test_nombre_provincia_decodifica_codigo_con_padding_real_de_access():
    assert nombre_provincia("V ") == "T.del Fgo"  # mismo caso citado en FacturaService.letra_comprobante
    assert nombre_provincia("B ") == "Bs. As."
    assert nombre_provincia("X") == "Córdoba"


def test_nombre_provincia_es_case_insensitive():
    assert nombre_provincia("v") == "T.del Fgo"


def test_nombre_provincia_vacio_o_desconocido_da_string_vacio():
    assert nombre_provincia(None) == ""
    assert nombre_provincia("") == ""
    assert nombre_provincia("  ") == ""
    assert nombre_provincia("I") == ""  # hueco real del array legacy (FCMENU.bas:359)
    assert nombre_provincia("9") == ""  # código inexistente
