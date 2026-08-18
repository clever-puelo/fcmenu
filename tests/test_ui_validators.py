"""Tests de migration/ui/validators.py.

`cuit_valido` es una traducción 1:1 de FCMENU.bas Function Cuit() — se
verifica contra CUITs reales conocidos y contra los casos borde propios
del algoritmo (verificador 11->0, 10->9).
"""

from __future__ import annotations

from migration.ui.validators import cuit_valido


def test_cuit_valido_responsable_inscripto_real():
    # CUIT de AFIP (organismo) usado como ejemplo de referencia pública.
    assert cuit_valido("30500010912") is True


def test_cuit_invalido_por_digito_verificador_incorrecto():
    assert cuit_valido("30500010911") is False


def test_cuit_invalido_por_longitud():
    assert cuit_valido("123") is False
    assert cuit_valido("") is False
    assert cuit_valido("300500010912") is False


def test_cuit_invalido_por_caracteres_no_numericos():
    assert cuit_valido("3050001091A") is False
    assert cuit_valido("30-50001091") is False


def test_cuit_caso_borde_verificador_11_pasa_a_0():
    # Construido (por búsqueda exhaustiva) para que el cálculo módulo 11 dé
    # resto 0 -> verificador 11 -> se corrige a 0 (FCMENU.bas:526-528).
    assert cuit_valido("20123456050") is True


def test_cuit_caso_borde_verificador_10_pasa_a_9():
    # Construido para que el resto módulo 11 dé 1 -> verificador 10 -> se
    # corrige a 9 (FCMENU.bas:529-531).
    assert cuit_valido("20123456009") is True
