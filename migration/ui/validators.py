"""Validaciones portadas 1:1 desde FCMENU.bas.

`cuit_valido` es la traducción exacta del algoritmo de dígito verificador
de `Function Cuit(strCuit As String) As Boolean` (FCMENU.bas:494-537):
multiplicadores 5,4,3,2,7,6,5,4,3,2 sobre los primeros 10 dígitos, resto
módulo 11, con los casos especiales 11->0 y 10->9.
"""

from __future__ import annotations

_MULTIPLICADORES = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)


def cuit_valido(cuit: str) -> bool:
    """Valida el dígito verificador de un CUIT/CUIL de 11 dígitos.

    Réplica de FCMENU.bas Function Cuit(). Devuelve False si `cuit` no
    tiene exactamente 11 dígitos numéricos (igual que el legacy, que
    también rechaza cadenas vacías o de otra longitud).
    """
    if len(cuit) != 11 or not cuit.isdigit():
        return False

    digitos = [int(c) for c in cuit]
    calculo = sum(d * m for d, m in zip(digitos, _MULTIPLICADORES))
    verificador = 11 - (calculo % 11)

    if verificador == 11:
        verificador = 0
    elif verificador == 10:
        verificador = 9

    return verificador == digitos[10]


def email_valido(email: str) -> bool:
    """Chequeo básico, no estricto (pedido del usuario, 2026-08-18):
    tiene que tener un "@" con algo antes y algo después. No es un
    regex de RFC completo a propósito — sólo agarra el error de tipeo
    típico (usuario sin @, dominio vacío)."""
    if "@" not in email:
        return False
    usuario, _, dominio = email.partition("@")
    return bool(usuario) and bool(dominio) and "@" not in dominio
