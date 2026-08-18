"""Helpers de parseo/formato de montos — compartidos por la capa de
servicios y la capa de UI (por eso vive en `migration/`, no en
`migration/ui/`: `services.py` también los necesita, ej. para armar
mensajes de error, y `services` no debe depender de `ui`).

Nunca se usa `float` para dinero (regla del proyecto, ver CLAUDE.md), y
el operador tipea con coma decimal (cultura es-AR, igual que el legacy
VB6) — `parse_decimal()`/`format_decimal()` son la ida y vuelta
simétrica de esa convención.
"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation


def parse_decimal(texto: str) -> Decimal:
    """Convierte texto ingresado por el operador (coma o punto decimal) a Decimal.

    Texto vacío -> Decimal("0"). Nunca pasa por float en el camino.
    """
    texto = texto.strip().replace(".", "").replace(",", ".") if "," in texto else texto.strip()
    if not texto:
        return Decimal("0")
    try:
        return Decimal(texto)
    except InvalidOperation:
        return Decimal("0")


def format_decimal(valor: Decimal, decimales: int = 2) -> str:
    """Formatea un Decimal en convención es-AR (punto de miles, coma
    decimal — "1.234,56") — la contraparte simétrica de `parse_decimal()`.

    **Bug real encontrado con datos reales probando el Recibo (2026-08-15,
    cliente con Debe $681.240,23)**: escribir un monto con `f"{valor:,.2f}"`
    (formato US: coma de miles, punto decimal — "681,240.23") y después
    releerlo con `parse_decimal()` corrompe el valor en silencio para
    cualquier importe >= 1000, porque `parse_decimal()` siempre interpreta
    la coma como separador DECIMAL (convención del operador, que jamás
    tipea en formato US) — "681,240.23" terminaba parseado como 681.24,
    no 681240.23. Usar SIEMPRE esta función (nunca `f"{valor:,.2f}"`
    directo) para cualquier celda/campo que pueda volver a pasar por
    `parse_decimal()` — no sólo para mostrar, también para armar el texto
    que el propio código lee de vuelta.
    """
    texto_us = f"{Decimal(valor):,.{decimales}f}"  # "1,234.56"
    return texto_us.replace(",", "\x00").replace(".", ",").replace("\x00", ".")  # "1.234,56"
