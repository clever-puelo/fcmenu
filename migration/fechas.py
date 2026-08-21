"""Helper de formato de fecha "larga" en español ("dddd dd MMM yyyy",
ej. "lunes 15 ago 2026") — pedido del usuario (2026-08-15) para la
grilla de Pagos del Recibo.

Se arma con tablas fijas en vez de `locale`/`strftime("%A")`: el sistema
no depende de que el locale es-AR esté instalado en el SO donde corra
(riesgo real en un entorno de escritorio Windows con locale variable),
mismo criterio que el resto de la app (etiquetas en español hardcodeadas,
no traducciones dependientes del entorno).
"""

from __future__ import annotations

from datetime import date

_DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
_MESES_ABREV_ES = [
    "ene", "feb", "mar", "abr", "may", "jun",
    "jul", "ago", "sep", "oct", "nov", "dic",
]


def formatear_fecha_larga(fecha: date) -> str:
    """"dddd dd MMM yyyy" — mismo formato que usaban los `DTPicker` del
    legacy (`DetPago.frm`: `CustomFormat = "ddd dd MMM yyyy"`), acá con
    el nombre del día completo en vez de abreviado, tal como lo pidió
    el usuario."""
    return f"{_DIAS_ES[fecha.weekday()]} {fecha.day:02d} {_MESES_ABREV_ES[fecha.month - 1]} {fecha.year}"


_DIAS_ABREV_ES = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


def formatear_fecha_corta(fecha: date) -> str:
    """"ddd dd MMM yyyy" con día Y mes abreviados (ej. "Dom 16 Ago 2026") —
    el mismo `CustomFormat` literal de los `DTPicker` del legacy (ver
    `formatear_fecha_larga`), acá sin el nombre del día completo porque
    se usa en la cabecera de `MainMenuWindow` (`main_menu_window.py`),
    donde no entra."""
    dia = _DIAS_ABREV_ES[fecha.weekday()]
    mes = _MESES_ABREV_ES[fecha.month - 1].capitalize()
    return f"{dia} {fecha.day:02d} {mes} {fecha.year}"


def formatear_fecha_dd_mmm(fecha: date) -> str:
    """"dd mmm yyyy" en minúsculas, sin día de semana (ej. "25 jun
    2026") — réplica de `Format(fecha, "dd mmm yyyy")` bajo locale
    es-AR, usado en las tablas de comprobantes cancelados/aplicados de
    Recibo/Nota de Crédito Interna."""
    return f"{fecha.day:02d} {_MESES_ABREV_ES[fecha.month - 1]} {fecha.year}"


def formatear_fecha_impresion(fecha: date) -> str:
    """"ddd dd MMM yyyy" TODO en minúsculas (ej. "jue 20 ago 2026") —
    réplica exacta de `Format(Date, "ddd dd MMM yyyy")` (`EmiFact.frm`/
    `EmiRec.frm`/`NCInterna.frm`, todos los comprobantes impresos) bajo
    locale es-AR real: VB6 en ese locale imprime día/mes abreviados en
    minúscula, sin punto. Distinta de `formatear_fecha_corta` (que
    capitaliza, para la cabecera de la app) — acá se busca fidelidad
    con el papel impreso real, no legibilidad de pantalla."""
    dia = _DIAS_ES[fecha.weekday()][:3]
    return f"{dia} {fecha.day:02d} {_MESES_ABREV_ES[fecha.month - 1]} {fecha.year}"
