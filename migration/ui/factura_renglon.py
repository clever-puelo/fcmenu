"""Motor del renglón de Detalle de Factura — lógica pura, sin Qt.

Migración del mecanismo central de `DetFact.frm`: `Fctabla1(UM).NUMSD3`
("Posición", ver `tablas_window.py` y `Acttabla.frm`) se usa en el legacy
como índice de columna dinámica de la grilla FlexGrid (`FG1.Col =
Posi(i)`) — la Unidad de Medida configurada en cada ALF1-7 de la Sección
decide en qué "slot" físico (1, 2, 3, 4 o 6 — únicos valores reales del
enum, confirmado contra las 63 Secciones de `fcmenu_dev`) se tipea ese
segmento de código (ALF1-3) o esa cantidad/multiplicador de precio
(ALF4-7). Acá no hay grilla VB6, así que el "slot" no importa para el
layout — lo que sí hay que preservar es el ORDEN y la fórmula: el
Precio Unitario final depende de multiplicar los valores de ALF4-7 (los
presentes) en ese orden y dividir por la Unidad de Facturación (ALF7,
siempre obligatoria).

Reutiliza `formatear_segmento`/`armar_codigo` de `articulo_codigo.py`
para el armado del código compuesto (ALF1-3, mismo padding ya validado
en la fase de Artículos) — `SegmentoRenglon` es compatible por duck
typing (misma interfaz: `.libre`, `.numsd2`, `.ancho_total`).

Fuente: `DetFact.frm` — `Sub BuscaSeccion()` (líneas 929-1119, resuelve
los segmentos de una Sección), `Sub FG1_AfterEdit()` bloque `AlPrecio`
(líneas 599-621, calcula Precio Unitario/Importe).
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Optional

from migration.repository import FctablasRepository
from migration.ui.articulo_codigo import (
    COD_UNIDAD_LIBRE,
    CTAB_UNIDADES,
    armar_codigo,
    formatear_segmento,  # noqa: F401 (reexportado para quien arme un segmento a mano)
)

CTAB_SECCIONES = "SC"

# Código especial de Sección para un ítem de texto libre (sin Artículo real).
# DetFact.frm Sub BuscaSeccion(): `If Correa = "**" Then ...` — función real
# del negocio (cargar gastos varios directo en la factura), confirmada con
# el usuario, no se saca.
SECCION_ITEM_LIBRE = "**"


@dataclass(frozen=True)
class SegmentoRenglon:
    """Un segmento ALF (1 a 7) de una Sección ya resuelto contra `Fctabla1`
    CTAB='UM'. Los índices 1-3 son segmentos de CÓDIGO (arman el COD2 del
    Artículo); 4-7 son CANTIDAD/multiplicador de precio (4-6 opcionales,
    7 siempre presente — es la Unidad de Facturación)."""

    alf_index: int  # 1..7
    cod_unidad: str
    etiqueta: str
    numsd1: int  # dígitos de la parte entera
    numsd2: int  # dígitos de la parte decimal
    libre: bool  # True para "NRO": número libre, sin padding ni decimales
    posicion: int  # Fctabla1(UM).NUMSD3 — la columna dinámica del legacy

    @property
    def ancho_total(self) -> int:
        return self.numsd1 + self.numsd2

    @property
    def es_codigo(self) -> bool:
        return self.alf_index <= 3

    @property
    def es_facturacion(self) -> bool:
        return self.alf_index == 7


@dataclass(frozen=True)
class SeccionRenglon:
    """Resultado de resolver una Sección para armar un renglón de factura."""

    cod_seccion: str
    descripcion: str
    usar_descripcion_seccion: bool  # MCA1: True -> arma la descripción con los segmentos, False -> usa la del Artículo
    precio_en_dolares: bool  # MCA2
    segmentos_codigo: list[SegmentoRenglon]  # ALF1-3 presentes, en orden real
    segmentos_cantidad: list[SegmentoRenglon]  # ALF4-7 presentes, en orden real (ALF7 siempre última)
    item_libre: bool = False  # True sólo para la Sección especial "**"

    @property
    def segmento_facturacion(self) -> Optional[SegmentoRenglon]:
        """La Unidad de Facturación (ALF7) — la cantidad real vendida."""
        for segmento in self.segmentos_cantidad:
            if segmento.es_facturacion:
                return segmento
        return None


class SeccionInexistenteError(ValueError):
    """La Sección tipeada no existe en Fctabla1 (CTAB='SC').

    Réplica del MsgBox "La Sección ... No Existe en la Tabla"
    (DetFact.frm:964)."""


class SeccionSinUnidadFacturacionError(ValueError):
    """La Sección no tiene ALF7 (Unidad de Facturación) configurado.

    Réplica del MsgBox "No Tiene cargada la UNIDAD DE FACTURACIÓN"
    (DetFact.frm:973-975) — sin esto no se puede calcular el Precio
    Unitario (`Importe / Cantidad_ALF7`)."""


def _resolver_alf(fctablas: FctablasRepository, alf_index: int, alf_cod: Optional[str]) -> Optional[SegmentoRenglon]:
    cod_unidad = (alf_cod or "").strip()
    if not cod_unidad:
        return None
    unidad = fctablas.by_ctab_cod(CTAB_UNIDADES, cod_unidad)
    if unidad is None:
        return None
    return SegmentoRenglon(
        alf_index=alf_index,
        cod_unidad=cod_unidad,
        etiqueta=(unidad.DESCRI or "").strip() or "Código",
        numsd1=unidad.NUMSD1 or 0,
        numsd2=unidad.NUMSD2 or 0,
        libre=(cod_unidad == COD_UNIDAD_LIBRE),
        posicion=unidad.NUMSD3 or 0,
    )


def resolver_seccion_renglon(fctablas: FctablasRepository, cod_seccion: str) -> Optional[SeccionRenglon]:
    """Réplica de `DetFact.frm Sub BuscaSeccion()`.

    Devuelve `None` si la Sección no existe (a diferencia del legacy no
    se levanta excepción para este caso — el llamador decide qué hacer,
    ej. mostrar "Sección no encontrada" o disparar el selector de
    Artículo). Levanta `SeccionSinUnidadFacturacionError` si ALF7 está
    vacío, igual que el legacy.

    Caso especial `cod_seccion == "**"`: ítem de texto libre (sin
    Artículo real) — ver `SECCION_ITEM_LIBRE`. Réplica de
    `BuscaSeccion()` líneas 939-956: sin segmentos de código, una sola
    cantidad (Unidades) que multiplicada por el Precio da el Importe.
    """
    cod_seccion = cod_seccion.strip()

    if cod_seccion == SECCION_ITEM_LIBRE:
        return SeccionRenglon(
            cod_seccion=SECCION_ITEM_LIBRE,
            descripcion="",
            usar_descripcion_seccion=False,
            precio_en_dolares=False,
            segmentos_codigo=[],
            segmentos_cantidad=[
                SegmentoRenglon(
                    alf_index=7,
                    cod_unidad="UNID",
                    etiqueta="Cantidad",
                    numsd1=5,
                    numsd2=0,
                    libre=True,
                    posicion=6,
                )
            ],
            item_libre=True,
        )

    seccion = fctablas.by_ctab_cod(CTAB_SECCIONES, cod_seccion)
    if seccion is None:
        raise SeccionInexistenteError(cod_seccion)

    alf7 = (seccion.ALF7 or "").strip()
    if not alf7:
        raise SeccionSinUnidadFacturacionError(cod_seccion)

    segmentos_codigo = [
        segmento
        for segmento in (
            _resolver_alf(fctablas, idx, alf)
            for idx, alf in enumerate((seccion.ALF1, seccion.ALF2, seccion.ALF3), start=1)
        )
        if segmento is not None
    ]

    segmentos_cantidad = [
        segmento
        for segmento in (
            _resolver_alf(fctablas, idx, alf)
            for idx, alf in enumerate((seccion.ALF4, seccion.ALF5, seccion.ALF6, seccion.ALF7), start=4)
        )
        if segmento is not None
    ]

    return SeccionRenglon(
        cod_seccion=cod_seccion,
        descripcion=(seccion.DESCRI or "").strip(),
        usar_descripcion_seccion=bool(seccion.MCA1),
        precio_en_dolares=bool(seccion.MCA2),
        segmentos_codigo=segmentos_codigo,
        segmentos_cantidad=segmentos_cantidad,
    )


# Posiciones reales del enum de "Posición" (Fctabla1(UM).NUMSD3, ver
# tablas_window.py): 1=Nro/Pulg, 2=Mtr/Kg, 3=MM, 4=Telas, 6=Unidades
# (no existe la opción "5"). Se usa en las consultas de Ventas
# (ESTADIST/VTAXART) para mostrar sólo las columnas de Unidad de Medida
# que corresponden a la Sección elegida (feedback del usuario,
# 2026-08-16: "si es 'A' sólo debe estar 'Número', si es 'PL' 'MM' y
# 'Telas'... lo que indique la tabla de secciones").
POSICIONES_UM = (1, 2, 3, 4, 6)


def posiciones_activas_seccion(fctablas: FctablasRepository, cod_seccion: str) -> set[int]:
    """Posiciones que una Sección usa de verdad en su configuración real
    de ALF1-7 (unión de `segmentos_codigo` + `segmentos_cantidad`).
    Devuelve TODAS las posiciones como fallback seguro si la Sección no
    existe o no tiene Unidad de Facturación (`ALF7`) configurada — acá
    sólo se usa para decidir qué columnas mostrar, no para validar nada,
    así que un fallback permisivo es preferible a romper la consulta."""
    try:
        seccion = resolver_seccion_renglon(fctablas, cod_seccion)
    except (SeccionInexistenteError, SeccionSinUnidadFacturacionError):
        return set(POSICIONES_UM)
    posiciones = {s.posicion for s in seccion.segmentos_codigo} | {s.posicion for s in seccion.segmentos_cantidad}
    return posiciones or set(POSICIONES_UM)


def armar_codigo_renglon(seccion: SeccionRenglon, valores: list[Decimal]) -> str:
    """Arma el COD2 compuesto (RESTO en el legacy) a partir de los
    segmentos ALF1-3, reutilizando `armar_codigo()` de
    `articulo_codigo.py` (mismo padding). Vacío si la Sección no tiene
    segmentos de código (ítem libre, o Sección con sólo unidad "NRO"
    manejada aparte por el llamador)."""
    if not seccion.segmentos_codigo:
        return ""
    return armar_codigo(seccion.segmentos_codigo, valores)


def resolver_precio_articulo(
    precio_lista: Decimal, seccion_precio_en_dolares: bool, factura_en_dolares: bool, cotizacion: Decimal
) -> Decimal:
    """Conversión $/USD del precio de lista antes de aplicar los
    multiplicadores de cantidad — réplica exacta de `DetFact.frm Sub
    VerSiBusca` (líneas 841-853):

        Si la factura es en Dólares (`EnDolares`):
            Si la Sección tiene precio en USD (`MCA2`): precio = Precio (ya en USD)
            Si no:                                      precio = Precio / Cotización
        Si la factura es en Pesos:
            Si la Sección tiene precio en USD (`MCA2`): precio = Precio × Cotización
            Si no:                                      precio = Precio (ya en pesos)

    `precio_lista` es `Articulo.PREC` tal cual, sin convertir. El
    resultado es el `precio_base` que recibe `calcular_precio_e_importe()`.
    """
    precio_lista = Decimal(precio_lista)
    cotizacion = Decimal(cotizacion)
    if factura_en_dolares:
        return precio_lista if seccion_precio_en_dolares else precio_lista / cotizacion
    return precio_lista * cotizacion if seccion_precio_en_dolares else precio_lista


def calcular_importe(
    precio_unitario: Decimal,
    valores_cantidad: dict[int, Decimal],
    seccion: SeccionRenglon,
) -> Decimal:
    """Importe = Precio Unitario × ALF4 × ALF5 × ALF6 × ALF7 (los que
    existan) — a diferencia del mecanismo anterior (`calcular_precio_e_
    importe()`/`calcular_precio_preview()`, réplica fiel de `DetFact.frm
    FG1_AfterEdit` bloque `AlPrecio`, líneas 599-621, que excluía ALF7
    del "Precio Unitario" mostrado y lo iba RE-DERIVANDO de
    Importe/Cantidad a medida que se tipeaban los segmentos ALF4-6),
    acá el "Precio Unitario" NUNCA se deriva ni se toca solo: es
    siempre el precio de lista del Artículo o el que tipeó el operador
    a mano — decisión del usuario (2026-08-21), reportado como bug real
    ("en la PL, cuando carga las Pulgadas, [el Precio Unitario] se
    actualiza — no lo debe hacer"). El Importe sale de multiplicar ESE
    precio (`precio_unitario`, tal cual está en pantalla) por TODAS las
    cantidades, sin excluir la Unidad de Facturación.

    Levanta `ValueError` si la cantidad de la Unidad de Facturación es
    0 (no tendría sentido facturar cantidad 0)."""
    segmento_facturacion = seccion.segmento_facturacion
    if segmento_facturacion is None:
        raise SeccionSinUnidadFacturacionError(seccion.cod_seccion)

    cantidad_facturacion = Decimal(valores_cantidad.get(segmento_facturacion.alf_index, Decimal(0)))
    if cantidad_facturacion == 0:
        raise ValueError("La cantidad de la Unidad de Facturación no puede ser 0")

    importe = Decimal(precio_unitario)
    for segmento in seccion.segmentos_cantidad:
        importe *= Decimal(valores_cantidad.get(segmento.alf_index, Decimal(0)))
    return importe


def descripcion_con_segmentos(seccion: SeccionRenglon, descripcion_base: str, valores_codigo: list[Decimal]) -> str:
    """Arma la descripción final de un renglón cuando la Sección tiene
    `MCA1` activo ("Usar para Facturar" / usar descripción de Sección) —
    réplica de `EmiFact.frm` líneas 998-1032: concatena la descripción
    base con cada segmento de código que tenga valor > 0, en formato
    "<valor> <etiqueta_unidad>".

    Si `usar_descripcion_seccion` es False, el llamador debe usar
    directamente la descripción del Artículo (no ésta)."""
    partes = [descripcion_base]
    for segmento, valor in zip(seccion.segmentos_codigo, valores_codigo):
        if valor and valor > 0:
            partes.append(f"{valor} {segmento.etiqueta}")
    return " ".join(p for p in partes if p)
