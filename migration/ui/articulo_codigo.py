"""Generador del código compuesto de artículo (COD2) — lógica pura, sin Qt.

Migración de la parte más delicada de `AbmArt.frm`: el código de un
artículo no se tipea libre. Se arma concatenando 1 o 2 "segmentos"
numéricos cuyo ancho (dígitos enteros / decimales) sale de la tabla de
Unidades de Medida (`Fctabla1` CTAB='UM'), según cuál Unidad de Medida
tenga configurada la Sección elegida (`Fctabla1` CTAB='SC', columnas
ALF1/ALF2 → apuntan a códigos de UM).

Confirmado contra datos reales de `fcmenu_dev` (2026-08-14): de las 63
Secciones reales, ~48 usan la unidad "NRO" (número libre, sin padding —
caso `Lgo1=5/Lgo2=0` de Abmclte.frm) y 15 usan una unidad física real
(mangueras, gomas, cintas: MM, PULG, MTRS, etc.) que si necesita padding
con ceros a la izquierda. 4 Secciones (CT, GPN, GPS, PL) tienen las DOS
dimensiones (ALF1 y ALF2).

**Bug real encontrado en el legacy y NO replicado**: en
`AbmArt.frm Sub Grabacion()` (líneas 1007-1025), al paddear el segundo
segmento el código pisa la variable equivocada
(`Codi1 = String(...) & Codi2` en vez de `Codi2 = ...`), lo que
corrompería el código final si el segundo segmento necesitara padding
(no se manifestó nunca en los datos reales porque la única Unidad usada
como segundo segmento, TELAS, tiene ancho 1 y los operadores siempre
tipean ese único dígito). Acá se implementa la lógica correcta.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import Optional

from migration.models import Fctabla1
from migration.repository import FctablasRepository

CTAB_SECCIONES = "SC"
CTAB_UNIDADES = "UM"
COD_UNIDAD_LIBRE = "NRO"


@dataclass(frozen=True)
class SegmentoCodigo:
    """Un segmento del código compuesto (una Unidad de Medida configurada
    en la Sección)."""

    cod_unidad: str
    etiqueta: str
    numsd1: int  # dígitos de la parte entera
    numsd2: int  # dígitos de la parte decimal
    libre: bool  # True para "NRO": número libre, sin padding ni decimales

    @property
    def ancho_total(self) -> int:
        return self.numsd1 + self.numsd2


def resolver_segmentos(fctablas: FctablasRepository, cod_seccion: str) -> list[SegmentoCodigo]:
    """Devuelve los 1-2 segmentos que arman el código de una Sección.

    Réplica de Abmclte... digo, AbmArt.frm Sub BuscaSeccion() (líneas
    1121-1171): lee ALF1/ALF2 de la Sección (CTAB='SC') y por cada uno
    no vacío busca su definición en CTAB='UM'. ALF3 existe en el modelo
    y en el .frm pero nunca se usa en BuscaSeccion() — código muerto real,
    se ignora acá también (no se inventa un tercer segmento).
    """
    seccion = fctablas.by_ctab_cod(CTAB_SECCIONES, cod_seccion)
    if seccion is None:
        return []

    segmentos: list[SegmentoCodigo] = []
    for alf in (seccion.ALF1, seccion.ALF2):
        cod_unidad = (alf or "").strip()
        if not cod_unidad:
            continue
        unidad = fctablas.by_ctab_cod(CTAB_UNIDADES, cod_unidad)
        if unidad is None:
            continue
        segmentos.append(
            SegmentoCodigo(
                cod_unidad=cod_unidad,
                etiqueta=(unidad.DESCRI or "Código").strip() or "Código",
                numsd1=unidad.NUMSD1 or 0,
                numsd2=unidad.NUMSD2 or 0,
                libre=(cod_unidad == COD_UNIDAD_LIBRE),
            )
        )
    return segmentos


def formatear_segmento(valor: Decimal, segmento: SegmentoCodigo) -> str:
    """Convierte el valor tipeado por el operador en el bloque de dígitos
    con padding que le corresponde dentro del código final."""
    if segmento.libre:
        return str(int(valor))

    factor = Decimal(10) ** segmento.numsd2
    escalado = int((valor * factor).to_integral_value(rounding=ROUND_HALF_UP))
    return str(max(escalado, 0)).zfill(segmento.ancho_total)


def armar_codigo(segmentos: list[SegmentoCodigo], valores: list[Decimal]) -> str:
    """Concatena los segmentos ya formateados — este es el `COD2` final."""
    if len(segmentos) != len(valores):
        raise ValueError("La cantidad de valores no coincide con la de segmentos")
    return "".join(formatear_segmento(v, s) for v, s in zip(valores, segmentos))


def separar_codigo(cod2: str, segmentos: list[SegmentoCodigo]) -> list[Decimal]:
    """Inversa de `armar_codigo`: separa un COD2 ya guardado en sus valores
    por segmento, para poblar los campos al editar un artículo existente
    (AbmArt.frm Form_Load líneas 710-717: `Mid(RESTO, 1, Lgo1)` /
    `Mid(RESTO, Lgo1+1, Lgo2)`)."""
    cod2 = (cod2 or "").strip()
    valores: list[Decimal] = []
    pos = 0
    for segmento in segmentos:
        if segmento.libre:
            # Único segmento libre (sección "NRO"): el código completo es
            # el número, sin ancho fijo ni padding.
            trozo = cod2[pos:] if len(segmentos) == 1 else cod2[pos : pos + segmento.ancho_total]
            pos += len(trozo)
            try:
                valores.append(Decimal(int(trozo)) if trozo.strip() else Decimal(0))
            except (ValueError, ArithmeticError):
                valores.append(Decimal(0))
            continue
        trozo = cod2[pos : pos + segmento.ancho_total].rjust(segmento.ancho_total, "0")
        pos += segmento.ancho_total
        try:
            entero = int(trozo) if trozo.strip() else 0
        except ValueError:
            entero = 0
        valores.append(Decimal(entero) / (Decimal(10) ** segmento.numsd2))
    return valores


def descripcion_automatica(cod_seccion: str, cod2: str) -> str:
    """Descripción por defecto cuando el operador deja el campo en blanco
    (AbmArt.frm Grabacion(): `RgART!Descri = Combo2.Text & " " & Codi2`)."""
    return f"{cod_seccion.strip()} {cod2.strip()}"
