"""Tests de migration/ui/factura_renglon.py — el motor de renglón del
Detalle de Factura (DetFact.frm).

El caso "GPN" (Goma Plancha Nat., ALF1=MM1/ALF2=TELAS/ALF7=KG) reusa los
mismos datos reales que tests/test_articulo_codigo.py (COD2 "0300" para
3mm/0T), extendido acá con la Unidad de Facturación (ALF7) que AbmArt no
necesita pero DetFact sí.
"""

from __future__ import annotations

from decimal import Decimal

import pytest

from migration.models import Fctabla1
from migration.repository import RepositoryFactory
from migration.ui.factura_renglon import (
    SeccionInexistenteError,
    SeccionSinUnidadFacturacionError,
    armar_codigo_renglon,
    calcular_precio_e_importe,
    calcular_precio_preview,
    descripcion_con_segmentos,
    posiciones_activas_seccion,
    resolver_precio_articulo,
    resolver_seccion_renglon,
)


def _cargar_secciones_y_unidades(db):
    db.add_all(
        [
            # Secciones
            Fctabla1(
                CTAB="SC   ", COD="GPN  ", DESCRI="GOMA PLANCHA NAT.",
                ALF1="MM1  ", ALF2="TELAS", ALF7="KG   ", MCA1=1, MCA2=1,
            ),
            Fctabla1(
                CTAB="SC   ", COD="A    ", DESCRI='SECC. "A"',
                ALF7="UNID ", MCA1=0, MCA2=0,
            ),
            Fctabla1(CTAB="SC   ", COD="SINALF7", DESCRI="SIN UNIDAD DE FACTURACION"),
            # Sección con un multiplicador de precio (ALF4) además de la
            # Unidad de Facturación (ALF7) — usada para probar el preview
            # progresivo de `calcular_precio_preview`.
            Fctabla1(
                CTAB="SC   ", COD="MUL  ", DESCRI="SECC. CON MULTIPLICADOR",
                ALF4="PULG ", ALF7="KG   ", MCA1=0, MCA2=0,
            ),
            # Unidades de medida (NUMSD3 = "Posición", ver tablas_window.py)
            Fctabla1(CTAB="UM   ", COD="MM1  ", DESCRI="MM", NUMSD1=2, NUMSD2=1, NUMSD3=3),
            Fctabla1(CTAB="UM   ", COD="TELAS", DESCRI="TELAS", NUMSD1=1, NUMSD2=0, NUMSD3=4),
            Fctabla1(CTAB="UM   ", COD="KG   ", DESCRI="KILOS", NUMSD1=3, NUMSD2=2, NUMSD3=2),
            Fctabla1(CTAB="UM   ", COD="UNID ", DESCRI="UNIDADES", NUMSD1=5, NUMSD2=0, NUMSD3=6),
            Fctabla1(CTAB="UM   ", COD="PULG ", DESCRI="PULGADAS", NUMSD1=2, NUMSD2=0, NUMSD3=1),
        ]
    )
    db.commit()


def test_resolver_seccion_item_libre():
    fctablas = None  # no debe tocar la DB para "**"
    seccion = resolver_seccion_renglon(fctablas, "**")

    assert seccion.item_libre is True
    assert seccion.segmentos_codigo == []
    assert len(seccion.segmentos_cantidad) == 1
    assert seccion.segmento_facturacion.alf_index == 7
    assert seccion.segmento_facturacion.libre is True


def test_resolver_seccion_gpn_dos_segmentos_codigo_y_una_cantidad(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    seccion = resolver_seccion_renglon(fctablas, "GPN")

    assert seccion.usar_descripcion_seccion is True
    assert seccion.precio_en_dolares is True
    assert [s.cod_unidad for s in seccion.segmentos_codigo] == ["MM1", "TELAS"]
    assert [s.cod_unidad for s in seccion.segmentos_cantidad] == ["KG"]
    assert seccion.segmento_facturacion.cod_unidad == "KG"


def test_resolver_seccion_solo_unidad_facturacion_sin_codigo(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    seccion = resolver_seccion_renglon(fctablas, "A")

    assert seccion.segmentos_codigo == []
    assert [s.cod_unidad for s in seccion.segmentos_cantidad] == ["UNID"]


def test_resolver_seccion_inexistente_levanta_error(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    with pytest.raises(SeccionInexistenteError):
        resolver_seccion_renglon(fctablas, "NO_EXISTE")


def test_resolver_seccion_sin_alf7_levanta_error(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    with pytest.raises(SeccionSinUnidadFacturacionError):
        resolver_seccion_renglon(fctablas, "SINALF7")


def test_armar_codigo_renglon_gpn_3mm_0t(db):
    """Mismo caso real que test_articulo_codigo.py: 3mm/0T -> COD2 "0300"."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "GPN")

    cod2 = armar_codigo_renglon(seccion, [Decimal("3"), Decimal("0")])

    assert cod2 == "0300"


def test_calcular_precio_e_importe_un_solo_multiplicador(db):
    """Sección "A": sólo ALF7 (Unidades) — Importe = Precio × Cantidad,
    Precio Unitario final = el mismo Precio (único divisor = él mismo)."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "A")

    precio_unitario, importe = calcular_precio_e_importe(
        Decimal("10"), {7: Decimal("5")}, seccion
    )

    assert importe == Decimal("50")
    assert precio_unitario == Decimal("10")


def test_calcular_precio_e_importe_multiples_factores(db):
    """GPN: Importe = Precio × Cantidad(KG) (único factor ALF4-7 presente
    es ALF7 en este fixture) — Precio Unitario = Importe / Cantidad."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "GPN")

    precio_unitario, importe = calcular_precio_e_importe(
        Decimal("100"), {7: Decimal("3")}, seccion
    )

    assert importe == Decimal("300")
    assert precio_unitario == Decimal("100")


def test_calcular_precio_e_importe_cantidad_cero_levanta_error(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "A")

    with pytest.raises(ValueError):
        calcular_precio_e_importe(Decimal("10"), {7: Decimal("0")}, seccion)


# ---------------------------------------------------------------------------
# calcular_precio_preview — precio "en construcción" mientras se tipea
# (feedback del usuario, 2026-08-15: mostrar el precio unitario apenas se
# resuelve el Artículo, y actualizarlo a medida que se completan los
# segmentos de cálculo, sin esperar a que la fila esté completa)
# ---------------------------------------------------------------------------


def test_precio_preview_sin_cantidades_cargadas_es_el_precio_base(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "MUL")

    precio = calcular_precio_preview(Decimal("100"), {}, seccion)

    assert precio == Decimal("100")


def test_precio_preview_aplica_multiplicador_ya_tipeado(db):
    """ALF4 (PULG) ya tipeado multiplica el precio; ALF7 (KG, la Unidad
    de Facturación) NO participa del Precio Unitario — daría el Importe."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "MUL")

    precio = calcular_precio_preview(Decimal("100"), {4: Decimal("3")}, seccion)

    assert precio == Decimal("300")


def test_precio_preview_ignora_la_cantidad_de_facturacion_aunque_este_cargada(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "MUL")

    precio = calcular_precio_preview(Decimal("100"), {4: Decimal("3"), 7: Decimal("5")}, seccion)

    assert precio == Decimal("300")  # no 1500 — ALF7 no es un multiplicador de precio


def test_precio_preview_coincide_con_el_precio_unitario_final(db):
    """Una vez completos todos los segmentos, el preview debe coincidir
    con el `precio_unitario` que calcula `calcular_precio_e_importe`."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "MUL")
    valores = {4: Decimal("3"), 7: Decimal("5")}

    preview = calcular_precio_preview(Decimal("100"), valores, seccion)
    precio_unitario, _ = calcular_precio_e_importe(Decimal("100"), valores, seccion)

    assert preview == precio_unitario


def test_descripcion_con_segmentos_omite_valores_en_cero(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    seccion = resolver_seccion_renglon(fctablas, "GPN")

    descripcion = descripcion_con_segmentos(
        seccion, "GOMA PLANCHA NAT.", [Decimal("3"), Decimal("0")]
    )

    assert descripcion == "GOMA PLANCHA NAT. 3 MM"


# ---------------------------------------------------------------------------
# resolver_precio_articulo — conversión $/USD (DetFact.frm:841-853)
# ---------------------------------------------------------------------------


class TestResolverPrecioArticulo:
    def test_factura_pesos_seccion_pesos_no_convierte(self):
        precio = resolver_precio_articulo(
            Decimal("100"), seccion_precio_en_dolares=False, factura_en_dolares=False, cotizacion=Decimal("1000")
        )
        assert precio == Decimal("100")

    def test_factura_pesos_seccion_dolares_multiplica_por_cotizacion(self):
        precio = resolver_precio_articulo(
            Decimal("100"), seccion_precio_en_dolares=True, factura_en_dolares=False, cotizacion=Decimal("1000")
        )
        assert precio == Decimal("100000")

    def test_factura_dolares_seccion_dolares_no_convierte(self):
        precio = resolver_precio_articulo(
            Decimal("100"), seccion_precio_en_dolares=True, factura_en_dolares=True, cotizacion=Decimal("1000")
        )
        assert precio == Decimal("100")

    def test_factura_dolares_seccion_pesos_divide_por_cotizacion(self):
        precio = resolver_precio_articulo(
            Decimal("100000"), seccion_precio_en_dolares=False, factura_en_dolares=True, cotizacion=Decimal("1000")
        )
        assert precio == Decimal("100")


# ---------------------------------------------------------------------------
# posiciones_activas_seccion — columnas de UM a mostrar en Ventas
# (ESTADIST.frm/VTAXART.frm), feedback del usuario 2026-08-16
# ---------------------------------------------------------------------------


class TestPosicionesActivasSeccion:
    def test_seccion_compuesta_devuelve_las_posiciones_reales(self, db):
        _cargar_secciones_y_unidades(db)
        fctablas = RepositoryFactory(db).fctablas()

        # GPN: ALF1=MM1(pos3), ALF2=TELAS(pos4), ALF7=KG(pos2).
        assert posiciones_activas_seccion(fctablas, "GPN") == {2, 3, 4}

    def test_seccion_simple_solo_devuelve_la_posicion_de_facturacion(self, db):
        _cargar_secciones_y_unidades(db)
        fctablas = RepositoryFactory(db).fctablas()

        # "A": sólo ALF7=UNID (pos6), sin segmentos de código.
        assert posiciones_activas_seccion(fctablas, "A") == {6}

    def test_seccion_con_multiplicador_incluye_esa_posicion(self, db):
        _cargar_secciones_y_unidades(db)
        fctablas = RepositoryFactory(db).fctablas()

        # "MUL": ALF4=PULG(pos1), ALF7=KG(pos2).
        assert posiciones_activas_seccion(fctablas, "MUL") == {1, 2}

    def test_seccion_sin_alf7_cae_al_fallback_de_todas_las_posiciones(self, db):
        _cargar_secciones_y_unidades(db)
        fctablas = RepositoryFactory(db).fctablas()

        assert posiciones_activas_seccion(fctablas, "SINALF7") == {1, 2, 3, 4, 6}

    def test_seccion_inexistente_cae_al_fallback_de_todas_las_posiciones(self, db):
        _cargar_secciones_y_unidades(db)
        fctablas = RepositoryFactory(db).fctablas()

        assert posiciones_activas_seccion(fctablas, "NOEXISTE") == {1, 2, 3, 4, 6}
