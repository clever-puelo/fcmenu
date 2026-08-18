"""Tests de migration/ui/articulo_codigo.py — el generador de código
compuesto de Articulo.COD2.

Los casos de "GOMA PLANCHA NAT." (sección GPN) están tomados de datos
reales de `fcmenu_dev` (2026-08-14): COD2 "0300"/"0301" para 3mm con 0/1
telas, confirma el ancho de segmentos (MM1: 2+1 dígitos, TELAS: 1+0).
"""

from __future__ import annotations

from decimal import Decimal

from migration.models import Fctabla1
from migration.repository import RepositoryFactory
from migration.ui.articulo_codigo import (
    armar_codigo,
    descripcion_automatica,
    formatear_segmento,
    resolver_segmentos,
    separar_codigo,
)


def _cargar_secciones_y_unidades(db):
    db.add_all(
        [
            # Secciones
            Fctabla1(CTAB="SC   ", COD="A    ", DESCRI='SECC. "A"', ALF1="NRO  ", ALF2="     "),
            Fctabla1(CTAB="SC   ", COD="GPN  ", DESCRI="GOMA PLANCHA NAT.", ALF1="MM1  ", ALF2="TELAS"),
            Fctabla1(CTAB="SC   ", COD="AC   ", DESCRI="ACARREADOR", ALF1="MTRS ", ALF2="     "),
            # Unidades de medida
            Fctabla1(CTAB="UM   ", COD="NRO  ", DESCRI=None, NUMSD1=5, NUMSD2=0),
            Fctabla1(CTAB="UM   ", COD="MM1  ", DESCRI="MM", NUMSD1=2, NUMSD2=1),
            Fctabla1(CTAB="UM   ", COD="TELAS", DESCRI="TELAS", NUMSD1=1, NUMSD2=0),
            Fctabla1(CTAB="UM   ", COD="MTRS ", DESCRI="MTRS", NUMSD1=3, NUMSD2=0),
        ]
    )
    db.commit()


def test_resolver_segmentos_seccion_nro_es_un_solo_segmento_libre(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    segmentos = resolver_segmentos(fctablas, "A")

    assert len(segmentos) == 1
    assert segmentos[0].libre is True
    assert segmentos[0].cod_unidad == "NRO"


def test_resolver_segmentos_seccion_con_dos_dimensiones(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    segmentos = resolver_segmentos(fctablas, "GPN")

    assert len(segmentos) == 2
    assert segmentos[0].cod_unidad == "MM1"
    assert segmentos[0].numsd1 == 2 and segmentos[0].numsd2 == 1
    assert segmentos[1].cod_unidad == "TELAS"
    assert segmentos[1].numsd1 == 1 and segmentos[1].numsd2 == 0


def test_resolver_segmentos_seccion_una_dimension_fisica(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    segmentos = resolver_segmentos(fctablas, "AC")

    assert len(segmentos) == 1
    assert segmentos[0].libre is False
    assert segmentos[0].numsd1 == 3 and segmentos[0].numsd2 == 0


def test_resolver_segmentos_seccion_inexistente_devuelve_vacio(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()

    assert resolver_segmentos(fctablas, "NO_EXISTE") == []


def test_armar_codigo_gpn_3mm_0_telas_matchea_dato_real(db):
    """COD2 real en fcmenu_dev para GPN 3mm/0T es '0300'."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "GPN")

    codigo = armar_codigo(segmentos, [Decimal("3.0"), Decimal("0")])

    assert codigo == "0300"


def test_armar_codigo_gpn_3mm_1_tela_matchea_dato_real(db):
    """COD2 real en fcmenu_dev para GPN 3mm/1T es '0301'."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "GPN")

    codigo = armar_codigo(segmentos, [Decimal("3.0"), Decimal("1")])

    assert codigo == "0301"


def test_armar_codigo_seccion_nro_no_paddea(db):
    """Sección NRO: el código es el número tal cual lo tipeó el operador,
    sin ceros a la izquierda (dato real: sección 'A' código '100', no
    '00100')."""
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "A")

    assert armar_codigo(segmentos, [Decimal("100")]) == "100"


def test_armar_codigo_una_dimension_fisica_paddea_a_su_ancho(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "AC")

    assert armar_codigo(segmentos, [Decimal("7")]) == "007"


def test_armar_codigo_cantidad_de_valores_no_coincide_lanza_error(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "GPN")

    import pytest

    with pytest.raises(ValueError):
        armar_codigo(segmentos, [Decimal("3.0")])


def test_separar_codigo_es_inversa_de_armar_codigo_dos_segmentos(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "GPN")

    codigo = armar_codigo(segmentos, [Decimal("3.0"), Decimal("1")])
    valores = separar_codigo(codigo, segmentos)

    assert valores == [Decimal("3.0"), Decimal("1")]


def test_separar_codigo_seccion_nro(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmentos = resolver_segmentos(fctablas, "A")

    assert separar_codigo("100", segmentos) == [Decimal("100")]


def test_formatear_segmento_redondea_half_up(db):
    _cargar_secciones_y_unidades(db)
    fctablas = RepositoryFactory(db).fctablas()
    segmento = resolver_segmentos(fctablas, "GPN")[0]  # MM1: 2 enteros, 1 decimal

    assert formatear_segmento(Decimal("3.05"), segmento) == "031"  # 3,05 -> 3,1 (half up)


def test_descripcion_automatica_cuando_operador_deja_blanco():
    assert descripcion_automatica("GPN", "0300") == "GPN 0300"
