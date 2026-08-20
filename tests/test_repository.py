"""Tests de migration/repository.py.

Cubren especialmente los bugs encontrados y corregidos auditando el VB6 y
el schema real de fcmenu.mdb (ver data_layer_progress.md):
  - CtascteRepository.deuda_cliente() sumaba mal DEBE en vez de IMPTE/TIPO.
  - StockRepository.criticos() filtraba EST1 (no actualizado por la venta)
    en vez de STUNID.
  - ParametroRepository.get_ivains() buscaba CLAVE='IVAINS' (no existe) en
    vez de la fila única CLAVE='1'.
  - CotizacionRepository.ultima()/by_fecha() no pueden usar ORDER BY/=
    directo en SQL porque FECHA es texto "dd/mm/yyyy".
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from migration.models import (
    Articulo,
    Cheque,
    Cliente,
    Cotizacion,
    Ctascte,
    Despacho,
    DtoxClte,
    Fcestad1,
    Fctabla1,
    FcivaVta,
    NotaArticulo,
    NotaClte,
    Parametro,
    Stock,
)
from migration.repository import RepositoryFactory


# ---------------------------------------------------------------------------
# CtascteRepository.deuda_cliente
# ---------------------------------------------------------------------------


def test_deuda_cliente_suma_debe_y_resta_haber_segun_tipo(db):
    """Algoritmo real (CtaCte.frm:1941-1950): TIPO 0/1/3/7 suman, resto resta."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=0, IMPTE=Decimal("100")),  # sdo ant, suma
            Ctascte(CLTE=1, FECHA=date(2026, 1, 2), TIPO=1, IMPTE=Decimal("500")),  # factura, suma
            Ctascte(CLTE=1, FECHA=date(2026, 1, 3), TIPO=3, IMPTE=Decimal("50")),  # nota débito, suma
            Ctascte(CLTE=1, FECHA=date(2026, 1, 4), TIPO=7, IMPTE=Decimal("20")),  # recibo anulado, suma
            Ctascte(CLTE=1, FECHA=date(2026, 1, 5), TIPO=2, IMPTE=Decimal("30")),  # NC, resta
            Ctascte(CLTE=1, FECHA=date(2026, 1, 6), TIPO=4, IMPTE=Decimal("200")),  # recibo, resta
            Ctascte(CLTE=1, FECHA=date(2026, 1, 7), TIPO=5, IMPTE=Decimal("10")),  # pago a cta, resta
            Ctascte(CLTE=1, FECHA=date(2026, 1, 8), TIPO=6, IMPTE=Decimal("5")),  # descuento, resta
            Ctascte(CLTE=1, FECHA=date(2026, 1, 9), TIPO=8, IMPTE=Decimal("3")),  # NC interés, resta
            Ctascte(CLTE=1, FECHA=date(2026, 1, 10), TIPO=9, IMPTE=Decimal("2")),  # anulación, resta
        ]
    )
    db.commit()

    resultado = repos.ctascte().deuda_cliente(1)

    esperado = (100 + 500 + 50 + 20) - (30 + 200 + 10 + 5 + 3 + 2)
    assert resultado["saldo"] == Decimal(str(esperado))
    assert resultado["deuda"] == resultado["saldo"]


def test_deuda_cliente_no_usa_el_campo_debe(db):
    """Ctascte.DEBE es el saldo puntual de un comprobante (imputaciones
    parciales), no un monto "debe" contable — no debe influir en el saldo
    general del cliente."""
    repos = RepositoryFactory(db)
    db.add(
        Ctascte(
            CLTE=2,
            FECHA=date(2026, 1, 1),
            TIPO=1,
            IMPTE=Decimal("1000"),
            DEBE=Decimal("999999"),  # valor absurdo a propósito: no debe usarse
        )
    )
    db.commit()

    resultado = repos.ctascte().deuda_cliente(2)
    assert resultado["saldo"] == Decimal("1000")


def test_deuda_cliente_cliente_sin_movimientos(db):
    repos = RepositoryFactory(db)
    resultado = repos.ctascte().deuda_cliente(9999)
    assert resultado["saldo"] == Decimal("0")


# ---------------------------------------------------------------------------
# StockRepository.criticos
# ---------------------------------------------------------------------------


def test_criticos_filtra_por_stunid_no_por_est1(db):
    """STUNID es lo único que la venta actualiza (EmiFact.frm) — EST1 es
    legacy multi-depósito no sincronizado. criticos() debe ignorarlo."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            # Crítico real: STUNID bajo mínimo, aunque EST1 esté alto.
            Stock(COD1="A", COD2="", STMIN=Decimal("10"), STUNID=Decimal("5"), EST1=999),
            # No crítico: STUNID por encima del mínimo, aunque EST1 esté bajo
            # (si el bug reapareciera y filtrara por EST1, este ítem
            # aparecería incorrectamente como crítico).
            Stock(COD1="B", COD2="", STMIN=Decimal("10"), STUNID=Decimal("50"), EST1=0),
        ]
    )
    db.commit()

    criticos = repos.stock().criticos()
    codigos = {c.COD1 for c in criticos}
    assert codigos == {"A"}


# ---------------------------------------------------------------------------
# ParametroRepository.get_config / get_ivains
# ---------------------------------------------------------------------------


def test_get_config_lee_la_fila_unica_clave_1(db):
    """FCMENU.bas:333 siempre lee CLAVE='1', no una clave por nombre."""
    repos = RepositoryFactory(db)
    db.add(Parametro(CLAVE="1", IVAINS=Decimal("21"), IVANI=Decimal("0")))
    db.commit()

    config = repos.parametro().get_config()
    assert config is not None
    assert config.IVAINS == Decimal("21")

    # get_ivains() es un alias sobre la misma fila (antes buscaba CLAVE='IVAINS').
    assert repos.parametro().get_ivains() is config


def test_get_config_ausente_no_rompe(db):
    repos = RepositoryFactory(db)
    assert repos.parametro().get_config() is None


def test_get_config_tolera_el_padding_fijo_de_access(db):
    """Bug real encontrado corriendo Parámetros contra fcmenu_dev: CLAVE
    quedó como '1 ' (con espacio), no '1' — get_config() devolvía None
    contra datos reales pese a que la fila de configuración sí existe."""
    repos = RepositoryFactory(db)
    db.add(Parametro(CLAVE="1 ", NOMEMPR="ALESTEL SRL"))
    db.commit()

    config = repos.parametro().get_config()
    assert config is not None
    assert config.NOMEMPR == "ALESTEL SRL"


# ---------------------------------------------------------------------------
# CotizacionRepository.ultima / by_fecha (FECHA es texto "dd/mm/yyyy")
# ---------------------------------------------------------------------------


def test_ultima_ordena_cronologicamente_no_lexicograficamente(db):
    """'01/12/2026' es MENOR como texto que '02/01/2010', pero es 16 años
    posterior — ultima() debe devolver la fecha real más reciente."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Cotizacion(FECHA="02/01/2010", DOLAR=Decimal("3.8")),
            Cotizacion(FECHA="01/12/2026", DOLAR=Decimal("1200")),
            Cotizacion(FECHA="15/06/2020", DOLAR=Decimal("70")),
        ]
    )
    db.commit()

    ultima = repos.cotizacion().ultima()
    assert ultima.FECHA == "01/12/2026"
    assert ultima.DOLAR == Decimal("1200")


def test_by_fecha_recibe_date_de_python(db):
    repos = RepositoryFactory(db)
    db.add(Cotizacion(FECHA="05/03/2026", DOLAR=Decimal("900")))
    db.commit()

    encontrada = repos.cotizacion().by_fecha(date(2026, 3, 5))
    assert encontrada is not None
    assert encontrada.DOLAR == Decimal("900")

    assert repos.cotizacion().by_fecha(date(2026, 3, 6)) is None


def test_guardar_crea_si_no_existe_la_fecha(db):
    repos = RepositoryFactory(db)

    creada = repos.cotizacion().guardar(date(2026, 3, 5), {"DOLAR": Decimal("900")})

    assert creada.FECHA == "05/03/2026"
    assert repos.cotizacion().by_fecha(date(2026, 3, 5)).DOLAR == Decimal("900")


def test_guardar_actualiza_si_ya_existe_la_fecha(db):
    repos = RepositoryFactory(db)
    db.add(Cotizacion(FECHA="05/03/2026", DOLAR=Decimal("900")))
    db.commit()

    repos.cotizacion().guardar(date(2026, 3, 5), {"DOLAR": Decimal("950")})

    assert repos.cotizacion().by_fecha(date(2026, 3, 5)).DOLAR == Decimal("950")


# ---------------------------------------------------------------------------
# Sanity de CRUD genérico / búsquedas simples ya existentes
# ---------------------------------------------------------------------------


def test_cliente_by_codigo_y_by_nombre(db):
    repos = RepositoryFactory(db)
    db.add(Cliente(CODIGO=1001, NOMB="Juan Perez"))
    db.commit()

    assert repos.cliente().by_codigo(1001).NOMB == "Juan Perez"
    assert len(repos.cliente().by_nombre("perez")) == 1
    assert repos.cliente().by_codigo(9999) is None


def test_cliente_desde_codigo_ordena_ascendente_desde_el_numero_pedido(db):
    """`ClienteBusquedaWindow`: si el código tipeado no existe, se lista
    igual arrancando desde ahí (feedback del usuario, 2026-08-15)."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Cliente(CODIGO=10, NOMB="Diez"),
            Cliente(CODIGO=50, NOMB="Cincuenta"),
            Cliente(CODIGO=30, NOMB="Treinta"),
            Cliente(CODIGO=20, NOMB="Veinte"),
        ]
    )
    db.commit()

    resultado = repos.cliente().desde_codigo(20)

    assert [c.CODIGO for c in resultado] == [20, 30, 50]


def test_cliente_desde_codigo_respeta_el_limite(db):
    repos = RepositoryFactory(db)
    db.add_all([Cliente(CODIGO=i, NOMB=f"Cliente {i}") for i in range(1, 6)])
    db.commit()

    resultado = repos.cliente().desde_codigo(1, limite=2)

    assert [c.CODIGO for c in resultado] == [1, 2]


def test_articulo_by_cod1(db):
    repos = RepositoryFactory(db)
    db.add(Articulo(COD1="ART1", COD2="X", DESCRI="Tornillo"))
    db.commit()

    encontrado = repos.articulo().by_cod1("ART1")
    assert encontrado is not None
    assert encontrado.DESCRI == "Tornillo"


# ---------------------------------------------------------------------------
# ArticuloRepository / Fcestad1Repository — padding real de Access en COD1/COD2
# ---------------------------------------------------------------------------


def test_articulo_by_cod1_cod2_tolera_el_padding_fijo_de_access(db):
    """Bug real encontrado corriendo AbmArt contra fcmenu_dev: COD1/COD2
    quedaron con el padding fijo original ('GPN  ', '0300      ') igual
    que Fctabla1.CTAB. Buscar con los valores sin padding devolvía None."""
    repos = RepositoryFactory(db)
    db.add(Articulo(COD1="GPN  ", COD2="0300      ", DESCRI="GOMA 3MM 0T"))
    db.commit()

    encontrado = repos.articulo().by_cod1_cod2("GPN", "0300")
    assert encontrado is not None
    assert encontrado.DESCRI == "GOMA 3MM 0T"


def test_fcestad1_by_cod1_cod2_tolera_el_padding_fijo_de_access(db):
    repos = RepositoryFactory(db)
    db.add(
        Fcestad1(
            COD1="GPN  ", COD2="0300      ", TIPO=1, LETRA="A", PTOVTA=3, CPBTE=1,
            ITEM=1, CLTE=1, FECHA=date(2026, 1, 1),
        )
    )
    db.commit()

    encontrados = repos.fcestad1().by_cod1_cod2("GPN", "0300")
    assert len(encontrados) == 1


# ---------------------------------------------------------------------------
# DespachoRepository (DetFact.frm Sub VerSiBusca / StockDespa.frm) — mismo
# padding real de Access que Articulo/Fcestad1, más el filtro STOCK > 0.
# ---------------------------------------------------------------------------


def test_despacho_con_stock_tolera_el_padding_fijo_de_access(db):
    repos = RepositoryFactory(db)
    db.add(
        Despacho(
            COD1="GPN  ", COD2="0300      ", NRODESP="18001IC04080837W",
            CPBTE=90418, ENTRADA=Decimal("105"), SALIDAS=Decimal("43"),
            STOCK=Decimal("62"), FECENT=date(2018, 4, 13),
        )
    )
    db.commit()

    encontrados = repos.despacho().con_stock("GPN", "0300")
    assert len(encontrados) == 1
    assert encontrados[0].NRODESP == "18001IC04080837W"


def test_despacho_con_stock_excluye_lotes_agotados(db):
    repos = RepositoryFactory(db)
    db.add(Despacho(COD1="GPN", COD2="0300", NRODESP="AGOTADO", STOCK=Decimal("0")))
    db.add(Despacho(COD1="GPN", COD2="0300", NRODESP="DISPONIBLE", STOCK=Decimal("10")))
    db.commit()

    encontrados = repos.despacho().con_stock("GPN", "0300")
    assert [d.NRODESP for d in encontrados] == ["DISPONIBLE"]


def test_despacho_by_nrodesp(db):
    repos = RepositoryFactory(db)
    db.add(Despacho(COD1="GPN  ", COD2="0300      ", NRODESP="X1", STOCK=Decimal("5")))
    db.commit()

    encontrado = repos.despacho().by_nrodesp("GPN", "0300", "X1")
    assert encontrado is not None
    assert encontrado.STOCK == Decimal("5")


def test_despacho_resumen_lotes_agrupa_por_nrodesp_y_fecent_con_comprobante(db):
    """Columna "Comprobante" (pedido del usuario, 2026-08-19) — el CPBTE
    de ENTRADA de cada lote, junto a la cantidad de artículos que ya
    traía `resumen_lotes()`."""
    repos = RepositoryFactory(db)
    db.add(Despacho(COD1="AA", COD2="1", NRODESP="L001", CPBTE=500, FECENT=date(2026, 1, 5), STOCK=Decimal("10")))
    db.add(Despacho(COD1="AA", COD2="2", NRODESP="L001", CPBTE=500, FECENT=date(2026, 1, 5), STOCK=Decimal("20")))
    db.add(Despacho(COD1="BB", COD2="1", NRODESP="L002", CPBTE=600, FECENT=date(2026, 2, 1), STOCK=Decimal("5")))
    db.commit()

    resumen = {f["nrodesp"]: f for f in repos.despacho().resumen_lotes()}
    assert resumen["L001"]["cpbte"] == 500
    assert resumen["L001"]["cantidad"] == 2
    assert resumen["L002"]["cpbte"] == 600
    assert resumen["L002"]["cantidad"] == 1


def test_despacho_articulos_de_lote_filtra_por_fecent_si_el_nrodesp_se_repite(db):
    """Bug real encontrado por el usuario con datos reales (2026-08-19):
    "22001IC05005447M dice que tiene 2 artículos y al desplegar muestra
    más de 15 renglones" — ese NRODESP real tiene 27 filas con
    FECENT=17/03/2022 y 2 más con FECENT=05/01/2023 (2 lotes de
    importación distintos que reusan el número); sin filtrar también por
    `fecent`, `articulos_de_lote()` traía las 29 juntas sin importar cuál
    de los 2 renglones del resumen se hubiera clickeado."""
    repos = RepositoryFactory(db)
    db.add(Despacho(COD1="AA", COD2="1", NRODESP="REPETIDO", FECENT=date(2022, 3, 17)))
    db.add(Despacho(COD1="BB", COD2="1", NRODESP="REPETIDO", FECENT=date(2022, 3, 17)))
    db.add(Despacho(COD1="CC", COD2="1", NRODESP="REPETIDO", FECENT=date(2023, 1, 5)))
    db.commit()

    todos = repos.despacho().articulos_de_lote("REPETIDO")
    assert len(todos) == 3  # comportamiento previo, sin fecent, se mantiene disponible

    solo_2022 = repos.despacho().articulos_de_lote("REPETIDO", date(2022, 3, 17))
    assert sorted(d.COD1 for d in solo_2022) == ["AA", "BB"]

    solo_2023 = repos.despacho().articulos_de_lote("REPETIDO", date(2023, 1, 5))
    assert [d.COD1 for d in solo_2023] == ["CC"]


# ---------------------------------------------------------------------------
# ClienteRepository.proximo_codigo (Abmclte.frm Sub Blanquea, líneas 1839-1891)
# ---------------------------------------------------------------------------


def test_proximo_codigo_sin_clientes_devuelve_1(db):
    repos = RepositoryFactory(db)
    assert repos.cliente().proximo_codigo() == 1


def test_proximo_codigo_es_el_maximo_mas_uno(db):
    repos = RepositoryFactory(db)
    db.add_all([Cliente(CODIGO=10), Cliente(CODIGO=25), Cliente(CODIGO=3)])
    db.commit()

    assert repos.cliente().proximo_codigo() == 26


def test_proximo_codigo_ignora_codigos_reservados_desde_99900(db):
    """Los códigos >= 99900 son genéricos (p.ej. Consumidor Final de prueba)
    y no deben desplazar el autonumerado normal (Abmclte.frm:1841-1842)."""
    repos = RepositoryFactory(db)
    db.add_all([Cliente(CODIGO=50), Cliente(CODIGO=99900), Cliente(CODIGO=99999)])
    db.commit()

    assert repos.cliente().proximo_codigo() == 51


# ---------------------------------------------------------------------------
# FctablasRepository.by_ctab / by_ctab_cod — padding real de Access
# ---------------------------------------------------------------------------


def test_by_ctab_tolera_el_padding_fijo_de_access(db):
    """Bug real encontrado corriendo contra fcmenu_dev (Postgres): la
    migración desde Access dejó CTAB con el padding fijo original
    ('CV   ', 5 caracteres, igual que el literal SQL de Abmclte.frm
    Form_Load: WHERE CTAB = 'CV   '). Buscar por 'CV' sin padding contra
    esos datos reales devolvía 0 filas."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Fctabla1(CTAB="CV   ", COD="1 ", DESCRI="Contado"),
            Fctabla1(CTAB="CV   ", COD="2 ", DESCRI="Cuenta Cte."),
            Fctabla1(CTAB="VD   ", COD="1 ", DESCRI="Pérez"),
        ]
    )
    db.commit()

    resultado = repos.fctablas().by_ctab("CV")
    assert [f.DESCRI for f in resultado] == ["Contado", "Cuenta Cte."]


def test_by_ctab_cod_tolera_el_padding_fijo_de_access(db):
    repos = RepositoryFactory(db)
    db.add(Fctabla1(CTAB="ZN   ", COD="3 ", DESCRI="Norte"))
    db.commit()

    encontrado = repos.fctablas().by_ctab_cod("ZN", "3")
    assert encontrado is not None
    assert encontrado.DESCRI == "Norte"


# ---------------------------------------------------------------------------
# NotaClteRepository.by_cliente
# ---------------------------------------------------------------------------


def test_notaclte_by_cliente_encuentra_la_nota(db):
    repos = RepositoryFactory(db)
    db.add(NotaClte(CLTE=1, NOTA1="Cliente moroso", NOTA2="Llamar antes de las 18hs"))
    db.commit()

    nota = repos.notaclte().by_cliente(1)
    assert nota is not None
    assert nota.NOTA1 == "Cliente moroso"


def test_notaclte_by_cliente_sin_nota_devuelve_none(db):
    repos = RepositoryFactory(db)
    assert repos.notaclte().by_cliente(999) is None


# ---------------------------------------------------------------------------
# NotaArticuloRepository.by_articulo
# ---------------------------------------------------------------------------


def test_notaarticulo_by_articulo_encuentra_la_nota(db):
    repos = RepositoryFactory(db)
    db.add(NotaArticulo(COD1="GPN", COD2="0300", NOTA3="Sólo venta mayorista", NOTA4="Pedir seña"))
    db.commit()

    nota = repos.notaarticulo().by_articulo("GPN", "0300")
    assert nota is not None
    assert nota.NOTA3 == "Sólo venta mayorista"


def test_notaarticulo_by_articulo_tolera_padding(db):
    repos = RepositoryFactory(db)
    db.add(NotaArticulo(COD1="GPN  ", COD2="0300      ", NOTA3="Nota"))
    db.commit()

    assert repos.notaarticulo().by_articulo("GPN", "0300") is not None


def test_notaarticulo_by_articulo_sin_nota_devuelve_none(db):
    repos = RepositoryFactory(db)
    assert repos.notaarticulo().by_articulo("GPN", "9999") is None


# ---------------------------------------------------------------------------
# DtoxClteRepository.by_cliente / reemplazar_para_cliente
# ---------------------------------------------------------------------------


def test_dtoxclte_by_cliente_lista_todas_las_secciones(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            DtoxClte(CLTE=26, SECCION="A", DTO1=Decimal("48")),
            DtoxClte(CLTE=26, SECCION="B", DTO1=Decimal("48")),
            DtoxClte(CLTE=99, SECCION="A", DTO1=Decimal("10")),
        ]
    )
    db.commit()

    filas = repos.dtoxclte().by_cliente(26)
    assert {f.SECCION for f in filas} == {"A", "B"}


def test_reemplazar_para_cliente_borra_lo_anterior_e_inserta_lo_nuevo(db):
    repos = RepositoryFactory(db)
    db.add(DtoxClte(CLTE=26, SECCION="VIEJA", DTO1=Decimal("5")))
    db.commit()

    cantidad = repos.dtoxclte().reemplazar_para_cliente(
        26,
        [
            {"SECCION": "A", "DTO1": Decimal("48"), "DTO2": Decimal("0")},
            {"SECCION": "B", "DTO1": Decimal("10")},
        ],
    )

    assert cantidad == 2
    secciones = {f.SECCION for f in repos.dtoxclte().by_cliente(26)}
    assert secciones == {"A", "B"}


def test_reemplazar_para_cliente_deduplica_seccion_repetida_sin_abortar(db):
    """El legacy (DtosxClte.frm Grabar(), línea 856) abandonaba TODA la
    transacción sin guardar nada si el operador repetía una Sección en la
    grilla (un `Exit Sub` sin commit ni rollback). Acá se deduplica en vez
    de perder el resto de las filas."""
    repos = RepositoryFactory(db)

    cantidad = repos.dtoxclte().reemplazar_para_cliente(
        26,
        [
            {"SECCION": "A", "DTO1": Decimal("10")},
            {"SECCION": "A", "DTO1": Decimal("20")},  # repetida: gana esta
            {"SECCION": "B", "DTO1": Decimal("30")},
        ],
    )

    assert cantidad == 2
    filas = {f.SECCION: f.DTO1 for f in repos.dtoxclte().by_cliente(26)}
    assert filas == {"A": Decimal("20"), "B": Decimal("30")}


def test_reemplazar_para_cliente_con_lista_vacia_borra_todo(db):
    repos = RepositoryFactory(db)
    db.add(DtoxClte(CLTE=26, SECCION="A", DTO1=Decimal("5")))
    db.commit()

    cantidad = repos.dtoxclte().reemplazar_para_cliente(26, [])

    assert cantidad == 0
    assert repos.dtoxclte().by_cliente(26) == []


# ---------------------------------------------------------------------------
# CtascteRepository — extracto anual / saldo vencido / saldos de todos
# los clientes (CtaCte.frm, Consulta de Cuenta Corriente)
# ---------------------------------------------------------------------------


def test_saldo_inicial_usa_la_misma_formula_tipos_debe_del_saldo_corriente(db):
    """Corregido 2026-08-16 (cuenta real 407, ver pyqt6_ui_progress.md):
    el Saldo Inicial de un año usa la MISMA fórmula que `deuda_cliente()`
    (`TIPOS_DEBE = {0,1,3,7}` suma, el resto resta) — así siempre
    reconcilia con "mostrar todo". Antes usaba una fórmula distinta e
    incompleta (sólo TIPO 1/3 débito, 2/4 crédito) que dejaba afuera
    TIPO 5/6/8/9, generando saldos finales distintos según el año
    elegido para el mismo cliente."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Ctascte(CLTE=1, FECHA=date(2024, 1, 1), TIPO=1, IMPTE=Decimal("1000")),  # debe
            Ctascte(CLTE=1, FECHA=date(2024, 6, 1), TIPO=4, IMPTE=Decimal("300")),  # crédito
            Ctascte(CLTE=1, FECHA=date(2024, 7, 1), TIPO=2, IMPTE=Decimal("50")),  # crédito
            Ctascte(CLTE=1, FECHA=date(2024, 8, 1), TIPO=6, IMPTE=Decimal("20")),  # crédito (antes se ignoraba)
            Ctascte(CLTE=1, FECHA=date(2025, 1, 1), TIPO=1, IMPTE=Decimal("999")),  # otro año previo, debe
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("5000")),  # del año pedido, NO cuenta
        ]
    )
    db.commit()

    inicial = repos.ctascte().saldo_inicial(1, 2026)

    assert inicial == {"debe": Decimal("1999"), "credito": Decimal("370")}


def test_saldo_inicial_reconcilia_con_mostrar_todo_sin_importar_el_anio_elegido(db):
    """Regresión directa del bug real reportado por el usuario (cuenta
    407): el saldo final debe ser el MISMO sin importar qué año se elija
    como "desde" — antes daba resultados distintos para el mismo cliente
    según se filtrara desde el principio o desde un año posterior."""
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Ctascte(CLTE=1, FECHA=date(2010, 1, 1), TIPO=1, IMPTE=Decimal("1000")),
            Ctascte(CLTE=1, FECHA=date(2011, 1, 1), TIPO=6, IMPTE=Decimal("412.86")),  # descuento viejo
            Ctascte(CLTE=1, FECHA=date(2013, 1, 1), TIPO=1, IMPTE=Decimal("500")),
            Ctascte(CLTE=1, FECHA=date(2014, 1, 1), TIPO=4, IMPTE=Decimal("999.14")),
        ]
    )
    db.commit()
    repo = repos.ctascte()

    def saldo_final_desde(anio: int) -> Decimal:
        inicial = repo.saldo_inicial(1, anio)
        saldo = inicial["debe"] - inicial["credito"]
        for m in repo.extracto(1, anio):
            importe = m.IMPTE or Decimal("0")
            saldo += importe if m.TIPO in repo.TIPOS_DEBE else -importe
        return saldo

    # 2010-1000(debe) - 2011-412.86(crédito) + 2013-500(debe) - 2014-999.14(crédito) = 88.00
    assert saldo_final_desde(2005) == Decimal("88.00")
    assert saldo_final_desde(2013) == Decimal("88.00")
    assert saldo_final_desde(2014) == Decimal("88.00")


def test_extracto_ordena_por_fecha_cpbte_tipo_desde_el_anio_pedido(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Ctascte(CLTE=1, FECHA=date(2025, 12, 31), TIPO=1, CPBTE=1, IMPTE=Decimal("1")),  # año anterior, afuera
            Ctascte(CLTE=1, FECHA=date(2026, 3, 1), TIPO=4, CPBTE=2, IMPTE=Decimal("2")),
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, CPBTE=3, IMPTE=Decimal("3")),
        ]
    )
    db.commit()

    filas = repos.ctascte().extracto(1, 2026)

    assert [f.CPBTE for f in filas] == [3, 2]


def test_saldo_vencido_solo_tipo_1_3_con_debe_positivo_y_fecvto_pasada(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("700"), FECVTO=date(2026, 2, 1)),  # vencida
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("300"), FECVTO=date(2099, 1, 1)),  # futura
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, DEBE=Decimal("0"), FECVTO=date(2026, 1, 1)),  # ya saldada
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=4, DEBE=Decimal("999"), FECVTO=date(2026, 1, 1)),  # no es 1/3
        ]
    )
    db.commit()

    vencido = repos.ctascte().saldo_vencido(1, hoy=date(2026, 8, 16))

    assert vencido == Decimal("700")


def test_saldos_todos_clientes_ordena_por_nombre_e_incluye_clientes_sin_movimientos(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Cliente(CODIGO=2, NOMB="Beto"),
            Cliente(CODIGO=1, NOMB="Ana"),
            Ctascte(CLTE=1, FECHA=date(2026, 1, 1), TIPO=1, IMPTE=Decimal("100")),
        ]
    )
    db.commit()

    filas = repos.ctascte().saldos_todos_clientes()

    assert [f["nombre"] for f in filas] == ["Ana", "Beto"]
    assert filas[0]["saldo"] == Decimal("100")
    assert filas[1]["saldo"] == Decimal("0")  # sin movimientos, no debe romper (outer join)


# ---------------------------------------------------------------------------
# ChequeRepository — códigos reales de ESTADO (confirmados 2026-08-16
# contra VerCheq.frm/EmiRec.frm, no las palabras "COBRADO"/"RECHAZADO"
# que tenían por_cobrar()/rechazados() antes, sin ningún call site real)
# ---------------------------------------------------------------------------


def test_por_cobrar_usa_el_codigo_real_1_en_cartera(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Cheque(NROCHEQ=1, ESTADO="1", IMPORTE=Decimal("100")),
            Cheque(NROCHEQ=2, ESTADO="2", IMPORTE=Decimal("200")),
        ]
    )
    db.commit()

    assert [c.NROCHEQ for c in repos.cheque().por_cobrar()] == [1]


def test_en_cartera_de_cliente_filtra_por_cliente_y_estado_ordena_por_vencimiento(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            Cheque(NROCHEQ=10, CLTE=50, ESTADO="1", FECVTO=date(2026, 9, 1)),
            Cheque(NROCHEQ=11, CLTE=50, ESTADO="1", FECVTO=date(2026, 8, 1)),
            Cheque(NROCHEQ=12, CLTE=50, ESTADO="2", FECVTO=date(2026, 7, 1)),  # egresado, afuera
            Cheque(NROCHEQ=13, CLTE=99, ESTADO="1", FECVTO=date(2026, 6, 1)),  # otro cliente, afuera
        ]
    )
    db.commit()

    cheques = repos.cheque().en_cartera_de_cliente(50)

    assert [c.NROCHEQ for c in cheques] == [11, 10]


# ---------------------------------------------------------------------------
# FcivaVtaRepository.ultima_venta
# ---------------------------------------------------------------------------


def test_ultima_venta_devuelve_la_fecha_mas_reciente_del_cliente(db):
    repos = RepositoryFactory(db)
    db.add_all(
        [
            FcivaVta(FECHA=date(2026, 1, 1), CLTE=50, TIPO="1"),
            FcivaVta(FECHA=date(2026, 5, 1), CLTE=50, TIPO="1"),
            FcivaVta(FECHA=date(2099, 1, 1), CLTE=99, TIPO="1"),  # otro cliente, afuera
        ]
    )
    db.commit()

    assert repos.fciva_vta().ultima_venta(50) == date(2026, 5, 1)


def test_ultima_venta_sin_facturas_devuelve_none(db):
    repos = RepositoryFactory(db)
    assert repos.fciva_vta().ultima_venta(50) is None


# ---------------------------------------------------------------------------
# FcivaVtaRepository.by_comprobante
# ---------------------------------------------------------------------------


def test_fciva_vta_by_comprobante_tolera_el_padding_fijo_de_access(db):
    """Bug real encontrado probando "Ver factura" desde Cuentas Corrientes
    con datos reales (2026-08-19): `FcivaVta.TIPO` en `fcmenu_dev` real
    viene con padding fijo de Access (`'1 '`, no `'1'`) — el `==` exacto
    contra el TIPO limpio que arma `CtaCteWindow` nunca encontraba nada,
    devolvía `None` en silencio y el clic no hacía nada visible."""
    repos = RepositoryFactory(db)
    db.add(FcivaVta(TIPO="1 ", LETRA="A ", PTOVTA=1, CPBTE=20120, CLTE=3916))
    db.commit()

    encontrada = repos.fciva_vta().by_comprobante("1", "A", 1, 20120)
    assert encontrada is not None
    assert encontrada.CLTE == 3916


def test_fciva_vta_by_comprobante_sin_coincidencia_devuelve_none(db):
    repos = RepositoryFactory(db)
    assert repos.fciva_vta().by_comprobante("1", "A", 1, 99999) is None
