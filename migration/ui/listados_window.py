"""`ListadosWindow` — migración de `Listados.frm` ("Listados Varios",
menú `LISTADO1` de `FCMENU.frm`).

Una sola ventana con un selector de reporte (12 tipos reales — ver
`migration.services.ListadosService` para el detalle de qué se migró,
qué se reconstruyó con contenido real ["Deuda Pendiente"/"Lista de
Precios"] y qué se descartó por código muerto/abandonado) + un panel de
filtros que cambia según el reporte elegido (mismo patrón dinámico que
el legacy, Frame3/Frame4/Frame5), y un botón "Generar PDF" que arma el
reporte y lo muestra en el visor ya usado por Factura/Recibo
(`PdfPreviewDialog`) — reemplaza el visor genérico `IMPRE.frm`/
`VSPrinter`, que no se migra como pantalla propia.
"""

from __future__ import annotations

from datetime import date, timedelta
from decimal import Decimal

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QDateEdit,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.pdf import DIR_SALIDA_DEFAULT, generar_pdf_listado
from migration.provincias import NOMBRE_PROVINCIA, nombre_provincia
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import ListadosService

from .cliente_busqueda_window import ClienteBusquedaWindow
from .pdf_preview_dialog import PdfPreviewDialog
from .widgets import crear_boton_hoy

REPORTES = [
    ("clientes", "Clientes"),
    ("precios", "Lista de Precios"),
    ("subdiario_ventas", "Subdiario de Ventas"),
    ("subdiario_cobranzas", "Subd. de Cobranzas AFIP"),
    ("ingresos_brutos", "Ingresos Brutos"),
    ("deuda_pendiente", "Deuda Pendiente"),
    ("planilla_cobranzas", "Planilla de Cobranzas"),
    ("estado_cuenta", "Estado de Cuenta"),
    ("saldos", "Saldos de Cta. Cte."),
    ("percepciones_arba", "Percepciones ARBA"),
    ("comisiones_cobranzas", "Comis. x Cobr. (nvo.)"),
    ("subdiario_comisiones", "Subdiario Vtas. (Comisiones)"),
]
USA_FECHAS = {
    "subdiario_ventas", "subdiario_cobranzas", "ingresos_brutos",
    "percepciones_arba", "comisiones_cobranzas", "subdiario_comisiones",
}
USA_CLIENTE_FILTRO = {"clientes", "deuda_pendiente", "planilla_cobranzas", "estado_cuenta", "saldos"}
USA_VENDEDOR = {"comisiones_cobranzas", "subdiario_comisiones"}
USA_SECCION = {"precios"}
VENDEDOR_OBLIGATORIO = {"comisiones_cobranzas"}


class ListadosWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None, *, reporte_inicial: str):
        super().__init__(parent)
        # Ya no hay combo para elegir/cambiar el listado acá adentro
        # (feedback del usuario, 2026-08-18: "quitar el combo de la
        # ventana para seleccionar el listado, ya se selecciona desde
        # el boton superior") — cada botón de la barra de tareas de
        # "Listados" abre esta ventana con SU reporte ya elegido, sin
        # forma de cambiarlo desde acá adentro (para eso está el botón
        # de otro listado en la barra de tareas). El título de la
        # ventana ya queda como la etiqueta del botón clickeado (ver
        # `MainMenuWindow._mostrar`), pero por si se instancia standalone
        # (`main_listados.py`) también se arma un título razonable acá.
        self._reporte = reporte_inicial
        self.setWindowTitle(f"Listados Varios — {dict(REPORTES).get(reporte_inicial, reporte_inicial)}")
        # Ventana 100% más ancha (feedback del usuario, 2026-08-18).
        self.resize(1280, 420)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.servicio = ListadosService(self.db)

        self._construir_ui()
        self._poblar_combos_tablas()
        self._on_reporte_cambiado()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        lbl_reporte = QLabel(dict(REPORTES).get(self._reporte, self._reporte))
        lbl_reporte.setStyleSheet("font-size: 13pt; font-weight: bold;")
        layout.addWidget(lbl_reporte)

        # -- Fechas ------------------------------------------------------
        self.grupo_fechas = QGroupBox("Período")
        fila_f = QHBoxLayout(self.grupo_fechas)
        hoy = QDate.currentDate()
        fila_f.addWidget(QLabel("Desde :"))
        self.fecha_desde = QDateEdit()
        self.fecha_desde.setCalendarPopup(True)
        self.fecha_desde.setDate(QDate(hoy.year(), hoy.month(), 1))
        fila_f.addWidget(self.fecha_desde)
        fila_f.addWidget(crear_boton_hoy(self.fecha_desde))
        fila_f.addWidget(QLabel("Hasta :"))
        self.fecha_hasta = QDateEdit()
        self.fecha_hasta.setCalendarPopup(True)
        self.fecha_hasta.setDate(hoy)
        fila_f.addWidget(self.fecha_hasta)
        fila_f.addWidget(crear_boton_hoy(self.fecha_hasta))
        layout.addWidget(self.grupo_fechas)

        # -- Filtro de Cliente (Uno/Zona/Provincia/Todos) -----------------
        self.grupo_cliente = QGroupBox("Selección de Clientes")
        fila_c = QHBoxLayout(self.grupo_cliente)
        self.grupo_botones_filtro = QButtonGroup(self)
        self.radio_uno = QRadioButton("Uno Sólo")
        self.radio_zona = QRadioButton("Por Zona")
        self.radio_provincia = QRadioButton("Por Provincia")
        self.radio_todos = QRadioButton("Todos")
        for r in (self.radio_uno, self.radio_zona, self.radio_provincia, self.radio_todos):
            self.grupo_botones_filtro.addButton(r)
            fila_c.addWidget(r)
        self.radio_todos.setChecked(True)

        # Ya no se tipea a mano — lo llena `_on_elegir_cliente()` (sigue
        # existiendo, oculto, sólo como "backing store" del código
        # elegido, `_filtro_cliente()` lo sigue leyendo igual).
        self.spin_clte = QSpinBox()
        self.spin_clte.setRange(0, 999999)
        self.spin_clte.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.spin_clte.setVisible(False)
        fila_c.addWidget(self.spin_clte)
        self.lbl_cliente_elegido = QLabel("(sin cliente elegido)")
        fila_c.addWidget(self.lbl_cliente_elegido)
        # Ventana de selección de Cliente (feedback del usuario,
        # 2026-08-18: "En los listados que pide un cliente, aparezca la
        # ventana de selección de cliente") — se abre sola al elegir
        # "Uno Sólo" (mismo criterio ya usado en Facturador/Recibo/
        # CtaCte), y también con este botón para volver a cambiarlo sin
        # tener que destildar/tildar el radio de nuevo.
        self.btn_elegir_cliente = QPushButton("Elegir Cliente...")
        self.btn_elegir_cliente.clicked.connect(self._on_elegir_cliente)
        fila_c.addWidget(self.btn_elegir_cliente)

        self.combo_zona = QComboBox()
        fila_c.addWidget(self.combo_zona)

        self.combo_provincia = QComboBox()
        for letra, nombre in NOMBRE_PROVINCIA.items():
            self.combo_provincia.addItem(f"{letra}-{nombre}", letra)
        fila_c.addWidget(self.combo_provincia)

        for r in (self.radio_uno, self.radio_zona, self.radio_provincia, self.radio_todos):
            r.toggled.connect(self._on_filtro_cliente_cambiado)
        layout.addWidget(self.grupo_cliente)

        # -- Vendedor / Sección -------------------------------------------
        fila_extra = QHBoxLayout()
        self.lbl_vendedor = QLabel("Vendedor :")
        self.combo_vendedor = QComboBox()
        fila_extra.addWidget(self.lbl_vendedor)
        fila_extra.addWidget(self.combo_vendedor)

        self.lbl_seccion = QLabel("Sección :")
        self.combo_seccion = QComboBox()
        self.combo_seccion.addItem("-- Todas --", None)
        fila_extra.addWidget(self.lbl_seccion)
        fila_extra.addWidget(self.combo_seccion)
        fila_extra.addStretch()
        layout.addLayout(fila_extra)

        layout.addStretch()

        fila_botones = QHBoxLayout()
        fila_botones.addStretch()
        self.btn_generar = QPushButton("Generar PDF")
        self.btn_generar.clicked.connect(self._on_generar)
        fila_botones.addWidget(self.btn_generar)
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

        self._on_filtro_cliente_cambiado()

    def _poblar_combos_tablas(self) -> None:
        for f in self.repos.fctablas().by_ctab("ZN"):
            cod = (f.COD or "").strip()
            if cod.isdigit():
                self.combo_zona.addItem(f"{cod}-{(f.DESCRI or '').strip()}", int(cod))
        for f in self.repos.fctablas().by_ctab("VD"):
            cod = (f.COD or "").strip()
            if cod.isdigit():
                self.combo_vendedor.addItem(f"{cod}-{(f.DESCRI or '').strip()}", int(cod))
        for f in self.repos.fctablas().by_ctab("SC"):
            cod = (f.COD or "").strip()
            self.combo_seccion.addItem(f"{cod}-{(f.DESCRI or '').strip()}", cod)

    # ------------------------------------------------------------------
    def _on_reporte_cambiado(self) -> None:
        reporte = self._reporte
        self.grupo_fechas.setVisible(reporte in USA_FECHAS)
        self.grupo_cliente.setVisible(reporte in USA_CLIENTE_FILTRO)
        self.radio_provincia.setVisible(reporte == "clientes")
        self.lbl_vendedor.setVisible(reporte in USA_VENDEDOR)
        self.combo_vendedor.setVisible(reporte in USA_VENDEDOR)
        self.lbl_seccion.setVisible(reporte in USA_SECCION)
        self.combo_seccion.setVisible(reporte in USA_SECCION)

    def _on_filtro_cliente_cambiado(self) -> None:
        es_uno = self.radio_uno.isChecked()
        self.lbl_cliente_elegido.setVisible(es_uno)
        self.btn_elegir_cliente.setVisible(es_uno)
        self.combo_zona.setVisible(self.radio_zona.isChecked())
        self.combo_provincia.setVisible(self.radio_provincia.isChecked())
        if es_uno and self.spin_clte.value() == 0:
            # Recién se tildó "Uno Sólo" y todavía no hay cliente
            # elegido — se abre el buscador directo, no hace falta
            # clickear "Elegir Cliente..." a mano primero.
            self._on_elegir_cliente()

    def _on_elegir_cliente(self) -> None:
        dialogo = ClienteBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.cliente_elegido is None:
            return
        self.spin_clte.setValue(dialogo.cliente_elegido.CODIGO)
        self.lbl_cliente_elegido.setText(
            f"{dialogo.cliente_elegido.CODIGO} — {(dialogo.cliente_elegido.NOMB or '').strip()}"
        )

    def _filtro_cliente(self) -> tuple[str, object]:
        if self.radio_uno.isChecked():
            return ListadosService.FILTRO_UNO, self.spin_clte.value()
        if self.radio_zona.isChecked():
            return ListadosService.FILTRO_ZONA, self.combo_zona.currentData()
        if self.radio_provincia.isChecked():
            return ListadosService.FILTRO_PROVINCIA, self.combo_provincia.currentData()
        return ListadosService.FILTRO_TODOS, None

    def _descripcion_filtro_cliente(self) -> str:
        """Texto de la selección de operador (pedido del usuario,
        2026-08-18: "en el centro abajo del titulo, la seleccion del
        operador (Cliente, zona, todos, etc.)") — se agrega al
        `subtitulo` de cada listado que filtra por cliente."""
        if self.radio_uno.isChecked():
            return f"Cliente: {self.lbl_cliente_elegido.text()}"
        if self.radio_zona.isChecked():
            return f"Zona: {self.combo_zona.currentText()}"
        if self.radio_provincia.isChecked():
            return f"Provincia: {self.combo_provincia.currentText()}"
        return "Todos los Clientes"

    # ------------------------------------------------------------------
    def _on_generar(self) -> None:
        reporte = self._reporte
        if reporte in VENDEDOR_OBLIGATORIO and self.combo_vendedor.currentData() is None:
            QMessageBox.warning(self, "Listados Varios", "Elegí un Vendedor.")
            return

        try:
            ruta = self._generar_reporte(reporte)
        except ValueError as exc:
            QMessageBox.warning(self, "Listados Varios", str(exc))
            return

        if ruta is None:
            QMessageBox.information(self, "Listados Varios", "No hay datos para los filtros elegidos.")
            return

        etiqueta = dict(REPORTES)[reporte]
        PdfPreviewDialog(ruta, titulo=f"Listado — {etiqueta}", parent=self).exec()

    def _rango_fechas(self) -> tuple[date, date]:
        return self.fecha_desde.date().toPyDate(), self.fecha_hasta.date().toPyDate()

    def _rango_exclusivo(self) -> tuple[date, date]:
        desde, hasta = self._rango_fechas()
        return desde, hasta + timedelta(days=1)

    def _generar_reporte(self, reporte: str):
        if reporte == "clientes":
            return self._pdf_clientes()
        if reporte == "precios":
            return self._pdf_precios()
        if reporte == "subdiario_ventas":
            return self._pdf_subdiario_ventas()
        if reporte == "subdiario_cobranzas":
            return self._pdf_subdiario_cobranzas()
        if reporte == "ingresos_brutos":
            return self._pdf_ingresos_brutos()
        if reporte in ("deuda_pendiente", "planilla_cobranzas"):
            return self._pdf_deuda_pendiente(reporte)
        if reporte == "estado_cuenta":
            return self._pdf_estado_cuenta()
        if reporte == "saldos":
            return self._pdf_saldos()
        if reporte == "percepciones_arba":
            return self._pdf_percepciones_arba()
        if reporte == "comisiones_cobranzas":
            return self._pdf_comisiones_cobranzas()
        if reporte == "subdiario_comisiones":
            return self._pdf_subdiario_comisiones()
        raise ValueError(f"Listado no reconocido: {reporte}")

    # ------------------------------------------------------------------
    def _pdf_clientes(self):
        filtro, valor = self._filtro_cliente()
        filas = self.servicio.listado_clientes(filtro, valor)
        if not filas:
            return None
        datos = [
            [str(c.CODIGO), (c.NOMB or "").strip(), (c.TEL1 or "").strip(),
             (c.DIR or "").strip(), (c.LOC or "").strip(), nombre_provincia(c.PCIA)]
            for c in filas
        ]
        return generar_pdf_listado(
            titulo="Clientes", subtitulo=f"{self._descripcion_filtro_cliente()} — {len(filas)} clientes",
            columnas=["Código", "Nombre", "Teléfono", "Dirección", "Localidad", "Provincia"],
            filas=datos, columnas_derecha=(0,), nombre_archivo="listado_clientes.pdf",
        )

    def _pdf_precios(self):
        cod_seccion = self.combo_seccion.currentData()
        filas = self.servicio.lista_precios(cod_seccion)
        if not filas:
            return None
        datos = [
            [(a.COD1 or "").strip(), (a.COD2 or "").strip(), (a.DESCRI or "").strip(),
             f"$ {format_decimal(a.PREC or Decimal('0'))}"]
            for a in filas
        ]
        return generar_pdf_listado(
            titulo="Lista de Precios", subtitulo=f"{len(filas)} artículos",
            columnas=["Sección", "Código", "Descripción", "Precio"],
            filas=datos, columnas_derecha=(3,), nombre_archivo="lista_precios.pdf",
        )

    def _pdf_subdiario_ventas(self):
        desde, hasta = self._rango_fechas()
        filas = self.servicio.subdiario_ventas(desde, hasta)
        if not filas:
            return None
        datos = []
        total = Decimal("0")
        for f in filas:
            neto = (f.GRINS or Decimal("0")) + (f.IVAINS or Decimal("0"))
            total += neto
            datos.append([
                f.FECHA.strftime("%d/%m/%Y") if f.FECHA else "",
                ETIQUETAS_TIPO_CTASCTE.get(int(f.TIPO) if (f.TIPO or "").strip().isdigit() else -1, "--"),
                f"{(f.LETRA or '').strip()} {(f.PTOVTA or 0):04d}-{(f.CPBTE or 0):08d}",
                str(f.CLTE or ""), (f.NOMB or "").strip(),
                format_decimal(f.GRINS or Decimal("0")), format_decimal(f.IVAINS or Decimal("0")),
                format_decimal(neto),
            ])
        return generar_pdf_listado(
            titulo="Subdiario de Ventas", subtitulo=f"{desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}",
            columnas=["Fecha", "Tipo", "Comprobante", "Cód.", "Cliente", "Gravado", "IVA", "Total"],
            filas=datos, columnas_derecha=(0, 3, 5, 6, 7),
            pie=[("Total", f"$ {format_decimal(total)}")], nombre_archivo="subdiario_ventas.pdf",
        )

    def _pdf_subdiario_cobranzas(self):
        desde, hasta_excl = self._rango_exclusivo()
        movimientos = self.servicio.subdiario_cobranzas(desde, hasta_excl)
        if not movimientos:
            return None
        datos = []
        total = Decimal("0")
        for m in movimientos:
            importe = m.IMPTE or Decimal("0")
            total += importe
            cliente = self.repos.cliente().by_codigo(m.CLTE) if m.CLTE else None
            datos.append([
                m.FECHA.strftime("%d/%m/%Y") if m.FECHA else "",
                f"{(m.PREFIJO or 0):04d}-{(m.CPBTE or 0):08d}",
                str(m.CLTE or ""), (cliente.NOMB or "").strip() if cliente else "",
                format_decimal(importe),
            ])
        return generar_pdf_listado(
            titulo="Subdiario de Cobranzas", subtitulo=f"{desde.strftime('%d/%m/%Y')} al {self.fecha_hasta.date().toPyDate().strftime('%d/%m/%Y')}",
            columnas=["Fecha", "Recibo", "Cód.", "Cliente", "Importe"],
            filas=datos, columnas_derecha=(0, 2, 4),
            pie=[("Total", f"$ {format_decimal(total)}")], nombre_archivo="subdiario_cobranzas.pdf",
        )

    def _pdf_ingresos_brutos(self):
        desde, hasta_excl = self._rango_exclusivo()
        filas = self.servicio.ingresos_brutos(desde, hasta_excl)
        if not filas:
            return None
        datos = [
            [f.pcia_nombre, ETIQUETAS_TIPO_CTASCTE.get(f.tipo, "--"),
             format_decimal(f.gravado), format_decimal(f.iva), format_decimal(f.total)]
            for f in filas
        ]
        total = sum((f.total for f in filas), Decimal("0"))
        return generar_pdf_listado(
            titulo="Ingresos Brutos", subtitulo=f"{desde.strftime('%d/%m/%Y')} al {self.fecha_hasta.date().toPyDate().strftime('%d/%m/%Y')}",
            columnas=["Provincia", "Tipo", "Gravado", "IVA", "Total"],
            filas=datos, columnas_derecha=(2, 3, 4),
            pie=[("Total", f"$ {format_decimal(total)}")], nombre_archivo="ingresos_brutos.pdf",
        )

    def _pdf_deuda_pendiente(self, reporte: str):
        filtro, valor = self._filtro_cliente()
        grupos = self.servicio.deuda_pendiente(filtro, valor)
        if not grupos:
            return None
        datos = []
        total = Decimal("0")
        for cliente, pendientes in grupos:
            for p in pendientes:
                debe = p.DEBE or Decimal("0")
                total += debe
                datos.append([
                    str(cliente.CODIGO), (cliente.NOMB or "").strip(),
                    ETIQUETAS_TIPO_CTASCTE.get(p.TIPO, "--"), str(p.CPBTE or ""),
                    p.FECHA.strftime("%d/%m/%Y") if p.FECHA else "",
                    p.FECVTO.strftime("%d/%m/%Y") if p.FECVTO else "",
                    format_decimal(debe),
                ])
        titulo = "Deuda Pendiente" if reporte == "deuda_pendiente" else "Planilla de Cobranzas"
        return generar_pdf_listado(
            titulo=titulo, subtitulo=f"{self._descripcion_filtro_cliente()} — {len(grupos)} clientes con deuda",
            columnas=["Cód.", "Cliente", "Tipo", "Cpbte.", "Fecha", "Vto.", "Debe"],
            filas=datos, columnas_derecha=(0, 3, 6),
            pie=[("Total", f"$ {format_decimal(total)}")], nombre_archivo=f"{reporte}.pdf",
        )

    def _pdf_estado_cuenta(self):
        filtro, valor = self._filtro_cliente()
        grupos = self.servicio.estado_cuenta(filtro, valor)
        if not grupos:
            return None
        datos = []
        for cliente, movimientos in grupos:
            saldo = Decimal("0")
            for m in movimientos:
                importe = m.IMPTE or Decimal("0")
                saldo += importe if m.TIPO in (0, 1, 3, 7) else -importe
                datos.append([
                    str(cliente.CODIGO), (cliente.NOMB or "").strip(),
                    m.FECHA.strftime("%d/%m/%Y") if m.FECHA else "",
                    ETIQUETAS_TIPO_CTASCTE.get(m.TIPO, "--"), str(m.CPBTE or ""),
                    format_decimal(importe), format_decimal(saldo),
                ])
        return generar_pdf_listado(
            titulo="Estado de Cuenta", subtitulo=f"{self._descripcion_filtro_cliente()} — {len(grupos)} clientes",
            columnas=["Cód.", "Cliente", "Fecha", "Tipo", "Cpbte.", "Importe", "Saldo"],
            filas=datos, columnas_derecha=(0, 4, 5, 6), nombre_archivo="estado_cuenta.pdf",
        )

    def _pdf_saldos(self):
        filtro, valor = self._filtro_cliente()
        filas = self.servicio.saldos_cta_cte(filtro, valor)
        if not filas:
            return None
        datos = [[str(f["codigo"]), f["nombre"], format_decimal(f["saldo"])] for f in filas]
        total = sum((f["saldo"] for f in filas), Decimal("0"))
        return generar_pdf_listado(
            titulo="Saldos de Cta. Cte.", subtitulo=f"{self._descripcion_filtro_cliente()} — {len(filas)} clientes",
            columnas=["Cód.", "Cliente", "Saldo"],
            filas=datos, columnas_derecha=(0, 2),
            pie=[("Total", f"$ {format_decimal(total)}")], nombre_archivo="saldos_ctacte.pdf",
        )

    def _pdf_percepciones_arba(self):
        desde, hasta_excl = self._rango_exclusivo()
        resultado = self.servicio.percepciones_arba(desde, hasta_excl)
        if not resultado.filas:
            return None
        datos = [
            [f.fecha.strftime("%d/%m/%Y") if f.fecha else "", f.letra, f.comprobante, f.nombre, f.cuit,
             f.civa_label, format_decimal(f.imponible), format_decimal(f.percepcion)]
            for f in resultado.filas
        ]
        ruta = generar_pdf_listado(
            titulo="Percepciones ARBA", subtitulo=f"{desde.strftime('%d/%m/%Y')} al {self.fecha_hasta.date().toPyDate().strftime('%d/%m/%Y')}",
            columnas=["Fecha", "L.", "Comprobante", "Cliente", "CUIT", "IVA", "Imponible", "Percepción"],
            filas=datos, columnas_derecha=(0, 2, 6, 7),
            pie=[("Total Imponible", f"$ {format_decimal(resultado.total_imponible)}"),
                 ("Total Percepción", f"$ {format_decimal(resultado.total_percepcion)}")],
            nombre_archivo="percepciones_arba.pdf",
        )
        # Archivo de presentación .txt — réplica literal del formato del
        # legacy (ver ListadosService.percepciones_arba, alcance
        # confirmado con el usuario 2026-08-16: vigencia ante ARBA no
        # verificada, confirmar con contador antes de usar para una
        # presentación real).
        ruta_txt = DIR_SALIDA_DEFAULT / resultado.nombre_archivo_txt
        ruta_txt.parent.mkdir(parents=True, exist_ok=True)
        ruta_txt.write_text(resultado.contenido_txt, encoding="latin-1", errors="replace")
        return ruta

    def _pdf_comisiones_cobranzas(self):
        vend = self.combo_vendedor.currentData()
        desde, hasta_excl = self._rango_exclusivo()
        resultado = self.servicio.comisiones_cobranzas(vend, desde, hasta_excl)
        if not resultado.filas:
            return None
        datos = [
            [str(f.clte), f.nombre, str(f.cpbte), f.fecha.strftime("%d/%m/%Y") if f.fecha else "",
             format_decimal(f.importe), "ND Rechazado" if f.es_nd_rechazada else ""]
            for f in resultado.filas
        ]
        return generar_pdf_listado(
            titulo="Comisiones por Cobranzas", subtitulo=self.combo_vendedor.currentText(),
            columnas=["Cód.", "Cliente", "Cpbte.", "Fecha", "Importe", "Obs."],
            filas=datos, columnas_derecha=(0, 2, 4),
            pie=[
                ("Total Cobrado", f"$ {format_decimal(resultado.total_cobrado)}"),
                ("ND Rechazadas", f"{resultado.total_nd_rechazadas} — $ {format_decimal(resultado.total_nd_rechazadas_importe)}"),
                ("Neto (base comisión)", f"$ {format_decimal(resultado.neto)}"),
            ],
            nombre_archivo="comisiones_cobranzas.pdf",
        )

    def _pdf_subdiario_comisiones(self):
        desde, hasta_excl = self._rango_exclusivo()
        vend = self.combo_vendedor.currentData()
        filas = self.servicio.subdiario_ventas_comisiones(desde, hasta_excl, vend)
        if not filas:
            return None
        datos = [
            [f.fecha.strftime("%d/%m/%Y") if f.fecha else "", f.letra, f.comprobante, str(f.clte or ""),
             f.nombre, f.civa_label, format_decimal(f.gravado), format_decimal(f.exento),
             format_decimal(f.iva), format_decimal(f.percepcion_ib), format_decimal(f.total)]
            for f in filas
        ]
        total = sum((f.total for f in filas), Decimal("0"))
        return generar_pdf_listado(
            titulo="Subdiario Vtas. (Comisiones)",
            subtitulo=f"{desde.strftime('%d/%m/%Y')} al {self.fecha_hasta.date().toPyDate().strftime('%d/%m/%Y')}"
            + (f" — {self.combo_vendedor.currentText()}" if vend else " — Todos los Vendedores"),
            columnas=["Fecha", "L.", "Comprobante", "Cód.", "Cliente", "IVA", "Gravado", "Exento", "IVA $", "Perc.IB", "Total"],
            filas=datos, columnas_derecha=(0, 3, 6, 7, 8, 9, 10),
            pie=[("Total", f"$ {format_decimal(total)}")], nombre_archivo="subdiario_ventas_comisiones.pdf",
        )

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
