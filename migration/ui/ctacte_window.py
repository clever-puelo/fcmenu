"""`CtaCteWindow` — Consulta de Cuenta Corriente, migración de
`CtaCte.frm` (2521 líneas).

**Arquitectura confirmada con el usuario**: mismo criterio ya aplicado a
Facturador/Recibo — una sola ventana cohesiva (extracto + panel lateral
+ acciones), no los paneles/`PictureBox` superpuestos que usa el legacy
para simular varias pantallas dentro de una sola.

**4 sub-vistas del legacy, todas cubiertas**:
1. Extracto anual (`Sub DoVer3`/`CargaGrilla`) — grilla principal.
2. "Ver Cheques" (`Sub DoVer9`) — `ChequesClienteDialog`.
3. "Ver Saldos" (`Sub DoVer8`) — `SaldosClientesDialog`, permite además
   cambiar de cliente con un clic (cierra el diálogo y vuelve acá ya con
   ese cliente cargado).
4. Clic en una fila del extracto → drill-down a factura/NC/ND
   (`Sub MuestraDetalle`) — `FacturaEmitidaDetalleDialog` (mismo módulo
   que usa "Facturas Emitidas", con Imprimir), con los renglones reales
   ya incluidos (no hace falta un segundo clic "Ver Items").

**Panel lateral**: `CtaCte.frm` declaraba 6 campos (Contacto/Teléfono/
Fec. Alta/Val.Pend./Últ.Vta.Ctd./Últ.Vta.C.Cte./Sdo.Venc.) que ningún
código del proyecto poblaba jamás (confirmado con grep completo). El
usuario confirmó (2026-08-16) implementar los 5 que sí son calculables
con datos reales — ver `CuentaCorrienteService.resumen_cliente()` — y
descartar sólo "Contacto" (sin fuente de datos, mismo criterio que los
controles fantasma de `Abmclte.frm`). "Últ.Vta.Ctd."/"Últ.Vta.C.Cte." se
unificaron en un solo campo "Última Venta" (no hay forma confiable de
clasificar Contado/Cta.Cte en los datos reales de Cond. de Venta).

**Saldo Inicial de cada año reconciliado con "mostrar todo"** — el
legacy usaba ahí una fórmula distinta e incompleta a la del saldo
corriente; corregido tras confirmar con el usuario (2026-08-16, cuenta
407 real) que la discrepancia entre filtrar por año y mostrar toda la
historia no era aceptable. Ver docstring de
`CtascteRepository.saldo_inicial()`.
"""

from __future__ import annotations

from datetime import date

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.decimals import format_decimal
from migration.models import Cliente
from migration.repository import RepositoryFactory
from migration.services import CuentaCorrienteService, FilaExtracto

from .cheques_cliente_dialog import ChequesClienteDialog
from .cliente_busqueda_window import ClienteBusquedaWindow
from .factura_emitida_detalle_dialog import FacturaEmitidaDetalleDialog
from .nota_cliente_dialog import NotaClienteDialog
from .saldos_clientes_dialog import SaldosClientesDialog
from .widgets import redimensionar_pct_pantalla

ANIO_MINIMO = 2005  # CtaCte.frm Sub CargaAnos: "Ano1 = 2005"

COL_FECHA = 0
COL_TIPO = 1
COL_CPBTE = 2
COL_IMPUT1 = 3
COL_IMPUT2 = 4
COL_IMPUT3 = 5
COL_IMPUT4 = 6
COL_RESTO = 7
COL_DEBE = 8
COL_HABER = 9
COL_SALDO = 10
COLUMNAS_EXTRACTO = [
    "Fecha", "Tipo", "Comprobante", "Im.1", "Im.2", "Im.3", "Im.4", "Resto", "Debe", "Haber", "Saldo",
]
COLUMNAS_ALINEADAS_DERECHA = (COL_CPBTE, COL_RESTO, COL_DEBE, COL_HABER, COL_SALDO)
COLUMNAS_ALINEADAS_CENTRO = (COL_IMPUT1, COL_IMPUT2, COL_IMPUT3, COL_IMPUT4)

# TIPO de comprobante que puede tener cabecera en FCIVAVTA (Facturas/NC/ND
# — Recibos y el resto de los movimientos de Ctasctes no la tienen, un
# clic ahí no encuentra nada, igual que CtaCte.frm MuestraDetalle cuando
# RgFCIVA.EOF).
TIPOS_CON_DRILLDOWN = {1, 2, 3}


class CtaCteWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Cuentas Corrientes")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        redimensionar_pct_pantalla(self, 95, 90)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.cc_service = CuentaCorrienteService(self.db)

        self.cliente_actual: Cliente | None = None
        self._filas_extracto: list[FilaExtracto] = []

        self._construir_ui()
        # Primero se muestra la ventana, recién después se pide el
        # Cliente (feedback del usuario, 2026-08-18, segunda ronda) —
        # mismo criterio ya aplicado en Facturador/Recibo:
        # `QTimer.singleShot(0, ...)` deja que esta ventana se muestre
        # primero; llamado directo acá (dentro de `__init__`) el
        # buscador modal aparecía ANTES que la propia ventana.
        QTimer.singleShot(0, self._on_otro_cliente)

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_cabecera())

        cuerpo = QHBoxLayout()
        cuerpo.addWidget(self._armar_extracto(), stretch=1)
        cuerpo.addWidget(self._armar_panel_lateral())
        layout.addLayout(cuerpo, stretch=1)

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QWidget:
        contenedor = QWidget()
        fila = QHBoxLayout(contenedor)

        self.lbl_cliente = QLabel("(sin cliente elegido)")
        self.lbl_cliente.setStyleSheet("font-weight: bold;")
        fila.addWidget(self.lbl_cliente, stretch=1)

        self.btn_nota_cliente = QPushButton("Nota Clte.")
        self.btn_nota_cliente.setEnabled(False)
        self.btn_nota_cliente.clicked.connect(self._on_nota_cliente)
        fila.addWidget(self.btn_nota_cliente)

        fila.addWidget(QLabel("Desde:"))
        self.combo_anio = QComboBox()
        anio_actual = date.today().year
        for anio in range(ANIO_MINIMO, anio_actual + 1):
            self.combo_anio.addItem(str(anio), anio)
        # CtaCte.frm Sub CargaAnos: `Combo1.ListIndex = A1 - 5` -> el año
        # default seleccionado es siempre "año actual - 4" (ver docstring
        # del módulo/análisis previo), acotado a ANIO_MINIMO si el
        # historial todavía no llega tan atrás.
        anio_default = max(ANIO_MINIMO, anio_actual - 4)
        self.combo_anio.setCurrentIndex(anio_default - ANIO_MINIMO)
        self.combo_anio.currentIndexChanged.connect(self._cargar_extracto)
        fila.addWidget(self.combo_anio)

        self.btn_ver_cheques = QPushButton("Ver Cheques")
        self.btn_ver_cheques.clicked.connect(self._on_ver_cheques)
        fila.addWidget(self.btn_ver_cheques)

        btn_ver_saldos = QPushButton("Ver Saldos")
        btn_ver_saldos.clicked.connect(self._on_ver_saldos)
        fila.addWidget(btn_ver_saldos)

        btn_otro_cliente = QPushButton("Otro Cliente")
        btn_otro_cliente.clicked.connect(self._on_otro_cliente)
        fila.addWidget(btn_otro_cliente)

        # No tenía botón Cerrar (feedback del usuario, 2026-08-18).
        fila.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila.addWidget(btn_cerrar)

        return contenedor

    # ------------------------------------------------------------------
    # Extracto
    # ------------------------------------------------------------------
    def _armar_extracto(self) -> QGroupBox:
        grupo = QGroupBox("Saldos de Cuenta Corriente")
        layout = QVBoxLayout(grupo)

        self.tabla_extracto = QTableWidget(0, len(COLUMNAS_EXTRACTO))
        self.tabla_extracto.setHorizontalHeaderLabels(COLUMNAS_EXTRACTO)
        self.tabla_extracto.verticalHeader().setVisible(False)
        # Renglones más cortos para que entren más por pantalla (feedback
        # del usuario, 2026-08-18, tercera ronda) — mismo criterio ya
        # usado en `TablaBusqueda` (24px), acá un poco más chico todavía
        # porque esta grilla no es de búsqueda sino de extracto largo.
        self.tabla_extracto.verticalHeader().setDefaultSectionSize(22)
        self.tabla_extracto.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabla_extracto.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabla_extracto.itemClicked.connect(self._on_click_fila_extracto)
        layout.addWidget(self.tabla_extracto, stretch=1)

        return grupo

    def _cargar_extracto(self) -> None:
        self._filas_extracto = []
        self.tabla_extracto.setRowCount(0)
        if self.cliente_actual is None:
            self._refrescar_panel_totales(None)
            return

        anio = self.combo_anio.currentData()
        extracto = self.cc_service.extracto_anual(self.cliente_actual.CODIGO, anio)
        self._filas_extracto = extracto.filas

        for fila_idx, f in enumerate(extracto.filas):
            self.tabla_extracto.insertRow(fila_idx)
            valores = [
                f.fecha.strftime("%d/%m/%Y") if f.fecha else "",
                f.tipo_label,
                str(f.cpbte) if f.cpbte is not None else "",
                f.imput1, f.imput2, f.imput3, f.imput4,
                format_decimal(f.resto) if f.resto is not None else "---",
                format_decimal(f.debe) if f.debe else "",
                format_decimal(f.haber) if f.haber else "",
                format_decimal(f.saldo),
            ]
            for col, texto in enumerate(valores):
                item = QTableWidgetItem(texto)
                if col in COLUMNAS_ALINEADAS_DERECHA:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                elif col in COLUMNAS_ALINEADAS_CENTRO:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabla_extracto.setItem(fila_idx, col, item)
        self.tabla_extracto.resizeColumnsToContents()
        if self.tabla_extracto.rowCount():
            self.tabla_extracto.scrollToBottom()

        self._refrescar_panel_totales(extracto)

    def _on_click_fila_extracto(self, item: QTableWidgetItem) -> None:
        fila = item.row()
        if fila < 0 or fila >= len(self._filas_extracto):
            return
        info = self._filas_extracto[fila]
        if info.tipo not in TIPOS_CON_DRILLDOWN or info.cpbte is None:
            return
        factura = self.repos.fciva_vta().by_comprobante(
            str(info.tipo), info.letra or "", info.prefijo or 0, info.cpbte
        )
        if factura is None:
            return
        # Mismo módulo de drill-down que "Facturas Emitidas" (pedido del
        # usuario, 2026-08-19: "En la consulta de cuenta corriente se
        # pueda hacer lo mismo" — cabecera compacta, detalle grande con
        # scroll, botón Imprimir) — antes tenía su propio
        # `FacturaDetalleDialog`, más simple, ahora unificado en una sola
        # pantalla para no mantener 2 vistas de lo mismo.
        FacturaEmitidaDetalleDialog(self.repos, factura, parent=self).exec()

    # ------------------------------------------------------------------
    # Panel lateral
    # ------------------------------------------------------------------
    def _armar_panel_lateral(self) -> QGroupBox:
        grupo = QGroupBox("Cliente")
        grupo.setMinimumWidth(260)
        form = QFormLayout(grupo)

        self.lbl_telefono = QLabel("—")
        self.lbl_fecha_alta = QLabel("—")
        self.lbl_valores_pend = QLabel("—")
        self.lbl_ultima_venta = QLabel("—")
        self.lbl_saldo_venc = QLabel("—")
        form.addRow("Teléfono :", self.lbl_telefono)
        form.addRow("Fec. Alta :", self.lbl_fecha_alta)
        form.addRow("Val. Pend. :", self.lbl_valores_pend)
        form.addRow("Última Venta :", self.lbl_ultima_venta)
        form.addRow("Sdo. Venc. :", self.lbl_saldo_venc)

        self.lbl_tot_debe = QLabel("$ 0,00")
        self.lbl_tot_haber = QLabel("$ 0,00")
        self.lbl_saldo = QLabel("$ 0,00")
        self.lbl_saldo.setStyleSheet("font-weight: bold;")
        form.addRow("Tot. Debe :", self.lbl_tot_debe)
        form.addRow("Tot. Haber :", self.lbl_tot_haber)
        form.addRow("SALDO :", self.lbl_saldo)

        return grupo

    def _refrescar_panel_totales(self, extracto) -> None:
        if extracto is None:
            self.lbl_tot_debe.setText("$ 0,00")
            self.lbl_tot_haber.setText("$ 0,00")
            self.lbl_saldo.setText("$ 0,00")
            return
        self.lbl_tot_debe.setText(f"$ {format_decimal(extracto.total_debe)}")
        self.lbl_tot_haber.setText(f"$ {format_decimal(extracto.total_haber)}")
        self.lbl_saldo.setText(f"$ {format_decimal(extracto.saldo_final)}")

    def _refrescar_resumen_cliente(self) -> None:
        if self.cliente_actual is None:
            self.lbl_telefono.setText("—")
            self.lbl_fecha_alta.setText("—")
            self.lbl_valores_pend.setText("—")
            self.lbl_ultima_venta.setText("—")
            self.lbl_saldo_venc.setText("—")
            return
        resumen = self.cc_service.resumen_cliente(self.cliente_actual)
        self.lbl_telefono.setText(resumen.telefono or "—")
        self.lbl_fecha_alta.setText(resumen.fecha_alta.strftime("%d/%m/%Y") if resumen.fecha_alta else "—")
        self.lbl_valores_pend.setText(f"$ {format_decimal(resumen.valores_pendientes)}")
        self.lbl_ultima_venta.setText(resumen.ultima_venta.strftime("%d/%m/%Y") if resumen.ultima_venta else "—")
        self.lbl_saldo_venc.setText(f"$ {format_decimal(resumen.saldo_vencido)}")

    # ------------------------------------------------------------------
    # Cliente
    # ------------------------------------------------------------------
    def _on_otro_cliente(self) -> None:
        dialogo = ClienteBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.cliente_elegido is None:
            return
        # `dialogo.cliente_elegido` pertenece a la sesión propia de
        # ClienteBusquedaWindow — se re-busca en self.db, mismo criterio
        # ya aplicado en Facturador/Recibo.
        self.cliente_actual = self.repos.cliente().by_codigo(dialogo.cliente_elegido.CODIGO)
        self._refrescar_cliente()

    def _refrescar_cliente(self) -> None:
        cliente = self.cliente_actual
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.btn_nota_cliente.setEnabled(False)
            self.btn_nota_cliente.setStyleSheet("")
            self._cargar_extracto()
            self._refrescar_resumen_cliente()
            return

        self.lbl_cliente.setText(f"{cliente.CODIGO}  {(cliente.NOMB or '').strip()}")
        self.btn_nota_cliente.setEnabled(True)

        # Auto-popup de Nota Cliente — réplica de CtaCte.frm Form_Activate
        # (líneas 1673-1690), mismo criterio ya aplicado a Facturador/Recibo.
        self._actualizar_boton_nota_cliente()
        if self._tiene_nota_cliente:
            NotaClienteDialog(self.repos, cliente.CODIGO, parent=self).exec()
            self._actualizar_boton_nota_cliente()

        self._cargar_extracto()
        self._refrescar_resumen_cliente()

    def _actualizar_boton_nota_cliente(self) -> None:
        self._tiene_nota_cliente = self.cliente_actual is not None and (
            self.repos.notaclte().by_cliente(self.cliente_actual.CODIGO) is not None
        )
        self.btn_nota_cliente.setStyleSheet(
            "background-color: #ffd54f; font-weight: bold;" if self._tiene_nota_cliente else ""
        )

    def _on_nota_cliente(self) -> None:
        if self.cliente_actual is None:
            return
        NotaClienteDialog(self.repos, self.cliente_actual.CODIGO, parent=self).exec()
        self._actualizar_boton_nota_cliente()

    # ------------------------------------------------------------------
    # Ver Cheques / Ver Saldos
    # ------------------------------------------------------------------
    def _on_ver_cheques(self) -> None:
        if self.cliente_actual is None:
            return
        ChequesClienteDialog(self.repos, self.cliente_actual.CODIGO, parent=self).exec()

    def _on_ver_saldos(self) -> None:
        dialogo = SaldosClientesDialog(self.cc_service, parent=self)
        dialogo.exec()
        if dialogo.codigo_elegido is None:
            return
        self.cliente_actual = self.repos.cliente().by_codigo(dialogo.codigo_elegido)
        self._refrescar_cliente()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
