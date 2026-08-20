"""`NotaCreditoConceptoWindow` — Nota de Crédito/Débito de "Concepto
Libre", migración de `CabFact.frm` (TipoFac=2/3, Motivo != 1) +
`DetNC.frm` + `PieFact.frm` + `EmiFact.frm` (rama común TipoFac 2/3) +
`ImputFC.frm` (sólo Nota de Crédito).

**Alcance confirmado con el usuario (2026-08-19)**: el legacy usa la
MISMA Cabecera/Pie para dos circuitos de Detalle bien distintos, según
el Motivo elegido (`CabFact.frm Sub Combo5_Click`, líneas 726-746):
Motivo=1 ("DEV.MERC.", único código real con ese valor, confirmado
contra `Fctabla1`) abre `DetFact.Show` (grilla de artículos + reversa
de Stock) — pantalla aparte, `NotaCreditoMercaderiaWindow`, todavía no
migrada. Cualquier otro Motivo abre `DetNC.Show` (3 renglones de texto
libre + importe + Con/Sin IVA) — lo que migra ESTA ventana. El combo de
Motivo acá excluye "1" a propósito, ese código vive sólo en la otra
pantalla.

**Circuito**:
1. Cabecera: Tipo (Nota de Crédito / Nota de Débito, `CabFact.frm
   Option1` índices 2/3), Cliente (mismo buscador que Facturador/
   Recibo), Motivo (`Fctabla1` CTAB='MT', sin el código 1).
2. Sólo para Nota de Crédito: elegir el comprobante (Factura/ND con
   deuda real) contra el que se imputa — réplica de `ImputFC.frm`
   (`CuentaCorrienteService.facturas_pendientes()`, ya usado por
   Recibo). Nota de Débito no imputa nada, sólo suma a la deuda
   (`Graba()` línea 2244).
3. Conceptos: hasta 3 líneas de texto libre + importe + Con/Sin IVA.
4. Pie: Base Gravada/IVA/No Gravado/TOTAL en vivo
   (`NotaCreditoConceptoService.calcular_total`) — ver esa clase para
   el bug real de "SIN IVA" del legacy, corregido acá.
5. Emitir: pide el CAE (mismo `NumeracionYCAEProvider` que Factura) y
   recién con el CAE aprobado persiste todo
   (`EmisionNotaCreditoService.emitir_concepto`). Si AFIP rechaza, NO se
   persiste nada.

**Fuera de alcance de esta primera versión** (a diferencia de Factura/
Recibo): sin vista previa en PDF antes de grabar — se confirma con un
diálogo simple. Se puede agregar en una ronda posterior si hace falta,
no bloquea el circuito de AFIP/Cta.Cte., que es lo que importa
validar primero.
"""

from __future__ import annotations

import os
from datetime import date
from decimal import Decimal
from typing import Optional

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from migration.afip import (
    NumeracionYCAEProvider,
    ResultadoCAE,
    codigo_afip,
    condicion_iva_receptor,
    crear_afip_provider,
    punto_venta_por_tipo,
)
from migration.db import get_session
from migration.models import Cliente, Ctascte
from migration.repository import ETIQUETAS_TIPO_CTASCTE, RepositoryFactory
from migration.services import (
    ConceptoNotaCredito,
    CuentaCorrienteService,
    EmisionNotaCreditoService,
    FacturaService,
    NotaCreditoConceptoService,
)

from .cliente_busqueda_window import ClienteBusquedaWindow
from .comprobante_aplicar_dialog import ComprobanteAplicarDialog
from .decimals import format_decimal
from .nota_cliente_dialog import NotaClienteDialog
from .widgets import (
    MontoLineEdit,
    UpperCaseLineEdit,
    crear_recuadro_destacado,
    redimensionar_pct_pantalla,
    texto_contacto_cliente,
)

CUIT_EMISOR_DEFAULT = "33703467909"  # ídem FacturadorWindow, ver ese docstring

TIPO_NC = 2
TIPO_ND = 3


class NotaCreditoConceptoWindow(QMainWindow):
    TEXTO_BTN_ELEGIR_CLIENTE = "Elegir Cliente..."
    TEXTO_BTN_CAMBIAR_CLIENTE = "Cambiar Cliente"

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        afip: Optional[NumeracionYCAEProvider] = None,
        tipo_inicial: int = TIPO_NC,
    ):
        super().__init__(parent)
        titulo = "Nota de Débito" if tipo_inicial == TIPO_ND else "Nota de Crédito"
        self.setWindowTitle(f"{titulo} — Concepto Libre")
        redimensionar_pct_pantalla(self, 75, 70)
        self._tipo_inicial = tipo_inicial

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.factura_service = FacturaService(self.db)
        self.nc_service = NotaCreditoConceptoService(self.db)
        self.emision_service = EmisionNotaCreditoService(self.db)
        self.cuenta_corriente_service = CuentaCorrienteService(self.db)
        self.afip: NumeracionYCAEProvider = afip if afip is not None else crear_afip_provider()

        self.cliente_actual: Optional[Cliente] = None
        self.tiene_nota_cliente = False
        self._usuario = os.environ.get("USERNAME", "SISTEMA")[:6]
        # Comprobante elegido en la ventana flotante "Aplicar a" (ver
        # `ComprobanteAplicarDialog`) — reemplaza la tabla completa que
        # antes vivía siempre visible (feedback del usuario, 2026-08-19).
        self._comprobante_seleccionado: Optional[Ctascte] = None

        self._construir_ui()
        self._poblar_motivos()
        self._nueva()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_cabecera())
        layout.addWidget(self._armar_comprobante_imputar())
        layout.addWidget(self._armar_conceptos(), stretch=1)
        layout.addWidget(self._armar_pie())

    # ------------------------------------------------------------------
    # Cabecera
    # ------------------------------------------------------------------
    def _armar_cabecera(self) -> QGroupBox:
        grupo = QGroupBox("Cabecera")
        layout = QVBoxLayout(grupo)

        # "Nota Clte."/"Nueva" subidos a esta línea, junto a los radios
        # de Tipo (feedback del usuario, 2026-08-19: "subir los botones
        # nota clte y nueva a la linea de NCredito y nDeb.") — el lugar
        # que dejan libre en `fila_cliente` lo ocupa "Aplicar a".
        fila_tipo = QHBoxLayout()
        self.radio_nc = QRadioButton("Nota de Crédito")
        self.radio_nd = QRadioButton("Nota de Débito")
        if self._tipo_inicial == TIPO_ND:
            self.radio_nd.setChecked(True)
        else:
            self.radio_nc.setChecked(True)
        grupo_radios = QButtonGroup(self)
        grupo_radios.addButton(self.radio_nc)
        grupo_radios.addButton(self.radio_nd)
        self.radio_nc.toggled.connect(self._on_tipo_cambiado)
        fila_tipo.addWidget(self.radio_nc)
        fila_tipo.addWidget(self.radio_nd)
        fila_tipo.addStretch()
        # Cambiar-cliente antes de "Notas" (pedido del usuario,
        # 2026-08-20) — esta ventana ya tenía "Notas"/"Nueva" acá arriba
        # (ronda anterior, 2026-08-19), así que el botón de cliente se
        # suma a esta misma línea en vez de a `fila_cliente`.
        self.btn_elegir_cliente = QPushButton(self.TEXTO_BTN_ELEGIR_CLIENTE)
        self.btn_elegir_cliente.clicked.connect(self._on_elegir_cliente)
        fila_tipo.addWidget(self.btn_elegir_cliente)
        self.btn_nota_cliente = QPushButton("Nota Clte.")
        self.btn_nota_cliente.setEnabled(False)
        self.btn_nota_cliente.clicked.connect(self._on_nota_cliente)
        fila_tipo.addWidget(self.btn_nota_cliente)
        self.btn_nueva = QPushButton("Nueva")
        self.btn_nueva.clicked.connect(self._nueva)
        fila_tipo.addWidget(self.btn_nueva)
        layout.addLayout(fila_tipo)

        # Código+nombre a la izquierda; "Aplicar a" abre la ventana
        # flotante de selección de Factura/ND a cancelar
        # (`ComprobanteAplicarDialog`), sólo visible/habilitado para Nota
        # de Crédito (Nota de Débito no imputa nada, ver docstring del
        # módulo).
        fila_cliente = QHBoxLayout()
        self.lbl_cliente = QLabel("(sin cliente elegido)")
        fila_cliente.addWidget(self.lbl_cliente, stretch=1)
        self.btn_aplicar_a = QPushButton("Aplicar a...")
        self.btn_aplicar_a.setEnabled(False)
        self.btn_aplicar_a.clicked.connect(self._on_aplicar_a)
        fila_cliente.addWidget(self.btn_aplicar_a)
        layout.addLayout(fila_cliente)

        # Localidad/Teléfono/Email del cliente elegido (ídem).
        self.lbl_cliente_contacto = QLabel("—")
        layout.addWidget(self.lbl_cliente_contacto)

        fila_datos = QHBoxLayout()
        fila_datos.addWidget(QLabel("Motivo :"))
        self.combo_motivo = QComboBox()
        self.combo_motivo.currentIndexChanged.connect(self._recalcular_totales)
        fila_datos.addWidget(self.combo_motivo, stretch=1)
        fila_datos.addStretch()
        self.lbl_proximo_numero = QLabel("Próx. Nº: —")
        fila_datos.addWidget(self.lbl_proximo_numero)
        layout.addLayout(fila_datos)

        return grupo

    def _poblar_motivos(self) -> None:
        self.combo_motivo.clear()
        for fila in self.repos.fctablas().by_ctab("MT"):
            cod = (fila.COD or "").strip()
            # Motivo "1" (DEV.MERC.) es de la otra pantalla (grilla de
            # artículos + reversa de Stock) — ver docstring del módulo.
            if not cod.isdigit() or cod == "1":
                continue
            self.combo_motivo.addItem(f"{cod} - {fila.DESCRI or ''}", int(cod))

    # ------------------------------------------------------------------
    # Comprobante a Imputar (sólo Nota de Crédito) — panel achicado
    # (feedback del usuario, 2026-08-19: "El panel comprobantes a
    # imputar achicarlo y colocar el cpbte. seleccionado"), la selección
    # en sí ahora vive en la ventana flotante `ComprobanteAplicarDialog`
    # (botón "Aplicar a...", ver `_armar_cabecera`/`_on_aplicar_a`).
    # ------------------------------------------------------------------
    def _armar_comprobante_imputar(self) -> QGroupBox:
        self.grupo_imputar = QGroupBox("Comprobante a Imputar")
        layout = QHBoxLayout(self.grupo_imputar)
        self.lbl_comprobante_imputar = QLabel("(sin comprobante seleccionado — usá \"Aplicar a...\")")
        layout.addWidget(self.lbl_comprobante_imputar, stretch=1)
        return self.grupo_imputar

    def _refrescar_panel_comprobante(self) -> None:
        comprobante = self._comprobante_seleccionado
        if comprobante is None:
            self.lbl_comprobante_imputar.setText('(sin comprobante seleccionado — usá "Aplicar a...")')
            return
        etiqueta_tipo = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
        fecha_txt = comprobante.FECHA.strftime("%d/%m/%Y") if comprobante.FECHA else "—"
        self.lbl_comprobante_imputar.setText(
            f"{etiqueta_tipo} {(comprobante.PREFIJO or 0):04d}-{(comprobante.CPBTE or 0):08d} — {fecha_txt} — "
            f"Importe sin aplicar: $ {format_decimal(comprobante.DEBE or Decimal('0'))}"
        )

    def _on_aplicar_a(self) -> None:
        if self.cliente_actual is None:
            return
        elegido = ComprobanteAplicarDialog.elegir(
            self.repos, self.cuenta_corriente_service, self.cliente_actual.CODIGO, parent=self
        )
        if elegido is None:
            return
        self._comprobante_seleccionado = elegido
        self._refrescar_panel_comprobante()
        self._recalcular_totales()

    def _on_tipo_cambiado(self) -> None:
        es_nc = self.radio_nc.isChecked()
        self.grupo_imputar.setVisible(es_nc)
        self.btn_aplicar_a.setVisible(es_nc)
        self._refrescar_proximo_numero()
        self._recalcular_totales()

    # ------------------------------------------------------------------
    # Conceptos (DetNC.frm — 3 renglones fijos)
    # ------------------------------------------------------------------
    def _armar_conceptos(self) -> QGroupBox:
        self.grupo_conceptos = QGroupBox("Conceptos (hasta 3 renglones, con importe y Con/Sin IVA cada uno)")
        grupo = self.grupo_conceptos
        layout = QVBoxLayout(grupo)
        self.filas_concepto: list[tuple[UpperCaseLineEdit, MontoLineEdit, QCheckBox]] = []
        for indice in range(3):
            fila = QHBoxLayout()
            txt_desc = UpperCaseLineEdit()
            txt_desc.setPlaceholderText(f"Concepto {indice + 1}")
            fila.addWidget(txt_desc, stretch=1)
            fila.addWidget(QLabel("Importe :"))
            txt_importe = MontoLineEdit("0")
            txt_importe.setMaximumWidth(120)
            txt_importe.editingFinished.connect(self._recalcular_totales)
            fila.addWidget(txt_importe)
            chk_con_iva = QCheckBox("Con IVA")
            # Default "Sin IVA" — mismo estado inicial que DetNC.frm
            # (Check1.Value arranca en 0/sin marcar).
            chk_con_iva.stateChanged.connect(self._recalcular_totales)
            fila.addWidget(chk_con_iva)
            layout.addLayout(fila)
            self.filas_concepto.append((txt_desc, txt_importe, chk_con_iva))
        layout.addStretch()
        return grupo

    def _conceptos_actuales(self) -> list[ConceptoNotaCredito]:
        resultado = []
        for txt_desc, txt_importe, chk in self.filas_concepto:
            descripcion = txt_desc.text().strip()
            importe = txt_importe.decimal()
            if not descripcion or importe <= 0:
                continue
            resultado.append(ConceptoNotaCredito(descripcion=descripcion, importe=importe, con_iva=chk.isChecked()))
        return resultado

    # ------------------------------------------------------------------
    # Pie
    # ------------------------------------------------------------------
    def _armar_pie(self) -> QGroupBox:
        grupo = QGroupBox("Totales")
        layout = QHBoxLayout(grupo)

        layout.addWidget(QLabel("Base Gravada:"))
        self.lbl_base_gravada = QLabel("$ 0,00")
        layout.addWidget(self.lbl_base_gravada)
        layout.addWidget(QLabel("IVA:"))
        self.lbl_iva = QLabel("$ 0,00")
        layout.addWidget(self.lbl_iva)
        layout.addWidget(QLabel("No Gravado:"))
        self.lbl_no_gravado = QLabel("$ 0,00")
        layout.addWidget(self.lbl_no_gravado)
        layout.addStretch()

        recuadro_total, self.lbl_total = crear_recuadro_destacado("TOTAL:")
        layout.addWidget(recuadro_total)
        layout.addStretch()

        self.btn_emitir = QPushButton("Emitir")
        self.btn_emitir.setStyleSheet("font-weight: bold; padding: 8px;")
        self.btn_emitir.setEnabled(False)
        self.btn_emitir.clicked.connect(self._on_emitir)
        layout.addWidget(self.btn_emitir)

        layout.addSpacing(12)
        btn_salir = QPushButton("Salir")
        btn_salir.clicked.connect(self.close)
        layout.addWidget(btn_salir)

        return grupo

    # ------------------------------------------------------------------
    # Cliente
    # ------------------------------------------------------------------
    def _on_elegir_cliente(self) -> None:
        dialogo = ClienteBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.cliente_elegido is None:
            return
        # Misma sesión propia — mismo motivo/bug ya documentado en
        # FacturadorWindow._on_elegir_cliente.
        self.cliente_actual = self.repos.cliente().by_codigo(dialogo.cliente_elegido.CODIGO)
        self._refrescar_cliente()

    def _refrescar_cliente(self) -> None:
        cliente = self.cliente_actual
        # Cambiar de Cliente invalida cualquier comprobante ya elegido
        # en "Aplicar a..." (era de OTRO Cliente).
        self._comprobante_seleccionado = None
        self._refrescar_panel_comprobante()
        if cliente is None:
            self.lbl_cliente.setText("(sin cliente elegido)")
            self.lbl_cliente_contacto.setText("—")
            self.btn_nota_cliente.setEnabled(False)
            self.btn_aplicar_a.setEnabled(False)
            self.tiene_nota_cliente = False
            self.btn_nota_cliente.setStyleSheet("")
            self.btn_elegir_cliente.setText(self.TEXTO_BTN_ELEGIR_CLIENTE)
            self._refrescar_proximo_numero()
            self._recalcular_totales()
            return

        self.lbl_cliente.setText(f"{cliente.CODIGO} — {(cliente.NOMB or '').strip()} — CUIT {cliente.CUIT or 's/d'}")
        self.lbl_cliente_contacto.setText(texto_contacto_cliente(cliente))
        self.btn_nota_cliente.setEnabled(True)
        self.btn_aplicar_a.setEnabled(True)
        self.btn_elegir_cliente.setText(self.TEXTO_BTN_CAMBIAR_CLIENTE)

        self._actualizar_boton_nota_cliente()
        if self.tiene_nota_cliente:
            NotaClienteDialog(self.repos, cliente.CODIGO, parent=self).exec()
            self._actualizar_boton_nota_cliente()

        self._refrescar_proximo_numero()
        self._recalcular_totales()

    def _actualizar_boton_nota_cliente(self) -> None:
        self.tiene_nota_cliente = self.cliente_actual is not None and (
            self.repos.notaclte().by_cliente(self.cliente_actual.CODIGO) is not None
        )
        self.btn_nota_cliente.setStyleSheet(
            "background-color: #ffd54f; font-weight: bold;" if self.tiene_nota_cliente else ""
        )

    def _on_nota_cliente(self) -> None:
        if self.cliente_actual is None:
            return
        NotaClienteDialog(self.repos, self.cliente_actual.CODIGO, parent=self).exec()
        self._actualizar_boton_nota_cliente()

    # ------------------------------------------------------------------
    # Numeración / totales
    # ------------------------------------------------------------------
    def _tipo_actual(self) -> int:
        return TIPO_NC if self.radio_nc.isChecked() else TIPO_ND

    def _letra_actual(self) -> Optional[str]:
        if self.cliente_actual is None:
            return None
        return self.factura_service.letra_comprobante(self.cliente_actual.CIVA or 0, self.cliente_actual.PCIA)

    def _refrescar_proximo_numero(self) -> None:
        letra = self._letra_actual()
        if letra is None:
            self.lbl_proximo_numero.setText("Próx. Nº: —")
            return
        tipo_cbte = codigo_afip(letra, self._tipo_actual())
        punto_venta = punto_venta_por_tipo(self._tipo_actual())
        try:
            ultimo = self.afip.ultimo_comprobante(punto_venta, tipo_cbte)
            self.lbl_proximo_numero.setText(f"Próx. Nº: {punto_venta:04d}-{ultimo + 1:08d} (Letra {letra})")
        except Exception:  # noqa: BLE001 — sólo informativo, no bloquea la carga
            self.lbl_proximo_numero.setText(f"Próx. Nº: (no disponible, Letra {letra})")

    def _detalle_habilitado(self) -> bool:
        """Nota de Crédito necesita el comprobante a imputar elegido
        ANTES de poder cargar Conceptos (feedback del usuario,
        2026-08-19: "Si no selecciona la factura o ND no habilite
        detalle") — Nota de Débito no imputa nada, no tiene esa traba."""
        if not self.radio_nc.isChecked():
            return True
        return self._comprobante_seleccionado is not None

    def _recalcular_totales(self) -> None:
        self.grupo_conceptos.setEnabled(self._detalle_habilitado())

        conceptos = self._conceptos_actuales()
        civa = self.cliente_actual.CIVA if self.cliente_actual else 3
        total = self.nc_service.calcular_total(conceptos, civa_cliente=civa or 3)

        self.lbl_base_gravada.setText(f"$ {format_decimal(total.base_gravada)}")
        self.lbl_iva.setText(f"$ {format_decimal(total.iva)}")
        self.lbl_no_gravado.setText(f"$ {format_decimal(total.no_gravado)}")
        self.lbl_total.setText(f"$ {format_decimal(total.total)}")

        habilitado = self.cliente_actual is not None and total.total > 0 and self._detalle_habilitado()
        if self.radio_nc.isChecked():
            habilitado = habilitado and self._comprobante_seleccionado is not None
        self.btn_emitir.setEnabled(habilitado)

    # ------------------------------------------------------------------
    # Nueva
    # ------------------------------------------------------------------
    def _nueva(self) -> None:
        self.cliente_actual = None
        self._comprobante_seleccionado = None
        for txt_desc, txt_importe, chk in getattr(self, "filas_concepto", []):
            txt_desc.clear()
            txt_importe.set_decimal(Decimal("0"))
            chk.setChecked(False)
        self._refrescar_cliente()
        QTimer.singleShot(0, self._on_elegir_cliente)

    # ------------------------------------------------------------------
    # Emitir
    # ------------------------------------------------------------------
    def _on_emitir(self) -> None:
        if self.cliente_actual is None:
            QMessageBox.warning(self, "Nota de Crédito/Débito", "Elegí un Cliente.")
            return
        if not (self.cliente_actual.CUIT or "").strip():
            QMessageBox.warning(self, "Nota de Crédito/Débito", "El cliente no tiene CUIT cargado — no se puede solicitar CAE.")
            return

        conceptos = self._conceptos_actuales()
        if not conceptos:
            QMessageBox.warning(self, "Nota de Crédito/Débito", "Cargá al menos un concepto con importe.")
            return

        tipo = self._tipo_actual()
        comprobante_a_imputar = self._comprobante_seleccionado if tipo == TIPO_NC else None
        if tipo == TIPO_NC and comprobante_a_imputar is None:
            QMessageBox.warning(self, "Nota de Crédito", "Elegí el comprobante contra el cual se imputa esta Nota de Crédito.")
            return

        letra = self._letra_actual()
        tipo_cbte = codigo_afip(letra, tipo)
        punto_venta = punto_venta_por_tipo(tipo)
        total = self.nc_service.calcular_total(conceptos, civa_cliente=self.cliente_actual.CIVA or 3)
        motivo = self.combo_motivo.currentData() or 0
        fecha = date.today()

        try:
            numero = self.afip.ultimo_comprobante(punto_venta, tipo_cbte) + 1
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Nota de Crédito/Débito", f"Error al conectar con AFIP:\n{exc}")
            return

        etiqueta_tipo = "Nota de Crédito" if tipo == TIPO_NC else "Nota de Débito"
        respuesta = QMessageBox.question(
            self,
            "Nota de Crédito/Débito",
            f"¿Desea grabar la {etiqueta_tipo} {letra} por $ {format_decimal(total.total)}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        try:
            resultado_cae: ResultadoCAE = self.afip.solicitar_cae(
                tipo_cbte=tipo_cbte,
                punto_venta=punto_venta,
                cbte_nro=numero,
                cuit_receptor=self.cliente_actual.CUIT,
                importe_neto=total.base_gravada,
                importe_iva=total.iva,
                importe_iibb=Decimal("0"),
                importe_total=total.total,
                fecha_cbte=fecha,
                importe_no_gravado=total.no_gravado,
                condicion_iva_receptor_id=condicion_iva_receptor(self.cliente_actual.CIVA),
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Nota de Crédito/Débito", f"Error al conectar con AFIP:\n{exc}")
            return

        if not resultado_cae.aprobado:
            QMessageBox.critical(self, "Nota de Crédito/Débito", f"AFIP no aprobó el comprobante.\nMotivo: {resultado_cae.motivo}")
            return

        try:
            resultado = self.emision_service.emitir_concepto(
                cliente=self.cliente_actual,
                tipo=tipo,
                letra=letra,
                punto_venta=punto_venta,
                numero_comprobante=numero,
                conceptos=conceptos,
                total=total,
                motivo=motivo,
                usuario=self._usuario,
                comprobante_a_imputar=comprobante_a_imputar,
                fecha=fecha,
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self,
                "Nota de Crédito/Débito",
                f"AFIP aprobó el CAE {resultado_cae.cae} pero el comprobante NO se pudo grabar en el sistema:\n{exc}\n"
                "Anotar el CAE para grabarlo manualmente.",
            )
            return

        QMessageBox.information(
            self,
            "Nota de Crédito/Débito",
            f"{etiqueta_tipo} {letra} {punto_venta:04d}-{numero:08d} emitida.\n"
            f"CAE: {resultado_cae.cae}\nVencimiento: {resultado_cae.vencimiento}\nTotal: $ {format_decimal(total.total)}",
        )
        self._nueva()

    # ------------------------------------------------------------------
    def _hay_algo_cargado(self) -> bool:
        return bool(self._conceptos_actuales())

    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        if self._hay_algo_cargado():
            respuesta = QMessageBox.question(
                self,
                "Nota de Crédito/Débito",
                "Hay una Nota cargada — se va a perder el ingreso.\n¿Desea salir igual?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if respuesta != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        self.db.close()
        super().closeEvent(event)
