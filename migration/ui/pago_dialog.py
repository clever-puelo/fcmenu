"""Diálogo de alta/edición de una forma de pago (Cheque o Retención/
Anticipo) para el Recibo — reemplaza el panel emergente "Detalle del
Cheque"/"Anticipos y Retenciones" de `DetPago.frm` (`Frame3`, que
alternaba visibilidad de controles según la "Operación" elegida en el
mismo formulario) por un único diálogo modal, mismo criterio ya usado en
el resto de la migración (un diálogo por acción, en vez de controles que
aparecen/desaparecen en la misma pantalla).

**Modo edición** (feedback del usuario, 2026-08-15): pasando
`pago_existente` se precargan los campos, el título cambia a "Editar...",
el botón "Agregar" pasa a "Guardar" y aparece "Borrar" (antes era un
botón "Quitar" aparte en `ReciboWindow`, ahora vive acá — "si el renglón
existe y se quiere modificar"). `ReciboWindow` abre este diálogo con un
clic sobre una fila ya cargada de la grilla de Pagos.

Devuelve en `self.resultado` un `PagoCheque`/`PagoRetencion` (alta o
edición confirmada), o `self.eliminar = True` si se pidió Borrar — y
`self.resultado is None` sin `eliminar` si se canceló sin cambios.
"""

from __future__ import annotations

from PyQt6.QtCore import QDate, QTimer
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.repository import RepositoryFactory
from migration.services import PagoCheque, PagoRetencion

from .banco_busqueda_dialog import BancoBusquedaDialog
from .widgets import EnteroLineEdit, EnterAsTabFilter, MontoLineEdit, UpperCaseLineEdit, crear_boton_hoy

# Combo "Operación" — mismo orden que DetPago.frm Combo1 (TIPREG = índice+1).
TIPOS_RETENCION = [
    (1, "Retenc.Gan."),
    (2, "Retenc.IB"),
    (3, "Retenc.IVA"),
    (4, "Retenc.SUS"),
    (5, "Dep./Transf."),
    (6, "Tarj.Déb/Cr."),
    (7, "Cheques Vs."),
    (8, "Cheq.Electr."),
    (9, "Baja Incobr."),
]
# TIPREG=6 ("Tarj.Déb/Cr.") es el único caso, además de Cheque, donde
# DetPago.frm también pide "Banco" para una Retención/Anticipo.
TIPREG_TARJETA = 6
TIPREG_CHEQUE = 0  # sentinel local — "Cheque" no es un TIPREG real, va a la tabla Cheques, no MovimVS


class PagoDialog(QDialog):
    def __init__(
        self,
        repos: RepositoryFactory,
        pago_existente: PagoCheque | PagoRetencion | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.repos = repos
        self.pago_existente = pago_existente
        self.resultado: PagoCheque | PagoRetencion | None = None
        self.eliminar = False

        self.setWindowTitle("Editar Forma de Pago" if pago_existente is not None else "Agregar Forma de Pago")
        self.resize(440, 400)

        self._construir_ui()
        if pago_existente is not None:
            self._precargar(pago_existente)
        self._actualizar_visibilidad()
        # Enter avanza el foco como Tab (convención de sistema) — se
        # instala DESPUÉS de armar todos los campos (ver docstring de
        # EnterAsTabFilter, se autoinstala en los descendientes que ya
        # existan en ese momento). `boton_final=self.btn_agregar`:
        # Enter en el último campo (Observ.) lleva el foco al botón
        # "Agregar"/"Guardar" en vez de darle la vuelta a "Tipo"
        # (feedback del usuario, 2026-08-15).
        self._enter_as_tab = EnterAsTabFilter(self, boton_final=self.btn_agregar)
        # Foco en el primer campo REAL de carga al abrir (feedback del
        # usuario, 2026-08-15: "que se posicione en el numero de
        # cheque") — "Tipo" ya arranca en "Cheque" por default, no hace
        # falta pasar por ahí primero. Si el Tipo no es Cheque (modo
        # edición de una Retención), Nº Cheque está oculto y no tiene
        # sentido enfocarlo — cae en "Tipo". Diferido porque el diálogo
        # todavía no está mostrado en este punto del __init__.
        QTimer.singleShot(0, self._foco_inicial)

    def _foco_inicial(self) -> None:
        if self.txt_nro_cheque.isVisible():
            self.txt_nro_cheque.setFocus()
        else:
            self.combo_tipo.setFocus()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.combo_tipo = QComboBox()
        self.combo_tipo.addItem("Cheque", TIPREG_CHEQUE)
        for tipreg, etiqueta in TIPOS_RETENCION:
            self.combo_tipo.addItem(f"{tipreg}-{etiqueta}", tipreg)
        self.combo_tipo.currentIndexChanged.connect(self._actualizar_visibilidad)
        form.addRow("Tipo :", self.combo_tipo)

        self.txt_nro_cheque = EnteroLineEdit()
        self.lbl_nro_cheque = QLabel("Nº Cheque :")
        form.addRow(self.lbl_nro_cheque, self.txt_nro_cheque)

        fila_banco = QHBoxLayout()
        self.txt_banco = EnteroLineEdit()
        self.txt_banco.setMaximumWidth(60)
        self.txt_banco.editingFinished.connect(self._on_banco_confirmado)
        fila_banco.addWidget(self.txt_banco)
        self.lbl_nombre_banco = QLabel("")
        fila_banco.addWidget(self.lbl_nombre_banco, stretch=1)
        self.lbl_banco = QLabel("Banco (código) :")
        form.addRow(self.lbl_banco, fila_banco)

        fila_emision = QHBoxLayout()
        self.fecha_emision = QDateEdit()
        self.fecha_emision.setCalendarPopup(True)
        self.fecha_emision.setDate(QDate.currentDate())
        fila_emision.addWidget(self.fecha_emision)
        fila_emision.addWidget(crear_boton_hoy(self.fecha_emision))
        self.lbl_fecha_emision = QLabel("Fecha Emisión :")
        form.addRow(self.lbl_fecha_emision, fila_emision)

        fila_vencimiento = QHBoxLayout()
        self.fecha_vencimiento = QDateEdit()
        self.fecha_vencimiento.setCalendarPopup(True)
        self.fecha_vencimiento.setDate(QDate.currentDate())
        fila_vencimiento.addWidget(self.fecha_vencimiento)
        fila_vencimiento.addWidget(crear_boton_hoy(self.fecha_vencimiento))
        self.lbl_fecha_vencimiento = QLabel("Fecha Vto. :")
        form.addRow(self.lbl_fecha_vencimiento, fila_vencimiento)

        self.combo_tipo_cheque = QComboBox()
        self.combo_tipo_cheque.addItems(["Común (Papel)", "Electrónico"])
        self.lbl_tipo_cheque = QLabel("Tipo Cheque :")
        form.addRow(self.lbl_tipo_cheque, self.combo_tipo_cheque)

        self.chk_a_la_orden = QCheckBox("A la Orden")
        self.lbl_a_la_orden = QLabel("")
        form.addRow(self.lbl_a_la_orden, self.chk_a_la_orden)

        self.txt_importe = MontoLineEdit("0")
        form.addRow("Importe :", self.txt_importe)

        self.txt_datos_adicionales = UpperCaseLineEdit()
        form.addRow("Datos Adic. :", self.txt_datos_adicionales)

        self.txt_concepto = UpperCaseLineEdit()
        form.addRow("Concepto :", self.txt_concepto)

        self.txt_observaciones = UpperCaseLineEdit()
        form.addRow("Observ. :", self.txt_observaciones)

        layout.addLayout(form)
        layout.addLayout(self._armar_barra_botones())

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        self.btn_borrar = QPushButton("Borrar")
        self.btn_borrar.setVisible(self.pago_existente is not None)
        self.btn_borrar.clicked.connect(self._on_borrar)
        fila.addWidget(self.btn_borrar)

        fila.addStretch()
        self.btn_agregar = QPushButton("Guardar" if self.pago_existente is not None else "Agregar")
        self.btn_agregar.clicked.connect(self._on_agregar)
        fila.addWidget(self.btn_agregar)
        self.btn_cancelar = QPushButton("Cancelar")
        self.btn_cancelar.clicked.connect(self.reject)
        fila.addWidget(self.btn_cancelar)
        return fila

    # ------------------------------------------------------------------
    def _precargar(self, pago: PagoCheque | PagoRetencion) -> None:
        if isinstance(pago, PagoCheque):
            self.combo_tipo.setCurrentIndex(self.combo_tipo.findData(TIPREG_CHEQUE))
            self.txt_nro_cheque.setText(str(pago.nro_cheque))
            self.txt_banco.setText(pago.banco)
            self._resolver_nombre_banco()
            self.fecha_emision.setDate(QDate(pago.fecha_emision.year, pago.fecha_emision.month, pago.fecha_emision.day))
            self.fecha_vencimiento.setDate(
                QDate(pago.fecha_vencimiento.year, pago.fecha_vencimiento.month, pago.fecha_vencimiento.day)
            )
            self.combo_tipo_cheque.setCurrentIndex(pago.tipo_cheque)
            self.chk_a_la_orden.setChecked(pago.a_la_orden)
            self.txt_importe.set_decimal(pago.importe)
            self.txt_datos_adicionales.setText(pago.datos_adicionales)
            self.txt_concepto.setText(pago.concepto)
            self.txt_observaciones.setText(pago.observaciones)
        else:
            self.combo_tipo.setCurrentIndex(self.combo_tipo.findData(pago.tipreg))
            self.txt_banco.setText(pago.banco_o_dato)
            self._resolver_nombre_banco()
            self.txt_importe.set_decimal(pago.importe)
            self.txt_datos_adicionales.setText(pago.datos_adicionales)
            self.txt_concepto.setText(pago.concepto)
            self.txt_observaciones.setText(pago.observaciones)

    # ------------------------------------------------------------------
    def _es_cheque(self) -> bool:
        return self.combo_tipo.currentData() == TIPREG_CHEQUE

    def _actualizar_visibilidad(self) -> None:
        es_cheque = self._es_cheque()
        for widget in (
            self.txt_nro_cheque, self.lbl_nro_cheque,
            self.fecha_emision, self.lbl_fecha_emision,
            self.fecha_vencimiento, self.lbl_fecha_vencimiento,
            self.combo_tipo_cheque, self.lbl_tipo_cheque,
            self.chk_a_la_orden, self.lbl_a_la_orden,
        ):
            widget.setVisible(es_cheque)

        muestra_banco = es_cheque or self.combo_tipo.currentData() == TIPREG_TARJETA
        self.txt_banco.setVisible(muestra_banco)
        self.lbl_banco.setVisible(muestra_banco)
        self.lbl_nombre_banco.setVisible(muestra_banco)

    def _resolver_nombre_banco(self) -> None:
        """Sólo MUESTRA el nombre si el código ya está cargado — nunca
        abre la búsqueda sola (se usa al precargar un pago existente,
        donde un código en blanco es simplemente "sin banco cargado",
        no una acción del operador que haya que resolver)."""
        codigo = self.txt_banco.text().strip()
        if not codigo.isdigit():
            self.lbl_nombre_banco.setText("")
            return
        banco = self.repos.banco().by_cod(int(codigo))
        self.lbl_nombre_banco.setText((banco.NOMBRE or "").strip() if banco else "*** No existe en Tabla ***")

    def _on_banco_confirmado(self) -> None:
        """Wired a `editingFinished` de `txt_banco` — acá sí es una
        acción real del operador saliendo del campo."""
        codigo = self.txt_banco.text().strip()
        if not codigo:
            # Feedback del usuario (2026-08-15): dejar el código en blanco
            # ya no intenta "buscar" nada solo (eso daba un error
            # inconsistente) — abre directo la búsqueda por nombre.
            self._abrir_busqueda_banco()
            return
        self._resolver_nombre_banco()
        self._foco_despues_de_banco()

    def _abrir_busqueda_banco(self) -> None:
        dialogo = BancoBusquedaDialog(self.repos, parent=self)
        dialogo.exec()
        if dialogo.banco_elegido is None:
            self.lbl_nombre_banco.setText("")
            return
        self.txt_banco.setText(str(dialogo.banco_elegido.COD))
        self.lbl_nombre_banco.setText((dialogo.banco_elegido.NOMBRE or "").strip())
        self._foco_despues_de_banco()

    def _foco_despues_de_banco(self) -> None:
        """Feedback del usuario (2026-08-15): "luego de elgir banco
        salte a la fecha" — sólo tiene sentido cuando la Fecha de
        Emisión está visible (Cheque); para Tarjeta (el otro caso que
        pide Banco) no hay fechas, el orden normal de Tab ya alcanza."""
        if self.fecha_emision.isVisible():
            self.fecha_emision.setFocus()

    # ------------------------------------------------------------------
    def _on_borrar(self) -> None:
        respuesta = QMessageBox.question(
            self,
            "Forma de Pago",
            "¿Borrar esta forma de pago?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return
        self.eliminar = True
        self.accept()

    def _on_agregar(self) -> None:
        importe = self.txt_importe.decimal()
        if importe <= 0:
            QMessageBox.warning(self, "Forma de Pago", "El Importe tiene que ser mayor a cero.")
            return

        if self._es_cheque():
            nro_texto = self.txt_nro_cheque.text().strip()
            if not nro_texto.isdigit():
                QMessageBox.warning(self, "Forma de Pago", "El Nº de Cheque es obligatorio.")
                return
            self.resultado = PagoCheque(
                nro_cheque=int(nro_texto),
                banco=self.txt_banco.text().strip(),
                fecha_emision=self.fecha_emision.date().toPyDate(),
                fecha_vencimiento=self.fecha_vencimiento.date().toPyDate(),
                importe=importe,
                a_la_orden=self.chk_a_la_orden.isChecked(),
                datos_adicionales=self.txt_datos_adicionales.text().strip(),
                concepto=self.txt_concepto.text().strip(),
                observaciones=self.txt_observaciones.text().strip(),
                tipo_cheque=self.combo_tipo_cheque.currentIndex(),
            )
        else:
            self.resultado = PagoRetencion(
                tipreg=self.combo_tipo.currentData(),
                importe=importe,
                banco_o_dato=self.txt_banco.text().strip(),
                datos_adicionales=self.txt_datos_adicionales.text().strip(),
                concepto=self.txt_concepto.text().strip(),
                observaciones=self.txt_observaciones.text().strip(),
            )

        self.accept()
