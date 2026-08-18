"""`ArregloSubdiarioWindow` — migración de `CargaFC.frm` ("Arreglos >
Subdiario Ventas", menú oculto `Complementos`/`RECICLA1`/`ARRESUBDV1`
de `FCMENU.frm`).

Editor directo de la CABECERA de un comprobante en `FcivaVta` — ver
`ArregloSubdiarioService` para el alcance ("romper vidrio", sin las
validaciones del Facturador) y el bug real corregido (Cod.IVA con el
código real 1-5, no el `ListIndex` 0-4 del legacy).
"""

from __future__ import annotations

from decimal import Decimal

from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.repository import RepositoryFactory
from migration.services import ArregloSubdiarioService

from .cliente_busqueda_window import ClienteBusquedaWindow
from .cliente_detalle_dialog import CIVA_OPCIONES
from .facturador_window import FORMAS_PEDIDO
from .widgets import (
    EnterAsTabFilter,
    EnteroLineEdit,
    MontoLineEdit,
    UpperCaseLineEdit,
    crear_boton_hoy,
    redimensionar_pct_pantalla,
)

# Tipo Cpbte — réplica de CargaFC.frx (Combo1): a diferencia de
# CargaCC.frm no hay "0-Saldo Anterior" (concepto exclusivo de
# CtasCtes) y sí hay "9-Anulada".
TIPOS_CPBTE = [
    (1, "Factura"), (2, "Nota Crédito"), (3, "Nota Débito"),
    (4, "Recibo"), (5, "Pago a Cta."), (6, "Descuento"), (9, "Anulada"),
]
TIPOS_CON_FORMA_PEDIDO = {1, 4}  # Factura/Recibo -> Combo8; el resto -> Motivo (Combo7)

# Letra — réplica de CargaFC.frx (Combo2): a diferencia de CargaCC.frm
# incluye "E" (Exportación).
LETRAS = ["A", "B", "E", "X"]


class ArregloSubdiarioWindow(QMainWindow):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Arreglos — Subdiario Ventas")
        # % de la pantalla real (convención de sistema #11, feedback del
        # usuario, 2026-08-19) — valor sugerido, a ajustar tras probar.
        # El panel "Comprobante" y el de "Cliente" siguen comprimidos a
        # menos filas (ver `_armar_cabecera`/`_armar_cliente`) y los
        # campos angostados acorde al tamaño real de sus datos.
        redimensionar_pct_pantalla(self, 55, 55)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.service = ArregloSubdiarioService(self.db)

        self._construir_ui()
        self._poblar_combos_tablas()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_cabecera())
        layout.addWidget(self._armar_cliente())
        layout.addWidget(self._armar_importes())

        fila_botones = QHBoxLayout()
        self.btn_borrar = QPushButton("Borrar")
        self.btn_borrar.setVisible(False)
        self.btn_borrar.clicked.connect(self._on_borrar)
        fila_botones.addWidget(self.btn_borrar)
        fila_botones.addStretch()
        self.btn_grabar = QPushButton("Grabar")
        self.btn_grabar.setEnabled(False)
        self.btn_grabar.clicked.connect(self._on_grabar)
        fila_botones.addWidget(self.btn_grabar)
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(self.close)
        fila_botones.addWidget(btn_cerrar)
        layout.addLayout(fila_botones)

        self.combo_tipo.currentIndexChanged.connect(self._on_tipo_cambiado)
        self.combo_tipo.currentIndexChanged.connect(self._buscar_fila)
        self.combo_letra.currentIndexChanged.connect(self._buscar_fila)
        self.txt_prefijo.editingFinished.connect(self._buscar_fila)
        self.txt_cpbte.editingFinished.connect(self._buscar_fila)
        self.spin_clte.editingFinished.connect(self._on_cliente_confirmado)

        self._on_tipo_cambiado()
        self._enter_as_tab = EnterAsTabFilter(self, boton_final=self.btn_grabar)

    @staticmethod
    def _form_compacto(panel: QWidget) -> QFormLayout:
        """`QFormLayout` con espaciado/márgenes mínimos — pedido del
        usuario, 2026-08-18, tercera ronda: "panel comprobante achicar,
        comprimir el contenido al mínimo posible" (y lo mismo para
        "Cliente", más abajo) — mismo criterio ya usado en
        `TablasWindow._form_compacto`."""
        form = QFormLayout(panel)
        form.setVerticalSpacing(3)
        form.setContentsMargins(6, 4, 6, 4)
        return form

    def _armar_cabecera(self) -> QGroupBox:
        grupo = QGroupBox("Comprobante")
        form = self._form_compacto(grupo)

        # Fecha y Tipo Cpbte. en la misma línea (feedback del usuario,
        # 2026-08-18, tercera ronda: "comprimir el contenido al mínimo
        # posible") — antes 2 filas propias.
        self.fecha = QDateEdit()
        self.fecha.setCalendarPopup(True)
        self.fecha.setDate(QDate.currentDate())
        self.fecha.setMaximumWidth(110)
        self.combo_tipo = QComboBox()
        for codigo, etiqueta in TIPOS_CPBTE:
            self.combo_tipo.addItem(f"{codigo}-{etiqueta}", codigo)
        self.combo_tipo.setMaximumWidth(160)
        fila_fecha = QHBoxLayout()
        fila_fecha.addWidget(self.fecha)
        fila_fecha.addWidget(crear_boton_hoy(self.fecha))
        fila_fecha.addWidget(QLabel("Tipo Cpbte. :"))
        fila_fecha.addWidget(self.combo_tipo)
        fila_fecha.addStretch()
        form.addRow("Fecha :", fila_fecha)

        # Letra, Prefijo y Nro. Cpbte. en la misma línea (feedback del
        # usuario, 2026-08-18, segunda ronda) — antes 3 filas propias.
        self.combo_letra = QComboBox()
        self.combo_letra.addItems(LETRAS)
        self.combo_letra.setMaximumWidth(70)
        self.txt_prefijo = EnteroLineEdit("0001")
        self.txt_prefijo.setMaximumWidth(70)
        self.txt_cpbte = EnteroLineEdit()
        self.txt_cpbte.setMaximumWidth(90)
        fila_cpbte = QHBoxLayout()
        fila_cpbte.addWidget(self.combo_letra)
        fila_cpbte.addWidget(QLabel("Prefijo (Pto.Vta.) :"))
        fila_cpbte.addWidget(self.txt_prefijo)
        fila_cpbte.addWidget(QLabel("Nro. Cpbte. :"))
        fila_cpbte.addWidget(self.txt_cpbte)
        fila_cpbte.addStretch()
        form.addRow("Letra :", fila_cpbte)

        self.lbl_estado = QLabel()
        self.lbl_estado.setStyleSheet("color: #b00020; font-weight: bold;")
        form.addRow("", self.lbl_estado)

        return grupo

    def _armar_cliente(self) -> QGroupBox:
        grupo = QGroupBox("Cliente")
        form = self._form_compacto(grupo)

        self.spin_clte = QSpinBox()
        self.spin_clte.setRange(0, 999999)
        self.spin_clte.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.spin_clte.setMaximumWidth(90)
        # Botón de búsqueda de Cliente (feedback del usuario, 2026-08-18,
        # segunda ronda) — mismo buscador modal ya usado en Facturador/
        # Recibo/Cta.Cte./Arreglo Cta.Cte.
        self.btn_buscar_cliente = QPushButton("Buscar...")
        self.btn_buscar_cliente.clicked.connect(self._on_buscar_cliente)
        fila_clte = QHBoxLayout()
        fila_clte.addWidget(self.spin_clte)
        fila_clte.addWidget(self.btn_buscar_cliente)
        fila_clte.addStretch()
        form.addRow("Cliente :", fila_clte)

        self.txt_nombre = UpperCaseLineEdit()
        self.txt_nombre.setMaxLength(100)
        self.txt_nombre.setMaximumWidth(320)
        form.addRow("Razón Social :", self.txt_nombre)

        self.txt_pcia = UpperCaseLineEdit()
        self.txt_pcia.setMaxLength(1)
        self.txt_pcia.setMaximumWidth(40)
        form.addRow("Pcia. (letra) :", self.txt_pcia)

        # Cod.IVA y CUIT en la misma línea (feedback del usuario,
        # 2026-08-18, segunda ronda) — antes 2 filas propias.
        self.txt_cuit = EnteroLineEdit()
        self.txt_cuit.setMaximumWidth(110)
        self.combo_civa = QComboBox()
        for codigo, etiqueta in CIVA_OPCIONES:
            self.combo_civa.addItem(etiqueta, codigo)
        self.combo_civa.setMaximumWidth(190)
        fila_civa_cuit = QHBoxLayout()
        fila_civa_cuit.addWidget(self.txt_cuit)
        fila_civa_cuit.addWidget(QLabel("Cod. IVA :"))
        fila_civa_cuit.addWidget(self.combo_civa)
        fila_civa_cuit.addStretch()
        form.addRow("CUIT :", fila_civa_cuit)

        # Vendedor y Zona en la misma línea (feedback del usuario,
        # 2026-08-18, segunda ronda) — antes 2 filas propias.
        self.combo_vend = QComboBox()
        self.combo_vend.setMaximumWidth(170)
        self.combo_zona = QComboBox()
        self.combo_zona.setMaximumWidth(130)
        fila_vend_zona = QHBoxLayout()
        fila_vend_zona.addWidget(self.combo_vend)
        fila_vend_zona.addWidget(QLabel("Zona :"))
        fila_vend_zona.addWidget(self.combo_zona)
        fila_vend_zona.addStretch()
        form.addRow("Vendedor :", fila_vend_zona)

        # Cond.Vta. y Forma Ped./Motivo en la misma línea (feedback del
        # usuario, 2026-08-18, tercera ronda: "achicar ese panel al
        # mínimo también") — antes Cond.Vta. tenía su propia fila y
        # Forma Ped./Motivo otras 2 (una por combo, alternados con
        # `setVisible`, con una fila fantasma de etiqueta vacía cuando
        # el otro combo estaba oculto). Ahora los 3 comparten una sola
        # fila; `self.lbl_motivo` sigue cambiando de texto en
        # `_on_tipo_cambiado` ("Forma Ped. :"/"Motivo :").
        self.combo_cvta = QComboBox()
        self.combo_cvta.setMaximumWidth(170)
        self.lbl_motivo = QLabel("Forma Ped. :")
        self.combo_forma_ped = QComboBox()
        self.combo_forma_ped.addItems(FORMAS_PEDIDO)
        self.combo_forma_ped.setMaximumWidth(150)
        self.combo_motivo = QComboBox()
        self.combo_motivo.setMaximumWidth(190)
        fila_cvta_motivo = QHBoxLayout()
        fila_cvta_motivo.addWidget(self.combo_cvta)
        fila_cvta_motivo.addWidget(self.lbl_motivo)
        fila_cvta_motivo.addWidget(self.combo_forma_ped)
        fila_cvta_motivo.addWidget(self.combo_motivo)
        fila_cvta_motivo.addStretch()
        form.addRow("Cond. Vta. :", fila_cvta_motivo)

        return grupo

    def _armar_importes(self) -> QGroupBox:
        grupo = QGroupBox("Importes")
        form = self._form_compacto(grupo)

        # Campos acordes al tamaño de los datos (feedback del usuario,
        # 2026-08-18, tercera ronda) — importes topeados como en
        # `ArregloCtaCteWindow`, enteros chicos (Items/Tot.Unid.) más
        # angostos todavía.
        self.txt_grins = MontoLineEdit()
        self.txt_grins.setMaximumWidth(120)
        form.addRow("Gravado :", self.txt_grins)
        self.txt_ivains = MontoLineEdit()
        self.txt_ivains.setMaximumWidth(120)
        form.addRow("IVA Inscr. :", self.txt_ivains)
        self.txt_ivanoins = MontoLineEdit()
        self.txt_ivanoins.setMaximumWidth(120)
        form.addRow("IVA No Insc. :", self.txt_ivanoins)
        self.txt_exento = MontoLineEdit()
        self.txt_exento.setMaximumWidth(120)
        form.addRow("Exento :", self.txt_exento)
        self.txt_bon = MontoLineEdit()
        self.txt_bon.setMaximumWidth(120)
        form.addRow("Descuentos :", self.txt_bon)
        self.txt_porcib = MontoLineEdit()
        self.txt_porcib.setMaximumWidth(80)
        form.addRow("% I.B. :", self.txt_porcib)
        self.txt_totib = MontoLineEdit()
        self.txt_totib.setMaximumWidth(120)
        form.addRow("Impte. I.B. :", self.txt_totib)
        self.txt_items = EnteroLineEdit("0")
        self.txt_items.setMaximumWidth(70)
        form.addRow("Items :", self.txt_items)
        self.txt_totcan = EnteroLineEdit("0")
        self.txt_totcan.setMaximumWidth(70)
        form.addRow("Tot.Unid. :", self.txt_totcan)

        return grupo

    # ------------------------------------------------------------------
    def _poblar_combos_tablas(self) -> None:
        for cod, descri in self._opciones_tabla("VD"):
            self.combo_vend.addItem(f"{cod}-{descri}", cod)
        for cod, descri in self._opciones_tabla("ZN"):
            self.combo_zona.addItem(f"{cod}-{descri}", cod)
        for cod, descri in self._opciones_tabla("CV"):
            self.combo_cvta.addItem(f"{cod}-{descri}", cod)
        for cod, descri in self._opciones_tabla("MT"):
            self.combo_motivo.addItem(f"{cod}-{descri}", cod)

    def _opciones_tabla(self, ctab: str) -> list[tuple[int, str]]:
        filas = self.repos.fctablas().by_ctab(ctab)
        opciones = []
        for f in filas:
            cod_texto = (f.COD or "").strip()
            if cod_texto.isdigit():
                opciones.append((int(cod_texto), (f.DESCRI or "").strip()))
        return opciones

    def _on_tipo_cambiado(self) -> None:
        tipo = self.combo_tipo.currentData()
        es_forma_pedido = tipo in TIPOS_CON_FORMA_PEDIDO
        self.lbl_motivo.setText("Forma Ped. :" if es_forma_pedido else "Motivo :")
        self.combo_forma_ped.setVisible(es_forma_pedido)
        self.combo_motivo.setVisible(not es_forma_pedido)

    def _on_cliente_confirmado(self) -> None:
        clte = self.spin_clte.value()
        cliente = self.repos.cliente().by_codigo(clte) if clte else None
        if cliente is None:
            self.txt_nombre.setText("*** Cliente NO Existe ***")
            return
        self.txt_nombre.setText((cliente.NOMB or "").strip())
        self.txt_pcia.setText((cliente.PCIA or "").strip()[:1])
        self.txt_cuit.setText((cliente.CUIT or "").strip())
        idx_civa = self.combo_civa.findData(cliente.CIVA or 1)
        if idx_civa >= 0:
            self.combo_civa.setCurrentIndex(idx_civa)
        self._seleccionar_por_data(self.combo_vend, cliente.VEND)
        self._seleccionar_por_data(self.combo_zona, cliente.ZONA)
        self._seleccionar_por_data(self.combo_cvta, cliente.CVTA)

    def _on_buscar_cliente(self) -> None:
        dialogo = ClienteBusquedaWindow(parent=self, modo_seleccion=True)
        dialogo.exec()
        if dialogo.cliente_elegido is None:
            return
        self.spin_clte.setValue(dialogo.cliente_elegido.CODIGO)
        self._on_cliente_confirmado()

    @staticmethod
    def _seleccionar_por_data(combo: QComboBox, valor: int | None) -> None:
        if valor is None:
            return
        idx = combo.findData(valor)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    # ------------------------------------------------------------------
    def _buscar_fila(self) -> None:
        letra = self.combo_letra.currentText()
        tipo = self.combo_tipo.currentData()
        prefijo = self.txt_prefijo.entero()
        cpbte = self.txt_cpbte.entero()
        if prefijo is None or cpbte is None:
            return

        fila = self.service.buscar(letra, tipo, prefijo, cpbte)

        if fila is not None:
            self.lbl_estado.setText("")
            self.spin_clte.setValue(fila.CLTE or 0)
            self.txt_nombre.setText((fila.NOMB or "").strip())
            self.txt_pcia.setText((fila.PCIA or "").strip()[:1])
            self.txt_cuit.setText((fila.CUIT or "").strip())
            idx_civa = self.combo_civa.findData(int(fila.CIVA) if (fila.CIVA or "").strip().isdigit() else 1)
            self.combo_civa.setCurrentIndex(idx_civa if idx_civa >= 0 else 0)
            self._seleccionar_por_data(self.combo_vend, fila.VEND)
            self._seleccionar_por_data(self.combo_zona, fila.ZONA)
            self._seleccionar_por_data(self.combo_cvta, fila.CVTA)
            moti_texto = (fila.MOTI or "").strip()
            if moti_texto.isdigit():
                if tipo in TIPOS_CON_FORMA_PEDIDO:
                    idx = int(moti_texto) - 1
                    if 0 <= idx < self.combo_forma_ped.count():
                        self.combo_forma_ped.setCurrentIndex(idx)
                else:
                    self._seleccionar_por_data(self.combo_motivo, int(moti_texto))

            self.txt_grins.set_decimal(fila.GRINS or Decimal("0"))
            self.txt_ivains.set_decimal(fila.IVAINS or Decimal("0"))
            self.txt_ivanoins.set_decimal(fila.IVANOINS or Decimal("0"))
            self.txt_exento.set_decimal(fila.EXENTO or Decimal("0"))
            self.txt_bon.set_decimal(fila.BON or Decimal("0"))
            self.txt_porcib.set_decimal(fila.PORCIB or Decimal("0"))
            self.txt_totib.set_decimal(fila.TOTIB or Decimal("0"))
            self.txt_items.setText(str(fila.ITEMS or 0))
            self.txt_totcan.setText(str(fila.TOTCAN or 0))
            self.btn_borrar.setVisible(True)
        else:
            self.lbl_estado.setText("N U E V O")
            self.spin_clte.setValue(0)
            self.txt_nombre.setText("")
            self.txt_pcia.setText("")
            self.txt_cuit.setText("")
            for campo in (
                self.txt_grins, self.txt_ivains, self.txt_ivanoins, self.txt_exento,
                self.txt_bon, self.txt_porcib, self.txt_totib,
            ):
                campo.set_decimal(Decimal("0"))
            self.txt_items.setText("0")
            self.txt_totcan.setText("0")
            self.btn_borrar.setVisible(False)

        self.btn_grabar.setEnabled(True)

    # ------------------------------------------------------------------
    def _moti_a_grabar(self) -> str:
        tipo = self.combo_tipo.currentData()
        if tipo in TIPOS_CON_FORMA_PEDIDO:
            return str(self.combo_forma_ped.currentIndex() + 1)
        return str(self.combo_motivo.currentData() or 0)

    def _on_grabar(self) -> None:
        respuesta = QMessageBox.question(
            self, "Carga de Datos a la Tabla", "¿Desea continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.service.grabar(
            tipo=self.combo_tipo.currentData(),
            letra=self.combo_letra.currentText(),
            ptovta=self.txt_prefijo.entero() or 0,
            cpbte=self.txt_cpbte.entero() or 0,
            fecha=self.fecha.date().toPyDate(),
            clte=self.spin_clte.value(),
            nomb=self.txt_nombre.text().strip(),
            pcia=self.txt_pcia.text().strip(),
            cuit=self.txt_cuit.text().strip(),
            grins=self.txt_grins.decimal(),
            ivains=self.txt_ivains.decimal(),
            ivanoins=self.txt_ivanoins.decimal(),
            exento=self.txt_exento.decimal(),
            bon=self.txt_bon.decimal(),
            porcib=self.txt_porcib.decimal(),
            totib=self.txt_totib.decimal(),
            items=self.txt_items.entero() or 0,
            totcan=self.txt_totcan.entero() or 0,
            civa=self.combo_civa.currentData(),
            vend=self.combo_vend.currentData(),
            zona=self.combo_zona.currentData(),
            cvta=self.combo_cvta.currentData(),
            moti=self._moti_a_grabar(),
        )
        QMessageBox.information(self, "Arreglos", "Registro grabado.")
        self._buscar_fila()

    def _on_borrar(self) -> None:
        respuesta = QMessageBox.question(
            self, "Eliminar Datos de la Tabla", "¿Desea continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.service.borrar(
            self.combo_letra.currentText(), self.combo_tipo.currentData(),
            self.txt_prefijo.entero() or 0, self.txt_cpbte.entero() or 0,
        )
        QMessageBox.information(self, "Arreglos", "Registro eliminado.")
        self._buscar_fila()

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
