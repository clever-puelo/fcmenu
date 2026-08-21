"""Diálogo de Detalle de Cliente — panel de Alta/Cambio/Baja de `Abmclte.frm`.

Convención de navegación confirmada con el usuario (2026-08-15), aplicada
a todo el sistema salvo aclaración específica: cada ABM tiene una
**ventana de Búsqueda** (lista de resultados + botón Alta) y una
**ventana/diálogo de Detalle** separada que sólo se abre al elegir un
resultado de la lista o al presionar Alta. Este módulo es el Detalle de
Clientes — la Búsqueda vive en `cliente_busqueda_window.py`.

Alcance de este panel (confirmado con el usuario el 2026-08-14, ver
memoria de sesión `pyqt6_ui_progress.md`):

- Se excluyen los "controles fantasma" del legacy que nunca se guardaban
  (Remito/Flete/Lista, Contacto de Venta, Transporte).
- "Email" mapea a `Cliente.EMAIL` (el legacy lo etiquetaba "Contacto de
  Cobranza" pero grababa ahí — bug de etiqueta, se corrigió el label).
- La Baja bloquea si el cliente tiene saldo≠0 o movimientos en Ctasctes.
- Al Grabar o dar de Baja con éxito, el diálogo se cierra (igual que el
  legacy: `Abmclte.frm` hace `Unload Me` después de `Grabacion()` tanto
  en Alta/Cambio como en Baja) — la ventana de Búsqueda que lo abrió
  refresca su lista al recuperar el control.
- El botón "Notas" (`NotaClte.frm`, migrado 2026-08-15) se habilita sólo
  con un cliente ya cargado (no durante Alta, igual que el legacy). Si
  el cliente tiene notas guardadas, el diálogo de notas se abre solo al
  cargarlo (réplica de `Abmclte.frm Form_Activate`, que muestra
  NOTACLTE.Show automáticamente cuando hay una nota — es una alerta real
  del negocio, ej. "cliente moroso", no un adorno cosmético).
"""

from __future__ import annotations

from typing import Callable, Optional
from urllib.parse import quote

from PyQt6.QtCore import QUrl
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from migration.models import Cliente
from migration.repository import RepositoryFactory
from migration.services import ClienteService

from .decimals import parse_decimal
from .dtos_cliente_dialog import DtosClienteDialog
from .iconos import icono
from .nota_cliente_dialog import NotaClienteDialog
from .validators import cuit_valido, email_valido
from .widgets import EnterAsTabFilter, LowerCaseLineEdit, UpperCaseLineEdit

CIVA_OPCIONES = [
    (1, "Inscripto"),
    (2, "Resp.No Insc. (desactualizado, AFIP lo eliminó en 2003)"),
    (3, "Cons.Final"),
    (4, "Exento"),
    (5, "Monotributo"),
]


class ClienteDetalleDialog(QDialog):
    """Alta / Cambio / Baja de UN cliente. Se abre desde
    `ClienteBusquedaWindow` (uso standalone, `main.py`), o directamente
    embebida en el MDI de `MainMenuWindow` (2026-08-16: la ficha aparece
    primero, vacía, y la búsqueda se abre encima — ver
    `main_menu_window._abrir_clientes`); en ese caso `on_buscar` conecta
    el botón "Buscar..." para volver a invocar la búsqueda sin cerrar la
    ficha, y `cargar_cliente()` es el punto de entrada público para que
    quien la abrió le entregue el cliente elegido después."""

    def __init__(
        self,
        repos: RepositoryFactory,
        cliente_service: ClienteService,
        cliente: Optional[Cliente] = None,
        parent: QWidget | None = None,
        *,
        on_buscar: Optional[Callable[[], None]] = None,
    ):
        super().__init__(parent)
        self.repos = repos
        self.cliente_service = cliente_service
        self.cliente_actual = cliente
        self.guardado_o_eliminado = False  # la ventana que abre el diálogo refresca si esto queda True
        self._on_buscar = on_buscar

        self.setWindowTitle("Detalle de Cliente")
        # Más chica que la primera pasada de compactación (feedback del
        # usuario, 2026-08-17, segunda ronda: "sigue siendo larga, sale
        # de la ventana... como para achicar la ventana") — Crédito+
        # Comisión y Vendedor+Zona ahora comparten línea, así que entra
        # todo en menos alto.
        self.resize(760, 540)

        self._construir_ui()
        self._cargar_combos()

        if cliente is not None:
            self._cargar_cliente(cliente)
        else:
            self._blanquear()
            self._set_modo("alta")
            self.txt_nomb.setFocus()

        self._enter_as_tab = EnterAsTabFilter(self)

    # ------------------------------------------------------------------
    # Construcción de la UI
    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        # Líneas más juntas (feedback del usuario, 2026-08-17: "sale del
        # cuadro y no se puede ver ni acceder a la parte de abajo") —
        # antes usaba el spacing default de Qt entre secciones.
        layout.setSpacing(6)

        self.frame_nominales = self._armar_frame_nominales()
        layout.addWidget(self.frame_nominales)

        fila_media = QHBoxLayout()
        self.frame_cobranzas = self._armar_frame_cobranzas()
        fila_media.addWidget(self.frame_cobranzas, stretch=3)
        self.frame_varios = self._armar_frame_varios()
        fila_media.addWidget(self.frame_varios, stretch=1)
        layout.addLayout(fila_media)

        layout.addStretch()
        layout.addLayout(self._armar_barra_botones())

    def _armar_frame_nominales(self) -> QGroupBox:
        box = QGroupBox("Datos Nominales")
        form = QFormLayout(box)
        # Celdas más cortas (feedback del usuario, 2026-08-17: "En
        # Clientes, comprimir las líneas... Sale del cuadro").
        form.setVerticalSpacing(4)

        self.txt_codigo = QSpinBox()
        self.txt_codigo.setRange(1, 999999)
        # Ancho para 5-6 dígitos, no todo el ancho del form — "no es
        # necesario y queda feo tan ancho" (feedback del usuario,
        # 2026-08-17).
        self.txt_codigo.setMaximumWidth(70)
        form.addRow("Código :", self.txt_codigo)

        self.txt_nomb = UpperCaseLineEdit()
        self.txt_nomb.setMaxLength(50)
        form.addRow("Razón Social :", self.txt_nomb)

        self.txt_dir = UpperCaseLineEdit()
        self.txt_dir.setMaxLength(50)
        # Achicado + botón "ver en el mapa" a la derecha (pedido del
        # usuario, 2026-08-21) — abre Google Maps en el navegador con
        # la Dirección + Localidad + Provincia ya cargadas, para no
        # tener que buscarla a mano.
        self.txt_dir.setMaximumWidth(260)
        fila_dir = QHBoxLayout()
        fila_dir.addWidget(self.txt_dir)
        self.btn_mapa = QPushButton()
        self.btn_mapa.setIcon(icono("mapa", 22))
        self.btn_mapa.setToolTip("Ver la Dirección en Google Maps")
        self.btn_mapa.setMaximumWidth(36)
        self.btn_mapa.clicked.connect(self._abrir_mapa)
        fila_dir.addWidget(self.btn_mapa)
        fila_dir.addStretch()
        form.addRow("Dirección :", fila_dir)

        self.txt_loc = UpperCaseLineEdit()
        self.txt_loc.setMaxLength(50)
        form.addRow("Localidad :", self.txt_loc)

        fila_cp_pcia = QHBoxLayout()
        self.txt_cp = UpperCaseLineEdit()
        self.txt_cp.setMaxLength(10)
        fila_cp_pcia.addWidget(self.txt_cp)
        # "Pcia." en vez de "Provincia" — quedaba muy apretado al lado
        # del campo (feedback del usuario, 2026-08-18).
        fila_cp_pcia.addWidget(QLabel("Pcia. :"))
        self.txt_pcia = UpperCaseLineEdit()
        self.txt_pcia.setMaxLength(2)
        fila_cp_pcia.addWidget(self.txt_pcia)
        form.addRow("Cod. Postal :", fila_cp_pcia)

        fila_tel = QHBoxLayout()
        self.txt_tel1 = UpperCaseLineEdit()
        self.txt_tel1.setMaxLength(20)
        fila_tel.addWidget(self.txt_tel1)
        self.txt_tel2 = UpperCaseLineEdit()
        self.txt_tel2.setMaxLength(20)
        fila_tel.addWidget(self.txt_tel2)
        self.txt_fax = UpperCaseLineEdit()
        self.txt_fax.setMaxLength(20)
        fila_tel.addWidget(self.txt_fax)
        form.addRow("Teléfonos / Fax :", fila_tel)

        # Condic. I.V.A. y C.U.I.T. en la misma línea (feedback del
        # usuario, 2026-08-17) — mismo patrón ya usado acá arriba para
        # Cod. Postal/Provincia.
        fila_civa_cuit = QHBoxLayout()
        self.combo_civa = QComboBox()
        for cod, descri in CIVA_OPCIONES:
            self.combo_civa.addItem(descri, cod)
        fila_civa_cuit.addWidget(self.combo_civa, stretch=2)
        fila_civa_cuit.addWidget(QLabel("C.U.I.T. :"))
        self.txt_cuit = QLineEdit()
        self.txt_cuit.setMaxLength(11)
        self.txt_cuit.setToolTip("Número CUIT/CUIL/DNI (sin guiones)")
        fila_civa_cuit.addWidget(self.txt_cuit, stretch=1)
        form.addRow("Condic. I.V.A. :", fila_civa_cuit)

        # Minúsculas (no mayúsculas como el resto de los campos de texto
        # libre) — un email en mayúsculas no es la convención real
        # (feedback del usuario, 2026-08-18).
        self.txt_email = LowerCaseLineEdit()
        self.txt_email.setMaxLength(120)
        self.txt_email.setToolTip("Email de contacto")
        form.addRow("Email :", self.txt_email)

        return box

    def _armar_frame_cobranzas(self) -> QGroupBox:
        box = QGroupBox("Datos de Cobranzas")
        form = QFormLayout(box)
        form.setVerticalSpacing(4)

        # Crédito + Comisión Espec. en la misma línea (feedback del
        # usuario, 2026-08-17: "sigue siendo larga, sale de la ventana...
        # como para achicar la ventana").
        fila_credito_comision = QHBoxLayout()
        self.txt_credito = QLineEdit("0")
        fila_credito_comision.addWidget(self.txt_credito, stretch=1)
        fila_credito_comision.addWidget(QLabel("Comisión Espec. :"))
        self.txt_comision = QLineEdit("0")
        self.txt_comision.setToolTip("Comisión Especial")
        fila_credito_comision.addWidget(self.txt_comision, stretch=1)
        form.addRow("Crédito :", fila_credito_comision)

        self.combo_cvta = QComboBox()
        form.addRow("Cond. Venta :", self.combo_cvta)

        # Vendedor + Zona en la misma línea (mismo pedido).
        fila_vend_zona = QHBoxLayout()
        self.combo_vend = QComboBox()
        fila_vend_zona.addWidget(self.combo_vend, stretch=1)
        fila_vend_zona.addWidget(QLabel("Zona :"))
        self.combo_zona = QComboBox()
        fila_vend_zona.addWidget(self.combo_zona, stretch=1)
        form.addRow("Vendedor :", fila_vend_zona)

        return box

    def _armar_frame_varios(self) -> QGroupBox:
        box = QGroupBox("Varios")
        layout = QVBoxLayout(box)
        self.chk_hab_fact = QCheckBox("Hab. Facturación")
        self.chk_hab_fact.setToolTip("Cliente habilitado para facturar (CANAL=9)")
        layout.addWidget(self.chk_hab_fact)
        layout.addStretch()
        return box

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()

        self.btn_editar = QPushButton("Cambio")
        self.btn_editar.clicked.connect(self._on_toggle_editar)
        fila.addWidget(self.btn_editar)

        self.btn_baja = QPushButton("Baja")
        self.btn_baja.clicked.connect(self._on_baja)
        fila.addWidget(self.btn_baja)

        if self._on_buscar is not None:
            self.btn_buscar = QPushButton("Buscar...")
            self.btn_buscar.setToolTip("Buscar otro cliente sin cerrar esta ficha")
            self.btn_buscar.clicked.connect(self._on_buscar)
            fila.addWidget(self.btn_buscar)

        fila.addStretch()

        # "Cargar Descuentos..." — bajado del GroupBox "Datos de Ventas"
        # (que ocupaba una fila entera arriba sólo para este botón) a la
        # barra de abajo, y ahora oculto (no sólo deshabilitado) hasta
        # que el cliente ya está dado de alta (feedback del usuario,
        # 2026-08-17). Se sacó del menú de A-B-M's — este es el único
        # camino ahora, `_set_modo` controla la visibilidad.
        self.btn_descuentos = QPushButton("Descuentos...")
        self.btn_descuentos.setVisible(False)
        self.btn_descuentos.clicked.connect(self._on_descuentos)
        fila.addWidget(self.btn_descuentos)

        self.btn_notas = QPushButton("Notas")
        self.btn_notas.setEnabled(False)
        self.btn_notas.clicked.connect(self._on_notas)
        fila.addWidget(self.btn_notas)

        self.btn_guardar = QPushButton("Grabar")
        self.btn_guardar.clicked.connect(self._on_guardar)
        fila.addWidget(self.btn_guardar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.close)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    # Carga de combos (Fctabla1: CV / VD / ZN — igual que Form_Load)
    # ------------------------------------------------------------------
    def _cargar_combos(self) -> None:
        fctablas = self.repos.fctablas()
        self._poblar_combo_tabla(self.combo_cvta, fctablas.by_ctab("CV"))
        self._poblar_combo_tabla(self.combo_vend, fctablas.by_ctab("VD"))
        self._poblar_combo_tabla(self.combo_zona, fctablas.by_ctab("ZN"))

    @staticmethod
    def _poblar_combo_tabla(combo: QComboBox, filas) -> None:
        combo.clear()
        for fila in filas:
            cod = (fila.COD or "").strip()
            combo.addItem(f"{cod}-{fila.DESCRI or ''}", cod)

    @staticmethod
    def _seleccionar_por_data(combo: QComboBox, valor) -> None:
        valor_str = str(valor).strip() if valor is not None else ""
        for i in range(combo.count()):
            if str(combo.itemData(i)).strip() == valor_str:
                combo.setCurrentIndex(i)
                return
        if combo.count():
            combo.setCurrentIndex(0)

    # ------------------------------------------------------------------
    # Modos: alta / ver / editar (Command4/Command5 del legacy)
    # ------------------------------------------------------------------
    def _set_modo(self, modo: str) -> None:
        self.modo = modo
        editable = modo in ("alta", "editar")

        for frame in (self.frame_nominales, self.frame_cobranzas, self.frame_varios):
            frame.setEnabled(editable)

        self.txt_codigo.setEnabled(modo == "alta")
        self.btn_guardar.setEnabled(editable)
        self.btn_editar.setText("Sólo Ver" if modo == "editar" else "Cambio")
        self.btn_baja.setEnabled(modo == "ver")
        self.btn_notas.setEnabled(self.cliente_actual is not None)
        # Oculto (no sólo deshabilitado) durante el Alta — "aparezca
        # cuando el cliente ya esta dado de alta" (feedback del usuario,
        # 2026-08-17).
        self.btn_descuentos.setVisible(self.cliente_actual is not None)

    def _blanquear(self) -> None:
        """Réplica de Sub Blanquea() (Abmclte.frm:1839-1891)."""
        self.txt_codigo.setValue(self.repos.cliente().proximo_codigo())
        self.txt_nomb.clear()
        self.txt_dir.clear()
        self.txt_loc.clear()
        self.txt_cp.clear()
        self.txt_pcia.clear()
        self.txt_tel1.clear()
        self.txt_tel2.clear()
        self.txt_fax.clear()
        self.txt_cuit.clear()
        self.txt_email.clear()
        self.combo_civa.setCurrentIndex(0)
        self.txt_credito.setText("0")
        self.txt_comision.setText("0")
        self.chk_hab_fact.setChecked(False)
        if self.combo_cvta.count():
            self.combo_cvta.setCurrentIndex(0)
        if self.combo_vend.count():
            self.combo_vend.setCurrentIndex(0)
        if self.combo_zona.count():
            self.combo_zona.setCurrentIndex(0)

    # ------------------------------------------------------------------
    def cargar_cliente(self, cliente: Cliente) -> None:
        """Punto de entrada público: carga `cliente` en una ficha ya
        visible (usado por `MainMenuWindow` después de que el operador
        elige uno en la búsqueda que se abre encima)."""
        self._cargar_cliente(cliente)

    def _cargar_cliente(self, cliente: Cliente) -> None:
        self.cliente_actual = cliente
        self.txt_codigo.setValue(cliente.CODIGO)
        self.txt_nomb.setText(cliente.NOMB or "")
        self.txt_dir.setText(cliente.DIR or "")
        self.txt_loc.setText(cliente.LOC or "")
        self.txt_cp.setText(cliente.CP or "")
        self.txt_pcia.setText(cliente.PCIA or "")
        self.txt_tel1.setText(cliente.TEL1 or "")
        self.txt_tel2.setText(cliente.TEL2 or "")
        self.txt_fax.setText(cliente.FAX or "")
        self.txt_cuit.setText(cliente.CUIT or "")
        self.txt_email.setText(cliente.EMAIL or "")

        civa = cliente.CIVA or 1
        idx = self.combo_civa.findData(civa)
        self.combo_civa.setCurrentIndex(idx if idx >= 0 else 0)

        self.txt_credito.setText(str(cliente.CREDIT if cliente.CREDIT is not None else 0))
        self.txt_comision.setText(str(cliente.PROIND if cliente.PROIND is not None else 0))
        self.chk_hab_fact.setChecked(cliente.CANAL == 9)

        self._seleccionar_por_data(self.combo_cvta, cliente.CVTA)
        self._seleccionar_por_data(self.combo_vend, cliente.VEND)
        self._seleccionar_por_data(self.combo_zona, cliente.ZONA)

        self._set_modo("ver")

        # Réplica de Abmclte.frm Form_Activate: si el cliente tiene una
        # nota guardada, se muestra sola al cargarlo (alerta real del
        # negocio, ej. "cliente moroso").
        if self.repos.notaclte().by_cliente(cliente.CODIGO) is not None:
            self._on_notas()

    def _abrir_mapa(self) -> None:
        partes = [self.txt_dir.text().strip(), self.txt_loc.text().strip()]
        pcia = self.txt_pcia.text().strip()
        if pcia:
            partes.append(pcia)
        direccion = ", ".join(p for p in partes if p)
        if not direccion:
            QMessageBox.information(self, "Mapa", "Cargá la Dirección primero.")
            return
        url = f"https://www.google.com/maps/search/?api=1&query={quote(direccion)}"
        QDesktopServices.openUrl(QUrl(url))

    def _on_notas(self) -> None:
        if self.cliente_actual is None:
            return
        dialogo = NotaClienteDialog(self.repos, self.cliente_actual.CODIGO, parent=self)
        dialogo.exec()

    def _on_descuentos(self) -> None:
        if self.cliente_actual is None:
            return
        dialogo = DtosClienteDialog(
            self.repos, self.cliente_actual.CODIGO, self.cliente_actual.NOMB or "", parent=self
        )
        dialogo.exec()

    def _on_toggle_editar(self) -> None:
        if self.modo == "editar":
            self._set_modo("ver")
        else:
            self._set_modo("editar")
            self.txt_nomb.setFocus()

    # ------------------------------------------------------------------
    # Grabar (Sub Grabacion() del legacy, sólo camino Alta/Cambio; la
    # Baja física tiene su propio flujo más abajo con validación nueva)
    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        nombre = self.txt_nomb.text().strip()
        if not nombre:
            QMessageBox.warning(self, "ABM de Clientes", "La Razón Social es obligatoria.")
            self.txt_nomb.setFocus()
            return

        cuit = self.txt_cuit.text().strip()
        if cuit and not cuit_valido(cuit):
            QMessageBox.warning(self, "ABM de Clientes", "El CUIT ingresado es inválido.")
            self.txt_cuit.setFocus()
            return

        email = self.txt_email.text().strip()
        if email and not email_valido(email):
            QMessageBox.warning(self, "ABM de Clientes", "El Email ingresado no es válido (falta @).")
            self.txt_email.setFocus()
            return

        codigo = self.txt_codigo.value()
        cliente_repo = self.repos.cliente()

        if self.modo == "alta":
            existente = cliente_repo.by_codigo(codigo)
            if existente is not None:
                QMessageBox.critical(self, "Error de Ingreso", "Código  YA  Existe")
                self.txt_codigo.setFocus()
                return

        respuesta = QMessageBox.question(
            self,
            "ABM de Clientes",
            "Desea Grabar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        datos = dict(
            CODIGO=codigo,
            NOMB=nombre,
            DIR=self.txt_dir.text().strip(),
            LOC=self.txt_loc.text().strip(),
            CP=self.txt_cp.text().strip(),
            PCIA=self.txt_pcia.text().strip(),
            TEL1=self.txt_tel1.text().strip(),
            TEL2=self.txt_tel2.text().strip(),
            FAX=self.txt_fax.text().strip(),
            CUIT=cuit,
            EMAIL=self.txt_email.text().strip(),
            CIVA=self.combo_civa.currentData(),
            CREDIT=parse_decimal(self.txt_credito.text()),
            PROIND=parse_decimal(self.txt_comision.text()),
            CANAL=9 if self.chk_hab_fact.isChecked() else 0,
            CVTA=self.combo_cvta.currentData(),
            VEND=self.combo_vend.currentData(),
            ZONA=self.combo_zona.currentData(),
        )

        if self.modo == "alta":
            nuevo = cliente_repo.create(datos)
            self.cliente_actual = nuevo
        else:
            cliente_repo.update(self.cliente_actual.id, datos)

        QMessageBox.information(self, "ABM de Clientes", "Cliente grabado correctamente.")
        self.guardado_o_eliminado = True
        self.close()

    # ------------------------------------------------------------------
    # Baja — con la validación nueva (confirmada con el usuario) que el
    # legacy no tenía: bloquear si hay saldo o movimientos en Ctasctes.
    # ------------------------------------------------------------------
    def _on_baja(self) -> None:
        if self.cliente_actual is None:
            return

        codigo = self.cliente_actual.CODIGO
        puede, motivo = self.cliente_service.puede_dar_de_baja(codigo)

        if not puede:
            QMessageBox.critical(self, "ABM de Clientes", f"No se puede dar de baja: {motivo}")
            return

        respuesta = QMessageBox.question(
            self,
            "ABM de Clientes",
            "Desea Eliminar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.repos.cliente().delete(self.cliente_actual.id)
        QMessageBox.information(self, "ABM de Clientes", "Cliente eliminado.")
        self.guardado_o_eliminado = True
        self.close()
