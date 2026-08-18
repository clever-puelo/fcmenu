"""ABM de Tablas — migración de `ABMTablas.frm` + `Acttabla.frm`.

`Tablas.frm` (archivo homónimo parecido) NO está en `FCMENU.VBP` — es un
muerto, no se migra. El form vivo es `ABMTablas.frm` (selector de
categoría) que abre `Acttabla.frm` (editor), invocado desde `FCMENU.frm`.

Las 8 categorías viven todas en `Fctabla1`, distinguidas por `CTAB`: SC
(Secciones), VD (Vendedores), ZN (Zonas), PV (Provincias), UM (Unidades
de Medida), CV (Condición de Venta), MT (Motivos de NC/ND), VS (Datos
Varios). Confirmado con el usuario (2026-08-14) el significado real de
los campos de Sección, que no es evidente leyendo sólo el código:

- **ALF1/ALF2/ALF3** ("Especificación"): arman el código del artículo
  (ver `migration/ui/articulo_codigo.py` — sólo ALF1/ALF2 se usan en la
  práctica, ALF3 está vacío en las 63 Secciones reales).
- **ALF4/ALF5/ALF6** ("Cálculo de Precio"): se multiplican para obtener
  el "Precio Unitario" de esa especificación. No se implementa el
  cálculo acá — eso es de la fase de facturación (DetFact.frm); esta
  pantalla sólo permite configurar QUÉ Unidad de Medida va en cada slot.
- **ALF7** ("Unid. de Facturación", obligatoria): la cantidad (mtrs/kg/
  unid) que multiplicada por el Precio Unitario da el Importe.
- **MCA1**/**MCA2**: "Usar para Facturar" / "Precio en Dólares" (el
  .frm tiene sus dos tooltips pegados por error — se usa el Caption).

Para Unidades de Medida, "Posición" (NUMSD3) no es un número de columna
libre: es una categoría fija de 5 valores, extraída del recurso binario
`Acttabla.frx` (1-Nro/Pulg, 2-Mtr/Kg, 3-MM, 4-Telas, 6-Unidades — nota:
no existe la opción 5).

**Baja**: igual que Cliente/Articulo, el legacy hacía DELETE físico sin
validar (Acttabla.frm Grabacion(), TIPOMov=2). Acá se bloquea según la
categoría vía `TablaService.puede_dar_de_baja()` (Sección con artículos,
Unidad de Medida usada por alguna Sección, Vendedor/Zona/Cond.Venta con
clientes asignados).
"""

from __future__ import annotations

from typing import Optional

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.models import Fctabla1
from migration.repository import RepositoryFactory
from migration.services import TablaService

from .decimals import parse_decimal
from .theme import Verde
from .widgets import EnterAsTabFilter, TablaBusqueda, UpperCaseLineEdit

# Límites en píxeles del panel superior "TABLA:" (25% del alto de la
# ventana) — mismo criterio de piso/techo ya usado en
# `MainMenuWindow._ajustar_proporciones` para que no quede inusable en
# ventanas muy chicas o gigantes.
# Mitad de lo que era antes (feedback del usuario, 2026-08-17, segunda
# ronda: "La ventana sale del panel central... Achicar el panel
# superior a la mitad").
ALTO_MIN_PANEL_TABLA = 55
ALTO_MAX_PANEL_TABLA = 110
PROPORCION_PANEL_TABLA = 0.125  # 25% -> 12.5% (la mitad)

# (CTAB, etiqueta del menú, el Código es texto alfanumérico [True] o
# numérico [False] — ver Acttabla.frm Grabacion(): LaOpcion 1/5 usan el
# código tal cual se tipea, el resto lo pasa por Val()).
CATEGORIAS: list[tuple[str, str, bool]] = [
    ("SC", "Secciones Correa", True),
    ("VD", "Vendedores", False),
    ("ZN", "Zonas", False),
    ("PV", "Provincias", False),
    ("UM", "Unid. Medida Prod.", True),
    ("CV", "Condic. de Venta", False),
    ("MT", "Motivos de N/C y D", False),
    ("VS", "Datos Varios", False),
]

# Posición de Unidad de Medida — enum fijo decodificado de Acttabla.frx.
POSICIONES_UM = [
    ("1", "1-Nro/Pulg"),
    ("2", "2-Mtr/Kg"),
    ("3", "3-MM"),
    ("4", "4-Telas"),
    ("6", "6-Unidades"),
]


class TablasWindow(QMainWindow):
    """Ventana de ABM de Tablas (ABMTablas + Acttabla)."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("ABM de Tablas")
        self.resize(880, 560)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)
        self.tabla_service = TablaService(self.db)

        self.fila_actual: Optional[Fctabla1] = None
        self.modo: Optional[str] = None  # "alta" | "ver" | "editar"

        self._construir_ui()
        self._ajustar_panel_tabla()
        self._on_categoria_cambiada()

        self._enter_as_tab = EnterAsTabFilter(self)

    # ------------------------------------------------------------------
    @property
    def categoria_actual(self) -> tuple[str, str, bool]:
        return CATEGORIAS[self.combo_categoria.currentIndex()]

    def resizeEvent(self, event) -> None:  # noqa: N802 (override de Qt)
        super().resizeEvent(event)
        self._ajustar_panel_tabla()

    def _ajustar_panel_tabla(self) -> None:
        """Panel superior — 12.5% del alto de la ventana (mitad del
        25% original, pedido del usuario, 2026-08-17), recalculado en
        cada resize — mismo criterio que
        `MainMenuWindow._ajustar_proporciones`."""
        self.panel_tabla.setFixedHeight(
            max(ALTO_MIN_PANEL_TABLA, min(ALTO_MAX_PANEL_TABLA, round(self.height() * PROPORCION_PANEL_TABLA)))
        )

    # ------------------------------------------------------------------
    # Construcción de la UI
    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        layout.addWidget(self._armar_panel_tabla())

        cuerpo_widget = QWidget()
        cuerpo = QHBoxLayout(cuerpo_widget)
        cuerpo.setContentsMargins(8, 8, 8, 8)
        # Panel izquierdo más ancho (feedback del usuario, 2026-08-18) —
        # antes 1:2 (33%/67%), ahora 2:3 (40%/60%).
        cuerpo.addWidget(self._armar_lista(), stretch=2)
        cuerpo.addWidget(self._armar_editor(), stretch=3)
        layout.addWidget(cuerpo_widget, stretch=1)

        barra_botones = self._armar_barra_botones()
        barra_botones.setContentsMargins(8, 4, 8, 8)
        layout.addLayout(barra_botones)

    def _armar_panel_tabla(self) -> QWidget:
        """Panel superior "TABLA:" — combo centrado, letra grande y en
        negrita, en vez de la fila chica "Categoría:" que había antes
        (feedback del usuario, 2026-08-17). Lo demás (comportamiento
        del combo, paneles de abajo) no cambia, sólo el espacio/estilo
        de esta franja."""
        self.panel_tabla = QWidget()
        self.panel_tabla.setStyleSheet(
            f"background-color: {Verde.PASTEL}; border-bottom: 1px solid {Verde.MEDIO_CLARO};"
        )
        layout = QVBoxLayout(self.panel_tabla)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        fila_cat = QHBoxLayout()
        fila_cat.addStretch()
        lbl_tabla = QLabel("TABLA:")
        lbl_tabla.setStyleSheet(f"color: {Verde.OSCURO}; font-size: 15pt; font-weight: bold;")
        fila_cat.addWidget(lbl_tabla)

        self.combo_categoria = QComboBox()
        for _, etiqueta, _ in CATEGORIAS:
            self.combo_categoria.addItem(etiqueta)
        self.combo_categoria.currentIndexChanged.connect(self._on_categoria_cambiada)
        fuente_combo = self.combo_categoria.font()
        # Letra al doble (pedido del usuario, 2026-08-17, segunda ronda
        # — antes era +3pt).
        fuente_combo.setPointSize(fuente_combo.pointSize() * 2)
        fuente_combo.setBold(True)
        self.combo_categoria.setFont(fuente_combo)
        self.combo_categoria.setMinimumWidth(280)
        fila_cat.addWidget(self.combo_categoria)
        fila_cat.addStretch()
        layout.addLayout(fila_cat)

        return self.panel_tabla

    def _armar_lista(self) -> QGroupBox:
        # "Registros" en vez de "Entradas de la categoría" (feedback del
        # usuario, 2026-08-18).
        box = QGroupBox("Registros")
        layout = QVBoxLayout(box)
        self.tabla = TablaBusqueda(["Código", "Descripción"])
        self.tabla.itemClicked.connect(lambda _: self._on_elegir_de_lista())
        layout.addWidget(self.tabla)
        return box

    @staticmethod
    def _form_compacto(panel: QWidget) -> QFormLayout:
        """`QFormLayout` con espaciado a la mitad del default de Qt —
        pedido del usuario, 2026-08-17, segunda ronda: "Achicar los
        paneles inferiores a la mitad" (junto con el panel superior más
        chico, 12.5%, así entra todo sin que la ventana se salga del
        panel central del MDI)."""
        form = QFormLayout(panel)
        form.setVerticalSpacing(4)
        form.setContentsMargins(6, 4, 6, 4)
        return form

    def _armar_editor(self) -> QGroupBox:
        box = QGroupBox("Datos")
        form = self._form_compacto(box)

        self.txt_codigo = QLineEdit()
        self.txt_codigo.setMaxLength(5)
        form.addRow("Código :", self.txt_codigo)

        self.txt_descri = UpperCaseLineEdit()
        self.txt_descri.setMaxLength(50)
        form.addRow("Descripción :", self.txt_descri)

        self.stack = QStackedWidget()
        for ctab, _, _ in CATEGORIAS:
            self.stack.addWidget(self._armar_panel(ctab))
        form.addRow(self.stack)

        return box

    def _armar_panel(self, ctab: str) -> QWidget:
        if ctab == "SC":
            return self._armar_panel_seccion()
        if ctab == "VD":
            return self._armar_panel_vendedor()
        if ctab == "PV":
            return self._armar_panel_provincia()
        if ctab == "UM":
            return self._armar_panel_um()
        if ctab == "CV":
            return self._armar_panel_cond_venta()
        if ctab == "VS":
            return self._armar_panel_varios()
        return QWidget()  # ZN, MT: sin datos adicionales

    def _armar_panel_seccion(self) -> QWidget:
        panel = QWidget()
        form = self._form_compacto(panel)

        self.combos_alf_codigo: list[QComboBox] = []
        for i in range(1, 4):
            combo = QComboBox()
            self.combos_alf_codigo.append(combo)
            form.addRow(f"Especificación {i} (código) :", combo)

        self.combos_alf_precio: list[QComboBox] = []
        for i in range(1, 4):
            combo = QComboBox()
            self.combos_alf_precio.append(combo)
            form.addRow(f"Cálc. de Precio {i} :", combo)

        self.combo_alf7 = QComboBox()
        form.addRow("Unid. de Facturación :", self.combo_alf7)

        self.chk_mca1 = QCheckBox("Usar para Facturar")
        form.addRow(self.chk_mca1)
        self.chk_mca2 = QCheckBox("Precio en Dólares")
        form.addRow(self.chk_mca2)

        return panel

    def _armar_panel_vendedor(self) -> QWidget:
        panel = QWidget()
        form = self._form_compacto(panel)
        self.spin_vd_zona = QSpinBox()
        self.spin_vd_zona.setRange(0, 999)
        form.addRow("Zona :", self.spin_vd_zona)
        self.txt_vd_pct_venta = QLineEdit("0")
        form.addRow("% Comisión Venta :", self.txt_vd_pct_venta)
        self.txt_vd_pct_cobranza = QLineEdit("0")
        form.addRow("% Comisión Cobranza :", self.txt_vd_pct_cobranza)
        return panel

    def _armar_panel_provincia(self) -> QWidget:
        panel = QWidget()
        form = self._form_compacto(panel)
        self.txt_pv_letra = UpperCaseLineEdit()
        self.txt_pv_letra.setMaxLength(1)
        form.addRow("Letra :", self.txt_pv_letra)
        return panel

    def _armar_panel_um(self) -> QWidget:
        panel = QWidget()
        form = self._form_compacto(panel)
        self.spin_um_enteros = QSpinBox()
        self.spin_um_enteros.setRange(0, 9)
        form.addRow("Enteros :", self.spin_um_enteros)
        self.spin_um_decimales = QSpinBox()
        self.spin_um_decimales.setRange(0, 9)
        form.addRow("Decimales :", self.spin_um_decimales)
        self.combo_um_posicion = QComboBox()
        for cod, etiqueta in POSICIONES_UM:
            self.combo_um_posicion.addItem(etiqueta, cod)
        form.addRow("Posición :", self.combo_um_posicion)
        return panel

    def _armar_panel_cond_venta(self) -> QWidget:
        panel = QWidget()
        form = self._form_compacto(panel)
        self.spin_cv_dias = QSpinBox()
        self.spin_cv_dias.setRange(0, 999)
        form.addRow("Cant. Días :", self.spin_cv_dias)
        return panel

    def _armar_panel_varios(self) -> QWidget:
        panel = QWidget()
        form = self._form_compacto(panel)
        self.spin_vs_num1 = QSpinBox()
        self.spin_vs_num1.setRange(0, 999999999)
        form.addRow("Número 1 :", self.spin_vs_num1)
        self.spin_vs_num2 = QSpinBox()
        self.spin_vs_num2.setRange(0, 999999999)
        form.addRow("Número 2 :", self.spin_vs_num2)
        self.txt_vs_pct1 = QLineEdit("0")
        form.addRow("Porcent. 1 :", self.txt_vs_pct1)
        self.txt_vs_pct2 = QLineEdit("0")
        form.addRow("Porcent. 2 :", self.txt_vs_pct2)
        return panel

    # Botones un poco más grandes (feedback del usuario, 2026-08-18) —
    # mismo estilo para los 5, sólo cambia el padding/alto respecto al
    # QPushButton default de la hoja de estilo global.
    _ESTILO_BOTON_GRANDE = "padding: 8px 18px; font-size: 10.5pt;"

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()
        fila.setSpacing(10)  # separados entre sí, no pegados

        self.btn_nuevo = QPushButton("Alta")
        self.btn_nuevo.setStyleSheet(self._ESTILO_BOTON_GRANDE)
        self.btn_nuevo.clicked.connect(self._on_nuevo)
        fila.addWidget(self.btn_nuevo)

        self.btn_editar = QPushButton("Cambio")
        self.btn_editar.setStyleSheet(self._ESTILO_BOTON_GRANDE)
        self.btn_editar.clicked.connect(self._on_toggle_editar)
        fila.addWidget(self.btn_editar)

        self.btn_baja = QPushButton("Baja")
        self.btn_baja.setStyleSheet(self._ESTILO_BOTON_GRANDE)
        self.btn_baja.clicked.connect(self._on_baja)
        fila.addWidget(self.btn_baja)

        # Espacio extra (además del addStretch) para separar bien el
        # grupo Alta/Cambio/Baja del grupo Grabar/Cerrar.
        fila.addSpacing(20)
        fila.addStretch()

        self.btn_guardar = QPushButton("Grabar")
        self.btn_guardar.setStyleSheet(self._ESTILO_BOTON_GRANDE)
        self.btn_guardar.clicked.connect(self._on_guardar)
        fila.addWidget(self.btn_guardar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.setStyleSheet(self._ESTILO_BOTON_GRANDE)
        self.btn_cerrar.clicked.connect(self.close)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    # Cambio de categoría
    # ------------------------------------------------------------------
    def _on_categoria_cambiada(self) -> None:
        idx = self.combo_categoria.currentIndex()
        self.stack.setCurrentIndex(idx)
        ctab, _, _ = self.categoria_actual

        if ctab == "SC":
            self._poblar_combos_um()

        self._refrescar_lista()
        self.fila_actual = None
        self._set_modo(None)
        # El panel de la derecha tiene que blanquearse al cambiar de
        # categoría (feedback del usuario, 2026-08-18) — sin esto
        # quedaban los datos de la categoría anterior a la vista, sin
        # relación con la nueva. Mismo patrón ya usado en `_on_baja`:
        # habilitar momentáneamente para poder tocar los combos/spins,
        # blanquear, volver a deshabilitar (ya lo dejó `_set_modo(None)`).
        self.stack.currentWidget().setEnabled(True)
        self._blanquear()
        self.stack.currentWidget().setEnabled(False)

    def _poblar_combos_um(self) -> None:
        unidades = self.repos.fctablas().by_ctab("UM")
        combos = [*self.combos_alf_codigo, *self.combos_alf_precio, self.combo_alf7]
        for combo in combos:
            combo.clear()
            combo.addItem("(sin usar)", "")
            for u in unidades:
                cod = (u.COD or "").strip()
                combo.addItem(f"{cod} - {u.DESCRI or ''}", cod)

    def _refrescar_lista(self) -> None:
        ctab, _, _ = self.categoria_actual
        filas = [
            ([(fila.COD or "").strip(), fila.DESCRI or ""], fila)
            for fila in self.tabla_service.buscar(ctab)
        ]
        self.tabla.cargar_filas(filas)

    def _on_elegir_de_lista(self) -> None:
        fila: Optional[Fctabla1] = self.tabla.objeto_seleccionado()
        if fila is not None:
            self._cargar_fila(fila)

    # ------------------------------------------------------------------
    # Modos: alta / ver / editar
    # ------------------------------------------------------------------
    def _set_modo(self, modo: Optional[str]) -> None:
        self.modo = modo
        hay_fila = modo is not None
        editable = modo in ("alta", "editar")

        self.txt_codigo.setEnabled(modo == "alta")
        self.txt_descri.setEnabled(editable)
        self.stack.currentWidget().setEnabled(editable)

        self.btn_guardar.setEnabled(editable)
        self.btn_editar.setEnabled(hay_fila)
        self.btn_editar.setText("Sólo Ver" if modo == "editar" else "Cambio")
        self.btn_baja.setEnabled(modo == "ver")

    # ------------------------------------------------------------------
    # Alta
    # ------------------------------------------------------------------
    def _on_nuevo(self) -> None:
        self.fila_actual = None
        self._blanquear()
        self._set_modo("alta")
        self.txt_codigo.setFocus()

    def _blanquear(self) -> None:
        self.txt_codigo.clear()
        self.txt_descri.clear()
        ctab, _, _ = self.categoria_actual
        if ctab == "SC":
            for combo in (*self.combos_alf_codigo, *self.combos_alf_precio, self.combo_alf7):
                combo.setCurrentIndex(0)
            self.chk_mca1.setChecked(False)
            self.chk_mca2.setChecked(False)
        elif ctab == "VD":
            self.spin_vd_zona.setValue(0)
            self.txt_vd_pct_venta.setText("0")
            self.txt_vd_pct_cobranza.setText("0")
        elif ctab == "PV":
            self.txt_pv_letra.clear()
        elif ctab == "UM":
            self.spin_um_enteros.setValue(0)
            self.spin_um_decimales.setValue(0)
            self.combo_um_posicion.setCurrentIndex(0)
        elif ctab == "CV":
            self.spin_cv_dias.setValue(0)
        elif ctab == "VS":
            self.spin_vs_num1.setValue(0)
            self.spin_vs_num2.setValue(0)
            self.txt_vs_pct1.setText("0")
            self.txt_vs_pct2.setText("0")

    # ------------------------------------------------------------------
    # Cargar / Cambio
    # ------------------------------------------------------------------
    def _cargar_fila(self, fila: Fctabla1) -> None:
        self.fila_actual = fila
        self.txt_codigo.setText((fila.COD or "").strip())
        self.txt_descri.setText((fila.DESCRI or "").strip())

        ctab, _, _ = self.categoria_actual
        if ctab == "SC":
            for combo, valor in zip(
                self.combos_alf_codigo, (fila.ALF1, fila.ALF2, fila.ALF3)
            ):
                self._seleccionar_por_data(combo, valor)
            for combo, valor in zip(
                self.combos_alf_precio, (fila.ALF4, fila.ALF5, fila.ALF6)
            ):
                self._seleccionar_por_data(combo, valor)
            self._seleccionar_por_data(self.combo_alf7, fila.ALF7)
            self.chk_mca1.setChecked(bool(fila.MCA1))
            self.chk_mca2.setChecked(bool(fila.MCA2))
        elif ctab == "VD":
            self.spin_vd_zona.setValue(fila.NUMSD1 or 0)
            self.txt_vd_pct_venta.setText(str(fila.NUMCD1 or 0))
            self.txt_vd_pct_cobranza.setText(str(fila.NUMCD2 or 0))
        elif ctab == "PV":
            self.txt_pv_letra.setText((fila.ALF1 or "").strip())
        elif ctab == "UM":
            self.spin_um_enteros.setValue(fila.NUMSD1 or 0)
            self.spin_um_decimales.setValue(fila.NUMSD2 or 0)
            idx = self.combo_um_posicion.findData(str(fila.NUMSD3 or ""))
            self.combo_um_posicion.setCurrentIndex(idx if idx >= 0 else 0)
        elif ctab == "CV":
            self.spin_cv_dias.setValue(fila.NUMSD3 or 0)
        elif ctab == "VS":
            self.spin_vs_num1.setValue(fila.NUMSD1 or 0)
            self.spin_vs_num2.setValue(fila.NUMSD2 or 0)
            self.txt_vs_pct1.setText(str(fila.NUMCD1 or 0))
            self.txt_vs_pct2.setText(str(fila.NUMCD2 or 0))

        self._set_modo("ver")

    @staticmethod
    def _seleccionar_por_data(combo: QComboBox, valor) -> None:
        valor_str = (valor or "").strip() if isinstance(valor, str) else str(valor or "")
        idx = combo.findData(valor_str)
        combo.setCurrentIndex(idx if idx >= 0 else 0)

    def _on_toggle_editar(self) -> None:
        if self.modo == "editar":
            self._set_modo("ver")
        else:
            self._set_modo("editar")
            self.txt_descri.setFocus()

    # ------------------------------------------------------------------
    # Grabar
    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        ctab, _, texto_cod = self.categoria_actual
        codigo_ingresado = self.txt_codigo.text().strip()
        descri = self.txt_descri.text().strip()

        if not codigo_ingresado:
            QMessageBox.warning(self, "ABM de Tablas", "El Código es obligatorio.")
            return
        if not descri:
            QMessageBox.warning(self, "ABM de Tablas", "La Descripción es obligatoria.")
            return

        if texto_cod:
            codigo = codigo_ingresado.upper()
        else:
            try:
                codigo = str(int(codigo_ingresado))
            except ValueError:
                QMessageBox.warning(self, "ABM de Tablas", "El Código debe ser numérico.")
                return

        if ctab == "SC" and not self.combo_alf7.currentData():
            QMessageBox.warning(
                self, "ABM de Tablas", "La Sección no tiene cargada la Unidad de Facturación."
            )
            return

        if self.modo == "alta":
            if self.repos.fctablas().by_ctab_cod(ctab, codigo) is not None:
                QMessageBox.critical(self, "Error de Ingreso", "Error en el Código")
                return

        respuesta = QMessageBox.question(
            self,
            "Grabación en Tablas",
            "Desea Grabar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        datos = self._recolectar_datos(ctab, descri)

        if self.modo == "alta":
            datos.update(CTAB=ctab, COD=codigo)
            nueva = self.repos.fctablas().create(datos)
            self.fila_actual = nueva
        else:
            self.repos.fctablas().update(self.fila_actual.id, datos)

        QMessageBox.information(self, "ABM de Tablas", "Grabado correctamente.")
        self._refrescar_lista()
        self._set_modo("ver")

    def _recolectar_datos(self, ctab: str, descri: str) -> dict:
        datos: dict = dict(DESCRI=descri)

        if ctab == "SC":
            datos.update(
                ALF1=self.combos_alf_codigo[0].currentData() or "",
                ALF2=self.combos_alf_codigo[1].currentData() or "",
                ALF3=self.combos_alf_codigo[2].currentData() or "",
                ALF4=self.combos_alf_precio[0].currentData() or "",
                ALF5=self.combos_alf_precio[1].currentData() or "",
                ALF6=self.combos_alf_precio[2].currentData() or "",
                ALF7=self.combo_alf7.currentData() or "",
                MCA1=1 if self.chk_mca1.isChecked() else 0,
                MCA2=1 if self.chk_mca2.isChecked() else 0,
            )
        elif ctab == "VD":
            datos.update(
                NUMSD1=self.spin_vd_zona.value(),
                NUMCD1=parse_decimal(self.txt_vd_pct_venta.text()),
                NUMCD2=parse_decimal(self.txt_vd_pct_cobranza.text()),
            )
        elif ctab == "PV":
            datos.update(ALF1=self.txt_pv_letra.text().strip())
        elif ctab == "UM":
            datos.update(
                NUMSD1=self.spin_um_enteros.value(),
                NUMSD2=self.spin_um_decimales.value(),
                NUMSD3=int(self.combo_um_posicion.currentData() or 0),
            )
        elif ctab == "CV":
            datos.update(NUMSD3=self.spin_cv_dias.value())
        elif ctab == "VS":
            datos.update(
                NUMSD1=self.spin_vs_num1.value(),
                NUMSD2=self.spin_vs_num2.value(),
                NUMCD1=parse_decimal(self.txt_vs_pct1.text()),
                NUMCD2=parse_decimal(self.txt_vs_pct2.text()),
            )

        return datos

    # ------------------------------------------------------------------
    # Baja
    # ------------------------------------------------------------------
    def _on_baja(self) -> None:
        if self.fila_actual is None:
            return

        ctab, _, _ = self.categoria_actual
        cod = (self.fila_actual.COD or "").strip()
        puede, motivo = self.tabla_service.puede_dar_de_baja(ctab, cod)

        if not puede:
            QMessageBox.critical(self, "ABM de Tablas", f"No se puede dar de baja: {motivo}")
            return

        respuesta = QMessageBox.question(
            self,
            "Grabación en Tablas",
            "Desea Eliminar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.repos.fctablas().delete(self.fila_actual.id)
        QMessageBox.information(self, "ABM de Tablas", "Eliminado.")
        self.fila_actual = None
        self._refrescar_lista()
        self._set_modo(None)
        self.stack.currentWidget().setEnabled(True)
        self._blanquear()
        self.stack.currentWidget().setEnabled(False)

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
