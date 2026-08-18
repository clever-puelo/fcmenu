"""Diálogo de Detalle de Artículo — panel de Alta/Cambio/Baja de `AbmArt.frm`.

Convención de navegación confirmada con el usuario (2026-08-15), aplicada
a todo el sistema salvo aclaración específica: cada ABM tiene una
**ventana de Búsqueda** (lista de resultados + botón Alta) y una
**ventana/diálogo de Detalle** separada que sólo se abre al elegir un
resultado de la lista o al presionar Alta. Este módulo es el Detalle de
Artículos — la Búsqueda vive en `articulo_busqueda_window.py`.

Alcance de este panel (confirmado con el usuario el 2026-08-14, ver
memoria de sesión `pyqt6_ui_progress.md`):

- El código de artículo (`Articulo.COD2`) no se tipea libre: se arma
  dinámicamente a partir de la Sección elegida (`Fctabla1` CTAB='SC')
  concatenando 1 o 2 segmentos numéricos cuyo ancho sale de la Unidad de
  Medida configurada (`Fctabla1` CTAB='UM'). Ver `articulo_codigo.py`.
- Stock / Stock Mínimo / Stock Anterior se sacaron de esta pantalla
  (Articulo.STOCK no es la fuente de verdad, es Stock.STUNID).
- "Precio Alt. (VTA1)": campo que ninguna venta real lee, se mantiene por
  compatibilidad con datos ya migrados, relabelado para no confundir.
- La Baja bloquea si el artículo tiene movimientos históricos en Fcestad1.
- Al Grabar o dar de Baja con éxito, el diálogo se cierra (igual que el
  legacy: `AbmArt.frm Sub Grabacion()` termina en `Blanquea: Unload Me`
  tanto para Alta/Cambio como para Baja) — la ventana de Búsqueda que lo
  abrió refresca su lista al recuperar el control.
- El botón "Notas" (`Notartic.frm`, migrado 2026-08-15) se habilita sólo
  con un artículo ya cargado (no durante Alta). Si el artículo tiene
  notas guardadas, el diálogo se abre solo al cargarlo (réplica de
  `AbmArt.frm Form_Activate`, igual criterio que Clientes).
"""

from __future__ import annotations

from decimal import Decimal
from typing import Callable, Optional

from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from migration.models import Articulo
from migration.repository import RepositoryFactory
from migration.services import ArticuloService

from .articulo_codigo import (
    SegmentoCodigo,
    armar_codigo,
    descripcion_automatica,
    resolver_segmentos,
    separar_codigo,
)
from .decimals import parse_decimal
from .nota_articulo_dialog import NotaArticuloDialog
from .widgets import EnterAsTabFilter, UpperCaseLineEdit

MAX_SEGMENTOS = 2  # el legacy nunca usa más de 2 (Text2/Text3, ALF3 es código muerto)


class ArticuloDetalleDialog(QDialog):
    """Alta / Cambio / Baja de UN artículo. Se abre desde
    `ArticuloBusquedaWindow`, nunca de forma independiente."""

    def __init__(
        self,
        repos: RepositoryFactory,
        articulo_service: ArticuloService,
        articulo: Optional[Articulo] = None,
        seccion_sugerida: Optional[str] = None,
        parent: QWidget | None = None,
        *,
        on_buscar: Optional[Callable[[], None]] = None,
    ):
        super().__init__(parent)
        self.repos = repos
        self.articulo_service = articulo_service
        self.articulo_actual = articulo
        self.segmentos_actuales: list[SegmentoCodigo] = []
        self.guardado_o_eliminado = False  # la ventana que abre el diálogo refresca si esto queda True
        self._on_buscar = on_buscar

        self.setWindowTitle("Detalle de Artículo")
        self.resize(560, 420)

        self._construir_ui()
        self._cargar_secciones()

        if articulo is not None:
            self._cargar_articulo(articulo, modo_inicial="ver")
        else:
            self._blanquear()
            self._set_modo("alta")
            if seccion_sugerida:
                idx = self.combo_seccion.findData(seccion_sugerida)
                if idx >= 0:
                    self.combo_seccion.setCurrentIndex(idx)
            self.combo_seccion.setFocus()

        self._enter_as_tab = EnterAsTabFilter(self)

    # ------------------------------------------------------------------
    # Construcción de la UI
    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        self.grupo_codigo = self._armar_grupo_codigo()
        layout.addWidget(self.grupo_codigo)

        self.frame_item = self._armar_frame_item()
        layout.addWidget(self.frame_item)

        layout.addStretch()
        layout.addLayout(self._armar_barra_botones())

    def _armar_grupo_codigo(self) -> QGroupBox:
        box = QGroupBox("Código de Artículo")
        fila = QHBoxLayout(box)

        fila.addWidget(QLabel("Sección :"))
        self.combo_seccion = QComboBox()
        self.combo_seccion.currentIndexChanged.connect(self._on_seccion_cambiada)
        fila.addWidget(self.combo_seccion)

        self.labels_segmento: list[QLabel] = []
        self.campos_segmento: list[QLineEdit] = []
        for _ in range(MAX_SEGMENTOS):
            etiqueta = QLabel()
            campo = QLineEdit()
            campo.setMaximumWidth(120)
            self.labels_segmento.append(etiqueta)
            self.campos_segmento.append(campo)
            fila.addWidget(etiqueta)
            fila.addWidget(campo)

        fila.addStretch()
        return box

    def _armar_frame_item(self) -> QGroupBox:
        box = QGroupBox("Datos del Ítem")
        form = QFormLayout(box)

        self.txt_descri = UpperCaseLineEdit()
        self.txt_descri.setMaxLength(25)
        self.txt_descri.setToolTip("Descripción (o dejar en blanco para autogenerarla)")
        form.addRow("Descripción :", self.txt_descri)

        self.txt_precio = QLineEdit("0")
        self.txt_precio.setToolTip("Precio de Venta (el que usa la facturación)")
        form.addRow("Precio :", self.txt_precio)

        self.txt_precio_alt = QLineEdit("0")
        self.txt_precio_alt.setToolTip(
            "Campo VTA1 del legacy. No lo usa ninguna pantalla de venta actual "
            "(sólo el importador batch de listas de precios)."
        )
        form.addRow("Precio Alt. (VTA1) :", self.txt_precio_alt)

        self.txt_costo = QLineEdit("0")
        self.txt_costo.setToolTip("Precio de Costo")
        form.addRow("Precio de Costo :", self.txt_costo)

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
            self.btn_buscar.setToolTip("Buscar otro artículo sin cerrar esta ficha")
            self.btn_buscar.clicked.connect(self._on_buscar)
            fila.addWidget(self.btn_buscar)

        fila.addStretch()

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
    # Carga de Secciones (Fctabla1 CTAB='SC' — igual que Form_Load)
    # ------------------------------------------------------------------
    def _cargar_secciones(self) -> None:
        secciones = self.repos.fctablas().by_ctab("SC")
        self.combo_seccion.clear()
        for fila in secciones:
            cod = (fila.COD or "").strip()
            self.combo_seccion.addItem(f"{cod} - {fila.DESCRI or ''}", cod)

    # ------------------------------------------------------------------
    # Segmentos dinámicos del código (Sub BuscaSeccion del legacy)
    # ------------------------------------------------------------------
    def _on_seccion_cambiada(self) -> None:
        cod_seccion = self.combo_seccion.currentData()
        if not cod_seccion:
            self.segmentos_actuales = []
        else:
            self.segmentos_actuales = resolver_segmentos(self.repos.fctablas(), cod_seccion)
        self._refrescar_campos_segmento()

    def _refrescar_campos_segmento(self, valores: Optional[list[Decimal]] = None) -> None:
        for i in range(MAX_SEGMENTOS):
            visible = i < len(self.segmentos_actuales)
            self.labels_segmento[i].setVisible(visible)
            self.campos_segmento[i].setVisible(visible)
            if visible:
                segmento = self.segmentos_actuales[i]
                self.labels_segmento[i].setText(f"{segmento.etiqueta} :")
                texto = ""
                if valores and i < len(valores):
                    texto = self._formatear_valor_segmento(valores[i], segmento)
                self.campos_segmento[i].setText(texto)
            else:
                self.campos_segmento[i].clear()

    @staticmethod
    def _formatear_valor_segmento(valor: Decimal, segmento: SegmentoCodigo) -> str:
        if segmento.libre or segmento.numsd2 == 0:
            return str(int(valor))
        return str(valor)

    # ------------------------------------------------------------------
    # Modos: alta / ver / editar (Command5/Command4 del legacy)
    # ------------------------------------------------------------------
    def _set_modo(self, modo: str) -> None:
        self.modo = modo
        editable = modo in ("alta", "editar")
        clave_editable = modo == "alta"

        self.frame_item.setEnabled(editable)
        self.combo_seccion.setEnabled(clave_editable)
        for campo in self.campos_segmento:
            campo.setEnabled(clave_editable)

        self.btn_guardar.setEnabled(editable)
        self.btn_editar.setText("Sólo Ver" if modo == "editar" else "Cambio")
        self.btn_baja.setEnabled(modo == "ver")
        self.btn_notas.setEnabled(self.articulo_actual is not None)

    def _blanquear(self) -> None:
        """Réplica de Sub Blanquea() (AbmArt.frm:1100-1119)."""
        if self.combo_seccion.count():
            self.combo_seccion.setCurrentIndex(0)
        self.txt_descri.clear()
        self.txt_precio.setText("0")
        self.txt_precio_alt.setText("0")
        self.txt_costo.setText("0")
        self._on_seccion_cambiada()

    # ------------------------------------------------------------------
    def cargar_articulo(self, articulo: Articulo) -> None:
        """Punto de entrada público: carga `articulo` en una ficha ya
        visible (usado por `MainMenuWindow` después de que el operador
        elige uno en la búsqueda que se abre encima)."""
        self._cargar_articulo(articulo, modo_inicial="ver")

    def _cargar_articulo(self, articulo: Articulo, modo_inicial: str = "ver") -> None:
        self.articulo_actual = articulo
        cod1 = (articulo.COD1 or "").strip()

        idx = self.combo_seccion.findData(cod1)
        self.combo_seccion.setCurrentIndex(idx if idx >= 0 else 0)
        self.segmentos_actuales = resolver_segmentos(self.repos.fctablas(), cod1)
        valores = separar_codigo(articulo.COD2 or "", self.segmentos_actuales)
        self._refrescar_campos_segmento(valores)

        self.txt_descri.setText((articulo.DESCRI or "").strip())
        self.txt_precio.setText(str(articulo.PREC if articulo.PREC is not None else 0))
        self.txt_precio_alt.setText(str(articulo.VTA1 if articulo.VTA1 is not None else 0))
        self.txt_costo.setText(str(articulo.PCOS if articulo.PCOS is not None else 0))

        self._set_modo(modo_inicial)

        # Réplica de AbmArt.frm Form_Activate: si el artículo tiene una
        # nota guardada, se muestra sola al cargarlo.
        if self.repos.notaarticulo().by_articulo(cod1, (articulo.COD2 or "").strip()) is not None:
            self._on_notas()

    def _on_notas(self) -> None:
        if self.articulo_actual is None:
            return
        cod1 = (self.articulo_actual.COD1 or "").strip()
        cod2 = (self.articulo_actual.COD2 or "").strip()
        dialogo = NotaArticuloDialog(self.repos, cod1, cod2, parent=self)
        dialogo.exec()

    def _on_toggle_editar(self) -> None:
        if self.modo == "editar":
            self._set_modo("ver")
        else:
            self._set_modo("editar")
            self.txt_descri.setFocus()

    # ------------------------------------------------------------------
    # Grabar (Sub Grabacion() del legacy, camino Alta/Cambio)
    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        cod1 = self.combo_seccion.currentData()
        if not cod1:
            QMessageBox.warning(self, "ABM de Artículos", "Elegí una Sección.")
            return

        if self.modo == "alta":
            if not self.segmentos_actuales:
                QMessageBox.warning(self, "ABM de Artículos", "La Sección elegida no tiene Unidad de Medida configurada.")
                return
            try:
                valores = [
                    parse_decimal(self.campos_segmento[i].text())
                    for i in range(len(self.segmentos_actuales))
                ]
                cod2 = armar_codigo(self.segmentos_actuales, valores)
            except (ValueError, ArithmeticError):
                QMessageBox.warning(self, "ABM de Artículos", "Código inválido.")
                return

            if self.repos.articulo().by_cod1_cod2(cod1, cod2) is not None:
                QMessageBox.critical(self, "Error de Ingreso", "Código  YA  Existe")
                return
        else:
            cod2 = (self.articulo_actual.COD2 or "").strip()

        descri = self.txt_descri.text().strip() or descripcion_automatica(cod1, cod2)

        respuesta = QMessageBox.question(
            self,
            "ABM de ITEMS",
            "Desea Grabar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        datos = dict(
            DESCRI=descri,
            PREC=parse_decimal(self.txt_precio.text()),
            VTA1=parse_decimal(self.txt_precio_alt.text()),
            PCOS=parse_decimal(self.txt_costo.text()),
        )

        if self.modo == "alta":
            datos.update(
                COD1=cod1,
                COD2=cod2,
                VTA2=Decimal("0"),
                DTO1=Decimal("0"),
                STOCK=0,
                STMIN=0,
                STANT=0,
                MCALIS=0,
            )
            nuevo = self.repos.articulo().create(datos)
            self.articulo_actual = nuevo
        else:
            self.repos.articulo().update(self.articulo_actual.id, datos)

        QMessageBox.information(self, "ABM de Artículos", "Artículo grabado correctamente.")
        self.guardado_o_eliminado = True
        self.close()

    # ------------------------------------------------------------------
    # Baja — con la validación nueva (confirmada con el usuario) que el
    # legacy no tenía: bloquear si hay movimientos en Fcestad1.
    # ------------------------------------------------------------------
    def _on_baja(self) -> None:
        if self.articulo_actual is None:
            return

        cod1 = (self.articulo_actual.COD1 or "").strip()
        cod2 = (self.articulo_actual.COD2 or "").strip()
        puede, motivo = self.articulo_service.puede_dar_de_baja(cod1, cod2)

        if not puede:
            QMessageBox.critical(self, "ABM de Artículos", f"No se puede dar de baja: {motivo}")
            return

        respuesta = QMessageBox.question(
            self,
            "ABM de ITEMS",
            "Desea Eliminar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        self.repos.articulo().delete(self.articulo_actual.id)
        QMessageBox.information(self, "ABM de Artículos", "Artículo eliminado.")
        self.guardado_o_eliminado = True
        self.close()
