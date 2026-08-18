"""Parámetros globales — migración de `Paramet.frm`.

A diferencia de Cliente/Artículo/Tablas, acá no hay lista para elegir: es
una única fila de configuración (`Parametro` CLAVE='1', ver
`ParametroRepository.get_config()`), así que la pantalla es un editor
directo — carga al abrir, Grabar guarda, Cerrar cierra.

Alcance confirmado con el usuario (2026-08-15, ver memoria de sesión
`pyqt6_ui_progress.md`):

- **Se omiten enteros Frame3 ("Datos": Driver/DSN/Carpeta de Backup) y
  Frame6 ("e-mail de Resguardo")** del legacy: no usan la tabla
  `Parametro` — se guardaban en el Registro de Windows
  (`HKEY_CURRENT_USER`) y sólo eran accesibles con un atajo secreto
  (Shift+F12). Era la configuración de un backup automático del `.mdb`
  por Gmail, confirmado en desuso y sin sentido con PostgreSQL. El
  usuario pidió explícitamente quitarlos; un backup automático moderno
  se diseñará aparte más adelante.
- **Se mantienen los 8 "Números de Comprobante" tal cual** (Factura A/B,
  N.Créd. A/B, N.Déb. A/B, Recibo, Cotización) aunque Factura/NC/ND van a
  terminar gobernados por AFIP/WSFEv1 (`FECompUltimoAutorizado`) cuando
  se migre el facturador — el usuario prefirió no sacarlos todavía.
- Los `NUME9..NUME19` del legacy (12 cajas ocultas, `Visible=False` en
  el propio .frm) son código muerto real — ni siquiera se muestran ahí,
  se descartan sin preguntar (mismo criterio que otros controles
  fantasma ya encontrados: invisibles+deshabilitados en el propio
  original = no hace falta migrarlos).
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from migration.db import get_session
from migration.repository import RepositoryFactory

from .decimals import parse_decimal
from .widgets import EnterAsTabFilter, UpperCaseLineEdit


class ParametrosWindow(QMainWindow):
    """Editor de Parámetros globales (equivalente a Paramet de Paramet.frm)."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Parámetros")
        self.resize(640, 420)

        self.db = get_session()
        self.repos = RepositoryFactory(self.db)

        self._construir_ui()
        self._cargar_datos()

        self._enter_as_tab = EnterAsTabFilter(self)

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(self._armar_grupo_comprobantes())

        fila_media = QHBoxLayout()
        fila_media.addWidget(self._armar_grupo_iva())
        fila_media.addWidget(self._armar_grupo_punto_venta())
        layout.addLayout(fila_media)

        layout.addWidget(self._armar_grupo_varios())

        layout.addStretch()
        layout.addLayout(self._armar_barra_botones())

    def _armar_grupo_comprobantes(self) -> QGroupBox:
        box = QGroupBox("Números de Comprobante")
        form = QFormLayout(box)

        self.txt_factura_a = QLineEdit("0")
        form.addRow("Factura A :", self.txt_factura_a)
        self.txt_nc_a = QLineEdit("0")
        form.addRow("N.Créd. A :", self.txt_nc_a)
        self.txt_nd_a = QLineEdit("0")
        form.addRow("N.Déb. A :", self.txt_nd_a)
        self.txt_cotizacion = QLineEdit("0")
        self.txt_cotizacion.setToolTip("Último Nro. de Cotización impreso")
        form.addRow("Cotizac. :", self.txt_cotizacion)
        self.txt_factura_b = QLineEdit("0")
        form.addRow("Factura B :", self.txt_factura_b)
        self.txt_nc_b = QLineEdit("0")
        form.addRow("N.Créd. B :", self.txt_nc_b)
        self.txt_nd_b = QLineEdit("0")
        form.addRow("N.Déb. B :", self.txt_nd_b)
        self.txt_recibo = QLineEdit("0")
        form.addRow("Recibo :", self.txt_recibo)

        return box

    def _armar_grupo_iva(self) -> QGroupBox:
        box = QGroupBox("Porcentaje de I.V.A.")
        form = QFormLayout(box)
        self.txt_iva_inscripto = QLineEdit("0")
        form.addRow("Inscripto :", self.txt_iva_inscripto)
        self.txt_iva_no_inscripto = QLineEdit("0")
        self.txt_iva_no_inscripto.setToolTip(
            "Sólo referencia histórica: Responsable No Inscripto fue eliminado por AFIP en 2003."
        )
        form.addRow("No Inscripto :", self.txt_iva_no_inscripto)
        return box

    def _armar_grupo_punto_venta(self) -> QGroupBox:
        box = QGroupBox("Punto de Venta")
        form = QFormLayout(box)
        self.txt_punto_venta = QLineEdit("0")
        form.addRow("Nro. :", self.txt_punto_venta)
        return box

    def _armar_grupo_varios(self) -> QGroupBox:
        box = QGroupBox("Varios")
        form = QFormLayout(box)

        self.txt_empresa = UpperCaseLineEdit()
        self.txt_empresa.setMaxLength(80)
        form.addRow("Empresa :", self.txt_empresa)

        self.txt_monto_deuda = QLineEdit("0")
        self.txt_monto_deuda.setToolTip("Monto Mínimo de Deuda para Aviso en el Facturador")
        form.addRow("Monto Deuda :", self.txt_monto_deuda)

        self.txt_dias_vto = QLineEdit("0")
        self.txt_dias_vto.setToolTip("Cantidad de días para aviso en el Facturador")
        form.addRow("Días Vto. :", self.txt_dias_vto)

        fila_percep = QHBoxLayout()
        self.radio_percep_si = QRadioButton("Sí")
        self.radio_percep_no = QRadioButton("No")
        self.radio_percep_no.setChecked(True)
        fila_percep.addWidget(self.radio_percep_si)
        fila_percep.addWidget(self.radio_percep_no)
        fila_percep.addStretch()
        form.addRow("C/Percep. IIBB :", fila_percep)

        return box

    def _armar_barra_botones(self) -> QHBoxLayout:
        fila = QHBoxLayout()
        fila.addStretch()

        self.btn_guardar = QPushButton("Grabar")
        self.btn_guardar.clicked.connect(self._on_guardar)
        fila.addWidget(self.btn_guardar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.close)
        fila.addWidget(self.btn_cerrar)

        return fila

    # ------------------------------------------------------------------
    def _cargar_datos(self) -> None:
        """Réplica de Sub CargaDatos() (Paramet.frm:1297-1344), sin la
        parte de Registro de Windows (Frame3/Frame6, ver docstring del
        módulo)."""
        config = self.repos.parametro().get_config()
        if config is None:
            return  # se deja todo en 0/blanco, igual que el legacy sin fila

        self.txt_factura_a.setText(str(config.NUME1 or 0))
        self.txt_nc_a.setText(str(config.NUME2 or 0))
        self.txt_nd_a.setText(str(config.NUME3 or 0))
        self.txt_cotizacion.setText(str(config.NUME4 or 0))
        self.txt_factura_b.setText(str(config.NUME5 or 0))
        self.txt_nc_b.setText(str(config.NUME6 or 0))
        self.txt_nd_b.setText(str(config.NUME7 or 0))
        self.txt_recibo.setText(str(config.NUME8 or 0))

        self.txt_iva_inscripto.setText(str(config.IVAINS if config.IVAINS is not None else 0))
        self.txt_iva_no_inscripto.setText(str(config.IVANI if config.IVANI is not None else 0))

        self.txt_punto_venta.setText((config.PTOVTA or "0").strip())

        self.txt_empresa.setText((config.NOMEMPR or "").strip())
        self.txt_monto_deuda.setText(str(config.LIMVTA if config.LIMVTA is not None else 0))
        self.txt_dias_vto.setText(str(config.NUME20 or 0))

        if config.MCAIB == 1:
            self.radio_percep_si.setChecked(True)
        else:
            self.radio_percep_no.setChecked(True)

    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        respuesta = QMessageBox.question(
            self,
            "Actualización de Parámetros",
            "Desea continuar ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        def entero(campo: QLineEdit) -> int:
            try:
                return int(campo.text().strip() or 0)
            except ValueError:
                return 0

        datos = dict(
            NUME1=entero(self.txt_factura_a),
            NUME2=entero(self.txt_nc_a),
            NUME3=entero(self.txt_nd_a),
            NUME4=entero(self.txt_cotizacion),
            NUME5=entero(self.txt_factura_b),
            NUME6=entero(self.txt_nc_b),
            NUME7=entero(self.txt_nd_b),
            NUME8=entero(self.txt_recibo),
            IVAINS=parse_decimal(self.txt_iva_inscripto.text()),
            IVANI=parse_decimal(self.txt_iva_no_inscripto.text()),
            PTOVTA=str(entero(self.txt_punto_venta)),
            NOMEMPR=self.txt_empresa.text().strip(),
            LIMVTA=parse_decimal(self.txt_monto_deuda.text()),
            NUME20=entero(self.txt_dias_vto),
            MCAIB=1 if self.radio_percep_si.isChecked() else 2,
        )

        parametro_repo = self.repos.parametro()
        config = parametro_repo.get_config()
        if config is None:
            parametro_repo.create({**datos, "CLAVE": "1"})
        else:
            parametro_repo.update(config.id, datos)

        QMessageBox.information(self, "Parámetros", "Parámetros grabados correctamente.")

    # ------------------------------------------------------------------
    def closeEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.db.close()
        super().closeEvent(event)
