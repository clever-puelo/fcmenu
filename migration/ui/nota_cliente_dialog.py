"""Diálogo de Notas de Cliente — migración de `NotaClte.frm`.

Se abre desde el botón "Notas" de `ClienteDetalleDialog` (ya no
deshabilitado). Réplica del comportamiento real: dos cuadros de texto
libre ("Principal"/`NOTA1` y "Adicional"/`NOTA2`); si al Grabar ambos
quedan vacíos, se borra la fila completa en vez de guardar una nota
vacía (NotaClte.frm:215-221).

El botón "Borrar" del legacy está `Visible=False` y no tiene ni siquiera
un `Command3_Click` definido — código muerto real, no se migra. Los
campos `TITULO1-4`/`PIE1-4`/`NOTA3-4` del modelo existen en el esquema
real pero esta pantalla nunca los usó — quedan para una futura función
de impresión, no se tocan acá.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from migration.repository import RepositoryFactory


class NotaClienteDialog(QDialog):
    """Notas libres de un cliente (equivalente a NOTACLTE de NotaClte.frm)."""

    def __init__(self, repos: RepositoryFactory, codigo_cliente: int, parent: QWidget | None = None):
        super().__init__(parent)
        self.repos = repos
        self.codigo_cliente = codigo_cliente

        self.setWindowTitle(f"Notas del Cliente {codigo_cliente}")
        self.resize(560, 300)

        self._construir_ui()
        self._cargar_nota()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        fila_textos = QHBoxLayout()

        grupo_principal = QGroupBox("Principal")
        layout_principal = QVBoxLayout(grupo_principal)
        self.txt_nota1 = QTextEdit()
        layout_principal.addWidget(self.txt_nota1)
        fila_textos.addWidget(grupo_principal)

        grupo_adicional = QGroupBox("Adicional")
        layout_adicional = QVBoxLayout(grupo_adicional)
        self.txt_nota2 = QTextEdit()
        layout_adicional.addWidget(self.txt_nota2)
        fila_textos.addWidget(grupo_adicional)

        layout.addLayout(fila_textos)
        layout.addLayout(self._armar_barra_botones())

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
    def _cargar_nota(self) -> None:
        nota = self.repos.notaclte().by_cliente(self.codigo_cliente)
        if nota is None:
            return
        self.txt_nota1.setPlainText(nota.NOTA1 or "")
        self.txt_nota2.setPlainText(nota.NOTA2 or "")

    # ------------------------------------------------------------------
    def _on_guardar(self) -> None:
        respuesta = QMessageBox.question(
            self,
            "Grabación de Notas",
            "Desea Grabar",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        nota1 = self.txt_nota1.toPlainText().strip()
        nota2 = self.txt_nota2.toPlainText().strip()
        nota_repo = self.repos.notaclte()
        existente = nota_repo.by_cliente(self.codigo_cliente)

        if not nota1 and not nota2:
            if existente is not None:
                nota_repo.delete(existente.id)
        elif existente is None:
            nota_repo.create(dict(CLTE=self.codigo_cliente, NOTA1=nota1, NOTA2=nota2))
        else:
            nota_repo.update(existente.id, dict(NOTA1=nota1, NOTA2=nota2))

        QMessageBox.information(self, "Notas del Cliente", "Notas grabadas correctamente.")
        # Feedback del usuario (2026-08-15): tras Grabar, la pantalla se
        # cierra sola — antes quedaba abierta esperando un "Cerrar" aparte.
        self.accept()
