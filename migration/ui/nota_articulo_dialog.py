"""Diálogo de Notas de Artículo — migración de `Notartic.frm`.

Se abre desde el botón "Notas" de `ArticuloDetalleDialog`. El legacy usa
`NOTA3`/`NOTA4` (etiquetadas "Ventas"/"Especial") pese a que el esquema
real tiene 5 juegos completos de `TITULO`/`NOTA`/`PIE` — ver docstring de
`NotaArticulo` en `models.py`. Si al Grabar ambos quedan vacíos, se borra
la fila (no se guarda una nota vacía), igual que `NotaClte.frm`.
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


class NotaArticuloDialog(QDialog):
    """Notas libres de un artículo (equivalente a NOTARTIC de Notartic.frm)."""

    def __init__(
        self, repos: RepositoryFactory, cod1: str, cod2: str, parent: QWidget | None = None
    ):
        super().__init__(parent)
        self.repos = repos
        self.cod1 = cod1
        self.cod2 = cod2

        self.setWindowTitle(f"Notas del Artículo {cod1}-{cod2}")
        self.resize(560, 300)

        self._construir_ui()
        self._cargar_nota()

    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)

        fila_textos = QHBoxLayout()

        grupo_ventas = QGroupBox("Ventas")
        layout_ventas = QVBoxLayout(grupo_ventas)
        self.txt_nota3 = QTextEdit()
        layout_ventas.addWidget(self.txt_nota3)
        fila_textos.addWidget(grupo_ventas)

        grupo_especial = QGroupBox("Especial")
        layout_especial = QVBoxLayout(grupo_especial)
        self.txt_nota4 = QTextEdit()
        layout_especial.addWidget(self.txt_nota4)
        fila_textos.addWidget(grupo_especial)

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
        nota = self.repos.notaarticulo().by_articulo(self.cod1, self.cod2)
        if nota is None:
            return
        self.txt_nota3.setPlainText(nota.NOTA3 or "")
        self.txt_nota4.setPlainText(nota.NOTA4 or "")

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

        nota3 = self.txt_nota3.toPlainText().strip()
        nota4 = self.txt_nota4.toPlainText().strip()
        nota_repo = self.repos.notaarticulo()
        existente = nota_repo.by_articulo(self.cod1, self.cod2)

        if not nota3 and not nota4:
            if existente is not None:
                nota_repo.delete(existente.id)
        elif existente is None:
            nota_repo.create(dict(COD1=self.cod1, COD2=self.cod2, NOTA3=nota3, NOTA4=nota4))
        else:
            nota_repo.update(existente.id, dict(NOTA3=nota3, NOTA4=nota4))

        QMessageBox.information(self, "Notas del Artículo", "Notas grabadas correctamente.")
        # Feedback del usuario (2026-08-15): tras Grabar, la pantalla se
        # cierra sola — antes quedaba abierta esperando un "Cerrar" aparte.
        self.accept()
