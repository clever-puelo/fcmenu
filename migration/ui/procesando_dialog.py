"""`ejecutar_con_progreso()` — helper reutilizable para "todo proceso
que implique demora" (pedido del usuario, 2026-08-18): mientras corre,
muestra un `QProgressDialog` ("Procesando...", barra indeterminada +
botón Cancelar). Convención de sistema #10 (ver `widgets.py` para las
9 anteriores) — cualquier pantalla nueva que dispare una consulta o un
reporte que pueda demorar debe pasar por acá en vez de bloquear la UI
en silencio.

**Cómo corre**: la función recibida corre en un `QThread` aparte (no en
el hilo de la UI) para que la ventana de "Procesando..." se pinte y
Cancelar sea clickeable de verdad mientras tanto. **Importante**: la
función NO debe tocar objetos Qt (widgets) ni la sesión de SQLAlchemy
del hilo principal — si necesita la base, tiene que abrir su propia
sesión (`migration.db.get_session()`), las `Session` de SQLAlchemy no
son seguras para compartir entre threads. Ver cada uso real
(`CobranzasZonaWindow._on_zona_cambiada`, `ListadosWindow`) para el
patrón exacto.

**Cancelar**: como las consultas reales de este sistema son contra una
base local (Postgres/SQLite en la misma red, sin datasets masivos), no
hay puntos de corte cooperativos naturales adentro de cada consulta —
Cancelar mata el hilo (`QThread.terminate()`), aceptable acá porque
todo lo que pasa por este helper es de sólo lectura (consultas/reportes,
nunca una grabación) y usa una sesión de base PROPIA del hilo (no la
del resto de la ventana), así que no deja nada a medio escribir."""

from __future__ import annotations

from typing import Callable, Optional, TypeVar

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import QApplication, QProgressDialog, QWidget

T = TypeVar("T")


class _HiloTrabajo(QThread):
    terminado = pyqtSignal(object)
    fallo = pyqtSignal(Exception)

    def __init__(self, funcion: Callable[[], T], parent: QWidget | None = None):
        super().__init__(parent)
        self._funcion = funcion

    def run(self) -> None:  # noqa: N802 (override de Qt/QThread)
        try:
            resultado = self._funcion()
        except Exception as exc:  # noqa: BLE001 — se re-lanza en el hilo principal, ver abajo
            self.fallo.emit(exc)
        else:
            self.terminado.emit(resultado)


def ejecutar_con_progreso(
    parent: QWidget,
    texto: str,
    funcion: Callable[[], T],
    *,
    titulo: str = "Procesando...",
) -> Optional[T]:
    """Corre `funcion` (sin argumentos) en un hilo aparte mostrando
    "Procesando..." con barra indeterminada + Cancelar. Devuelve el
    resultado de `funcion`, o `None` si el usuario canceló. Si
    `funcion` lanza una excepción real (no una cancelación), se
    re-lanza acá mismo después de cerrar el diálogo — el llamador la
    maneja igual que si hubiera llamado a `funcion()` directo."""
    dialogo = QProgressDialog(texto, "Cancelar", 0, 0, parent)
    dialogo.setWindowTitle(titulo)
    dialogo.setWindowModality(Qt.WindowModality.WindowModal)
    dialogo.setMinimumDuration(0)
    dialogo.setAutoClose(False)
    dialogo.setAutoReset(False)

    resultado_caja: dict = {}
    hilo = _HiloTrabajo(funcion, parent)

    def _on_terminado(resultado: object) -> None:
        resultado_caja["ok"] = resultado
        dialogo.close()

    def _on_fallo(exc: Exception) -> None:
        resultado_caja["error"] = exc
        dialogo.close()

    hilo.terminado.connect(_on_terminado)
    hilo.fallo.connect(_on_fallo)
    dialogo.canceled.connect(hilo.terminate)

    # Se muestra y se pinta ANTES de arrancar el hilo (feedback del
    # usuario, 2026-08-18, tercera ronda: "el aviso de que está
    # procesando lo tenía y ahora no") — carrera real posible: si
    # `funcion` es rápida, el hilo puede terminar (y `dialogo.close()`)
    # en la primera vuelta del loop de eventos que abre `dialogo.exec()`,
    # antes de que el diálogo llegue a pintarse en pantalla siquiera.
    # `show()` + `processEvents()` fuerza el primer pintado ANTES de que
    # el hilo tenga la chance de terminar.
    dialogo.show()
    QApplication.processEvents()

    hilo.start()
    dialogo.exec()
    hilo.wait(3000)

    if "error" in resultado_caja:
        raise resultado_caja["error"]
    return resultado_caja.get("ok")
