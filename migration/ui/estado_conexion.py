"""Chequeo de conectividad a Internet en segundo plano — usado por el
pie de `MainMenuWindow` para mostrar "Hay Internet"/"Sin Internet" sin
congelar la UI.

`socket.create_connection` es bloqueante (hasta `TIMEOUT_SEGUNDOS` si el
host no responde), por eso el chequeo corre en un `threading.Thread` en
vez del hilo principal. Deliberadamente NO se usa `QThreadPool`/
`QRunnable`: combinar `QObject`+`QRunnable` con `autoDelete` es un
patrón frágil en PyQt6 — la primera versión de este módulo lo hacía así
y un smoke test real (instanciar `MainMenuWindow` y dejar terminar el
proceso con un chequeo en vuelo) crasheaba el proceso de punta a punta
de forma reproducible (segfault). `MonitorInternet` es un `QObject`
*persistente* (una sola instancia, dueña de la señal, reusada en cada
chequeo en vez de crear un `QRunnable` descartable por tick) — patrón
mucho más robusto para este caso."""

from __future__ import annotations

import socket
import threading

from PyQt6.QtCore import QObject, pyqtSignal

# DNS público de Google, puerto 53 — chequeo clásico de conectividad que no
# depende de que la resolución de nombres funcione (a diferencia de pingear
# un hostname), ni de que un sitio en particular esté arriba.
HOST_CHEQUEO = ("8.8.8.8", 53)
TIMEOUT_SEGUNDOS = 1.5


class MonitorInternet(QObject):
    """Dispara chequeos puntuales de conectividad en segundo plano y
    avisa el resultado por la señal `resultado` (entregada en el hilo
    del receptor vía el mecanismo normal de señales/slots de Qt)."""

    resultado = pyqtSignal(bool)

    def verificar(self) -> None:
        """No bloquea — el resultado llega después por `resultado`."""
        threading.Thread(target=self._chequear, daemon=True).start()

    def _chequear(self) -> None:
        try:
            with socket.create_connection(HOST_CHEQUEO, timeout=TIMEOUT_SEGUNDOS):
                hay_internet = True
        except OSError:
            hay_internet = False
        try:
            self.resultado.emit(hay_internet)
        except RuntimeError:
            pass  # la ventana ya se cerró (objeto Qt destruido) mientras el chequeo estaba en vuelo
