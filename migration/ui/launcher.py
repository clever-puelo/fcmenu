"""Lanzador liviano — pantalla de bienvenida INMEDIATA, sin la pantalla
negra de antes, mientras arranca el programa principal en un proceso
aparte.

**Por qué un módulo/proceso aparte** (pedido del usuario, 2026-08-22:
*"La pantalla de bienvenida debe suplantar a la pantalla negra... ahora
primero aparece una pantalla negra y luego de un rato aparece y
desaparece inmediatamente la bienvenida, casi no se ve. El problema es
el tamaño y la demora en cargar... dividirlo en dos módulos es la
manera más elegante y efectiva. Un módulo pequeño, llamador/disparador
con la pantalla de bienvenida, carga el programa principal sin pantalla
negra/ciego, que elimina al llamador, que tiene una barra de progreso
de fantasía"*):

`SplashScreen` (ver `splash_screen.py`) ya existía, pero se creaba
DENTRO de `main_menu.py`, en el cuerpo de `main()` — y para cuando
Python llega a EJECUTAR el cuerpo de esa función, TODOS los imports de
nivel de módulo de `main_menu.py` ya corrieron, entre ellos `from
.main_menu_window import MainMenuWindow`, que en cascada trae TODA la
app real (PyQt6-WebEngine, SQLAlchemy, ReportLab, zeep, cryptography —
el .exe empaquetado pesa ~265MB). Ese import por sí solo tarda varios
segundos, tiempo en el que NINGUNA ventana Qt existe todavía — ahí es
donde Windows pinta la pantalla negra, ANTES de que `SplashScreen`
tuviera la más mínima chance de mostrarse. Ningún truco DENTRO del
mismo proceso evita esto: mientras el import pesado no termine, ese
proceso no puede pintar nada.

La solución real es la que pidió el usuario: DOS procesos.
1. Este módulo — sólo importa PyQt6 (`splash_screen.py`/`theme.py` no
   importan nada pesado, confirmado) — arranca casi instantáneo, pinta
   la bienvenida YA MISMO con la barra "de fantasía" (indeterminada,
   `SplashScreen` ya la tenía) y AHÍ lanza el programa real
   (`FcMenuII.exe`, o `-m migration.ui.main_menu` corriendo desde
   código fuente) como proceso hijo aparte.
2. El programa real, apenas muestra su propia ventana
   (`MainMenuWindow`), avisa que está listo escribiendo un archivo
   "sentinel" (`_avisar_launcher_listo()` en `main_menu.py`) — este
   módulo lo espera sondeando cada `_POLL_MS`, y ahí SÍ se cierra
   ("que elimina al llamador"). Si el aviso nunca llega (falla el
   arranque real, o corre sin el launcher) hay un tope de seguridad
   (`_TIMEOUT_MS`) para no quedar pegado para siempre.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

from .splash_screen import SplashScreen
from .theme import aplicar_tema

# Variable de entorno con la que este lanzador le pasa su propio PID al
# proceso hijo — `main_menu.py` la lee para saber DÓNDE escribir el
# sentinel de "ya estoy listo" (ver `sentinel_path()`).
VAR_ENTORNO_PID_LAUNCHER = "FCMENU_LAUNCHER_PID"

_POLL_MS = 150
_TIMEOUT_MS = 45_000  # tope de seguridad — nunca se queda pegado


def sentinel_path(pid_launcher: str | int) -> Path:
    """Ruta del archivo "sentinel" para un PID de lanzador dado — misma
    función usada acá (para sondearlo) y desde `main_menu.py` (para
    escribirlo), así los dos lados están siempre de acuerdo en dónde
    va."""
    return Path(tempfile.gettempdir()) / f"fcmenu_ready_{pid_launcher}.flag"


def _comando_programa_principal() -> list[str]:
    """Arma el comando del programa real — el .exe empaquetado
    (`FcMenuII.exe`, misma carpeta que este lanzador cuando los dos
    están congelados con PyInstaller) o, corriendo desde código fuente
    (dev), el mismo intérprete con `-m migration.ui.main_menu`."""
    if getattr(sys, "frozen", False):
        exe_principal = Path(sys.executable).parent / "FcMenuII.exe"
        return [str(exe_principal)]
    return [sys.executable, "-m", "migration.ui.main_menu"]


def main() -> int:
    app = QApplication(sys.argv)
    aplicar_tema(app)

    splash = SplashScreen()
    splash.show()
    splash.actualizar_estado("Iniciando…")
    app.processEvents()  # pintar YA — antes de gastar un solo ms más

    comando = _comando_programa_principal()
    if getattr(sys, "frozen", False) and not Path(comando[0]).exists():
        QMessageBox.critical(
            None, "FcMenu II", f"No se encontró el programa principal:\n{comando[0]}"
        )
        return 1

    pid_propio = os.getpid()
    sentinel = sentinel_path(pid_propio)
    sentinel.unlink(missing_ok=True)

    try:
        proceso = subprocess.Popen(
            comando, env={**os.environ, VAR_ENTORNO_PID_LAUNCHER: str(pid_propio)}
        )
    except OSError as exc:
        QMessageBox.critical(None, "FcMenu II", f"No se pudo iniciar el programa principal:\n{exc}")
        return 1

    splash.actualizar_estado("Cargando FcMenu II…")

    inicio = time.monotonic()

    def _chequear() -> None:
        # Ya se mostró la ventana real, o el proceso principal terminó
        # (con o sin error) — en cualquiera de los dos casos no tiene
        # sentido seguir mostrando la bienvenida.
        listo = sentinel.exists()
        terminado = proceso.poll() is not None
        vencido = (time.monotonic() - inicio) * 1000 > _TIMEOUT_MS
        if not (listo or terminado or vencido):
            return
        sentinel.unlink(missing_ok=True)
        app.quit()

    temporizador = QTimer()
    temporizador.timeout.connect(_chequear)
    temporizador.start(_POLL_MS)

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
