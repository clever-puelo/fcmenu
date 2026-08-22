"""Entry point real del .exe del lanzador liviano (`fcmenu_launcher.spec`).

Mismo motivo que `run_app.py`: `migration/ui/launcher.py` usa imports
relativos, así que PyInstaller necesita un script AFUERA del paquete
que importe `migration.ui.launcher` de forma absoluta.

No se usa corriendo desde código fuente (`.venv/Scripts/python.exe -m
migration.ui.launcher` sigue siendo la forma normal de probarlo en
dev) — sólo como target de `pyinstaller fcmenu_launcher.spec`.
"""
import sys

from migration.ui.launcher import main

if __name__ == "__main__":
    sys.exit(main())
