"""Entry point real del .exe de la app completa (`fcmenu.spec`).

`migration/ui/main_menu.py` usa imports relativos (`from .main_menu_window
import ...`, normal dentro de un paquete) — PyInstaller, al tomarlo como
script TOP-LEVEL directo, lo ejecuta sin contexto de paquete y esos
imports relativos rompen (`ImportError: attempted relative import with
no known parent package`, confirmado corriendo el .exe una primera vez).
Este wrapper es el fix estándar: un script de afuera del paquete que
importa `migration.ui.main_menu` de forma ABSOLUTA (sí funciona, carga
el paquete entero con su contexto normal) y sólo llama a `main()`.

No se usa corriendo desde código fuente (`.venv/Scripts/python.exe -m
migration.ui.main_menu` sigue siendo la forma normal de desarrollo) —
sólo como target de `pyinstaller fcmenu.spec`.
"""
import sys

from migration.ui.main_menu import main

if __name__ == "__main__":
    sys.exit(main())
