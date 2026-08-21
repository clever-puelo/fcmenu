# -*- mode: python ; coding: utf-8 -*-
# App completa (ERP) — reemplaza a FCMENU.frm/FCMENU.vbp, entry point real
# `migration/ui/main_menu.py` (MainMenuWindow). Ver `migrar_fcmenu.spec`
# para el migrador Access->Postgres (un .exe aparte, no esto).
#
# `console=True` a propósito en esta primera versión de prueba (ver
# docs/instalacion_pc_nueva.md) — deja ver cualquier traceback/aviso de
# conexión a Postgres en la ventana de consola mientras se valida en la
# PC nueva. Pasar a `console=False` más adelante, una vez confirmado que
# arranca bien, para la versión "limpia" sin caja negra de fondo.
from PyInstaller.utils.hooks import collect_all

datas = [("assets", "assets")]
binaries = []
hiddenimports = ["psycopg2"]

for paquete in ("PyQt6", "reportlab", "zeep", "cryptography"):
    _datas, _binaries, _hidden = collect_all(paquete)
    datas += _datas
    binaries += _binaries
    hiddenimports += _hidden


a = Analysis(
    ["run_app.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FcMenuII',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/Icon-Alestel.ico',
)
