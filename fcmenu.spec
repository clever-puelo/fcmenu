# -*- mode: python ; coding: utf-8 -*-
# App completa (ERP) — reemplaza a FCMENU.frm/FCMENU.vbp, entry point real
# `migration/ui/main_menu.py` (MainMenuWindow). Ver `migrar_fcmenu.spec`
# para el migrador Access->Postgres (un .exe aparte, no esto).
#
# `console=False` (2026-08-22 — antes `console=True` "a propósito en
# esta primera versión de prueba... pasar a console=False más adelante,
# una vez confirmado que arranca bien": ya se confirmó de punta a punta
# contra Postgres real, varias veces, en esta sesión). Bug real
# encontrado probando el lanzador (`launcher.py`/`fcmenu_launcher.spec`,
# "la pantalla negra sigue apareciendo... más de 5 segundos"): con
# `console=True`, Windows abre la consola (negra, vacía) apenas arranca
# el PROCESO — antes de que una sola línea de Python corra — y se queda
# ahí durante TODA la extracción del onefile + los imports pesados
# (WebEngine/SQLAlchemy/ReportLab/zeep/cryptography) + la conexión a la
# base, tapando/compitiendo con la bienvenida del lanzador durante ese
# rato largo. Sin consola, ese hueco queda cubierto de verdad por la
# bienvenida (que sigue arriba hasta que este proceso avisa que ya
# mostró su ventana real, ver `_avisar_launcher_listo()` en
# `main_menu.py`) — no aparece nada negro de por medio.
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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/Icon-Alestel.ico',
)
