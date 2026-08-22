# -*- mode: python ; coding: utf-8 -*-
# Lanzador liviano — pantalla de bienvenida INMEDIATA + arranque de
# `FcMenuII.exe` (`fcmenu.spec`) como proceso aparte. Ver el docstring
# de `migration/ui/launcher.py` para el motivo completo (pedido del
# usuario, 2026-08-22: reemplazar la pantalla negra del arranque).
#
# A proposito SIN `collect_all("PyQt6")` (el que si usa `fcmenu.spec`
# para la app completa) - `collect_all` empaqueta TODO PyQt6 de punta a
# punta (WebEngine, Qt3D, QML, todos los drivers SQL de Qt...) sin
# importar que se usa de verdad, y en la primera prueba (2026-08-22)
# eso dejo este .exe en ~220MB - practicamente el mismo peso que
# `FcMenuII.exe`, exactamente lo que este lanzador tiene que NO ser.
# Sin `collect_all`, PyInstaller sigue los imports reales de
# `launcher.py`/`splash_screen.py`/`theme.py` (QtCore/QtGui/QtWidgets
# nomas - sus hooks internos de PyQt6 alcanzan solos para esos 3
# modulos) y arma un .exe liviano de verdad (40MB). Solo los assets del
# splash (logo + icono) se agregan a mano.
#
# `FcMenuLauncher.exe` se instala en LA MISMA carpeta que `FcMenuII.exe`
# (`launcher.py` lo busca ahi, ver `_comando_programa_principal()`) -
# es el .exe que el acceso directo del escritorio debe apuntar de aca
# en adelante, no `FcMenuII.exe` directo.
#
# **onefile de nuevo** (2026-08-22, tercera vuelta): se probo pasar a
# onedir (carpeta con el .exe + una subcarpeta `_internal/` de DLLs
# sueltas, ~100MB en disco) pensando que la pantalla negra era la
# descompresion propia del onefile - resulto ser OTRA cosa (`FcMenuII.exe`
# seguia con `console=True`: Windows le abria una consola negra apenas
# arrancaba el proceso, bien antes de que corriera Python - ver el
# comentario de `console=False` en `fcmenu.spec`). Con esa causa real ya
# corregida, este lanzador onefile de 40MB arranca solo en una fraccion
# de segundo (confirmado, no hace falta separar en carpeta) - un solo
# .exe para copiar es mas simple de instalar que .exe+carpeta, y no
# tiene sentido cargar con esa complejidad para un problema que ya no
# existe.
datas = [("assets", "assets")]
binaries = []
hiddenimports = []

a = Analysis(
    ["run_launcher.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PyQt6.QtWebEngineWidgets", "PyQt6.QtWebEngineCore", "PyQt6.QtQml", "PyQt6.QtQuick",
        "PyQt6.Qt3DCore", "PyQt6.QtNetwork", "PyQt6.QtSql", "PyQt6.QtPdf", "PyQt6.QtMultimedia",
    ],
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
    name='FcMenuLauncher',
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
