# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

# 'logging.config': bug real encontrado probando el .exe en la PC nueva
# (2026-08-21) — `migration/alembic/env.py` (`from logging.config import
# fileConfig`) sólo se carga DINÁMICAMENTE en runtime (Alembic lo lee
# como archivo de datos, no como módulo importado normal, ver
# `datas=[...]` más abajo), así que el análisis estático de PyInstaller
# nunca ve ese import y no lo empaqueta solo — hay que declararlo a mano.
# Confirmado el fix reproduciendo el mismo `alembic upgrade head` dentro
# de un .exe de prueba descartable (contra SQLite, sin necesitar el
# Postgres real de la PC nueva): con este hidden-import corre 7
# migraciones reales sin error de import.
hiddenimports = ['psycopg2', 'pyodbc', 'logging.config']
hiddenimports += collect_submodules('alembic')
hiddenimports += collect_submodules('sqlalchemy')


a = Analysis(
    ['migration/etl/migrar_completo.py'],
    pathex=['.'],
    binaries=[],
    datas=[('alembic.ini', '.'), ('migration/alembic', 'migration/alembic')],
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
    name='migrar_fcmenu',
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
    # Mismo ícono real de la app (`assets/Icon-Alestel.png`, ver
    # `migration/ui/theme.py icono_app()`) — pedido del usuario
    # (2026-08-21). Convertido a .ico multi-resolución
    # (`assets/Icon-Alestel.ico`, generado con Pillow) porque el .exe
    # de Windows necesita ese formato, un .png solo no alcanza.
    icon='assets/Icon-Alestel.ico',
)
