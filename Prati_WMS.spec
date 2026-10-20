# -*- mode: python ; coding: utf-8 -*-

import struct
import sys


if sys.platform != 'win32':
    raise SystemExit('O Prati WMS deve ser compilado no Windows.')
if sys.version_info[:2] != (3, 12):
    raise SystemExit('Use Python 3.12 para manter a compatibilidade definida com Windows 10 e 11.')
if struct.calcsize('P') * 8 != 64:
    raise SystemExit('Use Python 3.12 de 64 bits para gerar o executavel.')


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    name='Prati_WMS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    manifest='Prati_WMS.manifest',
    uac_admin=False,
    uac_uiaccess=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
