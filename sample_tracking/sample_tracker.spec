# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect all for packages that need special handling
pandas_datas, pandas_binaries, pandas_hiddenimports = collect_all('pandas')
numpy_datas, numpy_binaries, numpy_hiddenimports = collect_all('numpy')
customtkinter_datas, customtkinter_binaries, customtkinter_hiddenimports = collect_all('customtkinter')
pyodbc_datas, pyodbc_binaries, pyodbc_hiddenimports = collect_all('pyodbc')
tkcalendar_datas, tkcalendar_binaries, tkcalendar_hiddenimports = collect_all('tkcalendar')

# Get the script directory
script_dir = os.path.dirname(os.path.abspath('sampletracking.py'))

a = Analysis(
    ['sampletracking.py'],
    pathex=[script_dir],
    binaries=pandas_binaries + numpy_binaries + customtkinter_binaries + pyodbc_binaries + tkcalendar_binaries,
    datas=pandas_datas + numpy_datas + customtkinter_datas + pyodbc_datas + tkcalendar_datas,
    hiddenimports=pandas_hiddenimports + numpy_hiddenimports + customtkinter_hiddenimports + pyodbc_hiddenimports + tkcalendar_hiddenimports + ['pyodbc', 'babel.numbers'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SampleTracker',
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
)