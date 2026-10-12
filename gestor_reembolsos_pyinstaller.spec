# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

# Recolectar datos y dependencias de librerías complejas (OCR, PDF, ONNX)
datas = [('templates', 'templates')]
binaries = []
hiddenimports = [
    'procesar_reembolsos',
    'rapidocr_onnxruntime',
    'onnxruntime',
    'pypdfium2',
    'pdfplumber',
    'pdfminer',
    'pdfminer.high_level',
    'pdfminer.layout',
    'cv2',
    'PIL',
    'numpy',
    'tqdm',
    'xml.etree.ElementTree',
    'http.server',
    'socketserver',
    'webbrowser',
    'urllib.parse',
]

# Recolectar submódulos y datos automáticamente
for pkg in ['rapidocr_onnxruntime', 'onnxruntime', 'pypdfium2', 'pdfminer']:
    try:
        pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
        datas += pkg_datas
        binaries += pkg_binaries
        hiddenimports += pkg_hidden
    except Exception as e:
        print(f"Aviso al recolectar {pkg}: {e}")

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'pandas', 'IPython', 'pytest'],
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
    name='Gestor_Reembolsos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Mantener consola visible para ver logs de arranque
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
