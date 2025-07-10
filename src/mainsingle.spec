# -*- mode: python ; coding: utf-8 -*-



a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('bin', 'bin'), ('lib', 'lib'), ('pdf.ico', '.')], 
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name='MinimalPDF Compress v1.3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,
    icon='pdf.ico',
)