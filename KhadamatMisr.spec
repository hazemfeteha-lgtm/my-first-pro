# -*- mode: python ; coding: utf-8 -*-
"""بناء نسخة ويندوز بملف تنفيذي واحد بدون الحاجة لتثبيت Python.

    py -3 -m PyInstaller --noconfirm --clean KhadamatMisr.spec

الناتج: dist\\KhadamatMisr.exe
قاعدة البيانات (center_management.db) تُنشأ/تُقرأ بجوار ملف exe نفسه.
"""

a = Analysis(
    ['center_management.py'],
    pathex=['.'],
    binaries=[],
    datas=[('img', 'img')],
    hiddenimports=['arabic_reshaper', 'bidi', 'bidi.algorithm', 'net'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'numpy', 'pandas', 'matplotlib', 'IPython', 'pydoc_data'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='KhadamatMisr',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # بدون نافذة سوداء
    disable_windowed_traceback=False,
    icon='logo.ico',
)
