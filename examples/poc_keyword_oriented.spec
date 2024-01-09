# -*- mode: python ; coding: utf-8 -*-


added_files = [
    ('D:\\Python\\Python311\\Lib\\site-packages\\tabula\\tabula-1.0.5-jar-with-dependencies.jar', 'tabula'),
    ('*.py', '.')]
a = Analysis(
    ['poc_keyword_oriented.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
a.datas += [
    ('stop_words.txt', 'stop_words.txt', 'DATA')
]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='poc_keyword_oriented',
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
)
