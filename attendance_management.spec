# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for attendance_management
# Produces a directory-based (onedir) Windows executable with no console window.
# Using onedir instead of onefile avoids extracting files to a temp directory on
# every launch, which significantly reduces startup time.

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "config",
        "attendance",
        "windows_events",
        "openpyxl",
        "openpyxl.cell._writer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    name="attendance_management",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="attendance_management",
)
