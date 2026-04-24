# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

project_root = Path(globals().get("SPEC", Path.cwd())).resolve().parent.parent

hiddenimports = []
hiddenimports += collect_submodules("config")
hiddenimports += collect_submodules("configs")
hiddenimports += collect_submodules("core")
hiddenimports += collect_submodules("src")
hiddenimports += collect_submodules("data")

datas = [
    (str(project_root / "templates"), "templates"),
    (str(project_root / "config"), "config"),
    (str(project_root / "configs"), "configs"),
]

a = Analysis(
    [str(project_root / "desktop_app.py")],
    pathex=[str(project_root)],
    binaries=[],
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
    [],
    exclude_binaries=True,
    name="DiplomaGenerator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name="DiplomaGenerator",
)
