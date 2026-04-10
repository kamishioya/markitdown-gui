# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_submodules


project_root = Path(SPECPATH).resolve().parent
entry_script = project_root / "src" / "markitdown_gui" / "__main__.py"
license_file = project_root / "LICENSE"

magika_datas, magika_binaries, magika_hiddenimports = collect_all("magika")
markitdown_hiddenimports = collect_submodules("markitdown")

hiddenimports = list(markitdown_hiddenimports) + list(magika_hiddenimports)
datas = list(magika_datas)
if license_file.exists():
    datas.append((str(license_file), "."))
binaries = list(magika_binaries)


a = Analysis(
    [str(entry_script)],
    pathex=[str(project_root / "src")],
    binaries=binaries,
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
    name="MarkItDownGUI",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="MarkItDownGUI",
)
