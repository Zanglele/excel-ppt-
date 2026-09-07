# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import os
import sys

root = Path(SPECPATH).parent
a = Analysis(
    [str(root / "app.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[],
    hiddenimports=["matplotlib.backends.backend_qtagg"],
    hookspath=[],
    hooksconfig={"matplotlib": {"backends": ["QtAgg"]}},
    runtime_hooks=[],
    excludes=["tkinter", "PyQt5", "PySide2", "PySide6"],
    noarchive=False,
)
# 外部软件的 DLL 混入会使开发环境可运行、冻结程序却加载失败。
allowed_roots = tuple(Path(path).resolve() for path in (sys.prefix, sys.base_prefix, os.environ["SystemRoot"]))
for target, source, kind in a.binaries:
    if not any(Path(source).resolve().is_relative_to(path) for path in allowed_roots):
        raise RuntimeError(f"打包依赖来自环境外部，请检查构建 PATH：{source}")
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ExcelPPT",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="ExcelPPT")
