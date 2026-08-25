# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, copy_metadata


datas = [
    ("app.py", "."),
    ("README_팀원용.txt", "."),
    (".streamlit/config.toml", ".streamlit"),
    ("app/data/standard.xlsx", "app/data"),
]
binaries = []
hiddenimports = []

for package in [
    "streamlit",
    "altair",
    "plotly",
    "pandas",
    "numpy",
    "openpyxl",
    "xlrd",
    "pyarrow",
    "pydeck",
    "watchdog",
]:
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hiddenimports

for package in ["streamlit", "altair", "pandas", "numpy", "pyarrow", "plotly"]:
    datas += copy_metadata(package)


a = Analysis(
    ["launcher.py"],
    pathex=[],
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
    name="AGM_Reduction_Analysis",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    name="AGM_Reduction_Analysis",
)
