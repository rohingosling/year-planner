# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for the Year Planner Generator — one-file console executable (dist/year-planner.exe).
#
# Bundled data (read at runtime via src/paths.py resource_path()):
#   - config/config.yaml                                       -> config/
#   - assets/images/instructions.png                           -> assets/images/
#   - assets/images/graph_paper_37x56_15_50_2161x3295px.png    -> assets/images/  (matches the default config's grid)
#
# If graph_paper.* / page size / margins change in config.yaml, regenerate and re-bundle the matching grid PNG;
# otherwise the exe safely falls back to generating it once into the writable cache dir on first run.
#
# Build with:  .venv/Scripts/python.exe scripts/build.py   (or: pyinstaller year-planner.spec, run from the repo root)

import os

# Resolve the repo root from the spec's own location, falling back to the working directory.

try:
    root = os.path.abspath(SPECPATH)
except NameError:
    root = os.path.abspath(os.getcwd())


a = Analysis(
    [os.path.join(root, 'src', 'main.py')],
    pathex=[root],
    binaries=[],
    datas=[
        (os.path.join(root, 'config', 'config.yaml'), 'config'),
        (os.path.join(root, 'assets', 'images', 'instructions.png'), 'assets/images'),
        (os.path.join(root, 'assets', 'images', 'graph_paper_37x56_15_50_2161x3295px.png'), 'assets/images'),
    ],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='year-planner',
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
