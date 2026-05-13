# PyInstaller spec for the Amrita Lab Billing Tool.
#
# Build:
#   pyinstaller build.spec --noconfirm --clean
#
# Output: dist/AmritaLabBilling/AmritaLabBilling.exe (one-folder mode).
#
# Notes
# -----
# • Live MRS, addendums, backups, output and app.log all live next to the .exe
#   (writable, persistent across app updates). See config._user_root().
# • Read-only templates (Working_Sheet.xlsx, Invoice_Bill.xlsx, the bundled
#   "first-run" copy of Master_Rate_Sheet.xlsx) ride inside the bundle and are
#   addressed via sys._MEIPASS.
# • pandas/numpy/openpyxl are pulled in with collect_all() to bypass a known
#   PyInstaller-modulegraph bug that crashes on pandas's bytecode with Python
#   3.10's dis module ("IndexError: tuple index out of range").

# -*- mode: python ; coding: utf-8 -*-

import os

# Workaround: PyInstaller's modulegraph uses dis to scan module bytecode.
# Some pandas/numpy submodules contain bytecode where a LOAD_CONST const_index
# falls outside dis's view of the constants tuple, producing
# "IndexError: tuple index out of range" mid-build. Patch dis to swallow
# the error so the scan can continue. See:
#   https://github.com/pyinstaller/pyinstaller/issues/8419
import dis as _dis
_orig_get_const_info = _dis._get_const_info
def _safe_get_const_info(const_index, const_list):
    try:
        return _orig_get_const_info(const_index, const_list)
    except (IndexError, TypeError):
        return ("?", "?")
_dis._get_const_info = _safe_get_const_info

from PyInstaller.utils.hooks import collect_all

block_cipher = None


# ── Templates + first-run MRS seed ────────────────────────────────────────────
DATAS = [
    ("data/Master_Rate_Sheet.xlsx", "data"),
    ("data/Working_Sheet.xlsx",     "data"),
    ("data/Invoice_Bill.xlsx",      "data"),
    ("CLAUDE.md",                   "."),
]

BINARIES = []
HIDDEN   = [
    "num2words.lang_EN_IN",
    "num2words.lang_EN",
]

# ── Collect-all the heavyweight third-parties ────────────────────────────────
# This bundles every submodule + data + binary they ship, sidestepping
# PyInstaller's per-module bytecode scanner (which trips on pandas).
for pkg in ("pandas", "numpy", "openpyxl", "num2words"):
    p_datas, p_bins, p_hidden = collect_all(pkg)
    DATAS    += p_datas
    BINARIES += p_bins
    HIDDEN   += p_hidden


a = Analysis(
    ["main.py"],
    pathex=[os.path.abspath(".")],
    binaries=BINARIES,
    datas=DATAS,
    hiddenimports=HIDDEN,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter",
        "matplotlib",
        "scipy",
        "IPython",
        "jupyter",
        "notebook",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="AmritaLabBilling",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX often flags AV; safer off for a billing exe
    console=False,             # GUI app — no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                 # add path to .ico here when available
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="AmritaLabBilling",
)
