# PyInstaller build for the Ghaf Woods Log Validation desktop app.
#
# Two things here are not automatic. The reference CSVs are opened by path at
# runtime, so PyInstaller's import analysis cannot see them and they have to be
# listed as data — desktop.app.bundled_reference() reads them back out of
# _MEIPASS. And openpyxl resolves parts of its writer lazily, so the submodules
# are named explicitly rather than trusted to be found.
#
# Built windowed, so there is no console to read. That is why --selftest also
# writes selftest.log beside the exe and the build script checks the exit code.
#
# A directory build rather than one file: the folder is copied to site as-is and
# run from there, so nothing has to be compiled again to deploy. It also starts
# noticeably faster, since a one-file exe unpacks itself to a temp folder on
# every launch.

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hidden = collect_submodules("openpyxl")
# pywin32 resolves several modules only when COM hands back a value, so the
# analysis never sees them. win32timezone in particular is imported the first
# time a COM date is converted, which is on the very first Workbooks.Open —
# the exe built fine and then died with "No module named 'win32timezone'".
hidden += [
    "win32com",
    "win32com.client",
    "win32com.client.dynamic",
    "win32com.client.gencache",
    "pythoncom",
    "pywintypes",
    "win32timezone",
    "win32api",
]

a = Analysis(
    ["desktop/app.py"],
    pathex=["."],
    binaries=[],
    datas=[("reference/*.csv", "reference")],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Not used by the app; excluding them keeps the exe from doubling in size.
        "pandas",
        "numpy",
        "matplotlib",
        "PyQt5",
        "PySide2",
        "pytest",
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
    name="ClearanceWIRLogValidation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ClearanceWIRLogValidation",
)
