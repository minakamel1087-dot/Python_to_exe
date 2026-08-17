# -*- mode: python ; coding: utf-8 -*-
"""
Build recipe for WIR Tools.

    pyinstaller --noconfirm "WIR Tools.spec"

The result must run on a machine with no Python and no pywin32, so this
file errs towards putting too much in rather than too little. Size is not
the constraint - a missing module is only discovered by the person using
it, halfway through a real batch of WIRs.

PyInstaller finds imports by reading the source. That misses anything
imported by name at runtime, which is exactly how pywin32 works:

  * pythoncom imports win32timezone the first time COM hands back a date,
    so nothing in the source ever mentions it
  * win32com.client.Dispatch builds its wrappers from strings

Both are listed below by hand for that reason. `--selftest` on the built
.exe proves it worked; build.bat runs it and refuses to hand over a build
that fails.
"""

from PyInstaller.utils.hooks import collect_submodules

# --- the parts PyInstaller cannot see -------------------------------------

hidden = [
    # pywin32. win32timezone is the one that has actually bitten us.
    "win32timezone",
    "pywintypes",
    "pythoncom",
    "win32api",
    "win32con",
    "win32gui",
    "win32process",
    "win32clipboard",
    "win32ui",
    # Excel, Word and Outlook are all reached through these.
    "win32com",
    "win32com.client",
    "win32com.client.dynamic",
    "win32com.client.gencache",
    "win32com.client.build",
    "win32com.client.makepy",
    "win32com.server.util",
]

# Dispatch resolves wrapper modules by name, so take the whole package.
hidden += collect_submodules("win32com")
hidden += collect_submodules("win32comext")
hidden += collect_submodules("pypdf")

# pypdf only imports cryptography when it meets a protected PDF - which a
# consultant's attachment often is. Include it when the build machine has
# it; a build without it still works, it just cannot merge those files.
try:
    hidden += collect_submodules("cryptography")
except Exception:                                    # noqa: BLE001
    print("NOTE: cryptography not installed - encrypted PDFs will not merge.")

hidden = sorted(set(hidden))

# --- things we know we never use ------------------------------------------
# Kept deliberately short. Every entry here is a chance to break the build
# in a way that only shows up on someone else's machine, so nothing is
# excluded unless the program plainly cannot reach it.

excludes = [
    "tkinter",
    "unittest",
    "pydoc_data",
    "test",
    "lib2to3",
]


a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=[],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name="WIR Tools",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,              # UPX corrupts some Qt DLLs; not worth the megabytes
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # --selftest still returns an exit code and writes selftest.log
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
