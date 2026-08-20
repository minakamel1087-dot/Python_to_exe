"""
Builds "Licence Signer.exe" with the private key inside it.

    python tools\\build_signer.py

The key is read from disk, substituted into a copy of
licence_signer_gui.py in a temporary folder, built, and the copy deleted.
Nothing carrying the key is ever written into the project, so there is
nothing to accidentally commit or deploy.

The finished .exe is a different matter: **it holds the signing key.**
Anyone who has it can issue licences for this program, and unpacking a
PyInstaller archive is not difficult. Keep it as you would keep the key
itself - it is the key, in a more convenient shape.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)

SOURCE = os.path.join(HERE, "licence_signer_gui.py")
PLACEHOLDER = "__PRIVATE_KEY_PEM__"

# Where the built .exe goes: alongside the project, not inside it.
DEFAULT_OUT = os.path.normpath(os.path.join(PROJECT, "..", "..", "Licence Signer"))

# The key is looked for beside the signer first - that is where it is kept -
# and then in the folder sign_licence.py originally created it.
KEY_CANDIDATES = [
    os.path.join(DEFAULT_OUT, "licence_private.pem"),
    os.path.join(
        os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or ".",
        "WIR Generate Tools", "licence_private.pem",
    ),
]


def find_key() -> str:
    for candidate in KEY_CANDIDATES:
        if os.path.isfile(candidate):
            return candidate
    return KEY_CANDIDATES[0]           # the one to name when nothing is found


def build(key_path: str, out_dir: str, python: str) -> int:
    if not os.path.isfile(key_path):
        print(f"No private key at:\n  {key_path}")
        print("\nGenerate one first:  python tools\\sign_licence.py --new-key")
        return 1

    with open(key_path, "r", encoding="utf-8") as handle:
        key = handle.read().strip()
    if "PRIVATE KEY" not in key:
        print(f"That file does not look like a PEM private key:\n  {key_path}")
        return 1

    with open(SOURCE, "r", encoding="utf-8") as handle:
        source = handle.read()
    if PLACEHOLDER not in source:
        print(f"{SOURCE} has no {PLACEHOLDER} to substitute.")
        return 1

    work = tempfile.mkdtemp(prefix="signer_")
    keyed = os.path.join(work, "licence_signer_gui.py")
    try:
        with open(keyed, "w", encoding="utf-8") as handle:
            handle.write(source.replace(PLACEHOLDER, key))

        print(f"key      : {key_path}")
        print(f"building : {os.path.abspath(out_dir)}")
        print()

        result = subprocess.run(
            [
                python, "-m", "PyInstaller",
                "--noconfirm", "--onefile", "--windowed",
                "--name", "Licence Signer",
                "--distpath", os.path.abspath(out_dir),
                "--workpath", os.path.join(work, "build"),
                "--specpath", os.path.join(work, "spec"),
                "--hidden-import", "cryptography",
                keyed,
            ],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(result.stdout[-3000:])
            print(result.stderr[-3000:])
            return result.returncode

        exe = os.path.join(os.path.abspath(out_dir), "Licence Signer.exe")
        size = os.path.getsize(exe) / 1_000_000 if os.path.exists(exe) else 0
        print(f"Built: {exe}  ({size:.1f} MB)")
        print()
        print("This executable CONTAINS the signing key. Do not share it.")
        return 0
    finally:
        # The keyed copy and every build artefact go with it.
        shutil.rmtree(work, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the Licence Signer executable.")
    parser.add_argument("--key", default=find_key(), help="the private key PEM")
    parser.add_argument("--output", default=DEFAULT_OUT, help="where to put the .exe")
    parser.add_argument("--python", default=sys.executable,
                        help="the Python that has PyInstaller installed")
    options = parser.parse_args()
    return build(options.key, options.output, options.python)


if __name__ == "__main__":
    sys.exit(main())
