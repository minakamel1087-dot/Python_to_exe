"""
Issues signed licence files.

Once, ever:

    python tools\\sign_licence.py --new-key

That writes the private key and prints the public key line to paste into
core/config.py. **Back the private key up.** Lose it and every machine
needs a rebuilt program, because the public key inside it changes.

Then, whenever the date needs extending:

    python tools\\sign_licence.py --expires 2027-06-30
    python tools\\sign_licence.py --expires 2027-06-30 --message "Call Mina"

It prints the signed file. Paste that into the GitHub file, and/or save
it as the local licence on machines that stay offline.

The private key never leaves this machine and is never distributed. Only
the public key goes into the program, and a public key can verify a
signature but never create one - which is what lets the licence file sit
in plain sight without being forgeable.
"""

from __future__ import annotations

import argparse
import base64
import os
import sys
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
sys.path.insert(0, PROJECT)

# Kept out of every project folder on purpose, so it cannot be committed,
# copied into a deployment, or shipped by accident.
# Beside the signer first - that is where the key is kept - then the
# folder --new-key originally created it in.
_CANDIDATES = [
    os.path.normpath(os.path.join(PROJECT, "..", "..", "Licence Signer",
                                  "licence_private.pem")),
    os.path.join(
        os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or ".",
        "WIR Generate Tools", "licence_private.pem",
    ),
]
DEFAULT_KEY = next((p for p in _CANDIDATES if os.path.isfile(p)), _CANDIDATES[0])


def _crypto():
    try:
        from cryptography.hazmat.primitives import serialization
        from cryptography.hazmat.primitives.asymmetric.ed25519 import (
            Ed25519PrivateKey,
        )
    except ImportError:
        print("The 'cryptography' package is required. It is in requirements.txt.")
        raise SystemExit(1)
    return serialization, Ed25519PrivateKey


def new_key(path: str) -> int:
    serialization, Ed25519PrivateKey = _crypto()

    if os.path.exists(path):
        print(f"A key already exists at:\n  {path}\n")
        print("Delete it first if you really mean to replace it - every machine")
        print("would then need a rebuilt program.")
        return 1

    private = Ed25519PrivateKey.generate()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(private.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        ))

    raw = private.public_key().public_bytes(
        encoding=serialization.Encoding.Raw,
        format=serialization.PublicFormat.Raw,
    )
    public = base64.b64encode(raw).decode("ascii")

    print(f"Private key written to:\n  {path}")
    print("\nBACK THIS UP, and keep it off any shared folder or repository.")
    print("\nPaste this line into core/config.py:\n")
    print(f'LICENCE_PUBLIC_KEY = "{public}"')
    return 0


def sign(path: str, expires: date, message: str) -> int:
    serialization, _ = _crypto()

    if not os.path.isfile(path):
        print(f"No private key at:\n  {path}\n\nRun with --new-key first.")
        return 1

    with open(path, "rb") as handle:
        private = serialization.load_pem_private_key(handle.read(), password=None)

    lines = ["# WIR Generate Tools", f"expires = {expires:%Y-%m-%d}"]
    if message:
        lines.append(f"message = {message}")

    # Signed exactly as core.licence.split_signed rebuilds it: normalised
    # line endings, no trailing spaces. Otherwise a file edited on GitHub
    # would verify differently from the same file saved on Windows.
    payload = "\n".join(lines).strip().encode("utf-8")
    signature = base64.b64encode(private.sign(payload)).decode("ascii")

    print("-" * 68)
    print("\n".join(lines))
    print(f"signature = {signature}")
    print("-" * 68)
    print(f"\nExpires {expires:%d %b %Y}"
          f" ({(expires - date.today()).days} days from today).")
    print("\nPaste the block above into the published file, and/or save it as")
    print("the local licence file.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create and sign licence files for WIR Generate Tools.",
    )
    parser.add_argument("--new-key", action="store_true",
                        help="generate the key pair (once, ever)")
    parser.add_argument("--expires", default="",
                        help="expiry date, YYYY-MM-DD")
    parser.add_argument("--message", default="internal Error Contact Mina Kamel",
                        help="what the user is told when it has expired")
    parser.add_argument("--key", default=DEFAULT_KEY,
                        help="the private key file")
    options = parser.parse_args()

    if options.new_key:
        return new_key(options.key)

    if not options.expires:
        parser.print_help()
        print("\nNothing to do: pass --expires or --new-key.")
        return 1

    try:
        expiry = datetime.strptime(options.expires, "%Y-%m-%d").date()
    except ValueError:
        print(f"Not a date: {options.expires!r}. Use YYYY-MM-DD.")
        return 1

    return sign(options.key, expiry, options.message)


if __name__ == "__main__":
    sys.exit(main())
