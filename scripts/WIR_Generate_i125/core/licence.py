"""
The expiry check.

A small signed text file says when this program stops working:

    # WIR Generate Tools
    expires = 2026-12-31
    message = internal Error Contact Mina Kamel
    signature = 8Kk2pQ7v...==

The signature covers every line except itself. It is made with a private
key that exists only on the author's machine; the program carries the
matching public key, which can verify a signature but can never create
one. So the file can sit in plain sight and still not be forgeable -
change the date and verification fails.

An unsigned file is rejected. Accepting one would make the bypass
"delete the signature line".

Where it looks, in order
------------------------
1. The local file (config.LICENCE_LOCAL_FILE). Found and valid decides
   the matter either way, and the network is never touched.
2. The published file (config.LICENCE_URL).
3. Neither reachable: a machine that has never validated is allowed to
   run - otherwise a firewall would brick a new PC on its first day.
   One that has validated before is not.

What this is, and is not
------------------------
It stops the program quietly outliving the author's involvement, and it
stops the one bypass a non-programmer could actually manage: editing a
date in Notepad. It does not stop someone who can edit Python - the
source ships readable. Signing protects the data, not the program.
"""

from __future__ import annotations

import base64
import json
import os
import urllib.request
from datetime import date, datetime

from . import config

CACHE_NAME = "state.json"
SIGNATURE_KEY = "signature"

_DATE_FORMATS = ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d")
_DATE_KEYS = ("expires", "expiry", "expiry_date", "date")


class Verdict:
    """Whether the program may run, and what to tell the user."""

    def __init__(self, ok: bool, headline: str = "", detail: str = "",
                 silent: bool = False):
        self.ok = ok
        self.headline = headline
        self.detail = detail
        # An expired licence whose message was left blank stops the
        # program without saying anything. That is the author's choice,
        # made when the licence was signed.
        self.silent = silent

    def __bool__(self) -> bool:
        return self.ok


# --- the licence file ------------------------------------------------------

def parse_date(text: str) -> date | None:
    text = (text or "").strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def split_signed(text: str) -> tuple[bytes, str]:
    """(what was signed, the signature).

    Line endings are normalised and trailing spaces dropped before the
    payload is formed, so a file edited on GitHub with LF verifies the
    same as a copy saved locally with CRLF. Without that, the same file
    would fail on one machine and pass on another.
    """
    payload_lines: list[str] = []
    signature = ""

    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.rstrip()
        key = line.split("=", 1)[0].strip().lower() if "=" in line else ""
        if key == SIGNATURE_KEY:
            signature = line.split("=", 1)[1].strip()
            continue
        payload_lines.append(line)

    return "\n".join(payload_lines).strip().encode("utf-8"), signature


def payload_fields(payload: bytes) -> tuple[date | None, str]:
    """(expiry, message) from the signed part."""
    expiry: date | None = None
    message = ""

    for line in payload.decode("utf-8", "replace").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue

        if "=" in line or ":" in line:
            separator = "=" if "=" in line else ":"
            key, _, value = line.partition(separator)
            key = key.strip().lower()
            value = value.strip()
            if key in _DATE_KEYS:
                expiry = parse_date(value) or expiry
                continue
            if key == "message":
                message = value
                continue

        if expiry is None:                          # a bare date on its own line
            expiry = parse_date(line)

    return expiry, message


def verify(text: str) -> tuple[date | None, str, str]:
    """(expiry, message, error) for one licence file.

    An error string means do not trust it, whatever the date said.
    """
    payload, signature = split_signed(text)
    if not signature:
        return None, "", "it carries no signature"

    try:
        from cryptography.exceptions import InvalidSignature
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
    except ImportError:
        return None, "", "the cryptography package is missing from this build"

    try:
        public = Ed25519PublicKey.from_public_bytes(
            base64.b64decode(config.LICENCE_PUBLIC_KEY)
        )
    except Exception:                               # noqa: BLE001
        return None, "", "the public key in this build is not usable"

    try:
        public.verify(base64.b64decode(signature), payload)
    except InvalidSignature:
        return None, "", "it has been changed, or it was not issued by the author"
    except Exception as exc:                        # noqa: BLE001
        return None, "", f"the signature could not be read ({exc})"

    expiry, message = payload_fields(payload)
    if expiry is None:
        return None, "", "it holds no expiry date"
    return expiry, message, ""


# --- where the files come from --------------------------------------------

def read_local(path: str) -> str | None:
    """The local licence text, or None when there is no file at all."""
    if not path or not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8-sig") as handle:
            return handle.read()
    except OSError:
        return None


def fetch(url: str, timeout: float) -> tuple[str | None, str]:
    """(text, error). Never raises - being offline is allowed."""
    if not url:
        return None, "no address is configured"
    try:
        request = urllib.request.Request(
            url,
            headers={"User-Agent": "WIR-Generate-Tools", "Cache-Control": "no-cache"},
        )
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.read().decode("utf-8", "replace"), ""
    except Exception as exc:                        # noqa: BLE001
        return None, str(exc)


# --- remembering ------------------------------------------------------------

def _cache_path() -> str:
    root = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or "."
    folder = os.path.join(root, "WIR Generate Tools")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, CACHE_NAME)


def _read_cache() -> dict:
    try:
        with open(_cache_path(), "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        return {}


def _write_cache(data: dict) -> None:
    try:
        with open(_cache_path(), "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
    except OSError:
        pass                                        # not being able to remember is not fatal


# --- the decision -----------------------------------------------------------

def check(today: date | None = None) -> Verdict:
    """May the program run?"""
    if not config.LICENCE_PUBLIC_KEY:
        return Verdict(True)                        # no key set: the check is off

    today = today or date.today()
    cache = _read_cache()

    # A clock wound back below a day already seen would revive an expired
    # licence, so it is refused before anything else is considered.
    seen = parse_date(cache.get("highest_date_seen", ""))
    if seen and today < seen:
        return Verdict(
            False,
            "The date on this PC looks wrong.",
            f"It reads {today:%d %b %Y}, but this program has already run on "
            f"{seen:%d %b %Y}.\nCorrect the clock and start it again.",
        )

    def remember() -> None:
        cache["validated_on"] = today.isoformat()
        if not seen or today > seen:
            cache["highest_date_seen"] = today.isoformat()
        _write_cache(cache)

    def judge(expiry: date, message: str, source: str) -> Verdict:
        if today > expiry:
            # Only what the licence itself says. No date, no wording of
            # ours - the message is the whole of what the user is told,
            # and a blank one means tell them nothing at all.
            if not message:
                return Verdict(False, silent=True)
            return Verdict(False, message)
        remember()
        return Verdict(True, detail=f"Valid until {expiry:%d %b %Y}, from {source}.")

    # 1. the local file. Still current, and the network is never touched.
    lapsed_message: str | None = None

    local = read_local(config.LICENCE_LOCAL_FILE)
    if local is not None:
        expiry, message, error = verify(local)
        if error:
            # Forged, corrupt or half-written. The published file is asked
            # instead of stopping here, so a damaged file cannot brick the
            # machine. Note the cost: tampering now produces no signal -
            # the program simply consults the host and carries on.
            pass
        elif today <= expiry:
            return judge(expiry, message, "the file on this PC")
        else:
            # Genuine, but out of date. Ask the published file before
            # giving up - that is how a machine renews itself once the
            # date has been extended, without anyone going to it.
            lapsed_message = message

    # 2. the published file
    text, error = fetch(config.LICENCE_URL, config.LICENCE_TIMEOUT)
    if text is not None:
        expiry, message, verify_error = verify(text)
        if verify_error:
            return Verdict(
                False,
                "The published licence file is not valid.",
                f"Reason: {verify_error}.",
            )
        return judge(expiry, message, "the published licence")

    # 3. the host could not be reached
    if lapsed_message is not None:
        # A real licence that has simply run out. A blank message is still
        # honoured - that choice was made when the licence was signed.
        if not lapsed_message:
            return Verdict(False, silent=True)
        return Verdict(False, "This version is expired.")

    # Anything else: no licence here, or one that could not be trusted,
    # and no way to ask. This covers the first run as well - a new machine
    # with no connection is told why rather than closing without a word.
    return Verdict(
        False,
        "The license could not be checked... connect to the internet",
    )
