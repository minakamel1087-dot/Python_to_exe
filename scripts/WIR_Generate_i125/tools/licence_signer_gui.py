"""
Licence Signer - the window version.

Type a date, press Generate, and it writes i125_WIR_Validation.txt beside
itself and shows the block to paste into the published file.

The private key is built into the executable by tools/build_signer.py,
which substitutes it for the placeholder below. This source file never
holds a key, so it is safe to keep in the project; the built .exe is not,
and must not be shared - anyone holding it can issue licences.

tkinter rather than PySide6 on purpose: it is in the standard library, so
this builds with any Python, including the 3.14 install that has no
PySide6 wheels.
"""

from __future__ import annotations

import base64
import os
import sys
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import messagebox, scrolledtext

# Replaced at build time. Left empty here so the source carries no secret.
PRIVATE_KEY_PEM = """__PRIVATE_KEY_PEM__"""

OUTPUT_NAME = "i125_WIR_Validation.txt"
DEFAULT_MESSAGE = "internal Error Contact Mina Kamel"

# Where WIR Generate Tools looks on each machine. Must stay in step with
# core/config.py LICENCE_LOCAL_FILE - this program is built standalone and
# cannot import it. Built from the environment, not spelled out, so it is
# right for whoever is logged in.
INSTALL_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or ".",
    "WGT", OUTPUT_NAME,
)

NAVY = "#1F4E78"
PAGE = "#F5F6F8"
MUTED = "#7A808C"
OK_GREEN = "#0F6E56"
FAIL_RED = "#C0392B"


def output_folder() -> str:
    """Beside the executable, not beside the script it was built from."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def sign(expires: date, message: str) -> str:
    """The signed licence block."""
    from cryptography.hazmat.primitives import serialization

    key = PRIVATE_KEY_PEM.strip().encode("utf-8")
    if not key or "PRIVATE KEY" not in PRIVATE_KEY_PEM:
        raise RuntimeError(
            "This build carries no signing key.\n"
            "It was not produced by tools/build_signer.py."
        )

    private = serialization.load_pem_private_key(key, password=None)

    lines = ["# WIR Generate Tools", f"expires = {expires:%Y-%m-%d}"]
    if message:
        lines.append(f"message = {message}")

    # Signed exactly as core.licence.split_signed rebuilds it - normalised
    # line endings, no trailing spaces - so a file edited on GitHub with LF
    # verifies the same as a copy saved on Windows with CRLF.
    payload = "\n".join(lines).strip().encode("utf-8")
    signature = base64.b64encode(private.sign(payload)).decode("ascii")
    return "\n".join(lines) + f"\nsignature = {signature}\n"


class Signer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Licence Signer")
        self.configure(bg=PAGE)
        self.resizable(False, False)

        tk.Label(self, text="Licence Signer", bg=PAGE, fg=NAVY,
                 font=("Segoe UI", 15, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 0))
        tk.Label(self, text="WIR Generate Tools", bg=PAGE, fg=MUTED,
                 font=("Segoe UI", 9)).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 12))

        tk.Label(self, text="Expires", bg=PAGE, font=("Segoe UI", 10)).grid(
            row=2, column=0, sticky="e", padx=(16, 8), pady=4)
        self.expires = tk.Entry(self, width=28, font=("Consolas", 11))
        self.expires.insert(0, f"{date.today() + timedelta(days=365):%Y-%m-%d}")
        self.expires.grid(row=2, column=1, sticky="w", padx=(0, 16), pady=4)

        tk.Label(self, text="Message", bg=PAGE, font=("Segoe UI", 10)).grid(
            row=3, column=0, sticky="e", padx=(16, 8), pady=4)
        self.message = tk.Entry(self, width=44, font=("Segoe UI", 10))
        self.message.insert(0, DEFAULT_MESSAGE)
        self.message.grid(row=3, column=1, sticky="w", padx=(0, 16), pady=4)

        tk.Label(self,
                 text="Leave the message empty and the program closes silently\n"
                      "when it has expired, with nothing shown to the user.",
                 bg=PAGE, fg=MUTED, font=("Segoe UI", 8), justify="left").grid(
            row=4, column=1, sticky="w", padx=(0, 16), pady=(0, 8))

        buttons = tk.Frame(self, bg=PAGE)
        buttons.grid(row=5, column=0, columnspan=2, pady=(4, 10))

        tk.Button(buttons, text="Generate", command=self.generate,
                  bg=NAVY, fg="white", activebackground="#2A6395",
                  activeforeground="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), width=18, height=1,
                  cursor="hand2").pack(side="left", padx=6)

        # The same licence, written where the program actually reads it -
        # otherwise every issue ends with copying a file by hand, which is
        # the step that gets forgotten.
        tk.Button(buttons, text="Install on this PC", command=self.install,
                  bg=OK_GREEN, fg="white", activebackground="#12866A",
                  activeforeground="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), width=18, height=1,
                  cursor="hand2").pack(side="left", padx=6)

        self.output = scrolledtext.ScrolledText(
            self, width=76, height=8, font=("Consolas", 9), wrap="none")
        self.output.grid(row=6, column=0, columnspan=2, padx=16)
        self.output.insert("1.0", "The signed licence appears here.")
        self.output.configure(state="disabled")

        self.status = tk.Label(self, text=f"Generate writes beside this program. "
                                         f"Install writes to {INSTALL_PATH}",
                               bg=PAGE, fg=MUTED, font=("Segoe UI", 9), anchor="w")
        self.status.grid(row=7, column=0, columnspan=2, sticky="we",
                         padx=16, pady=(8, 14))

        self.expires.focus_set()
        self.bind("<Return>", lambda _event: self.generate())

    def _say(self, text: str, colour: str = MUTED) -> None:
        self.status.configure(text=text, fg=colour)

    def _show(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.insert("1.0", text)
        self.output.configure(state="disabled")

    def _block(self) -> tuple[str, date] | None:
        """The signed licence for what is on screen, or None with the
        reason already shown."""
        raw = self.expires.get().strip()
        try:
            expires = datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            self._say(f"{raw!r} is not a date. Use YYYY-MM-DD.", FAIL_RED)
            return None

        if expires < date.today():
            if not messagebox.askyesno(
                "Licence Signer",
                f"{expires:%d %b %Y} is in the past.\n\n"
                "Anyone using this licence will be stopped immediately.\n\n"
                "Generate it anyway?",
            ):
                return None

        try:
            return sign(expires, self.message.get().strip()), expires
        except Exception as exc:                     # noqa: BLE001
            self._say("Could not sign.", FAIL_RED)
            messagebox.showerror("Licence Signer", str(exc))
            return None

    def _write(self, block: str, target: str) -> bool:
        try:
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with open(target, "w", encoding="utf-8") as handle:
                handle.write(block)
            return True
        except OSError as exc:
            self._say("Could not write the file.", FAIL_RED)
            messagebox.showerror("Licence Signer", f"{target}\n\n{exc}")
            return False

    def generate(self) -> None:
        made = self._block()
        if made is None:
            return
        block, expires = made

        target = os.path.join(output_folder(), OUTPUT_NAME)
        if not self._write(block, target):
            return

        self._show(block)
        days = (expires - date.today()).days
        self._say(f"Written to {target}  -  {days} day(s) from today.", OK_GREEN)

    def install(self) -> None:
        """Write it where WIR Generate Tools reads it on this machine."""
        made = self._block()
        if made is None:
            return
        block, expires = made

        if not self._write(block, INSTALL_PATH):
            return

        self._show(block)
        days = (expires - date.today()).days
        self._say(f"Installed on this PC  -  {days} day(s) from today. "
                  f"({INSTALL_PATH})", OK_GREEN)


if __name__ == "__main__":
    Signer().mainloop()
