"""
How a long task says where it has got to.

Deliberately tiny and Qt-free: a task takes a Progress or takes nothing,
and the window is the only thing that knows a dialog exists. That keeps
the tasks runnable from a script, from --check, and from a test, none of
which have a window to draw on.
"""

from __future__ import annotations

from typing import Callable


class Progress:
    """Passed into a task that may run for a while.

    `report` is called as the work proceeds. `cancelled` is checked by the
    task between items - it is set from the GUI thread and read from the
    worker thread, which is safe for a plain bool flag and does not need
    a lock: a missed read simply means one more item is processed.
    """

    def __init__(self, callback: Callable[[int, int, str], None] | None = None):
        self._callback = callback
        self.cancelled = False

    def report(self, done: int, total: int, message: str = "") -> None:
        if self._callback:
            self._callback(done, total, message)

    def cancel(self) -> None:
        self.cancelled = True


#: A Progress that reports nowhere, so tasks need no None checks.
SILENT = Progress()
