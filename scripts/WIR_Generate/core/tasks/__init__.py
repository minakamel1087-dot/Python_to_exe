"""
The tasks the buttons run.

Each one takes the live workbook, does its job against the open sheet, and
returns a RunResult the UI can render. None of them import anything from
the UI - the logic has to stay runnable without a window.
"""

from __future__ import annotations

from ..findings import RunResult
from ..reference import ReferenceData, load
from ..workbook import LiveWorkbook
from . import check_areas, fix_links, preflight

__all__ = ["fix_links", "check_areas", "preflight", "run_all", "TASKS"]


def run_all(book: LiveWorkbook, ref: ReferenceData) -> list[RunResult]:
    """Fix the links, tidy the areas, then check everything - in that
    order, because each step feeds the next: repaired links stop the
    attachment check reporting phantom failures, and the area pass puts
    its suggestions in K before pre-flight reads the rows."""
    results = [fix_links.run(book)]
    if not results[0].ok:
        return results

    results.append(check_areas.run(book, ref))

    # The area pass may have written to the sheet; reload the reference
    # tables so pre-flight sees the same workbook the user now has.
    results.append(preflight.run(book, load(book)))
    return results


# Which task each button runs, and what it says on the button.
TASKS = {
    "fix_links": fix_links,
    "check_areas": check_areas,
    "preflight": preflight,
}
