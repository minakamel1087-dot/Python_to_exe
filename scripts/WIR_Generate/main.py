"""
WIR Tools - entry point.

    python main.py            open the window
    python main.py --check    run every check once and print the result

The second form is there so the logic can be exercised without a window,
which is also how it gets tested.
"""

from __future__ import annotations

import sys


def _headless() -> int:
    from core import reference, tasks
    from core.workbook import LiveWorkbook

    book = LiveWorkbook.attach()
    with book:
        results = tasks.run_all(book, reference.load(book))

    for result in results:
        print(f"=== {result.title} ===")
        if not result.ok:
            print(f"  {result.headline}")
        for line in result.details:
            print(f"  {line}")
        print()

    return 1 if any(r.errors for r in results) else 0


def main() -> int:
    if "--check" in sys.argv:
        return _headless()

    from ui.window import launch
    return launch()


if __name__ == "__main__":
    sys.exit(main())
