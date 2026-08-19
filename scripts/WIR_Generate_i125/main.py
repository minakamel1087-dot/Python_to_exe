"""
WIR Generate Tools - entry point.

    python main.py               open the window
    python main.py --check       run every check once and print the result
    python main.py --selftest    prove the build is complete

The second form is there so the logic can be exercised without a window,
which is also how it gets tested. The third needs no workbook at all - it
is what the build script runs against the finished .exe.
"""

from __future__ import annotations

import os
import sys


def _add_project_to_path() -> None:
    """An embeddable Python builds sys.path from its ._pth file and does
    not add the script's own directory, so `import core` fails. Normal
    installs already have this entry; adding it twice is harmless."""
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)


def _register_pywin32_dlls() -> None:
    """pywin32 keeps pythoncom and pywintypes in its own folder and relies
    on a post-install step to make them findable. That step cannot run in
    an embeddable Python, and does not run at all inside a packaged .exe,
    so point the loader at every place they might be.

    Harmless where it does not apply - a folder that is absent is skipped,
    and one already on the path costs nothing to add again.
    """
    roots = [sys.prefix]
    bundle = getattr(sys, "_MEIPASS", None)     # where a one-file .exe unpacks
    if bundle:
        roots.insert(0, bundle)

    candidates = []
    for root in roots:
        candidates.append(root)
        candidates.append(os.path.join(root, "pywin32_system32"))
        candidates.append(os.path.join(root, "Lib", "site-packages", "pywin32_system32"))

    for folder in candidates:
        if not os.path.isdir(folder):
            continue
        try:
            os.add_dll_directory(folder)
        except OSError:
            pass
        os.environ["PATH"] = folder + os.pathsep + os.environ.get("PATH", "")


_add_project_to_path()
_register_pywin32_dlls()


def _probe() -> int:
    """Read-only. Attaches, loads the reference tables and counts the
    rows, and writes nothing at all - so the risky parts (finding the
    workbook, resolving tables by name, reading with Value2) can be
    proved before anything touches the sheet."""
    from core import config, reference
    from core.workbook import LiveWorkbook, text_of

    book = LiveWorkbook.attach()
    print(f"workbook     : {book.wb.Name}")
    print(f"active sheet : {book.excel.ActiveSheet.Name}")

    sheet = book.log_sheet
    last = book.last_row(sheet, config.Main.WIR_NO)
    rows = book.block(sheet, config.FIRST_DATA_ROW, last, config.Main.LAST_COL)

    filled = [r for r in rows if text_of(r[config.Main.WIR_NO - 1])]
    done = [
        r for r in filled
        if text_of(r[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower()
    ]
    print(f"log sheet    : {sheet.Name}, last row {last}")
    print(f"rows         : {len(filled)} with a WIR number, {len(done)} marked Done")

    register = book.sheet(config.REGISTER_SHEET)
    if register is None:
        print(f"register     : sheet '{config.REGISTER_SHEET}' NOT FOUND")
    else:
        print(f"register     : {book.last_row(register, config.Register.WIR_NO)} rows")

    ref = reference.load(book)
    print()
    print(f"Prev_Activities  : {len(ref.previous)} activities")
    print(f"Activity_Splits  : {len(ref.split_subs)} parents, {len(ref.split_parents)} subs")
    print(f"Floor_Apartments : {len(ref.apartments)} floor codes")
    print(f"WIR_Status_Codes : {len(ref.status)} codes -> "
          f"{sorted(c for c in ref.status if ref.is_approved(c))} count as approved")
    print(f"Area_Names       : {len(ref.area_names)} areas, "
          f"{len(ref.area_canonical)} spellings, "
          f"{sum(1 for v in ref.area_floors.values() if v)} floor-restricted")

    missing = [
        name for name, ok in (
            (config.TABLE_PREV_ACTIVITIES, ref.previous),
            (config.TABLE_FLOOR_APARTMENTS, ref.apartments),
            (config.NAME_AREA_NAMES, ref.area_names),
        ) if not ok
    ]
    print()
    print("NOT FOUND: " + ", ".join(missing) if missing else "All reference tables resolved.")
    return 1 if missing else 0


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


def _pause_if_asked() -> None:
    """Launched from a shortcut, a console window shuts the instant the
    program ends and takes the output with it."""
    if "--pause" in sys.argv:
        try:
            input("\nPress Enter to close.")
        except EOFError:
            pass


def _expired() -> bool:
    """The licence check, reported the way each caller can see it.

    --selftest deliberately runs before this: a machine that cannot start
    the program still needs to be able to prove why.
    """
    from core import licence

    verdict = licence.check()
    if verdict.ok:
        return False

    # A licence signed with a blank message stops the program without a
    # word - nothing printed, no dialog. Deliberate, and chosen at the
    # moment the licence was signed.
    if verdict.silent:
        return True

    message = "\n\n".join(part for part in (verdict.headline, verdict.detail) if part)
    print(message, file=sys.stderr)

    if len(sys.argv) == 1:                       # the windowed run
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(
                None, message, "WIR Generate Tools", 0x30      # warning icon
            )
        except Exception:                        # noqa: BLE001
            pass
    return True


def main() -> int:
    if "--selftest" in sys.argv:
        from core import selftest
        code = selftest.run()
        _pause_if_asked()
        return code

    if _expired():
        _pause_if_asked()
        return 2

    if "--probe" in sys.argv:
        return _probe()
    if "--check" in sys.argv:
        return _headless()

    from ui.window import launch
    return launch()


def _report_startup_failure(exc: BaseException) -> None:
    """Something went wrong before the window existed.

    The release build has no console, so a traceback goes nowhere a user
    can read it. ctypes is used rather than Qt because the most likely
    cause of getting here is that a piece of the bundle is missing, and
    Qt could easily be the missing piece.
    """
    # The packaging advice belongs only on a packaging failure. Attached to
    # something ordinary like "Excel is not running" it just buries the one
    # line that tells the user what to do.
    packaging = isinstance(exc, ImportError) or "No module named" in str(exc)
    advice = (
        "\n\nThe program was packaged incompletely. Run it once from a "
        "command prompt with --selftest and send the selftest.log it "
        "writes to whoever built it."
        if packaging else ""
    )
    message = f"WIR Generate Tools could not start.\n\n{type(exc).__name__}: {exc}{advice}"
    print(message, file=sys.stderr)

    # Only for the windowed run. The --check and --selftest forms are run
    # from a command prompt, where a modal box just blocks the output the
    # person is there to read.
    if len(sys.argv) > 1:
        return

    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, message, "WIR Generate Tools", 0x10)
    except Exception:                                  # noqa: BLE001
        pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException as exc:                       # noqa: BLE001
        _report_startup_failure(exc)
        sys.exit(1)
