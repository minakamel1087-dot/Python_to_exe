"""
Builds the portable "WIR Generate Tools" folder.

    python tools\\deploy.py

The result is a folder you copy next to i125-WIR Cover Generate.xlsm on
any machine. Nothing is installed, nothing is compiled, and there is no
.bat, .vbs or .exe of ours for IT to object to - only Python's own
python.exe, which is signed by the Python Software Foundation.

    WIR Generate Tools\\
        WIR Generate Tools.lnk         double-click this
        Check this PC.lnk     run once on a new machine to prove it works
        runtime\\              a copy of the embeddable Python
        app\\                  main.py, core, ui
        README.txt

A .lnk can only hold an absolute path - there is no portable form of one,
and no way to say "try here, then there". So the paths have to be agreed
in advance, and the deployment works like this:

    D:\\i125\\python Package        the runtime, same path on every machine
    <output folder>               app + shortcuts, same path on every machine

Build once with --link-runtime, then copy the small output folder to that
same path everywhere. Both halves are local, so the program starts with
no network at all - which matters, because a shortcut pointing at a
server simply fails when the server is down, before any of this program's
own "share unavailable" handling can run.

Without --link-runtime the runtime is copied into the folder instead:
one self-contained ~250 MB folder, no fixed Python path needed.
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)

# The embeddable Python. This exact path exists on every machine, which is
# what lets --link-runtime work offline: the shortcut points at a local
# folder, not the server, so nothing needs a network to start.
DEFAULT_RUNTIME = r"D:\i125\python Package"

# Copied into app\. Everything the program needs and nothing else.
APP_PARTS = ["main.py", "core", "ui"]

# Left out of runtime\. Development and documentation weight that a user
# machine never reads.
RUNTIME_SKIP = {
    "wheels",            # the pip wheels the runtime was built from
    "pythonwin",         # pywin32's IDE
    "PyWin32.chm",       # pywin32's help file
    "__pycache__",
}

README = """WIR Generate Tools
==================

Nothing here needs installing.

    1. Keep this folder where it was put. The shortcuts point at that
       exact path, so moving or renaming it stops them working - copy the
       shortcut to your desktop instead of moving the folder.
    2. Open i125-WIR Cover Generate.xlsm in Excel and select your WIR log
       sheet.
    3. Double-click "WIR Generate Tools".

On a machine that has not run this before, double-click "Check this PC"
first. It takes a few seconds, needs no workbook, and prints whether this
machine can run the program. It also writes app\\selftest.log, which is
the thing to send on if it reports a problem.

The app folder is the program itself, in readable source form. There is
nothing compiled here.

Python comes from the Python Software Foundation. Depending on how this
was built it is either in the runtime folder beside this file, or on the
server with the shortcuts pointing at it - in which case the server has
to be reachable for the program to start.

The full manual is "WIR Generate Tools - User Manual.docx".
"""


def _copy_tree(source: str, target: str, skip: set[str] | None = None) -> int:
    """Copies a folder and reports how many files it wrote. shutil's own
    ignore callback is avoided so the skip list can be matched case-
    insensitively, which is what a Windows user would expect."""
    skip = {name.lower() for name in (skip or set())}
    written = 0

    for folder, subfolders, files in os.walk(source):
        subfolders[:] = [name for name in subfolders if name.lower() not in skip]

        relative = os.path.relpath(folder, source)
        destination = target if relative == "." else os.path.join(target, relative)
        os.makedirs(destination, exist_ok=True)

        for name in files:
            if name.lower() in skip or name.endswith(".pyc"):
                continue
            shutil.copy2(os.path.join(folder, name), os.path.join(destination, name))
            written += 1

    return written


def _make_shortcut(path: str, target: str, arguments: str, description: str,
                   root: str) -> None:
    """Writes a .lnk pointing at where the folder will finally live.

    A .lnk always stores an absolute target - there is no portable form of
    one. Windows will sometimes re-find a moved target on its own, but
    that is a heuristic and not something to hand to other people's
    machines, so the final location is stated up front instead (see
    --final-path) and the shortcut is simply correct for it.

    The working directory is deliberately left empty. Explorer then starts
    the process in pythonw.exe's own folder, which is what makes the
    script argument `..\\app\\main.py` resolve - so only the runtime path
    has to be right, not two paths.
    """
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    link = shell.CreateShortCut(path)
    link.TargetPath = target
    link.Arguments = arguments
    link.Description = description
    link.WorkingDirectory = ""
    link.Save()


def build(runtime_source: str, output: str, final: str,
          link_runtime: bool = False) -> int:
    have_runtime = os.path.isfile(os.path.join(runtime_source, "pythonw.exe"))

    if not have_runtime:
        if not link_runtime:
            # It has to be copied, so it has to be here.
            print(f"Runtime not found: {runtime_source}")
            print("Point --runtime at the embeddable Python folder.")
            return 1

        # Linking only records the path. It is a statement about the
        # machines this will run on, and those are not this one - so warn
        # rather than refuse.
        print(f"WARNING: no pythonw.exe at {runtime_source}")
        print("         Building anyway - the shortcuts will expect it there")
        print("         on every machine that runs this.")
        print()

    print(f"runtime from : {runtime_source}")
    print(f"app from     : {PROJECT}")
    print(f"building     : {output}")
    print(f"will live at : {final}")
    if os.path.normcase(output) != os.path.normcase(final):
        print("               (shortcuts are written for that path, not this one)")
    print()

    if os.path.isdir(output):
        shutil.rmtree(output)
    os.makedirs(output)

    if link_runtime:
        # Not copied: the shortcuts point at the runtime where it already
        # is. One Python to maintain, and the folder built here is a few
        # hundred KB instead of 250 MB.
        print(f"  runtime : linked, not copied - {runtime_source}")
    else:
        count = _copy_tree(runtime_source, os.path.join(output, "runtime"),
                           RUNTIME_SKIP)
        print(f"  runtime : {count} files")

    app = os.path.join(output, "app")
    os.makedirs(app)
    count = 0
    for part in APP_PARTS:
        source = os.path.join(PROJECT, part)
        if os.path.isdir(source):
            count += _copy_tree(source, os.path.join(app, part), {"__pycache__"})
        elif os.path.isfile(source):
            shutil.copy2(source, os.path.join(app, part))
            count += 1
        else:
            print(f"  MISSING : {part}")
            return 1
    print(f"  app     : {count} files")

    manual = os.path.join(PROJECT, "WIR Generate Tools - User Manual.docx")
    if os.path.isfile(manual):
        shutil.copy2(manual, os.path.join(output, os.path.basename(manual)))
        print("  manual  : copied")

    with open(os.path.join(output, "README.txt"), "w", encoding="utf-8") as handle:
        handle.write(README)

    # Written for the final location, which is usually not where we are
    # building - so these paths come from `final`, not `output`.
    if link_runtime:
        # The runtime stays where it is, so the script cannot be named
        # relative to pythonw.exe's folder any more. Both paths absolute.
        final_runtime = runtime_source
        script = f'"{os.path.join(final, "app", "main.py")}"'
    else:
        final_runtime = os.path.join(final, "runtime")
        script = r'"..\app\main.py"'

    # pythonw for the window - python.exe would flash a console behind it.
    _make_shortcut(
        os.path.join(output, "WIR Generate Tools.lnk"),
        os.path.join(final_runtime, "pythonw.exe"),
        script,
        "Open WIR Generate Tools",
        final,
    )
    # python.exe for the check, so its output is visible, and --pause so
    # the window stays up long enough to read.
    _make_shortcut(
        os.path.join(output, "Check this PC.lnk"),
        os.path.join(final_runtime, "python.exe"),
        f"{script} --selftest --pause",
        "Prove this machine can run WIR Generate Tools",
        final,
    )
    print("  shortcuts: 2")

    total = sum(
        os.path.getsize(os.path.join(folder, name))
        for folder, _, files in os.walk(output)
        for name in files
    )
    print()
    print(f"Done. {total / 1_000_000:.1f} MB")
    if os.path.normcase(output) != os.path.normcase(final):
        print(f"Copy this folder to {final} - the shortcuts expect it there.")
    if link_runtime:
        print()
        print("The runtime is NOT in this folder. The shortcuts point at")
        print(f"  {runtime_source}")
        print("which has to stay put and be reachable from every machine.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build the portable WIR Generate Tools folder.",
        epilog="The shortcuts hold an absolute path, so --final-path must be "
               "where the folder ends up. Use one location for everyone: a "
               "folder on the share, or the same local path on each PC.",
    )
    parser.add_argument("--runtime", default=DEFAULT_RUNTIME,
                        help="the embeddable Python to copy")
    parser.add_argument("--output", default=os.path.join(PROJECT, "dist", "WIR Generate Tools"),
                        help="where to build it")
    parser.add_argument("--final-path", default=None,
                        help="where the folder will live once copied "
                             "(defaults to --output, i.e. built in place)")
    parser.add_argument("--link-runtime", action="store_true",
                        help="do not copy the runtime; point the shortcuts at "
                             "it where it already is. A few hundred KB instead "
                             "of 250 MB, and one Python to maintain - but that "
                             "path must stay put and be reachable everywhere")
    options = parser.parse_args()

    output = os.path.abspath(options.output)
    final = os.path.abspath(options.final_path) if options.final_path else output
    # Not abspath'd when linking: a UNC path is already absolute and
    # os.path.abspath would mangle it against the current drive.
    runtime = options.runtime if options.link_runtime else os.path.abspath(options.runtime)
    return build(runtime, output, final, options.link_runtime)


if __name__ == "__main__":
    sys.exit(main())
