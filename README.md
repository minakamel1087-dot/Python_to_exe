# py-to-exe

Turn any Python project into a standalone Windows `.exe`. Two ways in, same
rules, same result:

| | Needs | Good for |
| --- | --- | --- |
| **PyToExe Builder** (a local app) | Python on the machine | offline work, private code, fast iteration |
| **GitHub Actions** (push and download) | nothing installed | a clean machine, or handing builds to someone else |

## PyToExe Builder — compile any repo, locally

Build it once (see below), then run `PyToExe_Builder.exe`:

1. **Browse...** to any repo or project folder.
2. It lists what it could build; the best guess is already selected. Override
   it from the dropdown, or **Pick file...** for a layout it didn't anticipate.
3. Choose one-file or folder, console or windowed, an icon if you want one.
4. **Build .exe**, and watch the log. The result lands in that project's `dist\`.

**It needs Python installed on that machine.** PyInstaller compiles a program by
importing it and tracing what it pulls in, which takes a real interpreter with a
complete standard library — a frozen exe's own runtime is stripped and cannot do
it. Embedding a Python distribution to dodge that would be large and fragile. So
the app finds Python, keeps its own virtualenv in `%LOCALAPPDATA%\PyToExe`, and
runs PyInstaller as a subprocess. No Python found, and it says so plainly instead
of failing halfway through.

It touches nothing else: no registry, no admin, no services, no shell scripts.
Uninstalling is deleting `%LOCALAPPDATA%\PyToExe` and the exe.

## GitHub Actions — compile without installing anything

Drop a `.py` (or a whole project folder) into `scripts/`, commit, push:

```bash
git add scripts/my_tool
git commit -m "Add my_tool"
git push
```

The **Build scripts** workflow runs and the `.exe` waits under **Actions → that
run → Artifacts**. Every project in `scripts/` builds independently, so a broken
one doesn't hold up the rest.

For one-off options — an icon, windowed mode, extra PyInstaller flags — use
**Actions → Build exe (manual) → Run workflow** against any path in the repo.

## What gets built

Both routes pick the entry point the same way, best first:

1. **A `.spec` file** in the project folder. A project that ships one has already
   declared its hidden imports, data files and windowed mode; no flags here could
   second-guess it. It runs from its own folder, so the relative paths inside it
   (`Analysis(["main.py"])`, `pathex=["."]`) behave as they do by hand.
2. **A conventional entry point** — `main.py`, `__main__.py`, `app.py`, `run.py`,
   `cli.py`, `gui.py`, `start.py` — at the root or one or two levels down, so
   `src/myapp/__main__.py` is found. `.venv`, `build`, `dist`, `tests` and friends
   are skipped.
3. **Any other `.py` at the root**, so a flat pile of scripts still offers a list.

The exe is named after the folder when the entry point is `main.py`
(`myproject/main.py` → `myproject.exe`), otherwise after the file. In the local
app you can always override the choice; the workflow takes the first match.

## Dependencies

In this order, no configuration required:

1. **A `requirements.txt` beside the entry point** — used as-is. The reliable
   option; reach for it the moment a build guesses wrong.
2. **Otherwise the nearest one above it**, walking up to the project root — so
   one tool never picks up its neighbour's.
3. **Nothing found** — `pipreqs` reads the imports and generates one, printed in
   the log. Packages install one at a time, because pipreqs infers package names
   from import names and gets some wrong (`import win32com` really means
   `pywin32`, `import cv2` means `opencv-python`, a local module looks like a
   package). One bad guess doesn't stop the rest, and the log names what it
   skipped.

PyInstaller itself is always installed, pinned, whichever branch runs.

## Per-project flags: build.args

Drop a `build.args` next to the entry point for anything the defaults get wrong.
One flag per line, `#` starts a comment:

```
--windowed
--hidden-import win32timezone
--add-data "templates;templates"
```

## Three things that bite with PyInstaller

Nearly every "it built but won't run" is one of these:

- **Data files aren't detected.** Templates, images, `.json` config: add
  `--add-data "assets;assets"` (on Windows the separator is `;`, not `:`).
- **Dynamic imports are invisible.** Anything imported by name at runtime —
  plugin loaders, `importlib`, database drivers, all of pywin32 — is left out, so
  the exe builds cleanly and dies on startup. Fix with `--hidden-import`.
- **Paths change when frozen.** `__file__` points into a temp extraction folder,
  so writing next to it silently loses data. Use `sys.executable`'s folder for
  anything the user keeps, and check `getattr(sys, "frozen", False)` when the two
  differ.

Also expect a SmartScreen warning on any unsigned exe, 30-60 MB for anything
pulling in Qt or pandas, and the occasional antivirus false positive — a onefile
exe unpacks itself to `%TEMP%` and runs from there, which is what packed malware
does too. Folder mode avoids that pattern and starts faster; the cost is shipping
a folder instead of a single file.

## Building the builder

The workflow compiles it like any other project — push, then download the
**PyToExe_Builder** artifact. Or locally, if you already have Python:

```bash
powershell -ExecutionPolicy Bypass -File build-exe.ps1 -Script scripts\PyToExe_Builder\main.py
```

## Layout

| Path | What it is |
| --- | --- |
| `scripts/` | The drop folder — projects the workflow builds. |
| `scripts/PyToExe_Builder/` | The local app. Compiles any repo on your machine. |
| `build-exe.ps1` | Every build rule, in one place. Both workflows call it, so CI and local builds can't drift. |
| `.github/workflows/build-scripts.yml` | Builds everything in `scripts/` on push. |
| `.github/workflows/build-exe.yml` | Manual build of any path, full options. |
