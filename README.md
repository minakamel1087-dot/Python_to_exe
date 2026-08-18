# py-to-exe

Give it a Python script, get back a Windows `.exe`. No Python, no PyInstaller,
nothing to install — GitHub's Windows runners do the work.

## Use it

Drop a `.py` into `scripts/`, commit, push:

```bash
git add scripts/my_tool.py
git commit -m "Add my_tool"
git push
```

The **Build scripts** workflow runs, and the `.exe` is waiting under
**Actions → that run → Artifacts**. Every script in `scripts/` builds
independently, so a broken one doesn't hold up the rest.

Need options — an icon, no console window, extra PyInstaller flags? Use
**Actions → Build exe (manual) → Run workflow** and point it at any path in the
repo (`.spec` files accepted too).

## What goes in scripts/

```
scripts/
  my_tool.py            -> my_tool.exe
  bigger_tool/
    main.py             -> bigger_tool.exe   (entry point must be main.py)
    helpers.py
    requirements.txt    (optional, see below)
    build.args          (optional, extra PyInstaller flags)
```

A single file, or a folder with `main.py` as its entry point when the tool has
helper modules beside it.

**If the folder contains a `.spec` file it wins over `main.py`.** A project that
ships its own spec has already declared its hidden imports, data files and
windowed mode there, and no set of flags here could second-guess it. The spec
runs from its own folder, so the relative paths inside it (`Analysis(["main.py"])`,
`pathex=["."]`) work exactly as they do when you run PyInstaller by hand.

`build.args` is how a script in the drop folder gets non-default flags without
going through the manual workflow — one flag per line, `#` starts a comment:

```
--windowed
--hidden-import win32timezone
--add-data "template.xlsx;."
```

## How dependencies are handled

You don't have to do anything, but here's the order so nothing is a surprise:

1. **A `requirements.txt` next to your script** — used as-is. This is the
   reliable option; reach for it whenever the build gets it wrong.
2. **Otherwise, the nearest one above it** — walking up to the repo root. So
   `scripts/bigger_tool/main.py` picks up `scripts/bigger_tool/requirements.txt`,
   never a neighbouring tool's.
3. **Nothing found** — `pipreqs` reads your imports and generates one. It's
   printed in the build log. Packages install one at a time, so a line pipreqs
   guessed wrong (`import cv2` really means `opencv-python`, a local module
   mistaken for a package) doesn't stop the others — the log names what it
   skipped.
4. **`requirements: none`** in the manual workflow — install nothing, for a
   stdlib-only script.

PyInstaller itself is always installed, pinned, whichever branch runs.

## Building on your own machine

Same script the workflows call, so the result matches:

```bash
powershell -ExecutionPolicy Bypass -File build-exe.ps1 -Script scripts\hello.py
```

Dependencies go into a throwaway `.venv-build\`, not your normal Python. Needs
Python installed locally; the workflows don't.

Useful parameters — `-Name`, `-Mode onedir`, `-Console windowed`, `-Icon app.ico`,
`-Requirements <path|none>`, `-ExtraArgs`, `-SmokeArgs`.

## Three things that bite with PyInstaller

Worth reading before your first real build — these cause almost every "it built
but won't run".

- **Data files aren't detected.** Images, templates, `.json` config next to your
  script: add `--add-data "assets;assets"` (on Windows the separator between
  source and destination is `;`, not `:`).
- **Dynamic imports are invisible.** Anything imported by name at runtime —
  plugin loaders, `importlib`, some database drivers — is left out, so the exe
  builds cleanly and dies on startup. Fix with `--hidden-import package.module`,
  and pass `-SmokeArgs "--version"` so the build catches it instead of you.
- **Paths change when frozen.** `__file__` points inside a temp extraction
  folder at runtime, so writing next to it silently loses data. Use
  `sys.executable`'s folder for files the user should keep, and check
  `getattr(sys, "frozen", False)` when the two paths need to differ.

Also expect a SmartScreen warning on any unsigned exe, and 30-60 MB for anything
that imports pandas or similar. `-Mode onedir` starts faster than `onefile` and
trips antivirus less often, at the cost of shipping a folder instead of one file.

## Layout

| Path | What it is |
| --- | --- |
| `scripts/` | Your scripts. The drop folder. |
| `build-exe.ps1` | All the build logic. Both workflows call this, so CI and local builds can't drift apart. |
| `.github/workflows/build-scripts.yml` | Auto-builds everything in `scripts/` on push. |
| `.github/workflows/build-exe.yml` | Manual build of any path, with the full option set. |
