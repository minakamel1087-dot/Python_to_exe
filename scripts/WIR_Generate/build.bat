@echo off
REM One-file build. --noconsole hides the terminal; drop it while debugging.
pyinstaller --noconfirm --onefile --noconsole --name "WIR Tools" ^
    --hidden-import win32timezone ^
    main.py
