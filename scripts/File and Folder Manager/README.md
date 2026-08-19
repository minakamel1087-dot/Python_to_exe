# File and Folder Manager

Standalone Windows desktop replacement for `Files and Folders Manager2.xlsm`.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

The program has no password, expiry date, remote `text.txt` check, or licence gate.

## Included features

- Scan files or folders, either at one level or recursively.
- Filter results, edit a proposed name/destination, and batch rename.
- Copy or move selected items to one chosen destination, an item-specific destination,
  or a subfolder named in the **Subfolder** column.
- Send selected items to the Windows recycle bin after confirmation.
- Merge ordered PDF files without Adobe Acrobat.
- Export the visible results to Excel and retain a local activity log.

Operations affect only explicitly selected rows. When a destination already has the
same name, the application appends ` (1)`, ` (2)`, and so on rather than overwriting.
