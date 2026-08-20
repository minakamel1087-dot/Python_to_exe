# File and Folder Manager

Windows companion application for the live `Files and Folders Manager2.xlsm` workbook.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

Open the workbook in Microsoft Excel first, then run the application and choose
**Connect to active Excel workbook**. The app never runs a workbook macro. It reads
and writes the currently open `FileManager` and `FolderManager` sheets directly.

The program has no password, expiry date, remote `text.txt` check, or licence gate.

## Included features

- Scan files or folders, either at one level or recursively, and import their information into the open Excel sheet.
- Edit a proposed name/destination in Excel, then batch rename using those live cell values.
- Copy or move selected Excel rows to one chosen destination, a destination in the Excel row,
  or a subfolder named in the Excel row.
- Send selected or visible Excel rows to the Windows recycle bin after confirmation.
- Merge ordered PDF files without Adobe Acrobat.
- Retain a local activity log.

File operations affect explicitly selected Excel rows by default; the app can instead
process all visible (filtered) Excel rows. When a destination already has the same
name, it appends ` (1)`, ` (2)`, and so on rather than overwriting.
