import os
import shutil
import sys
import win32com.client as win32
from pypdf import PdfWriter
import tkinter as tk
from tkinter import filedialog, messagebox

# ============================================================
# CONFIG — mirrors the VBA constants
# ============================================================
COL_WIR_NO      = "B"
COL_REV_NO      = "O"
COL_ATTACHMENT1 = "U"
COL_ATTACHMENT2 = "W"
COL_ATTACHMENT3 = "Y"
COL_ATTACHMENT4 = "AA"
COL_STATUS      = "AC"
COL_CKL_SHEET   = "P"

SHEET_WIR_ELE  = "WIR-Form"
STATUS_DONE    = "Done"
OUTPUT_FOLDER  = "WIRs"
FIRST_DATA_ROW = 3

LIGHT_RED = 0xC8C8FF   # BGR order for win32com (= RGB 255,200,200)
SOLID_RED = 0x0000FF   # BGR order for win32com (= RGB 255,0,0)

xlUp                    = -4162
xlNone                  = -4142
xlCalculationManual     = -4135
xlCalculationAutomatic  = -4105
xlTypePDF               = 0


def safe_print(*args, **kwargs):
    if sys.stdout is not None:
        print(*args, **kwargs)


def col_letter_to_num(letter):
    num = 0
    for c in letter:
        num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num


def last_filled_row(sheet, col_letter, start_row=1):
    col_num = col_letter_to_num(col_letter)
    last_row = sheet.Cells(sheet.Rows.Count, col_num).End(xlUp).Row
    return max(last_row, start_row)


def process_attachment(src_sheet, row_idx, col_letter, row_folder, wir_no, suffix, fail_counter):
    col_num = col_letter_to_num(col_letter)
    cell = src_sheet.Cells(row_idx, col_num)
    src_path = str(cell.Text).strip()
    if not src_path:
        return False  # empty cell -- no-op

    if not os.path.isfile(src_path):
        cell.Interior.Color = LIGHT_RED
        fail_counter[0] += 1
        return True

    ext = os.path.splitext(src_path)[1]
    dest_path = os.path.join(row_folder, f"{wir_no}{suffix}{ext}")
    try:
        shutil.copyfile(src_path, dest_path)
        return False
    except Exception:
        cell.Interior.Color = LIGHT_RED
        fail_counter[0] += 1
        return True


def merge_and_cleanup(row_folder, output_root, group_name, fail_counter):
    # Merges all PDFs in row_folder into a single PDF named <group_name>.pdf
    # saved directly in output_root. Deletes row_folder only on success.
    # Sorted by filename for consistent ordering.
    pdf_files = sorted(
        f for f in os.listdir(row_folder) if f.lower().endswith(".pdf")
    )
    if not pdf_files:
        fail_counter[0] += 1
        return False

    merger = PdfWriter()
    try:
        for fname in pdf_files:
            merger.append(os.path.join(row_folder, fname))

        merged_path = os.path.join(output_root, f"{group_name}.pdf")
        with open(merged_path, "wb") as f:
            merger.write(f)
    except Exception:
        fail_counter[0] += 1
        return False
    finally:
        merger.close()

    try:
        shutil.rmtree(row_folder)
    except Exception:
        pass  # merge worked but cleanup failed -- not fatal

    return True


def ele_wir_generate(workbook_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.EnableEvents = False
    merge_fail_count = [0]

    pdf_count = skip_count = fail_count = ckl_issue_count = 0
    attach_fail_count = [0]
    wb = None

    try:
        wb = excel.Workbooks.Open(workbook_path)
        excel.ScreenUpdating = False
        excel.Calculation = xlCalculationManual

        source_sheet = wb.ActiveSheet
        wir_form_sh = wb.Sheets(SHEET_WIR_ELE)

        wb_dir = os.path.dirname(workbook_path)
        output_root = os.path.join(wb_dir, OUTPUT_FOLDER)
        os.makedirs(output_root, exist_ok=True)

        last_row = last_filled_row(source_sheet, COL_WIR_NO)

        for i in range(FIRST_DATA_ROW, last_row + 1):
            row_attach_failed = False
            ckl_issue = False

            for col in (COL_ATTACHMENT1, COL_ATTACHMENT2, COL_ATTACHMENT3, COL_ATTACHMENT4, COL_WIR_NO):
                source_sheet.Cells(i, col_letter_to_num(col)).Interior.ColorIndex = xlNone

            if source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value == STATUS_DONE:
                skip_count += 1
                continue

            wir_no = source_sheet.Cells(i, col_letter_to_num(COL_WIR_NO)).Value
            if not wir_no or str(wir_no).strip() == "":
                skip_count += 1
                continue
            wir_no = str(wir_no).strip()

            rev_raw = source_sheet.Cells(i, col_letter_to_num(COL_REV_NO)).Value or 0
            rev_no = f"{int(rev_raw):02d}"

            row_folder = os.path.join(output_root, f"{wir_no}-R{rev_no}")
            try:
                os.makedirs(row_folder, exist_ok=True)
            except Exception:
                source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value = "Folder create failed"
                fail_count += 1
                continue

            wir_form_sh.Range("BB1").Value = wir_no
            excel.Calculate()

            cover_pdf_path = os.path.normpath(os.path.join(row_folder, f"{wir_no}-R{rev_no}.pdf"))
            safe_print(f"Row {i}: folder exists = {os.path.isdir(row_folder)} -> {row_folder}")
            try:
                wir_form_sh.ExportAsFixedFormat(xlTypePDF, cover_pdf_path)
            except Exception as e:
                safe_print(f"Row {i} ({wir_no}) - PDF export failed: {e}")
                source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value = "PDF export failed"
                fail_count += 1
                continue
            pdf_count += 1

            ckl_sheet_name = str(source_sheet.Cells(i, col_letter_to_num(COL_CKL_SHEET)).Text).strip()
            ckl_form_sh = None

            if ckl_sheet_name == "":
                source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value = "CKL sheet name blank (col P)"
                source_sheet.Cells(i, col_letter_to_num(COL_WIR_NO)).Interior.Color = SOLID_RED
                ckl_issue = True
                ckl_issue_count += 1
            else:
                try:
                    ckl_form_sh = wb.Sheets(ckl_sheet_name)
                except Exception:
                    ckl_form_sh = None

                if ckl_form_sh is None:
                    source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value = f"CKL sheet not found: '{ckl_sheet_name}'"
                    source_sheet.Cells(i, col_letter_to_num(COL_WIR_NO)).Interior.Color = SOLID_RED
                    ckl_issue = True
                    ckl_issue_count += 1
                else:
                    ckl_form_sh.Range("BB1").Value = wir_no

                    if ckl_form_sh.Name == "Comments_Form":
                        cmt_last_row = max(last_filled_row(ckl_form_sh, "B"), 1)
                        ckl_form_sh.PageSetup.PrintArea = f"A1:J{cmt_last_row}"

                    excel.Calculate()

                    ckl_pdf_path = os.path.normpath(os.path.join(row_folder, f"{wir_no}-XCKL.pdf"))
                    try:
                        ckl_form_sh.ExportAsFixedFormat(xlTypePDF, ckl_pdf_path)
                    except Exception as e:
                        safe_print(f"Row {i} ({wir_no}) - CKL PDF export failed: {e}")
                        source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value = "CKL PDF export failed"
                        source_sheet.Cells(i, col_letter_to_num(COL_WIR_NO)).Interior.Color = SOLID_RED
                        ckl_issue = True
                        ckl_issue_count += 1

            for col, suffix in [
                (COL_ATTACHMENT2, "-XTTCH2"),
                (COL_ATTACHMENT1, "-XTTCH1"),
                (COL_ATTACHMENT3, "-XTTCH3"),
                (COL_ATTACHMENT4, "-XTTCH4"),
            ]:
                failed = process_attachment(source_sheet, i, col, row_folder, wir_no, suffix, attach_fail_count)
                row_attach_failed = row_attach_failed or failed

            if row_attach_failed and not ckl_issue:
                source_sheet.Cells(i, col_letter_to_num(COL_WIR_NO)).Interior.Color = LIGHT_RED

            # --- Merge all PDFs in this row's folder into one file, then cleanup ---
            group_name = f"{wir_no}-R{rev_no}"
            merge_and_cleanup(row_folder, output_root, group_name, merge_fail_count)

            if not ckl_issue:
                source_sheet.Cells(i, col_letter_to_num(COL_STATUS)).Value = STATUS_DONE

        wb.Save()

    finally:
        if wb is not None:
            excel.ScreenUpdating = True
            excel.Calculation = xlCalculationAutomatic
            excel.EnableEvents = True
            wb.Close(SaveChanges=False)
        excel.Quit()

    safe_print(f"Created: {pdf_count} | Skipped: {skip_count} | Failed: {fail_count} "
          f"| CKL issues: {ckl_issue_count} | Attach fails: {attach_fail_count[0]} "
          f"| Merge fails: {merge_fail_count[0]}")

    return {
        "created": pdf_count,
        "skipped": skip_count,
        "failed": fail_count,
        "ckl_issues": ckl_issue_count,
        "attach_fails": attach_fail_count[0],
        "merge_fails": merge_fail_count[0],
    }


def select_workbook_and_run():
    root = tk.Tk()
    root.withdraw()  # hide the empty root window, only show the dialogs

    workbook_path = filedialog.askopenfilename(
        title="Select the WIR Register Excel file",
        filetypes=[("All files", "*.*"), ("Excel Macro-Enabled Workbook", "*.xlsm"), ("Excel Workbook", "*.xlsx")]
    )

    if not workbook_path:
        # User clicked Cancel — exit quietly, no error
        root.destroy()
        sys.exit(0)

    workbook_path = os.path.normpath(workbook_path)

    try:
        results = ele_wir_generate(workbook_path)
        report = (
            "WIR Generation completed.\n\n"
            f"PDFs Created     : {results['created']}\n"
            f"Rows Skipped     : {results['skipped']}\n"
            f"Failed Rows      : {results['failed']}\n"
            f"CKL Issues       : {results['ckl_issues']}\n"
            f"Attachment Fails : {results['attach_fails']}\n"
            f"Merge Fails      : {results['merge_fails']}"
        )
        messagebox.showinfo("WIR Generate", report)
    except Exception as e:
        messagebox.showerror("WIR Generate - Error", f"An error occurred:\n\n{e}")
    finally:
        root.destroy()


if __name__ == "__main__":
    select_workbook_and_run()
