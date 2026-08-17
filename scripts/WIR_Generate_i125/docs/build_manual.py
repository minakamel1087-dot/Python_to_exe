"""
Builds the WIR Tools user manual as a .docx, via Word itself.

Uses the Selection API rather than Paragraphs.Add: the latter silently
dropped every style and reordered the content when driven over late
binding. Selection types the document in order, exactly as a person would.
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(ROOT, "WIR Tools - User Manual.docx")

sys.path.insert(0, ROOT)
from main import _register_pywin32_dlls  # noqa: E402

_register_pywin32_dlls()

import win32com.client  # noqa: E402

WD_STORY = 6
WD_PAGE_BREAK = 7
NAVY = 0x784E1F           # BGR of RGB(31, 78, 120)
GREY = 0x8C857A
HEADER_FILL = 0xF5EDE3
BORDER = 0xC8C8C8


class Manual:
    def __init__(self, word):
        word.Documents.Add()
        self.doc = word.ActiveDocument
        self.sel = word.Selection
        setup = self.doc.PageSetup
        setup.TopMargin = setup.BottomMargin = 56
        setup.LeftMargin = setup.RightMargin = 62

    def _end(self):
        self.sel.EndKey(WD_STORY)

    def para(self, text="", style="Normal", colour=None, size=None,
             bold=None, mono=False):
        self._end()
        self.sel.Style = style
        font = self.sel.Font
        font.Color = NAVY if style.startswith("Heading") or style == "Title" else (
            colour if colour is not None else 0
        )
        if colour is not None:
            font.Color = colour
        if mono:
            font.Name = "Consolas"
            font.Size = 9
        elif size:
            font.Size = size
        if bold is not None:
            font.Bold = bold
        self.sel.TypeText(text)
        self.sel.TypeParagraph()
        # Reset so the next paragraph is not inheriting this one's look.
        self.sel.Style = "Normal"
        self.sel.Font.Bold = False
        self.sel.Font.Name = "Calibri"
        self.sel.Font.Size = 11
        self.sel.Font.Color = 0

    def code(self, lines):
        for line in lines:
            self.para(line, mono=True, colour=0x505050)

    def bullets(self, items):
        for item in items:
            self.para(item, style="List Bullet")

    def table(self, headers, rows, widths):
        self._end()
        table = self.doc.Tables.Add(self.sel.Range, len(rows) + 1, len(headers))
        table.Borders.InsideLineStyle = 1
        table.Borders.OutsideLineStyle = 1
        table.Borders.InsideColor = BORDER
        table.Borders.OutsideColor = BORDER
        table.Range.Font.Size = 9.5
        table.Range.Font.Name = "Calibri"
        table.Range.ParagraphFormat.SpaceAfter = 3
        table.Range.ParagraphFormat.SpaceBefore = 3

        for index, heading in enumerate(headers, start=1):
            cell = table.Cell(1, index)
            cell.Range.Text = heading
            cell.Range.Font.Bold = True
            cell.Range.Font.Color = NAVY
            cell.Shading.BackgroundPatternColor = HEADER_FILL

        for r, row in enumerate(rows, start=2):
            for c, value in enumerate(row, start=1):
                table.Cell(r, c).Range.Text = str(value)

        for index, width in enumerate(widths, start=1):
            table.Columns(index).Width = width

        table.Rows(1).HeadingFormat = True
        self._end()
        self.sel.TypeParagraph()

    def page_break(self):
        self._end()
        self.sel.InsertBreak(WD_PAGE_BREAK)


# DispatchEx, not Dispatch: a private instance, so Quit() at the end cannot
# take down a Word the user already had open.
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

try:
    m = Manual(word)

    m.para("WIR Tools", style="Title")
    m.para("User manual", size=13, colour=GREY)
    m.para("i125 / Ghaf Woods  -  Work Inspection Requests", size=10, colour=GREY)
    m.para()
    m.para(
        "WIR Tools prepares, checks and generates Work Inspection Requests. It works "
        "directly on the Excel workbook you already have open: every change it makes "
        "appears in your window straight away, exactly as a macro would have done."
    )

    m.para("Contents", style="Heading 1")
    toc = None
    try:
        m._end()
        toc = m.doc.TablesOfContents.Add(m.sel.Range, True, 1, 2)
    except Exception as exc:                                    # noqa: BLE001
        print("TOC not created:", exc)
    m.page_break()

    # ---------------------------------------------------------------------
    m.para("1. What the program does", style="Heading 1")
    m.para(
        "The program replaces the macros that used to live inside the workbook. The "
        "workbook is now just a workbook - the data, the forms and the reference "
        "tables - and this program is the tool that works on it."
    )
    m.para("It does four kinds of job:")
    m.table(
        ["Job", "Buttons"],
        [
            ["Produce documents", "Generate WIRs"],
            ["Repair data", "Fix Attachment Links, Fix WIR Path Prefix, Check Areas"],
            ["Validate before submission", "Check WIRs Before Generate"],
            ["Move data around",
             "Import WIR Data, Import Previous WIR Paths, Copy To History, "
             "Extract PDFs, Clear sheets"],
        ],
        [130, 330],
    )
    m.para(
        "Excel is still used to turn the WIR form and the checklists into PDFs, "
        "because nothing else reproduces that layout. The difference is that the "
        "instructions now live in the program rather than inside the file."
    )

    # ---------------------------------------------------------------------
    m.para("2. Before you start", style="Heading 1")
    m.bullets([
        "Open the workbook i125-WIR Cover Generate.xlsm in Excel. The program will "
        "not open it for you, and nothing works if it is closed.",
        "Select your WIR log sheet. Whichever sheet is active when you press a "
        "button is the sheet the program works on.",
        "Connect to the server if you want attachment paths checked. Without it the "
        "program says so and skips those checks, rather than reporting every link "
        "as broken.",
    ])
    m.para(
        "The sheet you were on when a run starts is the sheet it keeps working on. "
        "Clicking another tab part-way through will not redirect it."
    )

    # ---------------------------------------------------------------------
    m.para("3. The window", style="Heading 1")
    m.para(
        "One window with two tabs. Actions holds the buttons; Last run shows what the "
        "most recent run did. The status line along the bottom gives the short version."
    )
    m.para("The buttons are grouped in three:")
    m.table(
        ["Group", "What it is for", "Buttons"],
        [
            ["Tools", "Getting data in and out",
             "Import WIR Data, Import Previous WIR Paths, Extract PDFs, Clear Any Sheet"],
            ["Verify and maintain", "Cleaning and checking before you submit",
             "Fix Attachment Links, Fix WIR Path Prefix, Check Areas, "
             "Check WIRs Before Generate"],
            ["Generate", "Producing and filing the documents",
             "Generate WIRs, Copy To History, Clear WIR Sheet"],
        ],
        [95, 150, 215],
    )
    m.para(
        "Run all checks, the wide button underneath, runs Fix Attachment Links, then "
        "Check Areas, then Check WIRs Before Generate. That order matters: repairing "
        "the links first stops the attachment check reporting failures that are not "
        "real, and the area pass writes its suggestions before the final check reads "
        "the rows."
    )

    # ---------------------------------------------------------------------
    m.para("4. The everyday workflow", style="Heading 1")
    m.para("A normal batch of WIRs goes like this.")
    m.table(
        ["Step", "Do this", "Why"],
        [
            ["1", "Import WIR Data",
             "Reloads the register so the checks can see recent approvals. Without "
             "this, work approved since your last import still looks outstanding."],
            ["2", "Fill in your rows",
             "WIR number, package, building, floor, area, activity, rev, and the "
             "attachment path."],
            ["3", "Import Previous WIR Paths",
             "Fills column AA with the approved WIRs that came before each row."],
            ["4", "Run all checks",
             "Repairs the links, suggests area corrections in column K, and produces "
             "the PreFlight report."],
            ["5", "Fix what it found",
             "Errors must be dealt with. Warnings are for you to judge."],
            ["6", "Generate WIRs",
             "Produces one merged PDF per WIR in the WIRs folder, and marks each row "
             "Done."],
            ["7", "Copy To History",
             "Files the visible rows in WIR_History, and can draft the submission "
             "email."],
        ],
        [40, 150, 270],
    )

    # ---------------------------------------------------------------------
    m.para("5. What each button does", style="Heading 1")

    m.para("Generate WIRs", style="Heading 2")
    m.para(
        "For every row not marked Done: renders the cover sheet, renders the checklist "
        "named in column P, copies in whatever the attachment columns point at, and "
        "merges the whole lot into a single PDF."
    )
    m.para(
        "Output goes to the WIRs folder beside the workbook. You are not asked where - "
        "it is always the same place. Each row is marked Done, so running it again "
        "skips them."
    )
    m.para("Merging behaves like this:")
    m.table(
        ["What was in the folder", "Merged PDF", "The folder"],
        [
            ["All PDFs", "Created", "Deleted - nothing left over"],
            ["A PDF and a photo", "Created", "Kept - the photo could not be merged"],
            ["A PDF that will not open", "Not created",
             "Kept, untouched - merge it by hand"],
            ["No PDFs at all", "Not created", "Kept"],
        ],
        [150, 105, 205],
    )
    m.para(
        "When a merge fails nothing half-finished is left behind, because a partly "
        "written PDF looks like a finished document and would eventually be submitted."
    )

    m.para("Fix Attachment Links", style="Heading 2")
    m.para(
        "Repairs the attachment paths in column N of the log. Three faults are common, "
        "and all three are handled:"
    )
    m.code([
        '"Z:\\WIR request\\..."    pasted with quotes around it',
        'S:\\WIR request\\...      the share mapped to a different letter',
        'Z:\\WIR request\\...      the \\Common\\ level missing',
    ])
    m.para(
        "A link that already works is never touched, and a local path - C: through H: - "
        "is never touched either, because those are real folders on your PC. A "
        "correction is only written back when the files are actually there afterwards. "
        "If the correction finds nothing either, your text is left exactly as you typed "
        "it and the cell turns red for you to look at."
    )

    m.para("Fix WIR Path Prefix", style="Heading 2")
    m.para(
        "Repoints the register's commented-WIR paths from the network share to the "
        "local extract folder, which is the copy that actually opens. This runs "
        "automatically at the end of Import WIR Data; the button is there for when the "
        "register has been edited by hand."
    )

    m.para("Check Areas", style="Heading 2")
    m.para(
        "Reads the Area column and writes into column K: a tidied version of the area "
        "when it can produce one, a short note when it cannot, or both. Column K is "
        "filled amber where there is a note."
    )
    m.para(
        "The Area column itself is never changed, so nothing is lost if a suggestion is "
        "wrong. You copy across what you agree with.", bold=True
    )
    m.para("Examples of what it tidies:")
    m.code([
        "301 302 303 304           ->  Apartments 301, 302, 303, 304",
        "101,102,103               ->  Apartments 101, 102, 103",
        "Flats 301, 302            ->  Apartments 301, 302",
        "Cooridor                  ->  Corridor",
        "Garabge room              ->  Garbage Room",
        "Staircase 1, 2            ->  Staircase 1, Staircase 2",
    ])
    m.para(
        "Rows on floors without an agreed area scheme are skipped - staircases, risers, "
        "podium and upper roof. Only BS, GF, 1F to 10F and RF are checked."
    )

    m.para("Check WIRs Before Generate", style="Heading 2")
    m.para(
        "The full pre-flight. It writes the PreFlight sheet and colours the log, and "
        "changes nothing else. Section 7 explains how to read it."
    )

    m.para("Import WIR Data", style="Heading 2")
    m.para(
        "Asks for the master register file, then rebuilds the WIRs sheet from its ELE "
        "and ELV sheets. Everything the checks know about what has been approved comes "
        "from here, so run it at the start of a batch."
    )

    m.para("Import Previous WIR Paths", style="Heading 2")
    m.para(
        "For each pending row, finds the approved WIRs for the activities that must "
        "come before it at the same package, building and floor, and writes their file "
        "paths into column AA."
    )
    m.para(
        "It understands split activities in both directions: if Wires_Pulling was "
        "raised as Wires_H.L plus Wire_PWR_Pulling that counts, and so does the reverse."
    )
    m.para(
        "Rows it skips - hidden, already Done, or missing the fields it matches on - "
        "keep whatever was already in column AA."
    )

    m.para("Copy To History", style="Heading 2")
    m.para(
        "Appends the visible rows to WIR_History with a timestamp. Only visible rows, "
        "because filtering the log is how you choose what goes out."
    )
    m.para(
        "It asks whether you also want the submission email. Answer yes and it "
        "opens a draft in Outlook with a table of the same rows, for you to check "
        "and address. It opens the draft; it never sends it."
    )

    m.para("Extract PDFs", style="Heading 2")
    m.para(
        "Asks for a folder, then copies whatever the visible rows' Link column points "
        "at into it. Original filenames are kept, and a clash is numbered rather than "
        "overwritten."
    )

    m.para("Clear WIR Sheet and Clear Any Sheet", style="Heading 2")
    m.para(
        "Clear Any Sheet wipes rows 3 to 1000 of whichever sheet is active. Clear WIR "
        "Sheet does the same to the log and then puts the template row back - the "
        "formulas, today's and tomorrow's dates, and Rev 0."
    )
    m.para("Both ask you to confirm first. Neither can be undone.", bold=True)

    # ---------------------------------------------------------------------
    m.para("6. Where the results appear", style="Heading 1")
    m.table(
        ["Where", "What it means", "Put there by"],
        [
            ["Column K", "Suggested area text, or a note about it. Amber where there "
                         "is a note.", "Check Areas"],
            ["Column N, red", "The attachment path could not be found, even after "
                              "correcting it.", "Fix Attachment Links"],
            ["Column B, red", "This row has at least one error.", "Check WIRs"],
            ["Column B, amber", "This row has warnings but no errors.", "Check WIRs"],
            ["Column I", "Coloured when the problem is with this row's area.",
             "Check WIRs"],
            ["Column AA", "Paths to the approved previous WIRs.",
             "Import Previous WIR Paths"],
            ["Column AC", "Done, once the row has been generated.", "Generate WIRs"],
            ["PreFlight sheet", "The full report, one line per row.", "Check WIRs"],
            ["WIRs folder", "One merged PDF per WIR.", "Generate WIRs"],
        ],
        [95, 240, 125],
    )
    m.para(
        "Only columns B and I are coloured by the checks, and both are cleared at the "
        "start of every run - so a problem you have fixed stops being flagged, and the "
        "attachment colouring in the other columns is never disturbed."
    )

    # ---------------------------------------------------------------------
    m.para("7. Reading the PreFlight report", style="Heading 1")
    m.para(
        "One line per log row, not one per problem. If a row has several problems they "
        "are all in the same cell, with the errors first. The row number links back to "
        "the log, and error rows are listed above warning rows."
    )

    m.para("Errors - fix these before generating", style="Heading 2")
    m.table(
        ["Check", "Meaning"],
        [
            ["WIR number", "Wrong length, or the package or trade inside the number "
                           "does not match columns F and E."],
            ["Duplicate number", "The same WIR number appears twice in the log."],
            ["Already submitted", "This number is already in the register. Raise the "
                                  "revision, or use a new number."],
            ["Missing fields", "Dep., Package, Building, Floor or Activity is empty."],
            ["Rev", "Column O is not a number."],
            ["Checklist", "Column P is empty, or names a sheet that is not in the "
                          "workbook."],
        ],
        [120, 340],
    )

    m.para("Warnings - judge these yourself", style="Heading 2")
    m.table(
        ["Check", "Meaning"],
        [
            ["Previous activity", "Something that must come first has no approved WIR "
                                  "covering every area on this row. The report names "
                                  "the areas, and says whether the earlier WIR was "
                                  "rejected, under review, or never submitted."],
            ["Area submitted before", "The same activity is already approved here for "
                                      "one of this row's areas, under another WIR "
                                      "number."],
            ["Apartment", "A unit number is not in the register for this floor."],
            ["Area name", "The area is not in Area_Names, or it is on a floor it "
                          "cannot be on."],
            ["Attachment", "A path that was given does not resolve."],
            ["Description", "Column C is empty, so the cover sheet will have no "
                            "description."],
        ],
        [120, 340],
    )
    m.para(
        "Attachment coverage at the top of the report is a count rather than a list of "
        "faults. The QC column in particular points at a folder that only exists once "
        "QC has filed something, so an empty one is normal."
    )

    # ---------------------------------------------------------------------
    m.para("8. The tables you maintain", style="Heading 1")
    m.para(
        "These live in the workbook and are yours to edit. The program finds them by "
        "name, so you can move them between sheets without breaking anything."
    )
    m.table(
        ["Table", "Where", "What it controls"],
        [
            ["Area_Names", "Project-Des.",
             "Valid area names, their accepted misspellings, and which floors an area "
             "may appear on. The first column is the spelling the program suggests."],
            ["Activity_Splits", "ELE_Data",
             "Activities raised as two or more sub-activities. When all the subs are "
             "approved, the parent counts as done."],
            ["WIR_Status_Codes", "Project-Des.",
             "What each register status letter means, and which count as approved. "
             "Only B counts today."],
            ["Prev_Activities", "ELE_Data",
             "Which activities must be approved before each activity."],
            ["Floor_Apartments", "Project-Des.",
             "The apartments that exist on each building and floor."],
        ],
        [110, 85, 265],
    )
    m.para(
        "If a genuine area keeps being reported as unknown, add it to Area_Names and "
        "the message stops. If a misspelling keeps appearing, add it as a variant on "
        "the row for the correct name and it will be corrected automatically from then "
        "on."
    )

    # ---------------------------------------------------------------------
    m.para("9. Troubleshooting", style="Heading 1")
    m.table(
        ["What you see", "What it means"],
        [
            ["Excel is not running",
             "Open the workbook first. The program attaches to Excel; it does not "
             "start it."],
            ["... is not open in Excel",
             "The workbook is not open, or a different file of that name is."],
            ["... is not a WIR log sheet",
             "You were on PreFlight, WIRs or another reserved sheet. Click your log "
             "sheet and try again."],
            ["Share Z:\\Common\\ is not reachable",
             "You are off the network. The link check is skipped rather than marking "
             "every row broken. Reconnect and run it again."],
            ["Every attachment shows as not found",
             "Almost always the same thing - the share is not connected."],
            ["Table 'Prev_Activities' not found",
             "The reference table has been renamed or deleted. It is found by name, "
             "not by position."],
            ["Sheet 'WIRs' not found",
             "The register has not been imported into this workbook yet."],
            ["Nothing was checked",
             "Every row was marked Done. Clear column AC on the rows you want looked "
             "at."],
        ],
        [175, 285],
    )

    # ---------------------------------------------------------------------
    m.para("10. Things to be aware of", style="Heading 1")
    m.bullets([
        "Typing 101,102,103 into the Area column makes Excel store it as the number "
        "101102103, before any program sees it. The program recovers it when the "
        "numbers match the floor, but typing Apartments first keeps the cell as text.",
        "Hyphens are not treated as separators between areas, because they are part of "
        "names like Part-5 and Mock-up. Use commas between areas.",
        "A value like Substation South _Attic is reported rather than corrected, "
        "because correcting it would mean guessing which half to keep.",
        "The checks only know what the register knows. Import WIR Data at the start of "
        "a batch, or recently approved work will still look outstanding.",
        "Only ELE and ELV are imported from the register, so plumbing rows are not "
        "precedence-checked.",
        "Generate marks rows Done as it goes. To regenerate a row, clear its Status "
        "cell first.",
    ])

    if toc is not None:
        try:
            toc.Update()
        except Exception as exc:                                # noqa: BLE001
            print("TOC not updated:", exc)

    m.doc.SaveAs2(OUTPUT, FileFormat=16)
    print("saved :", OUTPUT)
    print("pages :", m.doc.ComputeStatistics(2))
    print("words :", m.doc.ComputeStatistics(0))
    m.doc.Close(False)
finally:
    word.Quit()
