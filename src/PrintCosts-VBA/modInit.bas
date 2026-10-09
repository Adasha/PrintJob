Attribute VB_Name = "modInit"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' One-time setup, run once after the modules are imported.
'
' The workbook file is written outside Excel, which can lay out cells but
' cannot draw controls. This turns the placeholder labels into Form Control
' buttons, applies protection, and leaves the file ready to save as .xlsm.
'
' It is safe to run again at any time - it removes the buttons it created
' before redrawing them - and must be run again after a new print room sheet
' is duplicated.

Private Const BTN_TAG As String = "pcb_"

' Column where a location sheet's occasional-use buttons live (2026-09-25
' layout fix, see DrawLocationButtons) - well past the job table's own
' columns (34, S_SchemaVer) and the AZ1/52 marker, so ApplyColumnVisibility
' can never hide a button by hiding the column under it. Same pattern
' Summary's own buttons already use at column O, clear of its A:M table.
Private Const SIDE_PANEL_COL As Long = 47

' Snag list item 1e's reduced-clutter view: the default hidden-column list,
' used when the SET_LOC_REDUCED_COLUMNS setting is blank or missing (SettingText's
' default). PrintCosts.xlsx ships the setting row with this same list. Declared
' here with this module's other module-level constant (BTN_TAG) rather than down by the code that
' uses it - every other module in this project keeps its Const/Dim
' declarations clustered at the top, and a module-level Const declared after
' a Sub/Function has already appeared in the source left this workbook
' failing to compile ("Variable not defined") even though the declaration
' itself was syntactically fine on its own.
' Location sheet's toolbar row. PrintCosts.xlsx ships the final layout: the
' config block in A1:B11, a blank row 12, this toolbar on row 13 and the job
' table's header on row 15.
Private Const TOOLBAR_ROW As Long = 15
' The pass buttons' row (multi-pass costing): the blank row just above the table header.
Private Const PASS_ROW As Long = 16
' Header block rows on a location sheet (short blank gap rows sit between the groups).
Private Const ROW_NAME As Long = 1
Private Const ROW_DEPT As Long = 2
Private Const ROW_DEF_TECH As Long = 4
Private Const ROW_DEF_PRINTER As Long = 5
Private Const ROW_DEF_PAPER As Long = 6
Private Const ROW_DIS_PAPER As Long = 8
Private Const ROW_DIS_CONS As Long = 9
Private Const ROW_ROLL_UNIT As Long = 11
Private Const ROW_PRINTERS As Long = 12
Private Const ROW_JOB_COUNT As Long = 13
Private Const REDUCED_COLUMNS_DEFAULT As String = "Status;Job ID;Area m2;S_SchemaVer"
' Minimal view hides these IN ADDITION to the Reduced list (so Minimal always
' includes Reduced - the setting holds only the extras).
Private Const MINIMAL_EXTRA_DEFAULT As String = "Printer;Disregard Paper;Disregard Consumable;Print Width mm;Sheet size"

' The view row on a location sheet: a "Show columns:" label, an All/Reduced/
' Minimal drop-down and the Hide/Show cost detail button, laid out left to
' right from VIEW_FIRST_COL (D) on VIEW_ROW (2), each on a visible column - see
' RepositionViewButtons.
Private Const VIEW_ROW As Long = 2
Private Const VIEW_FIRST_COL As Long = 4
Private Const VIEW_LABEL_W As Double = 75
Private Const VIEW_DD_W As Double = 90
Private Const VIEW_COST_W As Double = 110
Private Const VIEW_BTN_GAP As Double = 2
Private Const VIEW_DD_NAME As String = "ddViewMode"

' SidePanelButtonLayout's own inputs (2026-09-27, cost-columns toggle) -
' declared here with this module's other module-level constants for exactly
' the reason REDUCED_COLUMNS_DEFAULT's own comment just above gives. Learn
' from that one: a first draft of this pair was declared mid-file, next to
' SidePanelButtonLayout itself, and hit the identical "Variable not defined"
' compile failure this comment already warns about.
Private Const SIDE_PANEL_BTN_COUNT As Long = 6

' Settings sheet's button band (2026-10-02): the three rows between the sheet's
' description (row 2) and the tables' subtitle row, one per button group. The
' shipped .xlsx has the rows, with tblSettings' header on row 8;
' EnsureSettingsButtonBand labels and sizes them. Declared here with the other
' module-level constants - see REDUCED_COLUMNS_DEFAULT's note.
Private Const SETTINGS_BAND_FIRST_ROW As Long = 3
Private Const SETTINGS_BAND_ROWS As Long = 3
Private Const SETTINGS_BAND_ROW_H As Double = 26
Private Const SETTINGS_BTN_W As Double = 130
Private Const SETTINGS_BTN_GAP As Double = 4
Private Const SIDE_PANEL_MIN_GAP As Double = 3

' Set by a setup run and consumed by the message Refresh Locations shows, so a
' full setup can say how many buttons it drew without a second dialog.
Public gButtonsDrawn As Long

Public Sub InitialiseWorkbook()
    On Error GoTo Fail
    Dim ws As Worksheet, n As Long

    AppOff
    If ThisWorkbook.Date1904 Then ThisWorkbook.Date1904 = False

    ' _Registry, _Audit and _Data are built here rather than shipped in the
    ' .xlsx, so the workbook file stays something VBA can reconstruct.
    EnsureSystemSheets

    ' The version and schema cells are stamped from the code. Setup does not
    ' stamp the build date - that is build.ps1's job, via StampBuild.
    StampVersionSettings
    EnsureSuppliedColumnValidation
    ' The Departments sheet (modDepartments) is likewise built here, not shipped.
    EnsureDepartmentsSheet
    ' So is the Consumables sheet (modConsumables): the colours a multi-pass printer uses.
    EnsureConsumablesSheet

    ' Runs before the per-sheet loop, not after it (2026-09-27). Row heights
    ' have to be settled before anything is positioned from them, and the
    ' Settings buttons are free-floating (DrawOne's note), so they do not
    ' follow a later row-height change. (They now sit above the tables, which
    ' this cannot move, but the ordering is still the safe one.)
    FormatSettingsNotes

    ' Summary and Reports, likewise built here. Their formulas read
    ' _Data, which RefreshLocations writes at the end of this run - until then
    ' they sit on their IFERROR fallbacks rather than showing errors.
    BuildReportSheets

    For Each ws In ThisWorkbook.Worksheets
        UnlockSheet ws
        ClearButtons ws
        If IsLocation(ws) Then
            DrawLocationButtons ws
            ConfigValidation ws
            ApplyStatusFormat ws
            ' Widths are set here, before RefreshLocations (which ends this run)
            ' applies the view and cost-column visibility. That pass also
            ' re-settles the buttons once visibility is final, as do the other
            ' per-sheet steps it repeats (defaults, roll unit, dropdowns), so
            ' none of them is duplicated in this loop.
            n = n + 1
        ElseIf StrComp(ws.Name, "Summary", vbTextCompare) = 0 Then
            ' Column O onwards, clear of the A:M report table.
            DrawOne ws, 1, 15, "Refresh Locations", "btnRefreshLocations", 130
            DrawOne ws, 3, 15, "Check workbook", "btnCheckWorkbook", 130
            ' "Go to Settings" removed (snag 3b, 2026-09-22): with four
            ' configuration sheets and no way to tell which one a given task
            ' needs, the button could only ever jump to one of them (Settings)
            ' - not a fix worth making target-aware when Hide/Show settings
            ' sheets already reveals all four and Excel's own sheet tabs reach
            ' any of them directly, the easiest fix the snag list itself named.
            ' btnToggleConfigSheets, not btnToggleSettingsSheets: the latter
            ' is exactly 32 characters with this row/column, which Button.Name
            ' silently truncates to 31 rather than erroring on (SetButtonName's
            ' comment) - caught by the length audit that found the Settings
            ' lookup-table buttons' own truncation.
            DrawOne ws, 5, 15, ConfigToggleCaption(), "btnToggleConfigSheets", 130
            ' The job planner box (modPlanner) at the top of the page.
            DrawOne ws, PLN_ADD_ROW, PLN_ADD_COL, "Add to print room", "btnPlannerAdd", 130
        ElseIf StrComp(ws.Name, "Printers", vbTextCompare) = 0 Then
            DrawOne ws, 4, 1, "Add row", "btnAddRowPrinters", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowPrinters", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearPrinters", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Papers", vbTextCompare) = 0 Then
            DrawOne ws, 4, 1, "Add row", "btnAddRowPapers", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowPapers", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearPapers", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Print Technicians", vbTextCompare) = 0 Then
            DrawOne ws, 4, 1, "Add row", "btnAddRowTechnicians", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowTechnicians", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearTechnicians", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Consumables", vbTextCompare) = 0 Then
            DrawOne ws, 4, 1, "Add row", "btnAddRowColours", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowColours", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearColours", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Departments", vbTextCompare) = 0 Then
            DrawOne ws, 4, 1, "Add row", "btnAddRowDepartments", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowDepartments", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearDepartments", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Settings", vbTextCompare) = 0 Then
            ' Buttons sit in a band ABOVE the tables (2026-10-02, direct user
            ' request; they used to sit below tblSettings, out of sight on a
            ' long table, and before that down column T). Three rows, one per
            ' logical group, each with a label in column A and its buttons
            ' packed left to right from column B with a small fixed gap
            ' (SETTINGS_BTN_GAP) rather than anchored three columns apart,
            ' which spread four buttons across the whole sheet:
            '   Print rooms - Add / Remove print room, Refresh Locations
            '   Data        - Export All, Import, Backup, Restore
            '   Workbook    - Check workbook, Re-stamp prices, About
            ' EnsureSettingsButtonBand writes the labels and row heights; the
            ' rows themselves ship in the .xlsx.
            EnsureSettingsButtonBand ws
            Dim bandRow As Long, bandLeft As Double
            bandRow = SETTINGS_BAND_FIRST_ROW
            bandLeft = ws.Cells(1, 2).Left
            DrawBandButton ws, bandRow, bandLeft, 0, "Add print room...", "btnAddPrintRoom"
            DrawBandButton ws, bandRow, bandLeft, 1, "Remove print room...", "btnRemovePrintRoom"
            DrawBandButton ws, bandRow, bandLeft, 2, "Refresh Locations", "btnRefreshLocations"

            DrawBandButton ws, bandRow + 1, bandLeft, 0, "Export All Locations...", "btnExportAll"
            DrawBandButton ws, bandRow + 1, bandLeft, 1, "Import (choose room)...", "btnImportGlobal"
            DrawBandButton ws, bandRow + 1, bandLeft, 2, "Backup workbook...", "btnBackupWorkbook"
            DrawBandButton ws, bandRow + 1, bandLeft, 3, "Restore workbook...", "btnRestoreWorkbook"

            DrawBandButton ws, bandRow + 2, bandLeft, 0, "Check workbook", "btnCheckWorkbook"
            DrawBandButton ws, bandRow + 2, bandLeft, 1, "Re-stamp prices...", "btnReStamp"
            DrawBandButton ws, bandRow + 2, bandLeft, 2, "About", "btnAbout"
            ' Small +/- buttons above the three lookup tables (snag list item
            ' 5), on the row two above each table's header - the row above
            ' the table's own subtitle ("Paper stock types" etc.), which
            ' differs from the blank row on Printers/Papers/Print Technicians.
            '
            ' "Size"/"Consumable" rather than "StandardSizes"/
            ' "Consumables": Button.Name silently TRUNCATES
            ' to 31 characters at 32 and raises 1004 outright at 33+ in this
            ' Excel/COM automation context (verified directly with a length
            ' sweep - "pcb_btnRemoveRowConsumables_3_18" was exactly 32 and
            ' came back with its trailing "8" dropped, not an error, which
            ' is a worse bug than a clean failure would have been. The other
            ' three, at 34 each, raised 1004 outright.
            Dim smallRow As Long
            smallRow = Tbl("tblSettings").Range.Row - 2
            DrawSmall ws, smallRow, 6, "+", "btnAddRowPaperTypes", 24
            DrawSmall ws, smallRow, 7, "-", "btnRemoveRowPaperTypes", 24
            DrawSmall ws, smallRow, 9, "+", "btnAddRowSize", 24
            DrawSmall ws, smallRow, 10, "-", "btnRemoveRowSize", 24
            DrawSmall ws, smallRow, 13, "+", "btnAddRowConsumable", 24
            DrawSmall ws, smallRow, 14, "-", "btnRemoveRowConsumable", 24
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Reports", vbTextCompare) = 0 Then
            ' Rows 1-3, column T (2026-09-26, moved from F): F sat directly
            ' above the results table/filter block (A:Q), so the buttons
            ' floated over the top of that whole area rather than beside it.
            ' T is clear on these rows - the breakdowns BuildBreakdowns draws
            ' there ("By print room") don't start until row 15 - and is
            ' already the column the breakdowns themselves use, so the
            ' buttons now sit in the same right-hand strip as that content
            ' instead of overlapping the table.
            ' 0.10.25: all Reports header buttons are drawn and re-placed by
            ' modReports.DrawReportsButtons / RepositionReportsButtons (they
            ' follow column visibility/width changes like the location
            ' sheets' buttons). The delete macro is now btnDeleteVisible, so
            ' the shape name needs no truncating tag.
            ' 0.10.26: the strip starts at column N (not T, as above) and runs
            ' Clear all filters | Selected record: | Mark all as... | Export /
            ' Delete; the two captions are free-floating labels drawn there too.
            DrawReportsButtons ws
            ' No freeze panes on Reports (0.10.28, direct user request): with the
            ' filters beside the results table there is nothing worth pinning, and
            ' a frozen top block would cost screen height. Clears any freeze left
            ' by an earlier build.
            SetFreeze ws, ""
        End If
    Next ws

    UnlockConfigInputs
    ReorderSheetTabs

    Invalidate
    ProtectAll
    AppOn

    ' Refresh does the rest - codes, table names, the registry and the
    ' reporting range - and reports on the whole run, so there is one message
    ' at the end rather than two.
    gButtonsDrawn = CountButtons
    RefreshLocations
    Exit Sub
Fail:
    AppReset
    ReportError "InitialiseWorkbook"
End Sub

Private Sub DrawLocationButtons(ByVal ws As Worksheet)
    ' The toolbar's row is TOOLBAR_ROW (15). The config block (A1:B13, with two
    ' short gap rows) sits above it. Two columns
    ' apart; buttons are drawn over the placeholders and the labels cleared.
    '
    ' Layout fix (2026-09-25, user-reported): a button anchored over a column
    ' the reduced-clutter view hides (§4.1's known gap) vanishes along with
    ' it - Remove Row sat on Printer, Clear All on Disregard Consumable, both
    ' hidden by REDUCED_COLUMNS_DEFAULT. Fix has two parts:
    ' 1. Only the two buttons used on every job entry stay here, close to
    '    hand - everything used occasionally rather than per-job moves to a
    '    side panel at SIDE_PANEL_COL, past the table's own columns entirely,
    '    so no column-hide can ever reach it (below).
    ' 2. The few buttons still anchored inside the table's column span
    '    (these two, plus Clear defaults) are covered by RelocateAtRiskButtons
    '    instead, which moves any of them off a column the CURRENT
    '    reduced-view setting hides - general protection, since
    '    SET_LOC_REDUCED_COLUMNS is user-editable and a fixed anchor choice
    '    can't stay safe forever (§4.1's own "not a general solution" note).
    ' Three buttons, two columns apart each (comment above) - Repeat Job
    ' slots into the same "skip a column, land the next button's overflow
    ' harmlessly on empty space" pattern Add Print Job/Now already use, one
    ' step further right: column 3 (Student No) rather than 5 (Printer) is
    ' what keeps it nestled between the other two instead of past both of
    ' them. Now moves from 3 to 5 to make room - its own 110pt overflows
    ' column 5 (Printer) the same harmless way it used to overflow column 3.
    DrawOne ws, TOOLBAR_ROW, 1, "Add Print Job", "btnAddPrintJob", 110
    DrawOne ws, TOOLBAR_ROW, 3, "Repeat Job", "btnRepeatJob", 110
    DrawOne ws, TOOLBAR_ROW, 5, "Now", "btnNow", 110

    ' Pass controls (multi-pass costing, 0.11.0), on the row between the job
    ' buttons and the table header so they read as a separate group. They are
    ' on every location sheet and shaded lighter when no multi-pass printer is
    ' set up (modPasses.RefreshPassButtons). Columns 1, 2, 4, 5 are all wide
    ' enough for the button and are kept on screen by RelocateAtRiskButtons.
    ' The job buttons above are 22pt tall, taller than their 14.5pt row, so they ran
    ' on into this one and the pass buttons drew over them. Both rows are made tall
    ' enough to hold their own buttons, with a 2pt margin inside the pass row.
    If ws.Rows(TOOLBAR_ROW).RowHeight < 24 Then ws.Rows(TOOLBAR_ROW).RowHeight = 24
    If ws.Rows(PASS_ROW).RowHeight < 26 Then ws.Rows(PASS_ROW).RowHeight = 26
    Dim passTop As Double
    passTop = ws.Rows(PASS_ROW).Top + 2
    DrawOneAtTop ws, 1, passTop, "Add pass", "btnAddPass", 100
    DrawOneAtTop ws, 2, passTop, "Remove pass", "btnRemovePass", 100
    DrawOneAtTop ws, 4, passTop, "Toggle passes", "btnTogglePasses", 100
    DrawOneAtTop ws, 5, passTop, "Toggle all passes", "btnToggleAllPasses", 110

    ' View row (2026-10-01): "Show columns:" label, an All/Reduced/Minimal
    ' drop-down and the Hide/Show cost detail button on row 2, from D.
    ' Positioned (and kept off hidden columns) by RepositionViewButtons; drawn
    ' here at their nominal columns.
    DrawViewDropDown ws
    DrawOne ws, VIEW_ROW, VIEW_FIRST_COL + 2, CostColumnsCaption(), "btnToggleCostColumns", VIEW_COST_W
    DrawOne ws, VIEW_ROW, VIEW_FIRST_COL + 3, MultiPassColsCaption(), "btnToggleMultiPass", VIEW_COST_W

    Dim c As Long
    For c = 1 To 15
        With ws.Cells(TOOLBAR_ROW, c)
            .ClearContents
            .Interior.Pattern = xlNone
        End With
    Next c

    ' Side panel - occasional-use buttons, one stack, matching the idiom
    ' Summary's own off-table buttons use. Grouped loosely by kind: setup/
    ' maintenance/view, then row-level correction, then whole-sheet data
    ' movement.
    '
    ' SIDE_PANEL_COL widened to fit a whole button (2026-09-25, user-
    ' reported): a Buttons.Add shape's own Width is independent of the
    ' underlying cell/column width it is anchored to, so at the column's
    ' previous (narrow, default) width every 140pt-wide button here sprawled
    ' across two or three columns rightward - straight over the relocated
    ' settings block at AL/AM, blocking its Yes/No dropdown arrows from
    ' being clicked (the shape sits above the cell in z-order and intercepts
    ' the click) as well as simply looking wrong. Widening the column to
    ' comfortably exceed the button's own width keeps every button's
    ' footprint contained within its one column, so nothing to its right is
    ' ever at risk regardless of what lands there later.
    UnlockSheet ws
    ws.Columns(SIDE_PANEL_COL).ColumnWidth = ColWidthForPx(210)
    RelockSheet ws
    '
    ' Reduced-clutter view toggle moved here 2026-09-25 (was row 10, column
    ' 7, over Unit): not used often enough to earn a spot near the table, and
    ' its previous position - however carefully chosen - was still only ever
    ' safe by construction against today's REDUCED_COLUMNS_DEFAULT, not
    ' against whatever SET_LOC_REDUCED_COLUMNS might later be edited to name.
    ' Living in the side panel sidesteps that question entirely rather than
    ' relying on RelocateAtRiskButtons to keep dodging it, so it's dropped
    ' from that Sub's list too (below). Caption still read fresh from the
    ' current setting each time this runs, same as ConfigToggleCaption's
    ' button does.
    '
    ' All eight now share one column (2026-09-25, user-reported: Import had
    ' drifted onto SIDE_PANEL_COL2, away from the rest; joined 2026-09-27 by
    ' the cost-columns toggle, moved here for the same reason the reduced-view
    ' toggle was - see CostColumnsCaption's own comment). Every-other-row
    ' spacing only had room for six before the table header - packed instead
    ' via DrawOneAtTop at even pixel steps spanning the space actually
    ' available above the header (whatever row that currently is), so the
    ' whole stack fits in one column with a consistent, if tighter, gap
    ' between buttons - narrower than a full spare row, but comfortably more
    ' than nothing.
    Dim lo As ListObject, headerTop As Double, stepPx As Double, btnH As Double, i As Long
    Set lo = JobsTable(ws)
    If Not lo Is Nothing Then
        headerTop = ws.Cells(lo.Range.Row, 1).Top
        SidePanelButtonLayout headerTop, btnH, stepPx
        Dim capts() As Variant, macros() As Variant
        capts = Array("Select printers...", "Check this sheet", "Remove Row", "Clear All", "Export...", "Import...")
        macros = Array("btnSelectPrinters", "btnCheckSheet", "btnRemoveRow", "btnClearAll", "btnExport", "btnImportLocation")
        For i = 0 To UBound(macros)
            DrawOneAtTop ws, SIDE_PANEL_COL, i * stepPx, CStr(capts(i)), CStr(macros(i)), 140, btnH
        Next i
    End If

    ' Clear defaults - column D, middle row of the three-row default-selector
    ' block it clears (A4:B6), without overlapping the labels themselves -
    ' column D is free there (D1:D2 are cleared by EnsureJobDefaults). Still anchored
    ' inside the table's column span, so RelocateAtRiskButtons covers it
    ' (below).
    DrawOne ws, ROW_DEF_PRINTER, 4, "Clear defaults", "btnClearDefaults", 110
End Sub

' The side panel's shared button height and vertical step, derived from
' whatever room is actually available above the table's CURRENT header row
' rather than a hardcoded figure (same "read at runtime" reasoning
' DrawOneAtTop's own comment gives for pixel-based stacking generally).
'
' Direct user report, 2026-09-27: adding the cost-columns toggle as an
' EIGHTH button at the original fixed 22pt height made every button overlap
' its neighbour by roughly half a point, compounding to several points by
' the last one - the seven-button layout had run flush against the
' available space with nothing to spare from the day it was built (7*22 +
' 6*3pt gaps already came to within half a point of the full budget), so
' adding one more at the same height literally could not fit, gap or no
' gap. Growing the available space isn't a small change here: the panel
' already starts at row 1 with nowhere to extend upward, and pushing the
' table header down another row touches far more of the sheet's layout than
' this warrants - see this section's own history of layout bugs (the
' row-scope bug, column drift, button-width corruption, all above) for why
' that is not a change to make lightly. Shrinking the button height instead
' is genuinely self-contained: nothing outside DrawLocationButtons and
' RepositionSidePanelButtons (which MUST use this same function, not a
' second hardcoded height, or every reposition call would silently grow the
' buttons back to 22pt and reopen this exact bug) reads or assumes a side
' panel button's height.
'
' MIN_GAP is a floor, not a target - a real gap, however small, so eight
' buttons still read as eight buttons rather than one fused block. Height
' shrinks only as far as fitting the current count actually requires,
' capped at the standard 22pt so seven-or-fewer buttons (or a future build
' with SET_LOC_REDUCED_COLUMNS-style flexibility removing one) look exactly
' as they always did.
Private Sub SidePanelButtonLayout(ByVal HeaderTop As Double, ByRef ButtonHeight As Double, ByRef StepPx As Double)
    Dim available As Double
    ' 8pt safety margin before the header - packing the last button's bottom
    ' edge flush against it left effectively no clearance at all.
    available = HeaderTop - 8
    ButtonHeight = (available - (SIDE_PANEL_BTN_COUNT - 1) * SIDE_PANEL_MIN_GAP) / SIDE_PANEL_BTN_COUNT
    If ButtonHeight > 22 Then ButtonHeight = 22
    StepPx = ButtonHeight + SIDE_PANEL_MIN_GAP
End Sub

' Yes/No validation on the two location defaults (spec 9.2). These seed each
' new print job and are plain cells in the workbook file, so the list is added
' here rather than shipped with it.
' Unlocks/relocks around its own edit, like every other per-sheet Ensure/
' Apply function here - added 0.9.14, when EnsureJobTableGap became the
' first thing the location loop runs and started leaving the sheet
' genuinely re-protected (its own RelockSheet) by the time this ran. Validation.Add
' raises 1004 under real protection even with UserInterfaceOnly - drawing
' buttons/clearing cells (DrawLocationButtons, between the two) does not,
' which is exactly why only this one broke.
Private Sub ConfigValidation(ByVal ws As Worksheet)
    UnlockSheet ws
    AddYesNo LocRange(ws, "LOC_DefDisPaper"), "Disregard paper cost", "Sets what new print jobs on this sheet start with. Changing it never alters jobs already recorded."
    AddYesNo LocRange(ws, "LOC_DefDisCons"), "Disregard consumable cost", "Sets what new print jobs on this sheet start with. Changing it never alters jobs already recorded."
    RelockSheet ws
End Sub

Private Sub AddYesNo(ByVal target As Range, ByVal Title As String, ByVal Msg As String)
    If target Is Nothing Then Exit Sub
    If Len(Trim$(CStr(target.Value))) = 0 Then target.Value = "No"
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Choose Yes or No."
    End With
End Sub

' ------------------------------------------------------- batch defaults ---
' Layout fix (2026-09-25, same pass as the button relocation above): these
' four rows - the three per-job selectors plus the roll-unit setting - used
' to live spread across columns A-F on their own dedicated row (requiring
' EnsureJobTableGap to open space for it), which is exactly what put
' "Default: paper" on top of the Printer column and made it vanish under
' reduced view (§4.1's cross-column collision, same root cause the button
' fix addressed for shapes). Fixed the same way in spirit: everything that
' must stay visible regardless of the reduced-view setting now lives ONLY in
' columns A/B, which ApplyColumnVisibility never touches (below) - stacked
' as one row each instead of three column-pairs on one row. Location code,
' the two disregard-cost defaults and the permitted-printers list moved out
' to the side panel to make room (rows 3-6 here used to be theirs) - none of
' them are used per-job the way these four are.
'
' Snag list item 1b's own reasoning still holds: a technician sets a
' Technician/Printer/Paper Stock once and has every subsequently added job
' pre-filled from them (modJobs.AddPrintJob), until Clear defaults empties
' them again (modJobs.ClearDefaults). Self-provisioned here rather than
' shipped in the .xlsx - same reasoning as LOC_Export (modExport): a
' sheet-scoped name copies cleanly with a duplicated sheet, and re-adding it
' every run means an older location sheet picks the feature up without hand
' surgery. Public: modRegistry.RefreshLocations also calls this for every
' location on every refresh, alongside BindColumns - the same "rebuild
' dependent dropdowns, self-heal an old sheet, re-point a duplicated one"
' reasoning RefreshExportStatus/EnsureExportName already applies to
' LOC_Export.
Public Sub EnsureJobDefaults(ByVal ws As Worksheet)
    UnlockSheet ws

    ' Room name and department are fixed once set (Add print room asks for
    ' both up front); a blank one stays editable so it can still be filled in.
    With ws.Cells(ROW_NAME, 2)
        .Locked = (Len(Trim$(CStr(.Value))) > 0)
    End With
    With ws.Cells(ROW_DEPT, 2)
        .Locked = (Len(Trim$(CStr(.Value))) > 0)
    End With

    SetHeaderLabel ws, ROW_DEF_TECH, "Default: technician"
    SetHeaderLabel ws, ROW_DEF_PRINTER, "Default: printer"
    SetHeaderLabel ws, ROW_DEF_PAPER, "Default: paper"

    EnsureLocName ws, "LOC_DefTech", "$B$" & ROW_DEF_TECH
    EnsureLocName ws, "LOC_DefPrinter", "$B$" & ROW_DEF_PRINTER
    EnsureLocName ws, "LOC_DefPaper", "$B$" & ROW_DEF_PAPER

    StyleDefaultCell ws.Cells(ROW_DEF_TECH, 2)
    StyleDefaultCell ws.Cells(ROW_DEF_PRINTER, 2)
    StyleDefaultCell ws.Cells(ROW_DEF_PAPER, 2)

    ' The two disregard-cost defaults (2026-09-29, moved here from the side
    ' panel). Yes/No, seeded onto each new job like the selectors above but
    ' not cleared by Clear defaults; start at No. Labels match the job
    ' table's own "Disregard Paper"/"Disregard Consumable" headers and are
    ' styled exactly like the selector labels.
    SetHeaderLabel ws, ROW_DIS_PAPER, "Disregard paper"
    SetHeaderLabel ws, ROW_DIS_CONS, "Disregard consumable"
    EnsureLocName ws, "LOC_DefDisPaper", "$B$" & ROW_DIS_PAPER
    EnsureLocName ws, "LOC_DefDisCons", "$B$" & ROW_DIS_CONS
    StyleDefaultCell ws.Cells(ROW_DIS_PAPER, 2)
    StyleDefaultCell ws.Cells(ROW_DIS_CONS, 2)
    If Len(Trim$(CStr(ws.Cells(ROW_DIS_PAPER, 2).Value))) = 0 Then ws.Cells(ROW_DIS_PAPER, 2).Value = "No"
    If Len(Trim$(CStr(ws.Cells(ROW_DIS_CONS, 2).Value))) = 0 Then ws.Cells(ROW_DIS_CONS, 2).Value = "No"
    RelockSheet ws

    ' Both directions of spec 1a's filtering apply here too (spec 1b: "these
    ' selectors should implement the same filtering and autofill principles
    ' as the table cells").
    BindDefaultCells ws
End Sub

' Same label look for every row of the header block's left column.
Private Sub SetHeaderLabel(ByVal ws As Worksheet, ByVal Row As Long, ByVal Text As String)
    Dim ref As Range
    Set ref = ws.Cells(ROW_DEF_TECH, 1)
    With ws.Cells(Row, 1)
        .Value = Text
        .Font.Name = ref.Font.Name
        .Font.Size = ref.Font.Size
        .Font.Color = ref.Font.Color
        .Font.Bold = True
        .Font.Italic = False
        .HorizontalAlignment = ref.HorizontalAlignment
        .VerticalAlignment = ref.VerticalAlignment
        .IndentLevel = ref.IndentLevel
        .Interior.Pattern = xlNone
    End With
End Sub

' Input cell with the same border on all four edges (the blue left edge of
' StyleInputCell plus a thin grey frame), so the default cells match.
Private Sub StyleDefaultCell(ByVal target As Range)
    StyleInputCell target
    Dim e As Variant
    For Each e In Array(xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With target.Borders(e)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(166, 166, 166)
        End With
    Next e
End Sub

' Per-location roll-stock entry unit: some print rooms prefer to type a roll
' job's length in centimetres rather than metres. Conceptually part of the
' same "regularly used, must stay visible" group as the three selectors
' above (user feedback, 2026-09-25), so it joins them at A6/B6 rather than
' sitting apart - moved from its original D4/E4 spot in the same layout fix
' described above. Self-provisioned and re-applied on every
' InitialiseWorkbook AND RefreshLocations run, same "re-point a duplicated
' sheet, self-heal an old one" reasoning EnsureJobDefaults' own comment
' gives for LOC_DefTech/LOC_Export - this setting postdates the last
' hand-edit of the .xlsx, so a shipped name (the way LOC_DefDisPaper is
' done) is not an option.
'
' Display unit only (2026-10-01): Qty and the Unit column show roll lengths
' in this unit on THIS sheet, and the Unit/Area m2/Paper Cost formulas (which
' ship in the .xlsx) convert back to metres, as does the consolidated _Data range
' (modRegistry.WriteConsolidated) - so reports, Summary and every other
' location are untouched. Changing the setting rescales the existing roll
' rows (modValidation.ApplyRollUnitChange).
Public Sub EnsureRollUnitSetting(ByVal ws As Worksheet)
    UnlockSheet ws
    SetHeaderLabel ws, ROW_ROLL_UNIT, "Roll length unit"

    EnsureLocName ws, "LOC_RollUnit", "$B$" & ROW_ROLL_UNIT

    Dim c As Range
    Set c = ws.Cells(ROW_ROLL_UNIT, 2)
    If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = "Metres"
    StyleDefaultCell c
    With c.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Metres,Centimetres"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Roll length unit"
        .InputMessage = "The unit roll lengths are shown and typed in on this sheet (Qty and the Unit column). Changing it converts the roll lengths already entered. Reports, Summary and other locations always use metres - Sheet stock is unaffected."
        .ErrorTitle = "Roll length unit"
        .ErrorMessage = "Choose Metres or Centimetres."
    End With
    RelockSheet ws
End Sub

' Qty's validation rule, worded for the room's own roll length unit. Split
' out so a unit change can swap just this rule without re-running the whole
' EnsureJobColumnValidation (which also clears the Printer/Paper Stock lists).
Private Sub AddQtyRule(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim qtyMsg As String
    If StrComp(RollUnitOf(ws), "Centimetres", vbTextCompare) = 0 Then
        qtyMsg = "Sheets for sheet stock, centimetres for roll stock (see 'Roll length unit' above). Must be greater than zero."
    Else
        qtyMsg = "Sheets for sheet stock, metres for roll stock. Must be greater than zero."
    End If
    AddRule lo, "Qty", xlValidateDecimal, "0", False, "Qty", qtyMsg, qtyMsg, xlGreater
End Sub

Public Sub RefreshQtyValidation(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    UnlockSheet ws
    lo.ListColumns("Qty").DataBodyRange.Validation.Delete
    AddQtyRule ws, lo
    RelockSheet ws
End Sub

' "Centimetres" or "Metres" (blank/anything else reads as Metres, the default).
Public Function RollUnitOf(ByVal ws As Worksheet) As String
    If StrComp(LocValue(ws, "LOC_RollUnit"), "Centimetres", vbTextCompare) = 0 Then
        RollUnitOf = "Centimetres"
    Else
        RollUnitOf = "Metres"
    End If
End Function

' LOC_RollUnitApplied: a sheet-scoped constant name recording which unit the
' stored roll Qty values are actually in, so a change to LOC_RollUnit (which
' fires Change even when the same item is re-picked) knows what to convert
' FROM. Absent on a sheet that has never been through this - reads as Metres,
' which is what every sheet before 2026-10-01 stored.
Public Function AppliedRollUnit(ByVal ws As Worksheet) As String
    Dim r As String
    On Error Resume Next
    r = ws.Names("LOC_RollUnitApplied").RefersTo
    On Error GoTo 0
    If InStr(1, r, "Centimetres", vbTextCompare) > 0 Then AppliedRollUnit = "Centimetres" Else AppliedRollUnit = "Metres"
End Function

Public Sub SetAppliedRollUnit(ByVal ws As Worksheet, ByVal Unit As String)
    On Error Resume Next
    ws.Names("LOC_RollUnitApplied").Delete
    On Error GoTo 0
    ws.Names.Add Name:="LOC_RollUnitApplied", RefersTo:="=""" & Unit & """", Visible:=False
End Sub

' Friendly display of the permitted-printers list (2026-09-25, user
' request) - A7/B7, the row Sheet status vacated when it swapped places with
' this (below). LOC_Printers itself - the raw semicolon-delimited list the
' Select printers... picker writes and modCatalog's compatibility checks
' read (modPicker.bas:189, modCatalog.bas:149/168) - stays exactly as it
' always was functionally, just relocated to the side panel (AM7, "Permitted
' printers (raw list)") so the picker's plain .Value write there is
' unaffected; this is a SEPARATE, read-only formula cell, never written to
' directly, that renders it for people rather than code: "N printers
' (comma, separated, list)", singular "1 printer" handled explicitly. A
' genuine Excel formula (LET, like modReports' own summary formulas), not a
' VBA-computed string, so it stays in sync with the raw list with no code
' involved - re-written every run only because LOC_Printers' name could in
' principle be repointed by a future edit, same defensive reasoning as the
' job count below.
Public Sub EnsurePrintersDisplay(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Cells(ROW_PRINTERS, 1).Value = "Printers at this location"
    ws.Cells(ROW_PRINTERS, 1).Font.Bold = True

    ' Built with Chr(34) rather than a hand-escaped string literal - the
    ' quote-doubling needed to embed this many nested string arguments in a
    ' single VBA literal is too easy to get subtly wrong to trust by eye.
    ' Target formula (each q below is one literal "):
    '   =LET(list,LOC_Printers,n,IF(list="",0,LEN(list)-LEN(SUBSTITUTE(list,
    '   ";",""))+1),IF(n=0,"0 printers ()",n&" printer"&IF(n=1,"","s")&
    '   " ("&SUBSTITUTE(list,";",", ")&")"))
    Dim q As String
    q = Chr(34)
    ws.Cells(ROW_PRINTERS, 2).Formula2 = "=LET(list,LOC_Printers,n,IF(list=" & q & q & _
        ",0,LEN(list)-LEN(SUBSTITUTE(list," & q & ";" & q & "," & q & q & _
        "))+1),IF(n=0," & q & "0 printers ()" & q & ",n&" & q & " printer" & _
        q & "&IF(n=1," & q & q & "," & q & "s" & q & ")&" & q & " (" & q & _
        "&SUBSTITUTE(list," & q & ";" & q & "," & q & ", " & q & ")&" & q & _
        ")" & q & "))"
    ws.Cells(ROW_PRINTERS, 2).Locked = True
    RelockSheet ws
End Sub

' Live count of print jobs recorded on this sheet (2026-09-25, user
' request) - A8/B8, the row Export vacated when it moved to the side panel
' (below) - originally A9/B9, the row freed up by no longer needing
' EnsureJobTableGap, until Sheet status/Export/Printers' 2026-09-25 swap
' rearranged rows 7-9 again. A genuine Excel formula, not a VBA-computed
' value, so it stays accurate as rows are added/removed with no code
' involved - ROWS() on a bare table reference counts its data rows,
' excluding the header. Re-written (not just created-once) on every run
' because the table's own name can change (RefreshLocations, AT-13) and a
' stale table name in the formula text would silently stop updating.
Public Sub EnsureJobCountDisplay(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    UnlockSheet ws
    ws.Cells(ROW_JOB_COUNT, 1).Value = "Print jobs"
    ws.Cells(ROW_JOB_COUNT, 1).Font.Bold = True
    ws.Cells(ROW_JOB_COUNT, 1).HorizontalAlignment = xlRight
    ' Jobs, not rows: a multi-pass job's colour passes are rows of the table but
    ' not print jobs of their own.
    ws.Cells(ROW_JOB_COUNT, 2).Formula = "=ROWS(" & lo.Name & ")-COUNTIFS(" & lo.Name & "[Row Type]," & Chr$(34) & ROW_PASS & Chr$(34) & ")"
    ws.Cells(ROW_JOB_COUNT, 2).Locked = True
    RelockSheet ws
End Sub

Private Sub EnsureLocName(ByVal ws As Worksheet, ByVal Nm As String, ByVal Addr As String)
    On Error Resume Next
    ws.Names(Nm).Delete
    On Error GoTo 0
    ws.Names.Add Name:=Nm, RefersTo:="='" & ws.Name & "'!" & Addr
End Sub

' Same white-fill, blue-left-border treatment every unlocked input cell gets
' elsewhere (§11's visual design table; modReports.CritCell is the same
' pattern for the Reports page's own filter cells). Public: modCatalog's
' printer-capacity/supplied-by-student column setup reuses it too.
Public Sub StyleInputCell(ByVal target As Range)
    target.Locked = False
    target.Interior.Color = RGB(255, 255, 255)
    target.Borders(xlEdgeLeft).Color = RGB(46, 100, 168)
    target.Borders(xlEdgeLeft).Weight = xlMedium
End Sub

' NOTE: SET_LOC_REDUCED_COLUMNS must not name a column that shares a sheet
' column with the location config block or batch-defaults rows above the table
' (hiding a column hides it on every row). Status/Job ID sit clear of A/B for that
' reason; a more robust decoupling is still an open idea.

' Re-applies the job table's fixed validation rules (Date/Time, Qty, Print Width,
' Sheet size, Disregard flags, Paid) after clearing the table body, called from
' modLists.BindColumns so setup, Refresh Locations, Check workbook/sheet and the
' picker all restore them. Qty's message depends on the room's LOC_RollUnit, which
' is why this is code rather than only what the .xlsx ships.
'
' Every rule is bound to its column by HEADER NAME, never by letter or position, and the
' table body is cleared before they go on - so reordering, inserting or removing a job-table
' column cannot leave a rule on the wrong column (Excel's Cut+Insert spreads a rule onto the
' cells it shifts, and a positional rule goes stale the moment a column moves). Validation on
' cells BELOW the table is a separate matter: the shipped template has none, and nothing here
' adds any. Rows the table grows into take their validation from the row above, i.e. from here.
Public Sub EnsureJobColumnValidation(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim dateMsg As String
    If lo.DataBodyRange Is Nothing Then Exit Sub

    UnlockSheet ws
    lo.DataBodyRange.Validation.Delete

    dateMsg = "Date and time the print was produced. Use the Now button to stamp the current date and time."
    AddRule lo, "Date/Time", xlValidateDate, "01/01/2000", True, "Date and time", dateMsg, dateMsg, xlGreater

    AddQtyRule ws, lo

    AddRule lo, "Print Width mm", xlValidateDecimal, "0", True, "Print width", _
        "Optional, roll stock only, in millimetres. Leave blank to use the full width of the roll. It must not exceed the stock width. Required for 'Supplied (Roll)'.", _
        "Optional, roll stock only, in millimetres. Leave blank to use the full width of the roll. It must not exceed the stock width.", xlGreater

    AddRule lo, "Sheet size", xlValidateList, "=RNG_STD_SIZES", True, "Sheet size", _
        "Required only for 'Supplied (Sheet)' - the nearest standard size to the sheet the student brought. Ignored for every other paper stock.", _
        "Choose one of the standard sizes listed on Settings."

    AddRule lo, "Disregard Paper", xlValidateList, "Yes,No", True
    AddRule lo, "Disregard Consumable", xlValidateList, "Yes,No", True

    AddRule lo, "Paid", xlValidateList, "Yes,No", True, "Paid", _
        "Whether this chargeable cost has been paid. Blank means not recorded either way and counts as unpaid in totals.", _
        "Choose Yes or No."

    RelockSheet ws
End Sub

' One validation rule on one job-table column. Title is used for both the input and
' error dialogs; Title, Prompt (the input message) and ErrText (the error message) are
' left unset when blank.
Private Sub AddRule(ByVal lo As ListObject, ByVal ColName As String, ByVal RuleType As Long, _
        ByVal Formula1 As String, ByVal IgnoreBlank As Boolean, Optional ByVal Title As String, _
        Optional ByVal Prompt As String, Optional ByVal ErrText As String, _
        Optional ByVal Op As Long = xlBetween)
    With lo.ListColumns(ColName).DataBodyRange.Validation
        .Add Type:=RuleType, AlertStyle:=xlValidAlertStop, Operator:=Op, Formula1:=Formula1
        .IgnoreBlank = IgnoreBlank
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        If Len(Title) > 0 Then
            .InputTitle = Title
            .ErrorTitle = Title
        End If
        If Len(Prompt) > 0 Then .InputMessage = Prompt
        If Len(ErrText) > 0 Then .ErrorMessage = ErrText
    End With
End Sub

' Restates the whole H_Issues calculated column, so an old or duplicated sheet always
' carries the current rules - including the two required-field checks the student-supplied
' stock options need (Print Width mm on "Supplied (Roll)", Sheet size on "Supplied (Sheet)"),
' which have no catalogue size to fall back on. Called from modLists.BindColumns, right after
' EnsureJobColumnValidation, so it runs on every setup, Refresh Locations, Check workbook/sheet
' and picker call, and rewrites the whole column each time (repairing a single damaged row).
'
' There used to be a compare-first guard here. It never matched - the old text was built in
' the long tbl[[#This Row],[Field]] form, which Excel reads back as [@Field] - so it only ever
' looked like a saving. The text below is the form Excel reads back (and does not depend on the
' table's name), so a guard would now work, but would stop repairing a damaged row below the
' first; left out until the rewrite cost is shown to matter. Formula2, not Formula, because the
' text is past the 255 characters Range.Formula accepts (1004).
Public Sub EnsureJobIssuesFormula(ByVal ws As Worksheet, ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    UnlockSheet ws
    lo.ListColumns("H_Issues").DataBodyRange.Formula2 = BuildJobIssuesFormula()
    ' The milder level beside it (multi-pass design decision 35): same stance,
    ' asserted here so an old or duplicated sheet always carries it.
    If ColumnExists(lo, "H_Notices") Then lo.ListColumns("H_Notices").DataBodyRange.Formula2 = BuildJobNoticesFormula()
    RelockSheet ws
End Sub

' Notices are not problems: a pass whose colour this workbook does not define
' (it came in on an import from a workbook that does) is kept and counted, and
' shows yellow with this note instead of orange. Prefixed "Notice:" so the
' Status conditional format can tell the two levels apart from the text alone.
' Kept to a hyphen, not a semicolon, because Status splits its note on "; ".
Private Function BuildJobNoticesFormula() As String
    Const DQ As String = """"
    If Tbl("tblColours") Is Nothing Then
        BuildJobNoticesFormula = "=" & DQ & DQ
        Exit Function
    End If
    BuildJobNoticesFormula = "=IF(AND(" & RowRef("Row Type") & "=" & DQ & ROW_PASS & DQ & "," & RefFilled(RowRef("Colour")) & _
        ",COUNTIFS(tblColours[Colour]," & RowRef("Colour") & ")=0)," & DQ & "; Notice: colour not defined in this workbook - type not checked" & DQ & "," & DQ & DQ & ")"
End Function

' One IF per rule - IF(condition,"; message","") - joined with & in a fixed order, all under
' an outer IF that leaves the cell blank on a row with no Job ID.
Private Function BuildJobIssuesFormula() As String
    Const DQ As String = """"
    Dim qty As String, width As String, stock As String, measure As String
    Dim jobChain As String, passChain As String, colType As String, prnType As String, isMulti As String
    qty = RowRef("Qty")
    width = RowRef("Print Width mm")
    stock = RowRef("Paper Stock")
    measure = RowRef("S_Measure")

    jobChain = _
        IssueIf(RefBlank(RowRef("Date/Time")), "Date and time required") & _
        "&" & IssueIf("AND(" & RefBlank(RowRef("Student Name")) & "," & RefBlank(RowRef("Student No")) & ")", "Student name or number required") & _
        "&" & IssueIf(RefBlank(RowRef("Technician")), "Technician required") & _
        "&" & IssueIf(RefBlank(RowRef("Printer")), "Printer required") & _
        "&" & IssueIf(RefBlank(stock), "Paper stock required") & _
        "&IF(" & RefBlank(qty) & "," & DQ & "; Quantity required" & DQ & "," & IssueIf(qty & "<=0", "Quantity must be greater than zero") & ")" & _
        "&" & IssueIf("AND(" & measure & "=" & DQ & "Sheet" & DQ & "," & RefFilled(qty) & "," & qty & "<>INT(" & qty & "))", "Sheet quantity must be a whole number") & _
        "&" & IssueIf("AND(" & measure & "=" & DQ & "Sheet" & DQ & "," & RefFilled(width) & ")", "Print width does not apply to sheet stock") & _
        "&" & IssueIf("AND(" & RefFilled(width) & "," & width & ">" & RowRef("S_StockWidth_mm") & ")", "Print width exceeds stock width") & _
        "&" & IssueIf("AND(" & stock & "=" & DQ & SUPPLIED_ROLL & DQ & "," & RefBlank(width) & ")", "Print width required for student-supplied roll stock") & _
        "&" & IssueIf("AND(" & stock & "=" & DQ & SUPPLIED_SHEET & DQ & "," & RefBlank(RowRef("Sheet size")) & ")", "Sheet size required for student-supplied sheet stock")

    ' A multi-pass job needs at least one colour pass (decision 10). Reads the
    ' printer's current Colour mode: a status, not a cost, so it may follow the
    ' catalogue.
    isMulti = "XLOOKUP(" & RowRef("Printer") & ",tblPrinters[Model],tblPrinters[Colour mode]," & DQ & DQ & ")=" & DQ & COLOUR_MODE_MULTI & DQ
    jobChain = jobChain & "&" & IssueIf("AND(" & RowRef("Printer") & "<>" & DQ & DQ & "," & isMulti & "," & RefBlank(RowRef("Passes")) & ")", "At least one colour pass is required")

    ' A pass row (decisions 10, 24, 30): a colour, a parent that exists, a number in
    ' sequence, and a colour of the printer's consumable type when both are known.
    passChain = IssueIf(RefBlank(RowRef("Colour")), "Colour required") & _
        "&" & IssueIf("OR(" & RefBlank(RowRef("Parent")) & ",COUNTIFS([Job ID]," & RowRef("Parent") & ")=0)", "Parent job not found") & _
        "&" & IssueIf(RowRef("Pass") & "<>COUNTIFS(INDEX([Parent],1):" & RowRef("Parent") & "," & RowRef("Parent") & ")", "Pass number out of sequence")
    If Not Tbl("tblColours") Is Nothing Then
        colType = "XLOOKUP(" & RowRef("Colour") & ",tblColours[Colour],tblColours[Consumable type]," & DQ & DQ & ")"
        prnType = "XLOOKUP(XLOOKUP(" & RowRef("Parent") & ",[Job ID],[Printer]," & DQ & DQ & "),tblPrinters[Model],tblPrinters[Consumable type]," & DQ & DQ & ")"
        passChain = passChain & "&" & IssueIf("AND(" & colType & "<>" & DQ & DQ & "," & prnType & "<>" & DQ & DQ & "," & colType & "<>" & prnType & ")", _
            "Colour is a different consumable type from the printer's")
    End If

    BuildJobIssuesFormula = "=IF(" & RowRef("Row Type") & "=" & DQ & ROW_PASS & DQ & "," & passChain & "," & _
        "IF(" & RefBlank(RowRef("Job ID")) & "," & DQ & DQ & "," & jobChain & "))"
End Function

Private Function RowRef(ByVal Field As String) As String
    If Field Like "*[!A-Za-z0-9]*" Then RowRef = "[@[" & Field & "]]" Else RowRef = "[@" & Field & "]"
End Function

Private Function RefBlank(ByVal ref As String) As String
    RefBlank = ref & "=" & Chr$(34) & Chr$(34)
End Function

Private Function RefFilled(ByVal ref As String) As String
    RefFilled = ref & "<>" & Chr$(34) & Chr$(34)
End Function

Private Function IssueIf(ByVal Condition As String, ByVal Message As String) As String
    IssueIf = "IF(" & Condition & "," & Chr$(34) & "; " & Message & Chr$(34) & "," & Chr$(34) & Chr$(34) & ")"
End Function

Private Function CountButtons() As Long
    Dim ws As Worksheet, i As Long, t As Long
    For Each ws In ThisWorkbook.Worksheets
        For i = 1 To ws.Buttons.Count
            If Left$(ws.Buttons(i).Name, Len(BTN_TAG)) = BTN_TAG Then t = t + 1
        Next i
    Next ws
    CountButtons = t
End Function

' NameTag, optional: the shape's own identifying Name is BTN_TAG & NameTag
' (or & Macro, when NameTag is blank) & "_" & RowNo & "_" & ColNo, capped at
' 31 characters by Excel/COM with no error at 32+ (SetButtonName's own
' comment; caught before for "btnToggleSettingsSheets" at row/col 5/15).
' Macro itself - what actually runs on click, via .OnAction below - is never
' shortened; NameTag only trims the internal uniqueness label for an
' already-long macro name moved to a two-digit column, without renaming the
' real Sub everywhere else it's referenced (modMain, test scripts).
Public Sub DrawOne(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal ColNo As Long, ByVal Caption As String, ByVal Macro As String, ByVal W As Single, Optional ByVal NameTag As String = "")
    Dim b As Button, c As Range
    Dim tag As String
    tag = Macro
    If Len(NameTag) > 0 Then tag = NameTag
    Set c = ws.Cells(RowNo, ColNo)
    Set b = ws.Buttons.Add(c.Left, c.Top, W, 22)
    ' xlFreeFloating (2026-09-26, user-reported): Buttons.Add defaults to
    ' Placement:=xlMoveAndSize, which silently resizes/repositions the shape
    ' whenever the column(s) beneath it change width - not just an explicit
    ' .Left/.Width assignment, but ANY width change, including
    ' ApplyColumnVisibility's Hidden toggling for the reduced-clutter view.
    ' That is what let the Now button end up straddling the D/E border with
    ' only its "N" visible after one hide-then-reveal cycle: hiding Printer
    ' shrank the button's own Width (not just moved it), and
    ' RelocateAtRiskButtons' RelocateButton only ever restored .Left, so the
    ' shrunk width stuck. Free-floating detaches the shape from the
    ' underlying cells entirely, so it only ever moves when this code moves
    ' it.
    b.Placement = xlFreeFloating
    SetButtonName b, BTN_TAG & tag & "_" & RowNo & "_" & ColNo
    b.Caption = Caption
    b.OnAction = Macro
    b.Characters.Font.Size = 10
End Sub

' Same as DrawOne, but takes an explicit pixel Top rather than a row number -
' for the location side panel (2026-09-25, user-reported: Import had drifted
' off to a second column, away from the rest of the stack). Row height is a
' whole-row property, so every column sharing row 12 with the main A/B
' block is stuck at that row's ~14.5pt height - too short to give a 22pt
' button its usual one-row gap the way every-other-row spacing does
' elsewhere, which is what forced Import out to SIDE_PANEL_COL2 in the first
' place. Pixel-based stacking sidesteps the row grid entirely, so seven
' buttons fit in one column, evenly spaced, with room to spare above the
' table header - see SidePanelButtonLayout for the actual spacing/height
' calculation. H defaults to the standard 22pt but SidePanelButtonLayout can
' hand back something shorter when the current button count needs it
' (2026-09-27) - see that function's own comment for why.
Public Sub DrawOneAtTop(ByVal ws As Worksheet, ByVal ColNo As Long, ByVal Top As Double, ByVal Caption As String, ByVal Macro As String, ByVal W As Single, Optional ByVal H As Double = 22)
    Dim b As Button
    Set b = ws.Buttons.Add(ws.Cells(1, ColNo).Left, Top, W, H)
    b.Placement = xlFreeFloating ' see DrawOne's comment - same column-resize drift risk
    SetButtonName b, BTN_TAG & Macro & "_px" & CLng(Top) & "_" & ColNo
    b.Caption = Caption
    b.OnAction = Macro
    b.Characters.Font.Size = 10
End Sub

' Labels and sizes the Settings button band. The rows themselves ship in the
' .xlsx (tblSettings' header is on row 8); ClearButtons has already removed this
' sheet's buttons and the caller redraws them.
Private Sub EnsureSettingsButtonBand(ByVal ws As Worksheet)
    Dim r As Long, labels As Variant, i As Long
    UnlockSheet ws
    labels = Array("Print rooms", "Data", "Workbook")
    For i = 0 To SETTINGS_BAND_ROWS - 1
        r = SETTINGS_BAND_FIRST_ROW + i
        With ws.Rows(r)
            .ClearFormats
            .RowHeight = SETTINGS_BAND_ROW_H
        End With
        With ws.Cells(r, 1)
            .Value = labels(i)
            .Font.Bold = True
            .Font.Size = 10
            .VerticalAlignment = xlCenter
        End With
    Next i
    RelockSheet ws
End Sub

' One button in the Settings band: Slot is its position (0-based) within the
' row, laid out from BandLeft at a fixed width and gap, vertically centred on
' the row. Name is BTN_TAG & Macro & "_" & row & "_" & slot (<= 31 chars - see
' DrawOne's note on the truncation).
Private Sub DrawBandButton(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal BandLeft As Double, ByVal Slot As Long, ByVal Caption As String, ByVal Macro As String)
    Dim b As Button
    Set b = ws.Buttons.Add(BandLeft + Slot * (SETTINGS_BTN_W + SETTINGS_BTN_GAP), ws.Rows(RowNo).Top + (ws.Rows(RowNo).Height - 22) / 2, SETTINGS_BTN_W, 22)
    b.Placement = xlFreeFloating ' see DrawOne's comment - same column-resize drift risk
    SetButtonName b, BTN_TAG & Macro & "_" & RowNo & "_" & Slot
    b.Caption = Caption
    b.OnAction = Macro
    b.Characters.Font.Size = 10
End Sub

' A compact square button (snag list item 5) - the "+"/"-" row buttons on the
' Settings sheet's four lookup tables, narrow enough to sit above a table
' without needing a row of its own the way the wider text buttons
' on Printers/Papers/Print Technicians do.
Public Sub DrawSmall(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal ColNo As Long, ByVal Caption As String, ByVal Macro As String, ByVal W As Single)
    Dim b As Button, c As Range
    Set c = ws.Cells(RowNo, ColNo)
    Set b = ws.Buttons.Add(c.Left, c.Top, W, 16)
    b.Placement = xlFreeFloating ' see DrawOne's comment - same column-resize drift risk
    SetButtonName b, BTN_TAG & Macro & "_" & RowNo & "_" & ColNo
    b.Caption = Caption
    b.OnAction = Macro
    b.Characters.Font.Size = 10
    b.Characters.Font.Bold = True
End Sub

' Button.Name silently truncates to 31 characters at exactly 32, and raises
' 1004 ("Unable to set the Name property of the Button class") at 33+ -
' verified directly with a length sweep against this Excel/COM automation
' context. Neither limit is documented anywhere found. That is what caused
' this to fire for three of the small Settings-lookup-table buttons before
' their macro names were shortened (DrawSmall's caller comment). The retry
' below stays as a safety net for a genuinely transient COM rejection - the
' same class build.ps1's own SaveAs retries for - but it cannot fix a name
' that is simply too long; that has to be fixed at the call site, and a
' truncation is not even something this retry could detect, let alone fix.
Private Sub SetButtonName(ByVal b As Button, ByVal Nm As String)
    Dim tries As Long
    For tries = 1 To 5
        On Error Resume Next
        Err.Clear
        b.Name = Nm
        If Err.Number = 0 Then Exit Sub
        On Error GoTo 0
        DoEvents
    Next tries
    b.Name = Nm ' final attempt: let a genuine failure raise for real
End Sub

Private Sub ClearButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Buttons.Count To 1 Step -1
        If Left$(ws.Buttons(i).Name, Len(BTN_TAG)) = BTN_TAG Then ws.Buttons(i).Delete
    Next i
    For i = ws.DropDowns.Count To 1 Step -1
        If Left$(ws.DropDowns(i).Name, Len(BTN_TAG)) = BTN_TAG Then ws.DropDowns(i).Delete
    Next i
End Sub

' The All/Reduced/Minimal drop-down (a forms control, so it works on Mac as
' well as Windows, same reason the buttons are forms controls). Unlocked so
' it can be changed while the sheet is protected. Its OnAction is
' modMain.ddViewMode; RepositionViewButtons places it and keeps its
' selection in step with the workbook-wide setting.
Private Sub DrawViewDropDown(ByVal ws As Worksheet)
    Dim d As DropDown, c As Range
    Set c = ws.Cells(VIEW_ROW, VIEW_FIRST_COL + 1)
    Set d = ws.DropDowns.Add(c.Left, c.Top, VIEW_DD_W, 22)
    d.Placement = xlFreeFloating ' see DrawOne's comment - same column-resize drift risk
    d.Name = BTN_TAG & VIEW_DD_NAME
    d.AddItem "All"
    d.AddItem "Reduced"
    d.AddItem "Minimal"
    d.ListIndex = 1
    d.DropDownLines = 3
    d.Locked = False
    d.OnAction = "ddViewMode"
End Sub

Private Function FindDropDown(ByVal ws As Worksheet, ByVal Prefix As String) As DropDown
    Dim i As Long
    For i = 1 To ws.DropDowns.Count
        If Left$(ws.DropDowns(i).Name, Len(Prefix)) = Prefix Then
            Set FindDropDown = ws.DropDowns(i)
            Exit Function
        End If
    Next i
End Function

' modMain.ddViewMode's target. Reads the choice from the drop-down that was
' just used (Application.Caller names it) and applies it workbook-wide. When
' there is no caller (run from code) it uses the active sheet's drop-down.
Public Sub ViewDropDownChanged()
    Dim ws As Worksheet, d As DropDown, nm As String
    Set ws = ActiveSheet
    On Error Resume Next
    nm = CStr(Application.Caller)
    On Error GoTo 0
    If Len(nm) > 0 Then
        On Error Resume Next
        Set d = ws.DropDowns(nm)
        On Error GoTo 0
    End If
    If d Is Nothing Then Set d = FindDropDown(ws, BTN_TAG & VIEW_DD_NAME)
    If d Is Nothing Then Exit Sub
    Select Case d.ListIndex
        Case 2: SetViewMode "Reduced"
        Case 3: SetViewMode "Minimal"
        Case Else: SetViewMode "All"
    End Select
End Sub

' ------------------------------------------------------------- freeze panes ---
' Freeze panes are a per-window view setting with no non-UI object model
' property - Excel only exposes it through the active window - so, like
' ProtectAll re-applying UserInterfaceOnly on every run, it is re-asserted
' here every setup run rather than trusted to whatever the .xlsx last shipped
' with. Anchor is the cell that becomes the new top-left of the scrolling
' area; "" clears any existing freeze instead (snag list item 8, for the
' Settings/Technicians/Printers/Papers sheets - unneeded visual clutter on
' sheets that are just a handful of short tables).
Private Sub SetFreeze(ByVal ws As Worksheet, ByVal Anchor As String)
    Dim prevWs As Worksheet, prevSel As Range

    On Error Resume Next
    Set prevWs = ActiveSheet
    Set prevSel = Selection
    On Error GoTo 0

    ws.Activate
    ActiveWindow.FreezePanes = False
    If Len(Anchor) > 0 Then
        ws.Range(Anchor).Select
        ActiveWindow.FreezePanes = True
    End If

    On Error Resume Next
    If Not prevWs Is Nothing Then
        prevWs.Activate
        If Not prevSel Is Nothing Then prevSel.Select
    End If
    On Error GoTo 0
End Sub

' ------------------------------------------------------------- tab order ---
' Enforces the tab order the snag list settled on (item 12): Summary,
' Reports, every print room, then the four configuration sheets in a fixed
' order. Hidden system sheets (_Data, _Registry, _Audit, _Work, _Picker,
' _Export) are xlSheetVeryHidden and never show in the tab bar, so their
' position is left alone.
Private Sub ReorderSheetTabs()
    Dim after As Worksheet, ws As Worksheet
    Dim locs As Collection, v As Variant, nm As Variant

    On Error Resume Next
    Set after = ThisWorkbook.Worksheets("Reports")
    On Error GoTo 0
    If after Is Nothing Then Exit Sub

    Set locs = LocationSheets()
    For Each v In locs
        Set ws = v
        ws.Move After:=after
        Set after = ws
    Next v

    For Each nm In ConfigSheetNames()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Move After:=after
            Set after = ws
        End If
    Next nm
End Sub

' -------------------------------------------------------- config sheets ---
' The four sheets snag list items 6, 8, 12 and 13 all refer to by name.
' "Print Technicians" is the sheet's real name (ThisWorkbook.cls Case list) -
' the snag list's "Technicians" is shorthand for it.
Private Function ConfigSheetNames() As Variant
    ConfigSheetNames = Array("Print Technicians", "Printers", "Papers", "Consumables", "Departments", "Settings")
End Function

' Config sheets ship in the .xlsx with every cell at Excel's default Locked
' state, so ProtectAll's blanket Contents:=True previously left every table
' on them uneditable (snag list item 6, corrected from an earlier plan to
' unprotect the sheets outright: they stay protected, only the cells users
' fill in unlock). Unlocked here rather than hand-edited into the .xlsx, for
' the same reproducibility reason EnsureSystemSheets gives for building
' structure in VBA rather than by hand.
Private Sub UnlockConfigInputs()
    UnlockTableBody "tblTechnicians"
    UnlockTableBody "tblPrinters"
    UnlockTableBody "tblPapers"
    UnlockTableBody "tblDepartments"
    UnlockTableBody "tblColours"
    UnlockTableBody "tblPaperTypes"
    UnlockTableBody "tblStandardSizes"
    UnlockTableBody "tblConsumables"
    UnlockSettingsValues
End Sub

Private Sub UnlockTableBody(ByVal TableName As String)
    Dim lo As ListObject
    Dim idHdr As String, idCode As String, nameHdr As String, hwmKey As String
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    UnlockSheet lo.Parent
    lo.DataBodyRange.Locked = False
    ' Technicians, printers and papers have a managed ID column (modCatalog
    ' "catalogue IDs"): VBA writes it, nobody types it, so it stays locked
    ' while the rest of the row is open. A row added later inherits this
    ' from the row above it, and Setup re-applies it on every run.
    If CatalogIdSpec(TableName, idHdr, idCode, nameHdr, hwmKey) Then
        lo.ListColumns(idHdr).DataBodyRange.Locked = True
    End If
    ' tblDepartments' "Match list" is a calculated column the Department
    ' lookup on _Data reads; nobody types into it.
    If StrComp(TableName, "tblDepartments", vbTextCompare) = 0 Then
        If ColumnExists(lo, "Match list") Then lo.ListColumns("Match list").DataBodyRange.Locked = True
    End If
    RelockSheet lo.Parent
End Sub

' tblSettings mixes user-editable settings with system-derived rows (schema
' version, last-refresh stamp, build stamp) whose Notes column is written as
' "Read-only. ..." (in PrintCosts.xlsx) - that text is the one place
' the two kinds are already told apart, so it drives which Value cells unlock
' rather than a second hard-coded list of keys that could drift from it.
' The same loop styles the Value cell so the two kinds look different at a
' glance: an editable value is a white, bordered, blue-text input box; a locked
' one is flat grey italic with no border. Styling it here, from the same test
' that sets .Locked, means the look can never disagree with what is actually
' editable.
Private Sub UnlockSettingsValues()
    Dim lo As ListObject, i As Long, notes As String, v As Range, isLocked As Boolean
    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    UnlockSheet lo.Parent
    For i = 1 To lo.ListRows.Count
        notes = Trim$(CStr(CellIn(lo, i, "Notes").Value))
        isLocked = (Left$(notes, 9) = "Read-only")
        Set v = CellIn(lo, i, "Value")
        v.Locked = isLocked
        StyleSettingValue v, isLocked
    Next i
    RelockSheet lo.Parent
End Sub

Private Sub StyleSettingValue(ByVal v As Range, ByVal IsLocked As Boolean)
    Dim e As Variant
    If IsLocked Then
        v.Interior.Color = RGB(222, 222, 222)
        v.Font.Color = RGB(110, 110, 110)
        v.Font.Italic = True
        For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            v.Borders(e).LineStyle = xlNone
        Next e
    Else
        v.Interior.Color = RGB(255, 255, 255)
        v.Font.Color = RGB(0, 51, 153)
        v.Font.Italic = False
        For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            With v.Borders(e)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(91, 155, 213)
            End With
        Next e
    End If
End Sub

' Snag list item 16: long notes were clipped to one line. WrapText plus a row
' AutoFit is enough - the table is narrow and short, so nothing here needs a
' fixed row height that would break at a different zoom or font.
Private Sub FormatSettingsNotes()
    Dim lo As ListObject
    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    UnlockSheet lo.Parent
    With lo.ListColumns("Notes").DataBodyRange
        .WrapText = True
        .EntireRow.AutoFit
    End With
    RelockSheet lo.Parent
End Sub

' --------------------------------------------------- hide/show settings ---
' Snag list item 13: a Summary-page toggle for the four configuration sheets,
' xlSheetHidden rather than VeryHidden so a determined user can still reach
' Unhide by hand - the button is the friendly path, not the only one.
' Snag 3a (2026-09-22): the button lives on Summary, so Summary is normally
' already the active sheet when this runs - but hiding a sheet that HAPPENS
' to be active forces Excel to activate whatever is next in tab order, and
' nothing here guaranteed the active sheet was Summary rather than one of
' the four being hidden (called other than by a direct click on Summary's
' own button - e.g. Application.Run - can leave any sheet active). Capturing
' Summary explicitly and re-activating it unconditionally at the end is
' correct either way: a no-op when Summary was already active, and the fix
' when it was not.
Public Sub ToggleConfigSheets()
    Dim hideThem As Boolean, nm As Variant, ws As Worksheet, summaryWs As Worksheet
    hideThem = Not ConfigSheetsHidden()

    On Error Resume Next
    Set summaryWs = ThisWorkbook.Worksheets("Summary")
    On Error GoTo 0

    For Each nm In ConfigSheetNames()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Visible = IIf(hideThem, xlSheetHidden, xlSheetVisible)
        End If
    Next nm

    If Not summaryWs Is Nothing Then summaryWs.Activate

    RelabelConfigToggleButton
End Sub

' Read from the first configuration sheet found rather than requiring all
' four to agree, so a sheet renamed or deleted by hand cannot make the button
' appear stuck.
Private Function ConfigSheetsHidden() As Boolean
    Dim nm As Variant, ws As Worksheet
    For Each nm In ConfigSheetNames()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ConfigSheetsHidden = (ws.Visible = xlSheetHidden)
            Exit Function
        End If
    Next nm
End Function

Private Function ConfigToggleCaption() As String
    ConfigToggleCaption = IIf(ConfigSheetsHidden(), "Show settings sheets", "Hide settings sheets")
End Function

' Updates the caption on the Summary sheet's toggle button without a full
' InitialiseWorkbook rerun, so ToggleConfigSheets can flip label and
' visibility together.
Private Sub RelabelConfigToggleButton()
    Dim ws As Worksheet, i As Long, prefix As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Summary")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    prefix = BTN_TAG & "btnToggleConfigSheets"
    For i = 1 To ws.Buttons.Count
        If Left$(ws.Buttons(i).Name, Len(prefix)) = prefix Then
            ws.Buttons(i).Caption = ConfigToggleCaption()
            Exit For
        End If
    Next i
End Sub

' ------------------------------------------------------ reduced-clutter ---
' Snag list item 1e: a workbook-wide toggle (not per-sheet - simpler, and it
' means every location sheet stays in the same state as every other rather
' than risking drift) that hides a short list of columns most day-to-day
' entry doesn't need. The list itself lives in a SETTING rather than a VBA
' constant (REDUCED_COLUMNS_DEFAULT, declared with this module's other
' constants at the top), so it can be edited without a rebuild if the
' shortlist changes later - the snag list's own "keep it flexible".
' View mode (2026-10-01, replaces the single Reduce clutter toggle): "All",
' "Reduced" or "Minimal", held in SET_LOC_REDUCED_VIEW (anything
' unrecognised reads as All). Reduced hides SET_LOC_REDUCED_COLUMNS;
' Minimal hides those AND SET_LOC_MINIMAL_COLUMNS, so it always includes the
' Reduced list.
Private Function ViewMode() As String
    Select Case LCase$(SettingText("LOC_REDUCED_VIEW", "All"))
        Case "reduced": ViewMode = "Reduced"
        Case "minimal": ViewMode = "Minimal"
        Case Else: ViewMode = "All"
    End Select
End Function

Private Function ReducedColumnsText() As String
    ReducedColumnsText = SettingText("LOC_REDUCED_COLUMNS", REDUCED_COLUMNS_DEFAULT)
End Function

Private Function MinimalExtraText() As String
    MinimalExtraText = SettingText("LOC_MINIMAL_COLUMNS", MINIMAL_EXTRA_DEFAULT)
End Function

' Applies the CURRENT view mode to one location sheet's table - called on
' every InitialiseWorkbook/RefreshLocations run (so a freshly duplicated
' sheet, or one predating the feature, always ends up in sync) and again
' from SetViewMode for every location sheet at once. Only the columns either
' view can hide are touched - never a blanket reset (see ApplyColumnVisibility).
Public Sub ApplyReducedView(ByVal ws As Worksheet)
    Dim lo As ListObject, hideText As String, showText As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    Select Case ViewMode()
        Case "Reduced"
            hideText = ReducedColumnsText()
            showText = MinimalExtraText()
        Case "Minimal"
            hideText = ReducedColumnsText() & LIST_SEP & MinimalExtraText()
        Case Else
            showText = ReducedColumnsText() & LIST_SEP & MinimalExtraText()
    End Select
    ' Show first, hide second: a column named in both lists ends up hidden.
    ApplyColumnVisibility lo, SplitList(showText), False
    ApplyColumnVisibility lo, SplitList(hideText), True
    ' The pass columns follow their own toggle and nothing else: the view neither
    ' hides nor shows them, so this runs last and has the final word.
    ApplyMultiPassVisibility ws
    RelocateAtRiskButtons ws, lo
    ' User-reported, 2026-09-26: hiding columns left of SIDE_PANEL_COL shifts
    ' that column's pixel position left, but the side panel's buttons are
    ' Placement:=xlFreeFloating and no longer follow, so they ended up sitting
    ' right of column 36 rather than on it. SetViewMode reaches this Sub
    ' directly, without InitialiseWorkbook's separate RepositionLocationButtons
    ' call, so the side panel needs re-settling here too - on every call.
    RepositionSidePanelButtons ws
End Sub

' The view drop-down's target (ViewDropDownChanged). Stores
' the mode once, then re-applies it to every location sheet in one pass so all
' of them change together.
Public Sub SetViewMode(ByVal Mode As String)
    Dim ws As Worksheet
    SetSetting "LOC_REDUCED_VIEW", Mode

    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then ApplyReducedView ws
    Next ws
    AppOn
End Sub

' --------------------------------------------------------- cost columns ---
' Replaces the native-outline cost-columns group (docs/ARCHITECTURE.md
' §16.3, "button lag" writeup, 2026-09-27). A workbook-wide toggle, same
' shape as reduced-clutter view just above (simpler than per-sheet, and
' keeps every location in step) - "Yes" hides Paper Cost, Consumable Cost,
' Gross Cost and Disregarded on every location sheet; Chargeable Cost and
' Paid are never hidden by this, same columns the old outline group already
' kept outside it (2026-09-22 snag list item 1d).
Private Function CostColumnsHidden() As Boolean
    CostColumnsHidden = (StrComp(SettingText("COST_COLS_HIDDEN", "No"), "Yes", vbTextCompare) = 0)
End Function

Private Function CostColumnsCaption() As String
    CostColumnsCaption = IIf(CostColumnsHidden(), "Show cost detail", "Hide cost detail")
End Function

' Applies the CURRENT setting to one location sheet's table - called on
' every InitialiseWorkbook/RefreshLocations run (so a freshly duplicated
' sheet, or one predating the feature, always ends up in sync) and again
' from ToggleCostColumns for every location sheet at once. Repositions
' buttons in the same call for exactly the reason this replaced the native
' outline control in the first place - see this section's own opening
' comment - rather than relying on Workbook_SheetActivate/
' Workbook_SheetSelectionChange to catch up on the next click.
Public Sub ApplyCostColumnsVisibility(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    ApplyColumnVisibility lo, Array("Paper Cost", "Consumable Cost", "Set-up Cost", "Gross Cost", "Disregarded"), CostColumnsHidden()
    RelocateAtRiskButtons ws, lo
    RepositionSidePanelButtons ws
End Sub

' modMain.btnToggleCostColumns' target. Flips the setting once, then
' re-applies it to every location sheet and relabels every toggle button in
' one pass, same shape as SetViewMode.
Public Sub ToggleCostColumns()
    Dim hideIt As Boolean, ws As Worksheet
    hideIt = Not CostColumnsHidden()
    SetSetting "COST_COLS_HIDDEN", IIf(hideIt, "Yes", "No")

    ' ApplyCostColumnsVisibility re-places the row-2 controls, which also
    ' relabels the toggle (RepositionViewButtons).
    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then ApplyCostColumnsVisibility ws
    Next ws
    AppOn
End Sub

' General fix for the column-hide-takes-a-button-with-it gap (§4.1) - only
' the handful of buttons DrawLocationButtons still anchors inside the job
' table's own column span (everything else lives in the side panel, immune
' by construction) need this. Called every time column visibility can have
' changed - InitialiseWorkbook, RefreshLocations and SetViewMode all
' reach it via ApplyReducedView - so a button self-heals back onto a visible
' column whichever way the setting just moved, including back to its own
' preferred column once reduced view is switched off again.
Private Sub RelocateAtRiskButtons(ByVal ws As Worksheet, ByVal lo As ListObject)
    RelocateButton ws, lo, "btnAddPrintJob", "Date/Time"
    RelocateButton ws, lo, "btnRepeatJob", "Student No"
    RelocateButton ws, lo, "btnNow", "Printer"
    RelocateButton ws, lo, "btnClearDefaults", "Technician"
    RelocateButton ws, lo, "btnAddPass", "Date/Time"
    RelocateButton ws, lo, "btnRemovePass", "Student Name"
    RelocateButton ws, lo, "btnTogglePasses", "Technician"
    RelocateButton ws, lo, "btnToggleAllPasses", "Printer"
    RepositionViewButtons ws, lo
End Sub

' Lays the "Show columns:" label, the All/Reduced/Minimal drop-down and the
' Hide/Show cost detail button out left to right on row VIEW_ROW, starting at
' the first visible column from VIEW_FIRST_COL. Each item takes the first
' usable (visible, not H_Issues or a snapshot column) table column whose left
' edge clears the previous item, so when a column under one is hidden it slides
' to the next visible one instead of vanishing or piling onto its neighbour.
' The label is cell text, so it is moved by rewriting the cell. The drop-down's
' selection and the toggle's caption are re-read from the settings every time.
Private Sub RepositionViewButtons(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim col As Long, minLeft As Double, last As Long
    Dim b As Button, d As DropDown, rowH As Double

    last = lo.Range.Column + lo.ListColumns.Count - 1
    If last >= SIDE_PANEL_COL Then last = SIDE_PANEL_COL - 1
    rowH = ws.Cells(VIEW_ROW, 1).Height - 1
    If rowH < 12 Then rowH = 12

    col = NextViewColumn(ws, lo, last, -1)
    If col = 0 Then Exit Sub
    UnlockSheet ws
    ws.Range(ws.Cells(VIEW_ROW, VIEW_FIRST_COL), ws.Cells(VIEW_ROW, SIDE_PANEL_COL - 1)).ClearContents
    With ws.Cells(VIEW_ROW, col)
        .Value = "Show columns:"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    RelockSheet ws
    minLeft = ws.Cells(1, col).Left + VIEW_LABEL_W + VIEW_BTN_GAP

    Set d = FindDropDown(ws, BTN_TAG & VIEW_DD_NAME)
    col = NextViewColumn(ws, lo, last, minLeft)
    If Not d Is Nothing And col > 0 Then
        d.Left = ws.Cells(1, col).Left
        d.Top = ws.Cells(VIEW_ROW, 1).Top + 0.5
        d.Width = VIEW_DD_W
        d.Height = rowH
        d.Placement = xlFreeFloating
        Select Case ViewMode()
            Case "Reduced": d.ListIndex = 2
            Case "Minimal": d.ListIndex = 3
            Case Else: d.ListIndex = 1
        End Select
        minLeft = d.Left + VIEW_DD_W + VIEW_BTN_GAP
    End If

    Set b = FindButton(ws, BTN_TAG & "btnToggleCostColumns")
    col = NextViewColumn(ws, lo, last, minLeft)
    If Not b Is Nothing And col > 0 Then
        b.Left = ws.Cells(1, col).Left
        b.Top = ws.Cells(VIEW_ROW, 1).Top + 0.5
        b.Width = VIEW_COST_W
        b.Height = rowH
        b.Placement = xlFreeFloating
        b.Caption = CostColumnsCaption()
        minLeft = b.Left + VIEW_COST_W + VIEW_BTN_GAP
    End If

    ' The multi-pass column toggle sits after the cost-detail one.
    Set b = FindButton(ws, BTN_TAG & "btnToggleMultiPass")
    col = NextViewColumn(ws, lo, last, minLeft)
    If Not b Is Nothing And col > 0 Then
        b.Left = ws.Cells(1, col).Left
        b.Top = ws.Cells(VIEW_ROW, 1).Top + 0.5
        b.Width = VIEW_COST_W
        b.Height = rowH
        b.Placement = xlFreeFloating
        b.Caption = MultiPassColsCaption()
    End If
End Sub

' First table column from VIEW_FIRST_COL to ToCol that is safe to anchor on
' (visible, not H_Issues or a snapshot column) and whose left edge is at or
' past MinLeft. Returns 0 if there is none.
Private Function NextViewColumn(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal ToCol As Long, ByVal MinLeft As Double) As Long
    Dim c As Long
    For c = VIEW_FIRST_COL To ToCol
        If c >= lo.Range.Column Then
            If IsSafeAnchorColumn(lo, c - lo.Range.Column + 1) Then
                If ws.Cells(1, c).Left >= MinLeft - 0.5 Then
                    NextViewColumn = c
                    Exit Function
                End If
            End If
        End If
    Next c
End Function

Private Function FindButton(ByVal ws As Worksheet, ByVal Prefix As String) As Button
    Dim i As Long
    For i = 1 To ws.Buttons.Count
        If Left$(ws.Buttons(i).Name, Len(Prefix)) = Prefix Then
            Set FindButton = ws.Buttons(i)
            Exit Function
        End If
    Next i
End Function

Private Sub RelocateButton(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal Macro As String, ByVal PreferredHeader As String)
    Dim b As Button, prefix As String, i As Long, col As Long
    prefix = BTN_TAG & Macro
    For i = 1 To ws.Buttons.Count
        If Left$(ws.Buttons(i).Name, Len(prefix)) = prefix Then
            Set b = ws.Buttons(i)
            Exit For
        End If
    Next i
    If b Is Nothing Then Exit Sub

    col = NearestVisibleColumn(lo, PreferredHeader)
    If col = 0 Then Exit Sub
    b.Left = ws.Cells(1, col).Left
End Sub

' Re-settles a location sheet's buttons once column widths and visibility are
' final. DrawLocationButtons positions every button from the widths at the moment
' it runs, and the buttons are xlFreeFloating (so hide/reveal cycles cannot
' corrupt their size) - so they do not follow later width changes by themselves.
' Called after the per-sheet setup loop's width-changing steps, after the reduced
' view / cost-columns toggles, and from ThisWorkbook on sheet activate and
' selection change.
Public Sub RepositionLocationButtons(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If Not lo Is Nothing Then RelocateAtRiskButtons ws, lo
    RepositionSidePanelButtons ws
    ' Add pass / Remove pass grey with the selection (modPasses).
    RefreshPassButtons ws
End Sub

' The eight occasional-use buttons DrawLocationButtons stacks in
' SIDE_PANEL_COL - see RepositionLocationButtons' comment for why they can
' end up sitting to the right of that column instead of on it.
Private Sub RepositionSidePanelButtons(ByVal ws As Worksheet)
    Dim macros As Variant, i As Long, j As Long, b As Button, prefix As String
    Dim targetLeft As Double, lo As ListObject, btnH As Double, stepPx As Double
    macros = Array("btnSelectPrinters", "btnCheckSheet", "btnRemoveRow", "btnClearAll", "btnExport", "btnImportLocation")
    targetLeft = ws.Cells(1, SIDE_PANEL_COL).Left

    ' Must use the SAME height DrawLocationButtons drew these at (2026-09-27)
    ' - not a second hardcoded 22, which ran every button back to full size
    ' on the very next InitialiseWorkbook/RefreshLocations/toggle call and
    ' silently reopened the overlap SidePanelButtonLayout exists to prevent.
    btnH = 22
    Set lo = JobsTable(ws)
    If Not lo Is Nothing Then SidePanelButtonLayout ws.Cells(lo.Range.Row, 1).Top, btnH, stepPx

    For i = LBound(macros) To UBound(macros)
        prefix = BTN_TAG & CStr(macros(i))
        Set b = Nothing
        For j = 1 To ws.Buttons.Count
            If Left$(ws.Buttons(j).Name, Len(prefix)) = prefix Then
                Set b = ws.Buttons(j)
                Exit For
            End If
        Next j
        If Not b Is Nothing Then
            b.Left = targetLeft
            ' Top too: with the reduced-view toggle gone (2026-10-01) the stack
            ' closes up, and a sheet drawn before that must close up with it.
            b.Top = i * stepPx
            b.Width = 140
            b.Height = btnH
            b.Placement = xlFreeFloating
        End If
    Next i
End Sub

' The button's own preferred column if it's currently visible, else the
' nearest column (checked right then left, one step further out each pass)
' that is neither hidden by the reduced-clutter view nor one of the table's
' permanently-hidden columns (H_Issues, the snapshot block) - landing a
' relocated button on one of those would just move the problem rather than
' solve it. Returns a worksheet column number, not a ListColumn index -
' correct even though they coincide today (the table starts at column A).
Private Function NearestVisibleColumn(ByVal lo As ListObject, ByVal PreferredHeader As String) As Long
    Dim preferred As Long, dist As Long, side As Long, candidate As Long
    On Error Resume Next
    preferred = ColIdx(lo, PreferredHeader)
    On Error GoTo 0
    If preferred = 0 Then Exit Function

    If IsSafeAnchorColumn(lo, preferred) Then
        NearestVisibleColumn = lo.ListColumns(preferred).Range.Column
        Exit Function
    End If

    For dist = 1 To lo.ListColumns.Count
        For side = 1 To -1 Step -2
            candidate = preferred + side * dist
            If candidate >= 1 And candidate <= lo.ListColumns.Count Then
                If IsSafeAnchorColumn(lo, candidate) Then
                    NearestVisibleColumn = lo.ListColumns(candidate).Range.Column
                    Exit Function
                End If
            End If
        Next side
    Next dist
End Function

Private Function IsSafeAnchorColumn(ByVal lo As ListObject, ByVal col As Long) As Boolean
    Dim header As String
    header = lo.ListColumns(col).Name
    If Left$(header, 2) = "H_" Then Exit Function      ' hidden working columns (H_Issues, H_Ink ...)
    If Left$(header, 2) = "S_" Then Exit Function
    IsSafeAnchorColumn = Not CBool(lo.ListColumns(col).Range.EntireColumn.Hidden)
End Function



' Hides or shows exactly the named columns, by header, and touches nothing
' else - some columns (H_Issues) are hidden permanently, shipped that way in
' the .xlsx and never toggled by this or any other code, so this must never
' do a blanket "show everything then hide the list" reset. A header not
' currently on the table (a typo in the setting, say) is skipped rather than
' raising, since ColIdx would otherwise abort the whole pass over one bad
' name. Public: modReports.BuildReports reuses this for the Reports page's
' own fixed minimum-columns view (snag list item 2d).
Public Sub ApplyColumnVisibility(ByVal lo As ListObject, ByVal Headers As Variant, ByVal Hide As Boolean)
    Dim ws As Worksheet, i As Long, col As Long
    Set ws = lo.Parent
    UnlockSheet ws
    For i = LBound(Headers) To UBound(Headers)
        col = 0
        On Error Resume Next
        col = ColIdx(lo, Trim$(CStr(Headers(i))))
        On Error GoTo 0
        If col > 0 Then lo.ListColumns(col).Range.EntireColumn.Hidden = Hide
    Next i
    RelockSheet ws
End Sub

' ------------------------------------------------------- status colour ---
' Phase 8's "warning state" (design doc §11 / architecture §5): the Status
' column is a formula ("OK" or a semicolon-joined issue list from H_Issues,
' behind the scenes) - there is no Worksheet_Change-style hook that fires
' when a formula's result changes, so real conditional formatting is the
' only mechanism that can colour it live as a row goes bad or gets fixed.
'
' Added to the table's DataBodyRange rather than a fixed row range, so Excel's
' normal table auto-extend carries the rule onto rows ListRows.Add creates
' later - the same mechanic AddPrintJob already relies on for the Status
' formula itself.
'
' Re-run safe: unlike BuildSummary (which clears the whole sheet before
' redrawing), this loop never wipes the location sheets, so a second
' InitialiseWorkbook run must delete the rule it drew last time before
' re-adding it - otherwise every re-run stacks another identical one.
Private Sub ApplyStatusFormat(ByVal ws As Worksheet)
    Dim lo As ListObject, rng As Range, fc As FormatCondition, stCol As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = lo.ListColumns("Status").DataBodyRange

    UnlockSheet ws
    rng.FormatConditions.Delete
    ' Two levels (multi-pass design decision 35): a problem is orange, a notice
    ' (Status text starting "Notice:") is a milder yellow. Written against the
    ' whole Status column with INDEX(..., ROW()) so neither rule depends on which
    ' cell is active when it is added.
    stCol = rng.Cells(1, 1).EntireColumn.Address(True, True)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(INDEX(" & stCol & ",ROW())<>"""",INDEX(" & stCol & ",ROW())<>""OK"",LEFT(INDEX(" & stCol & ",ROW()),7)<>""Notice:"")")
    fc.Interior.Color = RGB(255, 192, 0)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=LEFT(INDEX(" & stCol & ",ROW()),7)=""Notice:""")
    fc.Interior.Color = RGB(255, 242, 153)
    ApplyPassFormat ws
    RefreshStatusNotes ws
    RelockSheet ws
End Sub

' The Status column is narrow (0.10.9), so a row's full message - possibly
' several issues joined by "; " - lives in a hover note on the cell instead.
' Excel has no dynamic cell tooltip, and a note's text is static, so this
' re-syncs every note with its cell's current Status: adds one where a row
' has a problem, rewrites it when the message has changed, and deletes it
' when the row is OK or blank (so ordinary rows carry no red marker).
' Diff-only, so a run over an unchanged sheet touches nothing; the scan is a
' single array read. Called by ApplyStatusFormat (setup/Refresh Locations),
' Workbook_SheetChange and Workbook_SheetSelectionChange - the last catches
' status changes that come from anywhere else (added rows, imports, catalogue
' edits) the next time the person clicks. Public: ThisWorkbook calls it.
Public Sub RefreshStatusNotes(ByVal ws As Worksheet)
    Dim lo As ListObject, rng As Range, v As Variant, i As Long
    Dim txt As String, want As String, c As Range, cm As Comment
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Status") Then Exit Sub
    Set rng = lo.ListColumns("Status").DataBodyRange

    ReDim v(1 To rng.Rows.Count, 1 To 1)
    If rng.Rows.Count = 1 Then
        v(1, 1) = rng.Cells(1, 1).Value2
    Else
        v = rng.Value2
    End If

    On Error Resume Next
    For i = 1 To UBound(v, 1)
        txt = ""
        If Not IsError(v(i, 1)) Then txt = CStr(v(i, 1))
        want = ""
        If Len(txt) > 0 And StrComp(txt, "OK", vbBinaryCompare) <> 0 Then
            want = ChrW(8226) & " " & Replace(txt, "; ", vbLf & ChrW(8226) & " ")
        End If

        Set c = rng.Cells(i, 1)
        Set cm = Nothing
        Set cm = c.Comment
        If Len(want) = 0 Then
            If Not cm Is Nothing Then c.ClearComments
        ElseIf cm Is Nothing Then
            c.AddComment want
            SizeStatusNote c.Comment, want
        ElseIf cm.Text <> want Then
            cm.Text want
            SizeStatusNote cm, want
        End If
    Next i
    On Error GoTo 0
End Sub

' Fixed width, height from an estimate of the wrapped line count (about 50
' characters to a line at 300pt), so the whole message shows without the
' note needing to be opened and resized by hand.
Private Sub SizeStatusNote(ByVal cm As Comment, ByVal Text As String)
    Dim parts() As String, i As Long, lines As Long
    parts = Split(Text, vbLf)
    For i = LBound(parts) To UBound(parts)
        lines = lines + 1 + (Len(parts(i)) - 1) \ 50
    Next i
    With cm.Shape
        .TextFrame.AutoSize = False
        .Width = 300
        .Height = 14 * lines + 10
    End With
End Sub
