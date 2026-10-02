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
Private Const SIDE_PANEL_COL As Long = 36

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
Private Const GAP_ROW_HEIGHT As Double = 6
' Header block rows on a location sheet (see EnsureHeaderGaps for the gap rows).
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
' What shipped before the Reduced/Minimal split (2026-10-01): the old single list.
' EnsureViewSettings swaps it for the new default if it finds it untouched.
Private Const REDUCED_COLUMNS_LEGACY As String = "Status;Job ID;Printer;Area m2;Disregard Paper;Disregard Consumable;S_SchemaVer"
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

    ' The version rows likewise. Setup does not stamp the build date - that
    ' is build.ps1's job, via StampBuild.
    EnsureVersionSettings
    EnsureSchemaSetting
    EnsureCatalogIdSettings
    EnsureViewSettings
    EnsureStdSizesName
    RemoveLegacySuppliedRows
    NormaliseSuppliedFlags
    EnsureCatalogIds
    EnsureSuppliedColumnValidation

    ' Moved here from after the per-sheet loop below (2026-09-27,
    ' user-reported: Settings-sheet buttons rendering over the top of
    ' tblSettings). FormatSettingsNotes word-wraps and AutoFits tblSettings'
    ' Notes column - which can grow several of its rows considerably taller
    ' than the sheet's default - but the Settings-sheet buttons a few lines
    ' into that loop (Refresh Locations etc.) are positioned from
    ' TablesBottom(ws), read at THAT moment. Every EnsureXSetting call that
    ' can add a row to tblSettings has already run by this point, so its
    ' final row count - and hence its final row heights, once this runs - are
    ' both settled before the loop ever measures TablesBottom for Settings.
    ' Previously this went unnoticed because Buttons.Add's default
    ' Placement:=xlMoveAndSize silently followed the table's growth spurt
    ' when this ran afterward, same as every other button-drift bug fixed
    ' this session; DrawOne now sets Placement:=xlFreeFloating, so a button
    ' positioned before a later row-height change no longer tracks it.
    FormatSettingsNotes

    ' Summary and Reports, likewise built here. Their formulas read
    ' _Data, which RefreshLocations writes at the end of this run - until then
    ' they sit on their IFERROR fallbacks rather than showing errors.
    BuildReportSheets

    For Each ws In ThisWorkbook.Worksheets
        UnlockSheet ws
        ClearButtons ws
        If IsLocation(ws) Then
            EnsureHeaderGaps ws
            DrawLocationButtons ws
            ConfigValidation ws
            EnsureJobDefaults ws
            EnsureRollUnitSetting ws
            EnsureRollUnitFormulas ws
            EnsurePrintersDisplay ws
            EnsureJobCountDisplay ws
            ClearBelowTableValidation ws
            BindColumns ws
            ApplyStatusFormat ws
            ApplyReducedView ws
            ApplyCostColumnsVisibility ws
            ApplyJobColumnWidths ws
            ' Must run after every call above that can change column widths -
            ' see RepositionLocationButtons' own comment.
            RepositionLocationButtons ws
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
        ElseIf StrComp(ws.Name, "Settings", vbTextCompare) = 0 Then
            ' The eight action buttons used to run down the right-hand side of
            ' the sheet (column T), out of the way but also out of sight on a
            ' normal-width window. They now sit in the space the removed About
            ' block (modVersion.WriteAbout, retired) used to occupy, directly
            ' below whichever table on this sheet runs deepest - tblSettings
            ' once EnsureSetting has added its self-provisioned rows, in every
            ' build seen so far. TablesBottom() finds that row the same way
            ' WriteAbout used to. The old About text lived in columns A:D at
            ' bottom+1 upward; clearing that band first means a workbook built
            ' before this change loses the stale text on its next setup run,
            ' not just on a from-scratch rebuild.
            Dim settingsBottom As Long
            settingsBottom = TablesBottom(ws)
            UnlockSheet ws
            ws.Range(ws.Cells(settingsBottom + 1, 1), ws.Cells(settingsBottom + 20, 4)).Clear
            RelockSheet ws

            ' Two logical groups, four buttons each, one row per group so both
            ' read at a glance: everyday workbook actions first, then the
            ' data-movement actions (export/import/backup/restore). Columns
            ' three apart (roughly 190px at the sheet's default column width)
            ' so a 130px-wide button never crowds its neighbour.
            DrawOne ws, settingsBottom + 2, 1, "Refresh Locations", "btnRefreshLocations", 130
            DrawOne ws, settingsBottom + 2, 4, "Check workbook", "btnCheckWorkbook", 130
            DrawOne ws, settingsBottom + 2, 7, "Re-stamp prices...", "btnReStamp", 130
            DrawOne ws, settingsBottom + 2, 10, "About", "btnAbout", 130

            DrawOne ws, settingsBottom + 4, 1, "Export All Locations...", "btnExportAll", 130
            DrawOne ws, settingsBottom + 4, 4, "Import (choose room)...", "btnImportGlobal", 130
            DrawOne ws, settingsBottom + 4, 7, "Backup workbook...", "btnBackupWorkbook", 130
            DrawOne ws, settingsBottom + 4, 10, "Restore workbook...", "btnRestoreWorkbook", 130
            ' A third row, direct user request (2026-09-27): automates the
            ' documented manual add-a-location procedure (ARCHITECTURE.md
            ' §4.4 - modRegistry.AddPrintRoom's own comment). Its own row
            ' rather than a fifth slot on either group above - it is neither
            ' an "everyday workbook action" nor a data-movement action, and
            ' the two existing rows are already full at four buttons each.
            ' Remove print room (2026-09-27, same day) sits right next to it
            ' rather than opening a fourth row - Add/Remove read as one pair,
            ' and this is the "later look at button arrangement" §16.3
            ' flagged: putting the new button beside its natural counterpart
            ' instead of picking its own spot in isolation.
            DrawOne ws, settingsBottom + 6, 1, "Add print room...", "btnAddPrintRoom", 130
            DrawOne ws, settingsBottom + 6, 4, "Remove print room...", "btnRemovePrintRoom", 130
            ' Small +/- buttons above the three lookup tables (snag list item
            ' 5). Row 4 is already the table's own subtitle ("Paper stock
            ' types" etc.) on this sheet, unlike the blank row 4 on Printers/
            ' Papers/Print Technicians, so these sit in row 3 instead.
            '
            ' "Size"/"Consumable" rather than "StandardSizes"/
            ' "Consumables": Button.Name silently TRUNCATES
            ' to 31 characters at 32 and raises 1004 outright at 33+ in this
            ' Excel/COM automation context (verified directly with a length
            ' sweep - "pcb_btnRemoveRowConsumables_3_18" was exactly 32 and
            ' came back with its trailing "8" dropped, not an error, which
            ' is a worse bug than a clean failure would have been. The other
            ' three, at 34 each, raised 1004 outright.
            DrawSmall ws, 3, 6, "+", "btnAddRowPaperTypes", 24
            DrawSmall ws, 3, 7, "-", "btnRemoveRowPaperTypes", 24
            DrawSmall ws, 3, 9, "+", "btnAddRowSize", 24
            DrawSmall ws, 3, 10, "-", "btnRemoveRowSize", 24
            DrawSmall ws, 3, 13, "+", "btnAddRowConsumable", 24
            DrawSmall ws, 3, 14, "-", "btnRemoveRowConsumable", 24
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
            ' "DelVis" tag: "btnDeleteVisibleReports" (23 chars) + "_3_20"
            ' (5, now two-digit) would put the shape's own Name at exactly
            ' 32 characters - silently truncated by Excel/COM, per DrawOne's
            ' own comment - without shortening what actually runs on click.
            DrawOne ws, 1, 20, "Export report...", "btnExportReport", 140
            DrawOne ws, 3, 20, "Delete visible records...", "btnDeleteVisibleReports", 140, "DelVis"
            ' "Mark all as..." cluster (2026-09-29, direct user request): O1
            ' holds the label (modReports.BuildReports - see there for why
            ' column O), then one button per row - Paid in row 2, Unpaid in
            ' row 3. Sized to the ROW (a shade under its height) rather than
            ' DrawOne's 22pt, which would spill each button into the row
            ' below and overlap its neighbour. Row heights are read here, not
            ' assumed: BuildReportSheets has already run, so A1's 16pt title
            ' has already settled row 1. Width 78pt sits inside column O's
            ' ~80pt, clear of the T buttons (848pt) beyond S.
            DrawOneAtTop ws, 15, ws.Cells(2, 15).Top + 0.5, "Paid", "btnMarkPaid", 78, ws.Cells(2, 15).Height - 1
            DrawOneAtTop ws, 15, ws.Cells(3, 15).Top + 0.5, "Unpaid", "btnMarkUnpaid", 78, ws.Cells(3, 15).Height - 1
            ' Freezes above the print-job results table (row 15) so its
            ' header row and the filter/totals area above stay visible while
            ' scrolling through matches - snag list item 9.
            SetFreeze ws, "A15"
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
    ' short gap rows added by EnsureHeaderGaps) sits above it. Two columns
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

    ' View row (2026-10-01): "Show columns:" label, an All/Reduced/Minimal
    ' drop-down and the Hide/Show cost detail button on row 2, from D.
    ' Positioned (and kept off hidden columns) by RepositionViewButtons; drawn
    ' here at their nominal columns.
    DrawViewDropDown ws
    DrawOne ws, VIEW_ROW, VIEW_FIRST_COL + 2, CostColumnsCaption(), "btnToggleCostColumns", VIEW_COST_W

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
    ' D1:D2 used to hold a heading and explanatory prose; cleared, styles
    ' included, so nothing of them lingers on sheets built before this.
    ws.Range("D1:D2").Clear

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

' Opens a short blank gap row between the department and the defaults, and
' between the defaults and the disregard options. Idempotent: detected from
' where the named default cells sit, so it migrates older sheets once and then
' does nothing. Insert is entire-row, so named cells, the job table and the
' side-panel cells all shift together; buttons are free-floating, so any at or
' below the first gap are moved down by hand.
Public Sub EnsureHeaderGaps(ByVal ws As Worksheet)
    Dim tech As Range, paper As Range, dis As Range
    Set tech = LocRange(ws, "LOC_DefTech")
    Set paper = LocRange(ws, "LOC_DefPaper")
    Set dis = LocRange(ws, "LOC_DefDisPaper")
    If tech Is Nothing Or paper Is Nothing Or dis Is Nothing Then Exit Sub

    UnlockSheet ws
    If tech.Row = ROW_DEF_TECH - 1 Then InsertGapRow ws, ROW_DEF_TECH - 1
    If dis.Row = paper.Row + 1 Then InsertGapRow ws, paper.Row + 1
    RelockSheet ws
End Sub

Private Sub InsertGapRow(ByVal ws As Worksheet, ByVal Row As Long)
    Dim cutoff As Double, i As Long
    cutoff = ws.Rows(Row).Top
    ws.Rows(Row).Insert Shift:=xlDown
    With ws.Rows(Row)
        .ClearFormats
        .Validation.Delete
        .RowHeight = GAP_ROW_HEIGHT
    End With
    For i = 1 To ws.Buttons.Count
        If ws.Buttons(i).Top >= cutoff Then ws.Buttons(i).Top = ws.Buttons(i).Top + GAP_ROW_HEIGHT
    Next i
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
' in this unit on THIS sheet, and EnsureRollUnitFormulas makes Unit/Area m2/
' Paper Cost convert back to metres, as does the consolidated _Data range
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

' Unit/Area m2/Paper Cost, restated so a Centimetres room's Qty is read as
' centimetres: Unit says "cm", and the two calculations divide a "cm" row's
' Qty by 100 so they still work in metres - Area m2 and money are the same
' whatever the sheet shows. Called from InitialiseWorkbook and Refresh
' Locations, after EnsureRollUnitSetting (the formulas name LOC_RollUnit).
'
' One-time migration: a sheet whose Unit formula predates this (no mention of
' LOC_RollUnit) and is set to Centimetres holds metres in Qty, since the old
' build converted on entry - so its roll rows are scaled x100 here, once. The
' formula rewrite below is what makes it once.
Public Sub EnsureRollUnitFormulas(ByVal ws As Worksheet)
    Dim lo As ListObject, legacy As Boolean, m As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Unit") Or Not ColumnExists(lo, "Qty") Then Exit Sub

    UnlockSheet ws
    legacy = (InStr(1, lo.ListColumns("Unit").DataBodyRange.Cells(1, 1).Formula2, "LOC_RollUnit", vbTextCompare) = 0)
    If legacy Then
        If StrComp(RollUnitOf(ws), "Centimetres", vbTextCompare) = 0 Then ScaleRollQty lo, 100
        SetAppliedRollUnit ws, RollUnitOf(ws)
    End If

    m = "IF([@Unit]=""cm"",100,1)"
    lo.ListColumns("Unit").DataBodyRange.Formula2 = _
        "=IF([@[S_Measure]]="""","""",IF([@[S_Measure]]=""Sheet"",""sheets"",IF(LOC_RollUnit=""Centimetres"",""cm"",""metres"")))"
    If ColumnExists(lo, "Area m2") Then
        lo.ListColumns("Area m2").DataBodyRange.Formula2 = _
            "=IF([@Qty]="""","""",IF([@[S_Measure]]=""Sheet"",([@[S_StockWidth_mm]]/1000)*([@[S_SheetHeight_mm]]/1000)*[@Qty]," & _
            "(IF([@[Print Width mm]]="""",[@[S_StockWidth_mm]],[@[Print Width mm]])/1000)*[@Qty]/" & m & "))"
    End If
    If ColumnExists(lo, "Paper Cost") Then
        lo.ListColumns("Paper Cost").DataBodyRange.Formula2 = _
            "=IF([@Qty]="""","""",ROUND([@Qty]/" & m & "*[@[S_UnitCost]],SET_ROUND_DP))"
    End If
    RelockSheet ws
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
    ws.Cells(ROW_JOB_COUNT, 2).Formula = "=ROWS(" & lo.Name & ")"
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
' cells BELOW the table is a separate matter - see ClearBelowTableValidation, run once per
' sheet by setup and Refresh Locations, not here (this runs on every row added). Rows the
' table grows into take their validation from the row above, i.e. from here.
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
        "Choose one of the standard sizes listed on Settings.", , True

    AddRule lo, "Disregard Paper", xlValidateList, "Yes,No", True
    AddRule lo, "Disregard Consumable", xlValidateList, "Yes,No", True

    AddRule lo, "Paid", xlValidateList, "Yes,No", True, "Paid", _
        "Whether this chargeable cost has been paid. Blank means not recorded either way and counts as unpaid in totals.", _
        "Choose Yes or No.", , True

    RelockSheet ws
End Sub

' Strips validation from the cells under the job table, across the table's own columns, down
' to the end of the sheet. Validation there is not part of the table: the shipped template
' once carried hand-placed rules on rows 28-2010 (a Yes/No list on what had become Sheet
' size) that EnsureJobColumnValidation never saw and rows added to the table could inherit.
'
' Called once per sheet from setup and Refresh Locations - NOT from BindColumns. Run on every
' row added (every BindColumns), it left Excel rejecting the very next COM call and
' test-suppliedstock failed 6 of 6. Finds the validated cells first (SpecialCells covers the
' used range only) so a clean sheet is left untouched.
Public Sub ClearBelowTableValidation(ByVal ws As Worksheet)
    Dim lo As ListObject, firstRow As Long, firstCol As Long, lastCol As Long, stray As Range
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    firstRow = lo.Range.Row + lo.Range.Rows.Count
    If firstRow > ws.Rows.Count Then Exit Sub
    firstCol = lo.Range.Column
    lastCol = firstCol + lo.Range.Columns.Count - 1

    On Error Resume Next    ' SpecialCells raises when nothing matches
    Set stray = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(ws.Rows.Count, lastCol)).SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0
    If stray Is Nothing Then Exit Sub

    UnlockSheet ws
    stray.Validation.Delete
    RelockSheet ws
End Sub

' One validation rule on one job-table column. Title is used for both the input and
' error dialogs; Title, Prompt (the input message) and ErrText (the error message) are
' left unset when blank. SkipIfMissing is for columns an older sheet may not have.
Private Sub AddRule(ByVal lo As ListObject, ByVal ColName As String, ByVal RuleType As Long, _
        ByVal Formula1 As String, ByVal IgnoreBlank As Boolean, Optional ByVal Title As String, _
        Optional ByVal Prompt As String, Optional ByVal ErrText As String, _
        Optional ByVal Op As Long = xlBetween, Optional ByVal SkipIfMissing As Boolean)
    If SkipIfMissing Then
        If Not ColumnExists(lo, ColName) Then Exit Sub
    End If
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
    If Not ColumnExists(lo, "H_Issues") Then Exit Sub
    If Not ColumnExists(lo, "Sheet size") Then Exit Sub

    UnlockSheet ws
    lo.ListColumns("H_Issues").DataBodyRange.Formula2 = BuildJobIssuesFormula()
    RelockSheet ws
End Sub

' One IF per rule - IF(condition,"; message","") - joined with & in a fixed order, all under
' an outer IF that leaves the cell blank on a row with no Job ID.
Private Function BuildJobIssuesFormula() As String
    Const DQ As String = """"
    Dim qty As String, width As String, stock As String, measure As String
    qty = RowRef("Qty")
    width = RowRef("Print Width mm")
    stock = RowRef("Paper Stock")
    measure = RowRef("S_Measure")

    BuildJobIssuesFormula = "=IF(" & RefBlank(RowRef("Job ID")) & "," & DQ & DQ & "," & _
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
        "&" & IssueIf("AND(" & stock & "=" & DQ & SUPPLIED_SHEET & DQ & "," & RefBlank(RowRef("Sheet size")) & ")", "Sheet size required for student-supplied sheet stock") & _
        ")"
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

' The row just below whichever ListObject on this sheet runs deepest - e.g.
' tblSettings on Settings, which grows every time modVersion.EnsureSetting
' adds a self-provisioned row. Used to anchor content that must sit clear of
' every table on the sheet regardless of how many rows each currently has.
Private Function TablesBottom(ByVal ws As Worksheet) As Long
    Dim lo As ListObject, bottom As Long
    For Each lo In ws.ListObjects
        If lo.Range.Row + lo.Range.Rows.Count - 1 > bottom Then bottom = lo.Range.Row + lo.Range.Rows.Count - 1
    Next lo
    TablesBottom = bottom
End Function

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
    ConfigSheetNames = Array("Print Technicians", "Printers", "Papers", "Settings")
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
        If ColumnExists(lo, idHdr) Then lo.ListColumns(idHdr).DataBodyRange.Locked = True
    End If
    RelockSheet lo.Parent
End Sub

' tblSettings mixes user-editable settings with system-derived rows (schema
' version, last-refresh stamp, build stamp) whose Notes column is written as
' "Read-only. ..." (modVersion.EnsureSetting) - that text is the one place
' the two kinds are already told apart, so it drives which Value cells unlock
' rather than a second hard-coded list of keys that could drift from it.
Private Sub UnlockSettingsValues()
    Dim lo As ListObject, i As Long, notes As String
    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    UnlockSheet lo.Parent
    For i = 1 To lo.ListRows.Count
        notes = Trim$(CStr(CellIn(lo, i, "Notes").Value))
        CellIn(lo, i, "Value").Locked = (Left$(notes, 9) = "Read-only")
    Next i
    RelockSheet lo.Parent
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
' "Reduced" or "Minimal", held in SET_LOC_REDUCED_VIEW (a legacy "Yes" reads as
' Reduced, anything else as All). Reduced hides SET_LOC_REDUCED_COLUMNS;
' Minimal hides those AND SET_LOC_MINIMAL_COLUMNS, so it always includes the
' Reduced list.
Private Function ViewMode() As String
    Select Case LCase$(SettingText("LOC_REDUCED_VIEW", "All"))
        Case "reduced", "yes": ViewMode = "Reduced"
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

' Self-provisions the view settings. The two shipped in PrintCosts.xlsx
' (LOC_REDUCED_VIEW, LOC_REDUCED_COLUMNS) are kept and re-labelled; the old
' Yes/No value and the old seven-column default are migrated once.
Public Sub EnsureViewSettings()
    Dim c As Range

    Set c = EnsureSetting("LOC_REDUCED_VIEW", "Location column view", "")
    SetSettingText c, "Location column view", "Which columns location sheets show: All, Reduced or Minimal. Set by the Show columns drop-down on each location sheet; applies to every location."
    Select Case LCase$(Trim$(CStr(c.Value)))
        Case "all", "reduced", "minimal"
        Case "yes": SetSettingValue c, "Reduced"
        Case Else: SetSettingValue c, "All"
    End Select

    Set c = EnsureSetting("LOC_REDUCED_COLUMNS", "Reduced view - hidden columns", "")
    SetSettingText c, "Reduced view - hidden columns", "Semicolon-separated column headers hidden by the Reduced view. Edit to change what it hides - no rebuild needed. Minimal hides these as well."
    If Len(Trim$(CStr(c.Value))) = 0 Or StrComp(Trim$(CStr(c.Value)), REDUCED_COLUMNS_LEGACY, vbTextCompare) = 0 Then
        SetSettingValue c, REDUCED_COLUMNS_DEFAULT
    End If

    Set c = EnsureSetting("LOC_MINIMAL_COLUMNS", "Minimal view - extra hidden columns", "")
    SetSettingText c, "Minimal view - extra hidden columns", "Semicolon-separated column headers the Minimal view hides IN ADDITION to the Reduced list above (Minimal always includes Reduced). Edit to change - no rebuild needed."
    If Len(Trim$(CStr(c.Value))) = 0 Then SetSettingValue c, MINIMAL_EXTRA_DEFAULT
End Sub

' Label (column B) and notes (column D) of a settings row, from its Value cell.
Private Sub SetSettingText(ByVal ValueCell As Range, ByVal Label As String, ByVal Notes As String)
    UnlockSheet ValueCell.Parent
    ValueCell.Offset(0, -1).Value = Label
    ValueCell.Offset(0, 1).Value = Notes
    RelockSheet ValueCell.Parent
End Sub

Private Sub SetSettingValue(ByVal ValueCell As Range, ByVal Value As String)
    UnlockSheet ValueCell.Parent
    ValueCell.Value = Value
    RelockSheet ValueCell.Parent
End Sub

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
    ApplyColumnVisibility lo, Array("Paper Cost", "Consumable Cost", "Gross Cost", "Disregarded"), CostColumnsHidden()
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
    ' Self-heal (2026-09-26, user-reported): restores Width/Height/Placement
    ' too, not just Left, for any button built before DrawOne started setting
    ' Placement:=xlFreeFloating. Those older buttons still have
    ' Placement:=xlMoveAndSize (Excel's own default), so a reduced-view
    ' hide/reveal cycle shrinks their Width along with moving them - fixing
    ' only .Left left them stuck narrow (the reported Now button showing just
    ' its "N" at the D/E border). All four buttons this Sub relocates are
    ' drawn 110pt/22pt by DrawLocationButtons.
    b.Width = 110
    b.Height = 22
    b.Placement = xlFreeFloating
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
    If StrComp(header, "H_Issues", vbTextCompare) = 0 Then Exit Function
    If Left$(header, 2) = "S_" Then Exit Function
    IsSafeAnchorColumn = Not CBool(lo.ListColumns(col).Range.EntireColumn.Hidden)
End Function

' Narrower default widths for the columns that need the least room to show
' their actual content, freeing screen space for Student Name/Notes/etc.
' Job ID gets a fixed 100px rather than a cap - it holds a multi-digit
' correlation ID that genuinely needs it, so there is no narrower "good
' enough" to allow. Date/Time (180px), Print Width mm (120px) and the two
' Disregard columns (130px each) are likewise fixed widths, not caps - a full
' "dd/mm/yyyy hh:mm" stamp and the Yes/No dropdown labels were both getting
' clipped at their old, narrower widths.
' modUtils.ColWidthForPx does the character-unit conversion; a header not
' present (e.g. a much older sheet mid-migration) is skipped via ColIdx's own
' error rather than aborting the rest.
Public Sub ApplyJobColumnWidths(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    UnlockSheet ws
    SetJobColWidth lo, "Unit", ColWidthForPx(50)
    SetJobColWidth lo, "Qty", ColWidthForPx(50)
    SetJobColWidth lo, "Area m2", ColWidthForPx(60)
    SetJobColWidth lo, "Paid", ColWidthForPx(40)
    SetJobColWidth lo, "Job ID", ColWidthForPx(100)
    SetJobColWidth lo, "Date/Time", ColWidthForPx(180)
    SetJobColWidth lo, "Print Width mm", ColWidthForPx(120)
    SetJobColWidth lo, "Disregard Paper", ColWidthForPx(130)
    SetJobColWidth lo, "Disregard Consumable", ColWidthForPx(130)
    ' 2026-09-29 (user-reported: several columns excessively wide). Every
    ' remaining sized column now gets an explicit width too, rather than
    ' keeping whatever PrintCosts.xlsx shipped (Chargeable Cost and Paper
    ' Stock were 220px, Paper Cost/Disregarded/Sheet size 115px). Headers
    ' wrap (row 13 is two lines tall), so a column only needs its longest
    ' word plus the filter button; the widths below are that or the widest
    ' typical content, whichever is larger. Notes is left alone (free text).
    ' Status is short (120px) because RefreshStatusNotes puts the full
    ' message in a hover note on every row that has a problem.
    SetJobColWidth lo, "Status", ColWidthForPx(120)
    SetJobColWidth lo, "Technician", ColWidthForPx(120)
    SetJobColWidth lo, "Printer", ColWidthForPx(160)
    SetJobColWidth lo, "Paper Stock", ColWidthForPx(170)
    SetJobColWidth lo, "Sheet size", ColWidthForPx(90)
    SetJobColWidth lo, "Paper Cost", ColWidthForPx(90)
    SetJobColWidth lo, "Consumable Cost", ColWidthForPx(100)
    SetJobColWidth lo, "Gross Cost", ColWidthForPx(90)
    SetJobColWidth lo, "Disregarded", ColWidthForPx(100)
    SetJobColWidth lo, "Chargeable Cost", ColWidthForPx(100)
    RelockSheet ws
End Sub

Private Sub SetJobColWidth(ByVal lo As ListObject, ByVal Header As String, ByVal Width As Double)
    On Error Resume Next
    lo.ListColumns(Header).Range.EntireColumn.ColumnWidth = Width
    On Error GoTo 0
End Sub

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
    Dim lo As ListObject, rng As Range, fc As FormatCondition
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = lo.ListColumns("Status").DataBodyRange

    UnlockSheet ws
    rng.FormatConditions.Delete
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(" & rng.Cells(1, 1).Address(False, False) & "<>""""," & _
                  rng.Cells(1, 1).Address(False, False) & "<>""OK"")")
    fc.Interior.Color = RGB(255, 192, 0)
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
