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
' used to seed the SET_LOC_REDUCED_COLUMNS setting the first time and as the
' fallback if that setting is ever cleared. Declared here with this module's
' other module-level constant (BTN_TAG) rather than down by the code that
' uses it - every other module in this project keeps its Const/Dim
' declarations clustered at the top, and a module-level Const declared after
' a Sub/Function has already appeared in the source left this workbook
' failing to compile ("Variable not defined") even though the declaration
' itself was syntactically fine on its own.
Private Const REDUCED_COLUMNS_DEFAULT As String = "Status;Job ID;Printer;Area m2;Disregard Paper;Disregard Consumable;S_SchemaVer"

' SidePanelButtonLayout's own inputs (2026-09-27, cost-columns toggle) -
' declared here with this module's other module-level constants for exactly
' the reason REDUCED_COLUMNS_DEFAULT's own comment just above gives. Learn
' from that one: a first draft of this pair was declared mid-file, next to
' SidePanelButtonLayout itself, and hit the identical "Variable not defined"
' compile failure this comment already warns about.
Private Const SIDE_PANEL_BTN_COUNT As Long = 8
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
    EnsureExportSettings
    EnsureReducedViewSettings
    EnsureCostColumnsSetting
    EnsureCatalogIdSettings
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
            EnsureToolbarGap ws
            DrawLocationButtons ws
            ConfigValidation ws
            EnsureJobDefaults ws
            EnsureRollUnitSetting ws
            EnsurePrintersDisplay ws
            EnsureJobCountDisplay ws
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
            EnsureTableGap ws, "tblPrinters", 6
            DrawOne ws, 4, 1, "Add row", "btnAddRowPrinters", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowPrinters", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearPrinters", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Papers", vbTextCompare) = 0 Then
            EnsureTableGap ws, "tblPapers", 6
            DrawOne ws, 4, 1, "Add row", "btnAddRowPapers", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowPapers", 110
            DrawOne ws, 4, 5, "Clear table", "btnClearPapers", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Print Technicians", vbTextCompare) = 0 Then
            EnsureTableGap ws, "tblTechnicians", 6
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
            ' Papers/Print Technicians, so these sit in row 3 instead rather
            ' than moving the tables the way EnsureTableGap does for those.
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
    ' The placeholder labels sit on row 10 in the shipped .xlsx, but
    ' EnsureToolbarGap (called just before this, per-sheet) has already
    ' inserted one blank row above it - direct user feedback, testing the
    ' 2026-09-25 rework below, that the toolbar needed breathing room from
    ' the config block ending at row 9 - so by the time this runs the
    ' toolbar's real row is 11. Two columns apart; buttons are drawn over the
    ' (now-shifted) placeholders and the labels cleared.
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
    DrawOne ws, 11, 1, "Add Print Job", "btnAddPrintJob", 110
    DrawOne ws, 11, 3, "Repeat Job", "btnRepeatJob", 110
    DrawOne ws, 11, 5, "Now", "btnNow", 110

    Dim c As Long
    For c = 1 To 15
        With ws.Cells(11, c)
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
        capts = Array("Select printers...", "Check this sheet", ReducedViewCaption(), CostColumnsCaption(), "Remove Row", "Clear All", "Export...", "Import...")
        macros = Array("btnSelectPrinters", "btnCheckSheet", "btnToggleReducedView", "btnToggleCostColumns", "btnRemoveRow", "btnClearAll", "btnExport", "btnImportLocation")
        For i = 0 To 7
            DrawOneAtTop ws, SIDE_PANEL_COL, i * stepPx, CStr(capts(i)), CStr(macros(i)), 140, btnH
        Next i
    End If

    ' Clear defaults - column D, row 4 (over Technician), roughly centred
    ' against the three-row default-selector block it clears (A3:B5) without
    ' overlapping the labels themselves - column D is free there (D1:D2 hold
    ' the block's own explanatory prose, D3 onward is empty). Still anchored
    ' inside the table's column span, so RelocateAtRiskButtons covers it
    ' (below).
    DrawOne ws, 4, 4, "Clear defaults", "btnClearDefaults", 110
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

' A single blank row between the configuration block (ends row 9, since the
' 2026-09-25 layout rework, above) and the toolbar - direct user feedback
' after testing that rework: with the table's header reverting to its
' originally-shipped row 12 (EnsureJobTableGap removed), the toolbar sat
' flush against Print jobs (row 9) with no breathing room. Narrower in scope
' than the old EnsureJobTableGap this replaces - one row, not two, since
' only the toolbar needs separating from the config block now, not a whole
' extra defaults row - but the same idiom: checked-first via the table's own
' row, so it is safe to call on a sheet that has already been migrated.
' Setup-time only (InitialiseWorkbook), not RefreshLocations - a duplicated
' sheet already carries the gap with it, same as the block above it.
Public Sub EnsureToolbarGap(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.Range.Row >= 13 Then Exit Sub

    UnlockSheet ws
    ws.Rows(10).Insert Shift:=xlDown
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
    ws.Range("A3").Value = "Default: technician"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "Default: printer"
    ws.Range("A4").Font.Bold = True
    ws.Range("A5").Value = "Default: paper"
    ws.Range("A5").Font.Bold = True

    EnsureLocName ws, "LOC_DefTech", "$B$3"
    EnsureLocName ws, "LOC_DefPrinter", "$B$4"
    EnsureLocName ws, "LOC_DefPaper", "$B$5"

    StyleInputCell ws.Range("B3")
    StyleInputCell ws.Range("B4")
    StyleInputCell ws.Range("B5")
    RelockSheet ws

    ' Both directions of spec 1a's filtering apply here too (spec 1b: "these
    ' selectors should implement the same filtering and autofill principles
    ' as the table cells").
    BindDefaultCells ws
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
' Only Qty's stored value ever changes to reflect this (modValidation.
' OnQtyChanged, converting a Centimetres-location's raw entry to metres on
' the cell itself) - Area m2/Paper Cost/every other formula on the row reads
' Qty exactly as before and has no idea this setting exists.
Public Sub EnsureRollUnitSetting(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Range("A6").Value = "Roll length unit"
    ws.Range("A6").Font.Bold = True

    EnsureLocName ws, "LOC_RollUnit", "$B$6"

    Dim c As Range
    Set c = ws.Range("B6")
    If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = "Metres"
    StyleInputCell c
    With c.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Metres,Centimetres"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Roll length unit"
        .InputMessage = "How a roll job's length is typed into Qty on this sheet. Always converted to, and stored as, metres regardless of this setting - Sheet stock is unaffected."
        .ErrorTitle = "Roll length unit"
        .ErrorMessage = "Choose Metres or Centimetres."
    End With
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
    ws.Range("A7").Value = "Printers at this location"
    ws.Range("A7").Font.Bold = True

    ' Built with Chr(34) rather than a hand-escaped string literal - the
    ' quote-doubling needed to embed this many nested string arguments in a
    ' single VBA literal is too easy to get subtly wrong to trust by eye.
    ' Target formula (each q below is one literal "):
    '   =LET(list,LOC_Printers,n,IF(list="",0,LEN(list)-LEN(SUBSTITUTE(list,
    '   ";",""))+1),IF(n=0,"0 printers ()",n&" printer"&IF(n=1,"","s")&
    '   " ("&SUBSTITUTE(list,";",", ")&")"))
    Dim q As String
    q = Chr(34)
    ws.Range("B7").Formula2 = "=LET(list,LOC_Printers,n,IF(list=" & q & q & _
        ",0,LEN(list)-LEN(SUBSTITUTE(list," & q & ";" & q & "," & q & q & _
        "))+1),IF(n=0," & q & "0 printers ()" & q & ",n&" & q & " printer" & _
        q & "&IF(n=1," & q & q & "," & q & "s" & q & ")&" & q & " (" & q & _
        "&SUBSTITUTE(list," & q & ";" & q & "," & q & ", " & q & ")&" & q & _
        ")" & q & "))"
    ws.Range("B7").Locked = True
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
    ws.Range("A8").Value = "Print jobs"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").HorizontalAlignment = xlRight
    ws.Range("B8").Formula = "=ROWS(" & lo.Name & ")"
    ws.Range("B8").Locked = True
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
Public Sub EnsureJobColumnValidation(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim paidCol As ListColumn
    If lo.DataBodyRange Is Nothing Then Exit Sub

    UnlockSheet ws
    lo.DataBodyRange.Validation.Delete

    With lo.ListColumns("Date/Time").DataBodyRange.Validation
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="01/01/2000"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Date and time"
        .InputMessage = "Date and time the print was produced. Use the Now button to stamp the current date and time."
        .ErrorTitle = "Date and time"
        .ErrorMessage = "Date and time the print was produced. Use the Now button to stamp the current date and time."
    End With

    Dim qtyMsg As String
    If StrComp(LocValue(ws, "LOC_RollUnit"), "Centimetres", vbTextCompare) = 0 Then
        qtyMsg = "Sheets for sheet stock, centimetres for roll stock (converted and stored as metres - see 'Roll length unit' above). Must be greater than zero."
    Else
        qtyMsg = "Sheets for sheet stock, metres for roll stock. Must be greater than zero."
    End If
    With lo.ListColumns("Qty").DataBodyRange.Validation
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Qty"
        .InputMessage = qtyMsg
        .ErrorTitle = "Qty"
        .ErrorMessage = qtyMsg
    End With

    With lo.ListColumns("Print Width mm").DataBodyRange.Validation
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Print width"
        .InputMessage = "Optional, roll stock only, in millimetres. Leave blank to use the full width of the roll. It must not exceed the stock width. Required for 'Supplied (Roll)'."
        .ErrorTitle = "Print width"
        .ErrorMessage = "Optional, roll stock only, in millimetres. Leave blank to use the full width of the roll. It must not exceed the stock width."
    End With

    Dim sheetSizeCol As ListColumn
    On Error Resume Next
    Set sheetSizeCol = lo.ListColumns("Sheet size")
    On Error GoTo 0
    If Not sheetSizeCol Is Nothing Then
        With sheetSizeCol.DataBodyRange.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=RNG_STD_SIZES"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
            .InputTitle = "Sheet size"
            .InputMessage = "Required only for 'Supplied (Sheet)' - the nearest standard size to the sheet the student brought. Ignored for every other paper stock."
            .ErrorTitle = "Sheet size"
            .ErrorMessage = "Choose one of the standard sizes listed on Settings."
        End With
    End If

    With lo.ListColumns("Disregard Paper").DataBodyRange.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    With lo.ListColumns("Disregard Consumable").DataBodyRange.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    On Error Resume Next
    Set paidCol = lo.ListColumns("Paid")
    On Error GoTo 0
    If Not paidCol Is Nothing Then
        With paidCol.DataBodyRange.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
            .InputTitle = "Paid"
            .InputMessage = "Whether this chargeable cost has been paid. Blank means not recorded either way and counts as unpaid in totals."
            .ErrorTitle = "Paid"
            .ErrorMessage = "Choose Yes or No."
        End With
    End If

    RelockSheet ws
End Sub

' Extends the calculated H_Issues column PrintCosts.xlsx ships (§5, §7.3)
' with the two required-field checks the student-supplied stock options
' need: Print Width mm on a "Supplied (Roll)" row and Sheet size on a
' "Supplied (Sheet)" row have no catalogue size to fall back on the way an
' ordinary stock does, so a value left blank there would otherwise cost a
' real print job nothing rather than being caught by Check sheet/Check
' workbook the way every other required field already is.
'
' Asserted here rather than left as a static xlsx formula, the same
' reasoning EnsureJobColumnValidation (above) already applies to this
' table's validation rules: H_Issues is never hand-maintained per sheet, so
' restating its full text in VBA is what lets an old or duplicated sheet
' self-heal on the next refresh, rather than PrintCosts.xlsx itself needing
' a hand-edited calculated-column formula (§5's own note on how easy that
' class of edit is to get subtly wrong). Table-name-parameterised via
' lo.Name, the same idiom EnsurePrintersDisplay/EnsureJobCountDisplay
' already use for their own per-sheet formulas - so this keeps working
' whatever the table is currently named (tblJobs_MAIN, tblJobs_ANNEX, ...).
' Called from modLists.BindColumns, right after EnsureJobColumnValidation -
' the same entry point every other self-healing table-structure fix in this
' project already reaches from setup, Refresh Locations, Check workbook/
' sheet and the picker.
Public Sub EnsureJobIssuesFormula(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim f As String, firstCell As Range
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "H_Issues") Then Exit Sub
    If Not ColumnExists(lo, "Sheet size") Then Exit Sub

    f = BuildJobIssuesFormula(lo.Name)

    ' Checked-first: BindColumns calls this on every setup run, Refresh
    ' Locations, Check workbook/sheet and the picker (module comment above),
    ' none of which are rare - unlike EnsureJobColumnValidation's cheaper
    ' Delete+Add, writing a ~1.5KB formula across the whole column is real
    ' work, and doing it unconditionally on every one of those calls was
    ' observed to make Excel's own recalculation noticeably heavier on a
    ' table already carrying a dozen rows. Comparing against the first
    ' cell's current formula (structurally identical down every row, same
    ' as the shipped column always was) is enough to know the rest matches
    ' too.
    Set firstCell = lo.ListColumns("H_Issues").DataBodyRange.Cells(1, 1)
    If firstCell.Formula2 = f Then Exit Sub

    UnlockSheet ws
    ' Formula2, not Formula: this formula runs well past 255 characters, and
    ' Range.Formula raises 1004 for anything longer when set via VBA/COM -
    ' the same limit EnsurePrintersDisplay's own B7 formula (above) already
    ' works around the same way.
    lo.ListColumns("H_Issues").DataBodyRange.Formula2 = f
    RelockSheet ws
End Sub

Private Function BuildJobIssuesFormula(ByVal t As String) As String
    Dim q As String, f As String
    q = Chr(34)

    f = "=IF(" & t & "[[#This Row],[Job ID]]=" & q & q & "," & q & q & "," & _
        "IF(" & t & "[[#This Row],[Date/Time]]=" & q & q & "," & q & "; Date and time required" & q & "," & q & q & ")" & _
        "&IF(AND(" & t & "[[#This Row],[Student Name]]=" & q & q & "," & t & "[[#This Row],[Student No]]=" & q & q & ")," & q & "; Student name or number required" & q & "," & q & q & ")" & _
        "&IF(" & t & "[[#This Row],[Technician]]=" & q & q & "," & q & "; Technician required" & q & "," & q & q & ")" & _
        "&IF(" & t & "[[#This Row],[Printer]]=" & q & q & "," & q & "; Printer required" & q & "," & q & q & ")" & _
        "&IF(" & t & "[[#This Row],[Paper Stock]]=" & q & q & "," & q & "; Paper stock required" & q & "," & q & q & ")" & _
        "&IF(" & t & "[[#This Row],[Qty]]=" & q & q & "," & q & "; Quantity required" & q & ",IF(" & t & "[[#This Row],[Qty]]<=0," & q & "; Quantity must be greater than zero" & q & "," & q & q & "))" & _
        "&IF(AND(" & t & "[[#This Row],[S_Measure]]=" & q & "Sheet" & q & "," & t & "[[#This Row],[Qty]]<>" & q & q & "," & t & "[[#This Row],[Qty]]<>INT(" & t & "[[#This Row],[Qty]]))," & q & "; Sheet quantity must be a whole number" & q & "," & q & q & ")" & _
        "&IF(AND(" & t & "[[#This Row],[S_Measure]]=" & q & "Sheet" & q & "," & t & "[[#This Row],[Print Width mm]]<>" & q & q & ")," & q & "; Print width does not apply to sheet stock" & q & "," & q & q & ")" & _
        "&IF(AND(" & t & "[[#This Row],[Print Width mm]]<>" & q & q & "," & t & "[[#This Row],[Print Width mm]]>" & t & "[[#This Row],[S_StockWidth_mm]])," & q & "; Print width exceeds stock width" & q & "," & q & q & ")" & _
        "&IF(AND(" & t & "[[#This Row],[Paper Stock]]=" & q & SUPPLIED_ROLL & q & "," & t & "[[#This Row],[Print Width mm]]=" & q & q & ")," & q & "; Print width required for student-supplied roll stock" & q & "," & q & q & ")" & _
        "&IF(AND(" & t & "[[#This Row],[Paper Stock]]=" & q & SUPPLIED_SHEET & q & "," & t & "[[#This Row],[Sheet size]]=" & q & q & ")," & q & "; Sheet size required for student-supplied sheet stock" & q & "," & q & q & ")" & _
        ")"

    BuildJobIssuesFormula = f
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
' without needing EnsureTableGap's row insert the way the wider text buttons
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

' The .xlsx ships Printers/Papers/Print Technicians with their table starting
' at row 5, directly under the row 4 the Add row/Remove row buttons (22px, so
' taller than the default row height) are drawn on - which visually overlaps
' the table header. Inserting a row above the table once fixes it without
' hand-editing the binary .xlsx; checked first (the table's own current row)
' so re-running setup never shifts an already-shifted table again.
Private Sub EnsureTableGap(ByVal ws As Worksheet, ByVal TableName As String, ByVal FirstRow As Long)
    Dim lo As ListObject
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    If lo.Range.Row >= FirstRow Then Exit Sub
    UnlockSheet ws
    ws.Rows(lo.Range.Row).Insert Shift:=xlDown
    RelockSheet ws
End Sub

Private Sub ClearButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Buttons.Count To 1 Step -1
        If Left$(ws.Buttons(i).Name, Len(BTN_TAG)) = BTN_TAG Then ws.Buttons(i).Delete
    Next i
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
Private Sub EnsureReducedViewSettings()
    Dim c As Range
    Set c = EnsureSetting("LOC_REDUCED_VIEW", "Reduced location view", "Yes hides the columns named in the setting below on every location sheet (toggled by the button on each one). No shows every column.")
    If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = "No"
    Set c = EnsureSetting("LOC_REDUCED_COLUMNS", "Reduced view - hidden columns", "Semicolon-separated column headers hidden by the reduced view above. Edit this list to change which columns it hides - no rebuild needed.")
    If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = REDUCED_COLUMNS_DEFAULT
End Sub

Private Function ReducedViewOn() As Boolean
    ReducedViewOn = (StrComp(SettingText("LOC_REDUCED_VIEW", "No"), "Yes", vbTextCompare) = 0)
End Function

Private Function ReducedViewCaption() As String
    ReducedViewCaption = IIf(ReducedViewOn(), "Show all columns", "Reduce clutter")
End Function

' Applies the CURRENT setting to one location sheet's table - called on
' every InitialiseWorkbook/RefreshLocations run (so a freshly duplicated
' sheet, or one predating the feature, always ends up in sync) and again
' from ToggleReducedView for every location sheet at once.
Public Sub ApplyReducedView(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    ApplyColumnVisibility lo, SplitList(SettingText("LOC_REDUCED_COLUMNS", REDUCED_COLUMNS_DEFAULT)), ReducedViewOn()
    RelocateAtRiskButtons ws, lo
    ' User-reported, 2026-09-26: hiding the shortlist (all of it left of
    ' SIDE_PANEL_COL) shrinks the total width in front of the side panel, so
    ' column 36's own pixel position shifts left - but with the side panel's
    ' buttons now Placement:=xlFreeFloating (RepositionLocationButtons' own
    ' comment), they no longer follow that shift and end up sitting further
    ' right of column 36 than before, not on it. ToggleReducedView reaches
    ' this Sub directly, without going through InitialiseWorkbook's separate
    ' RepositionLocationButtons call, so the side panel needs re-settling
    ' here too - on every call, not just from setup.
    RepositionSidePanelButtons ws
End Sub

' --------------------------------------------------------- cost columns ---
' Replaces the native-outline cost-columns group (docs/ARCHITECTURE.md
' §16.3, "button lag" writeup, 2026-09-27). A workbook-wide toggle, same
' shape as reduced-clutter view just above (simpler than per-sheet, and
' keeps every location in step) - "Yes" hides Paper Cost, Consumable Cost,
' Gross Cost and Disregarded on every location sheet; Chargeable Cost and
' Paid are never hidden by this, same columns the old outline group already
' kept outside it (2026-09-22 snag list item 1d).
Private Sub EnsureCostColumnsSetting()
    Dim c As Range
    Set c = EnsureSetting("COST_COLS_HIDDEN", "Hide cost detail columns", "Yes hides Paper Cost/Consumable Cost/Gross Cost/Disregarded on every location sheet (toggled by the button on each one). No shows them. Chargeable Cost and Paid are never hidden by this.")
    If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = "No"
End Sub

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
' one pass, same shape as ToggleReducedView.
Public Sub ToggleCostColumns()
    Dim hideIt As Boolean, ws As Worksheet
    hideIt = Not CostColumnsHidden()
    SetSetting "COST_COLS_HIDDEN", IIf(hideIt, "Yes", "No")

    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            ApplyCostColumnsVisibility ws
            RelabelCostColumnsButton ws
        End If
    Next ws
    AppOn
End Sub

Private Sub RelabelCostColumnsButton(ByVal ws As Worksheet)
    Dim i As Long, prefix As String
    prefix = BTN_TAG & "btnToggleCostColumns"
    For i = 1 To ws.Buttons.Count
        If Left$(ws.Buttons(i).Name, Len(prefix)) = prefix Then
            ws.Buttons(i).Caption = CostColumnsCaption()
            Exit For
        End If
    Next i
End Sub

' General fix for the column-hide-takes-a-button-with-it gap (§4.1) - only
' the handful of buttons DrawLocationButtons still anchors inside the job
' table's own column span (everything else lives in the side panel, immune
' by construction) need this. Called every time column visibility can have
' changed - InitialiseWorkbook, RefreshLocations and ToggleReducedView all
' reach it via ApplyReducedView - so a button self-heals back onto a visible
' column whichever way the setting just moved, including back to its own
' preferred column once reduced view is switched off again.
Private Sub RelocateAtRiskButtons(ByVal ws As Worksheet, ByVal lo As ListObject)
    RelocateButton ws, lo, "btnAddPrintJob", "Date/Time"
    RelocateButton ws, lo, "btnRepeatJob", "Student No"
    RelocateButton ws, lo, "btnNow", "Printer"
    RelocateButton ws, lo, "btnClearDefaults", "Technician"
End Sub

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
    macros = Array("btnSelectPrinters", "btnCheckSheet", "btnToggleReducedView", "btnToggleCostColumns", "btnRemoveRow", "btnClearAll", "btnExport", "btnImportLocation")
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
    ' typical content, whichever is larger. Status and Notes are left alone:
    ' Status needs room for its issue text, and Notes is free text.
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

' modMain.btnToggleReducedView's target. Flips the setting once, then
' re-applies it to every location sheet and relabels every toggle button in
' one pass, so all of them change state together rather than one at a time.
Public Sub ToggleReducedView()
    Dim reduceIt As Boolean, ws As Worksheet
    reduceIt = Not ReducedViewOn()
    SetSetting "LOC_REDUCED_VIEW", IIf(reduceIt, "Yes", "No")

    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            ApplyReducedView ws
            RelabelReducedViewButton ws
        End If
    Next ws
    AppOn
End Sub

Private Sub RelabelReducedViewButton(ByVal ws As Worksheet)
    Dim i As Long, prefix As String
    prefix = BTN_TAG & "btnToggleReducedView"
    For i = 1 To ws.Buttons.Count
        If Left$(ws.Buttons(i).Name, Len(prefix)) = prefix Then
            ws.Buttons(i).Caption = ReducedViewCaption()
            Exit For
        End If
    Next i
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
    RelockSheet ws
End Sub
