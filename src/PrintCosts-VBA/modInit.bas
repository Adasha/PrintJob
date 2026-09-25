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

' Import pairs horizontally with Export at SIDE_PANEL_COL2 (2026-09-25 - see
' DrawLocationButtons) rather than stacking below it: the table's header row
' moved up to 12 once EnsureJobTableGap was removed, and ReorderJobColumns'
' Cut+Insert operates on the table's whole header+data row span - so a
' floating button anchored to row 12 or below (row 13 was Import's original
' spot, landing exactly on the table's first data row) gets silently dragged
' sideways the next time a column reorder runs. Every side-panel row must
' stay at 11 or above; pairing two buttons per row rather than adding an
' 8th stacked row keeps that margin without cramming the existing six.
Private Const SIDE_PANEL_COL2 As Long = 48

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
    EnsureStdSizeColumnName

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
            EnsurePaidColumn ws
            EnsureQtyColumnName ws
            ReorderJobColumns ws
            BindColumns ws
            GroupJobColumns ws
            ApplyStatusFormat ws
            ApplyReducedView ws
            ApplyJobColumnWidths ws
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
            DrawOne ws, 11, 8, "Select families...", "btnSelectFamilies", 130
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Papers", vbTextCompare) = 0 Then
            EnsureTableGap ws, "tblPapers", 6
            DrawOne ws, 4, 1, "Add row", "btnAddRowPapers", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowPapers", 110
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Print Technicians", vbTextCompare) = 0 Then
            EnsureTableGap ws, "tblTechnicians", 6
            DrawOne ws, 4, 1, "Add row", "btnAddRowTechnicians", 110
            DrawOne ws, 4, 3, "Remove row", "btnRemoveRowTechnicians", 110
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
            ' Small +/- buttons above the four lookup tables (snag list item
            ' 5). Row 4 is already the table's own subtitle ("Paper stock
            ' types" etc.) on this sheet, unlike the blank row 4 on Printers/
            ' Papers/Print Technicians, so these sit in row 3 instead rather
            ' than moving the tables the way EnsureTableGap does for those.
            '
            ' "Family"/"Size"/"Consumable" rather than "PaperFamilies"/
            ' "StandardSizes"/"Consumables": Button.Name silently TRUNCATES
            ' to 31 characters at 32 and raises 1004 outright at 33+ in this
            ' Excel/COM automation context (verified directly with a length
            ' sweep - "pcb_btnRemoveRowConsumables_3_18" was exactly 32 and
            ' came back with its trailing "8" dropped, not an error, which
            ' is a worse bug than a clean failure would have been. The other
            ' three, at 34 each, raised 1004 outright.
            DrawSmall ws, 3, 6, "+", "btnAddRowPaperTypes", 24
            DrawSmall ws, 3, 7, "-", "btnRemoveRowPaperTypes", 24
            DrawSmall ws, 3, 9, "+", "btnAddRowFamily", 24
            DrawSmall ws, 3, 10, "-", "btnRemoveRowFamily", 24
            DrawSmall ws, 3, 13, "+", "btnAddRowSize", 24
            DrawSmall ws, 3, 14, "-", "btnRemoveRowSize", 24
            DrawSmall ws, 3, 17, "+", "btnAddRowConsumable", 24
            DrawSmall ws, 3, 18, "-", "btnRemoveRowConsumable", 24
            SetFreeze ws, ""
        ElseIf StrComp(ws.Name, "Reports", vbTextCompare) = 0 Then
            ' Rows 1-3, column F: clear of the title text (A1:A2) and above
            ' the filter/sort boxes (rows 5+), so both buttons sit inside the
            ' first screenful on any normal window - no scrolling needed.
            DrawOne ws, 1, 6, "Export report...", "btnExportReport", 140
            DrawOne ws, 3, 6, "Delete visible records...", "btnDeleteVisibleReports", 140
            ' Freezes above the print-job results table (row 15) so its
            ' header row and the filter/totals area above stay visible while
            ' scrolling through matches - snag list item 9.
            SetFreeze ws, "A15"
        End If
    Next ws

    UnlockConfigInputs
    FormatSettingsNotes
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
    DrawOne ws, 11, 1, "Add Print Job", "btnAddPrintJob", 110
    DrawOne ws, 11, 3, "Now", "btnNow", 110

    Dim c As Long
    For c = 1 To 15
        With ws.Cells(11, c)
            .ClearContents
            .Interior.Pattern = xlNone
        End With
    Next c

    ' Side panel - occasional-use buttons, one per row, matching the
    ' single-column-stacked idiom Summary's own off-table buttons already
    ' use. Grouped loosely by kind: setup/maintenance/view, then row-level
    ' correction, then whole-sheet data movement.
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
    DrawOne ws, 1, SIDE_PANEL_COL, "Select printers...", "btnSelectPrinters", 140
    DrawOne ws, 3, SIDE_PANEL_COL, "Check this sheet", "btnCheckSheet", 140
    DrawOne ws, 5, SIDE_PANEL_COL, ReducedViewCaption(), "btnToggleReducedView", 140
    DrawOne ws, 7, SIDE_PANEL_COL, "Remove Row", "btnRemoveRow", 140
    DrawOne ws, 9, SIDE_PANEL_COL, "Clear All", "btnClearAll", 140
    DrawOne ws, 11, SIDE_PANEL_COL, "Export...", "btnExport", 140
    DrawOne ws, 11, SIDE_PANEL_COL2, "Import...", "btnImportLocation", 140

    ' Clear defaults - column D, row 4 (over Technician), roughly centred
    ' against the three-row default-selector block it clears (A3:B5) without
    ' overlapping the labels themselves - column D is free there (D1:D2 hold
    ' the block's own explanatory prose, D3 onward is empty). Still anchored
    ' inside the table's column span, so RelocateAtRiskButtons covers it
    ' (below).
    DrawOne ws, 4, 4, "Clear defaults", "btnClearDefaults", 110
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
' pattern for the Reports page's own filter cells).
Private Sub StyleInputCell(ByVal target As Range)
    target.Locked = False
    target.Interior.Color = RGB(255, 255, 255)
    target.Borders(xlEdgeLeft).Color = RGB(46, 100, 168)
    target.Borders(xlEdgeLeft).Weight = xlMedium
End Sub

' -------------------------------------------------------------- Paid col ---
' Snag list item 1c: a genuine new job-row column (SCHEMA_VER bumped to 1.1,
' modUtils), so this only ever ADDS the column - it never runs against a
' sheet that already has it (checked first, so re-running setup is still
' idempotent). Positioned right after Chargeable Cost, ahead of Notes and the
' locked snapshot block - column order isn't load-bearing anywhere (§5.1),
' every consumer resolves it by header name.
'
' Left blank on existing rows deliberately, not force-defaulted to "No": a
' blank Paid means "not recorded either way" for a job that predates the
' column, and every consumer (Summary/Reports totals, export, import) treats
' blank the same as "No" rather than requiring a value. New rows still
' default to "No" explicitly - modJobs.AddPrintJob, same as the disregard
' flags.
Public Sub EnsurePaidColumn(ByVal ws As Worksheet)
    Dim lo As ListObject, lc As ListColumn, i As Long
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, "Paid", vbTextCompare) = 0 Then Exit Sub
    Next i

    UnlockSheet ws
    Set lc = lo.ListColumns.Add(ColIdx(lo, "Chargeable Cost") + 1)
    lc.Name = "Paid"
    If Not lc.DataBodyRange Is Nothing Then
        StyleInputCell lc.DataBodyRange
        With lc.DataBodyRange.Validation
            .Delete
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

' A pure rename (0.9.14, "shorten headers to save space"), not a new column -
' ListColumns(...).Name = handles the Table's internal bookkeeping (structured
' references, the _Data consolidation's own copy of this header) the same way
' Excel would if a person renamed it by hand, so every C("Qty")/SumBy("Qty")-
' style lookup elsewhere just needs to ask for the new name; nothing here
' migrates old data since nothing about the DATA changed, only its header
' text. Checked-first, like EnsurePaidColumn, so a sheet already renamed is
' left alone.
Public Sub EnsureQtyColumnName(ByVal ws As Worksheet)
    Dim lo As ListObject, i As Long
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, "Qty", vbTextCompare) = 0 Then Exit Sub
    Next i

    UnlockSheet ws
    On Error Resume Next
    lo.ListColumns("Quantity").Name = "Qty"
    On Error GoTo 0
    RelockSheet ws
End Sub

' ---------------------------------------------------------- column order ---
' Moves Status and Job ID from the start of the table (columns 1-2) to just
' after Paid, ahead of Notes/H_Issues/the snapshot block. Fixes a real bug
' found in the reduced-clutter view (snag 1e, above): Status and Job ID are
' two of its six hidden-by-default columns, but they used to sit at sheet
' columns A/B - the SAME columns the location config block above the table
' (room name, department, code, defaults - §4.1) occupies in rows 1-9.
' Hiding a column hides the WHOLE column, every row, not just the table's -
' so toggling reduced view was also blanking the room name/department/code
' the user needs to keep sight of. Moving Status/Job ID off columns A/B
' removes the collision without touching the config block at all.
'
' Column order is not load-bearing anywhere in this project (§5.1) -
' formulas use structured references, everything else resolves columns by
' header name - so this is free to do purely for layout reasons. The one
' exception needing a matching fix: modRegistry's consolidated-range span
' bounds itself by column NAME ("Job ID" to "Notes"), so moving Job ID away
' from being the leftmost column meant that bound had to move too (now
' "Date/Time" to "Notes" - still spans every real column, Status and Job ID
' included, since they now sit inside that span rather than starting it).
'
' NOTE (revisit): this fixes today's specific collision (columns A/B) but
' isn't a general solution - if SET_LOC_REDUCED_COLUMNS is ever edited to
' name a column that collides with the config block or the batch-defaults
' row for some other reason, the same class of bug could resurface. A more
' robust fix (decouple the config block's columns from the table's
' entirely, or keep them in sync some other way) is worth doing properly
' later rather than patching column-by-column.
' Static, non-catalog-driven validation for the job table (2026-09-25 bug
' report: every column from Unit through Chargeable Cost was showing a
' dropdown to pick a technician's name, and picking one overwrote the cell
' - breaking that row's formulas).
'
' Root cause, confirmed against the built .xlsm: ReorderJobColumns (below)
' moves columns via repeated Range.Cut + Range.Insert Shift:=xlToRight
' inside the table. Excel extends a validated column's rule onto cells
' newly shifted in beside it as each Insert runs, and the ~30 inserts one
' full reorder performs compound that into a wide, wrong span - Technician's
' own list (meant for one column) ended up the Formula1 on Unit through
' Chargeable Cost, and the Date/Time rule ended up on Student Name and
' Student No too. The values re-applied below are exactly what
' PrintCosts.xlsx ships on Date/Time, Qty, Print Width mm, Disregard Paper
' and Disregard Consumable - the same shipped rules ReorderJobColumns can
' disturb, restated in code so a refresh can restore them.
'
' Called from modLists.BindColumns - the same "rebuild dependent dropdowns,
' self-heal, never trust what a previous run left behind" reasoning that
' function already applies to Technician/Printer/Paper Stock covers these
' columns too, and reaches every existing BindColumns caller (setup,
' Refresh Locations, Check workbook/sheet, the picker) for free. Clearing
' the whole table body first means no stray rule from any past reorder can
' survive a refresh, whichever columns it ended up on - and cheap enough to
' do unconditionally, since none of those callers fire on every keystroke
' (OnPrinterChanged/OnStockChanged rebind a single row, never the table).
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
        .InputMessage = "Optional, roll stock only, in millimetres. Leave blank to use the full width of the roll. It must not exceed the stock width."
        .ErrorTitle = "Print width"
        .ErrorMessage = "Optional, roll stock only, in millimetres. Leave blank to use the full width of the roll. It must not exceed the stock width."
    End With

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

Public Sub ReorderJobColumns(ByVal ws As Worksheet)
    Dim lo As ListObject, order As Variant, i As Long, want As String, have As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    order = Array( _
        "Date/Time", "Student Name", "Student No", "Technician", "Printer", _
        "Paper Stock", "Unit", "Qty", "Print Width mm", "Disregard Paper", _
        "Disregard Consumable", "Area m2", "Paper Cost", "Consumable Cost", _
        "Gross Cost", "Disregarded", "Chargeable Cost", "Paid", "Status", "Job ID", _
        "Notes", "H_Issues", _
        "S_PrinterID", "S_StockID", "S_TechID", "S_Family", "S_Measure", _
        "S_UnitCost", "S_StockWidth_mm", "S_SheetHeight_mm", "S_ConsRate", _
        "S_StampedAt", "S_StampedBy", "S_SchemaVer")

    UnlockSheet ws
    For i = 1 To UBound(order) - LBound(order) + 1
        want = CStr(order(LBound(order) + i - 1))
        have = lo.ListColumns(i).Name
        If StrComp(have, want, vbTextCompare) <> 0 Then
            ' Cut+insert scoped to the table's own range (ListColumn.Range is
            ' header+data only, never the full column) - rows 1-11 above the
            ' table are never touched by this, whichever column is moving.
            lo.ListColumns(want).Range.Cut
            lo.ListColumns(i).Range.Insert Shift:=xlToRight
        End If
    Next i
    RelockSheet ws
End Sub

Private Function CountButtons() As Long
    Dim ws As Worksheet, i As Long, t As Long
    For Each ws In ThisWorkbook.Worksheets
        For i = 1 To ws.Buttons.Count
            If Left$(ws.Buttons(i).Name, Len(BTN_TAG)) = BTN_TAG Then t = t + 1
        Next i
    Next ws
    CountButtons = t
End Function

Public Sub DrawOne(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal ColNo As Long, ByVal Caption As String, ByVal Macro As String, ByVal W As Single)
    Dim b As Button, c As Range
    Set c = ws.Cells(RowNo, ColNo)
    Set b = ws.Buttons.Add(c.Left, c.Top, W, 22)
    SetButtonName b, BTN_TAG & Macro & "_" & RowNo & "_" & ColNo
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
    UnlockTableBody "tblPaperFamilies"
    UnlockTableBody "tblStandardSizes"
    UnlockTableBody "tblConsumables"
    UnlockSettingsValues
End Sub

Private Sub UnlockTableBody(ByVal TableName As String)
    Dim lo As ListObject
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    UnlockSheet lo.Parent
    lo.DataBodyRange.Locked = False
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
    RelocateButton ws, lo, "btnNow", "Student No"
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
        For side = -1 To 1 Step 2
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
' Date/Time and Job ID get a fixed 100px rather than a cap - both hold
' content that genuinely needs it (a full "dd/mm/yyyy hh:mm" stamp, or a
' multi-digit correlation ID), so there is no narrower "good enough" to allow.
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
    SetJobColWidth lo, "Date/Time", ColWidthForPx(100)
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

' -------------------------------------------------------- column groups ---
' Snag list item 14: the calculated cost columns collapse together, and so do
' the S_ snapshot columns - neither is referenced day to day, and grouping
' lets a user hide the detail without hiding the columns outright.
Private Sub GroupJobColumns(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    ' Clear whatever grouping already exists on the old (wrong) span first -
    ' re-running this against a workbook built before the 2026-09-22 fix
    ' would otherwise leave Chargeable Cost nested in both the old and new
    ' outline. Harmless no-op on a workbook that never had the old group.
    UngroupColumnRange lo, "Paper Cost", "Chargeable Cost"

    ' Paper Cost, Consumable Cost, Gross Cost and Disregarded collapse
    ' together; Chargeable Cost and Paid stay outside the group and always
    ' visible even when it's collapsed (2026-09-22 snag list item 1d - the
    ' previous range ran one column too far, to Chargeable Cost itself,
    ' which hid the one cost figure a collapsed view most needs to keep
    ' showing).
    GroupColumnRange lo, "Paper Cost", "Disregarded"

    ' Notes through S_SchemaVer: flattened to NO grouping at all (2026-09-22),
    ' not re-grouped - this used to also GroupColumnRange S_PrinterID through
    ' S_SchemaVer, but PrintCosts.xlsx already ships that whole span
    ' (H_Issues onward) pre-grouped and hidden, so the explicit re-group
    ' nested a SECOND outline level on top of it - one clean lvl-2 span
    ' became a stray lvl-2 group over Notes+H_Issues (Notes wrongly pulled
    ' in and hidden - see below) plus a separate lvl-3 group over most of
    ' the snapshot block, two extra collapsible groups cluttering the
    ' outline pane right next to the one cost group that matters day to
    ' day. Flattening removes the outline controls entirely; H_Issues and
    ' the snapshot columns stay invisible regardless, via their own
    ' .Hidden state (§3.4/§5), which needs no outline group to hold it.
    FlattenOutline lo, "Notes", "S_SchemaVer"
End Sub

Private Sub GroupColumnRange(ByVal lo As ListObject, ByVal FirstHeader As String, ByVal LastHeader As String)
    Dim c1 As Long, c2 As Long, ws As Worksheet
    c1 = ColIdx(lo, FirstHeader)
    c2 = ColIdx(lo, LastHeader)
    Set ws = lo.Parent
    UnlockSheet ws
    ws.Range(lo.HeaderRowRange.Cells(1, c1), lo.HeaderRowRange.Cells(1, c2)).EntireColumn.Group
    RelockSheet ws
End Sub

Private Sub UngroupColumnRange(ByVal lo As ListObject, ByVal FirstHeader As String, ByVal LastHeader As String)
    Dim c1 As Long, c2 As Long, ws As Worksheet
    c1 = ColIdx(lo, FirstHeader)
    c2 = ColIdx(lo, LastHeader)
    Set ws = lo.Parent
    UnlockSheet ws
    ' Ungroup raises 1004 outright when the range was never grouped (the
    ' ordinary case on a fresh build, which never had the old wider group to
    ' begin with) - expected and harmless, so this is the one place a bare
    ' On Error Resume Next is warranted rather than a real failure to report.
    On Error Resume Next
    ws.Range(lo.HeaderRowRange.Cells(1, c1), lo.HeaderRowRange.Cells(1, c2)).EntireColumn.Ungroup
    On Error GoTo 0
    RelockSheet ws
End Sub

' Sets OutlineLevel back to 1 (no grouping at all) across the given span,
' regardless of what it was before - a single deterministic reset rather
' than a bare .Ungroup, which only removes one level at a time and would
' leave a doubly-nested span (2026-09-22's bug, see GroupJobColumns) still
' one level deep. Also resets Notes specifically back to visible: it's a
' genuine input column, not part of the historical/snapshot block, and can
' end up swept into H_Issues/the snapshot block's hidden state as a side
' effect of where Paid gets inserted right next to it (EnsurePaidColumn).
' Every OTHER column in the span keeps whatever .Hidden state it already
' has - H_Issues and the snapshot columns ship hidden in PrintCosts.xlsx
' (§3.4/§5) and that is untouched here, on purpose.
Private Sub FlattenOutline(ByVal lo As ListObject, ByVal FirstHeader As String, ByVal LastHeader As String)
    Dim c1 As Long, c2 As Long, ws As Worksheet
    c1 = ColIdx(lo, FirstHeader)
    c2 = ColIdx(lo, LastHeader)
    Set ws = lo.Parent
    UnlockSheet ws
    ws.Range(lo.HeaderRowRange.Cells(1, c1), lo.HeaderRowRange.Cells(1, c2)).EntireColumn.OutlineLevel = 1
    lo.ListColumns(FirstHeader).Range.EntireColumn.Hidden = False
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
    RelockSheet ws
End Sub
