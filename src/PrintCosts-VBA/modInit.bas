Attribute VB_Name = "modInit"
Option Explicit

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
            EnsureJobDefaults ws
            EnsurePaidColumn ws
            ReorderJobColumns ws
            BindColumns ws
            GroupJobColumns ws
            ApplyStatusFormat ws
            ApplyReducedView ws
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
    ' The placeholder labels sit on row 10, two columns apart. Buttons are drawn
    ' over them and the labels cleared.
    DrawOne ws, 10, 1, "Add Print Job", "btnAddPrintJob", 110
    DrawOne ws, 10, 3, "Now", "btnNow", 110
    DrawOne ws, 10, 5, "Remove Row", "btnRemoveRow", 110
    DrawOne ws, 10, 7, "Select printers...", "btnSelectPrinters", 110
    DrawOne ws, 10, 9, "Check this sheet", "btnCheckSheet", 110
    DrawOne ws, 10, 11, "Clear All", "btnClearAll", 110
    DrawOne ws, 10, 13, "Export...", "btnExport", 110
    DrawOne ws, 10, 15, "Import...", "btnImportLocation", 110

    Dim c As Long
    For c = 1 To 15
        With ws.Cells(10, c)
            .ClearContents
            .Interior.Pattern = xlNone
        End With
    Next c
    ws.Cells(10, 13).ClearContents

    ' Reduced-clutter view toggle (row 9, snag list item 1e) - column 7,
    ' over Unit (just past the batch defaults' own Printer/Paper Stock cells
    ' at columns A-F, §4.1) - not one of the columns the toggle itself can
    ' hide. Originally drawn at column 13, which happened to land on
    ' whichever column Disregard Consumable was sitting at - column 13 is
    ' safe now that ReorderJobColumns has moved Status/Job ID away from the
    ' front of the table, but the anchor is deliberately independent of that
    ' fix (a hidden-column collision here would take out the one button that
    ' undoes it). Caption read fresh from the current setting each time this
    ' runs, same as ConfigToggleCaption's button does.
    DrawOne ws, 9, 7, ReducedViewCaption(), "btnToggleReducedView", 130

    ' Clear defaults - to the right of the toggle above, clear of both.
    DrawOne ws, 9, 9, "Clear defaults", "btnClearDefaults", 110
End Sub

' Yes/No validation on the two location defaults (spec 9.2). These seed each
' new print job and are plain cells in the workbook file, so the list is added
' here rather than shipped with it.
Private Sub ConfigValidation(ByVal ws As Worksheet)
    AddYesNo LocRange(ws, "LOC_DefDisPaper"), "Disregard paper cost", "Sets what new print jobs on this sheet start with. Changing it never alters jobs already recorded."
    AddYesNo LocRange(ws, "LOC_DefDisCons"), "Disregard consumable cost", "Sets what new print jobs on this sheet start with. Changing it never alters jobs already recorded."
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

' ----------------------------------------------------- batch defaults ---
' Snag list item 1b: three cells above the toolbar (row 9 - the config block
' ends at row 8, the toolbar starts at row 10) let a technician set a
' Technician/Printer/Paper Stock once and have every subsequently added job
' pre-filled from them (modJobs.AddPrintJob), until Clear defaults empties
' them again (modJobs.ClearDefaults). Self-provisioned here rather than
' shipped in the .xlsx - same reasoning as LOC_Export (modExport): a
' sheet-scoped name copies cleanly with a duplicated sheet, and re-adding it
' every run means an older location sheet picks the feature up without hand
' surgery.
' Public: modRegistry.RefreshLocations also calls this for every location on
' every refresh, alongside BindColumns - the same "rebuild dependent
' dropdowns, self-heal an old sheet, re-point a duplicated one" reasoning
' RefreshExportStatus/EnsureExportName already applies to LOC_Export.
Public Sub EnsureJobDefaults(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Range("A9").Value = "Default: technician"
    ws.Range("A9").Font.Bold = True
    ws.Range("C9").Value = "Default: printer"
    ws.Range("C9").Font.Bold = True
    ws.Range("E9").Value = "Default: paper"
    ws.Range("E9").Font.Bold = True

    EnsureLocName ws, "LOC_DefTech", "$B$9"
    EnsureLocName ws, "LOC_DefPrinter", "$D$9"
    EnsureLocName ws, "LOC_DefPaper", "$F$9"

    StyleInputCell ws.Range("B9")
    StyleInputCell ws.Range("D9")
    StyleInputCell ws.Range("F9")
    RelockSheet ws

    ' Both directions of spec 1a's filtering apply here too (spec 1b: "these
    ' selectors should implement the same filtering and autofill principles
    ' as the table cells").
    BindDefaultCells ws
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
Public Sub ReorderJobColumns(ByVal ws As Worksheet)
    Dim lo As ListObject, order As Variant, i As Long, want As String, have As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    order = Array( _
        "Date/Time", "Student Name", "Student No", "Technician", "Printer", _
        "Paper Stock", "Unit", "Quantity", "Print Width mm", "Disregard Paper", _
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
