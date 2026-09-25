Attribute VB_Name = "modRegistry"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Location discovery, and the consolidated job range the reports read.
'
' Spec 17 lets a non-technical user create a print room by duplicating a
' worksheet, which makes duplication the most fragile thing in the workbook.
' Excel silently renames the copied table to tblJobs_MAIN1, the copied buttons
' can acquire a workbook-file prefix on their OnAction, the new sheet arrives
' unprotected, and it still holds the original's records.
'
' RefreshLocations is the repair pass for all of that. It is safe to run at any
' time, and it is the only thing a user has to do after duplicating a sheet.
'
' Two departures from the design document, both forced by what was already
' built:
'
'   - The consolidated range lives on its own sheet, _Data, not on _Work.
'     modLists claims _Work column by column for validation staging and will
'     walk over anything else parked there.
'
'   - A sheet is recognised by its marker CELL, not by a sheet-scoped
'     LOC_Marker name, matching modUtils.IsLocation. One moving part, and
'     Excel cannot quietly rename it.

Private Const REG_SHEET As String = "_Registry"
Private Const DATA_SHEET As String = "_Data"
Private Const AUDIT_SHEET As String = "_Audit"
Private Const REG_TABLE As String = "tblLocations"
Private Const AUDIT_TABLE As String = "tblAudit"
' The consolidated range's column span, bounded by NAME rather than
' position: "every column from FIRST_JOB_COL to LAST_JOB_COL, inclusive" via
' a structured reference (WriteConsolidated/WriteHeaders below). FIRST_JOB_COL
' was "Job ID" until 2026-09-22, when modInit.ReorderJobColumns moved Status
' and Job ID away from the front of the table (columns 1-2) to fix a
' reduced-clutter-view bug (they used to collide with the location config
' block, which also lives in columns A/B - see modInit's own comment).
' Job ID is no longer the leftmost real column, so the bound had to move to
' whatever IS - Date/Time - or the span would have started mid-table and
' silently dropped everything before it. Status and Job ID are still
' INSIDE the span either way, since they now sit between Date/Time and
' Notes rather than starting it.
Private Const FIRST_JOB_COL As String = "Date/Time"
Private Const LAST_JOB_COL As String = "Notes"
Private Const HDR_ROW As Long = 9
Private Const DATA_ROW As Long = 10
Private Const CODE_MAX As Long = 8

' ---------------------------------------------------------------- refresh ---
Public Sub RefreshLocations()
    Dim sheets As Collection, codes As clsDict, wasThere As clsDict
    Dim ws As Worksheet, v As Variant
    Dim problems As String, gone As String, n As Long

    On Error GoTo Fail
    AppOff
    EnsureSystemSheets

    Set wasThere = RegisteredSheets()

    ' Every jobs table is parked on a temporary name first. Renaming straight
    ' to the final name fails whenever the name is still held by a table this
    ' pass has not reached yet - which is exactly the case after a duplication.
    ParkJobTables

    Set sheets = New Collection
    Set codes = New clsDict
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then sheets.Add ws
    Next ws

    ' Workbook order matters: Excel inserts a duplicate immediately after its
    ' original, so the original is reached first and keeps the code. The copy
    ' is the one that gets a new one, which is what you want - the original's
    ' job IDs already carry it.
    For Each v In sheets
        Set ws = v
        UnlockSheet ws
        problems = problems & AssignCode(ws, codes)
        NameJobTable ws, CStr(codes.Item("#" & ws.Name))
        FixButtons ws
        EnsureJobDefaults ws
        EnsureRollUnitSetting ws
        EnsurePrintersDisplay ws
        EnsureJobCountDisplay ws
        BindColumns ws
        ApplyReducedView ws
        ProtectSheet ws
        n = n + 1
    Next v

    gone = MissingSheets(wasThere, codes)
    WriteRegistry sheets, codes
    WriteConsolidated sheets, codes

    ' The Reports page's Technician/Printer/Paper Stock dropdowns are a VBA
    ' snapshot of _Data's distinct values, not a live formula (a literal
    ' Formula1 list is capped at 255 characters), so they need refreshing
    ' here too - the same reason BindColumns refreshes each location's own
    ' dropdowns on every call rather than only once at setup.
    Dim repWs As Worksheet
    Set repWs = ReportsSheet()
    If Not repWs Is Nothing Then RefreshReportFilterLists repWs

    ' After WriteRegistry, because the status is read out of the registry.
    For Each v In sheets
        Set ws = v
        RefreshExportStatus ws
        If NeedsExport(ws) Then
            problems = problems & "- " & LocValue(ws, "LOC_Name") & ": " & ExportStatusText(ws) & vbCrLf
        End If
    Next v

    Invalidate
    SetSetting "LASTREF", Now
    AppOn

    Announce n, gone, problems
    Exit Sub
Fail:
    AppReset
    ReportError "Refresh Locations"
End Sub

' ------------------------------------------------------------------ codes ---
' The dictionary holds both directions: "<CODE>" -> sheet name, so a clash can
' be spotted, and "#<sheet name>" -> code, so the second pass can look the
' code back up without re-deriving it.
Private Function AssignCode(ByVal ws As Worksheet, ByVal codes As clsDict) As String
    Dim r As Range, want As String, code As String

    Set r = LocRange(ws, "LOC_Code")
    If r Is Nothing Then
        ' Job IDs are built from this cell, so a sheet without it cannot number
        ' its records. Worth naming rather than silently working around.
        AssignCode = "- '" & ws.Name & "' has no LOC_Code cell, so new job IDs on it will read LOC." & vbCrLf
        want = ""
    Else
        want = CleanCode(CStr(r.Value))
    End If

    If Len(want) = 0 Then want = CleanCode(ws.Name)
    If Len(want) = 0 Then want = "LOC"

    code = Uniquify(want, codes)
    codes.Add code, ws.Name
    codes.Add "#" & ws.Name, code

    ' Only written when it actually differs, so a refresh does not dirty the
    ' workbook for no reason.
    If Not r Is Nothing Then
        If StrComp(CStr(r.Value), code, vbBinaryCompare) <> 0 Then r.Value = code
    End If
End Function

Private Function CleanCode(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
        If Len(out) >= CODE_MAX Then Exit For
    Next i
    CleanCode = out
End Function

Private Function Uniquify(ByVal want As String, ByVal codes As clsDict) As String
    Dim n As Long, try As String
    If Not codes.Exists(want) Then
        Uniquify = want
        Exit Function
    End If
    For n = 2 To 99
        try = Left$(want, CODE_MAX - Len(CStr(n))) & CStr(n)
        If Not codes.Exists(try) Then
            Uniquify = try
            Exit Function
        End If
    Next n
    Uniquify = Left$(want, CODE_MAX - 6) & Format$(Timer * 100 Mod 100000, "00000")
End Function

' ------------------------------------------------------------- job tables ---
Private Sub ParkJobTables()
    Dim ws As Worksheet, lo As ListObject, n As Long
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            Set lo = JobsTable(ws)
            If Not lo Is Nothing Then
                n = n + 1
                UnlockSheet ws
                lo.Name = JOBS_PREFIX & "PARK" & Format$(n, "000")
            End If
        End If
    Next ws
End Sub

Private Sub NameJobTable(ByVal ws As Worksheet, ByVal code As String)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    lo.Name = JOBS_PREFIX & code
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------- buttons ---
' A copied Form Control can come back with its OnAction qualified by the
' workbook file name, which stops working the moment the file is renamed. The
' macro name is recovered from the button's own name - modInit writes it there
' as pcb_<macro>_<row>_<col> - and falls back to stripping the prefix.
Private Sub FixButtons(ByVal ws As Worksheet)
    Dim i As Long, b As Button, want As String
    For i = 1 To ws.Buttons.Count
        Set b = ws.Buttons(i)
        If Left$(b.Name, 4) = "pcb_" Then
            want = MacroFromName(b.Name)
            If Len(want) = 0 Then want = StripBook(CStr(b.OnAction))
            If Len(want) > 0 Then
                If StrComp(CStr(b.OnAction), want, vbBinaryCompare) <> 0 Then b.OnAction = want
            End If
        End If
    Next i
End Sub

' Excel re-qualifies every button's OnAction with the workbook file name when it
' saves a macro-enabled file, so the buttons in the shipped .xlsm read
' "PrintCosts.xlsm!btnAddPrintJob". That works until somebody renames the file,
' at which point every button reports a missing macro.
'
' Called from Workbook_Open. It compares the prefix against the file's current
' name and only rewrites when they disagree, so an untouched workbook is never
' marked dirty just by being opened - which would prompt every user to save a
' file they had only looked at.
Public Sub HealButtons()
    Dim ws As Worksheet, i As Long, b As Button, cur As String, want As String
    Dim mine As String
    mine = ThisWorkbook.Name & "!"

    For Each ws In ThisWorkbook.Worksheets
        For i = 1 To ws.Buttons.Count
            Set b = ws.Buttons(i)
            If Left$(b.Name, 4) = "pcb_" Then
                cur = CStr(b.OnAction)
                If InStr(cur, "!") > 0 Then
                    If InStr(1, cur, mine, vbTextCompare) <> 1 Then
                        want = MacroFromName(b.Name)
                        If Len(want) = 0 Then want = StripBook(cur)
                        If Len(want) > 0 Then b.OnAction = want
                    End If
                End If
            End If
        Next i
    Next ws
End Sub

Private Function MacroFromName(ByVal s As String) As String
    ' pcb_btnAddPrintJob_10_1 -> btnAddPrintJob
    Dim p As Long
    s = Mid$(s, 5)
    p = InStrRev(s, "_")
    If p > 0 Then s = Left$(s, p - 1)
    p = InStrRev(s, "_")
    If p > 0 Then s = Left$(s, p - 1)
    MacroFromName = s
End Function

Private Function StripBook(ByVal s As String) As String
    Dim p As Long
    p = InStrRev(s, "!")
    If p > 0 Then s = Mid$(s, p + 1)
    StripBook = Replace$(s, "'", "")
End Function

' ---------------------------------------------------------- system sheets ---
' Created here rather than shipped in the .xlsx. The workbook file is written
' outside Excel and has no generator script, so anything structural that can be
' built in VBA is built in VBA - it stays reproducible, and it works the same
' on Mac.
Public Sub EnsureSystemSheets()
    Dim ws As Worksheet

    ' All but _Data ship in the .xlsx already. These calls are here so the
    ' workbook can still be rebuilt if one is ever lost, and the headers match
    ' what the file ships exactly - inventing new ones silently orphans the
    ' table that is already there.
    Set ws = SheetOrNew(REG_SHEET)
    EnsureTable ws, REG_TABLE, Array("SheetName", "Code", "Name", "Department", "Rows", "First date", "Last date", "State")

    ' The shipped .xlsx predates the export feature, so these two are added to
    ' the existing table rather than being part of its creation.
    EnsureColumn REG_TABLE, "Last export"
    EnsureColumn REG_TABLE, "Export sig"

    ' Likewise predates the persisted Job ID high-water mark (see NextJobId).
    EnsureColumn REG_TABLE, "Job ID HWM"

    Set ws = SheetOrNew(AUDIT_SHEET)
    EnsureTable ws, AUDIT_TABLE, Array("When", "User", "Action", "Location", "Detail")

    Set ws = SheetOrNew(DATA_SHEET)
End Sub

Private Function SheetOrNew(ByVal Nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Nm)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = Nm
    End If

    UnlockSheet ws
    ' xlSheetVeryHidden keeps it off the tab context menu's Unhide list, so it
    ' cannot be brought into view and edited by accident.
    If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
    Set SheetOrNew = ws
End Function

Private Sub EnsureColumn(ByVal TableName As String, ByVal Header As String)
    Dim lo As ListObject, i As Long, lc As ListColumn
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, Header, vbTextCompare) = 0 Then Exit Sub
    Next i
    UnlockSheet lo.Parent
    Set lc = lo.ListColumns.Add
    lc.Name = Header
    lo.Range.Columns.AutoFit
    RelockSheet lo.Parent
End Sub

Private Sub EnsureTable(ByVal ws As Worksheet, ByVal TableName As String, ByVal Headers As Variant)
    Dim lo As ListObject, i As Long
    If Not Tbl(TableName) Is Nothing Then Exit Sub

    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(1, i - LBound(Headers) + 1).Value = Headers(i)
    Next i
    ws.Rows(1).Font.Bold = True

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(Headers) - LBound(Headers) + 1)), , xlYes)
    lo.Name = TableName
End Sub

' ---------------------------------------------------------------- registry ---
Private Function RegisteredSheets() As clsDict
    Dim lo As ListObject, i As Long, d As clsDict, s As String
    Set d = New clsDict
    Set lo = Tbl(REG_TABLE)
    If lo Is Nothing Then
        Set RegisteredSheets = d
        Exit Function
    End If
    For i = 1 To lo.ListRows.Count
        s = Trim$(CStr(CellIn(lo, i, "SheetName").Value))
        If Len(s) > 0 Then d.Add s, CStr(CellIn(lo, i, "Name").Value)
    Next i
    Set RegisteredSheets = d
End Function

' Excel raises no event when a sheet is deleted, so a vanished print room can
' only be noticed after the fact. Saying what was lost, by name and record
' count, is the difference between an explanation and a silent hole.
Private Function MissingSheets(ByVal wasThere As clsDict, ByVal codes As clsDict) As String
    Dim k As Variant, out As String
    For Each k In wasThere.Keys
        If Not codes.Exists("#" & CStr(k)) Then
            out = out & "- '" & CStr(wasThere.Item(CStr(k))) & "' (sheet " & CStr(k) & ") is no longer in the workbook. Its records went with it." & vbCrLf
        End If
    Next k
    MissingSheets = out
End Function

Private Sub WriteRegistry(ByVal sheets As Collection, ByVal codes As clsDict)
    Dim lo As ListObject, ws As Worksheet, v As Variant, r As ListRow
    Dim i As Long, lo_jobs As ListObject, n As Long, fd As Double, ld As Double

    Set lo = Tbl(REG_TABLE)
    If lo Is Nothing Then Exit Sub
    UnlockSheet lo.Parent

    ' The export stamp has to survive a refresh, and a refresh rebuilds every
    ' row from scratch. Captured by sheet name first, put back below.
    '
    ' A sheet that has been renamed loses its stamp and reads as never
    ' exported. That is the conservative direction and is left deliberately:
    ' warning about an export that was in fact done costs a spare file, while
    ' the reverse costs the records.
    ' The Job ID high-water mark has to survive a refresh for the same reason
    ' and by the same means - it is what NextJobId reads, and it must never
    ' move backwards (see NextJobId).
    Dim keepWhen As clsDict, keepSig As clsDict, keepHwm As clsDict, k As String
    Set keepWhen = New clsDict
    Set keepSig = New clsDict
    Set keepHwm = New clsDict
    For i = 1 To lo.ListRows.Count
        k = Trim$(CStr(CellIn(lo, i, "SheetName").Value))
        If Len(k) > 0 Then
            keepWhen.Add k, CellIn(lo, i, "Last export").Value
            keepSig.Add k, CStr(CellIn(lo, i, "Export sig").Value)
            keepHwm.Add k, NumOf(CellIn(lo, i, "Job ID HWM"))
        End If
    Next i

    For i = lo.ListRows.Count To 1 Step -1
        lo.ListRows(i).Delete
    Next i

    For Each v In sheets
        Set ws = v
        Set lo_jobs = JobsTable(ws)
        n = 0: fd = 0: ld = 0
        If Not lo_jobs Is Nothing Then SpanOf lo_jobs, n, fd, ld

        Set r = lo.ListRows.Add
        r.Range.Cells(1, 1).Value = ws.Name
        r.Range.Cells(1, 2).Value = CStr(codes.Item("#" & ws.Name))
        r.Range.Cells(1, 3).Value = LocValue(ws, "LOC_Name")
        r.Range.Cells(1, 4).Value = LocValue(ws, "LOC_Dept")
        r.Range.Cells(1, 5).Value = n
        If fd > 0 Then r.Range.Cells(1, 6).Value = CDate(fd)
        If ld > 0 Then r.Range.Cells(1, 7).Value = CDate(ld)
        r.Range.Cells(1, 8).Value = "Active"
        If keepSig.Exists(ws.Name) Then
            CellIn(lo, r.Index, "Last export").Value = keepWhen.Item(ws.Name)
            CellIn(lo, r.Index, "Last export").NumberFormat = "dd/mm/yyyy hh:mm"
            CellIn(lo, r.Index, "Export sig").Value = CStr(keepSig.Item(ws.Name))
        End If
        If keepHwm.Exists(ws.Name) Then
            CellIn(lo, r.Index, "Job ID HWM").Value = CDbl(keepHwm.Item(ws.Name))
        End If
    Next v

    ' DataBodyRange is Nothing when the table has no rows, which is what a
    ' workbook with no print rooms leaves behind: every registry row is deleted
    ' above and the loop adds none back. Unguarded, the two NumberFormat lines
    ' raise 91 and take RefreshLocations to its error handler - in exactly the
    ' design 3.3 case where it most needs to finish and name what went missing.
    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns("First date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        lo.ListColumns("Last date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    End If
    lo.Range.Columns.AutoFit

    RelockSheet lo.Parent
End Sub

Private Sub SpanOf(ByVal lo As ListObject, ByRef n As Long, ByRef fd As Double, ByRef ld As Double)
    Dim i As Long, d As Double
    n = RowCount(lo)
    For i = 1 To n
        d = DateSerialOf(CellIn(lo, i, "Date/Time"))
        If d > 0 Then
            If fd = 0 Or d < fd Then fd = d
            If d > ld Then ld = d
        End If
    Next i
End Sub

' ------------------------------------------------------ consolidated range ---
' Every print room's records stacked into one spill range, each row tagged with
' its location code. This one formula is VBA's entire involvement in reporting:
' the Summary and Reports sheets are live formulas over it, so there
' is nothing to refresh and nothing that can go stale.
'
' Shape: Location, then Job ID .. Notes as the job table lays them out.
Private Sub WriteConsolidated(ByVal sheets As Collection, ByVal codes As clsDict)
    Dim dws As Worksheet, ws As Worksheet, v As Variant
    Dim lo As ListObject, blocks As String, t As String, code As String

    Set dws = ThisWorkbook.Worksheets(DATA_SHEET)
    UnlockSheet dws
    dws.Cells.Clear
    dws.Cells(1, 1).Value = "Consolidated print jobs. Written by Refresh Locations - do not edit."

    For Each v In sheets
        Set ws = v
        Set lo = JobsTable(ws)
        If Not lo Is Nothing Then
            code = CStr(codes.Item("#" & ws.Name))
            t = lo.Name
            If Len(blocks) > 0 Then blocks = blocks & ","
            blocks = blocks & "HSTACK(IF(SEQUENCE(ROWS(" & t & "[" & FIRST_JOB_COL & "]))>0," & _
                     """" & code & """)," & t & "[[" & FIRST_JOB_COL & "]:[" & LAST_JOB_COL & "]])"
        End If
    Next v

    If Len(blocks) = 0 Then
        dws.Cells(HDR_ROW, 1).Value = "Location"
        RelockSheet dws
        Exit Sub
    End If

    WriteHeaders dws, JobsTable(sheets(1))

    ' Blank table rows are filtered out here so no report has to know that an
    ' empty print room still holds one placeholder row. A deleted sheet turns
    ' the whole expression into an error, which IFERROR degrades to empty
    ' rather than letting #REF! cascade through every report; the next refresh
    ' rebuilds it and names what went missing.
    '
    ' .Formula2 is required. Writing a dynamic array through .Formula applies
    ' implicit intersection and silently stores a single value.
    dws.Cells(DATA_ROW, 1).Formula2 = "=LET(raw,VSTACK(" & blocks & _
        "),IFERROR(FILTER(raw,INDEX(raw,,2)<>""""),""""))"

    RelockSheet dws
End Sub

Private Sub WriteHeaders(ByVal dws As Worksheet, ByVal lo As ListObject)
    Dim i As Long, c1 As Long, c2 As Long
    dws.Cells(HDR_ROW, 1).Value = "Location"
    If lo Is Nothing Then Exit Sub
    c1 = ColIdx(lo, FIRST_JOB_COL)
    c2 = ColIdx(lo, LAST_JOB_COL)
    For i = c1 To c2
        dws.Cells(HDR_ROW, i - c1 + 2).Value = lo.HeaderRowRange.Cells(1, i).Value
    Next i
    dws.Rows(HDR_ROW).Font.Bold = True
End Sub

' --------------------------------------------------------------- reporting ---
Private Sub Announce(ByVal n As Long, ByVal gone As String, ByVal problems As String)
    Dim what As String, why As String

    If n = 0 Then
        Say "No print rooms were found.", _
            "A sheet counts as a print room when cell " & MARKER_CELL & " on it reads " & MARKER & ".", _
            "Column " & Left$(MARKER_CELL, 2) & " is hidden, so unhide it to check."
        Exit Sub
    End If

    what = n & " print room" & IIf(n = 1, "", "s") & " registered."
    why = "Job tables renamed to match, buttons re-pointed, dropdowns rebuilt, sheets protected, and the reporting range rewritten."
    If gButtonsDrawn > 0 Then
        why = gButtonsDrawn & " buttons drawn." & vbCrLf & why
        gButtonsDrawn = 0
    End If
    If Len(gone) > 0 Then why = why & vbCrLf & vbCrLf & "Gone since the last refresh:" & vbCrLf & gone
    If Len(problems) > 0 Then why = why & vbCrLf & "Worth looking at:" & vbCrLf & problems

    Say what, why, "Run this again after duplicating a sheet to create a new print room."
End Sub

' ------------------------------------------------------------------ lookup ---
' The registry is the only place that knows which sheets are print rooms
' without walking the workbook, so reports and the export block read it here.
Public Function LocationSheets() As Collection
    Dim out As Collection, ws As Worksheet
    Set out = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then out.Add ws
    Next ws
    Set LocationSheets = out
End Function

' Every print room currently registered (tblLocations), for the Reports
' page's Location filter (modReports.RefreshReportFilterLists) - the same
' "every catalogue entry, not just what has been used" behaviour as
' modCatalog.AllActivePrinters/AllActiveStocks, so a room added today but not
' yet logged against is still choosable here.
Public Function AllLocationCodes() As Collection
    Dim out As New Collection, lo As ListObject, i As Long, code As String
    Set lo = Tbl(REG_TABLE)
    If lo Is Nothing Then
        Set AllLocationCodes = out
        Exit Function
    End If
    For i = 1 To lo.ListRows.Count
        code = Trim$(CStr(CellIn(lo, i, "Code").Value))
        If Len(code) > 0 Then out.Add code
    Next i
    Set AllLocationCodes = SortedTextCollection(out)
End Function

Public Function ConsolidatedRange() As Range
    Dim dws As Worksheet
    On Error Resume Next
    Set dws = ThisWorkbook.Worksheets(DATA_SHEET)
    If dws Is Nothing Then Exit Function
    Set ConsolidatedRange = dws.Cells(DATA_ROW, 1).SpillingToRange
End Function

' ------------------------------------------------------------ job ID hwm ---
' The persisted Job ID high-water mark, one per location, and the allocator
' that reads it.
'
' Before this, IDs were allocated from a scan of rows currently on the sheet.
' Deleting the highest-numbered job then silently reissued its ID, breaking
' the "globally unique, never re-issued" invariant export, import and any
' future aggregation all depend on. The fix: allocate from a value that only
' ever increases, never from what happens to be on the sheet right now.
'
' The scan below still runs, but only as a FLOOR under the persisted value,
' never as the source of truth - Max(persisted, scan) cannot move backwards
' when a row is deleted, because deleting only lowers what the scan finds,
' never the persisted figure. It earns its keep in two cases the persisted
' value alone cannot cover: seeding it the first time this column exists on
' an already-populated sheet, and raising it after an import has just written
' rows under this exact prefix with higher numbers than anything allocated
' locally so far (see modImport.ApplyImport).
Public Function NextJobId(ByVal ws As Worksheet, ByVal lo As ListObject) As String
    Dim site As String, code As String, prefix As String, hi As Long
    site = SettingText("SITE_ID", "SITE")
    code = LocValue(ws, "LOC_Code")
    If Len(code) = 0 Then code = "LOC"
    prefix = site & "-" & code & "-"

    hi = JobIdHWM(ws)
    hi = CLng(Application.WorksheetFunction.Max(hi, ScanMaxSuffix(lo, "Job ID", prefix)))
    hi = hi + 1
    SetJobIdHWM ws, hi
    NextJobId = prefix & Format$(hi, "00000")
End Function

' The safety-net scan NextJobId takes a floor from. Digits only, so Val() is
' safe here - the locale trap modUtils.NumOf guards against is specific to a
' fractional part, and a Job ID suffix never has one.
Private Function ScanMaxSuffix(ByVal lo As ListObject, ByVal IdHeader As String, ByVal Prefix As String) As Long
    Dim i As Long, n As Long, hi As Long, s As String
    For i = 1 To lo.ListRows.Count
        s = CStr(CellIn(lo, i, IdHeader).Value)
        If Len(s) > Len(Prefix) Then
            If StrComp(Left$(s, Len(Prefix)), Prefix, vbTextCompare) = 0 Then
                n = Val(Mid$(s, Len(Prefix) + 1))
                If n > hi Then hi = n
            End If
        End If
    Next i
    ScanMaxSuffix = hi
End Function

Private Function JobIdHWM(ByVal ws As Worksheet) As Long
    Dim lo As ListObject, i As Long
    Set lo = Tbl(REG_TABLE)
    If lo Is Nothing Then Exit Function
    i = RegRowFor(lo, ws.Name)
    If i = 0 Then Exit Function
    JobIdHWM = CLng(NumOf(CellIn(lo, i, "Job ID HWM")))
End Function

Private Sub SetJobIdHWM(ByVal ws As Worksheet, ByVal Value As Long)
    Dim lo As ListObject, i As Long
    Set lo = Tbl(REG_TABLE)
    If lo Is Nothing Then Exit Sub
    i = RegRowFor(lo, ws.Name)
    If i = 0 Then Exit Sub
    UnlockSheet lo.Parent
    CellIn(lo, i, "Job ID HWM").Value = Value
    RelockSheet lo.Parent
End Sub

Private Function RegRowFor(ByVal lo As ListObject, ByVal SheetName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "SheetName").Value)), SheetName, vbTextCompare) = 0 Then
            RegRowFor = i
            Exit Function
        End If
    Next i
End Function

' Location code -> its sheet. Used by Reports-page delete (Problem 3) to map
' a visible row's Location back to the sheet holding its actual table row.
Public Function SheetForCode(ByVal Code As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In LocationSheets()
        If StrComp(LocValue(ws, "LOC_Code"), Code, vbTextCompare) = 0 Then
            Set SheetForCode = ws
            Exit Function
        End If
    Next ws
End Function
