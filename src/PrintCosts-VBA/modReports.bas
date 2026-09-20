Attribute VB_Name = "modReports"
Option Explicit

' Builds the Summary and Reports sheets.
'
' This module writes LAYOUT AND FORMULAS ONCE. It does no reporting: every
' figure on both sheets is a live worksheet formula over the consolidated
' range on _Data, so the reports cannot go stale and there is no refresh
' button to forget. The design document's claim that "modReports does not
' exist" was half right - what does not exist is a VBA reporting engine.
'
' Both builders are idempotent. They are re-run by InitialiseWorkbook, which
' is also what picks up a change to the job table's columns.
'
' Columns are addressed by HEADER NAME, never by position:
'
'     INDEX(_Data!$A$10#,,MATCH("Quantity",_Data!$A$9:$AZ$9,0))
'
' It is wordier than INDEX(...,,10) and it is the reason the job table's
' columns can be reordered without touching a single report formula.

Private Const SHEET_SUM As String = "Summary"
Private Const SHEET_REPORTS As String = "Reports"
Private Const DATA_SPILL As String = "_Data!$A$10#"
Private Const DATA_HDR As String = "_Data!$A$9:$AZ$9"
Private Const EXPORT_SIG_CELL As String = "AN1"
Private Const EXPORT_WHEN_CELL As String = "AN2"

' A column of the consolidated range, found by its header text.
Private Function C(ByVal Header As String) As String
    C = "INDEX(" & DATA_SPILL & ",,MATCH(""" & Header & """," & DATA_HDR & ",0))"
End Function

' SUMIFS of one column, grouped by the Location/Printer/Paper stock key. The
' criteria are arrays, which is what makes the result spill down alongside the
' keys instead of needing a formula per row.
Private Function SumBy(ByVal Header As String) As String
    SumBy = "SUMIFS(" & C(Header) & "," & C("Location") & ",lo," & C("Printer") & ",pr," & C("Paper Stock") & ",st)"
End Function

Public Sub BuildReportSheets()
    BuildSummary
    BuildReports
End Sub

' ============================================================== Summary ===
Public Sub BuildSummary()
    Dim ws As Worksheet, f As String

    Set ws = SheetNamed(SHEET_SUM, 1)
    UnlockSheet ws
    ws.Cells.Clear
    ClearSpillArea ws

    ws.Range("A1").Value = "Summary"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Formula = "=SET_ORG&"" - ""&SET_DEPT&""  |  ""&SET_SITE_NAME"
    ws.Range("A3").Value = "Every print job recorded in this workbook, by print room, printer and paper stock. " & _
        "Figures are live - they update as jobs are added and need no refreshing."

    ' --- totals, above the table ------------------------------------------
    ' Above, not below: the detail spills to an unknown height, so anything
    ' placed under it would be overwritten the moment a job is added.
    ws.Range("A5").Value = "Totals"
    ws.Range("A5").Font.Bold = True
    TotalCell ws, "B6", "Jobs", "=IFERROR(COUNTA(" & C("Job ID") & "),0)"
    TotalCell ws, "D6", "Gross", "=IFERROR(SUM(" & C("Gross Cost") & "),0)"
    TotalCell ws, "F6", "Disregarded", "=IFERROR(SUM(" & C("Disregarded") & "),0)"
    TotalCell ws, "H6", "Chargeable", "=IFERROR(SUM(" & C("Chargeable Cost") & "),0)"

    ' --- headers -----------------------------------------------------------
    WriteHeaderRow ws, 9, Array("Location", "Printer", "Paper stock", "Type", "Family", "Unit", _
                                "Jobs", "Quantity", "Area m2", "Paper cost", _
                                "Consumable cost", "Gross", "Disregarded", "Chargeable")

    ' --- the one formula -----------------------------------------------------
    ' Three-column key - Location, Printer, Paper Stock - rather than the
    ' two-column Location x Paper Stock key this replaces. Papers common to
    ' more than one printer model are deliberately not combined: two printers
    ' sharing a stock still get their own row each, because the cost per
    ' print differs by printer (consumable rate) even when the paper does not.
    f = "=IFERROR(LET(" & _
        "keys,SORT(UNIQUE(HSTACK(" & C("Location") & "," & C("Printer") & "," & C("Paper Stock") & "))),"
    f = f & "lo,INDEX(keys,,1),pr,INDEX(keys,,2),st,INDEX(keys,,3),"
    f = f & "HSTACK(lo,pr,st," & _
        "IFNA(XLOOKUP(st,tblPapers[Description],tblPapers[Paper type]),""(not in Papers)""),"
    f = f & "IFNA(XLOOKUP(st,tblPapers[Description],tblPapers[Family]),""(not in Papers)""),"
    f = f & "IFNA(IF(XLOOKUP(st,tblPapers[Description],tblPapers[Measure])=""Sheet"",""sheets"",""metres""),""-""),"
    f = f & "COUNTIFS(" & C("Location") & ",lo," & C("Printer") & ",pr," & C("Paper Stock") & ",st),"
    f = f & SumBy("Quantity") & "," & SumBy("Area m2") & "," & SumBy("Paper Cost") & ","
    f = f & SumBy("Consumable Cost") & "," & SumBy("Gross Cost") & ","
    f = f & SumBy("Disregarded") & "," & SumBy("Chargeable Cost") & ")),"
    f = f & """No print jobs have been recorded yet."")"
    ws.Range("A10").Formula2 = f

    FormatSummary ws
    RelockSheet ws
End Sub

Private Sub FormatSummary(ByVal ws As Worksheet)
    ws.Range("A9:N9").Interior.Color = RGB(222, 232, 244)
    ws.Range("H10:I2000").NumberFormat = "#,##0.00"
    ws.Range("J10:N2000").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("D6").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("F6").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("H6").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Columns("A:N").ColumnWidth = 14
    ws.Columns("A:C").ColumnWidth = 24
    ws.Rows(9).Font.Bold = True
End Sub

' ============================================================= Reports ===
' The criteria collapse to TRUE when blank, which is what gives the multi-
' case behaviour AT-12 tests without a single branch in the formula.
Private Function Criteria() As String
    Dim s As String

    ' The leading term exists to fix the SHAPE of the result, and without it
    ' the whole sheet breaks in its most ordinary state. A blank criterion
    ' makes IF(...) return a scalar TRUE, so with all four boxes empty the
    ' product is the scalar 1 - and FILTER(column, 1) is #CALC!, because the
    ' include argument has to be as tall as the data. Job ID is never blank in
    ' the consolidated range (7.1 filters those rows out), so this is an array
    ' of TRUE the right height, and every criterion below multiplies into it.
    s = "(" & C("Job ID") & "<>"""")"

    s = s & "*IF($C$5="""",TRUE,ISNUMBER(SEARCH($C$5," & C("Student Name") & ")))"
    s = s & "*IF($C$6="""",TRUE,TRIM(" & C("Student No") & "&"""")=TRIM($C$6&""""))"

    ' Dates are coerced with *1 on BOTH sides. A user typing 16/09/2026 into
    ' an unformatted cell can leave text there, and number >= text compares as
    ' FALSE for every row without raising anything - a filter that silently
    ' returns nothing. Text that will not coerce is treated as no filter at
    ' all rather than as a date of zero.
    '
    ' The column holds date AND time, so the end test is < end+1: otherwise a
    ' job logged at 16:30 on the closing date falls outside its own range.
    s = s & "*IF($C$7="""",TRUE,IF(ISERROR($C$7*1),TRUE," & _
        "IFERROR(" & C("Date/Time") & "*1,0)>=$C$7*1))"
    s = s & "*IF($C$8="""",TRUE,IF(ISERROR($C$8*1),TRUE," & _
        "IFERROR(" & C("Date/Time") & "*1,0)<$C$8*1+1))"

    ' Technician, Printer and Paper Stock are dropdowns (RefreshReportFilterLists),
    ' not free text, so an exact match is what "choose one from the list"
    ' means - unlike Student name, there is no fragment to search for.
    s = s & "*IF($G$5="""",TRUE," & C("Technician") & "=$G$5)"
    s = s & "*IF($G$6="""",TRUE," & C("Printer") & "=$G$6)"
    s = s & "*IF($G$7="""",TRUE," & C("Paper Stock") & "=$G$7)"

    ' Quantity: exact match, blank ignored, *1-coerced the same way the date
    ' boxes are so a value left as text by an unformatted cell is treated as
    ' no filter rather than as a quantity of zero.
    s = s & "*IF($G$8="""",TRUE,IF(ISERROR($G$8*1),TRUE," & _
        "IFERROR(" & C("Quantity") & "*1,0)=$G$8*1))"

    Criteria = s
End Function

' The sortable columns, in results-table order (Job ID excluded - it is a
' hidden correlation column, not something offered in the sort-by list).
Private Function ResultHeaders() As Variant
    ResultHeaders = Array("Date/Time", "Location", "Printer", "Paper stock", _
                          "Quantity", "Unit", "Area m2", "Paper cost", _
                          "Consumable cost", "Gross", "Disregarded", "Chargeable", _
                          "Technician", "Notes")
End Function

Private Function QuotedList(ByVal items As Variant) As String
    Dim i As Long, s As String
    For i = LBound(items) To UBound(items)
        If Len(s) > 0 Then s = s & ","
        s = s & """" & items(i) & """"
    Next i
    QuotedList = s
End Function

Public Sub BuildReports()
    Dim ws As Worksheet, f As String, ok As String, hdrs As Variant

    Set ws = SheetNamed(SHEET_REPORTS, 2)
    UnlockSheet ws

    ' The Export report stamp (Problem 2's refinement) has to survive this
    ' rebuild the same way tblLocations' Last export/Export sig survive
    ' RefreshLocations' registry rebuild - captured before Cells.Clear, put
    ' back once the sheet exists again.
    Dim savedSig As String, savedWhen As Variant
    savedSig = CStr(ws.Range(EXPORT_SIG_CELL).Value)
    savedWhen = ws.Range(EXPORT_WHEN_CELL).Value

    ws.Cells.Clear
    ok = Criteria
    hdrs = ResultHeaders

    ws.Range("A1").Value = "Reports"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Find and filter print jobs across every room in this workbook. Results update as you type - " & _
        "there is no search button. Leave a box empty to ignore it."

    CritCell ws, "A5", "C5", "Student name", "Part of a name is enough - ""Smith"" finds ""Jane Smith""."
    CritCell ws, "A6", "C6", "Student number", "Matched exactly."
    CritCell ws, "A7", "C7", "From date", "Pick a date, or leave blank for no start date."
    CritCell ws, "A8", "C8", "To date", "Jobs logged at any time on this date are included."
    ws.Range("C7:C8").NumberFormat = "dd/mm/yyyy"
    AddDateValidation ws.Range("C7"), "From date", "Leave blank for no start date."
    AddDateValidation ws.Range("C8"), "To date", "Jobs logged at any time on this date are included."

    CritCell ws, "E5", "G5", "Technician", "Choose from the list, or leave blank for all."
    CritCell ws, "E6", "G6", "Printer", "Choose from the list, or leave blank for all."
    CritCell ws, "E7", "G7", "Paper stock", "Choose from the list, or leave blank for all."
    CritCell ws, "E8", "G8", "Quantity", "Matched exactly."

    CritCell ws, "I5", "K5", "Sort by", "Leave blank for no sorting."
    CritCell ws, "I6", "K6", "Sort direction", "Ascending is the default."
    AddList ws.Range("K5"), QuotedList(hdrs), "Sort by", "Which column to sort the results by."
    AddList ws.Range("K6"), """Ascending"",""Descending""", "Sort direction", "Which way to sort."

    ' Spec 14.1: both criteria given, neither matching the other.
    ws.Range("A9").Formula2 = "=IF(OR($C$5="""",$C$6=""""),""""," & _
        "IF(IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)=0," & _
        """That name and that number do not appear together on any record - check both."",""""))"
    ws.Range("A9").Font.Color = RGB(176, 0, 32)
    ws.Range("A9").Font.Bold = True

    ' --- totals for the current selection ---------------------------------
    ws.Range("A11").Value = "Matching"
    ws.Range("A11").Font.Bold = True
    TotalCell ws, "B12", "Jobs", "=IFERROR(ROWS(FILTER(" & C("Job ID") & "," & ok & ")),0)"
    TotalCell ws, "D12", "Gross", "=IFERROR(SUM(FILTER(" & C("Gross Cost") & "," & ok & ")),0)"
    TotalCell ws, "F12", "Disregarded", "=IFERROR(SUM(FILTER(" & C("Disregarded") & "," & ok & ")),0)"
    TotalCell ws, "H12", "Chargeable", "=IFERROR(SUM(FILTER(" & C("Chargeable Cost") & "," & ok & ")),0)"
    ws.Range("D12").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("F12").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("H12").NumberFormat = ChrW(163) & "#,##0.00"

    ' --- the records ------------------------------------------------------
    ' Job ID is appended after Notes and hidden - the correlation key that
    ' maps a visible row back to its source location sheet and table row for
    ' the Reports-page delete. Nothing else moves, so existing column
    ' positions are untouched.
    WriteHeaderRow ws, 14, hdrs
    ws.Cells(14, UBound(hdrs) - LBound(hdrs) + 2).Value = "Job ID"

    ' Sorting: the whole FILTER result is bound to res once via LET, then
    ' re-ordered by whichever column K5 names, found by matching its header
    ' text against the same header list the table itself uses so the two
    ' can never drift apart. IFERROR around the SORTBY step is what keeps
    ' FILTER's own "no rows matched"/"no data at all" fallback text intact
    ' when there is nothing to sort - SORTBY on that text would otherwise
    ' error and the outer IFERROR would show the wrong one of the two
    ' messages.
    f = "=IFERROR(LET(" & _
        "res,FILTER(HSTACK(" & C("Date/Time") & "," & C("Location") & "," & C("Printer") & _
        "," & C("Paper Stock") & "," & C("Quantity") & "," & C("Unit") & "," & C("Area m2") & _
        "," & C("Paper Cost") & "," & C("Consumable Cost") & "," & C("Gross Cost") & _
        "," & C("Disregarded") & "," & C("Chargeable Cost") & "," & C("Technician") & _
        "," & C("Notes") & "," & C("Job ID") & ")," & ok & ",""No print jobs match those criteria.""),"
    f = f & "hdrs,{" & QuotedList(hdrs) & "},"
    f = f & "sortIdx,IFERROR(MATCH($K$5,hdrs,0),0),"
    f = f & "dir,IF($K$6=""Descending"",-1,1),"
    f = f & "IF(sortIdx=0,res,IFERROR(SORTBY(res,INDEX(res,,sortIdx),dir),res))"
    f = f & "),""No print jobs have been recorded yet."")"
    ws.Range("A15").Formula2 = f
    ws.Columns(UBound(hdrs) - LBound(hdrs) + 2).Hidden = True

    ws.Range(EXPORT_SIG_CELL).Value = savedSig
    If IsDate(savedWhen) Then
        ws.Range(EXPORT_WHEN_CELL).Value = savedWhen
        ws.Range(EXPORT_WHEN_CELL).NumberFormat = "dd/mm/yyyy hh:mm"
    End If
    ws.Range(EXPORT_SIG_CELL).EntireColumn.Hidden = True

    BuildBreakdowns ws, ok
    FormatReports ws

    ' Last, deliberately. ApplyTo (called from RefreshReportFilterLists) does
    ' its own Unlock/RelockSheet on ws as a self-contained operation - called
    ' any earlier, its RelockSheet would re-protect the sheet mid-build and
    ' every Validation.Add after it (the sort dropdowns above) would fail
    ' with a bare 1004 on the now-protected cells.
    RefreshReportFilterLists ws

    RelockSheet ws
End Sub

' Breakdowns sit to the RIGHT of the record list, not beneath it. The list
' spills to a height nobody can predict, so anything below it would be
' displaced the moment one more job matched.
Private Sub BuildBreakdowns(ByVal ws As Worksheet, ByVal ok As String)
    ws.Range("Q14").Value = "By print room"
    ws.Range("Q14").Font.Bold = True
    ws.Range("Q15").Formula2 = GroupFormula(ok, "Location")

    ws.Range("U14").Value = "By paper stock"
    ws.Range("U14").Font.Bold = True
    ws.Range("U15").Formula2 = GroupFormula(ok, "Paper Stock")

    ' Each block spills as key | Jobs | Gross | Chargeable, so the money
    ' columns are the third and fourth - Jobs is a count and must not be
    ' formatted as currency.
    ws.Range("R16:R2000").NumberFormat = "#,##0"
    ws.Range("S16:T2000").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("V16:V2000").NumberFormat = "#,##0"
    ws.Range("W16:X2000").NumberFormat = ChrW(163) & "#,##0.00"
End Sub

' Group the filtered records by one column. SUMIFS cannot be used here: its
' arguments must be ranges, and once FILTER has been applied these are arrays.
' BYROW with a LAMBDA does the per-key aggregation instead.
Private Function GroupFormula(ByVal ok As String, ByVal KeyHeader As String) As String
    Dim s As String
    ' The criteria are bound once as `ok` rather than pasted into each FILTER.
    ' Three copies of a 700-character expression is not just wasteful - it is
    ' three places for the four criteria to drift apart.
    s = "=IFERROR(LET(" & _
        "ok," & ok & "," & _
        "k,FILTER(" & C(KeyHeader) & ",ok)," & _
        "g,FILTER(" & C("Gross Cost") & ",ok)," & _
        "c,FILTER(" & C("Chargeable Cost") & ",ok)," & _
        "u,SORT(UNIQUE(k))," & _
        "VSTACK(HSTACK(""" & KeyHeader & """,""Jobs"",""Gross"",""Chargeable"")," & _
        "HSTACK(u," & _
        "BYROW(u,LAMBDA(x,SUM(--(k=x))))," & _
        "BYROW(u,LAMBDA(x,SUM(FILTER(g,k=x,0))))," & _
        "BYROW(u,LAMBDA(x,SUM(FILTER(c,k=x,0))))))" & _
        "),"""")"
    GroupFormula = s
End Function

Private Sub FormatReports(ByVal ws As Worksheet)
    ws.Range("A14:N14").Interior.Color = RGB(222, 232, 244)
    ws.Range("A15:A2000").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("E15:G2000").NumberFormat = "#,##0.00"
    ws.Range("H15:L2000").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Columns("A:N").ColumnWidth = 14
    ws.Columns("B:D").ColumnWidth = 22
    ws.Columns("N").ColumnWidth = 30
    ws.Columns("Q").ColumnWidth = 22
    ws.Columns("U").ColumnWidth = 22
    ws.Rows(14).Font.Bold = True
End Sub

' ======================================================== delete visible ===
' Problem 3: bulk pruning, deliberately confined to the Reports page and
' operating on whatever is currently visible under the active filters.
' modJobs.RemoveRow (single row) and ClearAll (whole sheet) on individual
' location sheets are sufficient and stay as they are - this is the
' higher-friction, encourages-a-backup-first path for pruning across rooms.
Public Sub DeleteVisibleReports()
    Dim ws As Worksheet, rng As Range, n As Long
    Dim locCol As Long, jobCol As Long, i As Long, loc As String
    Dim rooms As clsDict, k As Variant, breakdown As String, staleWarn As String
    Dim curSig As String, lastSig As String

    Set ws = ReportsSheet()
    If ws Is Nothing Then
        Say "The Reports sheet could not be found.", "Run Refresh Locations first."
        Exit Sub
    End If

    On Error Resume Next
    Set rng = ws.Range("A15").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        Say "There is nothing to delete.", "The Reports sheet has no results under the current filters."
        Exit Sub
    End If
    ' A15 may be spilling FILTER's own "no jobs match"/"no jobs recorded"
    ' fallback TEXT rather than real rows - see FilteredSig's own comment.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        Say "There is nothing to delete.", CStr(rng.Cells(1, 1).Value)
        Exit Sub
    End If
    n = rng.Rows.Count

    locCol = ColByHeader(ws, 14, "Location")
    jobCol = ColByHeader(ws, 14, "Job ID")
    If locCol = 0 Or jobCol = 0 Then
        Say "The Reports sheet layout looks wrong.", "The Location or Job ID column could not be found.", "Rebuild the report sheets (Refresh Locations), then try again."
        Exit Sub
    End If

    ' Broken down by room, following the same "show what you're about to
    ' lose" pattern as modJobs.RemoveRow/ClearAll.
    Set rooms = New clsDict
    For i = 1 To n
        loc = CStr(rng.Cells(i, locCol).Value)
        If rooms.Exists(loc) Then
            rooms.Add loc, CLng(rooms.Item(loc)) + 1
        Else
            rooms.Add loc, 1
        End If
    Next i
    For Each k In rooms.Keys
        breakdown = breakdown & "  " & CStr(k) & ": " & rooms.Item(CStr(k)) & vbCrLf
    Next k

    ' Warn when the current filtered set does not match what the last
    ' Export report run actually captured - never exported, or the filters
    ' or underlying data have changed since. Replaces the original "warn if
    ' not exported to CSV" wording, since deletion here is keyed off the
    ' Export report signature rather than per-location CSV export status.
    curSig = FilteredSig(ws)
    lastSig = ReportsExportSig(ws)
    If Len(lastSig) = 0 Then
        staleWarn = vbCrLf & vbCrLf & "Export report has never been run for a filtered set like this one - " & _
            "there is no up-to-date report or backup of what is about to be deleted."
    ElseIf StrComp(lastSig, curSig, vbBinaryCompare) <> 0 Then
        staleWarn = vbCrLf & vbCrLf & "The filters or underlying data have changed since the last Export report - " & _
            "there is no up-to-date report or backup of what is about to be deleted."
    End If

    If Not Ask("Delete " & n & " visible record" & IIf(n = 1, "", "s") & " from the Reports page?" & vbCrLf & vbCrLf & _
        breakdown & staleWarn & vbCrLf & vbCrLf & "This cannot be undone.", "Delete visible records") Then Exit Sub

    DeleteVisibleReportsConfirmed ws, rng, locCol, jobCol, n
End Sub

' The actual deletion. Public so a test can call it directly, bypassing the
' Ask() gate above - Ask() always declines under SetQuiet, same reason
' modJobs.RemoveRow/ClearAll and modImport.ApplyImportConfirmed are split
' this way.
Public Sub DeleteVisibleReportsConfirmed(ByVal ws As Worksheet, ByVal rng As Range, _
                                         ByVal LocCol As Long, ByVal JobCol As Long, ByVal n As Long)
    Dim i As Long, targetWs As Worksheet, lo As ListObject, rowIdx As Long
    Dim deleted As Long, missing As Long, detail As String
    Dim jobIds() As String, locs() As String

    On Error GoTo Fail

    ' Copied into memory FIRST. rng points at a live spilled formula range -
    ' the moment the first row is deleted from a source table, _Data (and
    ' this sheet's own FILTER) shrink, and rng's cells would be reading a
    ' moving target for every row after the first.
    ReDim jobIds(1 To n)
    ReDim locs(1 To n)
    For i = 1 To n
        jobIds(i) = CStr(rng.Cells(i, JobCol).Value)
        locs(i) = CStr(rng.Cells(i, LocCol).Value)
    Next i

    AppOff
    For i = 1 To n
        Set targetWs = SheetForCode(locs(i))
        rowIdx = 0
        If Not targetWs Is Nothing Then
            Set lo = JobsTable(targetWs)
            If Not lo Is Nothing Then rowIdx = FindReportRow(lo, jobIds(i))
        End If
        If rowIdx > 0 Then
            UnlockSheet targetWs
            lo.ListRows(rowIdx).Delete
            RelockSheet targetWs
            deleted = deleted + 1
        Else
            missing = missing + 1
        End If
    Next i
    AppOn

    detail = deleted & " record" & IIf(deleted = 1, "", "s") & " deleted from the Reports page."
    If missing > 0 Then detail = detail & " " & missing & " could not be found (already removed?)."
    LogAudit "Delete visible (Reports)", "(multiple rooms)", detail

    Say detail, "A note of what was removed has been kept in the workbook's audit log."
    Exit Sub
Fail:
    AppReset
    ReportError "Delete visible records"
End Sub

Private Function FindReportRow(ByVal lo As ListObject, ByVal JobId As String) As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Job ID").Value)), JobId, vbTextCompare) = 0 Then
            FindReportRow = i
            Exit Function
        End If
    Next i
End Function

Private Function ColByHeader(ByVal ws As Worksheet, ByVal HdrRow As Long, ByVal Header As String) As Long
    Dim c As Long
    For c = 1 To 100
        If Len(Trim$(CStr(ws.Cells(HdrRow, c).Value))) = 0 Then Exit Function
        If StrComp(CStr(ws.Cells(HdrRow, c).Value), Header, vbTextCompare) = 0 Then
            ColByHeader = c
            Exit Function
        End If
    Next c
End Function

' ============================================================== helpers ===
Private Function SheetNamed(ByVal Nm As String, ByVal Position As Long) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Nm)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = Nm
    End If

    ' Move Before: is only correct when the sheet currently sits LATER than
    ' the target. Moving a sheet that is already earlier "before" the sheet at
    ' Position is a no-op - taking it out shifts everything after it down by
    ' one, so it lands back where it started. That is why every build left
    ' Reports on tab 1 and Summary on tab 2, the reverse of what this function
    ' asks for: the .xlsx ships neither report sheet, so build.ps1 always
    ' takes the first-run path, where Reports is created at index 1 and then
    ' asked to move to 2. Moving later needs After:.
    If ws.Index > Position Then
        ws.Move Before:=ThisWorkbook.Worksheets(Position)
    ElseIf ws.Index < Position Then
        ws.Move After:=ThisWorkbook.Worksheets(Position)
    End If

    Set SheetNamed = ws
End Function

' A spill cannot be written into a range that still holds an old one, and
' Cells.Clear does not always release a stale spill parent on a rebuild.
Private Sub ClearSpillArea(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("A10:BZ5000").ClearContents
    On Error GoTo 0
End Sub

Private Sub TotalCell(ByVal ws As Worksheet, ByVal ValueAddr As String, ByVal Label As String, ByVal f As String)
    Dim lab As Range
    Set lab = ws.Range(ValueAddr).Offset(0, -1)
    lab.Value = Label
    lab.Font.Bold = True
    lab.HorizontalAlignment = xlRight
    ws.Range(ValueAddr).Formula2 = f
    ws.Range(ValueAddr).Font.Bold = True
End Sub

Private Sub CritCell(ByVal ws As Worksheet, ByVal LabelAddr As String, ByVal InputAddr As String, _
                     ByVal Label As String, ByVal Hint As String)
    ws.Range(LabelAddr).Value = Label
    ws.Range(LabelAddr).Font.Bold = True
    With ws.Range(InputAddr)
        .Locked = False
        .Interior.Color = RGB(255, 255, 255)
        .Borders(xlEdgeLeft).Color = RGB(46, 100, 168)
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    ws.Range(InputAddr).Offset(0, 1).Value = Hint
    ws.Range(InputAddr).Offset(0, 1).Font.Italic = True
    ws.Range(InputAddr).Offset(0, 1).Font.Color = RGB(110, 110, 110)
End Sub

Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal Headers As Variant)
    Dim i As Long
    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(RowNo, i - LBound(Headers) + 1).Value = Headers(i)
    Next i
End Sub

' ==================================================== export report sig ===
' A signature of the Reports sheet's CURRENT filtered/sorted results - same
' shape as modExport.ExportSig (row count, cell count, a cost sum, a max
' date) - so "Export report" can stamp what it captured, and the Reports-page
' delete can warn when the visible set has moved on since.
Public Function FilteredSig(ByVal ws As Worksheet) As String
    Dim rng As Range, n As Long, cellCount As Double, chg As Double, last As Double, bad As String

    On Error Resume Next
    Set rng = ws.Range("A15").SpillingToRange
    On Error GoTo 0
    If rng Is Nothing Then
        FilteredSig = "empty"
        Exit Function
    End If

    ' A15 may be spilling the "no jobs match"/"no jobs recorded" fallback
    ' TEXT rather than real rows. Row count alone reads 1 either way, but the
    ' first cell of a real row is always a Date/Time serial number, so
    ' ISNUMBER is what actually tells the two apart.
    If Not IsNumeric(rng.Cells(1, 1).Value2) Then
        FilteredSig = "empty"
        Exit Function
    End If

    n = rng.Rows.Count
    cellCount = Application.WorksheetFunction.CountA(rng)
    chg = ResultsAgg(rng, 12, "Sum", bad)   ' column L = Chargeable
    last = ResultsAgg(rng, 1, "Max", bad)   ' column A = Date/Time

    FilteredSig = n & "|" & Format$(cellCount, "0") & "|" & Format$(chg, "0.####") & _
                  "|" & Format$(last, "0.######") & bad
End Function

Private Function ResultsAgg(ByVal rng As Range, ByVal ColOffset As Long, ByVal What As String, ByRef Bad As String) As Double
    On Error GoTo Failed
    Select Case What
        Case "Sum": ResultsAgg = Application.WorksheetFunction.Sum(rng.Columns(ColOffset))
        Case "Max": ResultsAgg = Application.WorksheetFunction.Max(rng.Columns(ColOffset))
        Case Else: GoTo Failed
    End Select
    Exit Function
Failed:
    ResultsAgg = 0
    Bad = Bad & "|?" & What
End Function

Public Function ReportsSheet() As Worksheet
    On Error Resume Next
    Set ReportsSheet = ThisWorkbook.Worksheets(SHEET_REPORTS)
End Function

Public Function ReportsExportSig(ByVal ws As Worksheet) As String
    ReportsExportSig = CStr(ws.Range(EXPORT_SIG_CELL).Value)
End Function

Public Function ReportsExportWhen(ByVal ws As Worksheet) As Variant
    ReportsExportWhen = ws.Range(EXPORT_WHEN_CELL).Value
End Function

' Called by modExport.ExportReportSnapshot once the .xlsx has actually been
' written - never optimistically, for the same reason modExport.StampExported
' only runs after modExport confirms the file exists.
Public Sub StampReportsExport(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Range(EXPORT_SIG_CELL).Value = FilteredSig(ws)
    ws.Range(EXPORT_WHEN_CELL).Value = Now
    ws.Range(EXPORT_WHEN_CELL).NumberFormat = "dd/mm/yyyy hh:mm"
    RelockSheet ws
End Sub

' A dropdown from a literal comma-separated Formula1 list, e.g. Ascending,
' Descending - fine here because both lists (report headers, Asc/Desc) are
' short and fixed, well under xlValidateList's 255-character literal limit.
Private Sub AddList(ByVal target As Range, ByVal QuotedItems As String, ByVal Title As String, ByVal Msg As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Replace$(QuotedItems, """", "")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Choose one of the listed options, or leave it blank."
    End With
End Sub

' A cell with Date-type validation gets Excel's own calendar picker (the
' small icon that appears on selection, in Excel for Microsoft 365) with no
' ActiveX involved - consistent with modPicker's Mac-safe, no-ActiveX rule.
' Validation only fires on manual entry, never on a value set from VBA/COM,
' so this adds a UI convenience without narrowing what modImport or a test
' script can write, and the Criteria() formula's own tolerance for a text
' date left by an unformatted cell is still needed and unchanged.
'
' The lower bound is a real DATE() formula, not a literal date string - a
' string like "01/01/2000" is read back through the machine's locale, the
' same class of trap modUtils.NumOf's comment warns about for numbers.
Private Sub AddDateValidation(ByVal target As Range, ByVal Title As String, ByVal Msg As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="=DATE(2000,1,1)"
        .IgnoreBlank = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Enter a valid date on or after 1 January 2000, or leave blank."
    End With
End Sub

' Technician/Printer/Paper Stock dropdowns, sourced from what has actually
' been recorded (the consolidated _Data range) rather than the current
' Papers/Printers/Technicians catalogue - so a job against a since-renamed
' or deactivated printer or a technician no longer active is still findable,
' and every choice offered is guaranteed to match at least one record.
'
' Public, and also called from modRegistry.RefreshLocations after it
' rewrites _Data - not just from here. _Data does not exist yet the first
' time BuildReports runs (InitialiseWorkbook builds the report sheets before
' RefreshLocations ever writes it), and it changes on every later refresh
' too, so a one-time snapshot taken only at build time would read empty on
' a fresh workbook and go stale the moment a job is added anywhere.
Public Sub RefreshReportFilterLists(ByVal ws As Worksheet)
    ApplyTo ws, ws.Range("G5"), DistinctValues("Technician"), "REP|Technician", _
        "Technician", "Choose a technician, or leave blank to include all."
    ApplyTo ws, ws.Range("G6"), DistinctValues("Printer"), "REP|Printer", _
        "Printer", "Choose a printer, or leave blank to include all."
    ApplyTo ws, ws.Range("G7"), DistinctValues("Paper Stock"), "REP|Paper stock", _
        "Paper stock", "Choose a paper stock, or leave blank to include all."
End Sub

' Distinct, sorted, non-blank values of one column of the consolidated range,
' read directly in VBA rather than as a formula - ApplyTo's staging column
' needs a Collection of plain values, not a spilled array.
Private Function DistinctValues(ByVal Header As String) As Collection
    Dim out As Collection, seen As clsDict
    Dim rng As Range, hdrRng As Range, m As Variant
    Dim hdrCol As Long, i As Long, j As Long, v As String, tmp As String
    Dim vals() As String, n As Long

    Set out = New Collection
    Set rng = ConsolidatedRange()
    If rng Is Nothing Then
        Set DistinctValues = out
        Exit Function
    End If

    ' DATA_HDR is already sheet-qualified ("_Data!$A$9:$AZ$9"), and
    ' Worksheet.Range does not reliably accept a qualified address string -
    ' Application.Range does.
    Set hdrRng = Application.Range(DATA_HDR)
    m = Application.Match(Header, hdrRng, 0)
    If IsError(m) Then
        Set DistinctValues = out
        Exit Function
    End If
    hdrCol = CLng(m)   ' Both rng and hdrRng start at column A, so this offset lines up with rng directly.

    Set seen = New clsDict
    ReDim vals(1 To rng.Rows.Count)
    For i = 1 To rng.Rows.Count
        v = Trim$(CStr(rng.Cells(i, hdrCol).Value))
        If Len(v) > 0 Then
            If Not seen.Exists(v) Then
                seen.Add v, True
                n = n + 1
                vals(n) = v
            End If
        End If
    Next i

    ' An insertion sort is plenty for a list this size (printers, paper
    ' stocks, technicians number in the tens, not thousands).
    For i = 2 To n
        tmp = vals(i)
        j = i - 1
        Do While j >= 1
            If StrComp(vals(j), tmp, vbTextCompare) <= 0 Then Exit Do
            vals(j + 1) = vals(j)
            j = j - 1
        Loop
        vals(j + 1) = tmp
    Next i

    For i = 1 To n
        out.Add vals(i)
    Next i
    Set DistinctValues = out
End Function
