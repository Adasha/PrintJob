Attribute VB_Name = "modReports"
Option Explicit

' Builds the Summary and Cost Calculations sheets.
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
Private Const SHEET_CALC As String = "Cost Calculations"
Private Const DATA_SPILL As String = "_Data!$A$10#"
Private Const DATA_HDR As String = "_Data!$A$9:$AZ$9"

' A column of the consolidated range, found by its header text.
Private Function C(ByVal Header As String) As String
    C = "INDEX(" & DATA_SPILL & ",,MATCH(""" & Header & """," & DATA_HDR & ",0))"
End Function

' SUMIFS of one column, grouped by the Location/Paper stock key pair. The
' criteria are arrays, which is what makes the result spill down alongside the
' keys instead of needing a formula per row.
Private Function SumBy(ByVal Header As String) As String
    SumBy = "SUMIFS(" & C(Header) & "," & C("Location") & ",lo," & C("Paper Stock") & ",st)"
End Function

Public Sub BuildReportSheets()
    BuildSummary
    BuildCostCalc
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
    ws.Range("A3").Value = "Every print job recorded in this workbook, by print room and paper stock. " & _
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
    WriteHeaderRow ws, 9, Array("Location", "Paper stock", "Type", "Family", "Unit", _
                                "Jobs", "Quantity", "Area m2", "Paper cost", _
                                "Consumable cost", "Gross", "Disregarded", "Chargeable")

    ' --- the one formula ---------------------------------------------------
    f = "=IFERROR(LET(" & _
        "keys,SORT(UNIQUE(HSTACK(" & C("Location") & "," & C("Paper Stock") & "))),"
    f = f & "lo,INDEX(keys,,1),st,INDEX(keys,,2),"
    f = f & "HSTACK(lo,st," & _
        "IFNA(XLOOKUP(st,tblPapers[Description],tblPapers[Paper type]),""(not in Papers)""),"
    f = f & "IFNA(XLOOKUP(st,tblPapers[Description],tblPapers[Family]),""(not in Papers)""),"
    f = f & "IFNA(IF(XLOOKUP(st,tblPapers[Description],tblPapers[Measure])=""Sheet"",""sheets"",""metres""),""-""),"
    f = f & "COUNTIFS(" & C("Location") & ",lo," & C("Paper Stock") & ",st),"
    f = f & SumBy("Quantity") & "," & SumBy("Area m2") & "," & SumBy("Paper Cost") & ","
    f = f & SumBy("Consumable Cost") & "," & SumBy("Gross Cost") & ","
    f = f & SumBy("Disregarded") & "," & SumBy("Chargeable Cost") & ")),"
    f = f & """No print jobs have been recorded yet."")"
    ws.Range("A10").Formula2 = f

    FormatSummary ws
    RelockSheet ws
End Sub

Private Sub FormatSummary(ByVal ws As Worksheet)
    ws.Range("A9:M9").Interior.Color = RGB(222, 232, 244)
    ws.Range("G10:H2000").NumberFormat = "#,##0.00"
    ws.Range("I10:M2000").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("D6").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("F6").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Range("H6").NumberFormat = ChrW(163) & "#,##0.00"
    ws.Columns("A:M").ColumnWidth = 14
    ws.Columns("A:B").ColumnWidth = 24
    ws.Rows(9).Font.Bold = True
End Sub

' ==================================================== Cost Calculations ===
' The four criteria collapse to TRUE when blank, which is what gives the
' four-case behaviour AT-12 tests without a single branch in the formula.
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

    Criteria = s
End Function

Public Sub BuildCostCalc()
    Dim ws As Worksheet, f As String, ok As String

    Set ws = SheetNamed(SHEET_CALC, 2)
    UnlockSheet ws
    ws.Cells.Clear
    ok = Criteria

    ws.Range("A1").Value = "Cost Calculations"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Find what a student has spent. Results update as you type - there is no search button. " & _
        "Leave a box empty to ignore it."

    CritCell ws, "A5", "C5", "Student name", "Part of a name is enough - ""Smith"" finds ""Jane Smith""."
    CritCell ws, "A6", "C6", "Student number", "Matched exactly."
    CritCell ws, "A7", "C7", "From date", "Leave blank for no start date."
    CritCell ws, "A8", "C8", "To date", "Jobs logged at any time on this date are included."
    ws.Range("C7:C8").NumberFormat = "dd/mm/yyyy"

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
    WriteHeaderRow ws, 14, Array("Date/Time", "Location", "Printer", "Paper stock", _
                                 "Quantity", "Unit", "Area m2", "Paper cost", _
                                 "Consumable cost", "Gross", "Disregarded", "Chargeable", _
                                 "Technician", "Notes")
    f = "=IFERROR(FILTER(HSTACK(" & C("Date/Time") & "," & C("Location") & "," & C("Printer") & _
        "," & C("Paper Stock") & "," & C("Quantity") & "," & C("Unit") & "," & C("Area m2") & _
        "," & C("Paper Cost") & "," & C("Consumable Cost") & "," & C("Gross Cost") & _
        "," & C("Disregarded") & "," & C("Chargeable Cost") & "," & C("Technician") & _
        "," & C("Notes") & ")," & ok & ",""No print jobs match those criteria.""),"
    f = f & """No print jobs have been recorded yet."")"
    ws.Range("A15").Formula2 = f

    BuildBreakdowns ws, ok
    FormatCostCalc ws
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

Private Sub FormatCostCalc(ByVal ws As Worksheet)
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
    ' Cost Calculations on tab 1 and Summary on tab 2, the reverse of what
    ' this function asks for: the .xlsx ships neither report sheet, so
    ' build.ps1 always takes the first-run path, where Cost Calculations is
    ' created at index 1 and then asked to move to 2. Moving later needs After:.
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
