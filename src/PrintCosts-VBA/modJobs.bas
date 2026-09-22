Attribute VB_Name = "modJobs"
Option Explicit

' Adding, stamping, removing and clearing print jobs.

Public Sub AddPrintJob(ByVal ws As Worksheet)
    Dim lo As ListObject, r As ListRow, n As Long
    Set lo = JobsTable(ws)
    If lo Is Nothing Then
        Say "This sheet has no print job table.", "Add Print Job only works on a print room sheet.", "Switch to a print room tab and try again."
        Exit Sub
    End If

    AppOff
    UnlockSheet ws
    If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
        Set r = lo.ListRows(1)
    Else
        Set r = lo.ListRows.Add
    End If
    n = r.Index

    CellIn(lo, n, "Job ID").Value = NewJobId(ws, lo)
    CellIn(lo, n, "Date/Time").Value = Now

    ' Spec 9.2 and 10.10: the location defaults seed the row, and from this
    ' moment the row's own values are independent of them. Copying rather than
    ' referencing is the whole of AT-07 and AT-08.
    CellIn(lo, n, "Disregard Paper").Value = DefaultOrNo(ws, "LOC_DefDisPaper")
    CellIn(lo, n, "Disregard Consumable").Value = DefaultOrNo(ws, "LOC_DefDisCons")

    CellIn(lo, n, "S_SchemaVer").Value = SCHEMA_VER
    RelockSheet ws
    BindStockCell ws, lo, n
    BindPrinterCell ws, lo, n
    AppOn

    CellIn(lo, n, "Student Name").Select
End Sub

Private Function DefaultOrNo(ByVal ws As Worksheet, ByVal RefName As String) As String
    Dim v As String
    v = LocValue(ws, RefName)
    If StrComp(v, "Yes", vbTextCompare) = 0 Then DefaultOrNo = "Yes" Else DefaultOrNo = "No"
End Function

Private Function NewJobId(ByVal ws As Worksheet, ByVal lo As ListObject) As String
    ' <SITE>-<LOCATION>-00001. The site component matters because records from
    ' several workbook copies may later be collated, and two copies of the same
    ' template would otherwise both hold a location coded MAIN.
    '
    ' Allocated from modRegistry.NextJobId's persisted high-water mark, not
    ' from a scan of this sheet's rows - see that function's comment for why.
    NewJobId = NextJobId(ws, lo)
End Function

Public Sub StampNow(ByVal ws As Worksheet)
    Dim lo As ListObject, n As Long
    Set lo = JobsTable(ws)
    n = SelectedRow(ws, lo)
    If n = 0 Then Exit Sub
    CellIn(lo, n, "Date/Time").Value = Now
End Sub

Public Sub RemoveRow(ByVal ws As Worksheet)
    Dim lo As ListObject, n As Long, detail As String
    Set lo = JobsTable(ws)
    n = SelectedRow(ws, lo)
    If n = 0 Then Exit Sub

    detail = CStr(CellIn(lo, n, "Job ID").Value) & "  " & Format$(CellIn(lo, n, "Date/Time").Value, "dd/mm/yyyy hh:mm") & vbCrLf & "Student: " & NzText(CellIn(lo, n, "Student Name").Value, "(not given)") & "  " & CStr(CellIn(lo, n, "Student No").Value) & vbCrLf & "Printer: " & CStr(CellIn(lo, n, "Printer").Value) & vbCrLf & "Stock: " & CStr(CellIn(lo, n, "Paper Stock").Value) & vbCrLf & "Cost: " & Format$(CellIn(lo, n, "Gross Cost").Value, CurrencySymbol() & "#,##0.00")

    ' Never a bare "are you sure" - spec 10.12 requires the user to see what
    ' they are about to lose.
    If Not Ask("Delete this print job?" & vbCrLf & vbCrLf & detail & vbCrLf & vbCrLf & "This cannot be undone.", "Remove print job") Then Exit Sub

    AppOff
    LogAudit "Remove row", LocValue(ws, "LOC_Name"), detail
    UnlockSheet ws
    lo.ListRows(n).Delete
    RelockSheet ws
    AppOn
End Sub

Public Sub ClearAll(ByVal ws As Worksheet)
    Dim lo As ListObject, n As Long, i As Long
    Dim lo_first As Double, lo_last As Double, d As Double, span As String

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    n = RowCount(lo)
    If n = 0 Then
        Say "There is nothing to clear.", "'" & LocValue(ws, "LOC_Name") & "' has no print jobs recorded."
        Exit Sub
    End If

    For i = 1 To n
        d = DateSerialOf(CellIn(lo, i, "Date/Time"))
        If d > 0 Then
            If lo_first = 0 Or d < lo_first Then lo_first = d
            If d > lo_last Then lo_last = d
        End If
    Next i
    If lo_first > 0 Then
        span = " dated " & Format$(lo_first, "dd/mm/yyyy") & " to " & Format$(lo_last, "dd/mm/yyyy")
    End If

    ' Spec 11 is explicit about what the confirmation must name: the location,
    ' the number of records, and the date range.
    If Not Ask("Clear All - " & LocValue(ws, "LOC_Name") & vbCrLf & vbCrLf & "This will permanently delete " & n & " print job" & IIf(n = 1, "", "s") & span & "." & vbCrLf & vbCrLf & "Records on other print room sheets are not affected." & vbCrLf & "This cannot be undone. Continue?", "Clear All") Then Exit Sub

    AppOff
    LogAudit "Clear All", LocValue(ws, "LOC_Name"), n & " records deleted" & span
    UnlockSheet ws
    For i = lo.ListRows.Count To 1 Step -1
        lo.ListRows(i).Delete
    Next i
    lo.ListRows.Add
    RelockSheet ws
    AppOn

    Say n & " print job" & IIf(n = 1, "", "s") & " deleted from '" & LocValue(ws, "LOC_Name") & "'.", "A note of what was removed has been kept in the workbook's audit log."
End Sub

' ------------------------------------------------------------------ shared ---
Public Function SelectedRow(ByVal ws As Worksheet, ByRef lo As ListObject) As Long
    If lo Is Nothing Then Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim c As Range
    Set c = Application.Intersect(Selection.Cells(1, 1).EntireRow, lo.DataBodyRange)
    If c Is Nothing Then
        Say "No print job row is selected.", "This command works on the row the cursor is in.", "Click any cell in the print job you want, then try again."
        Exit Function
    End If
    SelectedRow = c.Row - lo.DataBodyRange.Row + 1
End Function

Public Function NzText(ByVal v As Variant, ByVal IfBlank As String) As String
    If Len(Trim$(CStr(v))) = 0 Then NzText = IfBlank Else NzText = CStr(v)
End Function
