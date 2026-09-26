Attribute VB_Name = "modJobs"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Adding, stamping, removing and clearing print jobs.

Public Sub AddPrintJob(ByVal ws As Worksheet)
    Dim lo As ListObject, r As ListRow, n As Long
    Set lo = JobsTable(ws)
    If lo Is Nothing Then
        Say "This sheet has no print job table.", "Add Print Job only works on a print room sheet.", "Switch to a print room tab and try again."
        Exit Sub
    End If

    ' modValidation.EnsureDefaultsClean: don't trust that an incompatible
    ' printer/paper default was already caught at the moment it was set - see
    ' that sub's comment for why (Excel for Mac's dropdown selection).
    EnsureDefaultsClean ws

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

    ' Snag list item 1c: no location default for Paid, unlike the disregard
    ' flags above - every new job simply starts unpaid.
    CellIn(lo, n, "Paid").Value = "No"

    ' Spec 1b: the batch defaults above the toolbar seed Technician/Printer/
    ' Paper Stock too, same copy-not-reference principle - changing a default
    ' afterwards never touches a job already added from it.
    CopyDefault ws, "LOC_DefTech", CellIn(lo, n, "Technician")
    CopyDefault ws, "LOC_DefPrinter", CellIn(lo, n, "Printer")
    CopyDefault ws, "LOC_DefPaper", CellIn(lo, n, "Paper Stock")

    CellIn(lo, n, "S_SchemaVer").Value = SCHEMA_VER
    RelockSheet ws
    BindStockCell ws, lo, n
    BindPrinterCell ws, lo, n
    AppOn

    CellIn(lo, n, "Student Name").Select
End Sub

' Repeat Job (direct user request): duplicates the selected row into a new
' one rather than starting from the location's own batch defaults the way
' AddPrintJob does - same Student Name/No, Printer, Paper Stock, Unit, Qty,
' Print Width mm and both Disregard flags. Job ID, Date/Time and Paid are
' always reset (a copy is its own job, logged now, unpaid), Technician
' follows the copy-not-reference default AddPrintJob already uses (spec
' 9.2/10.10), and a note records what it was copied from.
'
' Every value carried over came from a row that was already a real, valid
' job a moment ago, so unlike AddPrintJob there is no printer/paper
' compatibility re-check to make. StampRow still runs, though - Application.
' EnableEvents is off for this whole operation (AppOff), so nothing else
' will re-cost the copy at today's rates the way OnPrinterChanged/
' OnStockChanged would if these same values had just been typed by hand.
Public Sub RepeatJob(ByVal ws As Worksheet)
    Dim lo As ListObject, srcRow As Long, r As ListRow, n As Long
    Dim oldJobId As String, srcNotes As String

    Set lo = JobsTable(ws)
    If lo Is Nothing Then
        Say "This sheet has no print job table.", "Repeat Job only works on a print room sheet.", "Switch to a print room tab and try again."
        Exit Sub
    End If

    srcRow = SelectedRow(ws, lo)
    If srcRow = 0 Then Exit Sub

    oldJobId = Trim$(CStr(CellIn(lo, srcRow, "Job ID").Value))
    If Len(oldJobId) = 0 Then
        Say "This row has no print job to repeat.", "Repeat Job duplicates an existing print job, and the selected row is blank.", "Click a print job you want to repeat, then try again."
        Exit Sub
    End If

    AppOff
    UnlockSheet ws
    Set r = lo.ListRows.Add
    n = r.Index

    CellIn(lo, n, "Student Name").Value = CellIn(lo, srcRow, "Student Name").Value
    CellIn(lo, n, "Student No").Value = CellIn(lo, srcRow, "Student No").Value
    CellIn(lo, n, "Printer").Value = CellIn(lo, srcRow, "Printer").Value
    CellIn(lo, n, "Paper Stock").Value = CellIn(lo, srcRow, "Paper Stock").Value
    CellIn(lo, n, "Unit").Value = CellIn(lo, srcRow, "Unit").Value
    CellIn(lo, n, "Qty").Value = CellIn(lo, srcRow, "Qty").Value
    CellIn(lo, n, "Print Width mm").Value = CellIn(lo, srcRow, "Print Width mm").Value
    CellIn(lo, n, "Disregard Paper").Value = CellIn(lo, srcRow, "Disregard Paper").Value
    CellIn(lo, n, "Disregard Consumable").Value = CellIn(lo, srcRow, "Disregard Consumable").Value

    CellIn(lo, n, "Job ID").Value = NewJobId(ws, lo)
    CellIn(lo, n, "Date/Time").Value = Now
    CellIn(lo, n, "Paid").Value = "No"

    CellIn(lo, n, "Technician").ClearContents
    CopyDefault ws, "LOC_DefTech", CellIn(lo, n, "Technician")

    srcNotes = Trim$(CStr(CellIn(lo, srcRow, "Notes").Value))
    If Len(srcNotes) > 0 Then
        CellIn(lo, n, "Notes").Value = srcNotes & vbCrLf & "Copy of " & oldJobId
    Else
        CellIn(lo, n, "Notes").Value = "Copy of " & oldJobId
    End If

    RelockSheet ws
    BindStockCell ws, lo, n
    BindPrinterCell ws, lo, n
    StampRow ws, n
    AppOn

    CellIn(lo, n, "Student Name").Select
End Sub

Private Function DefaultOrNo(ByVal ws As Worksheet, ByVal RefName As String) As String
    Dim v As String
    v = LocValue(ws, RefName)
    If StrComp(v, "Yes", vbTextCompare) = 0 Then DefaultOrNo = "Yes" Else DefaultOrNo = "No"
End Function

Private Sub CopyDefault(ByVal ws As Worksheet, ByVal RefName As String, ByVal target As Range)
    Dim v As String
    v = LocValue(ws, RefName)
    If Len(v) > 0 Then target.Value = v
End Sub

' Snag list item 1b's Clear button: empties the three batch-default cells for
' the next batch of jobs, without touching any job already added from them
' (spec 9.2/10.10's copy-not-reference principle already guarantees that).
Public Sub ClearDefaults(ByVal ws As Worksheet)
    AppOff
    UnlockSheet ws
    Dim c As Range
    Set c = LocRange(ws, "LOC_DefTech")
    If Not c Is Nothing Then c.ClearContents
    Set c = LocRange(ws, "LOC_DefPrinter")
    If Not c Is Nothing Then c.ClearContents
    Set c = LocRange(ws, "LOC_DefPaper")
    If Not c Is Nothing Then c.ClearContents
    RelockSheet ws
    BindDefaultCells ws
    AppOn
End Sub

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

    ' "Student/Department" (snag 4b, 2026-09-22): the Student Name/No columns
    ' are also used free-text for department charging (D19) - the wording
    ' here just needs to read sensibly either way, not the header name itself.
    detail = CStr(CellIn(lo, n, "Job ID").Value) & "  " & Format$(CellIn(lo, n, "Date/Time").Value, "dd/mm/yyyy hh:mm") & vbCrLf & "Student/Department: " & NzText(CellIn(lo, n, "Student Name").Value, "(not given)") & "  " & CStr(CellIn(lo, n, "Student No").Value) & vbCrLf & "Printer: " & CStr(CellIn(lo, n, "Printer").Value) & vbCrLf & "Stock: " & CStr(CellIn(lo, n, "Paper Stock").Value) & vbCrLf & "Cost: " & Format$(CellIn(lo, n, "Gross Cost").Value, CurrencySymbol() & "#,##0.00")

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
