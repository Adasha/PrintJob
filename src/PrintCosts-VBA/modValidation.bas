Attribute VB_Name = "modValidation"
Option Explicit

' Validation that a formula cannot express.
'
' The Status column already reports everything checkable within a row, and it
' keeps working with macros disabled. This module covers the two things it
' cannot do: reject a bad value at the moment of entry, and compare a row
' against every other row in the workbook.

' Called from Workbook_SheetChange. Returns True if it changed the cell.
Public Function OnCellChanged(ByVal ws As Worksheet, ByVal Target As Range) As Boolean
    Dim lo As ListObject, n As Long, hdr As String

    ' The batch-default cells above the toolbar (spec 1b) aren't part of the
    ' job table, but need the same AT-05 compatibility check and
    ' bidirectional rebind as a table row's own Printer/Paper Stock (spec 1a).
    If OnDefaultCellChanged(ws, Target) Then Exit Function

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Function
    If Target.Cells.Count > 1 Then Exit Function

    n = Target.Row - lo.DataBodyRange.Row + 1
    hdr = CStr(lo.HeaderRowRange.Cells(1, Target.Column - lo.Range.Column + 1).Value)

    Select Case hdr
        Case "Printer"
            OnPrinterChanged ws, lo, n
        Case "Paper Stock"
            OnStockChanged ws, lo, n
        Case "Print Width mm"
            OnWidthChanged ws, lo, n
        Case "Student Name", "Student No"
            CheckStudent ws, lo, n
        Case "Technician"
            StampRow ws, n
    End Select
End Function

' Mirrors OnPrinterChanged/OnStockChanged for the two default cells above the
' toolbar (spec 1b) - same AT-05 compatibility check, same bidirectional
' rebind, just addressed by name instead of by table row. Returns True when
' Target was one of the default cells, so OnCellChanged can skip the
' table-row handling entirely (they are never both true for the same edit).
Private Function OnDefaultCellChanged(ByVal ws As Worksheet, ByVal Target As Range) As Boolean
    Dim prnCell As Range, stkCell As Range, model As String, stk As String
    Dim s As clsStock, p As clsPrinterDef, changedPrinter As Boolean

    If Target.Cells.Count > 1 Then Exit Function
    Set prnCell = LocRange(ws, "LOC_DefPrinter")
    Set stkCell = LocRange(ws, "LOC_DefPaper")
    If prnCell Is Nothing Or stkCell Is Nothing Then Exit Function

    If Not Application.Intersect(Target, prnCell) Is Nothing Then
        changedPrinter = True
    ElseIf Application.Intersect(Target, stkCell) Is Nothing Then
        Exit Function
    End If
    OnDefaultCellChanged = True

    model = CStr(prnCell.Value)
    stk = CStr(stkCell.Value)
    If Len(model) > 0 And Len(stk) > 0 Then
        If Not Compatible(model, stk) Then
            Set s = Stock(stk)
            Set p = Prn(model)
            If changedPrinter Then
                stkCell.ClearContents
                Say "'" & stk & "' cannot be used on " & model & ".", "That printer does not support the " & s.Family & " paper family.", "The default paper stock has been cleared."
            Else
                prnCell.ClearContents
                Say "'" & stk & "' cannot be used on " & model & ".", "That stock is in the " & s.Family & " family, which this printer does not support.", "The default printer has been cleared."
            End If
        End If
    End If

    BindDefaultCells ws
End Function

' Either field can drive the other now (spec 1a). Both handlers end by
' rebinding BOTH cells' lists rather than just the other one: whichever field
' just changed may itself need widening back out, e.g. when the other field
' had to be cleared for incompatibility, or when this field was cleared
' outright and the other field's list was previously narrowed by it.
Private Sub OnPrinterChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim model As String, stk As String, s As clsStock
    model = CStr(CellIn(lo, n, "Printer").Value)
    stk = CStr(CellIn(lo, n, "Paper Stock").Value)

    ' A printer change can strand a stock that was valid a moment ago (AT-05).
    If Len(stk) > 0 And Len(model) > 0 Then
        If Not Compatible(model, stk) Then
            Set s = Stock(stk)
            CellIn(lo, n, "Paper Stock").ClearContents
            Say "'" & stk & "' cannot be used on " & model & ".", "That printer does not support the " & s.Family & " paper family.", "The paper stock has been cleared. Choose a stock this printer supports."
        End If
    End If

    BindStockCell ws, lo, n
    BindPrinterCell ws, lo, n
    StampRow ws, n
End Sub

Private Sub OnStockChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim model As String, stk As String, s As clsStock, p As clsPrinterDef
    model = CStr(CellIn(lo, n, "Printer").Value)
    stk = CStr(CellIn(lo, n, "Paper Stock").Value)
    If Len(stk) = 0 Then
        BindPrinterCell ws, lo, n
        Exit Sub
    End If

    ' A stock change can strand a printer that was valid a moment ago, now
    ' that Paper Stock can be chosen before Printer - the mirror of
    ' OnPrinterChanged's own AT-05 check.
    If Len(model) > 0 Then
        If Not Compatible(model, stk) Then
            Set s = Stock(stk)
            Set p = Prn(model)
            CellIn(lo, n, "Paper Stock").ClearContents
            Say "'" & stk & "' cannot be used on " & model & ".", "That stock is in the " & s.Family & " family, which this printer does " & "not support.", "Choose a stock in one of: " & p.Families
            BindStockCell ws, lo, n
            BindPrinterCell ws, lo, n
            Exit Sub
        End If
    End If

    BindStockCell ws, lo, n
    BindPrinterCell ws, lo, n
    StampRow ws, n

    ' Print width is meaningless for cut sheets (spec 10.6).
    Set s = Stock(stk)
    If s.Measure = "Sheet" Then
        If Len(CellIn(lo, n, "Print Width mm").Value) > 0 Then
            CellIn(lo, n, "Print Width mm").ClearContents
            Say "Print width does not apply to sheet stock.", "For sheets the whole sheet is treated as printed, so a print width " & "would have no effect.", "The value has been cleared."
        End If
    End If
End Sub

Private Sub OnWidthChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim w As Double, sw As Double, stk As String, s As clsStock
    If Len(CellIn(lo, n, "Print Width mm").Value) = 0 Then Exit Sub

    stk = CStr(CellIn(lo, n, "Paper Stock").Value)
    If Len(stk) = 0 Then Exit Sub
    Set s = Stock(stk)

    If s.Measure = "Sheet" Then
        CellIn(lo, n, "Print Width mm").ClearContents
        Say "Print width does not apply to sheet stock.", "'" & stk & "' is cut sheet, and the whole sheet counts as printed.", "The value has been cleared."
        Exit Sub
    End If

    w = NumOf(CellIn(lo, n, "Print Width mm"))
    sw = NumOf(CellIn(lo, n, "S_StockWidth_mm"))
    If sw = 0 Then sw = s.WidthMM

    ' AT-04.
    If w > sw Then
        CellIn(lo, n, "Print Width mm").ClearContents
        Say "Print width " & Format$(w, "#,##0") & " mm is wider than the stock.", "'" & stk & "' is " & Format$(sw, "#,##0") & " mm wide, so a print cannot " & "be wider than that.", "Enter " & Format$(sw, "#,##0") & " mm or less, or leave it blank to use " & "the full width of the roll."
    End If
End Sub

' ----------------------------------------------------------------- student ---
' There is no student register, so the existing records are the reference
' (AT-11). This warns and never blocks: the first entry of any student is
' unverifiable, and people legitimately change their names.
Public Sub CheckStudent(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim nm As String, no As String, other As String
    nm = Trim$(CStr(CellIn(lo, n, "Student Name").Value))
    no = Trim$(CStr(CellIn(lo, n, "Student No").Value))
    If Len(nm) = 0 Or Len(no) = 0 Then Exit Sub

    other = FindConflict(no, nm, CStr(CellIn(lo, n, "Job ID").Value))
    If Len(other) > 0 Then
        ' "Student/department" (snag 4b): Student Name/No is also used
        ' free-text for department charging (D19) - the wording here just
        ' needs to read sensibly either way, not the underlying field.
        Say "Student/department number " & no & " has been logged before under a different name.", "Earlier records show it as '" & other & "'; this row says '" & nm & "'.", "If this is the same person (or department) under a new name, carry on. If it is a typo, " & "correct it now - the cost report matches on both fields."
    End If
End Sub

Private Function FindConflict(ByVal StudentNo As String, ByVal StudentName As String, ByVal SkipJobId As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim n As String, nm As String
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            Set lo = JobsTable(ws)
            If Not lo Is Nothing Then
                For i = 1 To lo.ListRows.Count
                    If Not IsBlankRow(lo, i) Then
                        If CStr(CellIn(lo, i, "Job ID").Value) <> SkipJobId Then
                            n = Trim$(CStr(CellIn(lo, i, "Student No").Value))
                            nm = Trim$(CStr(CellIn(lo, i, "Student Name").Value))
                            If Len(n) > 0 And Len(nm) > 0 Then
                                If StrComp(n, StudentNo, vbTextCompare) = 0 Then
                                    If StrComp(nm, StudentName, vbTextCompare) <> 0 Then
                                        FindConflict = nm
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next ws
End Function

' ------------------------------------------------------------------ sweeps ---
' A row can be made invalid by a later edit somewhere else, so the Status
' column is not enough on its own.
Public Sub CheckSheet(ByVal ws As Worksheet)
    Dim msg As String
    BindColumns ws
    msg = SheetProblems(ws)
    If Len(msg) = 0 Then
        SetStatus ws, "Checked " & Format$(Now, "dd/mm/yyyy hh:mm") & " - no problems"
        Say "No problems found on '" & LocValue(ws, "LOC_Name") & "'."
    Else
        SetStatus ws, "Problems found - see the Status column"
        Say "Problems found on '" & LocValue(ws, "LOC_Name") & "':", msg, "The Status column on each row explains what to fix."
    End If
End Sub

Public Sub CheckWorkbook()
    Dim ws As Worksheet, all As String, part As String, n As Long
    Dim unexported As String, u As Long

    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            part = SheetProblems(ws)
            If Len(part) > 0 Then
                all = all & vbCrLf & LocValue(ws, "LOC_Name") & ":" & vbCrLf & part
                n = n + 1
                SetStatus ws, "Problems found - see the Status column"
            Else
                SetStatus ws, "Checked " & Format$(Now, "dd/mm/yyyy hh:mm") & " - no problems"
            End If

            ' Export state is reported alongside validity but counted apart.
            ' Unexported records are not a defect in the data - they are a
            ' risk to it, because deleting a sheet cannot be intercepted
            ' (design 3.3). Folding them into the problem count would make a
            ' perfectly valid workbook report problems.
            RefreshExportStatus ws
            If NeedsExport(ws) Then
                unexported = unexported & "- " & LocValue(ws, "LOC_Name") & ": " & ExportStatusText(ws) & vbCrLf
                u = u + 1
            End If
        End If
    Next ws

    If Len(unexported) > 0 Then
        unexported = vbCrLf & vbCrLf & "Not exported:" & vbCrLf & unexported & _
                     "Deleting a print room sheet destroys its records and cannot be undone. Export first."
    End If

    If n = 0 Then
        Say "No problems found.", "Every print room sheet checked out clean." & unexported, _
            IIf(u > 0, "Use the Export... button on each sheet listed.", "")
    Else
        Say "Problems found on " & n & " print room sheet" & IIf(n = 1, "", "s") & ".", _
            all & unexported, _
            "Go to each sheet and use the Status column to see what needs fixing."
    End If
End Sub

Private Function SheetProblems(ByVal ws As Worksheet) As String
    Dim lo As ListObject, i As Long, s As String, c As Long
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If Not IsBlankRow(lo, i) Then
            s = Trim$(CStr(CellIn(lo, i, "Status").Value))
            If Len(s) > 0 And StrComp(s, "OK", vbTextCompare) <> 0 Then
                c = c + 1
                If c <= 10 Then
                    SheetProblems = SheetProblems & "  " & CStr(CellIn(lo, i, "Job ID").Value) & " - " & s & vbCrLf
                End If
            End If
        End If
    Next i
    If c > 10 Then
        SheetProblems = SheetProblems & "  ... and " & (c - 10) & " more." & vbCrLf
    End If
End Function

Private Sub SetStatus(ByVal ws As Worksheet, ByVal Text As String)
    Dim r As Range
    Set r = LocRange(ws, "LOC_Status")
    If Not r Is Nothing Then r.Value = Text
End Sub
