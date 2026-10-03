Attribute VB_Name = "modValidation"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

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

    ' Roll length unit (A6/B6): rescale the table's existing roll lengths.
    If Target.Cells.Count = 1 Then
        If Not LocRange(ws, "LOC_RollUnit") Is Nothing Then
            If Not Application.Intersect(Target, LocRange(ws, "LOC_RollUnit")) Is Nothing Then
                ApplyRollUnitChange ws
                RefreshQtyValidation ws
                Exit Function
            End If
        End If
    End If

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
        Case "Qty"
        Case "Print Width mm"
            OnWidthChanged ws, lo, n
        Case "Sheet size"
            OnSheetSizeChanged ws, lo, n
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

    model = Clean(prnCell)
    stk = Clean(stkCell)
    If Len(model) > 0 And Len(stk) > 0 Then
        If Not Compatible(model, stk) Then
            Set s = Stock(stk)
            Set p = Prn(model)
            If changedPrinter Then
                stkCell.ClearContents
                Say "'" & stk & "' cannot be used on " & model & ".", "That printer takes " & CapacityText(p) & ".", "The default paper stock has been cleared."
            Else
                prnCell.ClearContents
                Say "'" & stk & "' cannot be used on " & model & ".", "'" & stk & "' doesn't fit what that printer can take (" & CapacityText(p) & ").", "The default printer has been cleared."
            End If
        End If
    End If

    BindDefaultCells ws
End Function

' Defensive re-check called from modJobs.AddPrintJob immediately before it
' copies the batch defaults into a new row. OnDefaultCellChanged (above)
' normally catches an incompatible printer/paper default the instant it is
' set, but Excel for Mac does not reliably fire Worksheet_Change for a value
' picked off an in-cell dropdown list - so a default chosen that way can sit
' incompatible, still carrying modLists.UNAVAILABLE_SUFFIX in the cell text,
' until something else happens to edit it. Add Print Job cannot depend on
' that event having fired, so it calls this to clean and re-check the
' defaults itself right before copying them into the new row.
Public Sub EnsureDefaultsClean(ByVal ws As Worksheet)
    Dim prnCell As Range, stkCell As Range, model As String, stk As String, s As clsStock, p As clsPrinterDef
    Set prnCell = LocRange(ws, "LOC_DefPrinter")
    Set stkCell = LocRange(ws, "LOC_DefPaper")
    If prnCell Is Nothing Or stkCell Is Nothing Then Exit Sub

    model = Clean(prnCell)
    stk = Clean(stkCell)
    If Len(model) > 0 And Len(stk) > 0 Then
        If Not Compatible(model, stk) Then
            Set s = Stock(stk)
            Set p = Prn(model)
            stkCell.ClearContents
            Say "'" & stk & "' cannot be used on " & model & ".", "That printer takes " & CapacityText(p) & ".", "The default paper stock has been cleared before adding this job."
            BindDefaultCells ws
        End If
    End If
End Sub

' Reads a Printer/Paper Stock cell (table row or default cell alike) and
' strips modLists.UNAVAILABLE_SUFFIX if the value just picked off the
' dropdown carries it (2026-09-23: the dropdown now lists every item, marking
' - not removing - ones the other field rules out). Rewrites the cell to the
' clean value so nothing downstream ever sees the marker; a value with no
' marker round-trips unchanged.
Private Function Clean(ByVal c As Range) As String
    Dim raw As String
    raw = CStr(c.Value)
    Clean = CleanPick(raw)
    If Clean <> raw Then c.Value = Clean
End Function

' Either field can drive the other now (spec 1a). Both handlers end by
' rebinding BOTH cells' lists rather than just the other one: whichever field
' just changed may itself need widening back out, e.g. when the other field
' had to be cleared for incompatibility, or when this field was cleared
' outright and the other field's list was previously narrowed by it.
Private Sub OnPrinterChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim model As String, stk As String, s As clsStock, p As clsPrinterDef
    model = Clean(CellIn(lo, n, "Printer"))
    stk = Clean(CellIn(lo, n, "Paper Stock"))

    ' A printer change can strand a stock that was valid a moment ago (AT-05).
    If Len(stk) > 0 And Len(model) > 0 Then
        Set p = Prn(model)
        If Not Compatible(model, stk) Then
            Set s = Stock(stk)
            CellIn(lo, n, "Paper Stock").ClearContents
            Say "'" & stk & "' cannot be used on " & model & ".", "That printer takes " & CapacityText(p) & ".", "The paper stock has been cleared. Choose a stock this printer supports."
        Else
            ' Compatible (roll/sheet-capable at all) doesn't mean the row's
            ' OWN entered size still fits, for student-supplied stock - the
            ' mirror of the order-dependency class of bug already fixed
            ' once in this module for Qty/centimetres (§16.3): the size was
            ' checked against the printer at the time it was entered, and
            ' nothing else re-checks it if the printer changes afterward.
            RevalidateSuppliedSize ws, lo, n, model
        End If
    End If

    BindStockCell ws, lo, n
    BindPrinterCell ws, lo, n
    StampRow ws, n
End Sub

' See OnPrinterChanged's own comment above for why this exists.
Private Sub RevalidateSuppliedSize(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long, ByVal Model As String)
    Dim stk As String, s As clsStock, p As clsPrinterDef
    Dim w As Double, h As Double, sizeName As String

    stk = CStr(CellIn(lo, n, "Paper Stock").Value)
    If Len(stk) = 0 Then Exit Sub
    Set s = Stock(stk)
    If Not s.PerJobSize Then Exit Sub
    Set p = Prn(Model)
    If Not p.Found Then Exit Sub

    If s.Measure = "Roll" Then
        If p.MaxRollWidthMM > 0 Then
            w = NumOf(CellIn(lo, n, "Print Width mm"))
            If w > 0 And w > p.MaxRollWidthMM Then
                CellIn(lo, n, "Print Width mm").ClearContents
                Say "Print width " & Format$(w, "#,##0") & " mm is wider than " & Model & " can take.", Model & "'s maximum roll width is " & Format$(p.MaxRollWidthMM, "#,##0") & " mm.", "Enter " & Format$(p.MaxRollWidthMM, "#,##0") & " mm or less."
            End If
        End If
    ElseIf s.Measure = "Sheet" Then
        sizeName = Trim$(CStr(CellIn(lo, n, "Sheet size").Value))
        If Len(sizeName) > 0 And Len(p.MaxSheetSize) > 0 Then
            If StdSizeDims(sizeName, w, h) Then
                If Not FitsWithinMaxSheet(w, h, p.MaxSheetWidthMM, p.MaxSheetHeightMM) Then
                    CellIn(lo, n, "Sheet size").ClearContents
                    Say "'" & sizeName & "' is bigger than " & Model & " can take.", Model & "'s maximum sheet size is " & p.MaxSheetSize & ".", "Choose a smaller size."
                End If
            End If
        End If
    End If
End Sub

Private Sub OnStockChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim model As String, stk As String, s As clsStock, p As clsPrinterDef
    model = Clean(CellIn(lo, n, "Printer"))
    stk = Clean(CellIn(lo, n, "Paper Stock"))
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
            Say "'" & stk & "' cannot be used on " & model & ".", "That printer takes " & CapacityText(p) & ".", "Choose a stock that fits, or a different printer."
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

    ' Sheet size is meaningless for anything except 'Supplied (Sheet)' - the
    ' same tidy-up as Print Width mm just above, mirrored for the sheet case.
    If s.Measure <> "Sheet" Or Not s.PerJobSize Then
        If Len(CellIn(lo, n, "Sheet size").Value) > 0 Then
            CellIn(lo, n, "Sheet size").ClearContents
            Say "Sheet size only applies to 'Supplied (Sheet)'.", "This row's paper stock is now '" & stk & "'.", "The value has been cleared."
        End If
    End If
End Sub

' Called when LOC_RollUnit itself is edited. Rescales every existing roll
' job's Qty so the sheet shows the same lengths in the new unit - Sheet
' stock counts sheets, and a row with no stock chosen yet is left alone
' (nothing says it is a length). LOC_RollUnitApplied (modInit) records the
' unit the stored values are currently in, so re-selecting the unit already
' showing - which still fires Change - is a no-op and a value is never
' scaled twice. Rounded to 6 places to shed the floating-point residue of
' x100 / x0.01.
Public Sub ApplyRollUnitChange(ByVal ws As Worksheet)
    Dim lo As ListObject, was As String, now As String
    now = RollUnitOf(ws)
    was = AppliedRollUnit(ws)
    If StrComp(was, now, vbTextCompare) = 0 Then Exit Sub

    Set lo = JobsTable(ws)
    If Not lo Is Nothing Then
        If StrComp(now, "Centimetres", vbTextCompare) = 0 Then
            ScaleRollQty lo, 100
        Else
            ScaleRollQty lo, 0.01
        End If
    End If
    SetAppliedRollUnit ws, now
End Sub

Public Sub ScaleRollQty(ByVal lo As ListObject, ByVal Factor As Double)
    Dim i As Long, c As Range
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If CStr(CellIn(lo, i, "S_Measure").Value) = "Roll" Then
            Set c = CellIn(lo, i, "Qty")
            If Not IsEmpty(c.Value2) Then
                If IsNumeric(c.Value2) Then c.Value2 = Round(CDbl(c.Value2) * Factor, 6)
            End If
        End If
    Next i
End Sub

Private Sub OnWidthChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim w As Double, sw As Double, stk As String, s As clsStock, model As String, p As clsPrinterDef
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

    If s.PerJobSize Then
        ' No catalogue width to compare against - the material is whatever
        ' the student brought. The only real limit left is the printer
        ' itself, so AT-04's own check is applied against its capacity
        ' rather than a stock's nominal width.
        model = Trim$(CStr(CellIn(lo, n, "Printer").Value))
        If Len(model) > 0 Then
            Set p = Prn(model)
            If p.Found And p.MaxRollWidthMM > 0 Then
                If w > p.MaxRollWidthMM Then
                    CellIn(lo, n, "Print Width mm").ClearContents
                    Say "Print width " & Format$(w, "#,##0") & " mm is wider than " & model & " can take.", model & "'s maximum roll width is " & Format$(p.MaxRollWidthMM, "#,##0") & " mm.", "Enter " & Format$(p.MaxRollWidthMM, "#,##0") & " mm or less."
                End If
            End If
        End If
        Exit Sub
    End If

    sw = NumOf(CellIn(lo, n, "S_StockWidth_mm"))
    If sw = 0 Then sw = s.WidthMM

    ' AT-04.
    If w > sw Then
        CellIn(lo, n, "Print Width mm").ClearContents
        Say "Print width " & Format$(w, "#,##0") & " mm is wider than the stock.", "'" & stk & "' is " & Format$(sw, "#,##0") & " mm wide, so a print cannot " & "be wider than that.", "Enter " & Format$(sw, "#,##0") & " mm or less, or leave it blank to use " & "the full width of the roll."
    End If
End Sub

' The sheet-side mirror of OnWidthChanged's student-supplied branch above:
' 'Supplied (Sheet)' has no catalogue size either, so the row's own "Sheet
' size" pick (the nearest standard size) is what stands in for it, checked
' against the chosen printer's Max sheet size the same way Print Width mm
' is checked against Max roll width mm.
Private Sub OnSheetSizeChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim sizeName As String, stk As String, s As clsStock, model As String, p As clsPrinterDef
    Dim w As Double, h As Double

    sizeName = Trim$(CStr(CellIn(lo, n, "Sheet size").Value))
    If Len(sizeName) = 0 Then Exit Sub

    stk = CStr(CellIn(lo, n, "Paper Stock").Value)
    Set s = Stock(stk)
    If Not s.PerJobSize Or s.Measure <> "Sheet" Then
        CellIn(lo, n, "Sheet size").ClearContents
        Say "Sheet size only applies to 'Supplied (Sheet)'.", "This row's paper stock is '" & stk & "'.", "The value has been cleared."
        Exit Sub
    End If

    If Not StdSizeDims(sizeName, w, h) Then Exit Sub

    model = Trim$(CStr(CellIn(lo, n, "Printer").Value))
    If Len(model) > 0 Then
        Set p = Prn(model)
        If p.Found And Len(p.MaxSheetSize) > 0 Then
            If Not FitsWithinMaxSheet(w, h, p.MaxSheetWidthMM, p.MaxSheetHeightMM) Then
                CellIn(lo, n, "Sheet size").ClearContents
                Say "'" & sizeName & "' is bigger than " & model & " can take.", model & "'s maximum sheet size is " & p.MaxSheetSize & ".", "Choose a smaller size."
                Exit Sub
            End If
        End If
    End If

    ' Re-stamp so Area m2/Paper Cost pick up the newly-chosen dimensions
    ' immediately - StampRow reads this cell whenever the stock is
    ' 'Supplied (Sheet)' (modSnapshot).
    StampRow ws, n
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
