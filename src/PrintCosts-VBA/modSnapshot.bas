Attribute VB_Name = "modSnapshot"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Freezing the rates a record was costed at (design doc section 5).
'
' This is what makes spec 12.6 and AT-09 hold. No formula on a job row reaches
' into tblPapers or tblPrinters; the row reads only its own inputs and the
' values stamped here. A price change therefore has no path back into history.

Public Sub StampRow(ByVal ws As Worksheet, ByVal RowNo As Long)
    Dim lo As ListObject, s As clsStock, p As clsPrinterDef
    Dim stockW As Double, stockH As Double
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    Set p = Prn(CStr(CellIn(lo, RowNo, "Printer").Value))
    Set s = Stock(CStr(CellIn(lo, RowNo, "Paper Stock").Value))

    CellIn(lo, RowNo, "S_PrinterID").Value = p.PrinterID
    CellIn(lo, RowNo, "S_ConsRate").Value = IIf(p.Found, p.RatePerM2, Empty)

    CellIn(lo, RowNo, "S_StockID").Value = s.StockID
    CellIn(lo, RowNo, "S_Measure").Value = s.Measure
    CellIn(lo, RowNo, "S_UnitCost").Value = IIf(s.Found, s.Cost, Empty)

    stockW = s.WidthMM
    stockH = s.HeightMM
    ' The built-in Supplied (Sheet) stock has no catalogue size - the row's own
    ' "Sheet size" pick (the nearest standard size) is what the Area
    ' formula's Sheet branch needs stamped here instead. Roll needs no
    ' equivalent: Print Width mm is required for 'Supplied (Roll)'
    ' (modValidation), and the Area formula always uses it in preference to
    ' S_StockWidth_mm once it's set, so a stamped 0 there is never read.
    If s.Found And s.PerJobSize And s.Measure = "Sheet" Then
        StdSizeDims CStr(CellIn(lo, RowNo, "Sheet size").Value), stockW, stockH
    End If

    CellIn(lo, RowNo, "S_StockWidth_mm").Value = IIf(s.Found, stockW, Empty)
    If s.Found And s.Measure = "Sheet" Then
        CellIn(lo, RowNo, "S_SheetHeight_mm").Value = stockH
    Else
        CellIn(lo, RowNo, "S_SheetHeight_mm").ClearContents
    End If

    CellIn(lo, RowNo, "S_TechID").Value = TechID(CStr(CellIn(lo, RowNo, "Technician").Value))
    CellIn(lo, RowNo, "S_StampedAt").Value = Now
    CellIn(lo, RowNo, "S_StampedBy").Value = CurrentUser
    CellIn(lo, RowNo, "S_SchemaVer").Value = SCHEMA_VER
End Sub

' Re-stamping exists for one case: a price was entered wrongly and records were
' logged against it. It is deliberately kept away from the location toolbars -
' it lives on Settings, confirms with a count, and writes an audit entry.
'
' It covers every print room in one go, because a mistyped price is a property
' of the configuration rather than of one room, and correcting it room by room
' would leave the workbook half-restated.
Public Sub ReStampAll()
    Dim ws As Worksheet, lo As ListObject, i As Long, n As Long, rooms As Long

    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            Set lo = JobsTable(ws)
            If Not lo Is Nothing Then
                n = n + RowCount(lo)
                rooms = rooms + 1
            End If
        End If
    Next ws

    If n = 0 Then
        Say "There are no records to re-stamp."
        Exit Sub
    End If

    If Not Ask("Re-stamp all " & n & " record" & IIf(n = 1, "", "s") & " across " & rooms & " print room" & IIf(rooms = 1, "", "s") & " at today's configured prices?" & vbCrLf & vbCrLf & "Every historical cost will be recalculated at current rates. Do this " & "only if a price was originally entered wrongly - otherwise it " & "rewrites history that was correct." & vbCrLf & vbCrLf & "This cannot be undone.", "Re-stamp prices") Then Exit Sub

    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then
            Set lo = JobsTable(ws)
            If Not lo Is Nothing Then
                For i = 1 To lo.ListRows.Count
                    If Not IsBlankRow(lo, i) Then StampRow ws, i
                Next i
            End If
        End If
    Next ws
    AppOn

    LogAudit "Re-stamp", "(all print rooms)", n & " records re-stamped at current prices"
    Say n & " record" & IIf(n = 1, "", "s") & " re-stamped.", "Their costs now reflect the prices currently set on the Papers and Printers sheets.", "Check a few rows against what you expected before relying on the figures."
End Sub

' Append-only record of destructive acts. VBA clears Excel's undo stack, so
' without this a mistaken deletion would be both unrecoverable and untraceable.
Public Sub LogAudit(ByVal Action As String, ByVal Location As String, ByVal Detail As String)
    Dim lo As ListObject, r As ListRow, aws As Worksheet
    ' Not eNum: VBA identifiers are case-insensitive, and "eNum" collapses to
    ' the reserved word Enum - a Dim of that name is a compile error (Syntax
    ' error), not a warning, and it never surfaced before because nothing
    ' called LogAudit until modImport did.
    Dim errNum As Long, errDesc As String

    ' The lookup is the only part entitled to fail quietly. A workbook with no
    ' audit table should not stop a deletion the user has already confirmed.
    On Error Resume Next
    Set lo = Tbl("tblAudit")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    Set aws = lo.Parent

    ' The writes are not. One blanket handler over the whole body is the
    ' pattern design 8.2 names after StampProperties: five independent writes
    ' report nothing when they fail, so a truncated audit entry looks exactly
    ' like a complete one - and a RelockSheet that never runs leaves _Audit
    ' unprotected for the rest of the session.
    On Error GoTo Fail
    UnlockSheet aws
    If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
        Set r = lo.ListRows(1)
    Else
        Set r = lo.ListRows.Add
    End If
    r.Range.Cells(1, 1).Value = Now
    r.Range.Cells(1, 2).Value = CurrentUser
    r.Range.Cells(1, 3).Value = Action
    r.Range.Cells(1, 4).Value = Location
    r.Range.Cells(1, 5).Value = Detail
    RelockSheet aws
    Exit Sub

Fail:
    ' Read Err FIRST. Every form of On Error resets the Err object, and
    ' RelockSheet runs two of them, so reading Err after reprotecting reports
    ' "0: " and loses the only description of what actually went wrong.
    errNum = Err.Number
    errDesc = Err.Description

    ' Reprotect whatever state we got to, then say so. This reports rather
    ' than re-raising: the caller is midway through a destructive operation
    ' the user has confirmed, and losing the audit note is worth knowing about
    ' but is not a reason to abandon the deletion half-done.
    If Not aws Is Nothing Then RelockSheet aws
    Say "The print job was removed, but the audit note could not be written.", _
        errNum & ": " & errDesc, _
        "The records themselves are unaffected. Pass this on to whoever maintains the workbook."
End Sub
