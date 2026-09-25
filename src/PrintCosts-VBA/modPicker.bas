Attribute VB_Name = "modPicker"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Multi-select for the two many-to-many fields: which paper families a printer
' supports, and which printers a print room has.
'
' A worksheet standing in for a dialog. The design document called for a
' UserForm, but a UserForm is a .frm plus a binary .frx, and a .frx cannot be
' written as text - the form would have to be rebuilt by hand before anything
' worked. This ships inside the workbook, uses no ActiveX, and behaves the same
' on Windows and Mac.
'
' It cannot be modal, because VBA cannot block waiting on a worksheet. So it is
' styled to look temporary, and it closes itself if the user navigates away.

' Not named SH: Excel's sheet events pass a parameter called Sh, and VBA
' identifiers are case-insensitive, so a constant called SH is shadowed
' inside any handler that takes one.
Private Const PICKER_SHEET As String = "_Picker"
Private Const COL_ITEM As Long = 2
Private Const COL_TICK As Long = 3
Private Const COL_STATE As Long = 8
Private Const R_TITLE As Long = 2
Private Const R_HELP As Long = 3
Private Const R_HEAD As Long = 5
Private Const FIRST_ROW As Long = 6

Public Sub PickPrinters(ByVal ws As Worksheet)
    Dim lo As ListObject, i As Long, items As Collection, chosen As String
    Set items = New Collection
    Set lo = Tbl("tblPrinters")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Model").Value) > 0 Then
            If StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0 Then
                items.Add CStr(CellIn(lo, i, "Model").Value)
            End If
        End If
    Next i
    chosen = LocValue(ws, "LOC_Printers")
    Show "PRINTERS", ws.Name, ws.Name, items, chosen, "Printers at " & LocValue(ws, "LOC_Name"), "Tick every printer this print room has. Only these can be chosen for its print jobs."
End Sub

Public Sub PickFamilies()
    Dim lo As ListObject, i As Long, items As Collection, n As Long, chosen As String
    If StrComp(ActiveSheet.Name, "Printers", vbTextCompare) <> 0 Then
        Say "Select families works on the Printers sheet.", "It sets which paper families the selected printer can use.", "Go to the Printers sheet, click the printer's row, then try again."
        Exit Sub
    End If

    Set lo = Tbl("tblPrinters")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Application.Intersect(Selection.Cells(1, 1).EntireRow, lo.DataBodyRange) Is Nothing Then
        Say "No printer row is selected.", "This sets the paper families for one printer.", "Click any cell in that printer's row, then try again."
        Exit Sub
    End If
    n = Selection.Cells(1, 1).Row - lo.DataBodyRange.Row + 1

    Set items = New Collection
    Dim fl As ListObject
    Set fl = Tbl("tblPaperFamilies")
    For i = 1 To fl.ListRows.Count
        If Len(CellIn(fl, i, "Family").Value) > 0 Then items.Add CStr(CellIn(fl, i, "Family").Value)
    Next i

    chosen = CStr(CellIn(lo, n, "Supported families").Value)
    Show "FAMILIES", CStr(n), "Printers", items, chosen, "Paper families for " & CStr(CellIn(lo, n, "Model").Value), "Tick every paper family this printer can print on. A stock can only be used on a printer that supports its family."
End Sub

Private Sub Show(ByVal Mode As String, ByVal Target As String, ByVal ReturnTo As String, ByVal items As Collection, ByVal Chosen As String, ByVal Title As String, ByVal Help As String)
    Dim ws As Worksheet, i As Long, last As Long, cb As CheckBox, c As Range

    Set ws = ThisWorkbook.Worksheets(PICKER_SHEET)
    AppOff
    UnlockSheet ws

    ws.Cells.Clear
    ClearControls ws

    ws.Cells(1, COL_STATE).Value = Mode
    ws.Cells(2, COL_STATE).Value = Target
    ws.Cells(3, COL_STATE).Value = ReturnTo
    ws.Columns(COL_STATE).Hidden = True

    ws.Cells(R_TITLE, COL_ITEM).Value = Title
    ws.Cells(R_TITLE, COL_ITEM).Font.Size = 14
    ws.Cells(R_TITLE, COL_ITEM).Font.Bold = True
    ws.Cells(R_TITLE, COL_ITEM).Font.Color = RGB(31, 56, 100)
    ws.Cells(R_HELP, COL_ITEM).Value = Help
    ws.Cells(R_HELP, COL_ITEM).Font.Italic = True
    ws.Cells(R_HELP, COL_ITEM).Font.Color = RGB(128, 128, 128)

    ws.Cells(R_HEAD, COL_ITEM).Value = "Option"
    ws.Cells(R_HEAD, COL_TICK).Value = "Include"
    ws.Range(ws.Cells(R_HEAD, COL_ITEM), ws.Cells(R_HEAD, COL_TICK)).Font.Bold = True
    ws.Range(ws.Cells(R_HEAD, COL_ITEM), ws.Cells(R_HEAD, COL_TICK)).Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' The option reads first and the tick sits to its right, which is the order
    ' people expect on a form.
    For i = 1 To items.Count
        ws.Cells(FIRST_ROW + i - 1, COL_ITEM).Value = items(i)
        Set c = ws.Cells(FIRST_ROW + i - 1, COL_TICK)
        Set cb = ws.CheckBoxes.Add(c.Left + (c.Width - 16) / 2, c.Top + 2, 16, 14)
        cb.Caption = ""
        cb.Name = "pick_" & (FIRST_ROW + i - 1)
        cb.Value = IIf(InList(Chosen, items(i)), xlOn, xlOff)
        cb.Locked = False
    Next i
    last = FIRST_ROW + items.Count - 1

    DrawOne ws, last + 2, COL_ITEM, "OK", "btnPickerOK", 90
    DrawOne ws, last + 2, COL_TICK, "Cancel", "btnPickerCancel", 90

    Frame ws, last

    ws.Columns(1).ColumnWidth = 3
    ws.Columns(COL_ITEM).ColumnWidth = 46
    ws.Columns(COL_TICK).ColumnWidth = 12
    ws.Columns(4).ColumnWidth = 3
    ws.Tab.Color = RGB(255, 192, 0)

    ws.Visible = xlSheetVisible
    RelockSheet ws
    AppOn
    ws.Activate
    ws.Cells(R_TITLE, COL_ITEM).Select
End Sub

' A tinted, bordered panel so the sheet reads as a temporary dialog rather than
' as part of the workbook.
Private Sub Frame(ByVal ws As Worksheet, ByVal LastItemRow As Long)
    Dim box As Range, inner As Range
    Set box = ws.Range(ws.Cells(1, COL_ITEM - 1), ws.Cells(LastItemRow + 4, COL_TICK + 1))
    box.Interior.Color = RGB(252, 252, 248)
    With box.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(191, 191, 191)
    End With
    box.Borders(xlInsideVertical).LineStyle = xlNone
    box.Borders(xlInsideHorizontal).LineStyle = xlNone
    Set inner = ws.Range(ws.Cells(FIRST_ROW, COL_ITEM), ws.Cells(LastItemRow, COL_TICK))
    inner.Interior.Color = RGB(255, 255, 255)
End Sub

Private Sub ClearControls(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.CheckBoxes.Count To 1 Step -1
        ws.CheckBoxes(i).Delete
    Next i
    For i = ws.Buttons.Count To 1 Step -1
        ws.Buttons(i).Delete
    Next i
End Sub

Public Sub btnPickerOK()
    On Error GoTo Fail
    Dim ws As Worksheet, cb As CheckBox, r As Long, out As String
    Dim mode As String, target As String, picked As Collection, k As Variant

    Set ws = ThisWorkbook.Worksheets(PICKER_SHEET)
    mode = CStr(ws.Cells(1, COL_STATE).Value)
    target = CStr(ws.Cells(2, COL_STATE).Value)

    ' Read in row order rather than control order, so the stored list matches
    ' what the user saw.
    Set picked = New Collection
    For r = FIRST_ROW To FIRST_ROW + 500
        If Len(ws.Cells(r, COL_ITEM).Value) = 0 Then Exit For
        For Each cb In ws.CheckBoxes
            If cb.TopLeftCell.Row = r Then
                If cb.Value = xlOn Then picked.Add CStr(ws.Cells(r, COL_ITEM).Value)
                Exit For
            End If
        Next cb
    Next r
    For Each k In picked
        If Len(out) > 0 Then out = out & LIST_SEP
        out = out & CStr(k)
    Next k

    AppOff
    If mode = "PRINTERS" Then
        Dim t As Worksheet
        Set t = ThisWorkbook.Worksheets(target)
        UnlockSheet t
        LocRange(t, "LOC_Printers").Value = out
        RelockSheet t
        BindColumns t
    ElseIf mode = "FAMILIES" Then
        Dim lo As ListObject
        Set lo = Tbl("tblPrinters")
        UnlockSheet ThisWorkbook.Worksheets("Printers")
        CellIn(lo, CLng(target), "Supported families").Value = out
        RelockSheet ThisWorkbook.Worksheets("Printers")
        Invalidate
    End If
    AppOn
    Close_
    Exit Sub
Fail:
    AppReset
    ReportError "applying your selection"
End Sub

Public Sub btnPickerCancel()
    Close_
End Sub

' Called when any other sheet is activated. Leaving the picker behind in the
' tab list would be confusing, so navigating away is treated as Cancel.
' Called when any other sheet is activated. Leaving the picker behind in the
' tab list would be confusing, so navigating away is treated as Cancel.
Public Sub HideIfLeft(ByVal ActivatedSheet As Object)
    Dim ws As Worksheet
    On Error GoTo Fail
    If StrComp(ActivatedSheet.Name, PICKER_SHEET, vbTextCompare) = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(PICKER_SHEET)
    If ws.Visible <> xlSheetVisible Then Exit Sub
    ' The user has already chosen where to go, so do not drag them back to
    ' wherever the picker was opened from.
    Close_ False
    Exit Sub
Fail:
    Application.EnableEvents = True
End Sub

Private Sub Close_(Optional ByVal GoBack As Boolean = True)
    Dim ws As Worksheet, back As String, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PICKER_SHEET)
    back = CStr(ws.Cells(3, COL_STATE).Value)
    ev = Application.EnableEvents
    Application.EnableEvents = False
    If GoBack And Len(back) > 0 Then
        ' A sheet cannot be hidden while it is the active one, so move first.
        ThisWorkbook.Worksheets(back).Activate
    End If
    ws.Visible = xlSheetVeryHidden
    Application.EnableEvents = ev
End Sub
