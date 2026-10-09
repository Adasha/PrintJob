Attribute VB_Name = "modConsumables"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' The Consumables sheet (0.11.0, multi-pass costing): the colours a multi-pass
' printer (a RISO duplicator, say) prints with, and what each costs per m2.
'
' Settings' tblConsumables stays the list of consumable TYPES (Ink, Toner,
' Risograph ...). A colour belongs to one type, and a pass row's Colour
' dropdown offers the colours of its printer's type (modCatalog.ColoursFor).
'
' Works like the Printers and Papers sheets: Add row / Remove row / Clear table
' buttons, a locked ID column (<SITE>-CLR-0001, counter COLOUR_ID_HWM), and
' deactivate-rather-than-delete. Colours are found by NAME, as printers and
' papers are (design decision 26a), so renaming a colour leaves pass rows that
' used the old name unable to find it - they keep their stamped rate.
'
' Built in VBA like the Departments sheet rather than shipped in the .xlsx.
' The table is deliberately plain - ID, Consumable type, Colour, Cost per m2,
' Active - and easy to extend (a colour swatch is a possible later column).

Public Const CONS_SHEET As String = "Consumables"
Public Const COLOUR_TABLE As String = "tblColours"
Private Const HDR_ROW As Long = 6
Private Const NAVY As Long = 6567967        ' RGB(31,56,100) - 1F3864, the config-sheet title/header colour
Private Const GREY_TXT As Long = 8421504    ' RGB(128,128,128)
Private Const GREY_FILL As Long = 15921906  ' RGB(242,242,242)

Public Sub EnsureConsumablesSheet()
    Dim ws As Worksheet, lo As ListObject

    Set lo = Tbl(COLOUR_TABLE)
    If lo Is Nothing Then
        BuildColoursTable
        Set lo = Tbl(COLOUR_TABLE)
    End If
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    EnsureColourIdSetting
    ApplyConsumablesFormat ws, lo
    RelockSheet ws
End Sub

Private Sub BuildColoursTable()
    Dim ws As Worksheet, lo As ListObject, hdrs As Variant, i As Long
    Dim prev As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error Resume Next
        Set prev = ThisWorkbook.Worksheets("Papers")
        On Error GoTo 0
        If prev Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=prev)
        End If
        ws.Name = CONS_SHEET
    End If
    UnlockSheet ws

    hdrs = Array("ColourID", "Consumable type", "Colour", "Cost per m2", "Active")
    For i = LBound(hdrs) To UBound(hdrs)
        ws.Cells(HDR_ROW, i - LBound(hdrs) + 1).Value = hdrs(i)
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + 1, UBound(hdrs) - LBound(hdrs) + 1)), , xlYes)
    lo.Name = COLOUR_TABLE
    lo.TableStyle = "TableStyleLight9"
End Sub

' Re-applied on every setup run (same self-healing stance as the other
' Ensure* routines), so the look and the dropdowns can never drift.
Private Sub ApplyConsumablesFormat(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim c As Long, w As Variant

    ws.Cells.Font.Name = "Calibri"
    With ws.Range("A1")
        .Value = "Consumables"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = NAVY
    End With
    ws.Rows(1).RowHeight = 21
    With ws.Range("A2")
        .Value = "The colours a multi-pass printer (a RISO duplicator, for example) prints with. The cost per square metre already accounts for typical coverage."
        .Font.Size = 10: .Font.Italic = True: .Font.Color = GREY_TXT
    End With
    With ws.Range("A3")
        .Value = "Each colour belongs to a consumable type (the list on Settings); a pass offers the colours of its printer's type. Set a colour to Inactive rather than deleting or renaming it - pass rows already recorded keep the rate they were stamped with."
        .Font.Size = 10: .Font.Italic = True: .Font.Color = GREY_TXT
    End With

    With lo.HeaderRowRange
        .Font.Bold = True: .Font.Color = vbWhite
        .Interior.Color = NAVY
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    w = Array(16, 20, 26, 14, 9)
    For c = 0 To UBound(w)
        ws.Columns(c + 1).ColumnWidth = w(c)
    Next c
    If lo.DataBodyRange Is Nothing Then Exit Sub

    With lo.ListColumns("ColourID").DataBodyRange
        .Interior.Color = GREY_FILL: .Font.Color = GREY_TXT
    End With
    lo.ListColumns("Cost per m2").DataBodyRange.NumberFormat = "#,##0.0000"

    With lo.ListColumns("Consumable type").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lstConsumables"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Consumable type"
        .InputMessage = "The kind of consumable this colour is - choose from the list on Settings (Risograph for a RISO duplicator). A multi-pass printer offers the colours of its own type."
        .ErrorTitle = "Consumable type"
        .ErrorMessage = "Choose a consumable type from the list on Settings."
    End With
    With lo.ListColumns("Cost per m2").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
        .IgnoreBlank = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Cost per m2"
        .InputMessage = "What this colour costs to print per square metre, ink and master included at typical coverage."
        .ErrorTitle = "Cost per m2"
        .ErrorMessage = "Enter a number, zero or more."
    End With
    With lo.ListColumns("Active").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Active"
        .InputMessage = "Yes: the colour is offered on new passes. Set to No instead of deleting it."
        .ErrorTitle = "Active"
        .ErrorMessage = "Choose Yes or No."
    End With
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

' COLOUR_ID_HWM is a Settings row like TECH_ID_HWM (the highest ColourID issued,
' only ever rises). Settings rows normally ship in the .xlsx and no setup routine
' adds one; this is the exception, for the same reason as DEPT_ID_HWM: the
' Consumables sheet it counts for is itself built in VBA.
Private Sub EnsureColourIdSetting()
    Dim lo As ListObject, i As Long, pos As Long, after As Long, k As String
    Dim r As ListRow, ws As Worksheet, c As Range

    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    UnlockSheet ws

    For i = 1 To lo.ListRows.Count
        k = Trim$(CStr(CellIn(lo, i, "Key").Value))
        If StrComp(k, "COLOUR_ID_HWM", vbTextCompare) = 0 Then pos = i
        If StrComp(k, "STOCK_ID_HWM", vbTextCompare) = 0 And after = 0 Then after = i
        If StrComp(k, "DEPT_ID_HWM", vbTextCompare) = 0 Then after = i
    Next i

    If pos = 0 Then
        If after > 0 Then Set r = lo.ListRows.Add(after + 1) Else Set r = lo.ListRows.Add
        pos = r.Index
        CellIn(lo, pos, "Key").Value = "COLOUR_ID_HWM"
        CellIn(lo, pos, "Setting").Value = "Last colour ID number"
        CellIn(lo, pos, "Value").Value = 0
        CellIn(lo, pos, "Notes").Value = "Read-only. The highest ColourID number issued at this site. Only ever rises, so an ID is never reused."
    End If

    Set c = Nothing
    On Error Resume Next
    Set c = ThisWorkbook.Names("SET_COLOUR_ID_HWM").RefersToRange
    On Error GoTo 0
    If c Is Nothing Then
        ThisWorkbook.Names.Add Name:="SET_COLOUR_ID_HWM", RefersTo:="='" & ws.Name & "'!" & CellIn(lo, pos, "Value").Address
    End If
    RelockSheet ws
End Sub

' ------------------------------------------------------------ sheet edit ---
' From ThisWorkbook.Workbook_SheetChange for an edit on the Consumables sheet,
' events off. Gives a row typed under the table its ID and its Active default,
' and warns (never blocks) when a colour name is used more than once - colours
' are looked up by name, so only the first would be found.
Public Sub OnColourEdited(ByVal Target As Range)
    Dim lo As ListObject, hit As Range, c As Range, n As Long, lastN As Long
    Dim nameCol As Long, clash As String, i As Long, nm As String

    Set lo = Tbl(COLOUR_TABLE)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Application.Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub

    OnCatalogEdited COLOUR_TABLE, Target      ' ColourID

    nameCol = ColIdx(lo, "Colour")
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If n <> lastN Then
            lastN = n
            If Len(Trim$(CStr(CellIn(lo, n, "Colour").Value))) > 0 Then DefaultActive lo, n
        End If
        If c.Column - lo.Range.Column + 1 = nameCol Then
            nm = Trim$(CStr(c.Value))
            If Len(nm) > 0 Then
                For i = 1 To lo.ListRows.Count
                    If i <> n Then
                        If StrComp(Trim$(CStr(CellIn(lo, i, "Colour").Value)), nm, vbTextCompare) = 0 Then
                            clash = clash & "'" & nm & "' is already used on another row." & vbCrLf
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
    Next c

    If Len(clash) > 0 Then
        Say "A colour name is used more than once.", clash, _
            "Colours are found by name, so only the first row listed with that name is used. Give each colour its own name."
    End If
End Sub
