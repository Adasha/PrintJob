Attribute VB_Name = "modDepartments"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' The Departments sheet (0.10.30): other departments whose print jobs are
' tracked but not charged.
'
' A job is tied to a department by its Student Name matching a department's
' Name or one of its Aliases (modCatalog.DepartmentFor - the single lookup, so
' that a later real Department column on the job row only has to change that
' function). When such a name is ENTERED on a job and the department is Free
' (and Active), both disregard columns are set to Yes: the job is costed in
' full - Gross is unchanged - but Chargeable is zero, so Reports and Summary
' still show the cost as "Disregarded". That is applied once, at the moment the
' name is typed or picked; nothing re-applies it afterwards, so a person can
' change either disregard by hand and it stays changed.
'
' Built in VBA like _Registry/_Audit rather than shipped in PrintCosts.xlsx, so
' the file stays reproducible and no hand-edited .xlsx change is needed.
' tblDepartments is shaped for what comes next (per-department charging
' percentages, allowances): the last four columns are RESERVED - present so the
' table's shape is stable, grey, and not read by anything yet.

Public Const DEPT_SHEET As String = "Departments"
Public Const DEPT_TABLE As String = "tblDepartments"
Private Const HDR_ROW As Long = 6
Private Const NAVY As Long = 6567967        ' RGB(31,56,100) - 1F3864, the config-sheet title/header colour
Private Const GREY_TXT As Long = 8421504    ' RGB(128,128,128)
Private Const GREY_FILL As Long = 15921906  ' RGB(242,242,242)
Private Const RESERVED_FILL As Long = 8355711 ' RGB(127,127,127)

Public Sub EnsureDepartmentsSheet()
    Dim ws As Worksheet, lo As ListObject

    Set lo = Tbl(DEPT_TABLE)
    If lo Is Nothing Then
        BuildDepartmentsTable
        Set lo = Tbl(DEPT_TABLE)
    End If
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    EnsureDeptIdSetting
    ApplyDepartmentsFormat ws, lo
    RelockSheet ws
End Sub

Private Sub BuildDepartmentsTable()
    Dim ws As Worksheet, lo As ListObject, hdrs As Variant, i As Long
    Dim prev As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEPT_SHEET)
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
        ws.Name = DEPT_SHEET
    End If
    UnlockSheet ws

    hdrs = Array("DeptID", "Name", "Aliases", "Free", "Active", "Notes", "Match list", _
                 "Dis Paper %", "Dis Cons %", "Allowance", "Allowance period")
    For i = LBound(hdrs) To UBound(hdrs)
        ws.Cells(HDR_ROW, i - LBound(hdrs) + 1).Value = hdrs(i)
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + 1, UBound(hdrs) - LBound(hdrs) + 1)), , xlYes)
    lo.Name = DEPT_TABLE
    lo.TableStyle = "TableStyleLight9"

    ' Lower-cased, ';'-delimited Name + Aliases, so a formula can test a whole
    ' word with FIND(";"&name&";", [Match list]). Spaces around the ';' are
    ' tolerated; OnDepartmentEdited tidies what a person types as well.
    lo.ListColumns("Match list").DataBodyRange.Formula = _
        "=IF(TRIM([@Name])="""","""",LOWER("";""&TRIM([@Name])&"";""&SUBSTITUTE(SUBSTITUTE(TRIM([@Aliases]),"" ;"","";""),""; "","";"")&"";""))"
End Sub

' Re-applied on every setup run (same self-healing stance as the other
' Ensure* routines), so the look and the dropdowns can never drift.
Private Sub ApplyDepartmentsFormat(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim c As Long, w As Variant

    ws.Cells.Font.Name = "Calibri"
    With ws.Range("A1")
        .Value = "Departments"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = NAVY
    End With
    ws.Rows(1).RowHeight = 21
    With ws.Range("A2")
        .Value = "Other departments whose print jobs are tracked here but not charged. A job whose Student Name is a department's Name or one of its Aliases (separate aliases with ;) is recognised as that department."
        .Font.Size = 10: .Font.Italic = True: .Font.Color = GREY_TXT
    End With
    With ws.Range("A3")
        .Value = "Free = Yes sets both disregard columns to Yes when the name is entered on a job: it is still costed, but not charged. Change either by hand and it stays changed. Set a department to Inactive rather than deleting it."
        .Font.Size = 10: .Font.Italic = True: .Font.Color = GREY_TXT
    End With

    With lo.HeaderRowRange
        .Font.Bold = True: .Font.Color = vbWhite
        .Interior.Color = NAVY
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ' Reserved columns: visibly "not yet".
    lo.ListColumns("Dis Paper %").Range.Cells(1, 1).Resize(1, 4).Interior.Color = RESERVED_FILL
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.ListColumns("Dis Paper %").DataBodyRange.Resize(, 4).Font.Color = GREY_TXT

    With lo.ListColumns("DeptID").DataBodyRange
        .Interior.Color = GREY_FILL: .Font.Color = GREY_TXT
    End With

    w = Array(14, 30, 30, 9, 9, 40, 14, 12, 12, 12, 16)
    For c = 0 To UBound(w)
        ws.Columns(c + 1).ColumnWidth = w(c)
    Next c
    ws.Columns(7).Hidden = True     ' Match list: working column

    AddYesNo lo, "Free", "Free", "Yes: a job for this department starts with Disregard Paper and Disregard Consumable both set to Yes (it is still costed, not charged). No: the department is recognised for reporting only."
    AddYesNo lo, "Active", "Active", "Yes: the department is recognised when a name is entered. Set to No instead of deleting it - past jobs stay grouped under it in Reports."

    With lo.ListColumns("Aliases").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputTitle = "Aliases"
        .InputMessage = "Other spellings of this department, separated by ; for example: Fine Art;FA. Matching ignores capitals and extra spaces."
        .ShowInput = True
    End With
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

Private Sub AddYesNo(ByVal lo As ListObject, ByVal Header As String, ByVal Title As String, ByVal Msg As String)
    With lo.ListColumns(Header).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Choose Yes or No."
    End With
End Sub

' DEPT_ID_HWM is a Settings row like TECH_ID_HWM (the highest DeptID issued,
' only ever rises). Settings rows normally ship in the .xlsx and no setup
' routine adds one; this is the exception because the Departments sheet is
' itself built in VBA, and an upgraded workbook must gain both together.
Private Sub EnsureDeptIdSetting()
    Dim lo As ListObject, i As Long, pos As Long, stk As Long, k As String
    Dim r As ListRow, ws As Worksheet, c As Range

    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    UnlockSheet ws

    For i = 1 To lo.ListRows.Count
        k = Trim$(CStr(CellIn(lo, i, "Key").Value))
        If StrComp(k, "DEPT_ID_HWM", vbTextCompare) = 0 Then pos = i
        If StrComp(k, "STOCK_ID_HWM", vbTextCompare) = 0 Then stk = i
    Next i

    If pos = 0 Then
        If stk > 0 Then Set r = lo.ListRows.Add(stk + 1) Else Set r = lo.ListRows.Add
        pos = r.Index
        CellIn(lo, pos, "Key").Value = "DEPT_ID_HWM"
        CellIn(lo, pos, "Setting").Value = "Last department ID number"
        CellIn(lo, pos, "Value").Value = 0
        CellIn(lo, pos, "Notes").Value = "Read-only. The highest DeptID number issued at this site. Only ever rises, so an ID is never reused."
    End If

    Set c = Nothing
    On Error Resume Next
    Set c = ThisWorkbook.Names("SET_DEPT_ID_HWM").RefersToRange
    On Error GoTo 0
    If c Is Nothing Then
        ThisWorkbook.Names.Add Name:="SET_DEPT_ID_HWM", RefersTo:="='" & ws.Name & "'!" & CellIn(lo, pos, "Value").Address
    End If
    RelockSheet ws
End Sub

' ------------------------------------------------------------ job entry ---
' Called from modValidation.OnCellChanged when a job's Student Name is edited
' (a single-cell edit - a multi-cell paste never reaches the per-row handlers,
' and RepeatJob/import/re-stamp/the Planner write cells with events off, so
' none of those apply it). Events are already off. Silent: the cells changing
' to Yes is the feedback.
'
' Only ever sets Yes. A name that is not (or no longer) a free department
' leaves the disregards exactly as they are.
Public Sub StampDepartmentDisregards(ByVal lo As ListObject, ByVal n As Long)
    Dim d As clsDept
    Set d = FreeDepartmentFor(CStr(CellIn(lo, n, "Student Name").Value))
    If Not d.Found Then Exit Sub
    CellIn(lo, n, "Disregard Paper").Value = "Yes"
    CellIn(lo, n, "Disregard Consumable").Value = "Yes"
End Sub

' ------------------------------------------------------------ sheet edit ---
' From ThisWorkbook.Workbook_SheetChange for an edit on the Departments sheet,
' events off. Gives a row typed under the table its ID and its Active/Free
' defaults, tidies the Aliases list, and warns (never blocks) when a name or
' alias is already claimed by another department.
Public Sub OnDepartmentEdited(ByVal Target As Range)
    Dim lo As ListObject, hit As Range, c As Range, n As Long, lastN As Long
    Dim nameCol As Long, aliasCol As Long, clash As String, tidy As String, cur As String

    Set lo = Tbl(DEPT_TABLE)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Application.Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub

    OnCatalogEdited DEPT_TABLE, Target      ' DeptID

    nameCol = ColIdx(lo, "Name"): aliasCol = ColIdx(lo, "Aliases")
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If n <> lastN Then
            lastN = n
            If Len(Trim$(CStr(CellIn(lo, n, "Name").Value))) > 0 Then
                DefaultActive lo, n
                DefaultFree lo, n
            End If
        End If
        If c.Column - lo.Range.Column + 1 = aliasCol Then
            cur = CStr(c.Value)
            tidy = TidyAliases(cur)
            If StrComp(cur, tidy, vbBinaryCompare) <> 0 Then c.Value = tidy
        End If
        If c.Column - lo.Range.Column + 1 = nameCol Or c.Column - lo.Range.Column + 1 = aliasCol Then
            clash = clash & ClashText(lo, n)
        End If
    Next c

    If Len(clash) > 0 Then
        Say "A department name or alias is used more than once.", clash, _
            "Only the first department listed with a name or alias is matched. Change one of them so each is used once."
    End If
End Sub

' "Fine Art ; FA;;  fa" -> "Fine Art;FA;fa": trimmed, no empties.
Private Function TidyAliases(ByVal s As String) As String
    Dim parts As Variant, i As Long, out As String, a As String
    parts = Split(s, ";")
    For i = LBound(parts) To UBound(parts)
        a = Trim$(CStr(parts(i)))
        If Len(a) > 0 Then
            If Len(out) > 0 Then out = out & ";"
            out = out & a
        End If
    Next i
    TidyAliases = out
End Function

' One line per name/alias of row n that another row also claims.
Private Function ClashText(ByVal lo As ListObject, ByVal n As Long) As String
    Dim keys As Collection, k As Variant, i As Long, parts As Variant, j As Long, hit As String

    Set keys = New Collection
    If Len(Trim$(CStr(CellIn(lo, n, "Name").Value))) = 0 Then Exit Function
    keys.Add Trim$(CStr(CellIn(lo, n, "Name").Value))
    parts = Split(CStr(CellIn(lo, n, "Aliases").Value), ";")
    For j = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(j)))) > 0 Then keys.Add Trim$(CStr(parts(j)))
    Next j

    For Each k In keys
        For i = 1 To lo.ListRows.Count
            If i <> n Then
                hit = ""
                If StrComp(Trim$(CStr(CellIn(lo, i, "Name").Value)), CStr(k), vbTextCompare) = 0 Then
                    hit = "name"
                Else
                    parts = Split(CStr(CellIn(lo, i, "Aliases").Value), ";")
                    For j = LBound(parts) To UBound(parts)
                        If StrComp(Trim$(CStr(parts(j))), CStr(k), vbTextCompare) = 0 Then hit = "alias"
                    Next j
                End If
                If Len(hit) > 0 Then
                    ClashText = ClashText & "'" & CStr(k) & "' is also the " & hit & " of '" & Trim$(CStr(CellIn(lo, i, "Name").Value)) & "'." & vbCrLf
                    Exit For
                End If
            End If
        Next i
    Next k
End Function
