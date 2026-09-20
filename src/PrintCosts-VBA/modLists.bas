Attribute VB_Name = "modLists"
Option Explicit

' Dropdown lists for the transaction table.
'
' Two design decisions worth knowing, both arrived at the hard way.
'
' 1. Validation is written when something CHANGES, never when a cell is
'    selected. Replacing a cell's validation closes an open dropdown, so
'    rebuilding from a SelectionChange handler made the arrow flash and vanish.
'
' 2. Every list gets its OWN staging column, claimed by a tag written in row 1.
'    A single shared column per list type meant the last sheet to be bound
'    overwrote the others, so one print room showed another's printers.
'
' Tags are: PRN|<sheet>   printers permitted at that print room
'           TEC           active technicians, the same everywhere
'           STK|<model>   stocks that printer can take - shared by every row
'                         using it, because the list depends only on the printer
'
' Staging is needed at all because a validation list supplied as a literal
' string is capped at 255 characters, which a real stock list exceeds.

Private Const STAGE_SHEET As String = "_Work"
Private Const FIRST_ITEM_ROW As Long = 2
Private Const MAX_COLS As Long = 200
Private Const MAX_ITEMS As Long = 500

' Only fills a gap: a stock cell that has never been given a list. It checks
' first and does nothing when one is present, so there is nothing to flash.
Public Sub OnSelection(ByVal ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject, n As Long, hdr As String
    If Target.Cells.Count > 1 Then Exit Sub
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Application.Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    hdr = CStr(lo.HeaderRowRange.Cells(1, Target.Column - lo.Range.Column + 1).Value)
    If hdr <> "Paper Stock" Then Exit Sub
    If HasValidation(Target) Then Exit Sub

    n = Target.Row - lo.DataBodyRange.Row + 1
    BindStockCell ws, lo, n
End Sub

Private Function HasValidation(ByVal c As Range) As Boolean
    Dim t As Long
    On Error GoTo No
    t = c.Validation.Type
    HasValidation = True
    Exit Function
No:
    HasValidation = False
End Function

Public Sub BindColumns(ByVal ws As Worksheet)
    Dim lo As ListObject, i As Long, model As String
    Dim groups As clsDict, k As Variant, rng As Range, c As Range

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    ApplyTo ws, lo.ListColumns(ColIdx(lo, "Printer")).DataBodyRange, PrintersFor(ws), "PRN|" & ws.Name, "Printer", "Only active printers available at this print room are listed."
    ApplyTo ws, lo.ListColumns(ColIdx(lo, "Technician")).DataBodyRange, ActiveTechnicians(), "TEC", "Technician", "Inactive technicians are not listed."

    ' Rows are grouped by printer so each distinct list is written once, rather
    ' than once per row.
    Set groups = New clsDict
    For i = 1 To lo.ListRows.Count
        model = Trim$(CStr(CellIn(lo, i, "Printer").Value))
        Set c = CellIn(lo, i, "Paper Stock")
        If groups.Exists(model) Then
            Set rng = groups.Obj(model)
            groups.Add model, Application.Union(rng, c)
        Else
            groups.Add model, c
        End If
    Next i

    For Each k In groups.Keys
        BindStockRange ws, groups.Obj(CStr(k)), CStr(k)
    Next k
End Sub

Public Sub BindStockCell(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal RowNo As Long)
    BindStockRange ws, CellIn(lo, RowNo, "Paper Stock"), Trim$(CStr(CellIn(lo, RowNo, "Printer").Value))
End Sub

' Paper stock cannot be chosen until a printer has been. The cell is locked
' while the printer is blank, which on a protected sheet makes it genuinely
' unavailable rather than merely validated - the incompatible combination
' stops being reachable instead of being caught afterwards.
Private Sub BindStockRange(ByVal ws As Worksheet, ByVal target As Range, ByVal Model As String)
    Dim items As Collection
    If target Is Nothing Then Exit Sub

    If Len(Model) = 0 Then
        Set items = New Collection
        items.Add "- choose a printer first -"
        ApplyTo ws, target, items, "STK|<none>", "Paper stock", "Choose the printer for this job first. The stocks it can take will then be available here."
        UnlockSheet ws
        target.Locked = True
        RelockSheet ws
        Exit Sub
    End If

    Set items = StocksFor(Model)
    If items.Count = 0 Then
        Set items = New Collection
        items.Add "- no active stock fits this printer -"
    End If
    ApplyTo ws, target, items, "STK|" & Model, "Paper stock", "Only stocks whose paper family " & Model & " supports are listed."
    UnlockSheet ws
    target.Locked = False
    RelockSheet ws
End Sub

' Claims a staging column for this tag, writing the list into it, and points
' the target range's validation at it.
Private Sub ApplyTo(ByVal ws As Worksheet, ByVal target As Range, ByVal items As Collection, ByVal Tag As String, ByVal Title As String, ByVal Msg As String)
    Dim st As Worksheet, col As Long, i As Long, last As Long, src As String
    If target Is Nothing Then Exit Sub
    Set st = ThisWorkbook.Worksheets(STAGE_SHEET)

    UnlockSheet st
    col = StageColumn(st, Tag)
    st.Range(st.Cells(FIRST_ITEM_ROW, col), st.Cells(st.Rows.Count, col)).ClearContents
    If items.Count = 0 Then
        st.Cells(FIRST_ITEM_ROW, col).Value = "- nothing available -"
        last = FIRST_ITEM_ROW
    Else
        For i = 1 To items.Count
            If i > MAX_ITEMS Then Exit For
            st.Cells(FIRST_ITEM_ROW + i - 1, col).Value = items(i)
        Next i
        last = FIRST_ITEM_ROW + Application.Min(items.Count, MAX_ITEMS) - 1
    End If
    src = "='" & STAGE_SHEET & "'!" & st.Range(st.Cells(FIRST_ITEM_ROW, col), st.Cells(last, col)).Address(True, True, xlA1)
    RelockSheet st

    UnlockSheet ws
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=src
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = Title
        .InputMessage = Msg
        .ErrorTitle = Title
        .ErrorMessage = "Choose one of the listed options. " & Msg
    End With
    RelockSheet ws
End Sub

Private Function StageColumn(ByVal st As Worksheet, ByVal Tag As String) As Long
    Dim c As Long
    For c = 1 To MAX_COLS
        If Len(Trim$(CStr(st.Cells(1, c).Value))) = 0 Then
            st.Cells(1, c).Value = Tag
            StageColumn = c
            Exit Function
        ElseIf StrComp(CStr(st.Cells(1, c).Value), Tag, vbTextCompare) = 0 Then
            StageColumn = c
            Exit Function
        End If
    Next c
    StageColumn = MAX_COLS
End Function
