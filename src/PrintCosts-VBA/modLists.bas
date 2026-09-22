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
' Tags are: PRN|<all>|<sheet>   every printer permitted at that print room -
'                                used when the row's Paper Stock is blank
'           PRN|<sheet>|<stock> printers at that print room supporting <stock>
'           TEC                 active technicians, the same everywhere
'           STK|<all>|<sheet>   every stock compatible with a printer at that
'                                print room - used when Printer is blank
'           STK|<model>         stocks <model> can take - shared by every row
'                                using it, because the list depends only on
'                                the printer, not the sheet
'
' Printer and Paper Stock filter each other (spec 1a): both cells start
' unlocked and fully populated, and whichever is chosen first narrows the
' other to compatible options, auto-filling it when only one remains.
'
' Staging is needed at all because a validation list supplied as a literal
' string is capped at 255 characters, which a real stock list exceeds.

Private Const STAGE_SHEET As String = "_Work"
Private Const FIRST_ITEM_ROW As Long = 2
Private Const MAX_COLS As Long = 200
Private Const MAX_ITEMS As Long = 500

' Only fills a gap: a Printer or Paper Stock cell that has never been given a
' list - reachable when Excel's own table auto-extend creates a row outside
' AddPrintJob. Checks first and does nothing when a list is already present,
' so there is nothing to flash.
Public Sub OnSelection(ByVal ws As Worksheet, ByVal Target As Range)
    Dim lo As ListObject, n As Long, hdr As String
    If Target.Cells.Count > 1 Then Exit Sub
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Application.Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    hdr = CStr(lo.HeaderRowRange.Cells(1, Target.Column - lo.Range.Column + 1).Value)
    If HasValidation(Target) Then Exit Sub
    n = Target.Row - lo.DataBodyRange.Row + 1

    Select Case hdr
        Case "Paper Stock"
            BindStockCell ws, lo, n
        Case "Printer"
            BindPrinterCell ws, lo, n
    End Select
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
    Dim lo As ListObject, i As Long, model As String, stk As String
    Dim stockGroups As clsDict, printerGroups As clsDict, k As Variant

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    ApplyTo ws, lo.ListColumns(ColIdx(lo, "Technician")).DataBodyRange, ActiveTechnicians(), "TEC", "Technician", "Inactive technicians are not listed."

    ' Printer and Paper Stock filter each other (spec 1a), so each one's list
    ' depends on the OTHER field's value. Rows are grouped by that other
    ' field's value so each distinct list is written once, not once per row -
    ' the same reasoning the single-direction version this replaces used for
    ' Paper Stock alone.
    Set stockGroups = New clsDict
    Set printerGroups = New clsDict
    For i = 1 To lo.ListRows.Count
        model = Trim$(CStr(CellIn(lo, i, "Printer").Value))
        stk = Trim$(CStr(CellIn(lo, i, "Paper Stock").Value))
        GroupCell stockGroups, model, CellIn(lo, i, "Paper Stock")
        GroupCell printerGroups, stk, CellIn(lo, i, "Printer")
    Next i

    For Each k In stockGroups.Keys
        BindStockRange ws, stockGroups.Obj(CStr(k)), CStr(k)
    Next k
    For Each k In printerGroups.Keys
        BindPrinterRange ws, printerGroups.Obj(CStr(k)), CStr(k)
    Next k
End Sub

' Folds a cell into the range already collected under Key, or starts a new
' one - shared by the Printer and Paper Stock grouping passes above.
Private Sub GroupCell(ByVal groups As clsDict, ByVal Key As String, ByVal c As Range)
    Dim rng As Range
    If groups.Exists(Key) Then
        Set rng = groups.Obj(Key)
        groups.Add Key, Application.Union(rng, c)
    Else
        groups.Add Key, c
    End If
End Sub

Public Sub BindStockCell(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal RowNo As Long)
    BindStockRange ws, CellIn(lo, RowNo, "Paper Stock"), Trim$(CStr(CellIn(lo, RowNo, "Printer").Value))
End Sub

Public Sub BindPrinterCell(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal RowNo As Long)
    BindPrinterRange ws, CellIn(lo, RowNo, "Printer"), Trim$(CStr(CellIn(lo, RowNo, "Paper Stock").Value))
End Sub

' Binds the three batch-default cells above the toolbar (spec 1b), using the
' same bidirectional filtering/autofill as the table's own Printer/Paper
' Stock columns (spec 1a) - a single conceptual "row" rather than a whole
' table column, since there's only ever one of each.
Public Sub BindDefaultCells(ByVal ws As Worksheet)
    Dim techCell As Range, prnCell As Range, stkCell As Range
    Set techCell = LocRange(ws, "LOC_DefTech")
    Set prnCell = LocRange(ws, "LOC_DefPrinter")
    Set stkCell = LocRange(ws, "LOC_DefPaper")
    If techCell Is Nothing Or prnCell Is Nothing Or stkCell Is Nothing Then Exit Sub

    ApplyTo ws, techCell, ActiveTechnicians(), "TEC", "Default technician", "Pre-fills each new print job's Technician until Clear defaults is used."
    BindStockRange ws, stkCell, Trim$(CStr(prnCell.Value))
    BindPrinterRange ws, prnCell, Trim$(CStr(stkCell.Value))
End Sub

' Paper Stock choices: every active stock compatible with some printer
' permitted at this location when Printer is blank, or narrowed to what the
' row's own Printer supports once it's chosen (spec 1a). Never locked - the
' cell stays reachable either way, unlike the old one-directional version.
Private Sub BindStockRange(ByVal ws As Worksheet, ByVal target As Range, ByVal Model As String)
    Dim items As Collection, Tag As String, Msg As String
    If target Is Nothing Then Exit Sub

    If Len(Model) = 0 Then
        Set items = StocksForLocation(ws)
        Tag = "STK|<all>|" & ws.Name
        Msg = "Every active stock compatible with a printer at this print room. Choosing one narrows the Printer list to printers that support it."
    Else
        Set items = StocksFor(Model)
        Tag = "STK|" & Model
        Msg = "Only stocks whose paper family " & Model & " supports are listed."
    End If

    AutoFillIfSingle target, items
    If items.Count = 0 Then
        Set items = New Collection
        items.Add "- no active stock fits this printer -"
    End If
    ApplyTo ws, target, items, Tag, "Paper stock", Msg
    UnlockSheet ws
    target.Locked = False
    RelockSheet ws
End Sub

' Printer choices: every active printer permitted at this location when
' Paper Stock is blank, or narrowed to printers that support the row's own
' stock once it's chosen - the mirror of BindStockRange (spec 1a).
Private Sub BindPrinterRange(ByVal ws As Worksheet, ByVal target As Range, ByVal Description As String)
    Dim items As Collection, Tag As String, Msg As String
    If target Is Nothing Then Exit Sub

    If Len(Description) = 0 Then
        Set items = PrintersFor(ws)
        Tag = "PRN|<all>|" & ws.Name
        Msg = "Every active printer permitted at this print room. Choosing a paper stock first narrows this to printers that support it."
    Else
        Set items = PrintersForStock(ws, Description)
        Tag = "PRN|" & ws.Name & "|" & Description
        Msg = "Only printers at this print room that support '" & Description & "' are listed."
    End If

    AutoFillIfSingle target, items
    If items.Count = 0 Then
        Set items = New Collection
        items.Add "- no printer here supports this stock -"
    End If
    ApplyTo ws, target, items, Tag, "Printer", Msg
    UnlockSheet ws
    target.Locked = False
    RelockSheet ws
End Sub

' Every active stock compatible with at least one printer permitted at this
' location - the Paper Stock list's "nothing chosen yet" state.
Private Function StocksForLocation(ByVal ws As Worksheet) As Collection
    Dim out As Collection, seen As clsDict, models As Collection, m As Variant
    Dim s As Collection, i As Long
    Set out = New Collection
    Set seen = New clsDict
    Set models = PrintersFor(ws)
    For Each m In models
        Set s = StocksFor(CStr(m))
        For i = 1 To s.Count
            If Not seen.Exists(CStr(s(i))) Then
                seen.Add CStr(s(i)), True
                out.Add s(i)
            End If
        Next i
    Next m
    Set StocksForLocation = out
End Function

' If exactly one option remains once the other field has narrowed the list,
' fill it in rather than making the user pick the only choice (spec 1a). Only
' fires per-cell when that cell is currently blank, so it never overwrites a
' value the user - or an import - already put there. Checked against the
' real result, before BindStockRange/BindPrinterRange substitute an
' explanatory placeholder for an empty list, so a placeholder message is
' never mistaken for a value.
'
' target is grouped by BindColumns and can span several rows - a single
' contiguous range, several disjoint ones (a multi-Area Union), or one cell.
' Range.Value on anything but a single cell returns an ARRAY, not a scalar,
' so this must walk target.Cells and test/write each one individually rather
' than treating target as one value - CStr() on that array raises a Type
' Mismatch, which is exactly how this was first found (a Union spanning a
' gap in the sheet's row order, from AddPrintJob rows entered out of turn).
Private Sub AutoFillIfSingle(ByVal target As Range, ByVal items As Collection)
    Dim c As Range
    If items.Count <> 1 Then Exit Sub
    For Each c In target.Cells
        If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = items(1)
    Next c
End Sub

' Claims a staging column for this tag, writing the list into it, and points
' the target range's validation at it.
'
' Public so modReports can reuse it for the Reports page's Student name/no/
' Technician/Printer/Paper Stock filters - the same "own staging column per
' tag" reason this exists at all applies there too: a literal Formula1 list
' is capped at 255 characters, which a real name or paper stock list can
' exceed.
'
' Strict (default True, unchanged for every existing caller) rejects a typed
' value that is not on the list. Reports' Student name/no filters pass False:
' Student name matches on a fragment ("Smith" finds "Jane Smith" - see
' Criteria's ISNUMBER(SEARCH(...))) and Student no can legitimately be one
' that has not been logged before, so neither is a genuine "choose one of
' these" field the way Technician/Printer/Paper Stock are - the list is
' offered for browsing and autocomplete only, never enforced.
Public Sub ApplyTo(ByVal ws As Worksheet, ByVal target As Range, ByVal items As Collection, ByVal Tag As String, ByVal Title As String, ByVal Msg As String, Optional ByVal Strict As Boolean = True)
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
        .ShowError = Strict
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
