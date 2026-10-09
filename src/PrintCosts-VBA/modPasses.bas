Attribute VB_Name = "modPasses"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Multi-pass (per-colour) costing, 0.11.0 - docs/ARCHITECTURE.md section 17 and
' docs/multipass-costing-design.md (decision numbers below refer to it).
'
' A multi-pass job is one Job row with a Pass row per colour directly beneath
' it, in a collapsible row group (parent on top - the template ships
' summaryBelow="0"). A pass row carries no Job ID of its own: its Parent column
' holds the job's ID and its Pass column a number that is kept 1, 2, 3 ...
' with no gaps (decisions 1, 15, 24).
'
' Row groups belong to the sheet ROW, not the data, so every insert and delete
' here is a whole-row one (Range.EntireRow.Insert / .Delete), which keeps the
' levels lined up with the data; a cell-shifting ListRows.Delete would leave
' them behind. RebuildGroups repairs any misalignment (a paste, a manual
' delete) and runs from Check sheet / Check workbook and Refresh Locations.
'
' Money: a pass row's Consumable Cost and Set-up Cost come from its own
' snapshot (S_ConsRate, S_SetupCost) via the hidden H_Ink / H_Setup columns, and
' the Job row sums them (decision 16) - see the formulas in PrintCosts.xlsx
' and docs/ARCHITECTURE.md section 5.1. Nothing here recalculates a cost.

Public Const ROW_JOB As String = "Job"
Public Const ROW_PASS As String = "Pass"

Public Const MSG_SELECT_MULTIPASS As String = "Select a multi-pass job, or one of its passes, first."

' The text that marks ApplyPassFormat's own conditional format (the value it tests
' for), so a later run can find and replace it without touching any other rule.
Private Const PASS_CF_MARK As String = "Pass"

' Caption colour of a greyed pass button (RGB 150,150,150).
Private Const PASS_BTN_GREY As Long = 9868950

' --------------------------------------------------------------- lookups ---
' One column as a 1-based (rows x 1) array, whatever the row count - Range.Value2
' on a single cell is a scalar, not an array.
Public Function ColArr(ByVal lo As ListObject, ByVal Header As String) As Variant
    Dim rng As Range, a() As Variant
    Set rng = lo.ListColumns(ColIdx(lo, Header)).DataBodyRange
    If rng Is Nothing Then
        ReDim a(1 To 1, 1 To 1)
        ColArr = a
    ElseIf rng.Rows.Count = 1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = rng.Cells(1, 1).Value2
        ColArr = a
    Else
        ColArr = rng.Value2
    End If
End Function

Public Function TextOf(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    TextOf = Trim$(CStr(v))
End Function

' Whether table row n is a Pass row. A blank Row Type (a row that predates the
' column, or one just added) is a Job row.
Public Function IsPassRow(ByVal lo As ListObject, ByVal n As Long) As Boolean
    If n < 1 Or n > lo.ListRows.Count Then Exit Function
    IsPassRow = (StrComp(TextOf(CellIn(lo, n, "Row Type").Value2), ROW_PASS, vbTextCompare) = 0)
End Function

' Job rows that hold a real job: not blank, not a pass. What "N print jobs"
' means anywhere it is shown.
Public Function JobRowCount(ByVal lo As ListObject) As Long
    Dim rt As Variant, jid As Variant, i As Long, n As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    If RowCount(lo) = 0 Then Exit Function
    rt = ColArr(lo, "Row Type")
    For i = 1 To UBound(rt, 1)
        If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) <> 0 Then n = n + 1
    Next i
    JobRowCount = n
End Function

' The Job row (table row index) a row belongs to: itself for a Job row, its
' parent for a Pass row, 0 for an orphan pass or no row.
Public Function JobRowFor(ByVal lo As ListObject, ByVal n As Long) As Long
    Dim rt As Variant, pid As Variant, jid As Variant, i As Long, want As String
    If n < 1 Or n > lo.ListRows.Count Then Exit Function
    rt = ColArr(lo, "Row Type")
    If StrComp(TextOf(rt(n, 1)), ROW_PASS, vbTextCompare) <> 0 Then
        JobRowFor = n
        Exit Function
    End If
    pid = ColArr(lo, "Parent")
    jid = ColArr(lo, "Job ID")
    want = TextOf(pid(n, 1))
    If Len(want) = 0 Then Exit Function
    For i = n - 1 To 1 Step -1       ' the parent sits above its passes
        If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) <> 0 Then
            If StrComp(TextOf(jid(i, 1)), want, vbTextCompare) = 0 Then JobRowFor = i
            Exit Function
        End If
    Next i
End Function

' The pass rows of a job, in order, as table row indexes. Empty when none.
Public Function PassRowsOf(ByVal lo As ListObject, ByVal JobRow As Long) As Collection
    Dim out As New Collection, rt As Variant, pid As Variant, jid As Variant, i As Long, want As String
    rt = ColArr(lo, "Row Type")
    pid = ColArr(lo, "Parent")
    jid = ColArr(lo, "Job ID")
    want = TextOf(jid(JobRow, 1))
    If Len(want) > 0 Then
        For i = JobRow + 1 To UBound(rt, 1)
            If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) <> 0 Then Exit For
            If StrComp(TextOf(pid(i, 1)), want, vbTextCompare) = 0 Then out.Add i Else Exit For
        Next i
    End If
    Set PassRowsOf = out
End Function

' The selected table row, or 0 - with no message, unlike modJobs.SelectedRow:
' the pass buttons have their own wording.
Public Function SelectedTableRow(ByVal ws As Worksheet, ByVal lo As ListObject) As Long
    Dim c As Range
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    On Error Resume Next
    Set c = Application.Intersect(Selection.Cells(1, 1).EntireRow, lo.DataBodyRange)
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    SelectedTableRow = c.Row - lo.DataBodyRange.Row + 1
End Function

' The multi-pass job the selection is on (the job itself, or the job a selected
' pass belongs to), as a table row; 0 when it is not on one.
Private Function SelectedMultiPassJob(ByVal ws As Worksheet, ByVal lo As ListObject) As Long
    Dim sel As Long, jr As Long, p As clsPrinterDef
    sel = SelectedTableRow(ws, lo)
    If sel = 0 Then Exit Function
    jr = JobRowFor(lo, sel)
    If jr = 0 Then Exit Function
    If Len(TextOf(CellIn(lo, jr, "Job ID").Value2)) = 0 Then Exit Function
    Set p = Prn(TextOf(CellIn(lo, jr, "Printer").Value2))
    If Not p.Found Then Exit Function
    If Not p.MultiPass Then Exit Function
    SelectedMultiPassJob = jr
End Function

Public Function MultiPassPrinterExists() As Boolean
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblPrinters")
    If lo Is Nothing Then Exit Function
    If Not ColumnExists(lo, "Colour mode") Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(TextOf(CellIn(lo, i, "Colour mode").Value2), COLOUR_MODE_MULTI, vbTextCompare) = 0 Then
            MultiPassPrinterExists = True
            Exit Function
        End If
    Next i
End Function

' ------------------------------------------------------------- add / remove ---
' Add pass (decision 7): a new pass row at the foot of the selected multi-pass
' job's group, numbered next, with the printer's Template cost stamped on it
' (decision 23). The colour is chosen afterwards from the dropdown, which
' stamps the colour's rate.
Public Sub AddPass(ByVal ws As Worksheet)
    Dim lo As ListObject, jr As Long, passes As Collection, last As Long, n As Long
    Dim jobId As String, p As clsPrinterDef, r As ListRow, k As Variant

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    jr = SelectedMultiPassJob(ws, lo)
    If jr = 0 Then
        Say MSG_SELECT_MULTIPASS
        Exit Sub
    End If

    jobId = TextOf(CellIn(lo, jr, "Job ID").Value2)
    Set p = Prn(TextOf(CellIn(lo, jr, "Printer").Value2))
    Set passes = PassRowsOf(lo, jr)
    last = jr
    If passes.Count > 0 Then last = CLng(passes(passes.Count))

    AppOff
    UnlockSheet ws
    ' Expand the group first so the new row is not born inside a collapsed one.
    For Each k In passes
        lo.ListRows(CLng(k)).Range.EntireRow.Hidden = False
    Next k
    If last < lo.ListRows.Count Then
        lo.ListRows(last + 1).Range.EntireRow.Insert
        n = last + 1
    Else
        Set r = lo.ListRows.Add
        n = r.Index
    End If

    ' A whole-row insert inherits the row above's level and any constants are
    ' not carried, but be explicit about everything this row is.
    ClearTypedCells lo.ListRows(n).Range
    CellIn(lo, n, "Row Type").Value = ROW_PASS
    CellIn(lo, n, "Parent").Value = jobId
    CellIn(lo, n, "Pass").Value = passes.Count + 1
    CellIn(lo, n, "S_SetupCost").Value = IIf(p.TemplateCost <> 0, p.TemplateCost, Empty)
    CellIn(lo, n, "S_StampedAt").Value = Now
    CellIn(lo, n, "S_StampedBy").Value = CurrentUser
    CellIn(lo, n, "S_SchemaVer").Value = SCHEMA_VER
    lo.ListRows(n).Range.EntireRow.OutlineLevel = 2
    lo.ListRows(n).Range.EntireRow.Hidden = False
    RelockSheet ws
    BindPassColours ws, lo
    AppOn

    CellIn(lo, n, "Colour").Select
End Sub

' Remove pass (decision 7): the selected pass, or - with its job selected - the
' job's last pass. Confirms with what is about to go, then renumbers the rest
' (decision 24).
Public Sub RemovePass(ByVal ws As Worksheet)
    Dim lo As ListObject, jr As Long, passes As Collection, target As Long, sel As Long
    Dim detail As String, jobId As String

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    jr = SelectedMultiPassJob(ws, lo)
    If jr = 0 Then
        Say MSG_SELECT_MULTIPASS
        Exit Sub
    End If
    Set passes = PassRowsOf(lo, jr)
    If passes.Count = 0 Then
        Say "This job has no passes to remove.", "Add pass adds one."
        Exit Sub
    End If

    sel = SelectedTableRow(ws, lo)
    If IsPassRow(lo, sel) Then target = sel Else target = CLng(passes(passes.Count))
    jobId = TextOf(CellIn(lo, jr, "Job ID").Value2)

    detail = "Pass " & TextOf(CellIn(lo, target, "Pass").Value2) & " of " & jobId & ": " & _
        NzText(CellIn(lo, target, "Colour").Value, "(no colour chosen)")
    If Not Ask("Remove this pass?" & vbCrLf & vbCrLf & detail & vbCrLf & vbCrLf & "This cannot be undone.", "Remove pass") Then Exit Sub

    DeletePass ws, lo, jr, target, detail
End Sub

' The unprompted delete (the routine a test or a future caller reaches directly,
' as DeleteVisibleReportsConfirmed is for Reports): audit note, whole-row delete,
' renumber.
Public Sub DeletePass(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal JobRow As Long, _
                      ByVal PassRow As Long, ByVal Detail As String)
    AppOff
    LogAudit "Remove pass", LocValue(ws, "LOC_Name"), Detail
    UnlockSheet ws
    lo.ListRows(PassRow).Range.EntireRow.Delete
    RenumberPasses lo, JobRow
    RelockSheet ws
    AppOn
End Sub

' Passes 1, 2, 3 ... in row order, no gaps. Caller has the sheet unlocked.
Public Sub RenumberPasses(ByVal lo As ListObject, ByVal JobRow As Long)
    Dim passes As Collection, k As Variant, i As Long
    Set passes = PassRowsOf(lo, JobRow)
    For Each k In passes
        i = i + 1
        If NumOf(CellIn(lo, CLng(k), "Pass")) <> i Then CellIn(lo, CLng(k), "Pass").Value = i
    Next k
End Sub

' The rows that go when a Job row is deleted: itself and its passes, as table
' row indexes in ascending order.
Public Function GroupRows(ByVal lo As ListObject, ByVal JobRow As Long) As Collection
    Dim out As New Collection, k As Variant
    out.Add JobRow
    For Each k In PassRowsOf(lo, JobRow)
        out.Add CLng(k)
    Next k
    Set GroupRows = out
End Function

' ---------------------------------------------------------------- colours ---
' A colour was picked or typed on a pass row: stamp its rate and ID. A colour
' this workbook does not define stamps nothing (an imported row keeps what it
' came with); the milder Status notice covers it. Events are off.
Public Sub OnColourChanged(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim nm As String, c As clsColour
    If Not IsPassRow(lo, n) Then Exit Sub
    nm = TextOf(CellIn(lo, n, "Colour").Value2)
    If Len(nm) = 0 Then
        CellIn(lo, n, "S_ColourID").ClearContents
        CellIn(lo, n, "S_ConsRate").ClearContents
        Exit Sub
    End If
    Set c = Colour(nm)
    If c.Found Then
        CellIn(lo, n, "S_ColourID").Value = c.ColourID
        CellIn(lo, n, "S_ConsRate").Value = c.RatePerM2
        CellIn(lo, n, "S_StampedAt").Value = Now
        CellIn(lo, n, "S_StampedBy").Value = CurrentUser
    End If
End Sub

' Re-stamp one pass row at today's prices (Re-stamp prices, decisions 25 and
' 31): set-up cost from its job's printer, ink rate from its colour - but a
' colour this workbook cannot find leaves the stamped rate and ID alone.
Public Sub RestampPassRow(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    Dim jr As Long, p As clsPrinterDef, c As clsColour, nm As String
    jr = JobRowFor(lo, n)
    If jr > 0 Then
        Set p = Prn(TextOf(CellIn(lo, jr, "Printer").Value2))
        If p.Found Then CellIn(lo, n, "S_SetupCost").Value = IIf(p.TemplateCost <> 0, p.TemplateCost, Empty)
    End If
    nm = TextOf(CellIn(lo, n, "Colour").Value2)
    If Len(nm) > 0 Then
        Set c = Colour(nm)
        If c.Found Then
            CellIn(lo, n, "S_ColourID").Value = c.ColourID
            CellIn(lo, n, "S_ConsRate").Value = c.RatePerM2
        End If
    End If
    CellIn(lo, n, "S_StampedAt").Value = Now
    CellIn(lo, n, "S_StampedBy").Value = CurrentUser
    CellIn(lo, n, "S_SchemaVer").Value = SCHEMA_VER
End Sub

' Binds every pass row's Colour dropdown to the colours of its job's printer
' type (decision 3), one staged list per consumable type. Called from
' modLists.BindColumns, so it runs on setup, Refresh Locations and Check.
Public Sub BindPassColours(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim rt As Variant, pid As Variant, jid As Variant, prnCol As Variant, i As Long, j As Long
    Dim groups As clsDict, k As Variant, ctype As String, model As String, c As Range

    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Colour") Then Exit Sub
    rt = ColArr(lo, "Row Type")
    pid = ColArr(lo, "Parent")
    jid = ColArr(lo, "Job ID")
    prnCol = ColArr(lo, "Printer")

    Set groups = New clsDict
    For i = 1 To UBound(rt, 1)
        If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) = 0 Then
            ctype = ""
            model = ""
            For j = i - 1 To 1 Step -1
                If StrComp(TextOf(rt(j, 1)), ROW_PASS, vbTextCompare) <> 0 Then
                    If StrComp(TextOf(jid(j, 1)), TextOf(pid(i, 1)), vbTextCompare) = 0 Then model = TextOf(prnCol(j, 1))
                    Exit For
                End If
            Next j
            If Len(model) > 0 Then ctype = Prn(model).Consumable
            Set c = CellIn(lo, i, "Colour")
            If groups.Exists(ctype) Then
                groups.Add ctype, Application.Union(groups.Obj(ctype), c)
            Else
                groups.Add ctype, c
            End If
        End If
    Next i

    For Each k In groups.Keys
        BindColourRange ws, groups.Obj(CStr(k)), CStr(k)
    Next k
End Sub

Private Sub BindColourRange(ByVal ws As Worksheet, ByVal target As Range, ByVal ConsumableType As String)
    Dim items As Collection, Tag As String, Msg As String
    Set items = ColoursFor(ConsumableType)
    Tag = "CLR|" & ConsumableType
    If Len(ConsumableType) = 0 Then
        Msg = "Choose the colour of this pass."
    Else
        Msg = "The " & ConsumableType & " colours set up on the Consumables sheet."
    End If
    ApplyTo ws, target, items, Tag, "Colour", Msg
    UnlockSheet ws
    target.Locked = False
    RelockSheet ws
End Sub

' ------------------------------------------------------------ row groups ---
' Puts every row's outline level back where its Row Type says it belongs (Job 1,
' Pass 2) and keeps each group's collapsed state. Only rows that are wrong are
' touched. Caller need not have the sheet unlocked.
Public Sub RebuildGroups(ByVal ws As Worksheet)
    Dim lo As ListObject, rt As Variant, i As Long, want As Long, hid As Boolean
    Dim rowRng As Range, collapsed As Boolean, changed As Boolean

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Row Type") Then Exit Sub
    rt = ColArr(lo, "Row Type")

    UnlockSheet ws
    For i = 1 To UBound(rt, 1)
        Set rowRng = lo.ListRows(i).Range.EntireRow
        If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) = 0 Then
            want = 2
            ' A pass takes the collapsed state of the pass before it, else its own.
            If i > 1 Then
                If StrComp(TextOf(rt(i - 1, 1)), ROW_PASS, vbTextCompare) = 0 Then
                    hid = lo.ListRows(i - 1).Range.EntireRow.Hidden
                Else
                    hid = rowRng.Hidden
                End If
            Else
                hid = rowRng.Hidden
            End If
        Else
            want = 1
            hid = False
        End If
        If rowRng.OutlineLevel <> want Then rowRng.OutlineLevel = want
        If CBool(rowRng.Hidden) <> hid Then rowRng.Hidden = hid
    Next i
    RelockSheet ws
End Sub

' A light shade over every pass row so a group reads as one job and its colours.
' A conditional format rather than direct formatting, so a row added below a pass
' (which copies its look) never inherits the shade. Written with INDEX(..., ROW())
' on the whole Row Type column, so it does not depend on the active cell. Replaces
' only its own earlier rule (recognised by its formula), leaving any other
' conditional format on the table alone.
Public Sub ApplyPassFormat(ByVal ws As Worksheet)
    Dim lo As ListObject, body As Range, fc As FormatCondition, i As Long, rtCol As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Row Type") Then Exit Sub
    Set body = lo.DataBodyRange

    UnlockSheet ws
    For i = body.FormatConditions.Count To 1 Step -1
        On Error Resume Next
        If InStr(1, body.FormatConditions(i).Formula1, PASS_CF_MARK, vbTextCompare) > 0 Then body.FormatConditions(i).Delete
        On Error GoTo 0
    Next i
    rtCol = lo.ListColumns("Row Type").Range.Cells(1, 1).EntireColumn.Address(True, True)
    Set fc = body.FormatConditions.Add(Type:=xlExpression, Formula1:="=INDEX(" & rtCol & ",ROW())=""" & PASS_CF_MARK & """")
    fc.Interior.Color = RGB(236, 242, 250)
    RelockSheet ws
End Sub

' Per-sheet preparation, run by setup and Refresh Locations: row groups back in
' line, and no sort or filter on a location sheet (decision 12 - rows stay in
' the order added, and a sort would scatter the passes from their jobs; the
' Reports sheet is the place for both). The AutoFilter arrows are hidden and
' ProtectSheet withholds the sort/filter permissions on these sheets.
Public Sub PreparePassSheet(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    UnlockSheet ws
    On Error Resume Next
    If lo.ShowAutoFilter Then lo.ShowAutoFilter = False
    On Error GoTo 0
    RelockSheet ws
    RebuildGroups ws
End Sub

' ----------------------------------------------------- column toggle (view) ---
' The multi-pass columns toggle on top of the All / Reduced / Minimal drop-down
' (column-view-presets-design.md, decisions 1, 8): a column is hidden if either
' the view or the toggle hides it, and "All" shows everything except a group
' toggled off. Set-up Cost is NOT in this group - it follows the cost toggle.
Public Function MultiPassColumns() As Variant
    MultiPassColumns = Array("Passes", "Colour", "Row Type", "Parent", "Pass")
End Function

Public Function MultiPassColsHidden() As Boolean
    MultiPassColsHidden = (StrComp(SettingText("MULTIPASS_COLS_HIDDEN", "No"), "Yes", vbTextCompare) = 0)
End Function

Public Function MultiPassColsCaption() As String
    MultiPassColsCaption = IIf(MultiPassColsHidden(), "Show pass columns", "Hide pass columns")
End Function

' Re-asserts a hidden toggle after the view has been applied (ApplyReducedView
' calls this last). Only ever hides: showing is the view's decision.
Public Sub ApplyMultiPassVisibility(ByVal ws As Worksheet)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If Not MultiPassColsHidden() Then Exit Sub
    ApplyColumnVisibility lo, MultiPassColumns(), True
End Sub

' The toggle button: flips the setting and re-applies the view on every
' location sheet, which re-hides (or, via the view, re-shows) the columns.
Public Sub ToggleMultiPassColumns()
    Dim ws As Worksheet
    SetSetting "MULTIPASS_COLS_HIDDEN", IIf(MultiPassColsHidden(), "No", "Yes")
    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then ApplyReducedView ws
    Next ws
    AppOn
End Sub

' ------------------------------------------------------------ collapse / expand ---
' Toggle passes (decision 7): hide or show the selected job's pass rows - a job
' whose passes are all visible collapses, otherwise it expands. Works with the
' job or one of its passes selected. Hiding the rows is all the +/- control on
' the job row needs to flip with it.
Public Sub TogglePasses(ByVal ws As Worksheet)
    Dim lo As ListObject, jr As Long, passes As Collection, k As Variant, anyVisible As Boolean
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    jr = SelectedMultiPassJob(ws, lo)
    If jr = 0 Then
        Say MSG_SELECT_MULTIPASS
        Exit Sub
    End If
    Set passes = PassRowsOf(lo, jr)
    If passes.Count = 0 Then Exit Sub
    For Each k In passes
        If Not lo.ListRows(CLng(k)).Range.EntireRow.Hidden Then anyVisible = True
    Next k
    UnlockSheet ws
    For Each k In passes
        lo.ListRows(CLng(k)).Range.EntireRow.Hidden = anyVisible
    Next k
    RelockSheet ws
End Sub

' Toggle all passes (decision 7): groups that are a mix of expanded and
' collapsed are first all expanded so they match; all expanded collapses all;
' all collapsed expands all.
Public Sub ToggleAllPasses(ByVal ws As Worksheet)
    Dim lo As ListObject, rt As Variant, i As Long, vis As Long, hid As Long, hideThem As Boolean
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    rt = ColArr(lo, "Row Type")
    For i = 1 To UBound(rt, 1)
        If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) = 0 Then
            If lo.ListRows(i).Range.EntireRow.Hidden Then hid = hid + 1 Else vis = vis + 1
        End If
    Next i
    If vis + hid = 0 Then Exit Sub
    hideThem = (hid = 0)              ' all expanded -> collapse; a mix or all collapsed -> expand
    UnlockSheet ws
    For i = 1 To UBound(rt, 1)
        If StrComp(TextOf(rt(i, 1)), ROW_PASS, vbTextCompare) = 0 Then lo.ListRows(i).Range.EntireRow.Hidden = hideThem
    Next i
    RelockSheet ws
End Sub

' ------------------------------------------------------------ button state ---
' Add pass / Remove pass are clickable only with a multi-pass job (or one of its
' passes) selected; all four pass buttons are shaded lighter when no multi-pass
' printer is set up at all (column-view-presets-design.md, decisions 5-7). The
' grey is the caption's font colour - Button.Enabled does not visibly grey - and
' a click on a greyed button still shows the short message (AddPass /
' RemovePass guard themselves). Changing the font colour needs the sheet
' unprotected, so it is done only when the state actually changes.
Public Sub RefreshPassButtons(ByVal ws As Worksheet)
    Dim lo As ListObject, noMulti As Boolean, onJob As Boolean, i As Long, b As Button
    Dim greyIt As Boolean, isGrey As Boolean, nm As String, changed As Boolean

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    noMulti = Not MultiPassPrinterExists()
    If Not noMulti Then onJob = (SelectedMultiPassJob(ws, lo) > 0)

    For i = 1 To ws.Buttons.Count
        Set b = ws.Buttons(i)
        nm = b.Name
        If Left$(nm, 4) = "pcb_" Then
            If InStr(nm, "btnAddPass") = 5 Or InStr(nm, "btnRemovePass") = 5 Then
                greyIt = noMulti Or Not onJob
            ElseIf InStr(nm, "btnTogglePasses") = 5 Or InStr(nm, "btnToggleAllPasses") = 5 Or InStr(nm, "btnToggleMultiPass") = 5 Then
                greyIt = noMulti
            Else
                GoTo NextButton
            End If
            isGrey = (b.Characters.Font.Color = PASS_BTN_GREY)
            If isGrey <> greyIt Then
                If Not changed Then UnlockSheet ws: changed = True
                b.Characters.Font.Color = IIf(greyIt, PASS_BTN_GREY, 0)
            End If
        End If
NextButton:
    Next i
    If changed Then RelockSheet ws
End Sub
