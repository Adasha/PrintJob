Attribute VB_Name = "modColumns"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' The Show: checkboxes on row 2 of a print room (0.12.0, docs/column-toggles-design.md).
'
' Each box shows or hides one group of columns on its own room. Colour passes and
' Printer default from a rule on what the room has; a click beats the rule until
' the rule's answer changes. The clicks are kept in the room's LOC_Columns cell,
' for example "passes=off@on;cost=off" - only what has been clicked, and for the
' two rule groups what the rule said at the time, so a stale click is noticed
' whichever path changed the printers.
'
' The internal columns are one workbook-wide setting, not a box.

Public Const COL_BOX_TAG As String = "pcb_show_"
Public Const COL_BOX_MACRO As String = "chkShowColumn"
Public Const COL_INTERNAL_KEY As String = "SHOW_INTERNAL_COLS"
Private Const COL_CELL As String = "LOC_Columns"

Public Function ColumnGroups() As Variant
    ColumnGroups = Array("passes", "cost", "disregard", "size", "tech", "printer")
End Function

Public Function ColumnGroupCaption(ByVal Group As String) As String
    Select Case Group
        Case "passes": ColumnGroupCaption = "Colour passes"
        Case "cost": ColumnGroupCaption = "Cost detail"
        Case "disregard": ColumnGroupCaption = "Disregards"
        Case "size": ColumnGroupCaption = "Print size"
        Case "tech": ColumnGroupCaption = "Technician"
        Case "printer": ColumnGroupCaption = "Printer"
    End Select
End Function

' Wide enough for the caption beside the tick box on a Mac (the trial's widths,
' 14 to 16pt narrower, cut off a letter or two).
Public Function ColumnGroupWidth(ByVal Group As String) As Double
    Select Case Group
        Case "passes": ColumnGroupWidth = 108
        Case "cost": ColumnGroupWidth = 92
        Case "disregard": ColumnGroupWidth = 90
        Case "size": ColumnGroupWidth = 84
        Case "tech": ColumnGroupWidth = 94
        Case "printer": ColumnGroupWidth = 74
    End Select
End Function

Private Function GroupHeaders(ByVal Group As String) As Variant
    Select Case Group
        Case "passes": GroupHeaders = Array("Passes", "Colour")
        Case "cost": GroupHeaders = Array("Paper Cost", "Consumable Cost", "Set-up Cost", "Gross Cost", "Disregarded", "Area m2")
        Case "disregard": GroupHeaders = Array("Disregard Paper", "Disregard Consumable")
        Case "size": GroupHeaders = Array("Print Width mm", "Sheet size")
        Case "tech": GroupHeaders = Array("Technician")
        Case "printer": GroupHeaders = Array("Printer")
    End Select
End Function

Private Function InternalHeaders() As Variant
    InternalHeaders = Array("Job ID", "Row Type", "Parent", "Pass", "S_SchemaVer")
End Function

Private Function HasRule(ByVal Group As String) As Boolean
    HasRule = (Group = "passes" Or Group = "printer")
End Function

' What a group does when nobody has clicked it.
Private Function RuleSaysShow(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal Group As String) As Boolean
    Select Case Group
        Case "passes": RuleSaysShow = RoomUsesPasses(ws, lo)
        Case "printer": RuleSaysShow = (PrintersFor(ws).Count <> 1)
        Case Else: RuleSaysShow = True
    End Select
End Function

' A multi-pass printer is permitted here, or the room already holds pass rows
' (the printer may since have been removed from the room).
Private Function RoomUsesPasses(ByVal ws As Worksheet, ByVal lo As ListObject) As Boolean
    Dim k As Variant
    For Each k In PrintersFor(ws)
        If Prn(CStr(k)).MultiPass Then
            RoomUsesPasses = True
            Exit Function
        End If
    Next k
    RoomUsesPasses = (Application.WorksheetFunction.CountIf(lo.ListColumns("Row Type").DataBodyRange, ROW_PASS) > 0)
End Function

Private Function OnOff(ByVal Shown As Boolean) As String
    OnOff = IIf(Shown, "on", "off")
End Function

' Finds Group's stored click. Returns False if there is none.
Private Function StoredChoice(ByVal ws As Worksheet, ByVal Group As String, ByRef Saved As String, ByRef RuleThen As String) As Boolean
    Dim parts As Variant, i As Long, p As Long, entry As String
    parts = Split(LocValue(ws, COL_CELL), ";")
    For i = LBound(parts) To UBound(parts)
        entry = Trim$(CStr(parts(i)))
        If Left$(entry, Len(Group) + 1) = Group & "=" Then
            entry = Mid$(entry, Len(Group) + 2)
            p = InStr(entry, "@")
            If p > 0 Then
                Saved = Left$(entry, p - 1)
                RuleThen = Mid$(entry, p + 1)
            Else
                Saved = entry
                RuleThen = ""
            End If
            StoredChoice = True
            Exit Function
        End If
    Next i
End Function

' Whether a stored click still counts. One made when the rule said something else
' does not: the room has changed since, so the rule takes over again.
Private Function ChoiceIsStale(ByVal RuleThen As String, ByVal RuleNow As Boolean) As Boolean
    ChoiceIsStale = (Len(RuleThen) > 0 And RuleThen <> OnOff(RuleNow))
End Function

Public Function ColumnGroupShown(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal Group As String) As Boolean
    Dim saved As String, ruleThen As String, ruleNow As Boolean
    ruleNow = RuleSaysShow(ws, lo, Group)
    ColumnGroupShown = ruleNow
    If StoredChoice(ws, Group, saved, ruleThen) Then
        If Not ChoiceIsStale(ruleThen, ruleNow) Then ColumnGroupShown = (saved = "on")
    End If
End Function

' Replaces Group's entry in the room's cell. An empty Entry removes it.
Private Sub WriteChoice(ByVal ws As Worksheet, ByVal Group As String, ByVal Entry As String)
    Dim parts As Variant, i As Long, out As String, c As Range, wasOn As Boolean
    Set c = LocRange(ws, COL_CELL)
    If c Is Nothing Then Exit Sub
    parts = Split(CStr(c.Value), ";")
    For i = LBound(parts) To UBound(parts)
        If Left$(Trim$(CStr(parts(i))), Len(Group) + 1) <> Group & "=" Then
            If Len(Trim$(CStr(parts(i)))) > 0 Then out = out & IIf(Len(out) > 0, ";", "") & Trim$(CStr(parts(i)))
        End If
    Next i
    If Len(Entry) > 0 Then out = out & IIf(Len(out) > 0, ";", "") & Entry
    wasOn = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet ws
    c.Value = out
    RelockSheet ws
    Application.EnableEvents = wasOn
End Sub

' Records a choice for one room and applies it. The checkbox macro's target, and
' what the reveal rules below use.
Public Sub SetColumnGroup(ByVal ws As Worksheet, ByVal Group As String, ByVal Shown As Boolean)
    Dim lo As ListObject, entry As String
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    entry = Group & "=" & OnOff(Shown)
    If HasRule(Group) Then entry = entry & "@" & OnOff(RuleSaysShow(ws, lo, Group))
    WriteChoice ws, Group, entry
    ApplyRoomColumns ws
End Sub

' modMain.chkShowColumn's target. Application.Caller names the box that was clicked.
Public Sub ColumnBoxClicked()
    Dim ws As Worksheet, nm As String
    Set ws = ActiveSheet
    nm = CStr(Application.Caller)
    AppOff
    SetColumnGroup ws, Mid$(nm, Len(COL_BOX_TAG) + 1), (ws.CheckBoxes(nm).Value = xlOn)
    AppOn
End Sub

' Sets one room's columns and checkboxes from its stored clicks, its rules and the
' internal setting. Runs on activation as well as after a change, so it only
' unprotects the sheet when a column really has to move. By header, and only the
' columns named here: the H_ and snapshot columns stay as they ship.
Public Sub ApplyRoomColumns(ByVal ws As Worksheet)
    Dim lo As ListObject, g As Variant, shown As Boolean, wantShown As clsDict, h As Variant
    Dim saved As String, ruleThen As String

    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub

    Set wantShown = New clsDict
    For Each g In ColumnGroups()
        shown = ColumnGroupShown(ws, lo, CStr(g))
        For Each h In GroupHeaders(CStr(g))
            wantShown.Add CStr(h), shown
        Next h
        If StoredChoice(ws, CStr(g), saved, ruleThen) Then
            If ChoiceIsStale(ruleThen, RuleSaysShow(ws, lo, CStr(g))) Then WriteChoice ws, CStr(g), ""
        End If
        SetBox ws, CStr(g), shown
    Next g
    shown = InternalColumnsShown()
    For Each h In InternalHeaders()
        wantShown.Add CStr(h), shown
    Next h
    ' Never hideable now, but an older workbook's Reduced view could have hidden it.
    wantShown.Add "Status", True

    SetColumns ws, lo, wantShown
    RepositionLocationButtons ws
End Sub

Private Sub SetColumns(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal WantShown As clsDict)
    Dim k As Variant, col As Long, unlocked As Boolean
    For Each k In WantShown.Keys
        col = 0
        On Error Resume Next
        col = ColIdx(lo, CStr(k))
        On Error GoTo 0
        If col > 0 Then
            If CBool(lo.ListColumns(col).Range.EntireColumn.Hidden) = CBool(WantShown.Item(CStr(k))) Then
                If Not unlocked Then
                    UnlockSheet ws
                    unlocked = True
                End If
                lo.ListColumns(col).Range.EntireColumn.Hidden = Not CBool(WantShown.Item(CStr(k)))
            End If
        End If
    Next k
    If unlocked Then RelockSheet ws
End Sub

' A box that is not there (a sheet not redrawn yet) is skipped.
Private Sub SetBox(ByVal ws As Worksheet, ByVal Group As String, ByVal Shown As Boolean)
    Dim cb As CheckBox, want As Long
    On Error Resume Next
    Set cb = ws.CheckBoxes(COL_BOX_TAG & Group)
    On Error GoTo 0
    If cb Is Nothing Then Exit Sub
    want = IIf(Shown, xlOn, xlOff)
    If cb.Value <> want Then cb.Value = want
End Sub

' ------------------------------------------------- showing a column when needed ---
' A job must not get an error about a column nobody can see. These put the group
' back and tick its box, stored as a click.
Public Sub RevealColumnGroup(ByVal ws As Worksheet, ByVal Group As String)
    Dim lo As ListObject
    Set lo = JobsTable(ws)
    If lo Is Nothing Then Exit Sub
    If Not ColumnGroupShown(ws, lo, Group) Then SetColumnGroup ws, Group, True
End Sub

' After a new job's defaults and one-choice auto-fill have gone in.
Public Sub RevealBlankJobColumns(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal n As Long)
    If Len(TextOf(CellIn(lo, n, "Technician").Value2)) = 0 Then RevealColumnGroup ws, "tech"
    If Len(TextOf(CellIn(lo, n, "Printer").Value2)) = 0 Then RevealColumnGroup ws, "printer"
End Sub

' ------------------------------------------------------------- the setting ---
Public Function InternalColumnsShown() As Boolean
    InternalColumnsShown = (StrComp(SettingText(COL_INTERNAL_KEY, "No"), "Yes", vbTextCompare) = 0)
End Function

' From Workbook_SheetChange on the Settings sheet: every room follows the switch.
Public Sub OnColumnSettingEdited(ByVal Target As Range)
    Dim c As Range, ws As Worksheet
    On Error Resume Next
    Set c = ThisWorkbook.Names("SET_" & COL_INTERNAL_KEY).RefersToRange
    On Error GoTo 0
    If c Is Nothing Then Exit Sub
    If Application.Intersect(Target, c) Is Nothing Then Exit Sub
    AppOff
    For Each ws In ThisWorkbook.Worksheets
        If IsLocation(ws) Then ApplyRoomColumns ws
    Next ws
    AppOn
End Sub

Public Function IsRetiredSetting(ByVal Key As String) As Boolean
    Select Case UCase$(Trim$(Key))
        Case "LOC_REDUCED_VIEW", "LOC_REDUCED_COLUMNS", "LOC_MINIMAL_COLUMNS", "COST_COLS_HIDDEN", "MULTIPASS_COLS_HIDDEN"
            IsRetiredSetting = True
    End Select
End Function

' Added here rather than shipped, as EnsureSubsidySetting does. The settings it
' replaces are removed too, and the new row takes the first one's place. Those
' views are not converted: the old groups and the new ones do not map cleanly.
Public Sub EnsureColumnSettings()
    Dim lo As ListObject, ws As Worksheet, i As Long, pos As Long, gap As Long
    Dim k As String, r As ListRow, c As Range, nm As Variant

    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    UnlockSheet ws

    For i = lo.ListRows.Count To 1 Step -1
        k = Trim$(CStr(CellIn(lo, i, "Key").Value))
        If IsRetiredSetting(k) Then
            lo.ListRows(i).Delete
            gap = i
        End If
    Next i
    For Each nm In Array("LOC_REDUCED_VIEW", "LOC_REDUCED_COLUMNS", "LOC_MINIMAL_COLUMNS", "COST_COLS_HIDDEN", "MULTIPASS_COLS_HIDDEN")
        On Error Resume Next
        ThisWorkbook.Names("SET_" & CStr(nm)).Delete
        On Error GoTo 0
    Next nm

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Key").Value)), COL_INTERNAL_KEY, vbTextCompare) = 0 Then pos = i
    Next i
    If pos = 0 Then
        If gap > 0 And gap <= lo.ListRows.Count Then Set r = lo.ListRows.Add(gap) Else Set r = lo.ListRows.Add
        pos = r.Index
        CellIn(lo, pos, "Key").Value = COL_INTERNAL_KEY
        CellIn(lo, pos, "Setting").Value = "Show internal columns"
        CellIn(lo, pos, "Value").Value = "No"
        CellIn(lo, pos, "Notes").Value = "Yes shows Job ID, Row Type, Parent, Pass and S_SchemaVer on every print room. They are not needed day to day."
    End If

    With CellIn(lo, pos, "Value").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Show internal columns"
        .ErrorMessage = "Choose Yes or No."
    End With

    On Error Resume Next
    ThisWorkbook.Names("SET_" & COL_INTERNAL_KEY).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="SET_" & COL_INTERNAL_KEY, RefersTo:="='" & ws.Name & "'!" & CellIn(lo, pos, "Value").Address
    RelockSheet ws
End Sub
