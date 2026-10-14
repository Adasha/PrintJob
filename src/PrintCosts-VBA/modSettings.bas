Attribute VB_Name = "modSettings"
Option Explicit

' Settings key of the university subsidy (0.10.31), see the subsidy section below.
Public Const SUBSIDY_KEY As String = "SUBSIDY_PCT"

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Typed access to the global settings, so no other module addresses a cell.

Public Function SettingText(ByVal Key As String, Optional ByVal Dflt As String = "") As String
    On Error GoTo Missing
    Dim v As Variant
    v = ThisWorkbook.Names("SET_" & Key).RefersToRange.Value
    If Len(Trim$(CStr(v))) = 0 Then
        SettingText = Dflt
    Else
        SettingText = Trim$(CStr(v))
    End If
    Exit Function
Missing:
    SettingText = Dflt
End Function

Public Function SettingNum(ByVal Key As String, Optional ByVal Dflt As Double = 0) As Double
    Dim s As String
    s = SettingText(Key)
    If Len(s) = 0 Or Not IsNumeric(s) Then SettingNum = Dflt Else SettingNum = CDbl(s)
End Function

Public Sub SetSetting(ByVal Key As String, ByVal Value As Variant)
    On Error Resume Next
    ThisWorkbook.Names("SET_" & Key).RefersToRange.Value = Value
End Sub

Public Function RoundDP() As Long
    RoundDP = CLng(SettingNum("ROUND_DP", 2))
End Function

' The symbol every money NumberFormat and money Format$ call builds from -
' modReports, modExport and modJobs.RemoveRow all read this rather than
' hardcoding a symbol. Reads SET_CURRENCY so the setting is not merely
' decorative, and falls back to the pound sign the rest of the workbook
' assumes.
Public Function CurrencySymbol() As String
    CurrencySymbol = SettingText("CURRENCY", ChrW(163))
End Function

' The NumberFormat code every money cell uses. Deliberately wraps
' CurrencySymbol() in Excel's [$symbol] locale-currency bracket syntax rather
' than just concatenating it - verified directly (COM automation, this
' workbook): a NumberFormat set from VBA to a bare or quoted currency symbol
' followed by "#,##0.00" is silently canonicalised by Excel to the OS's own
' regional currency symbol, no matter what SET_CURRENCY actually holds. On a
' machine whose regional currency happens to be GBP this is invisible - the
' bare-£ code this replaced always "worked" purely because it already matched
' - but it would have silently defeated SET_CURRENCY for any other currency.
' [$symbol] (no "-LCID" needed) round-trips exactly as given and renders
' identically to the bare form for the unaffected case, so this is a strict
' fix, not a behaviour change for the existing £ default. Excel-specific:
' modJobs.RemoveRow's Format$ dialog text is VBA's own formatter, not this
' property, and does not have this problem - it keeps using CurrencySymbol()
' plain.
Public Function CurrencyFormatCode() As String
    CurrencyFormatCode = "[$" & CurrencySymbol() & "]#,##0.00"
End Function

' ------------------------------------------------------------- subsidy ---
' University subsidy (0.10.31): the share of the chargeable cost the university
' pays, one percentage for every print room. Stored on Settings as a fraction
' formatted as a percentage (25% is 0.25), so typing 25 or 25% both mean 25%.
' It only drives the "after subsidy" totals on Reports and in the exported
' report; no job row, cost or snapshot is changed, so there is no schema change.
' The subsidy as a fraction, 0 to 1. A blank, non-numeric or out-of-range
' value reads as no subsidy / the nearest limit rather than an error.
Public Function SubsidyFraction() As Double
    Dim f As Double
    f = SettingNum(SUBSIDY_KEY, 0)
    If f < 0 Then f = 0
    If f > 1 Then f = 1
    SubsidyFraction = f
End Function

' "25" or "12.5": the subsidy as a percentage for wording, never a trailing ".".
Public Function SubsidyPercentText() As String
    SubsidyPercentText = CStr(Round(SubsidyFraction() * 100, 2))
End Function

' Settings rows normally ship in PrintCosts.xlsx; this one is added here, like
' DEPT_ID_HWM, so the template (and any other branch's edits to it) is not
' touched. Idempotent: an existing row and name are left alone, so a value the
' user has set survives every re-run of setup. Called from InitialiseWorkbook.
Public Sub EnsureSubsidySetting()
    Dim lo As ListObject, ws As Worksheet, i As Long, pos As Long, anchor As Long
    Dim k As String, r As ListRow, c As Range

    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    UnlockSheet ws

    For i = 1 To lo.ListRows.Count
        k = Trim$(CStr(CellIn(lo, i, "Key").Value))
        If StrComp(k, SUBSIDY_KEY, vbTextCompare) = 0 Then pos = i
        If StrComp(k, "CURRENCY", vbTextCompare) = 0 Then anchor = i
    Next i

    If pos = 0 Then
        If anchor > 0 Then Set r = lo.ListRows.Add(anchor + 1) Else Set r = lo.ListRows.Add
        pos = r.Index
        CellIn(lo, pos, "Key").Value = SUBSIDY_KEY
        CellIn(lo, pos, "Setting").Value = "University subsidy"
        CellIn(lo, pos, "Value").Value = 0
        CellIn(lo, pos, "Notes").Value = "The percentage of the chargeable cost the university subsidises, one value for every print room (for example 25%). " & _
            "Reports and exported reports show the billable amount after subsidy next to the full chargeable figures. " & _
            "Recorded jobs are not changed. Leave at 0% for no subsidy."
    End If

    Set c = CellIn(lo, pos, "Value")
    c.NumberFormat = "0.##%"
    With c.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "University subsidy"
        .ErrorMessage = "Enter a percentage from 0% to 100%, for example 25%."
        .InputTitle = "University subsidy"
        .InputMessage = "Percentage of the chargeable cost the university pays, e.g. 25%."
    End With

    Set c = Nothing
    On Error Resume Next
    Set c = ThisWorkbook.Names("SET_" & SUBSIDY_KEY).RefersToRange
    On Error GoTo 0
    If c Is Nothing Then
        ThisWorkbook.Names.Add Name:="SET_" & SUBSIDY_KEY, RefersTo:="='" & ws.Name & "'!" & CellIn(lo, pos, "Value").Address
    End If
    RelockSheet ws
End Sub
