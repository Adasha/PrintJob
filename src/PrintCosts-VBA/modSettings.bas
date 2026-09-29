Attribute VB_Name = "modSettings"
Option Explicit

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
