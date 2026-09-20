Attribute VB_Name = "modSettings"
Option Explicit

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

' Currently unreferenced. Kept rather than deleted because wiring SET_CURRENCY
' up is phase 8 work (the symbol is hardcoded as ChrW(163) in twelve places
' across modReports and modJobs), and this is where that belongs - deleting it
' now would only mean writing it again. The Val() it used to contain is gone:
' Val on a currency value is the same locale trap as Val on a date, so the
' coercion goes through modUtils.NumOf's rule instead.
Public Function Money(ByVal v As Variant) As String
    Dim d As Double
    If IsNumeric(v) Then
        If Len(Trim$(CStr(v))) > 0 Then d = CDbl(v)
    End If
    Money = Format$(d, CurrencySymbol & "#,##0." & String$(RoundDP, "0"))
End Function

' The symbol the money formats use. Reads SET_CURRENCY so the setting is not
' merely decorative, and falls back to the pound sign the rest of the workbook
' assumes. Phase 8 is where the twelve hardcoded ChrW(163) sites come through
' here.
Public Function CurrencySymbol() As String
    CurrencySymbol = SettingText("CURRENCY", ChrW(163))
End Function
