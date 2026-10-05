Attribute VB_Name = "modVersion"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Version identity, and the About popup on the Settings sheet.
'
' There are TWO version numbers and they mean different things:
'
'   APP_VERSION   this build of the workbook - code, layout, formulas.
'                 Changes every time something is rebuilt.
'
'   SCHEMA_VER    the shape of the stored data, in modUtils. Changes only
'                 when the job table's columns change, because that is what
'                 decides whether one workbook's records can be read by
'                 another (spec 11 collation, and the S_SchemaVer stamped on
'                 every job row).
'
' A build can go 0.4.0 -> 0.5.0 with the schema staying at 1.0, and that is
' the normal case. Conflating the two would force a schema bump on every
' cosmetic change and make the collation check meaningless.
'
' APP_VERSION is 0.<phase>.<revision>, tracking the build sequence in the
' design document: 0.7.x is "phases 1-7 built", and it reaches 1.0.0 when
' acceptance passes at the end of the sequence.

Public Const APP_NAME As String = "Print Cost Management"
' Release history lives in docs/CHANGELOG.md (and git), not in this module.
Public Const APP_VERSION As String = "0.10.25"
Public Const APP_AUTHOR As String = "Adam Shailer"

Public Function VersionString() As String
    VersionString = APP_NAME & " v" & APP_VERSION & " (data schema " & SCHEMA_VER & ")"
End Function

' --------------------------------------------------------------- stamping ---
' Called by build.ps1 AFTER InitialiseWorkbook, not by it.
'
' "Built" has to mean when this file was produced, and InitialiseWorkbook is
' re-run by users every time they add a print room. Stamping from there would
' quietly reset the build date to whenever somebody last duplicated a sheet.
Public Sub StampBuild()
    On Error GoTo Fail
    AppOff
    StampVersionSettings
    SetSetting "BUILT", Now
    SetSetting "BUILT_BY", CurrentUser
    StampProperties
    AppOn
    Say "Stamped as " & VersionString & ".", "Built " & Format$(Now, "dd/mm/yyyy hh:mm") & " by " & CurrentUser & "."
    Exit Sub
Fail:
    AppReset
    ReportError "StampBuild"
End Sub

' --------------------------------------------------- file-level identity ---
' So the version is visible without opening the file: Explorer's details pane,
' Finder's Get Info, and File > Info in Excel all read these.
' Each property is set individually and failures are reported rather than
' swallowed. An earlier version wrapped the lot in On Error Resume Next and
' silently wrote nothing at all - the properties came back empty and the only
' way to find out was to go and look.
Private Sub StampProperties()
    Dim bad As String
    bad = SetProp("Title", APP_NAME & " v" & APP_VERSION)
    bad = bad & SetProp("Subject", "Print output and cost logging, data schema " & SCHEMA_VER)
    bad = bad & SetProp("Author", APP_AUTHOR)
    bad = bad & SetProp("Comments", VersionString & ", built " & Format$(Now, "dd/mm/yyyy hh:mm"))
    If Len(bad) > 0 Then
        Say "Some file properties could not be set.", bad, _
            "The version is still on the Settings sheet; only the Windows file details are affected."
    End If
End Sub

Private Function SetProp(ByVal Nm As String, ByVal Value As String) As String
    On Error GoTo Fail
    ThisWorkbook.BuiltinDocumentProperties(Nm).Value = Value
    Exit Function
Fail:
    SetProp = "- " & Nm & ": " & Err.Number & " " & Err.Description & vbCrLf
End Function

' Stamps the two cells that must always match the code in the workbook: APP_VER
' and SCHEMA. Importing a newer module set makes a newer workbook, and the cells
' should say so without waiting for a rebuild. BUILT and BUILT_BY are left to
' StampBuild. The rows (and their SET_ names) ship in PrintCosts.xlsx; upgrading
' means a fresh workbook plus an import, so nothing here adds or migrates a row.
Public Sub StampVersionSettings()
    SetSetting "APP_VER", APP_VERSION
    SetSetting "SCHEMA", SCHEMA_VER
End Sub

' ------------------------------------------------------------------ about ---
' Bound to the About button on the Settings sheet. The only place any of this
' is shown now that the on-sheet About block (WriteAbout) is gone, so it
' carries the brief description that used to live only in that block's prose
' as well as the identity/version details it always showed.
Public Sub ShowAbout()
    Dim built As String, who As String, ver As String

    ver = SettingText("APP_VER", "(not stamped)")
    built = SettingText("BUILT")
    who = SettingText("BUILT_BY")

    Say APP_NAME & " v" & ver, _
        "Logs print output and cost per student or department across the university's print bureaux. " & _
        "Each record freezes the prices it was costed at, so changing a paper or consumable price " & _
        "never alters what has already been charged." & vbCrLf & vbCrLf & _
        "Data schema " & SCHEMA_VER & IIf(Len(built) > 0, vbCrLf & "Built " & built, "") & _
        IIf(Len(who) > 0, " by " & who, ""), _
        "Specified by " & APP_AUTHOR & ". Built by Claude (Anthropic) working with " & APP_AUTHOR & ", September 2026.", _
        "About"
End Sub
