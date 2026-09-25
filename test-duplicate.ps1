# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# AT-13: duplicating a print room sheet must survive Refresh Locations.
#
# The duplication is done in VBA rather than over COM, because that is what a
# user actually does (right-click > Move or Copy) and it keeps Excel's own event
# handling inside one process. The test module is injected into the open
# workbook and never saved, so no test code reaches the deliverable.
#
# The workbook is closed WITHOUT saving.

$ErrorActionPreference = 'Stop'
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself.
#
# Neither Close($false) nor AutoSaveOn is the protection it looks like.
# Workbook_Open does real work on every open - ProtectAll unprotects and
# reprotects all fourteen sheets, plus Invalidate and HealButtons - so the
# workbook is dirty the instant it opens. Excel holds this file through its
# cloud-backed handle (Workbook.Path returns a d.docs.live.net URL), so
# AutoSave is on and COMMITS that immediately: $wb.Saved reads True on the
# very next line after Open, despite every sheet having just been modified.
# The bytes land when the handle closes - measured, the timestamp and the
# SHA-256 both move at Close/Quit, not at Open.
#
# Which is exactly why Close($false) does not help. It means "discard unsaved
# changes", and by then there are none: AutoSave has accepted them, so what
# is written at close is committed state. Setting AutoSaveOn = $false after
# Open is too late for the same reason, and Open offers no earlier hook - so
# no guard is attempted here, because none can work. (build.ps1 sets it, but
# only ever on a %TEMP% copy, which is not cloud-backed; its own comment notes
# it is usually a no-op there.) Tested: AutoSaveOn False, Saved forced True,
# nothing touched at all - the file still changes.
#
# None of this is the sync client. It reproduces with syncing paused and all
# sync operations settled. It is Excel-side only.
#
# verify.ps1 and probe.ps1 take the other way out: read-only, so Workbook_Open
# cannot save anything. These scripts need read-write to drive mutating
# macros, so a throwaway copy is what is left. Discarded at the end.
#
# Left unfixed, this rewrites the artefact under test: the phase 5-7 runs did,
# and a later script read an earlier one's edits back as though they were the
# build.
$deliverable = Join-Path $PSScriptRoot 'src\PrintCosts.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintCosts.xlsm'
Copy-Item $deliverable $f

$vba = @'
' Not named AT13: VBA and Application.Run both read a name of that shape as a
' cell reference (column AT, row 13) and refuse to resolve it as a procedure.
Public Function TestDuplicate() As String
    Dim r As String, ws As Worksheet, src As Worksheet
    On Error GoTo Fail
    Application.DisplayAlerts = False

    Set src = ThisWorkbook.Worksheets("Annexe")
    r = "job tables before:" & vbCrLf & TableList
    src.Copy After:=src
    Set ws = ActiveSheet
    r = r & "duplicated 'Annexe' -> '" & ws.Name & "'" & vbCrLf
    r = r & "job tables straight after the copy:" & vbCrLf & TableList

    SetQuiet True
    RefreshLocations
    r = r & "Refresh Locations said:" & vbCrLf & QuietLog
    SetQuiet False

    r = r & "job tables after refresh:" & vbCrLf & TableList
    r = r & "registry:" & vbCrLf & RegList
    r = r & "consolidated: " & SpillSize & vbCrLf
    TestDuplicate = r
    Exit Function
Fail:
    TestDuplicate = r & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
End Function

Private Function TableList() As String
    Dim ws As Worksheet, lo As ListObject, r As String
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If Left$(lo.Name, 8) = "tblJobs_" Then
                r = r & "   " & lo.Name & "  on '" & ws.Name & "'  LOC_Code=" & ws.Range("LOC_Code").Value & vbCrLf
            End If
        Next lo
    Next ws
    TableList = r
End Function

Private Function RegList() As String
    Dim lo As ListObject, i As Long, c As Long, r As String, ln As String
    Set lo = Tbl("tblLocations")
    If lo Is Nothing Then
        RegList = "   (no tblLocations)" & vbCrLf
        Exit Function
    End If
    For i = 1 To lo.ListRows.Count
        ln = ""
        For c = 1 To lo.ListColumns.Count
            ln = ln & CStr(lo.ListRows(i).Range.Cells(1, c).Text) & " | "
        Next c
        r = r & "   " & ln & vbCrLf
    Next i
    RegList = r
End Function

Private Function SpillSize() As String
    Dim rng As Range
    On Error GoTo No
    Set rng = ThisWorkbook.Worksheets("_Data").Range("A10").SpillingToRange
    SpillSize = rng.Rows.Count & " rows x " & rng.Columns.Count & " cols"
    Exit Function
No:
    SpillSize = "NO SPILL"
End Function
'@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($f)
try {
    $m = $wb.VBProject.VBComponents.Add(1)
    $m.Name = 'modDupTest'
    $m.CodeModule.AddFromString($vba)
    Write-Host $xl.Run('TestDuplicate')
}
finally {
    $wb.Close($false)
    $xl.Quit()
}
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
