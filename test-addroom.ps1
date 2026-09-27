# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Add print room (modRegistry.AddPrintRoom/CreatePrintRoom): duplicating a
# location sheet through the new automated command must produce a genuine
# fresh start - no job rows, batch defaults, disregard defaults, roll-unit
# setting or permitted-printers list carried over from the template sheet -
# and must still survive Refresh Locations (AT-13) exactly as a manual
# duplicate does, since CreatePrintRoom ends by calling it.
#
# CreatePrintRoom is called directly, not AddPrintRoom, so there is no
# InputBox to answer - same reasoning test-duplicate.ps1 gives for driving
# the duplication in VBA rather than over COM: it is what a user actually
# does, and it keeps Excel's own event handling inside one process.
#
# The workbook is closed WITHOUT saving.

$ErrorActionPreference = 'Stop'
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$vba = @'
Public Function TestAddRoom() As String
    Dim r As String, ws As Worksheet, src As Worksheet
    On Error GoTo Fail
    Application.DisplayAlerts = False

    Set src = ThisWorkbook.Worksheets("Example Print Room")
    ' Give the template some non-default state first, so clearing it is
    ' actually exercised rather than trivially true - a template that
    ' already has empty defaults would pass even a no-op reset.
    UnlockSheet src
    LocRange(src, "LOC_DefDisPaper").Value = "Yes"
    LocRange(src, "LOC_RollUnit").Value = "Centimetres"
    RelockSheet src

    r = "template before: Rows=" & JobsTable(src).ListRows.Count & _
        " DefDisPaper=" & LocValue(src, "LOC_DefDisPaper") & _
        " RollUnit=" & LocValue(src, "LOC_RollUnit") & _
        " Printers=" & LocValue(src, "LOC_Printers") & vbCrLf

    SetQuiet True
    Set ws = CreatePrintRoom("Test Room #1 / Annex:2", "Facilities")
    r = r & "Refresh Locations said:" & vbCrLf & QuietLog
    SetQuiet False

    If ws Is Nothing Then
        TestAddRoom = r & "FAIL: CreatePrintRoom returned Nothing" & vbCrLf
        Exit Function
    End If

    r = r & "new sheet name: '" & ws.Name & "'" & vbCrLf
    r = r & "new sheet: Rows=" & JobsTable(ws).ListRows.Count & _
        " LOC_Name=" & LocValue(ws, "LOC_Name") & _
        " LOC_Dept=" & LocValue(ws, "LOC_Dept") & _
        " LOC_Code=" & LocValue(ws, "LOC_Code") & _
        " DefTech=[" & LocValue(ws, "LOC_DefTech") & "]" & _
        " DefDisPaper=[" & LocValue(ws, "LOC_DefDisPaper") & "]" & _
        " RollUnit=" & LocValue(ws, "LOC_RollUnit") & _
        " Printers=[" & LocValue(ws, "LOC_Printers") & "]" & vbCrLf

    r = r & "template unchanged: Rows=" & JobsTable(src).ListRows.Count & _
        " LOC_Code=" & LocValue(src, "LOC_Code") & vbCrLf

    r = r & "registry:" & vbCrLf & RegList
    r = r & "consolidated: " & SpillSize & vbCrLf

    TestAddRoom = r
    Exit Function
Fail:
    TestAddRoom = r & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
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
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $m = $wb.VBProject.VBComponents.Add(1)
    $m.Name = 'modAddRoomTest'
    $m.CodeModule.AddFromString($vba)
    Write-Host $xl.Run('TestAddRoom')
}
finally {
    # $wb.Close() thrown from a null $wb would otherwise abort this block
    # before $xl.Quit() runs, orphaning the whole Excel process.
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
