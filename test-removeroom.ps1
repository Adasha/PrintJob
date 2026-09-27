# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Remove print room (modRegistry.RemovePrintRoom/DeletePrintRoom): the
# in-app counterpart to Add print room. DeletePrintRoom (the testable core)
# is called directly, not RemovePrintRoom - RemovePrintRoom's own picker and
# typed-name confirmation both go through Application.InputBox, which has no
# SetQuiet bypass (unlike Ask/Say) and would hang unattended COM automation
# waiting for an answer nobody can give. Same reasoning test-addroom.ps1
# gives for calling CreatePrintRoom directly instead of AddPrintRoom.
#
# Checks: DeletePrintRoom actually deletes the sheet, logs it to tblAudit,
# and the Refresh Locations pass it triggers reports the room as gone -
# the same "self-healing afterwards" story §4.3 documents for a manual
# deletion.
#
# The workbook is closed WITHOUT saving.

$ErrorActionPreference = 'Stop'
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$vba = @'
Public Function TestRemoveRoom() As String
    Dim r As String, room2 As Worksheet, before As Long, after As Long, removedName As String

    On Error GoTo Fail
    Application.DisplayAlerts = False

    before = ThisWorkbook.Worksheets.Count
    r = r & "sheets before: " & before & vbCrLf

    ' Give it a second room to remove - deleting the only one is
    ' RemovePrintRoom's own guard, upstream of DeletePrintRoom, and not
    ' exercised here for the InputBox reason above.
    SetQuiet True
    Set room2 = CreatePrintRoom("Test Room To Remove", "Facilities", "REMOVE")
    SetQuiet False
    If room2 Is Nothing Then
        TestRemoveRoom = r & "FAIL: CreatePrintRoom returned Nothing" & vbCrLf
        Exit Function
    End If
    removedName = room2.Name
    r = r & "second room created: '" & removedName & "' LOC_Code=" & LocValue(room2, "LOC_Code") & vbCrLf
    r = r & "sheets with second room: " & ThisWorkbook.Worksheets.Count & vbCrLf

    SetQuiet True
    DeletePrintRoom room2
    r = r & "Refresh Locations said:" & vbCrLf & QuietLog
    SetQuiet False

    after = ThisWorkbook.Worksheets.Count
    r = r & "sheets after removal: " & after & vbCrLf
    r = r & "removed sheet still present: " & SheetExistsPublic(removedName) & vbCrLf
    r = r & "registry after removal:" & vbCrLf & RegList
    r = r & "last audit row: " & LastAuditLine() & vbCrLf

    TestRemoveRoom = r
    Exit Function
Fail:
    TestRemoveRoom = r & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
End Function

Private Function SheetExistsPublic(ByVal Nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Nm)
    On Error GoTo 0
    SheetExistsPublic = Not ws Is Nothing
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

Private Function LastAuditLine() As String
    Dim lo As ListObject, i As Long, c As Long, ln As String
    Set lo = Tbl("tblAudit")
    If lo Is Nothing Then
        LastAuditLine = "(no tblAudit)"
        Exit Function
    End If
    i = lo.ListRows.Count
    If i = 0 Then
        LastAuditLine = "(no audit rows)"
        Exit Function
    End If
    For c = 1 To lo.ListColumns.Count
        ln = ln & CStr(lo.ListRows(i).Range.Cells(1, c).Text) & " | "
    Next c
    LastAuditLine = ln
End Function
'@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)

    $m = $null
    for ($attempt = 1; $attempt -le 5 -and -not $m; $attempt++) {
        try { $m = $wb.VBProject.VBComponents.Add(1) }
        catch {
            if ($attempt -eq 5) { throw }
            Start-Sleep -Seconds $attempt
        }
    }
    $m.Name = 'modRemoveRoomTest'
    $m.CodeModule.AddFromString($vba)
    Write-Host $xl.Run('TestRemoveRoom')
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
