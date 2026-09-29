# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# 0.10.9: the Status column is narrow and every row with a problem carries a
# hover note holding its full message (modInit.RefreshStatusNotes).
#
# Checks: Status is narrow; a note exists exactly on the rows whose Status is
# neither blank nor "OK", and its text matches; a note is removed when the row
# becomes OK. Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$failed = 0
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $ws = $wb.Worksheets('Example Print Room')
    $lo = $ws.ListObjects('tblJobs_MAIN')
    $ws.Activate()

    $sc = $lo.ListColumns('Status')
    $px = $sc.Range.Cells(1, 1).ColumnWidth * 7 + 5
    if ($px -gt 130) { Write-Host "FAIL: Status is $px px wide, expected about 120"; $failed++ }
    else { Write-Host "OK: Status is about $([int]$px) px wide" }

    # A new row is incomplete, so its Status lists issues.
    $xl.Run('btnAddPrintJob')
    $xl.Run('RefreshStatusNotes', $ws)

    function Check-Notes($label) {
        $bad = 0; $withNote = 0
        $rows = $lo.ListRows.Count
        for ($i = 1; $i -le $rows; $i++) {
            $c = $sc.DataBodyRange.Cells($i, 1)
            $txt = [string]$c.Value2
            $cm = $c.Comment
            $wantNote = ($txt -ne '' -and $txt -ne 'OK')
            if ($wantNote) {
                if ($null -eq $cm) { Write-Host "FAIL ($label): row $i status '$txt' has no note"; $bad++ }
                else {
                    $withNote++
                    $bullet = [string][char]0x2022 + ' '
                    $flat = (($cm.Text()).Replace($bullet, '').Replace("`n", '; ')).Trim()
                    if ($flat -ne $txt) { Write-Host "FAIL ($label): row $i note '$flat' <> status '$txt'"; $bad++ }
                }
            } elseif ($null -ne $cm) { Write-Host "FAIL ($label): row $i status '$txt' should have no note"; $bad++ }
        }
        Write-Host ("{0}: {1} rows, {2} with notes, {3} problems" -f $label, $rows, $withNote, $bad)
        return @($bad, $withNote)
    }

    $r = Check-Notes 'after add'
    $failed += $r[0]
    if ($r[1] -lt 1) { Write-Host 'FAIL: expected at least one row with a note after adding an incomplete job'; $failed++ }

    # Remove the note-bearing state: fill nothing, just re-run after a refresh.
    $xl.Run('btnRefreshLocations')
    $r = Check-Notes 'after refresh'
    $failed += $r[0]
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($failed -gt 0) { Write-Host "FAIL: $failed problem(s)"; exit 1 }
Write-Host 'OK: status notes match'
