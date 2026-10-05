# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# 0.10.9: the Status column is narrow and every row with a problem carries a
# hover note holding its full message (modInit.RefreshStatusNotes).
#
# Checks: Status is narrow; a note exists exactly on the rows whose Status is
# neither blank nor "OK", and its text matches; a note is removed when the row
# becomes OK. Also (moved here from the old test-phase8.ps1, 0.10.25): the
# amber warning fill on every non-"OK" Status cell, on every location sheet,
# and that re-running InitialiseWorkbook does not stack a second rule.
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
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
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.Run('SetQuiet', $true)
    $ws = $wb.Worksheets('Example Print Room')
    $lo = $ws.ListObjects('tblJobs_MAIN')
    $ws.Activate()

    $sc = $lo.ListColumns('Status')
    $px = $sc.Range.Cells(1, 1).ColumnWidth * 7 + 5
    if ($px -gt 110) { Write-Host "FAIL: Status is $px px wide, expected about 90"; $failed++ }
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

    # ---- warning fill: every non-"OK" Status cell, on every location sheet
    Write-Host ''
    Write-Host '=== Warning state: Status column conditional formatting ==='
    foreach ($sheetTable in @(@('Example Print Room', 'tblJobs_MAIN'), @('Annexe', 'tblJobs_ANNEX'))) {
        $wsx = $wb.Worksheets($sheetTable[0])
        $rng = $wsx.ListObjects($sheetTable[1]).ListColumns('Status').DataBodyRange
        $cnt = $rng.FormatConditions.Count
        if ($cnt -lt 1) { Write-Host "FAIL: $($sheetTable[0]) Status has no conditional format"; $failed++; continue }
        $fc = $rng.FormatConditions.Item(1)
        if (-not ($fc.Formula1 -like '*<>*OK*')) { Write-Host "FAIL: $($sheetTable[0]) rule is not a <> OK test ($($fc.Formula1))"; $failed++ }
        elseif ($fc.Interior.Color -ne 49407) { Write-Host "FAIL: $($sheetTable[0]) fill is $($fc.Interior.Color), expected amber 49407"; $failed++ }
        else { Write-Host "OK: $($sheetTable[0]) Status warns in amber when not OK" }
    }
    # The incomplete row added above reads something other than OK, so the
    # rule must actually fire on it, not just exist.
    $lastStatus = $lo.ListRows($lo.ListRows.Count).Range.Cells(1, $sc.Index)
    if ([string]$lastStatus.Text -eq 'OK') { Write-Host 'FAIL: the added incomplete row reads OK'; $failed++ }
    elseif ($lastStatus.DisplayFormat.Interior.Color -ne 49407) { Write-Host "FAIL: incomplete row Status is not amber (fill $($lastStatus.DisplayFormat.Interior.Color))"; $failed++ }
    else { Write-Host "OK: the incomplete row's Status ('$($lastStatus.Text)') is shown amber" }

    # Re-running setup must not stack a second rule.
    $xl.Run('InitialiseWorkbook')
    $cnt2 = $wb.Worksheets('Example Print Room').ListObjects('tblJobs_MAIN').ListColumns('Status').DataBodyRange.FormatConditions.Count
    if ($cnt2 -ne 1) { Write-Host "FAIL: Status has $cnt2 rules after a second InitialiseWorkbook (expected 1)"; $failed++ }
    else { Write-Host 'OK: re-running InitialiseWorkbook leaves one Status rule' }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($failed -gt 0) { Write-Host "FAIL: $failed problem(s)"; exit 1 }
Write-Host 'OK: status notes match'
