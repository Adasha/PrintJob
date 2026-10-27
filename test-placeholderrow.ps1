# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# An empty print room keeps one blank row. With a single printer permitted,
# re-binding the drop-downs (Select printers... OK, Refresh Locations) wrote
# that printer into the blank row, which then counted as a job: "1 print job -
# never exported", and New could not reuse the row (found in the 2026-10-27
# Mac checks; fixed in 0.11.4, modLists.BindColumns). A real job still gets
# the single printer filled in.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$room = 'One Printer Room'; $printer = 'Xerox Versant 180'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $ws = $xl.Run('CreatePrintRoom', $room, '', 'ONEPRN')
    Check ($null -ne $ws) "CreatePrintRoom made '$room'"
    $lo = $ws.ListObjects('tblJobs_ONEPRN')
    function ColTexts([string]$h) {
        $c = Col $lo $h
        return @(for ($i = 1; $i -le $lo.ListRows.Count; $i++) { [string]$lo.ListRows($i).Range.Cells(1, $c).Text })
    }

    # ------------------------------------------------------------- blank row
    Write-Host '=== A single permitted printer is not written into the blank row ==='
    $xl.Run('UnlockSheet', $ws)
    $ws.Names.Item('LOC_Printers').RefersToRange.Value2 = $printer
    $xl.Run('RelockSheet', $ws)
    $xl.Run('BindColumns', $ws)
    $p = ColTexts 'Printer'
    Check (($p -join '|') -eq '') "the blank row's Printer is still empty (got '$($p -join '|')')"
    $got = [string]$xl.Run('RowCount', $lo)
    Check ($got -eq '0') "the room still has no jobs (got '$got')"
    $got = [string]$xl.Run('ExportStatusText', $ws)
    Check ($got -eq 'No print jobs recorded.') "the export line says there is nothing to export (got '$got')"

    # -------------------------------------------------------------- real job
    Write-Host ''
    Write-Host '=== New still fills in the single printer, in the blank row ==='
    $xl.Run('AddPrintJob', $ws)
    Check ($lo.ListRows.Count -eq 1) "New reused the blank row (rows: $($lo.ListRows.Count))"
    $p = ColTexts 'Printer'
    Check (($p -join '|') -eq $printer) "the new job's Printer is $printer (got '$($p -join '|')')"
    $id = ColTexts 'Job ID'
    Check ($id[0].StartsWith('UNI-ONEPRN-')) "the new job has a Job ID (got '$($id -join '|')')"
    $xl.Run('BindColumns', $ws)
    $got = [string]$xl.Run('RowCount', $lo)
    Check ($got -eq '1') "re-binding leaves one job (got '$got')"

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'DONE'
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
