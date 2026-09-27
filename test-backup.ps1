# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Full workbook backup / restore (snag list item 4a), package 8 of the
# 2026-09-22 post-phase-8 snag list.
#
# Backs up a populated copy (A), corrupts a second fresh copy (B) by
# emptying its Printers table and Annexe's job table, then restores B from
# A's backup and checks it matches A's original state. Drives two COPIES in
# %TEMP%, never src\PrintJob.xlsm.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'

$workDirA = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDirA | Out-Null
$fA = Join-Path $workDirA 'PrintJob.xlsm'
Copy-Item $deliverable $fA

$workDirB = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDirB | Out-Null
$fB = Join-Path $workDirB 'PrintJob.xlsm'
Copy-Item $deliverable $fB

$anyFail = $false
function Check([bool]$cond, [string]$msg) {
    Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    if (-not $cond) { $script:anyFail = $true }
}

Write-Host '=== Back up copy A ==='
$xlA = New-Object -ComObject Excel.Application
$xlA.Visible = $false
$xlA.DisplayAlerts = $false
$wbA = $null
$origPrinterCount = 0
$origPaperCount = 0
$origAnnexeJobs = 0
try {
    $wbA = $xlA.Workbooks.Open($fA)
    Add-AnnexeFixture $xlA $wbA | Out-Null
    $xlA.Run('SetQuiet', $true)

    $origPrinterCount = $wbA.Worksheets('Printers').ListObjects('tblPrinters').ListRows.Count
    $origPaperCount = $wbA.Worksheets('Papers').ListObjects('tblPapers').ListRows.Count
    $origAnnexeJobs = $wbA.Worksheets('Annexe').ListObjects('tblJobs_ANNEX').ListRows.Count
    Write-Host "  original counts: Printers=$origPrinterCount Papers=$origPaperCount Annexe jobs=$origAnnexeJobs"

    $xlA.Run('BackupAll')
    Write-Host $xlA.Run('QuietLog')

    $xlA.Run('SetQuiet', $false)
}
finally {
    if ($wbA) { try { $wbA.Close($false) } catch {} }
    $xlA.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlA)
}

$catalogCsvs = Get-ChildItem $workDirA -Filter 'PrintCosts-*-CATALOG-*.csv'
Check ($catalogCsvs.Count -eq 8) "8 catalogue CSVs produced (got $($catalogCsvs.Count))"
$printersCsv = $catalogCsvs | Where-Object { $_.Name -like '*-CATALOG-tblPrinters-*' } | Select-Object -First 1
Check ($null -ne $printersCsv) "tblPrinters backup CSV found"

$locationCsvs = Get-ChildItem $workDirA -Filter 'PrintCosts-*.csv' | Where-Object { $_.Name -notlike '*-CATALOG-*' }
Check ($locationCsvs.Count -eq 2) "2 location job CSVs produced (got $($locationCsvs.Count))"

Write-Host ''
Write-Host '=== Corrupt copy B, then restore it from copy A''s backup ==='
$xlB = New-Object -ComObject Excel.Application
$xlB.Visible = $false
$xlB.DisplayAlerts = $false
$wbB = $null
try {
    $wbB = $xlB.Workbooks.Open($fB)
    Add-AnnexeFixture $xlB $wbB | Out-Null
    $xlB.Run('SetQuiet', $true)

    # Both sheets are protected (§9.4) - UnlockSheet/RelockSheet are the same
    # VBA-side unlock this workbook's own code uses for every write, called
    # directly here since this corruption step (unlike the restore itself)
    # has no VBA entry point of its own to reuse.
    $printersWs = $wbB.Worksheets('Printers')
    $xlB.Run('UnlockSheet', $printersWs)
    $loPrinters = $printersWs.ListObjects('tblPrinters')
    while ($loPrinters.ListRows.Count -gt 0) { $loPrinters.ListRows(1).Delete() }
    $xlB.Run('RelockSheet', $printersWs)

    $annexeWs = $wbB.Worksheets('Annexe')
    $xlB.Run('UnlockSheet', $annexeWs)
    $loAnnexe = $annexeWs.ListObjects('tblJobs_ANNEX')
    while ($loAnnexe.ListRows.Count -gt 0) { $loAnnexe.ListRows(1).Delete() }
    $xlB.Run('RelockSheet', $annexeWs)

    Check ($loPrinters.ListRows.Count -eq 0) "tblPrinters is empty before restore"
    Check ($loAnnexe.ListRows.Count -eq 0) "Annexe job table is empty before restore"

    $xlB.Run('RestoreFromFileConfirmed', $printersCsv.FullName)
    Write-Host $xlB.Run('QuietLog')

    # Settling pause (2026-09-27, reproduced directly): reading a ListObject's
    # ListRows.Count immediately after RestoreFromFileConfirmed returns can
    # come back stale (observed tblPapers.ListRows.Count as 0 right after the
    # call, then correctly 9 moments later with no code in between other than
    # a Write-Host) - the same class of Excel/COM property-read timing gotcha
    # ToggleReducedView's own callers already pause for elsewhere in this
    # test suite.
    Start-Sleep -Milliseconds 300

    $loPrinters = $wbB.Worksheets('Printers').ListObjects('tblPrinters')
    $loPapers = $wbB.Worksheets('Papers').ListObjects('tblPapers')
    $loAnnexe = $wbB.Worksheets('Annexe').ListObjects('tblJobs_ANNEX')

    Check ($loPrinters.ListRows.Count -eq $origPrinterCount) "tblPrinters restored to $origPrinterCount rows (got $($loPrinters.ListRows.Count))"
    Check ($loPapers.ListRows.Count -eq $origPaperCount) "tblPapers unchanged at $origPaperCount rows (got $($loPapers.ListRows.Count)) - was not part of this restore's file set"
    Check ($loAnnexe.ListRows.Count -eq $origAnnexeJobs) "Annexe job table restored to $origAnnexeJobs row(s) (got $($loAnnexe.ListRows.Count))"

    $aud = $wbB.Worksheets('_Audit').ListObjects('tblAudit')
    $lastRow = $aud.ListRows.Count
    $lastAction = [string]$aud.ListRows($lastRow).Range.Cells(1, 3).Value2
    Check ($lastAction -eq 'Restore workbook') "audit log entry recorded (got '$lastAction')"

    Write-Host ''
    Write-Host '=== Restoring again is idempotent (no duplicate rows) ==='
    $xlB.Run('RestoreFromFileConfirmed', $printersCsv.FullName)
    $loPrinters = $wbB.Worksheets('Printers').ListObjects('tblPrinters')
    Check ($loPrinters.ListRows.Count -eq $origPrinterCount) "tblPrinters still $origPrinterCount rows after a second restore (got $($loPrinters.ListRows.Count))"

    $xlB.Run('SetQuiet', $false)
}
finally {
    if ($wbB) { try { $wbB.Close($false) } catch {} }
    $xlB.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlB)
}

Write-Host ''
if ($anyFail) {
    Write-Host 'FAIL'
} else {
    Write-Host 'PASS'
}
Remove-Item $workDirA -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $workDirB -Recurse -Force -ErrorAction SilentlyContinue
if ($anyFail) { exit 1 }
