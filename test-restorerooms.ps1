# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Restore workbook creates a print room the backup has and this workbook
# lacks (it used to skip that room's job records with "no matching print room
# sheet found").
#
# Backs up a populated copy (A: Example Print Room + Annexe), then restores
# that backup into a second FRESH copy (B) that has no Annexe at all, and
# checks the room was created - sheet, code, name, department - and its job
# records arrived. Restoring a second time must not create a second room.
# Drives COPIES in %TEMP%, never src\PrintJob.xlsm.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'

function New-WorkDir {
    $d = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $d | Out-Null
    return $d
}

$workDirA = New-WorkDir
$fA = Join-Path $workDirA 'PrintJob.xlsm'
Copy-Item $deliverable $fA
$workDirB = New-WorkDir
$fB = Join-Path $workDirB 'PrintJob.xlsm'
Copy-Item $deliverable $fB

$anyFail = $false

Write-Host '=== Back up copy A (Example Print Room + Annexe) ==='
$xlA = New-Object -ComObject Excel.Application
$xlA.Visible = $false
$xlA.DisplayAlerts = $false
$wbA = $null
$origAnnexeJobs = 0
$origDept = ''
try {
    $wbA = $xlA.Workbooks.Open($fA)
    Add-AnnexeFixture $xlA $wbA | Out-Null
    $xlA.Run('SetQuiet', $true)

    # A department on the room, so the backup has one to carry.
    $annexeWs = $wbA.Worksheets('Annexe')
    $xlA.Run('UnlockSheet', $annexeWs)
    $annexeWs.Names.Item('LOC_Dept').RefersToRange.Value2 = 'Fine Art'
    $xlA.Run('RelockSheet', $annexeWs)
    $origDept = [string]$xlA.Run('LocValue', $annexeWs, 'LOC_Dept')
    $origAnnexeJobs = $wbA.Worksheets('Annexe').ListObjects('tblJobs_ANNEX').ListRows.Count
    Write-Host "  Annexe: $origAnnexeJobs job row(s), department '$origDept'"

    $xlA.Run('BackupAll')
    $xlA.Run('SetQuiet', $false)
}
finally {
    if ($wbA) { try { $wbA.Close($false) } catch {} }
    $xlA.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlA)
}

Check ($origAnnexeJobs -ge 1) "the fixture's Annexe has job rows to back up (got $origAnnexeJobs)"
$annexeCsv = Get-ChildItem $workDirA -Filter 'PrintCosts-*.csv' | Where-Object { $_.Name -notlike '*-CATALOG-*' -and $_.Name -like '*ANNEX*' } | Select-Object -First 1
Check ($null -ne $annexeCsv) 'Annexe job CSV found'
$anyCsv = Get-ChildItem $workDirA -Filter 'PrintCosts-*-CATALOG-tblPrinters-*.csv' | Select-Object -First 1
Check ($null -ne $anyCsv) 'a backup file to point Restore at found'

# The department is written beside the location name in the header block.
$hdr = Get-Content $annexeCsv.FullName -TotalCount 8
Check (($hdr[5] -like '*Location department*') -and ($hdr[5] -like "*$origDept*")) "job CSV header names the department (row 6: $($hdr[5]))"

Write-Host ''
Write-Host '=== Restore into fresh copy B, which has no Annexe ==='
$xlB = New-Object -ComObject Excel.Application
$xlB.Visible = $false
$xlB.DisplayAlerts = $false
$wbB = $null
try {
    $wbB = $xlB.Workbooks.Open($fB)
    $xlB.Run('SetQuiet', $true)

    $names = @(); foreach ($ws in $wbB.Worksheets) { $names += $ws.Name }
    Check ($names -notcontains 'Annexe') 'fresh copy has no Annexe sheet before the restore'

    $xlB.Run('RestoreFromFileConfirmed', $anyCsv.FullName)
    Write-Host $xlB.Run('QuietLog')
    Start-Sleep -Milliseconds 300

    $names = @(); foreach ($ws in $wbB.Worksheets) { $names += $ws.Name }
    Check ($names -contains 'Annexe') 'Annexe sheet was created by the restore'
    $annexeWs = $wbB.Worksheets('Annexe')
    Check ([string]$xlB.Run('LocValue', $annexeWs, 'LOC_Code') -eq 'ANNEX') "created room has code ANNEX (got '$($xlB.Run('LocValue', $annexeWs, 'LOC_Code'))')"
    Check ([string]$xlB.Run('LocValue', $annexeWs, 'LOC_Name') -eq 'Annexe') 'created room has name Annexe'
    Check ([string]$xlB.Run('LocValue', $annexeWs, 'LOC_Dept') -eq $origDept) "created room has department '$origDept' (got '$($xlB.Run('LocValue', $annexeWs, 'LOC_Dept'))')"
    $lo = $annexeWs.ListObjects('tblJobs_ANNEX')
    Check ($lo.ListRows.Count -eq $origAnnexeJobs) "Annexe job table holds $origAnnexeJobs row(s) (got $($lo.ListRows.Count))"

    $aud = $wbB.Worksheets('_Audit').ListObjects('tblAudit')
    $actions = @(); for ($i = 1; $i -le $aud.ListRows.Count; $i++) { $actions += [string]$aud.ListRows($i).Range.Cells(1, 3).Value2 }
    Check ($actions -contains 'Add print room') 'audit log records the room creation'
    Check ($actions -contains 'Restore workbook') 'audit log records the restore'

    Write-Host ''
    Write-Host '=== Restoring again creates no second room ==='
    $xlB.Run('RestoreFromFileConfirmed', $anyCsv.FullName)
    Start-Sleep -Milliseconds 300
    $count = 0; foreach ($ws in $wbB.Worksheets) { if ($ws.Name -like 'Annexe*') { $count++ } }
    Check ($count -eq 1) "still exactly one Annexe sheet after a second restore (got $count)"
    $lo = $wbB.Worksheets('Annexe').ListObjects('tblJobs_ANNEX')
    Check ($lo.ListRows.Count -eq $origAnnexeJobs) "Annexe still $origAnnexeJobs row(s) (got $($lo.ListRows.Count))"

    $xlB.Run('SetQuiet', $false)
}
finally {
    if ($wbB) { try { $wbB.Close($false) } catch {} }
    $xlB.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlB)
}

Write-Host ''
if ($anyFail) { Write-Host 'FAIL' } else { Write-Host 'PASS' }
Remove-Item $workDirA -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $workDirB -Recurse -Force -ErrorAction SilentlyContinue
if ($anyFail) { exit 1 }
