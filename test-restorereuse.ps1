# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# An older workbook whose real print room kept the shipped sample's code (MAIN)
# but was renamed: Restore workbook used to import its jobs INTO the sample
# room, which kept the name "Example Print Room", its department and its
# permitted printers. It now reuses the sample's sheet for the room, so the
# restored room comes back with its own name, department, defaults and printers.
# Drives COPIES in %TEMP%, never src\PrintJob.xlsm.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'

function New-WorkDir {
    $d = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $d | Out-Null
    return $d
}

$dirA = New-WorkDir; $fA = Join-Path $dirA 'Old.xlsm'; Copy-Item $deliverable $fA
$dirB = New-WorkDir; $fB = Join-Path $dirB 'PrintJob.xlsm'; Copy-Item $deliverable $fB

Write-Host '=== Older workbook: sample room renamed, with its own settings ==='
$xlA = New-Object -ComObject Excel.Application
$xlA.Visible = $false; $xlA.DisplayAlerts = $false
$wbA = $null
$jobsA = 0
try {
    $wbA = $xlA.Workbooks.Open($fA)
    Start-Sleep -Seconds 2
    $xlA.Run('SetQuiet', $true)
    $ws = Invoke-ComRetry { $wbA.Worksheets('Example Print Room') }
    Invoke-ComRetry { $xlA.Run('UnlockSheet', $ws) } | Out-Null
    $ws.Names.Item('LOC_Name').RefersToRange.Value2 = 'Photo Studio'
    $ws.Names.Item('LOC_Dept').RefersToRange.Value2 = 'Photography'
    $ws.Names.Item('LOC_Printers').RefersToRange.Value2 = 'Epson SureColor P9500'
    $ws.Names.Item('LOC_DefPrinter').RefersToRange.Value2 = 'Epson SureColor P9500'
    Invoke-ComRetry { $xlA.Run('RelockSheet', $ws) } | Out-Null
    $jobsA = $ws.ListObjects('tblJobs_MAIN').ListRows.Count
    Check ($jobsA -ge 1) "old room has job rows ($jobsA)"
    $wbA.Save()
    $xlA.Run('SetQuiet', $false)
}
finally {
    if ($wbA) { try { $wbA.Close($false) } catch {} }
    $xlA.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlA)
}

Write-Host '=== Restore into a fresh copy (still has the sample room) ==='
$xlB = New-Object -ComObject Excel.Application
$xlB.Visible = $false; $xlB.DisplayAlerts = $false
$wbB = $null
try {
    $wbB = $xlB.Workbooks.Open($fB)
    Start-Sleep -Seconds 2
    $xlB.Run('SetQuiet', $true)
    Invoke-ComRetry -Attempts 1 { $xlB.Run('RestoreFromFileConfirmed', $fA) } | Out-Null
    Write-Host $xlB.Run('QuietLog')
    Start-Sleep -Milliseconds 300

    $names = @(); foreach ($s in $wbB.Worksheets) { $names += $s.Name }
    Check ($names -notcontains 'Example Print Room') 'no sheet is still called Example Print Room'
    Check ($names -contains 'Photo Studio') "the restored room has its own sheet name (sheets: $($names -join ', '))"
    $room = $wbB.Worksheets('Photo Studio')
    Check ([string]$xlB.Run('LocValue', $room, 'LOC_Code') -eq 'MAIN') 'code kept as MAIN'
    Check ([string]$xlB.Run('LocValue', $room, 'LOC_Name') -eq 'Photo Studio') 'room name restored'
    Check ([string]$xlB.Run('LocValue', $room, 'LOC_Dept') -eq 'Photography') 'department restored'
    Check ([string]$xlB.Run('LocValue', $room, 'LOC_Printers') -eq 'Epson SureColor P9500') "permitted printers restored (got '$($xlB.Run('LocValue', $room, 'LOC_Printers'))')"
    Check ([string]$xlB.Run('LocValue', $room, 'LOC_DefPrinter') -eq 'Epson SureColor P9500') 'default printer restored'
    $lo = $room.ListObjects('tblJobs_MAIN')
    Check ($lo.ListRows.Count -eq $jobsA) "job rows match the old room ($jobsA, got $($lo.ListRows.Count))"
    if ($lo.ListRows.Count -ne $jobsA) { for ($i=1; $i -le $lo.ListRows.Count; $i++) { Write-Host ('  row {0}: [{1}] [{2}] [{3}]' -f $i, $lo.ListRows($i).Range.Cells(1,1).Text, $lo.ListRows($i).Range.Cells(1,2).Text, $lo.ListRows($i).Range.Cells(1,3).Text) } }
    $xlB.Run('SetQuiet', $false)
}
finally {
    if ($wbB) { try { $wbB.Close($false) } catch {} }
    $xlB.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlB)
}

Write-Host ''
if ($anyFail) { Write-Host 'FAIL' } else { Write-Host 'PASS' }
Remove-Item $dirA, $dirB -Recurse -Force -ErrorAction SilentlyContinue
if ($anyFail) { exit 1 }
