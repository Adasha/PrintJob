# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# An older workbook whose real print room kept the shipped sample's code (MAIN),
# whether or not it was renamed: Restore workbook used to import its jobs INTO
# the sample room, which kept its name, department and permitted printers. It now
# reuses the sample's sheet for the room, so the restored room comes back with
# its own settings (here: two permitted printers, one of them not in the sample).
# Drives COPIES in %TEMP%, never src\PrintJob.xlsm.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'

function New-WorkDir {
    $d = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $d | Out-Null
    return $d
}

function Restore-Case([string]$label, [string]$roomName, [string]$sheetExpected) {
    Write-Host "=== $label ==="
    $dirA = New-WorkDir; $fA = Join-Path $dirA 'Old.xlsm'; Copy-Item $deliverable $fA
    $dirB = New-WorkDir; $fB = Join-Path $dirB 'PrintJob.xlsm'; Copy-Item $deliverable $fB
    $printers = ''
    $xlA = New-Object -ComObject Excel.Application
    $xlA.Visible = $false; $xlA.DisplayAlerts = $false
    $wbA = $null
    $jobsA = 0
    try {
        $wbA = $xlA.Workbooks.Open($fA)
        Start-Sleep -Seconds 2
        $xlA.Run('SetQuiet', $true)
        $ws = Invoke-ComRetry { $wbA.Worksheets('Example Print Room') }
        $lo = $wbA.Worksheets('Printers').ListObjects('tblPrinters')
        $c = $lo.ListColumns('Model').DataBodyRange
        $printers = ([string]$c.Cells(1, 1).Value2) + ';' + ([string]$c.Cells(2, 1).Value2)
        Invoke-ComRetry { $xlA.Run('UnlockSheet', $ws) } | Out-Null
        $ws.Names.Item('LOC_Name').RefersToRange.Value2 = $roomName
        $ws.Names.Item('LOC_Dept').RefersToRange.Value2 = 'Photography'
        $ws.Names.Item('LOC_Printers').RefersToRange.Value2 = $printers
        $ws.Names.Item('LOC_DefPrinter').RefersToRange.Value2 = ([string]$c.Cells(1, 1).Value2)
        Invoke-ComRetry { $xlA.Run('RelockSheet', $ws) } | Out-Null
        $jobsA = $ws.ListObjects('tblJobs_MAIN').ListRows.Count
        Check ($jobsA -ge 1) "old room has job rows ($jobsA)"
        Check ($printers -match ';.') "old room permits two printers ($printers)"
        $wbA.Save()
        $xlA.Run('SetQuiet', $false)
    }
    finally {
        if ($wbA) { try { $wbA.Close($false) } catch {} }
        $xlA.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlA)
    }

    $xlB = New-Object -ComObject Excel.Application
    $xlB.Visible = $false; $xlB.DisplayAlerts = $false
    $wbB = $null
    try {
        $wbB = $xlB.Workbooks.Open($fB)
        Start-Sleep -Seconds 2
        $xlB.Run('SetQuiet', $true)
        Invoke-ComRetry -Attempts 1 { $xlB.Run('RestoreFromFileConfirmed', $fA) } | Out-Null
        Start-Sleep -Milliseconds 300

        $names = @(); foreach ($s in $wbB.Worksheets) { $names += $s.Name }
        Check ($names -contains $sheetExpected) "the restored room's sheet is '$sheetExpected' (sheets: $($names -join ', '))"
        $room = $wbB.Worksheets($sheetExpected)
        Check ([string]$xlB.Run('LocValue', $room, 'LOC_Code') -eq 'MAIN') 'code kept as MAIN'
        Check ([string]$xlB.Run('LocValue', $room, 'LOC_Name') -eq $roomName) 'room name restored'
        Check ([string]$xlB.Run('LocValue', $room, 'LOC_Dept') -eq 'Photography') 'department restored'
        Check ([string]$xlB.Run('LocValue', $room, 'LOC_Printers') -eq $printers) "permitted printers restored (got '$($xlB.Run('LocValue', $room, 'LOC_Printers'))')"
        $lo = $room.ListObjects('tblJobs_MAIN')
        Check ($lo.ListRows.Count -eq $jobsA) "job rows match the old room ($jobsA, got $($lo.ListRows.Count))"
        $xlB.Run('SetQuiet', $false)
    }
    finally {
        if ($wbB) { try { $wbB.Close($false) } catch {} }
        $xlB.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlB)
    }
    Remove-Item $dirA, $dirB -Recurse -Force -ErrorAction SilentlyContinue
}

Restore-Case 'Sample room renamed in the older workbook' 'Photo Studio' 'Photo Studio'
Restore-Case 'Sample room NOT renamed, only given its own printers' 'Example Print Room' 'Example Print Room'

Write-Host ''
if ($anyFail) { Write-Host 'FAIL' } else { Write-Host 'PASS' }
if ($anyFail) { exit 1 }
