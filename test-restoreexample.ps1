# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Restore workbook removes the shipped sample room ("Example Print Room", code
# MAIN) once the backup's own print rooms are in - but only while it is still
# the untouched sample:
#   1. backup holds Annexe only          -> Example Print Room removed
#   2. backup holds MAIN and Annexe      -> Example Print Room kept (it IS a restored room)
#   3. backup holds Annexe only, but the example room has a real job -> kept
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

function Sheet-Names($wb) { $n = @(); foreach ($ws in $wb.Worksheets) { $n += $ws.Name }; return $n }

# Restores $csv into a fresh copy and returns the sheet names afterwards.
# $prep, when given, runs against the open fresh workbook first.
function Restore-Fresh([string]$csv, [scriptblock]$prep) {
    $dir = New-WorkDir
    $f = Join-Path $dir 'PrintJob.xlsm'
    Copy-Item $deliverable $f
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $null
    try {
        $wb = $xl.Workbooks.Open($f)
        $xl.Run('SetQuiet', $true)
        if ($prep) { & $prep $xl $wb }
        $xl.Run('RestoreFromFileConfirmed', $csv)
        Write-Host $xl.Run('QuietLog')
        Start-Sleep -Milliseconds 300
        $names = Sheet-Names $wb
        $xl.Run('SetQuiet', $false)
        return $names
    }
    finally {
        if ($wb) { try { $wb.Close($false) } catch {} }
        $xl.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host '=== Back up a copy holding Example Print Room + Annexe ==='
$dirA = New-WorkDir
$fA = Join-Path $dirA 'PrintJob.xlsm'
Copy-Item $deliverable $fA
$xlA = New-Object -ComObject Excel.Application
$xlA.Visible = $false
$xlA.DisplayAlerts = $false
$wbA = $null
try {
    $wbA = $xlA.Workbooks.Open($fA)
    Add-AnnexeFixture $xlA $wbA | Out-Null
    $xlA.Run('SetQuiet', $true)
    $xlA.Run('BackupAll')
    $xlA.Run('SetQuiet', $false)
}
finally {
    if ($wbA) { try { $wbA.Close($false) } catch {} }
    $xlA.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlA)
}

$full = Get-ChildItem $dirA -Filter 'PrintCosts-*-CATALOG-tblPrinters-*.csv' | Select-Object -First 1
Check ($null -ne $full) 'a backup file to point Restore at found'
$mainCsv = Get-ChildItem $dirA -Filter 'PrintCosts-*.csv' | Where-Object { $_.Name -notlike '*-CATALOG-*' -and $_.Name -like '*MAIN*' } | Select-Object -First 1
Check ($null -ne $mainCsv) 'MAIN job CSV found'

Write-Host ''
Write-Host '=== 2. Backup holds MAIN and Annexe: Example Print Room stays ==='
$names = Restore-Fresh $full.FullName $null
Check ($names -contains 'Example Print Room') 'Example Print Room kept when the backup includes MAIN'
Check ($names -contains 'Annexe') 'Annexe created'

# Take MAIN out of the backup set for the remaining cases.
$mainCsv | Remove-Item -Force

Write-Host ''
Write-Host '=== 1. Backup holds Annexe only: Example Print Room is removed ==='
$names = Restore-Fresh $full.FullName $null
Check ($names -contains 'Annexe') 'Annexe created'
Check ($names -notcontains 'Example Print Room') "Example Print Room removed (sheets: $($names -join ', '))"

Write-Host ''
Write-Host '=== 3. Example room holds a real job: it stays ==='
$names = Restore-Fresh $full.FullName {
    param($xl, $wb)
    $ws = $wb.Worksheets('Example Print Room')
    $lo = $ws.ListObjects('tblJobs_MAIN')
    $xl.Run('UnlockSheet', $ws)
    $row = $lo.ListRows.Add()
    $row.Range.Cells(1, $lo.ListColumns('Job ID').Index).Value2 = 'UNI-MAIN-00013'
    $xl.Run('RelockSheet', $ws)
}
Check ($names -contains 'Annexe') 'Annexe created'
Check ($names -contains 'Example Print Room') 'Example Print Room kept because it holds a job that is not a sample'

Write-Host ''
if ($anyFail) { Write-Host 'FAIL' } else { Write-Host 'PASS' }
Remove-Item $dirA -Recurse -Force -ErrorAction SilentlyContinue
if ($anyFail) { exit 1 }
