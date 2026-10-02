# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# LOC_RollUnit as a DISPLAY unit (docs/ARCHITECTURE.md §16.3, 2026-10-01).
#
# On a Centimetres location Qty holds centimetres and the Unit column reads
# "cm"; Area m2 and Paper Cost still come out the same as for the equivalent
# length in metres, and the consolidated _Data range (what every report reads)
# always carries metres and "metres". Changing the setting rescales the
# existing roll rows, never twice, never touching Sheet stock.
#
# Covers:
#   - Metres sheet: Unit "metres", Qty as typed
#   - switching to Centimetres: roll Qty x100, Unit "cm", Area/Paper Cost unchanged
#   - re-picking the same unit does not rescale again
#   - Sheet stock Qty and Unit ("sheets") untouched by the setting
#   - a roll row typed on a Centimetres sheet: cost matches the metres equivalent
#   - _Data stays in metres for the cm rows (Qty/100, Unit "metres")
#   - switching back to Metres restores the original lengths
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself - see
# test-validation.ps1's header comment for why. Closes WITHOUT saving.

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
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Example Print Room')
    $lo = $main.ListObjects('tblJobs_MAIN')
    $data = $wb.Worksheets('_Data')

    $stkCol = $lo.Range.Column + (Col $lo 'Paper Stock') - 1
    $qtyCol = $lo.Range.Column + (Col $lo 'Qty') - 1
    $unitCol = $lo.Range.Column + (Col $lo 'Unit') - 1
    $areaCol = $lo.Range.Column + (Col $lo 'Area m2') - 1
    $costCol = $lo.Range.Column + (Col $lo 'Paper Cost') - 1
    $idCol = $lo.Range.Column + (Col $lo 'Job ID') - 1

    function New-Row {
        [void]$main.Activate()
        [void]$xl.Run('btnAddPrintJob')
        return $lo.ListRows($lo.ListRows.Count).Range.Row
    }
    function Set-Cell([int]$row, [int]$col, $v) { $c = $main.Cells($row, $col); if ($v -is [string]) { $c.Value2 = [string]$v } else { $c.Value2 = [double]$v }; Start-Sleep -Milliseconds 300 }
    function Near($a, $b) { return ([Math]::Abs([double]$a - [double]$b) -lt 0.000001) }
    # Finds a job's row in _Data (header row 9, data from row 10) by Job ID.
    function Data-Row($jobId) {
        $hdr = 9
        $idc = 1..60 | Where-Object { $data.Cells($hdr, $_).Value2 -eq 'Job ID' } | Select-Object -First 1
        for ($r = 10; $r -lt 400; $r++) { if ($data.Cells($r, $idc).Value2 -eq $jobId) { return $r } }
        return 0
    }
    function Data-Col($name) { return (1..60 | Where-Object { $data.Cells(9, $_).Value2 -eq $name } | Select-Object -First 1) }

    $rollUnit = $main.Range('LOC_RollUnit')
    $rollUnit.Value2 = 'Metres'
    Start-Sleep -Milliseconds 300

    # -------------------------------------------------------- Metres baseline
    Write-Host '=== Metres sheet ==='
    $r1 = New-Row
    Set-Cell $r1 $stkCol 'Canvas 914mm roll'
    Set-Cell $r1 $qtyCol 2.5
    $area1 = [double]$main.Cells($r1, $areaCol).Value2
    $cost1 = [double]$main.Cells($r1, $costCol).Value2
    Check ($main.Cells($r1, $unitCol).Value2 -eq 'metres') 'Unit reads "metres"'
    Check (Near $main.Cells($r1, $qtyCol).Value2 2.5) 'Qty is 2.5 as typed'
    Check ($cost1 -gt 0) "Paper Cost calculated ($cost1)"

    $r2 = New-Row
    Set-Cell $r2 $stkCol 'Gloss 200gsm SRA3 sheet'
    Set-Cell $r2 $qtyCol 40
    $costSheet = [double]$main.Cells($r2, $costCol).Value2

    # ------------------------------------------------- switch to Centimetres
    Write-Host ''
    Write-Host '=== Switch to Centimetres: existing roll rows rescale ==='
    $rollUnit.Value2 = 'Centimetres'
    Start-Sleep -Milliseconds 500
    Check (Near $main.Cells($r1, $qtyCol).Value2 250) "roll Qty 2.5 m now shows 250 (got $($main.Cells($r1, $qtyCol).Value2))"
    Check ($main.Cells($r1, $unitCol).Value2 -eq 'cm') 'roll Unit reads "cm"'
    Check (Near $main.Cells($r1, $areaCol).Value2 $area1) 'Area m2 unchanged (still metres-based)'
    Check (Near $main.Cells($r1, $costCol).Value2 $cost1) 'Paper Cost unchanged'
    Check (Near $main.Cells($r2, $qtyCol).Value2 40) 'Sheet Qty untouched (40)'
    Check ($main.Cells($r2, $unitCol).Value2 -eq 'sheets') 'Sheet Unit still "sheets"'
    Check (Near $main.Cells($r2, $costCol).Value2 $costSheet) 'Sheet Paper Cost unchanged'

    # Reports' Paper type filter on a Centimetres room: _Data converts roll
    # rows back to "metres", so Roll must still find the roll job and Sheet the
    # sheet job.
    $rep = $wb.Worksheets('Reports')
    $xl.CalculateFullRebuild()
    $rep.Range('F9').Value2 = 'Roll';  $xl.CalculateFullRebuild(); $rollJobs = [int]$rep.Range('B15').Text
    $rep.Range('F9').Value2 = 'Sheet'; $xl.CalculateFullRebuild(); $sheetJobs = [int]$rep.Range('B15').Text
    $rep.Range('F9').ClearContents() | Out-Null; $xl.CalculateFullRebuild()
    Check ($rollJobs -ge 1) "Reports Paper type = Roll finds the roll job on a cm room (got $rollJobs)"
    Check ($sheetJobs -ge 1) "Reports Paper type = Sheet finds the sheet job on a cm room (got $sheetJobs)"
    Check (($rollJobs + $sheetJobs) -eq [int]$rep.Range('B15').Text) 'Roll + Sheet account for every job on a cm room'

    # ---------------------------------------------- same unit re-picked
    Write-Host ''
    Write-Host '=== Re-picking the same unit does not rescale again ==='
    $rollUnit.Value2 = 'Centimetres'
    Start-Sleep -Milliseconds 500
    Check (Near $main.Cells($r1, $qtyCol).Value2 250) 'Qty still 250'

    # ------------------------------------------- new roll row typed in cm
    Write-Host ''
    Write-Host '=== Roll row typed in cm costs the same as its metres equivalent ==='
    $r3 = New-Row
    Set-Cell $r3 $stkCol 'Canvas 914mm roll'
    Set-Cell $r3 $qtyCol 250
    Check (Near $main.Cells($r3, $qtyCol).Value2 250) 'Qty stays 250 as typed (no rewrite)'
    Check ($main.Cells($r3, $unitCol).Value2 -eq 'cm') 'Unit reads "cm"'
    Check (Near $main.Cells($r3, $costCol).Value2 $cost1) 'Paper Cost equals the 2.5 m job'
    Check (Near $main.Cells($r3, $areaCol).Value2 $area1) 'Area m2 equals the 2.5 m job'

    # ----------------------------------------------------------- _Data
    Write-Host ''
    Write-Host '=== _Data (reports) always in metres ==='
    $xl.Calculate()
    $qCol = Data-Col 'Qty'; $uCol = Data-Col 'Unit'
    $id1 = $main.Cells($r1, $idCol).Value2
    $dr = Data-Row $id1
    Check ($dr -gt 0) "job $id1 found in _Data"
    if ($dr -gt 0) {
        Check (Near $data.Cells($dr, $qCol).Value2 2.5) "_Data Qty is 2.5 m (got $($data.Cells($dr, $qCol).Value2))"
        Check ($data.Cells($dr, $uCol).Value2 -eq 'metres') '_Data Unit is "metres"'
    }
    $id2 = $main.Cells($r2, $idCol).Value2
    $dr2 = Data-Row $id2
    if ($dr2 -gt 0) {
        Check (Near $data.Cells($dr2, $qCol).Value2 40) '_Data sheet Qty is 40'
        Check ($data.Cells($dr2, $uCol).Value2 -eq 'sheets') '_Data sheet Unit is "sheets"'
    } else { Check $false "sheet job $id2 found in _Data" }

    # ----------------------------------------------------- back to Metres
    Write-Host ''
    Write-Host '=== Back to Metres restores the lengths ==='
    $rollUnit.Value2 = 'Metres'
    Start-Sleep -Milliseconds 500
    Check (Near $main.Cells($r1, $qtyCol).Value2 2.5) 'row 1 back to 2.5'
    Check (Near $main.Cells($r3, $qtyCol).Value2 2.5) 'row 3 back to 2.5'
    Check ($main.Cells($r1, $unitCol).Value2 -eq 'metres') 'Unit back to "metres"'
    Check (Near $main.Cells($r2, $qtyCol).Value2 40) 'Sheet Qty still 40'

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
if ($script:anyFail) { exit 1 }
