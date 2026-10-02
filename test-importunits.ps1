# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mixed-unit import (docs/ARCHITECTURE.md §16.2 punch list; conversion added
# in 0.10.11, modImport.ConvertImportedRollQty).
#
# An export carries Qty and Unit as the room displayed them: a Centimetres
# room writes "cm" with the length in centimetres, a Metres room "metres".
# Import reads the file's own Unit and converts a roll Qty into the target
# room's unit. Sheet stock never converts.
#
# Setup: Example Print Room stays on Metres, the Annexe fixture is set to
# Centimetres.
#   - metres file -> cm room: roll Qty x100, Unit "cm", costs unchanged
#   - cm file -> metres room: roll Qty /100, Unit "metres", costs unchanged
#   - same-unit control: a metres file back into its own metres room is untouched
#   - sheet stock Qty/Unit unchanged in both directions
#
# PickImportFile opens an OS dialog, so this drives ReadImportRows +
# ApplyImportConfirmed directly, as test-import.ps1 does. Drives a COPY in
# %TEMP%, never src\PrintJob.xlsm itself - see test-validation.ps1's header
# comment for why. Closes WITHOUT saving.

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
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Example Print Room')
    $annexe = $wb.Worksheets('Annexe')
    $loM = $main.ListObjects('tblJobs_MAIN')
    $loA = $annexe.ListObjects('tblJobs_ANNEX')

    function Near($a, $b) { return ([Math]::Abs([double]$a - [double]$b) -lt 0.000001) }
    function Set-Cell($ws, [int]$row, [int]$col, $v) { $c = $ws.Cells($row, $col); if ($v -is [string]) { $c.Value2 = [string]$v } else { $c.Value2 = [double]$v }; Start-Sleep -Milliseconds 300 }
    function Cc($lo, $name) { return $lo.Range.Column + (Col $lo $name) - 1 }
    function New-Row($ws, $lo) {
        [void]$ws.Activate()
        [void]$xl.Run('btnAddPrintJob')
        return $lo.ListRows($lo.ListRows.Count).Range.Row
    }
    # Sheet row holding $jobId, or 0.
    function Find-Row($ws, $lo, $jobId) {
        $idc = Cc $lo 'Job ID'
        for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
            $r = $lo.ListRows($i).Range.Row
            if ([string]$ws.Cells($r, $idc).Value2 -eq $jobId) { return $r }
        }
        return 0
    }
    function Newest-Csv($code) {
        return Get-ChildItem $workDir -Filter "PrintCosts-*-$code-*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    function Import-Csv-Into($ws, $csv) {
        $rows = $xl.Run('ReadImportRows', $csv.FullName)
        $xl.Run('ApplyImportConfirmed', $ws, $rows)
        # ApplyImportConfirmed's Check-sheet sweep can leave Excel busy; see test-import.ps1.
        Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
        Start-Sleep -Milliseconds 300
    }

    $main.Range('LOC_RollUnit').Value2 = 'Metres'
    $annexe.Range('LOC_RollUnit').Value2 = 'Centimetres'
    Start-Sleep -Milliseconds 500

    # ---------------------------------------- jobs typed in Main (metres)
    Write-Host '=== Setup: roll 2.5 m and sheet x40 in the Metres room ==='
    $mStk = Cc $loM 'Paper Stock'; $mQty = Cc $loM 'Qty'; $mUnit = Cc $loM 'Unit'
    $mArea = Cc $loM 'Area m2'; $mCost = Cc $loM 'Paper Cost'; $mId = Cc $loM 'Job ID'

    $rRoll = New-Row $main $loM
    Set-Cell $main $rRoll $mStk 'Canvas 914mm roll'
    Set-Cell $main $rRoll $mQty 2.5
    $rollId = [string]$main.Cells($rRoll, $mId).Value2
    $rollArea = [double]$main.Cells($rRoll, $mArea).Value2
    $rollCost = [double]$main.Cells($rRoll, $mCost).Value2
    Check ($rollCost -gt 0) "roll Paper Cost calculated ($rollCost)"

    $rSheet = New-Row $main $loM
    Set-Cell $main $rSheet $mStk 'Gloss 200gsm SRA3 sheet'
    Set-Cell $main $rSheet $mQty 40
    $sheetId = [string]$main.Cells($rSheet, $mId).Value2
    $sheetCost = [double]$main.Cells($rSheet, $mCost).Value2

    # --------------------------------------------- metres file -> cm room
    Write-Host ''
    Write-Host '=== Metres file imported into the Centimetres room ==='
    [void]$main.Activate()
    [void]$xl.Run('btnExport')
    $csvM = Newest-Csv 'MAIN'
    Check ($null -ne $csvM) 'Main export written'
    Import-Csv-Into $annexe $csvM

    $aQty = Cc $loA 'Qty'; $aUnit = Cc $loA 'Unit'; $aArea = Cc $loA 'Area m2'
    $aCost = Cc $loA 'Paper Cost'; $aId = Cc $loA 'Job ID'; $aStk = Cc $loA 'Paper Stock'

    $ar = Find-Row $annexe $loA $rollId
    Check ($ar -gt 0) "roll job $rollId landed in Annexe"
    if ($ar -gt 0) {
        Check (Near $annexe.Cells($ar, $aQty).Value2 250) "roll Qty 2.5 m became 250 cm (got $($annexe.Cells($ar, $aQty).Value2))"
        Check ($annexe.Cells($ar, $aUnit).Value2 -eq 'cm') 'Unit reads "cm"'
        Check (Near $annexe.Cells($ar, $aArea).Value2 $rollArea) 'Area m2 unchanged'
        Check (Near $annexe.Cells($ar, $aCost).Value2 $rollCost) 'Paper Cost unchanged'
    }
    $as = Find-Row $annexe $loA $sheetId
    Check ($as -gt 0) "sheet job $sheetId landed in Annexe"
    if ($as -gt 0) {
        Check (Near $annexe.Cells($as, $aQty).Value2 40) 'sheet Qty not converted (40)'
        Check ($annexe.Cells($as, $aUnit).Value2 -eq 'sheets') 'sheet Unit "sheets"'
        Check (Near $annexe.Cells($as, $aCost).Value2 $sheetCost) 'sheet Paper Cost unchanged'
    }

    # --------------------------------------------- cm file -> metres room
    Write-Host ''
    Write-Host '=== Centimetres file imported into the Metres room ==='
    $rCm = New-Row $annexe $loA
    Set-Cell $annexe $rCm $aStk 'Canvas 914mm roll'
    Set-Cell $annexe $rCm $aQty 150
    $cmId = [string]$annexe.Cells($rCm, $aId).Value2
    $cmArea = [double]$annexe.Cells($rCm, $aArea).Value2
    $cmCost = [double]$annexe.Cells($rCm, $aCost).Value2
    Check ($annexe.Cells($rCm, $aUnit).Value2 -eq 'cm') 'Annexe row typed in cm reads "cm"'
    Check ($cmCost -gt 0) "cm job Paper Cost calculated ($cmCost)"
    [void]$annexe.Activate()
    [void]$xl.Run('btnExport')
    $csvA = Newest-Csv 'ANNEX'
    Check ($null -ne $csvA) 'Annexe export written'
    Import-Csv-Into $main $csvA

    $mr = Find-Row $main $loM $cmId
    Check ($mr -gt 0) "cm job $cmId landed in Main"
    if ($mr -gt 0) {
        Check (Near $main.Cells($mr, $mQty).Value2 1.5) "roll Qty 150 cm became 1.5 m (got $($main.Cells($mr, $mQty).Value2))"
        Check ($main.Cells($mr, $mUnit).Value2 -eq 'metres') 'Unit reads "metres"'
        Check (Near $main.Cells($mr, $mArea).Value2 $cmArea) 'Area m2 unchanged'
        Check (Near $main.Cells($mr, $mCost).Value2 $cmCost) 'Paper Cost unchanged'
    }
    # The cm file also carried Main's original jobs (held in Annexe at 250 cm);
    # overwriting them must land back at the original metres, not 250 or 0.025.
    $rr = Find-Row $main $loM $rollId
    Check ($rr -gt 0 -and (Near $main.Cells($rr, $mQty).Value2 2.5)) "original 2.5 m roll round-trips through cm back to 2.5"
    if ($rr -gt 0) { Check (Near $main.Cells($rr, $mCost).Value2 $rollCost) 'round-tripped Paper Cost unchanged' }
    $rs = Find-Row $main $loM $sheetId
    Check ($rs -gt 0 -and (Near $main.Cells($rs, $mQty).Value2 40)) 'sheet job still 40 after the round trip'

    # ----------------------------------------------- same-unit control
    Write-Host ''
    Write-Host '=== Same unit: Metres file back into the Metres room ==='
    Import-Csv-Into $main $csvM
    $rr = Find-Row $main $loM $rollId
    Check ($rr -gt 0 -and (Near $main.Cells($rr, $mQty).Value2 2.5)) 'metres into metres: Qty untouched (2.5)'
    if ($rr -gt 0) { Check ($main.Cells($rr, $mUnit).Value2 -eq 'metres') 'Unit still "metres"' }

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
