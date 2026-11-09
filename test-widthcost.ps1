# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# 0.12.2: print width scales paper cost.
#
#   - A roll job's Paper Cost is Qty x unit cost x (Print Width mm / stock
#     width); blank or full width costs the whole roll.
#   - Rows from before 0.12.2 (S_WidthRule blank) keep the full-width cost,
#     and neither StampRow nor re-picking the paper moves them.
#   - New jobs (Add, Repeat, the planner) are marked 'scaled'.
#   - The planner estimate equals the row it adds.
#
# Mac port: mac/test-widthcost.py. Drives a COPY in %TEMP%, never
# src\PrintJob.xlsm itself. Closes WITHOUT saving.

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
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Example Print Room')
    $lo = $main.ListObjects('tblJobs_EXAMPLE')
    $sum = $wb.Worksheets('Summary')
    $dp = 2
    try { $dp = [int]$wb.Names.Item('SET_ROUND_DP').RefersToRange.Value2 } catch {}

    function Cell($n, $h) { return $lo.ListRows($n).Range.Cells(1, (Col $lo $h)) }
    function Txt($n, $h) { return [string](Cell $n $h).Text }
    function Num($n, $h) { $v = (Cell $n $h).Value2; if ($null -eq $v -or "$v" -eq '') { return 0.0 }; return [double]$v }
    function Rnd([double]$x) { return [Math]::Round($x, $dp, [MidpointRounding]::AwayFromZero) }

    Write-Host '=== schema and column ==='
    $schema = [string]$wb.Names.Item('SET_SCHEMA').RefersToRange.Text
    Check ($schema -eq '1.5') "SET_SCHEMA reads 1.5 (got '$schema')"
    $wr = Col $lo 'S_WidthRule'
    Check ($wr -gt 0) 'S_WidthRule column exists'
    Check (($wr -eq (Col $lo 'S_SchemaVer') - 1) -and ((Col $lo 'S_SchemaVer') -eq $lo.ListColumns.Count)) 'S_WidthRule sits just before S_SchemaVer, which is still last'

    Write-Host ''
    Write-Host '=== rows from before 0.12.2 keep the full-width paper cost ==='
    $oldNarrow = 0; $rollRow = 0; $narrowOld = 0
    for ($r = 1; $r -le $lo.ListRows.Count; $r++) {
        if ((Txt $r 'Row Type') -eq 'Pass' -or (Txt $r 'Qty') -eq '') { continue }
        Check ((Txt $r 'S_WidthRule') -eq '') "row $r has no S_WidthRule"
        if ((Txt $r 'S_Measure') -eq 'Sheet') { continue }
        if ($rollRow -eq 0) { $rollRow = $r }
        $pw = Num $r 'Print Width mm'; $sw = Num $r 'S_StockWidth_mm'
        if ($pw -gt 0 -and $pw -lt $sw) {
            $oldNarrow++
            if ($narrowOld -eq 0) { $narrowOld = $r }
            $want = Rnd ((Num $r 'Qty') * (Num $r 'S_UnitCost'))
            Check ([Math]::Abs((Num $r 'Paper Cost') - $want) -lt 0.005) "row ${r}: narrower width ($pw of $sw mm) still costs the whole roll ($(Num $r 'Paper Cost') = $want)"
        }
    }
    Check ($oldNarrow -gt 0) "the sample rows include a roll job with a narrower width ($oldNarrow)"
    Check ($rollRow -gt 0) 'the sample rows include a roll job'

    $prn = Txt $rollRow 'Printer'; $stk = Txt $rollRow 'Paper Stock'
    $stockW = Num $rollRow 'S_StockWidth_mm'

    Write-Host ''
    Write-Host '=== a new roll job pays for the width it uses ==='
    $main.Activate()
    [void]$xl.Run('btnAddPrintJob')
    $new = $lo.ListRows.Count
    (Cell $new 'Printer').Value2 = $prn
    (Cell $new 'Paper Stock').Value2 = $stk
    (Cell $new 'Qty').Value2 = 10
    Start-Sleep -Milliseconds 300
    $unitCost = Num $new 'S_UnitCost'
    Check ($unitCost -gt 0) "unit cost stamped ($unitCost)"
    Check ((Txt $new 'S_WidthRule') -eq 'scaled') "new row is marked 'scaled' (got '$(Txt $new 'S_WidthRule')')"
    $full = Rnd (10 * $unitCost)
    Check ([Math]::Abs((Num $new 'Paper Cost') - $full) -lt 0.005) "no width: full roll ($(Num $new 'Paper Cost') = $full)"

    (Cell $new 'Print Width mm').Value2 = $stockW / 2
    Start-Sleep -Milliseconds 300
    $half = Rnd (10 * $unitCost * 0.5)
    Check ([Math]::Abs((Num $new 'Paper Cost') - $half) -lt 0.005) "half the width: half the paper ($(Num $new 'Paper Cost') = $half)"
    Check ([Math]::Abs((Num $new 'Gross Cost') - ((Num $new 'Paper Cost') + (Num $new 'Consumable Cost') + (Num $new 'Set-up Cost'))) -lt 0.005) 'Gross Cost follows the scaled paper cost'

    (Cell $new 'Print Width mm').Value2 = $stockW
    Start-Sleep -Milliseconds 300
    Check ([Math]::Abs((Num $new 'Paper Cost') - $full) -lt 0.005) "full width: full roll ($(Num $new 'Paper Cost') = $full)"

    Write-Host ''
    Write-Host '=== re-stamping never changes which rule a row is on ==='
    (Cell $new 'Print Width mm').Value2 = $stockW / 2
    Start-Sleep -Milliseconds 300
    [void]$xl.Run('StampRow', $main, $new)
    Check ((Txt $new 'S_WidthRule') -eq 'scaled') 'a scaled row stays scaled after StampRow'
    Check ([Math]::Abs((Num $new 'Paper Cost') - $half) -lt 0.005) 'and keeps its cost'
    $before = Num $narrowOld 'Paper Cost'
    [void]$xl.Run('StampRow', $main, $narrowOld)
    (Cell $narrowOld 'Paper Stock').Value2 = (Txt $narrowOld 'Paper Stock')
    Start-Sleep -Milliseconds 300
    Check ((Txt $narrowOld 'S_WidthRule') -eq '') 'an old row stays on the old rule after StampRow and re-picking its paper'
    Check ([Math]::Abs((Num $narrowOld 'Paper Cost') - $before) -lt 0.005) "and keeps its cost ($before)"

    Write-Host ''
    Write-Host '=== Repeat Job gives a new row the new rule ==='
    (Cell $narrowOld 'Qty').Select() | Out-Null
    [void]$xl.Run('btnRepeatJob')
    Start-Sleep -Milliseconds 500
    $rep = $lo.ListRows.Count
    Check ((Txt $rep 'S_WidthRule') -eq 'scaled') 'a repeated old job is a new job: scaled'
    Check ([Math]::Abs((Num $rep 'Paper Cost') - $before * (Num $rep 'Print Width mm') / (Num $rep 'S_StockWidth_mm')) -lt 0.01) "and costs the width it uses ($(Num $rep 'Paper Cost'))"

    Write-Host ''
    Write-Host '=== the job planner estimates the same ==='
    $sum.Range('B7').Value2 = $prn
    $sum.Range('B8').Value2 = $stk
    $sum.Range('B9').Value2 = 10
    $sum.Range('B10').Value2 = $stockW / 2
    [void]$xl.Run('PlannerRecalc', $sum)
    Check ([Math]::Abs([double]$sum.Range('B14').Value2 - $half) -lt 0.005) "planner paper cost for half width ($($sum.Range('B14').Value2) = $half)"
    $sum.Range('B19').Value2 = [string]$xl.Run('LocValue', $main, 'LOC_Name')
    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 500
    $added = $lo.ListRows.Count
    Check ((Txt $added 'S_WidthRule') -eq 'scaled') 'a planned job is scaled'
    Check ([Math]::Abs((Num $added 'Paper Cost') - [double]$sum.Range('B14').Value2) -lt 0.005) "the added row's Paper Cost equals the planner's ($(Num $added 'Paper Cost'))"

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
