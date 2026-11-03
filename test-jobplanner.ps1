# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# The job planner box at the top of the Summary sheet (modPlanner, 0.10.29).
#
#   1. layout: the box, its inputs and the shifted Summary table
#   2. dropdowns: printers/papers/rooms are offered, rooms from every print room
#   3. incomplete or conflicting input gives a message and no estimate
#   4. THE estimate equals what the job row computes: plan a job, add it to a
#      room, and the new row's Gross Cost must equal the planner's Overall
#   5. the new row is prefilled (Printer, Paper Stock, Qty) and stamped
#   6. nothing is added when the plan is incomplete or no room is chosen
#   7. a multi-pass printer (0.11.7): row 10 becomes five colour dropdowns, the
#      estimate prices each pass and its set-up, and the job is added with its
#      passes; a single-pass printer gets row 10 back
#
# Drives a COPY of src\PrintJob.xlsm in %TEMP%, closes without saving. Needs a
# second room for (2), so uses the shared Annexe fixture.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')

function Get-ListText($cell) {
    try { return [string]$cell.Validation.Formula1 } catch { return '' }
}
# The items behind a list validation that points at a staged range on _Work.
function Get-ListItems($wb, $cell) {
    $f = (Get-ListText $cell).TrimStart('=')
    if ($f -eq '') { return @() }
    $rng = $wb.Application.Range($f)
    $out = @()
    foreach ($c in $rng.Cells) { if ([string]$c.Text -ne '') { $out += [string]$c.Text } }
    return $out
}

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
    $wb = Invoke-ComRetry -Attempts 5 { $xl.Workbooks.Open($f) }
    [void](Add-AnnexeFixture $xl $wb)
    $xl.Run('SetQuiet', $true)
    $sum = Invoke-ComRetry -Attempts 5 { $wb.Worksheets('Summary') }
    $ws = Invoke-ComRetry -Attempts 5 { $wb.Worksheets('Example Print Room') }
    $lo = $ws.ListObjects('tblJobs_EXAMPLE')

    Write-Host '=== layout ==='
    Check ([string]$sum.Range('A5').Text -eq 'Plan a print job') 'box title in A5'
    Check ([string]$sum.Range('A7').Text -eq 'Printer' -and [string]$sum.Range('A8').Text -eq 'Paper') 'Printer and Paper labels'
    Check ([string]$sum.Range('A23').Text -eq 'Totals') 'Totals label moved below the box (A23)'
    Check ([string]$sum.Cells($SUM_HDR_ROW, 1).Text -eq 'Location') "Summary table header is on row $SUM_HDR_ROW"
    $caps = @(); for ($i = 1; $i -le $sum.Buttons().Count; $i++) { $caps += [string]$sum.Buttons($i).Caption }
    Check ($caps -contains 'Add to print room') 'Add to print room button exists'
    Check ($caps -contains 'Refresh Locations') 'existing Summary buttons are still there'

    Write-Host '=== dropdowns ==='
    $prnItems = Get-ListItems $wb $sum.Range('B7')
    $stkItems = Get-ListItems $wb $sum.Range('B8')
    $roomItems = Get-ListItems $wb $sum.Range('B19')
    Check ($prnItems.Count -ge 1) "printers offered ($($prnItems.Count))"
    Check ($stkItems.Count -ge 1) "papers offered ($($stkItems.Count))"
    Check ($roomItems -contains 'Annexe') 'room list includes the second room (Annexe)'
    Check ($roomItems.Count -ge 2) "room list has every room ($($roomItems.Count))"

    Write-Host '=== incomplete input ==='
    $sum.Range('B7:B11').ClearContents() | Out-Null
    $xl.Run('PlannerRecalc', $sum)
    Check ([string]$sum.Range('A17').Text -eq 'Choose a printer.') "no printer -> prompt (got '$($sum.Range('A17').Text)')"
    Check ([string]$sum.Range('B16').Text -eq '') 'no estimate without inputs'

    # Use the printer and paper of the shipped sample job: known to be a valid pair.
    $prn = [string]$lo.ListRows(1).Range.Cells(1, (Col $lo 'Printer')).Value2
    $stk = [string]$lo.ListRows(1).Range.Cells(1, (Col $lo 'Paper Stock')).Value2
    Check ($prn -ne '' -and $stk -ne '') "sample job gives a printer/paper pair ($prn / $stk)"

    $sum.Range('B7').Value2 = $prn
    $xl.Run('PlannerRecalc', $sum)
    Check ([string]$sum.Range('A17').Text -eq 'Choose a paper.') 'printer only -> asks for paper'
    $sum.Range('B8').Value2 = $stk
    $xl.Run('PlannerRecalc', $sum)
    Check ([string]$sum.Range('B16').Text -eq '') 'no amount -> no estimate'

    Write-Host '=== estimate equals the job row ==='
    $isSheet = ([string]$sum.Range('A9').Text -eq 'Number of sheets')
    $sum.Range('B9').Value2 = 7
    $xl.Run('PlannerRecalc', $sum)
    $total = [double]$sum.Range('B16').Value2
    $paper = [double]$sum.Range('B14').Value2
    $ink = [double]$sum.Range('B15').Value2
    Check ($total -gt 0) "an estimate is produced ($total)"
    Check ([Math]::Abs($paper + $ink - $total) -lt 0.0001) "paper + ink = overall ($paper + $ink)"

    $roomName = [string]$xl.Run('LocValue', $ws, 'LOC_Name')
    $sum.Range('B19').Value2 = $roomName
    $before = $lo.ListRows.Count
    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 500
    $after = $lo.ListRows.Count
    Check ($after -eq $before + 1) "a row was added ($before -> $after)"

    $n = $after
    $gross = [double]$lo.ListRows($n).Range.Cells(1, (Col $lo 'Gross Cost')).Value2
    Check ([Math]::Abs($gross - $total) -lt 0.005) "new row Gross Cost ($gross) equals the planner's Overall ($total)"

    Write-Host '=== the added row is prefilled and stamped ==='
    Check ([string]$lo.ListRows($n).Range.Cells(1, (Col $lo 'Printer')).Value2 -eq $prn) 'Printer filled in'
    Check ([string]$lo.ListRows($n).Range.Cells(1, (Col $lo 'Paper Stock')).Value2 -eq $stk) 'Paper Stock filled in'
    $unit = [string]$xl.Run('RollUnitOf', $ws)
    $wantQty = 7
    if (-not $isSheet -and $unit -eq 'Centimetres') { $wantQty = 700 }
    Check ([double]$lo.ListRows($n).Range.Cells(1, (Col $lo 'Qty')).Value2 -eq $wantQty) "Qty is $wantQty (room unit $unit)"
    Check ([string]$lo.ListRows($n).Range.Cells(1, (Col $lo 'Paid')).Value2 -eq 'No') 'Paid is No'
    Check ([string]$lo.ListRows($n).Range.Cells(1, (Col $lo 'Job ID')).Value2 -ne '') 'Job ID allocated'
    Check ([string]$lo.ListRows($n).Range.Cells(1, (Col $lo 'S_StampedAt')).Value2 -ne '') 'row is stamped (S_StampedAt)'

    Write-Host '=== nothing added when it should not be ==='
    $sum.Range('B19').ClearContents() | Out-Null
    $before = $lo.ListRows.Count
    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 300
    Check ($lo.ListRows.Count -eq $before) 'no room chosen -> no row added'

    $sum.Range('B19').Value2 = $roomName
    $sum.Range('B9').ClearContents() | Out-Null
    $xl.Run('PlannerRecalc', $sum)
    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 300
    Check ($lo.ListRows.Count -eq $before) 'incomplete plan -> no row added'

    Write-Host '=== a multi-pass printer ==='
    # Added here, not with the Annexe: everything above is the single-pass box.
    $cws = $wb.Worksheets('Consumables'); $clo = $cws.ListObjects('tblColours'); [void]$cws.Activate()
    foreach ($row in @(@('Black', 'Risograph', 0.30), @('Blue', 'Risograph', 0.40), @('Cyan', 'Ink', 0.10))) {
        [void]$xl.Run('AddCatalogRow', 'tblColours'); [void]$xl.Run('UnlockSheet', $cws)
        $r = $clo.ListRows($clo.ListRows.Count).Range
        $r.Cells(1, (Col $clo 'Colour')).Value2 = $row[0]; $r.Cells(1, (Col $clo 'Consumable type')).Value2 = $row[1]; $r.Cells(1, (Col $clo 'Cost per m2')).Value2 = [double]$row[2]
    }
    $pws = $wb.Worksheets('Printers'); $plo = $pws.ListObjects('tblPrinters'); [void]$pws.Activate()
    [void]$xl.Run('AddCatalogRow', 'tblPrinters'); [void]$xl.Run('UnlockSheet', $pws)
    $pr = $plo.ListRows($plo.ListRows.Count).Range
    $pr.Cells(1, (Col $plo 'Model')).Value2 = 'Test RISO'; $pr.Cells(1, (Col $plo 'Consumable type')).Value2 = 'Risograph'
    $pr.Cells(1, (Col $plo 'Colour mode')).Value2 = 'multi-pass'; $pr.Cells(1, (Col $plo 'Template cost')).Value2 = [double]3
    $pr.Cells(1, (Col $plo 'Max sheet size')).Value2 = 'A1'; $pr.Cells(1, (Col $plo 'Active')).Value2 = 'Yes'
    [void]$xl.Run('UnlockSheet', $ws)
    $loc = $ws.Names.Item('LOC_Printers').RefersToRange; $loc.Value2 = ([string]$loc.Value2 + ';Test RISO')
    [void]$xl.Run('RefreshSummaryPlanner')
    [void]$sum.Activate()

    function T([string]$a1) { return [string]$sum.Range($a1).Text }
    Check ((Get-ListItems $wb $sum.Range('B7')) -contains 'Test RISO') 'the multi-pass printer is offered'
    $sum.Range('B7:B11').ClearContents() | Out-Null
    $sum.Range('B7').Value2 = 'Test RISO'
    $sum.Range('B8').Value2 = 'Gloss 200gsm SRA3 sheet'
    $sum.Range('B9').Value2 = 100
    $xl.Run('PlannerRecalc', $sum)
    Check ((T 'A10') -eq 'Colours (one per pass)') "row 10 asks for colours (got '$(T 'A10')')"
    Check ((T 'A15') -eq 'Ink and set-up') 'the ink line includes set-up'
    $bItems = @(Get-ListItems $wb $sum.Range('B10')); $fItems = @(Get-ListItems $wb $sum.Range('F10'))
    Check (($bItems -join ',') -eq 'Black,Blue') "first slot lists this printer's colours (got $($bItems -join ','))"
    Check (($fItems -join ',') -eq 'Black,Blue') "fifth slot lists the same colours (got $($fItems -join ','))"
    Check ((T 'C10') -eq '') 'the width hint is gone'
    Check ((T 'A17') -eq 'Choose at least one colour.') "no colour -> prompt (got '$(T 'A17')')"
    Check ((T 'B16') -eq '') 'no colour -> no estimate'

    $sum.Range('B19').Value2 = $roomName
    $before = $lo.ListRows.Count
    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 300
    Check ($lo.ListRows.Count -eq $before) 'no colour -> no row added'

    $sum.Range('B10').Value2 = 'Cyan'
    $xl.Run('PlannerRecalc', $sum)
    Check ((T 'A17') -eq "'Cyan' is not a colour Test RISO can use.") "another type's colour -> message (got '$(T 'A17')')"

    Write-Host '=== colours are priced one by one and can repeat ==='
    $sum.Range('B10').Value2 = 'Black'
    $sum.Range('D10').Value2 = 'Blue'          # a gap in the strip is not a pass
    $sum.Range('F10').Value2 = 'Black'
    $xl.Run('PlannerRecalc', $sum)
    $total = [double]$sum.Range('B16').Value2
    $paper = [double]$sum.Range('B14').Value2
    $inkSetup = [double]$sum.Range('B15').Value2
    Check ($total -gt 0) "an estimate is produced ($total)"
    Check ([Math]::Abs($paper + $inkSetup - $total) -lt 0.0001) "paper + ink and set-up = overall ($paper + $inkSetup)"
    Check ((T 'C15').EndsWith('set-up (3 passes)')) "the breakdown names the passes (got '$(T 'C15')')"

    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 500
    $after = $lo.ListRows.Count
    Check ($after -eq $before + 4) "a job and three passes were added ($before -> $after)"

    Write-Host '=== the added job and its passes ==='
    $job = $before + 1
    function JV([int]$rowNo, [string]$h) { return $lo.ListRows($rowNo).Range.Cells(1, (Col $lo $h)).Value2 }
    $gross = [double](JV $job 'Gross Cost'); $cons = [double](JV $job 'Consumable Cost'); $setup = [double](JV $job 'Set-up Cost')
    Check ([Math]::Abs($gross - $total) -lt 0.005) "the job's Gross Cost ($gross) equals the planner's Overall ($total)"
    Check ([Math]::Abs($cons + $setup - $inkSetup) -lt 0.005) "its ink and set-up ($cons + $setup) equal the planner's ($inkSetup)"
    Check ([Math]::Abs($setup - 9) -lt 0.005) "set-up is 3 passes at 3.00 ($setup)"
    Check ([int](JV $job 'Passes') -eq 3) 'Passes is 3'
    Check ([int]$lo.ListRows($job).Range.EntireRow.OutlineLevel -eq 1) 'the job row is outline level 1'
    $jobId = [string](JV $job 'Job ID')
    $i = 0
    foreach ($want in @(@('Black', 0.30), @('Blue', 0.40), @('Black', 0.30))) {
        $i++
        $p = $job + $i
        Check ([string](JV $p 'Row Type') -eq 'Pass' -and [string](JV $p 'Parent') -eq $jobId -and [int](JV $p 'Pass') -eq $i) "pass $i belongs to the job"
        Check ([string](JV $p 'Colour') -eq $want[0] -and [Math]::Abs([double](JV $p 'S_ConsRate') - $want[1]) -lt 0.0001) "pass $i is $($want[0]) at $($want[1])"
        Check ([int]$lo.ListRows($p).Range.EntireRow.OutlineLevel -eq 2) "pass $i is outline level 2"
    }

    [void]$sum.Activate()
    $sum.Range('D10').ClearContents() | Out-Null
    $sum.Range('F10').ClearContents() | Out-Null
    $xl.Run('PlannerRecalc', $sum)
    Check ((T 'C15').EndsWith('set-up (1 pass)')) "one colour reads as one pass (got '$(T 'C15')')"
    $before = $lo.ListRows.Count
    [void]$xl.Run('btnPlannerAdd')
    Start-Sleep -Milliseconds 500
    [void]$sum.Activate()
    Check ($lo.ListRows.Count -eq $before + 2) 'a second job with one pass was added'
    Check ([int]$lo.ListRows($before + 1).Range.EntireRow.OutlineLevel -eq 1) 'added below a pass, the job row is still level 1'

    Write-Host '=== row 10 goes back for a single-pass printer ==='
    $sum.Range('C10').Value2 = 'Blue'
    $sum.Range('B7').Value2 = $prn
    $sum.Range('B8').Value2 = $stk
    $sum.Range('B9').Value2 = 7
    $xl.Run('PlannerRecalc', $sum)
    Check ((T 'A10') -eq 'Print width mm') 'the Print width label is back'
    Check ((T 'A15') -eq 'Ink (consumable)') 'the ink label is back'
    Check ((T 'B10') -eq '') 'the colour in B10 is gone'
    Check ((T 'C10') -ne 'Blue' -and (T 'D10') -eq '' -and (T 'F10') -eq '') 'and the other slots'
    Check ((T 'C15') -eq '') 'no set-up breakdown'
    Check ([double]$sum.Range('B16').Value2 -gt 0) 'the estimate still works'
    $sum.Range('B7').Value2 = 'Test RISO'
    $xl.Run('PlannerRecalc', $sum)
    Check ((T 'B10') -eq '' -and (T 'C10') -eq '') 'back on the multi-pass printer, the strip starts empty'
    $sum.Range('B8').Value2 = 'Matte 180gsm 1067mm roll'
    $sum.Range('B10').Value2 = 'Black'
    $xl.Run('PlannerRecalc', $sum)
    Check ((T 'A17') -like '*cannot take*') "a roll paper on the multi-pass printer -> message (got '$(T 'A17')')"
    Check ((T 'B16') -eq '') 'and no estimate'

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
