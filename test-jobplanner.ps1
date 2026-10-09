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
    $lo = $ws.ListObjects('tblJobs_MAIN')

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
