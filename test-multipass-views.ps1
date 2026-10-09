# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Prerequisite: multi-pass costing (0.11.0). The reporting and view side
# (docs/multipass-costing-design.md decisions 12, 17, 19, 22, 28, 32; column-view-
# presets-design.md):
#
#   1. Reports: Passes after Printer, Set-up cost after Consumable cost; the job's
#      Passes and Set-up in its row; the Pass type filter (F11) and Passes as a sort
#      choice; the Matching totals still reconcile; passes never listed.
#   2. Summary: Passes (C) and Set-up cost (L) per key, totals unchanged in meaning.
#   3. Column toggles: the pass-column toggle layers under All / Reduced / Minimal
#      ("All" shows everything except a group toggled off); Set-up Cost follows the
#      cost toggle.
#   4. Add pass / Remove pass grey unless a multi-pass job or one of its passes is
#      selected.
#   5. Location sheets: no AutoFilter arrows, sorting and filtering withheld.
#   6. The job planner does not offer multi-pass printers.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$script:anyFail = $false
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $main = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    $lo = $main.ListObjects('tblJobs_MAIN')
    function J([int]$rowNo, [string]$h) { return (Invoke-ComRetry -RetryOnNull { $lo.ListRows($rowNo).Range.Cells(1, (Col $lo $h)) }) }
    function W([int]$rowNo, [string]$h, $v) { [void]$xl.Run('UnlockSheet', $main); if ($v -is [int] -or $v -is [double]) { $v = ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }; $cell = J $rowNo $h; Invoke-ComRetry { $cell.Value2 = $v } | Out-Null }
    function Pick([int]$rowNo) { [void]$main.Activate(); [void]((J $rowNo 'Date/Time').Select()) }
    function Near([double]$a, [double]$b) { return ([Math]::Abs($a - $b) -lt 0.001) }
    function ColHidden([string]$h) { return [bool]$lo.ListColumns($h).Range.EntireColumn.Hidden }

    # ---- setup -------------------------------------------------------------------
    $cws = $wb.Worksheets('Consumables'); $clo = $cws.ListObjects('tblColours'); [void]$cws.Activate()
    foreach ($row in @(@('Black', 0.30), @('Blue', 0.40))) {
        [void]$xl.Run('AddCatalogRow', 'tblColours'); $n = $clo.ListRows.Count; [void]$xl.Run('UnlockSheet', $cws)
        $r = $clo.ListRows($n).Range
        $r.Cells(1, (Col $clo 'Colour')).Value2 = $row[0]; $r.Cells(1, (Col $clo 'Consumable type')).Value2 = 'Risograph'; $r.Cells(1, (Col $clo 'Cost per m2')).Value2 = [double]$row[1]
    }
    $pws = $wb.Worksheets('Printers'); $plo = $pws.ListObjects('tblPrinters'); [void]$pws.Activate()
    [void]$xl.Run('AddCatalogRow', 'tblPrinters'); $pr = $plo.ListRows($plo.ListRows.Count).Range; [void]$xl.Run('UnlockSheet', $pws)
    $pr.Cells(1, (Col $plo 'Model')).Value2 = 'Test RISO'; $pr.Cells(1, (Col $plo 'Consumable type')).Value2 = 'Risograph'
    $pr.Cells(1, (Col $plo 'Colour mode')).Value2 = 'multi-pass'; $pr.Cells(1, (Col $plo 'Template cost')).Value2 = [double]3
    $pr.Cells(1, (Col $plo 'Max sheet size')).Value2 = 'A1'; $pr.Cells(1, (Col $plo 'Active')).Value2 = 'Yes'
    [void]$xl.Run('UnlockSheet', $main)
    $loc = $main.Names.Item('LOC_Printers').RefersToRange; $loc.Value2 = ([string]$loc.Value2 + ';Test RISO')
    $sample = $lo.ListRows.Count

    [void]$main.Activate(); [void]$xl.Run('AddPrintJob', $main); $job = $lo.ListRows.Count
    W $job 'Student Name' 'Views Tester'; W $job 'Technician' ([string](J 1 'Technician').Value2)
    W $job 'Printer' 'Test RISO'; W $job 'Paper Stock' 'Gloss 200gsm SRA3 sheet'; W $job 'Qty' 100
    foreach ($c in 'Black', 'Blue') { Pick $job; [void]$xl.Run('AddPass', $main); W $lo.ListRows.Count 'Colour' $c }
    $jobId = [string](J $job 'Job ID').Value2
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Check ((Near ([double](J $job 'Set-up Cost').Value2) 6.0) -and (Near ([double](J $job 'Consumable Cost').Value2) 10.08)) 'setup: job has set-up 6.00 and ink 4.32 + 5.76'

    Write-Host '=== Reports ==='
    $rep = $wb.Worksheets('Reports'); [void]$rep.Activate(); $xl.Run('SetQuiet', $true)
    $hdr = @(); for ($i = 1; $i -le 20; $i++) { $hdr += [string]$rep.Cells(10, 7 + $i).Text }
    Write-Host ('  ' + ($hdr -join ' | '))
    Check (($hdr.IndexOf('Passes') -eq $hdr.IndexOf('Printer') + 1) -and ($hdr.IndexOf('Set-up cost') -eq $hdr.IndexOf('Consumable cost') + 1)) 'Passes follows Printer; Set-up cost follows Consumable cost'
    Check ($hdr.IndexOf('Paper stock') -eq 6 -and $hdr.IndexOf('Chargeable') -eq 15 -and $hdr.IndexOf('Paid') -eq 16) 'Paper stock, Chargeable and Paid are where the constants say (7, 16, 17)'
    Start-Sleep -Seconds 2
    # A spill range read once into a 2-D array (no per-cell COM calls), retried while Excel is busy.
    function SpillVals($anchor) { for ($a = 1; $a -le 4; $a++) { try { $v = $anchor.SpillingToRange.Value2; if ($v -is [array]) { return ,$v } } catch { Start-Sleep -Seconds ($a * 2) } }; throw 'spill range could not be read' }
    $sv = SpillVals $rep.Cells(11, 8)
    $total = [int]$rep.Range('I9').Value2
    Check ($total -eq $sample + 1) "Matching jobs counts jobs only ($total of $($sample + 1))"
    Check ($sv.GetLength(0) -eq $sample + 1) 'the results list has one row per job: no pass rows'
    $pc = $hdr.IndexOf('Passes') + 1; $sc = $hdr.IndexOf('Set-up cost') + 1; $jc = $hdr.IndexOf('Student name') + 1
    $found = 0
    for ($r = 1; $r -le $sv.GetLength(0); $r++) { if ([string]$sv[$r, $jc] -eq 'Views Tester') { $found = $r } }
    Check ($found -gt 0 -and [int]$sv[$found, $pc] -eq 2 -and (Near ([double]$sv[$found, $sc]) 6.0)) 'the multi-pass job shows Passes 2 and Set-up cost 6.00'
    $charge = [double]$rep.Range('N9').Value2
    Check ($charge -gt 0) "Matching chargeable (N9) reads a figure ($charge)"

    $rep.Range('F11').Value2 = 'Multi-pass'; Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Check ([int]$rep.Range('I9').Value2 -eq 1) 'Pass type = Multi-pass finds the one multi-pass job'
    $rep.Range('F11').Value2 = 'Single pass'; Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Check ([int]$rep.Range('I9').Value2 -eq $sample) 'Pass type = Single pass finds the rest'
    Check ([bool]$xl.Run('HasActiveFilter', $rep)) 'a set Pass type counts as an active filter'
    $rep.Range('F11').Value2 = ''; Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Check ([int]$rep.Range('I9').Value2 -eq $sample + 1) 'blank Pass type = everything'
    $rep.Range('I7').Value2 = 'Passes'; $rep.Range('K7').Value2 = 'Descending'; Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Start-Sleep -Seconds 1
    $sv = SpillVals $rep.Cells(11, 8)
    Check ([int]$sv[1, $pc] -eq 2) 'sorting by Passes (descending) puts the multi-pass job first'
    $rep.Range('I7').Value2 = ''; $rep.Range('K7').Value2 = ''

    Write-Host ''
    Write-Host '=== Summary ==='
    $sum = $wb.Worksheets('Summary'); [void]$sum.Activate()
    $sh = @(); for ($i = 1; $i -le 15; $i++) { $sh += [string]$sum.Cells($SUM_HDR_ROW, $i).Text }
    Check ($sh[2] -eq 'Passes' -and $sh[11] -eq 'Set-up cost' -and $sh[14] -eq 'Chargeable') "Summary headers: Passes in C, Set-up cost in L ($($sh -join ', '))"
    Start-Sleep -Seconds 1
    $sp = SpillVals $sum.Cells($SUM_FIRST_ROW, 1)
    $row = 0; for ($r = 1; $r -le $sp.GetLength(0); $r++) { if ([string]$sp[$r, 2] -eq 'Test RISO') { $row = $r } }
    Check ($row -gt 0 -and [int]$sp[$row, 3] -eq 2 -and (Near ([double]$sp[$row, 12]) 6.0)) 'the Test RISO row shows Passes 2 and Set-up cost 6.00'
    $gross = 0.0; for ($r = 1; $r -le $sp.GetLength(0); $r++) { $gross += [double]$sp[$r, 13] }
    Check ((Near $gross ([double]$sum.Range("D$SUM_TOT_ROW").Value2))) 'the table''s Gross column adds up to the Gross total'

    Write-Host ''
    Write-Host '=== Column toggles layer under the view ==='
    $pass = 'Passes', 'Colour', 'Row Type', 'Parent', 'Pass'
    [void]$main.Activate()
    function AllHidden($names) { foreach ($n in $names) { if (-not (ColHidden $n)) { return $false } }; return $true }
    function NoneHidden($names) { foreach ($n in $names) { if (ColHidden $n) { return $false } }; return $true }
    [void]$xl.Run('SetViewMode', 'All')
    Check (NoneHidden $pass) 'view All, toggle on: the pass columns show'
    [void]$xl.Run('ToggleMultiPassColumns')
    Check (AllHidden $pass) 'toggle off: they hide'
    [void]$xl.Run('SetViewMode', 'All')
    Check (AllHidden $pass) '...and choosing All does not switch the group back on'
    [void]$xl.Run('ToggleMultiPassColumns')
    Check (NoneHidden $pass) 'toggle on again: they show (view All)'
    [void]$xl.Run('SetViewMode', 'Reduced')
    Check (AllHidden $pass) 'Reduced hides them even with the toggle on'
    [void]$xl.Run('SetViewMode', 'All')
    Check (NoneHidden $pass) 'back to All shows them'
    Check (-not (ColHidden 'Set-up Cost')) 'Set-up Cost is visible'
    [void]$xl.Run('ToggleCostColumns')
    Check ((ColHidden 'Set-up Cost') -and (ColHidden 'Consumable Cost')) 'the cost toggle hides Set-up Cost with the other cost detail'
    Check (NoneHidden $pass) '...without touching the pass columns'
    [void]$xl.Run('ToggleCostColumns')

    Write-Host ''
    Write-Host '=== Add pass / Remove pass grey with the selection ==='
    function GreyOf([string]$prefix) {
        foreach ($b in $main.Buttons()) { if ($b.Name.StartsWith($prefix)) { return ([int]$b.Characters().Font.Color -eq 9868950) } }
        throw "button $prefix not found"
    }
    Pick 1; [void]$xl.Run('RefreshPassButtons', $main)
    Check ((GreyOf 'pcb_btnAddPass') -and (GreyOf 'pcb_btnRemovePass')) 'a single-pass job selected: Add pass and Remove pass are grey'
    Check (-not (GreyOf 'pcb_btnTogglePasses')) '...Toggle passes is not (a multi-pass printer exists)'
    Pick $job; [void]$xl.Run('RefreshPassButtons', $main)
    Check ((-not (GreyOf 'pcb_btnAddPass')) -and (-not (GreyOf 'pcb_btnRemovePass'))) 'the multi-pass job selected: both are live'
    Pick ($job + 1); [void]$xl.Run('RefreshPassButtons', $main)
    Check (-not (GreyOf 'pcb_btnAddPass')) 'one of its passes selected: still live'
    Check ($null -ne ($main.Buttons() | Where-Object { $_.Name.StartsWith('pcb_btnToggleMultiPass') })) 'the pass-column toggle button is drawn'

    Write-Host ''
    Write-Host '=== No sort or filter on a location sheet ==='
    Check (-not [bool]$lo.ShowAutoFilter) 'the AutoFilter arrows are hidden'
    Check ((-not [bool]$main.Protection.AllowSorting) -and (-not [bool]$main.Protection.AllowFiltering)) 'protection withholds sorting and filtering'
    Check ([bool]$rep.Protection.AllowFiltering) 'Reports keeps its filtering'

    Write-Host ''
    Write-Host '=== The planner does not offer multi-pass printers ==='
    [void]$xl.Run('RefreshSummaryPlanner')
    $v = $sum.Range('B7').Validation.Formula1
    $items = @(); $listRange = $xl.Evaluate($v.TrimStart('=')); foreach ($c in $listRange.Cells) { $items += [string]$c.Value2 }
    Check (($items | Where-Object { [string]$_ -eq 'Test RISO' }).Count -eq 0 -and ($items | Where-Object { [string]$_ -like '*Xerox*' }).Count -gt 0) "Test RISO is not in the planner's printer list (offered: $($items -join ', '))"
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($script:anyFail) { Write-Host 'FAIL'; exit 1 }
Write-Host 'multi-pass views checks complete'
