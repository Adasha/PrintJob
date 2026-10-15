# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# University subsidy (0.10.31): one percentage on Settings, applied to the
# chargeable total on Reports and in the exported report.
#
#   1. The SUBSIDY_PCT Settings row, its SET_ name, percent format, 0-100%
#      validation, unlocked value cell, default 0; re-running setup neither
#      duplicates the row nor resets a value the user set.
#   2. Reports: the label and value sit directly above Unpaid (rows 6 and 7);
#      at 0% the value equals Chargeable; at 25% it is Chargeable x 0.75
#      rounded to the currency decimals, with the label naming the percentage;
#      Chargeable / Paid / Unpaid and the per-job amounts stay at the full figure;
#      the value follows the filters; blank, over 100% and negative read safely.
#   3. Export report: a "Billable after N% subsidy" line follows Still owed and
#      carries the same value; Total chargeable is still the full amount.
#
# Cell addresses come from the header text of the Matching row, not fixed
# letters, so this survives the Reports columns moving (they do in 0.11).
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
    $rep = Invoke-ComRetry { $wb.Worksheets('Reports') }
    $set = $wb.Worksheets('Settings')

    # Column (sheet column number) of a label in the Matching row, row 8.
    function MatchCol([string]$label) {
        for ($col = 8; $col -le 40; $col++) { if ([string]$rep.Cells(8, $col).Text -eq $label) { return $col } }
        return 0
    }
    function SetPct($v) {
        $r = $wb.Names.Item('SET_SUBSIDY_PCT').RefersToRange
        [void]$xl.Run('UnlockSheet', $set)
        if ($null -eq $v) { $r.ClearContents() | Out-Null } else { $r.Value2 = [double]$v }
        $xl.Calculate()
    }

    Write-Host '=== The Settings row ==='
    $lo = $set.ListObjects('tblSettings')
    $keyCol = Col $lo 'Key'
    $rows = @(); for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if ([string]$lo.DataBodyRange.Cells($i, $keyCol).Value2 -eq 'SUBSIDY_PCT') { $rows += $i } }
    Check ($rows.Count -eq 1) 'exactly one SUBSIDY_PCT row'
    $val = $wb.Names.Item('SET_SUBSIDY_PCT').RefersToRange
    Check ([double]$val.Value2 -eq 0) 'defaults to 0'
    Check ($val.NumberFormat -like '*%') "value cell is formatted as a percentage ($($val.NumberFormat))"
    Check ($val.Locked -eq $false) 'value cell is unlocked (editable)'
    Check ($val.Validation.Type -eq 2) 'value cell has decimal validation'
    $notes = [string]$lo.DataBodyRange.Cells($rows[0], (Col $lo 'Notes')).Value2
    Check ($notes -notlike 'Read-only*') 'row is not marked Read-only (so a restore carries it)'

    SetPct 0.25
    [void]$xl.Run('EnsureSubsidySetting'); [void]$xl.Run('EnsureSubsidySetting')
    $n = 0; for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if ([string]$lo.DataBodyRange.Cells($i, $keyCol).Value2 -eq 'SUBSIDY_PCT') { $n++ } }
    Check ($n -eq 1) 're-running setup does not duplicate the row'
    Check ([math]::Abs([double]$wb.Names.Item('SET_SUBSIDY_PCT').RefersToRange.Value2 - 0.25) -lt 1e-9) 're-running setup keeps a value the user set'
    SetPct 0

    Write-Host '=== Reports totals ==='
    $cCh = MatchCol 'Chargeable'; $cPd = MatchCol 'Paid'; $cUn = MatchCol 'Unpaid'
    Check ($cCh -gt 0 -and $cPd -gt 0 -and $cUn -gt 0) 'Chargeable / Paid / Unpaid totals found in the Matching row'
    Check (-not $rep.Columns($cUn).Hidden) 'the Unpaid column is visible (so the subsidy total above it is)'
    $chg = { [double]$rep.Cells(9, $cCh).Value2 }
    $bill = { [double]$rep.Cells(7, $cUn).Value2 }
    $lab = { [string]$rep.Cells(6, $cUn).Text }

    Check ((& $chg) -gt 0) "there is a chargeable total to test with ($(& $chg))"
    Check ([math]::Abs((& $bill) - (& $chg)) -lt 0.005) "0%: billable equals chargeable ($(& $bill))"
    Check ((& $lab) -eq 'After 0% subsidy') "0%: label reads '$(& $lab)'"

    SetPct 0.25
    $full = & $chg
    $expect = [math]::Round($full * 0.75, 2)
    Check ([math]::Abs((& $bill) - $expect) -lt 0.005) "25%: billable is chargeable x 0.75 ($(& $bill) vs $expect)"
    Check ((& $lab) -eq 'After 25% subsidy') "25%: label reads '$(& $lab)'"
    Check ([math]::Abs((& $chg) - $full) -lt 0.005) 'Chargeable total is still the full amount'
    Check ([math]::Abs([double]$rep.Cells(9, $cUn).Value2 - $full) -lt 0.005) 'Unpaid is still the full amount (nothing paid yet)'
    Check ($rep.Cells(7, $cUn).NumberFormat -like '*#,##0.00') "billable is formatted as money ($($rep.Cells(7, $cUn).NumberFormat))"

    SetPct 0.125
    Check ((& $lab) -eq 'After 12.5% subsidy') "12.5%: label reads '$(& $lab)'"
    Check ([math]::Abs((& $bill) - [math]::Round($full * 0.875, 2)) -lt 0.005) '12.5%: billable is chargeable x 0.875'

    # Per-job amounts are untouched: the results' Chargeable column still sums to the full total.
    $hdrCol = 0; for ($col = 8; $col -le 40; $col++) { if ([string]$rep.Cells(10, $col).Text -eq 'Chargeable') { $hdrCol = $col } }
    $sum = [double]$xl.WorksheetFunction.Sum($rep.Range($rep.Cells(11, $hdrCol), $rep.Cells(400, $hdrCol)))
    Check ([math]::Abs($sum - $full) -lt 0.005) "per-job Chargeable column still sums to the full amount ($sum)"

    SetPct 1.5
    Check ([math]::Abs((& $bill)) -lt 0.005) 'over 100% reads as 100% (billable 0), not an error'
    SetPct -0.1
    Check ([math]::Abs((& $bill) - $full) -lt 0.005) 'a negative value reads as no subsidy'
    SetPct $null
    Check ([math]::Abs((& $bill) - $full) -lt 0.005) 'a blank value reads as no subsidy'

    Write-Host '=== Follows the filters ==='
    SetPct 0.5
    $rep.Range('E7').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    $filtered = & $chg
    Check ($filtered -gt 0 -and $filtered -lt $full) "the filter narrows Chargeable ($filtered of $full)"
    Check ([math]::Abs((& $bill) - [math]::Round($filtered * 0.5, 2)) -lt 0.011) "billable follows the filtered total ($(& $bill))"

    Write-Host '=== Export report ==='
    SetPct 0.25
    $xl.Run('SetExportNames', $false)
    $xl.CalculateFullRebuild()
    $expBill = & $bill; $expChg = & $chg
    $rep.Activate()
    $xl.Run('btnExportReport')
    Start-Sleep -Milliseconds 300
    $xlsx = Get-ChildItem $workDir -Filter 'PrintCosts-Report-*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Check ($null -ne $xlsx) 'an .xlsx was produced'
    if ($xlsx) {
        $xl2 = New-Object -ComObject Excel.Application; $xl2.Visible = $false; $xl2.DisplayAlerts = $false
        try {
            $out = $xl2.Workbooks.Open($xlsx.FullName); $w = $out.Worksheets(1)
            $r = @{}; for ($i = 1; $i -le 20; $i++) { $k = [string]$w.Cells($i, 1).Value2; if ($k -ne '') { $r[$k] = $i } }
            Check ($r.ContainsKey('Billable after 25% subsidy')) 'a "Billable after 25% subsidy" line is present'
            if ($r.ContainsKey('Billable after 25% subsidy')) {
                $row = $r['Billable after 25% subsidy']
                Check ($row -eq $r['Still owed'] + 1) 'it follows Still owed'
                Check ([math]::Abs([double]$w.Cells($row, 2).Value2 - $expBill) -lt 0.005) "it carries the Reports value ($($w.Cells($row, 2).Value2) vs $expBill)"
                Check ([math]::Abs([double]$w.Cells($r['Total chargeable'], 2).Value2 - $expChg) -lt 0.005) 'Total chargeable is still the full amount'
                Check ($w.Cells($row, 2).NumberFormat -like '*#,##0.00') 'it is formatted as money'
                Check (-not $w.Cells($row, 1).Value2.ToString().StartsWith('=')) 'the label is plain text'
            }
            $tbl = 0; for ($i = 1; $i -le 40; $i++) { if ([string]$w.Cells($i, 1).Value2 -eq 'Date/Time') { $tbl = $i; break } }
            Check ($tbl -gt 0) "the results table still follows the header block (row $tbl)"
            $out.Close($false)
        } finally { $xl2.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2) }
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($script:anyFail) { Write-Host 'FAILED'; exit 1 } else { Write-Host 'All subsidy checks passed.' }
