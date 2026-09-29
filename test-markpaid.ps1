# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Reports page "Mark all as..." Paid / Unpaid cluster (2026-09-29), and the
# at-least-one-filter safeguard it shares with Delete visible records.
#
#   - The cluster exists: label at O1, Paid and Unpaid buttons in rows 2 and 3.
#   - HasActiveFilter mirrors Criteria(): blank, spaces, and text that will not
#     coerce to a date/number do NOT count; a real value does.
#   - With no filter, MarkVisibleReportsConfirmed and DeleteVisibleReportsConfirmed
#     both refuse and change nothing.
#   - With a filter, only the visible rows change; hidden rows are untouched;
#     the filter boxes are exactly as they were; the room sheets end protected.
#   - The interactive entry point still asks first (Ask declines under SetQuiet).
#
# MarkVisibleReportsConfirmed / DeleteVisibleReportsConfirmed are called
# directly, bypassing Ask() - see test-deletereports.ps1 for why.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself.

$ErrorActionPreference = 'Stop'
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
    $wb = $xl.Workbooks.Open($f)
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.Run('SetQuiet', $true)
    $rep = $wb.Worksheets('Reports')
    $main = $wb.Worksheets('Example Print Room')
    $annexe = $wb.Worksheets('Annexe')
    $loMain = $main.ListObjects('tblJobs_MAIN')
    $loAnnexe = $annexe.ListObjects('tblJobs_ANNEX')

    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }

    # Job ID -> Paid, across both rooms. Trimmed text; blank stays ''.
    function Snapshot {
        $h = @{}
        foreach ($lo in @($loMain, $loAnnexe)) {
            $ids = $lo.ListColumns.Item('Job ID').DataBodyRange.Value2
            $paid = $lo.ListColumns.Item('Paid').DataBodyRange.Value2
            for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
                $h[($lo.Name + '|' + [string]$ids[$i, 1])] = ([string]$paid[$i, 1]).Trim()
            }
        }
        return $h
    }
    # The visible set, as table|jobId keys, plus the columns the routine needs.
    function VisibleInfo {
        $rng = $rep.Range('A16').SpillingToRange
        $locCol = 0; $jobCol = 0
        for ($c = 1; $c -le 20; $c++) {
            $hd = [string]$rep.Cells(15, $c).Value2
            if ($hd -eq 'Location') { $locCol = $c }
            if ($hd -eq 'Job ID') { $jobCol = $c }
        }
        $keys = @()
        for ($r = 1; $r -le $rng.Rows.Count; $r++) {
            $code = [string]$rng.Cells($r, $locCol).Value2
            $tbl = 'tblJobs_' + $code
            $keys += ($tbl + '|' + [string]$rng.Cells($r, $jobCol).Value2)
        }
        return @{ Rng = $rng; LocCol = $locCol; JobCol = $jobCol; N = $rng.Rows.Count; VKeys = $keys }
    }
    $filterCells = @('B4','B5','B7','B8','F4','F5','F6','F7','F8','B10','F10','O10')
    function FilterState { ($filterCells | ForEach-Object { '{0}={1}' -f $_, [string]$rep.Range($_).Value2 }) -join ';' }
    function ClearFilters { foreach ($a in @('B4','B5','B7','B8','F4','F5','F6','F7','F8')) { $rep.Range($a).ClearContents() | Out-Null }; $xl.CalculateFullRebuild() }

    # ------------------------------------------------------------ the cluster
    Write-Host '=== The cluster on the Reports header ==='
    Check ([string]$rep.Range('O1').Text -eq 'Mark all as...') "O1 label reads 'Mark all as...' (got '$($rep.Range('O1').Text)')"
    $paidBtn = $null; $unpaidBtn = $null
    foreach ($s in $rep.Shapes) {
        if ($s.Name -like 'pcb_btnMarkPaid_*') { $paidBtn = $s }
        if ($s.Name -like 'pcb_btnMarkUnpaid_*') { $unpaidBtn = $s }
    }
    Check ($null -ne $paidBtn) 'Paid button exists'
    Check ($null -ne $unpaidBtn) 'Unpaid button exists'
    if ($paidBtn -and $unpaidBtn) {
        Check ($paidBtn.TextFrame.Characters().Text -eq 'Paid') "Paid button caption is 'Paid'"
        Check ($unpaidBtn.TextFrame.Characters().Text -eq 'Unpaid') "Unpaid button caption is 'Unpaid'"
        Check ([string]$paidBtn.OnAction -like '*btnMarkPaid') "Paid button runs btnMarkPaid (got '$($paidBtn.OnAction)')"
        Check ([string]$unpaidBtn.OnAction -like '*btnMarkUnpaid') "Unpaid button runs btnMarkUnpaid (got '$($unpaidBtn.OnAction)')"
        $r2 = $rep.Range('O2'); $r3 = $rep.Range('O3'); $r4 = $rep.Range('O4')
        Check (($paidBtn.Top -ge $r2.Top) -and (($paidBtn.Top + $paidBtn.Height) -le $r3.Top)) 'Paid button sits inside row 2'
        Check (($unpaidBtn.Top -ge $r3.Top) -and (($unpaidBtn.Top + $unpaidBtn.Height) -le $r4.Top)) 'Unpaid button sits inside row 3'
        Check ([math]::Abs($paidBtn.Left - $rep.Range('O1').Left) -lt 1 -and [math]::Abs($unpaidBtn.Left - $rep.Range('O1').Left) -lt 1) 'both buttons are left-aligned under the O1 label'
        Check (($paidBtn.Left -le 900) -and ($unpaidBtn.Left -le 900)) "both buttons are within the first screenful (Left <= 900pt; Paid at $([int]$paidBtn.Left))"
        Check ($paidBtn.Top -lt $unpaidBtn.Top) 'Paid is the top button, Unpaid the bottom'
    }

    # ------------------------------------------------- HasActiveFilter rules
    Write-Host ''
    Write-Host '=== HasActiveFilter mirrors Criteria() ==='
    ClearFilters
    Check (-not [bool]$xl.Run('HasActiveFilter', $rep)) 'no filter set -> false'
    $rep.Range('B4').Value2 = '   '
    Check (-not [bool]$xl.Run('HasActiveFilter', $rep)) 'spaces only in the name box -> false'
    $rep.Range('B4').ClearContents() | Out-Null
    $rep.Range('B7').Value2 = 'not a date'
    Check (-not [bool]$xl.Run('HasActiveFilter', $rep)) 'From date text that will not coerce (ignored by Criteria) -> false'
    $rep.Range('B7').ClearContents() | Out-Null
    $rep.Range('F8').Value2 = 'lots'
    Check (-not [bool]$xl.Run('HasActiveFilter', $rep)) 'Quantity text that will not coerce -> false'
    $rep.Range('F8').ClearContents() | Out-Null
    $rep.Range('B10').Value2 = 'Qty'
    Check (-not [bool]$xl.Run('HasActiveFilter', $rep)) 'Sort by set, no filter -> false (sort is not a filter)'
    $rep.Range('B10').ClearContents() | Out-Null
    $rep.Range('B7').Value2 = [double](Get-Date '2020-01-01').ToOADate()
    Check ([bool]$xl.Run('HasActiveFilter', $rep)) 'a real From date -> true'
    $rep.Range('B7').ClearContents() | Out-Null
    $rep.Range('F6').Value2 = 'Epson SureColor P9500'
    Check ([bool]$xl.Run('HasActiveFilter', $rep)) 'Printer set -> true'
    ClearFilters

    # ---------------------------------------------- refused with no filter
    Write-Host ''
    Write-Host '=== No filter: everything refuses and changes nothing ==='
    $xl.CalculateFullRebuild()
    $rep.Activate()
    $all = VisibleInfo
    Write-Host ("  unfiltered sheet shows {0} records" -f $all.N)
    $before = Snapshot
    $rowsBefore = $loMain.ListRows.Count + $loAnnexe.ListRows.Count

    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $all.Rng, $all.LocCol, $all.JobCol, $all.N, 'Yes')
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*Set at least one filter before you mark records as Paid*') 'Mark Paid said to set a filter first'
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReports', 'No')
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*Set at least one filter before you mark records as Unpaid*') 'interactive Mark Unpaid said to set a filter first'
    Check ($log -notlike '*DECLINED*') 'interactive path refused BEFORE reaching the confirmation prompt'
    $xl.Run('SetQuiet', $true)
    $xl.Run('DeleteVisibleReportsConfirmed', $rep, $all.Rng, $all.LocCol, $all.JobCol, $all.N)
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*Set at least one filter before you delete records*') 'Delete visible said to set a filter first'
    $xl.Run('SetQuiet', $true)
    $xl.Run('DeleteVisibleReports')
    $log = [string]$xl.Run('QuietLog')
    Check (($log -like '*Set at least one filter before you delete records*') -and ($log -notlike '*DECLINED*')) 'interactive Delete refused BEFORE its confirmation prompt'

    $afterRefused = Snapshot
    $rowsAfter = $loMain.ListRows.Count + $loAnnexe.ListRows.Count
    Check ($rowsAfter -eq $rowsBefore) "no rows deleted ($rowsBefore -> $rowsAfter)"
    $diff = @($before.Keys | Where-Object { $before[$_] -ne $afterRefused[$_] }).Count
    Check ($diff -eq 0) 'no Paid value changed'

    # ------------------------------------------------- filtered: Paid = Yes
    Write-Host ''
    Write-Host '=== With a filter: only the visible rows change ==='
    $rep.Range('F6').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    $vis = VisibleInfo
    Write-Host ("  filtered to Printer = 'Epson SureColor P9500': {0} of {1} records visible" -f $vis.N, $before.Count)
    if ($vis.N -eq 0) { Write-Host 'FAIL: expected at least one visible record'; exit 1 }
    if ($vis.N -ge $before.Count) { Write-Host 'FAIL: the filter must leave some records hidden for this test to mean anything'; exit 1 }
    $visSet = @{}; foreach ($k in $vis.VKeys) { $visSet[$k] = $true }
    $filtersBefore = FilterState

    # Start from a known state: every hidden row Unpaid so a stray write is visible.
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $vis.Rng, $vis.LocCol, $vis.JobCol, $vis.N, 'No')
    $baseline = Snapshot
    Check (@($vis.VKeys | Where-Object { $baseline[$_] -ne 'No' }).Count -eq 0) 'visible rows now all read No (baseline)'

    $xl.Run('SetQuiet', $true)
    $vis = VisibleInfo
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $vis.Rng, $vis.LocCol, $vis.JobCol, $vis.N, 'Yes')
    $log = [string]$xl.Run('QuietLog')
    Write-Host $log
    $after = Snapshot
    $visWrong = @($vis.VKeys | Where-Object { $after[$_] -ne 'Yes' }).Count
    Check ($visWrong -eq 0) "all $($vis.N) visible rows are now Yes"
    $hiddenChanged = @($after.Keys | Where-Object { -not $visSet.ContainsKey($_) -and $after[$_] -ne $baseline[$_] }).Count
    Check ($hiddenChanged -eq 0) 'no hidden row was touched'
    Check ((FilterState) -eq $filtersBefore) 'filter and sort boxes are exactly as they were'
    Check (([string]$rep.Range('F6').Value2) -eq 'Epson SureColor P9500') 'Printer filter still set'
    Check ($log -like '*marked Paid*') 'confirmation message names Paid'
    Check ($main.ProtectContents -and $annexe.ProtectContents) 'room sheets are protected again afterwards'

    # ---------------------------------------------- filtered: Paid -> Unpaid
    Write-Host ''
    Write-Host '=== Unpaid, and already-at-target rows are left alone ==='
    $xl.CalculateFullRebuild()
    $vis = VisibleInfo
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $vis.Rng, $vis.LocCol, $vis.JobCol, $vis.N, 'No')
    $after2 = Snapshot
    Check (@($vis.VKeys | Where-Object { $after2[$_] -ne 'No' }).Count -eq 0) 'all visible rows are now No'
    Check (@($after2.Keys | Where-Object { -not $visSet.ContainsKey($_) -and $after2[$_] -ne $baseline[$_] }).Count -eq 0) 'no hidden row was touched'
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $vis.Rng, $vis.LocCol, $vis.JobCol, $vis.N, 'No')
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like "*0 records marked Unpaid*already were Unpaid*" -or $log -like "*0 records marked Unpaid*already was Unpaid*") 'a repeat run reports 0 changed, all already Unpaid'

    # ---------------------------------------------- interactive still asks
    Write-Host ''
    Write-Host '=== The interactive path still confirms first ==='
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReports', 'Yes')
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*DECLINED (quiet mode)*Mark visible records as Paid*') 'Ask() was reached for Paid'
    Check ($log -like "*Mark $($vis.N) visible record*as Paid?*") 'prompt states the record count'
    $after3 = Snapshot
    Check (@($after3.Keys | Where-Object { $after3[$_] -ne $after2[$_] }).Count -eq 0) 'declining changed nothing'

    # ----------------------------------------------------------- audit log
    Write-Host ''
    Write-Host '=== Audit log ==='
    $aud = $wb.Worksheets('_Audit').ListObjects('tblAudit')
    $actions = @(); for ($i = 1; $i -le $aud.ListRows.Count; $i++) { $actions += [string]$aud.ListRows($i).Range.Cells(1, 3).Value2 }
    Check ($actions -contains 'Mark visible Paid (Reports)') "audit has 'Mark visible Paid (Reports)'"
    Check ($actions -contains 'Mark visible Unpaid (Reports)') "audit has 'Mark visible Unpaid (Reports)'"

    # ------------------------------------- delete still works once filtered
    Write-Host ''
    Write-Host '=== Delete visible still works with a filter set ==='
    $xl.CalculateFullRebuild()
    $vis = VisibleInfo
    $rowsBefore = $loMain.ListRows.Count + $loAnnexe.ListRows.Count
    $xl.Run('SetQuiet', $true)
    $xl.Run('DeleteVisibleReportsConfirmed', $rep, $vis.Rng, $vis.LocCol, $vis.JobCol, $vis.N)
    $rowsAfter = $loMain.ListRows.Count + $loAnnexe.ListRows.Count
    Check (($rowsBefore - $rowsAfter) -eq $vis.N) "exactly $($vis.N) rows deleted ($rowsBefore -> $rowsAfter)"

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'DONE'
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
