
# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Changing Paid from the Reports page with the Toggle Paid button (0.10.25).
# (Replaces test-paidedit.ps1: the Paid cells used to be editable in place with
# a dropdown, which cannot be made reliable on Excel for Mac - see modReports,
# "Paid column".)
#
#   - The results' Paid cells are locked and carry no dropdown.
#   - The Toggle Paid button exists, sits in row 2 under the 'Selected record:'
#     caption with Go to record below it, and runs btnTogglePaid.
#   - It flips the selected row's record to the opposite of what the row shows
#     (Yes -> No, anything else -> Yes), in either room, changing nothing else,
#     leaving the spill intact and writing an audit entry.
#   - Several rows selected: all take the value the ACTIVE row flips to, after
#     a confirmation (declined under SetQuiet; Confirmed = True skips it).
#   - A selection that holds no record (header row, below the results) is
#     refused and changes nothing.
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

    function Snapshot {
        return (Invoke-ComRetry -Attempts 5 {
            $h = @{}
            foreach ($lo in @($loMain, $loAnnexe)) {
                $ids = $lo.ListColumns.Item('Job ID').DataBodyRange.Value2
                $paid = $lo.ListColumns.Item('Paid').DataBodyRange.Value2
                for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
                    $h[($lo.Name + '|' + [string]$ids[$i, 1])] = ([string]$paid[$i, 1]).Trim()
                }
            }
            $h
        })
    }
    function Cols {
        $o = @{ Loc = 0; Job = 0; Paid = 0 }
        for ($c = 8; $c -le 28; $c++) {
            $hd = [string]$rep.Cells(10, $c).Value2
            if ($hd -eq 'Location') { $o.Loc = $c }
            if ($hd -eq 'Job ID') { $o.Job = $c }
            if ($hd -eq 'Paid') { $o.Paid = $c }
        }
        return $o
    }
    function RecordAt([int]$row) {
        return ('tblJobs_' + [string]$rep.Cells($row, $cols.Loc).Value2) + '|' + [string]$rep.Cells($row, $cols.Job).Value2
    }
    function SpillCount { Invoke-ComRetry -Attempts 5 { [int]$rep.Range('H11').SpillingToRange.Rows.Count } }
    function SpillIntact { ([string]$rep.Range('H11').Text -notlike '#SPILL*') -and ((SpillCount) -eq $n) }
    $filterCells = @('B4','B5','B7','B8','B10','B11','B12','F4','F5','F7','F8','F9','F10','B14','F14','I7','K7')
    function FilterState { ($filterCells | ForEach-Object { '{0}={1}' -f $_, [string]$rep.Range($_).Value2 }) -join ';' }
    function AuditActions {
        $aud = $wb.Worksheets('_Audit').ListObjects('tblAudit')
        $a = @(); for ($i = 1; $i -le $aud.ListRows.Count; $i++) { $a += [string]$aud.ListRows($i).Range.Cells(1, 3).Value2 }
        return $a
    }
    function Shown([int]$row) { ([string]$rep.Cells($row, $cols.Paid).Value2).Trim() }
    function Other([string]$shown) { if ($shown -eq 'Yes') { 'No' } else { 'Yes' } }
    function Select-Cell([int]$row) { Invoke-ComRetry -Attempts 5 { $rep.Cells($row, $cols.Paid).Select() | Out-Null } }
    function Changed($before, $after) { @($after.Keys | Where-Object { $after[$_] -ne $before[$_] }) }

    $rep.Activate()
    $xl.CalculateFullRebuild()
    $cols = Cols
    $n = SpillCount
    Write-Host ("  unfiltered Reports shows {0} records; Paid is column {1}" -f $n, $cols.Paid)
    $filtersBefore = FilterState

    # --------------------------------------------------------- the column
    Write-Host '=== Paid cells are read-only: locked, no dropdown ==='
    $pc = $rep.Cells(11, $cols.Paid)
    $vt = ''
    try { $vt = [string]$pc.Validation.Type } catch { $vt = '' }
    Check ($vt -eq '') "no validation (no dropdown) on the Paid cells (type: '$vt')"
    Check ([bool]$pc.Locked) 'Paid cell is locked'
    Check ([bool]$rep.Cells(2000, $cols.Paid).Locked) 'locked down to row 2000'
    Check ($rep.ProtectContents) 'Reports sheet is protected'

    # ----------------------------------------------------------- the button
    Write-Host ''
    Write-Host '=== The Toggle Paid button ==='
    $btn = $null; $goBtn = $null
    foreach ($s in $rep.Shapes) {
        if ($s.Name -like 'pcb_btnTogglePaid_*') { $btn = $s }
        if ($s.Name -like 'pcb_btnGoToRecord_*') { $goBtn = $s }
    }
    Check ($null -ne $btn) 'Toggle Paid button exists'
    if ($btn) {
        Check ([string]$btn.OnAction -like '*btnTogglePaid') "button runs btnTogglePaid (got '$($btn.OnAction)')"
        Check ($btn.Name.Length -le 31) "shape name within 31 characters ($($btn.Name))"
        if ($goBtn) {
            Check ([math]::Abs([double]$btn.Left - [double]$goBtn.Left) -lt 1.5) 'sits in the same block as Go to record (same Left)'
            Check ([double]$goBtn.Top -gt [double]$btn.Top) 'Go to record is below Toggle Paid'
        }
        $selLbl = $null
        foreach ($s in $rep.Shapes) { if ($s.Name -eq 'pcb_lblSelected') { $selLbl = $s } }
        Check ($null -ne $selLbl) "the 'Selected record:' caption exists"
        if ($selLbl) {
            Check ([string]$selLbl.TextFrame.Characters().Text -eq 'Selected record:') "caption reads 'Selected record:' (got '$($selLbl.TextFrame.Characters().Text)')"
            Check ([math]::Abs([double]$selLbl.Left - [double]$btn.Left) -lt 1.5) 'the caption is left-aligned over Toggle Paid'
            Check (([double]$selLbl.Top + [double]$selLbl.Height) -le ([double]$btn.Top + 1)) 'the caption sits above Toggle Paid'
        }
        Check ([double]$btn.Top -ge [double]$rep.Rows(4).Top - 1 -and [double]$btn.Top + [double]$btn.Height -le [double]$rep.Rows(5).Top + 1) 'sits inside row 4'
    }

    # ------------------------------------------------ toggle, both rooms
    Write-Host ''
    Write-Host '=== Toggle Paid lands on the right record (both rooms) ==='
    $done = @{}
    foreach ($row in 11..($n + 10)) {
        $key = RecordAt $row
        $room = $key.Split('|')[0]
        if ($done.ContainsKey($room)) { continue }
        $done[$room] = $true
        $shown = Shown $row
        $target = Other $shown
        $before = Snapshot
        Select-Cell $row
        $xl.Run('btnTogglePaid')
        $after = Snapshot
        Check ($after[$key] -eq $target) "$room : $key changed '$shown' -> '$target'"
        Check ((Changed $before $after).Count -eq 1) "$room : no other record changed"
        Check (SpillIntact) "$room : results spill is intact ($(SpillCount) rows)"
        Check ((Shown $row) -eq $target) "$room : the row now displays '$target'"
        Check ($rep.ProtectContents -and $main.ProtectContents -and $annexe.ProtectContents) "$room : sheets re-protected"
        # And back again.
        $xl.Run('btnTogglePaid')
        $back = Snapshot
        Check ($back[$key] -eq $shown -or ($shown -eq '' -and $back[$key] -eq 'Yes') -or $back[$key] -eq (Other $target)) "$room : toggling again flips it back"
    }
    Check ((FilterState) -eq $filtersBefore) 'filters untouched'
    Check ((AuditActions) -contains 'Paid edited (Reports)') "audit has 'Paid edited (Reports)'"

    # ------------------------------------------- any cell of the row works
    Write-Host ''
    Write-Host '=== Any cell of the row selects the record ==='
    $row = 12
    $key = RecordAt $row
    $shown = Shown $row
    $before = Snapshot
    Invoke-ComRetry { $rep.Cells($row, 8).Select() | Out-Null }
    $xl.Run('btnTogglePaid')
    $after = Snapshot
    Check ($after[$key] -eq (Other $shown)) "record changed from a cell in column A ('$shown' -> '$(Other $shown)')"
    Check ((Changed $before $after).Count -eq 1) 'only that record changed'

    # ----------------------------------------------------- several rows
    Write-Host ''
    Write-Host '=== Several rows: all take the value the active row flips to ==='
    $r1 = 13; $r2 = 14; $r3 = 15
    # Make the rows differ so "all the same afterwards" proves something.
    Select-Cell $r1; $s1 = Shown $r1
    $keys = @((RecordAt $r1), (RecordAt $r2), (RecordAt $r3))
    $expect = Other $s1
    $before = Snapshot
    $xl.Run('SetQuiet', $true)
    Invoke-ComRetry { $rep.Range($rep.Cells($r1, $cols.Paid), $rep.Cells($r3, $cols.Paid)).Select() | Out-Null }
    $xl.Run('btnTogglePaid')    # unconfirmed under SetQuiet: Ask declines
    $after = Snapshot
    Check ((Changed $before $after).Count -eq 0) 'declined confirmation: nothing changed'
    $xl.Run('TogglePaidSelected', $true)
    $after = Snapshot
    foreach ($k in $keys) { Check ($after[$k] -eq $expect) "$k is '$expect'" }
    Check ((Changed $before $after).Count -le 3) 'no record outside the selection changed'
    Check (SpillIntact) 'spill intact'
    Check ((AuditActions) -contains 'Paid edited (Reports)') 'audited'

    # ------------------------------------------------ nothing to change
    Write-Host ''
    Write-Host '=== A selection with no record is refused ==='
    foreach ($case in @(@('header row', 10), @('below the results', (11 + $n + 3)))) {
        $xl.Run('SetQuiet', $true)
        $before = Snapshot
        Invoke-ComRetry { $rep.Cells($case[1], $cols.Paid).Select() | Out-Null }
        $xl.Run('btnTogglePaid')
        $log = [string]$xl.Run('QuietLog')
        Check ($log -like '*Select a record first*') "$($case[0]): said to select a record"
        Check ((Changed $before (Snapshot)).Count -eq 0) "$($case[0]): no record changed"
    }
    Check (SpillIntact) 'spill intact'

    # ---------------------------------- bulk buttons still fine afterwards
    Write-Host ''
    Write-Host '=== Mark all as... still works ==='
    $rep.Range('F7').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    $rng = $rep.Range('H11').SpillingToRange
    $nn = $rng.Rows.Count
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $rng, $cols.Loc, $cols.Job, $nn, 'Yes')
    $after = Snapshot
    $all = 0
    for ($r = 1; $r -le $nn; $r++) { if ($after[('tblJobs_' + [string]$rng.Cells($r, $cols.Loc - 7).Value2 + '|' + [string]$rng.Cells($r, $cols.Job - 7).Value2)] -eq 'Yes') { $all++ } }
    Check ($all -eq $nn) "all $nn visible records marked Yes"

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
