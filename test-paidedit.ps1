# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Editing Paid directly on the Reports page (2026-09-29).
#
# The results are one spilled formula, so typing into a cell of it blocks the
# spill and blanks the row; modReports remembers which job the selected Paid
# cell shows (on selection) and applies the edit to that job from the Change
# event, then clears the constant. This drives it the way a user would: select
# the cell (real SelectionChange event), then write a value (real Change event).
#
#   - Paid cells carry the Yes/No dropdown and are unlocked; the rest of the
#     results table stays locked.
#   - A choice lands on the right job record, in either room; nothing else
#     changes; the spill is intact afterwards; an audit entry is written.
#   - Re-picking the value already shown changes nothing.
#   - A cell below the results, an invalid value, a multi-cell write and a
#     record that changed since it was selected are all refused, leave no
#     constant behind, and change nothing.
#   - The previous selection's note is honoured (Excel can move the selection
#     before raising Change).
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

    # Retried: right after an event-driven select/write Excel can still be
    # settling and reject the next COM call (see TestCommon.ps1).
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
        for ($c = 1; $c -le 20; $c++) {
            $hd = [string]$rep.Cells(17, $c).Value2
            if ($hd -eq 'Location') { $o.Loc = $c }
            if ($hd -eq 'Job ID') { $o.Job = $c }
            if ($hd -eq 'Paid') { $o.Paid = $c }
        }
        return $o
    }
    function RecordAt([int]$row) {
        return ('tblJobs_' + [string]$rep.Cells($row, $cols.Loc).Value2) + '|' + [string]$rep.Cells($row, $cols.Job).Value2
    }
    function SpillCount { Invoke-ComRetry -Attempts 5 { [int]$rep.Range('A18').SpillingToRange.Rows.Count } }
    function SpillIntact { ([string]$rep.Range('A18').Text -notlike '#SPILL*') -and ((SpillCount) -eq $n) }
    $filterCells = @('B4','B5','B7','B8','B10','F4','F5','F7','F8','F9','B12','F12')
    function FilterState { ($filterCells | ForEach-Object { '{0}={1}' -f $_, [string]$rep.Range($_).Value2 }) -join ';' }
    function AuditActions {
        $aud = $wb.Worksheets('_Audit').ListObjects('tblAudit')
        $a = @(); for ($i = 1; $i -le $aud.ListRows.Count; $i++) { $a += [string]$aud.ListRows($i).Range.Cells(1, 3).Value2 }
        return $a
    }
    function Select-Cell([int]$row) { Invoke-ComRetry -Attempts 5 { $rep.Cells($row, $cols.Paid).Select() | Out-Null }; Start-Sleep -Milliseconds 300 }
    function Write-Cell([int]$row, $value) { Invoke-ComRetry -Attempts 5 { $rep.Cells($row, $cols.Paid).Value2 = $value }; Start-Sleep -Milliseconds 300 }
    function Other([string]$shown) { if ($shown -eq 'Yes') { 'No' } else { 'Yes' } }

    $rep.Activate()
    $xl.CalculateFullRebuild()
    $cols = Cols
    $n = SpillCount
    Write-Host ("  unfiltered Reports shows {0} records; Paid is column {1}" -f $n, $cols.Paid)
    $filtersBefore = FilterState

    # -------------------------------------------------------- the dropdown
    Write-Host '=== Paid cells: dropdown, unlocked; the rest stays locked ==='
    $pc = $rep.Cells(18, $cols.Paid)
    Check ($pc.Validation.Type -eq 3) 'Paid cell has list validation'
    Check ([string]$pc.Validation.Formula1 -eq 'Yes,No') "list is Yes,No (got '$($pc.Validation.Formula1)')"
    Check ([bool]$pc.Validation.InCellDropdown) 'in-cell dropdown arrow on'
    Check (-not [bool]$pc.Locked) 'Paid cell is unlocked'
    Check (-not [bool]$rep.Cells(2000, $cols.Paid).Locked) 'unlocked down to row 2000'
    Check ([bool]$rep.Cells(18, 1).Locked) 'Date/Time cell beside it is still locked'
    Check ([bool]$rep.Cells(18, $cols.Job).Locked) 'Job ID cell is still locked'
    Check ($rep.ProtectContents) 'Reports sheet is protected'

    # ------------------------------------------------ edit, both rooms
    Write-Host ''
    Write-Host '=== Editing Paid lands on the right record (both rooms) ==='
    $done = @{}
    foreach ($row in 18..($n + 17)) {
        $key = RecordAt $row
        $room = $key.Split('|')[0]
        if ($done.ContainsKey($room)) { continue }
        $done[$room] = $true
        $shown = ([string]$rep.Cells($row, $cols.Paid).Value2).Trim()
        $target = Other $shown
        $before = Snapshot
        Select-Cell $row
        Write-Cell $row $target
        $after = Snapshot
        Check ($after[$key] -eq $target) "$room : $key changed '$shown' -> '$target'"
        Check (@($after.Keys | Where-Object { $_ -ne $key -and $after[$_] -ne $before[$_] }).Count -eq 0) "$room : no other record changed"
        Check (SpillIntact) "$room : results spill is intact ($(SpillCount) rows)"
        Check ($rep.Cells($row, $cols.Paid).HasSpill) "$room : the edited cell is part of the spill again, not a constant"
        Check (([string]$rep.Cells($row, $cols.Paid).Value2) -eq $target) "$room : the row now displays '$target'"
        Check ($rep.ProtectContents -and $main.ProtectContents -and $annexe.ProtectContents) "$room : sheets re-protected"
    }
    Check ((FilterState) -eq $filtersBefore) 'filters untouched'
    Check ((AuditActions) -contains 'Paid edited (Reports)') "audit has 'Paid edited (Reports)'"

    # ----------------------------------------------- same value re-picked
    Write-Host ''
    Write-Host '=== Re-picking the value already shown changes nothing ==='
    $row = 19
    $key = RecordAt $row
    $shown = ([string]$rep.Cells($row, $cols.Paid).Value2).Trim()
    # A blank Paid (never recorded) first has to become a real value, through a
    # proper select-then-edit like any other; the re-pick below is then a no-op.
    if ($shown -eq '') { Select-Cell $row; Write-Cell $row 'No'; $shown = 'No' }
    $auditBefore = (AuditActions).Count
    $before = Snapshot
    Select-Cell $row
    Write-Cell $row $shown
    $after = Snapshot
    Check (@($after.Keys | Where-Object { $after[$_] -ne $before[$_] }).Count -eq 0) 'no record changed'
    Check ((AuditActions).Count -eq $auditBefore) 'no audit entry for a no-op'
    Check (SpillIntact) 'spill intact'

    # ------------------------------------------------------ not a record
    Write-Host ''
    Write-Host '=== A cell below the results is refused and cleared ==='
    $xl.Run('SetQuiet', $true)
    $orphanRow = 18 + $n + 2
    $before = Snapshot
    Select-Cell $orphanRow
    Write-Cell $orphanRow 'Yes'
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*not a job record*') 'said it is not a job record'
    Check ([string]$rep.Cells($orphanRow, $cols.Paid).Formula -eq '') 'the typed value was cleared (no orphan constant)'
    $after = Snapshot
    Check (@($after.Keys | Where-Object { $after[$_] -ne $before[$_] }).Count -eq 0) 'no record changed'
    Check (SpillIntact) 'spill intact'

    # ------------------------------------------------------ invalid value
    Write-Host ''
    Write-Host '=== A value that is not Yes/No is refused ==='
    $xl.Run('SetQuiet', $true)
    $row = 20
    $before = Snapshot
    Select-Cell $row
    Write-Cell $row 'Maybe'
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*Paid must be Yes or No*') 'said Paid must be Yes or No'
    $after = Snapshot
    Check (@($after.Keys | Where-Object { $after[$_] -ne $before[$_] }).Count -eq 0) 'no record changed'
    Check (SpillIntact) 'spill intact'
    Check ($rep.Cells($row, $cols.Paid).HasSpill) 'cell is part of the spill, not a constant'

    # ---------------------------------------------------------- many cells
    Write-Host ''
    Write-Host '=== Writing several Paid cells at once is refused ==='
    $xl.Run('SetQuiet', $true)
    $before = Snapshot
    Select-Cell 19
    Invoke-ComRetry { $rep.Range($rep.Cells(19, $cols.Paid), $rep.Cells(21, $cols.Paid)).Value2 = 'Yes' }
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*Change one Paid cell at a time*') 'said to change one cell at a time'
    $after = Snapshot
    Check (@($after.Keys | Where-Object { $after[$_] -ne $before[$_] }).Count -eq 0) 'no record changed'
    Check (SpillIntact) 'spill intact'

    # -------------------------------------------------------- stale record
    Write-Host ''
    Write-Host '=== A record that changed since it was selected is refused ==='
    $xl.Run('SetQuiet', $true)
    $row = 21
    $key = RecordAt $row
    $room = $key.Split('|')[0]
    $lo = if ($room -eq 'tblJobs_MAIN') { $loMain } else { $loAnnexe }
    $shown = ([string]$rep.Cells($row, $cols.Paid).Value2).Trim()
    Select-Cell $row
    # Change the record behind the report's back (events off, so the report
    # redraws but no selection change re-notes the cell).
    $ids = $lo.ListColumns.Item('Job ID').DataBodyRange.Value2
    $idx = 0; for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if (('tblJobs_' + $room.Substring(8) + '|' + [string]$ids[$i, 1]) -eq $key -or ($room + '|' + [string]$ids[$i, 1]) -eq $key) { $idx = $i } }
    Check ($idx -gt 0) 'found the record to disturb'
    $behind = Other $shown
    $xl.EnableEvents = $false
    $lo.ListColumns.Item('Paid').DataBodyRange.Cells($idx, 1).Value2 = $behind
    $xl.EnableEvents = $true
    $xl.Calculate()
    Write-Cell $row (Other $behind)
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*has changed since it was drawn*') 'said the record has changed'
    $now = Snapshot
    Check ($now[$key] -eq $behind) "the record kept the value set behind the report ('$behind')"
    Check (SpillIntact) 'spill intact'

    # ---------------------------------- previous selection honoured (order)
    Write-Host ''
    Write-Host '=== Selection moved on before Change arrived: previous cell honoured ==='
    $xl.Run('SetQuiet', $true)
    $rowA = 20; $rowB = 21
    $keyA = RecordAt $rowA; $keyB = RecordAt $rowB
    $shownA = ([string]$rep.Cells($rowA, $cols.Paid).Value2).Trim()
    $targetA = Other $shownA
    $before = Snapshot
    Select-Cell $rowA
    Select-Cell $rowB
    Write-Cell $rowA $targetA
    $after = Snapshot
    Check ($after[$keyA] -eq $targetA) "record A changed to '$targetA' via the previous note"
    Check ($after[$keyB] -eq $before[$keyB]) 'record B (the current selection) untouched'
    Check (SpillIntact) 'spill intact'

    # ---------------------------------- bulk buttons still fine afterwards
    Write-Host ''
    Write-Host '=== Mark all as... still works after edits ==='
    $rep.Range('F7').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    $rng = $rep.Range('A18').SpillingToRange
    $nn = $rng.Rows.Count
    $xl.Run('SetQuiet', $true)
    $xl.Run('MarkVisibleReportsConfirmed', $rep, $rng, $cols.Loc, $cols.Job, $nn, 'Yes')
    $after = Snapshot
    $all = 0
    for ($r = 1; $r -le $nn; $r++) { if ($after[('tblJobs_' + [string]$rng.Cells($r, $cols.Loc).Value2 + '|' + [string]$rng.Cells($r, $cols.Job).Value2)] -eq 'Yes') { $all++ } }
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
