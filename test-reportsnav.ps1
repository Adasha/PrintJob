# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# 0.10.25 Reports additions and fixes (from the Mac test notes):
#
#   - Filters: "Has notes" (Yes/No) and the Min/Max Chargeable cost range.
#   - Clear all filters: empties every filter box, leaves Sort by alone.
#   - Go to record: from a selected results row to that record's own sheet and
#     row (both rooms).
#   - Header buttons stay in place when a column is hidden (the same job the
#     location sheets' buttons already did), and the label follows its buttons.
#   - Button names: every shape name is <= 31 characters and agrees with its
#     macro, and HealButtons restores a macro even after a file rename -
#     the root cause of "Cannot run the macro ...DelVis" on Delete visible.
#   - Toggle Paid while the view is filtered to Paid = No leaves a sound
#     report.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself (AutoSave on a
# OneDrive-backed handle commits regardless of Close($false)). Closes WITHOUT
# saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$HDR = 10      # modReports.REP_HDR_ROW
$FIRST = 11    # modReports.REP_FIRST_ROW
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$anyFail = $false
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.CalculateFullRebuild()
    Start-Sleep -Seconds 3
    $xl.Run('SetQuiet', $true)
    $rep = Invoke-ComRetry { $wb.Worksheets('Reports') }
    $main = $wb.Worksheets('Example Print Room')
    $annexe = $wb.Worksheets('Annexe')
    $loMain = $main.ListObjects('tblJobs_MAIN')
    $loAnnexe = $annexe.ListObjects('tblJobs_ANNEX')
    $rep.Activate()

    function RepCol([string]$header) {
        for ($c = 8; $c -le 48; $c++) { if ([string]$rep.Cells($HDR, $c).Value2 -like $header) { return $c } }
        return 0
    }
    function Jobs { $xl.CalculateFullRebuild(); [int]$rep.Range('I9').Text }
    function SpillOk { [string]$rep.Cells($FIRST, 8).Text -notlike '#SPILL*' }
    function Clear-AllBoxes {
        foreach ($a in 'B4','B5','B7','B8','B10','B11','B12','F4','F5','F7','F8','F9','F10','B14','F14','I7','K7') { $rep.Range($a).ClearContents() | Out-Null }
        $xl.CalculateFullRebuild()
    }
    function Shape([string]$macro) {
        foreach ($s in $rep.Shapes) { if ($s.Name -like "pcb_${macro}_*") { return $s } }
        return $null
    }

    $total = Jobs
    Check ($total -ge 5) "unfiltered Reports shows $total records"

    # ------------------------------------------------------------ Has notes
    Write-Host ''
    Write-Host '=== Has notes filter ==='
    # The sample data already carries notes on most rows, so start from none.
    $loAnnexe.ListColumns('Notes').DataBodyRange.ClearContents() | Out-Null
    $loMain.ListColumns('Notes').DataBodyRange.ClearContents() | Out-Null
    $notesMain = $loMain.ListColumns('Notes').DataBodyRange
    $notesMain.Cells(1, 1).Value2 = 'check this one'
    $notesMain.Cells(2, 1).Value2 = '   '            # spaces only: counts as no notes
    $total = Jobs
    $rep.Range('F14').Value2 = 'Yes'
    Check ((Jobs) -eq 1) 'Has notes = Yes shows only the record with real notes (spaces-only is not a note)'
    $rep.Range('F14').Value2 = 'No'
    Check ((Jobs) -eq ($total - 1)) 'Has notes = No shows every other record'
    $rep.Range('F14').ClearContents() | Out-Null
    Check ((Jobs) -eq $total) 'Has notes blank shows everything'
    $notesCol = RepCol 'Notes'
    if ($notesCol -gt 0) {
        $bad = 0
        $sp = $rep.Range("H$FIRST").SpillingToRange
        for ($r = 1; $r -le $sp.Rows.Count; $r++) { if ([string]$sp.Cells($r, $notesCol - 7).Text -eq '0') { $bad++ } }
        Check ($bad -eq 0) 'no blank note is displayed as 0'
    }

    # ---------------------------------------------------------- cost range
    Write-Host ''
    Write-Host '=== Min / Max chargeable cost ==='
    $cc = RepCol 'Chargeable*'
    Check ($cc -gt 0) "found the Chargeable cost column ($cc)"
    $sp = $rep.Range("H$FIRST").SpillingToRange
    $vals = @(); for ($r = 1; $r -le $sp.Rows.Count; $r++) { $vals += [math]::Round([double]$sp.Cells($r, $cc - 7).Value2, 4) }
    $sorted = $vals | Sort-Object
    $mid = [math]::Round([double]$sorted[[int]($sorted.Count / 2)], 4)
    $rep.Range('B10').Value2 = $mid
    $expMin = @($vals | Where-Object { $_ -ge $mid }).Count
    Check ((Jobs) -eq $expMin) "Min cost $mid shows $expMin records (>=, inclusive)"
    $rep.Range('B10').ClearContents() | Out-Null
    $rep.Range('B11').Value2 = $mid
    $expMax = @($vals | Where-Object { $_ -le $mid }).Count
    Check ((Jobs) -eq $expMax) "Max cost $mid shows $expMax records (<=, inclusive)"
    $rep.Range('B10').Value2 = $mid
    $expBoth = @($vals | Where-Object { $_ -eq $mid }).Count
    Check ((Jobs) -eq $expBoth) "Min = Max = $mid shows the $expBoth records at exactly that cost"
    Check ([bool]$xl.Run('HasActiveFilter', $rep)) 'a cost range counts as a filter'
    Clear-AllBoxes
    $rep.Range('B10').Value2 = 'abc'
    Check ((Jobs) -eq $total) 'non-numeric Min cost is ignored'
    Check (-not [bool]$xl.Run('HasActiveFilter', $rep)) 'non-numeric Min cost is not a filter'
    Clear-AllBoxes

    # -------------------------------------------------------- clear filters
    Write-Host ''
    Write-Host '=== Clear all filters ==='
    $rep.Range('B4').Value2 = 'Patel'
    $rep.Range('F7').Value2 = 'Epson SureColor P9500'
    $rep.Range('B10').Value2 = 1
    $rep.Range('B11').Value2 = 9999
    $rep.Range('F14').Value2 = 'No'
    $rep.Range('B12').Value2 = 'No'
    $rep.Range('B7').Value2 = [double](Get-Date '2020-01-01').ToOADate()
    $rep.Range('I7').Value2 = 'Qty'
    $rep.Range('K7').Value2 = 'Descending'
    $xl.Run('btnClearFilters')
    $left = @()
    foreach ($a in 'B4','B5','B7','B8','B10','B11','B12','F4','F5','F7','F8','F9','F10','B14','F14') {
        if ([string]$rep.Range($a).Formula -ne '') { $left += $a }
    }
    Check ($left.Count -eq 0) "every filter box is empty afterwards (still set: $($left -join ', '))"
    Check (([string]$rep.Range('I7').Value2 -eq 'Qty') -and ([string]$rep.Range('K7').Value2 -eq 'Descending')) 'Sort by / Sort direction are left alone'
    Check ((Jobs) -eq $total) 'all records are back'
    Check ($rep.Range('F7').Validation.Type -eq 3) 'a cleared box keeps its dropdown'
    Clear-AllBoxes

    # ---------------------------------------------------------- go to record
    Write-Host ''
    Write-Host '=== Go to record ==='
    $locCol = RepCol 'Location'
    $jobCol = RepCol 'Job ID'
    $sp = $rep.Range("H$FIRST").SpillingToRange
    $codes = @(); for ($r = 1; $r -le $sp.Rows.Count; $r++) { $codes += [string]$sp.Cells($r, $locCol - 7).Value2 }
    Check ((@($codes | Select-Object -Unique)).Count -ge 2) "results span both rooms ($((@($codes | Select-Object -Unique)) -join ', '))"
    foreach ($pick in @($codes[0], ($codes | Where-Object { $_ -ne $codes[0] } | Select-Object -First 1))) {
        if (-not $pick) { continue }
        $idx = [array]::IndexOf($codes, $pick)
        $row = $FIRST + $idx
        $wantJob = ([string]$rep.Cells($row, $jobCol).Value2).Trim()
        $rep.Activate()
        Invoke-ComRetry { $rep.Cells($row, 8).Select() | Out-Null }
        $xl.Run('btnGoToRecord')
        $act = $wb.ActiveSheet
        $lo = $act.ListObjects | Where-Object { $_.Name -eq ('tblJobs_' + $pick) } | Select-Object -First 1
        Check ($null -ne $lo) "room $pick : went to its own sheet ('$($act.Name)')"
        if ($lo) {
            $sel = $xl.Selection
            $jobIdx = Col $lo 'Job ID'
            $gotJob = ([string]$sel.Cells(1, $jobIdx).Value2).Trim()
            Check ($gotJob -eq $wantJob) "room $pick : selected the record $wantJob (got '$gotJob')"
            Check ($sel.Rows.Count -eq 1) "room $pick : the whole table row is selected"
            $vis = $xl.ActiveWindow.VisibleRange
            Check (($sel.Row -ge $vis.Row) -and ($sel.Row -le ($vis.Row + $vis.Rows.Count - 1))) "room $pick : the row is in view"
        }
    }
    # Nothing sensible selected: stays put and says so (quiet mode logs it).
    $rep.Activate()
    $xl.Run('SetQuiet', $true)
    Invoke-ComRetry { $rep.Range('A5').Select() | Out-Null }
    $xl.Run('btnGoToRecord')
    Check ($wb.ActiveSheet.Name -eq 'Reports') 'a cell above the results does not leave the Reports sheet'
    Check (([string]$xl.Run('QuietLog')) -like '*Select a record*') 'and asks for a record to be selected'

    # ------------------------------------------------------- button placement
    Write-Host ''
    Write-Host '=== Reports buttons stay in place as columns change ==='
    $names = 'btnClearFilters','btnTogglePaid','btnGoToRecord','btnMarkPaid','btnMarkUnpaid','btnExportReport','btnDeleteVisible'
    function Lbl([string]$n) {
        foreach ($s in $rep.Shapes) { if ($s.Name -eq $n) { return $s } }
        return $null
    }
    foreach ($n in $names) { Check ($null -ne (Shape $n)) "button $n exists" }
    Check ($null -ne (Lbl 'pcb_lblSelected')) "the 'Selected record:' caption exists"
    Check ($null -ne (Lbl 'pcb_lblMarkAll')) "the 'Mark all as...' caption exists"
    $before = @{}; foreach ($n in $names) { $before[$n] = [double](Shape $n).Left }
    $lblBefore = @{}; foreach ($n in 'pcb_lblSelected','pcb_lblMarkAll') { $lblBefore[$n] = [double](Lbl $n).Left }
    function Overlaps {
        $a = Shape 'btnClearFilters'; $b = Shape 'btnTogglePaid'; $m = Shape 'btnMarkPaid'; $e = Shape 'btnExportReport'
        return ((($a.Left + $a.Width) -gt ($b.Left + 0.5)) -or (($b.Left + $b.Width) -gt ($m.Left + 0.5)) -or (($m.Left + $m.Width) -gt ($e.Left + 0.5)))
    }
    function CaptionsOk {
        $d1 = [math]::Abs([double](Lbl 'pcb_lblSelected').Left - [double](Shape 'btnTogglePaid').Left)
        $d2 = [math]::Abs([double](Lbl 'pcb_lblMarkAll').Left - [double](Shape 'btnMarkPaid').Left)
        return (($d1 -lt 1.5) -and ($d2 -lt 1.5))
    }
    Check (-not (Overlaps)) 'default columns: the four blocks do not overlap'
    Check (CaptionsOk) 'default columns: each caption sits over its buttons'
    Check ([double](Shape 'btnGoToRecord').Left -eq [double](Shape 'btnTogglePaid').Left) 'Go to record is directly under Toggle Paid'
    Check ([double](Shape 'btnMarkUnpaid').Left -eq [double](Shape 'btnMarkPaid').Left) 'Unpaid is directly under Paid'
    $rep.Activate()
    $rep.Columns('H').Hidden = $true
    Invoke-ComRetry { $rep.Range('A5').Select() | Out-Null }
    Invoke-ComRetry { $rep.Range('A6').Select() | Out-Null }
    $cl = [double](Shape 'btnClearFilters').Left
    Check ($cl -ge ($rep.Columns('I').Left - 0.5)) "H hidden: the strip moved off the hidden column (Left $([int]$cl))"
    Check (-not (Overlaps)) 'H hidden: the blocks do not overlap'
    Check (CaptionsOk) 'H hidden: each caption still sits over its buttons'
    $rep.Columns('H').Hidden = $false
    Invoke-ComRetry { $rep.Range('A5').Select() | Out-Null }
    Invoke-ComRetry { $rep.Range('A6').Select() | Out-Null }
    $same = $true
    foreach ($n in $names) { if ([math]::Abs([double](Shape $n).Left - $before[$n]) -gt 1) { $same = $false; Write-Host "    $n at $([int](Shape $n).Left), was $([int]$before[$n])" } }
    foreach ($n in 'pcb_lblSelected','pcb_lblMarkAll') { if ([math]::Abs([double](Lbl $n).Left - $lblBefore[$n]) -gt 1) { $same = $false; Write-Host "    $n at $([int](Lbl $n).Left), was $([int]$lblBefore[$n])" } }
    Check $same 'N shown again: every button and caption is back where it was'

    # ------------------------------------------------- button names / macros
    Write-Host ''
    Write-Host '=== Button names agree with their macros ==='
    $bad = @()
    foreach ($ws in $wb.Worksheets) {
        foreach ($s in $ws.Shapes) {
            if ($s.Name -like 'pcb_*' -and $s.Type -eq 8 -and $s.FormControlType -eq 0) {   # buttons only (the view dropdown has no macro suffix)
                $macro = ([string]$s.OnAction); $macro = $macro.Substring($macro.IndexOf('!') + 1)
                if ($s.Name.Length -gt 31 -or $s.Name -notlike "pcb_${macro}_*") { $bad += "$($ws.Name): $($s.Name) -> $macro" }
            }
        }
    }
    Check ($bad.Count -eq 0) "every button name is <= 31 characters and holds its macro $($bad -join '; ')"
    $del = Shape 'btnDeleteVisible'
    Check ($null -ne $del -and ([string]$del.OnAction) -like '*btnDeleteVisible') 'Delete visible records runs btnDeleteVisible'
    $del.OnAction = 'OlderName.xlsm!btnDeleteVisible'
    $xl.Run('HealButtons')
    $healed = [string](Shape 'btnDeleteVisible').OnAction
    Check ($healed -notlike 'OlderName*' -and $healed -like '*btnDeleteVisible') "after a file rename HealButtons restores the macro (got '$healed')"

    # --------------------------------------- Toggle Paid under a Paid = No view
    Write-Host ''
    Write-Host '=== Toggle Paid while filtered to Paid = No ==='
    $paidCol = RepCol 'Paid'
    Clear-AllBoxes
    foreach ($lo in @($loMain, $loAnnexe)) { $lo.ListColumns('Paid').DataBodyRange.Value2 = 'No' }
    $rep.Range('B12').Value2 = 'No'
    $unpaid = Jobs
    Check ($unpaid -ge 5) "filtered to Paid = No: $unpaid records"
    $rep.Activate()
    Invoke-ComRetry { $rep.Cells($FIRST, $paidCol).Select() | Out-Null }
    $xl.Run('btnTogglePaid')
    $xl.CalculateFullRebuild()
    Check (SpillOk) 'the report is not left showing #SPILL!'
    Check ((Jobs) -eq ($unpaid - 1)) 'the record that became Yes drops out of the Paid = No view'
    $yes = 0; foreach ($lo in @($loMain, $loAnnexe)) { $yes += @($lo.ListColumns('Paid').DataBodyRange.Value2 | Where-Object { ([string]$_).Trim() -eq 'Yes' }).Count }
    Check ($yes -eq 1) "exactly one record is Yes in the data (found $yes)"
    Clear-AllBoxes
    Check ((Jobs) -eq $total) 'clearing the filter shows every record again'}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($anyFail) { Write-Host 'FAIL: see above'; exit 1 }
Write-Host 'OK: Reports navigation, filters and buttons'
