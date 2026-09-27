# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Column grouping cleanup (2026-09-22, follow-up to the reduced-clutter view
# work). Originally, only the cost-column group (Paper Cost..Disregarded)
# was meant to carry an outline level above 1; two other groups had crept in
# as a side effect of EnsurePaidColumn inserting Paid right next to the
# already-grouped H_Issues/snapshot block PrintCosts.xlsx ships:
#
#   - Notes + H_Issues ended up grouped together, and NOTES ITSELF was
#     wrongly hidden - a real bug (Notes is a genuine input column, never
#     meant to be hidden).
#   - S_PrinterID..S_StampedBy ended up double-grouped (nested a level
#     deeper than the shipped state), with S_SchemaVer oddly excluded from
#     that extra nesting.
#
# Fixed by modInit.FlattenOutline: Notes through S_SchemaVer are reset to
# outline level 1 outright (no re-grouping), and Notes is explicitly
# unhidden. H_Issues/the snapshot columns stay invisible via their own
# .Hidden state, which needs no outline group to hold it.
#
# 2026-09-27: the cost-column group itself was ALSO replaced by a plain
# .Hidden toggle (modInit.ToggleCostColumns/ApplyCostColumnsVisibility,
# docs/ARCHITECTURE.md §16.3's "button lag" writeup) - a native outline
# group has no VBA event to hook, so buttons anchored beside it lagged by
# one click after every collapse/expand. modInit.NormalizeJobColumnOutlines
# (renamed from GroupJobColumns) now flattens that span too, via the same
# FlattenOutline this test already exercises for Notes..S_SchemaVer. So the
# whole job table should now carry NO outline group anywhere - see
# test-costcolumns.ps1 for the toggle's own .Hidden behaviour.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT
# saving.

$ErrorActionPreference = 'Stop'
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
    $lo = $main.ListObjects('tblJobs_MAIN')

    function Col($lo, $name) {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            if ($lo.ListColumns($i).Name -eq $name) { return $i }
        }
        return 0
    }
    function SheetCol($header) { return $lo.Range.Column + (Col $lo $header) - 1 }
    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }

    Write-Host '=== The cost columns carry no outline group any more (2026-09-27) ==='
    foreach ($h in 'Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded', 'Chargeable Cost', 'Paid') {
        $c = $main.Columns((SheetCol $h))
        Check ($c.OutlineLevel -eq 1) "$h is at outline level 1, no group (got $($c.OutlineLevel))"
    }

    Write-Host ''
    Write-Host '=== Notes: not grouped, not hidden (was a real bug) ==='
    $notesCol = $main.Columns((SheetCol 'Notes'))
    Check ($notesCol.OutlineLevel -eq 1) "Notes is at outline level 1 (got $($notesCol.OutlineLevel))"
    Check (-not [bool]$notesCol.Hidden) "Notes is NOT hidden (got hidden=$([bool]$notesCol.Hidden))"

    Write-Host ''
    Write-Host '=== H_Issues / snapshot block: ungrouped but still invisible ==='
    foreach ($h in 'H_Issues', 'S_PrinterID', 'S_StampedBy') {
        $c = $main.Columns((SheetCol $h))
        Check ($c.OutlineLevel -eq 1) "$h is at outline level 1, no group (got $($c.OutlineLevel))"
        Check ([bool]$c.Hidden) "$h is still hidden via .Hidden, not a group (got hidden=$([bool]$c.Hidden))"
    }

    Write-Host ''
    Write-Host '=== No outline group anywhere in the table (2026-09-27: the last one, cost columns, was removed) ==='
    $maxLevel = 0
    for ($ci = 1; $ci -le $lo.ListColumns.Count; $ci++) {
        $sc = $lo.Range.Column + $ci - 1
        $lvl = $main.Columns($sc).OutlineLevel
        if ($lvl -gt $maxLevel) { $maxLevel = $lvl }
    }
    Check ($maxLevel -eq 1) "max outline level across the whole table is 1 (got $maxLevel)"

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
