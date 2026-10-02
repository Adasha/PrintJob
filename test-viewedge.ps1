# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Column-view edge cases (ARCHITECTURE.md §16.2 punch list). test-reducedview.ps1
# covers the views themselves; this covers the places they meet other features:
#
#   - Refresh Locations keeps the current view (Minimal stays Minimal).
#   - Add print room: the view is workbook-wide, so a new room takes the CURRENT
#     mode (not "All"), and follows later changes with every other room.
#   - Import into a sheet with hidden columns lands values in the hidden columns
#     and leaves them hidden.
#   - A Settings list naming a nonexistent header (or empty entries) fails soft:
#     the real headers in the list are still hidden and nothing throws.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
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
    $main = $wb.Worksheets('Example Print Room')
    $annex = $wb.Worksheets('Annexe')
    $lo = $main.ListObjects('tblJobs_MAIN')
    $aloAnnex = $annex.ListObjects('tblJobs_ANNEX')

    function IsHidden($ws, $lo, $header) {
        return [bool](Invoke-ComRetry -Attempts 5 {
            $c = $lo.Range.Column + (Col $lo $header) - 1
            if ($c -lt 1) { throw "column lookup failed for '$header'" }
            $ws.Columns($c).Hidden
        })
    }
    function FindDropDown($ws, $prefix) {
        $n = $ws.DropDowns().Count
        for ($i = 1; $i -le $n; $i++) {
            $d = $ws.DropDowns($i)
            if ([string]$d.Name -like "$prefix*") { return $d }
        }
        return $null
    }
    function ViewSetting { [string]$wb.Names.Item('SET_LOC_REDUCED_VIEW').RefersToRange.Text }
    function Settle { Start-Sleep -Milliseconds 300 }

    $reducedCols = 'Status', 'Job ID', 'Area m2'
    $minimalExtra = 'Printer', 'Disregard Paper', 'Disregard Consumable', 'Print Width mm', 'Sheet size'
    $allListed = $reducedCols + $minimalExtra

    # ------------------------------------------------- Refresh Locations
    Write-Host '=== Refresh Locations keeps Minimal ==='
    [void]$xl.Run('SetViewMode', 'Minimal'); Settle
    $xl.Run('btnRefreshLocations')
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Start-Sleep -Seconds 1
    Check ((ViewSetting) -eq 'Minimal') 'setting still reads Minimal after Refresh Locations'
    foreach ($h in $allListed) {
        Check (IsHidden $main $lo $h) "$h still hidden on Example Print Room"
        Check (IsHidden $annex $aloAnnex $h) "$h still hidden on Annexe"
    }
    Check ([int](FindDropDown $main 'pcb_ddViewMode').ListIndex -eq 3) 'drop-down still shows Minimal'

    # ------------------------------------------------------- Add print room
    Write-Host ''
    Write-Host '=== Add print room while Minimal: new room takes the current view ==='
    $newWs = $xl.Run('CreatePrintRoom', 'View Edge Room', 'Test', 'VEDGE')
    Settle
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    $newWs = $wb.Worksheets('View Edge Room')
    $newLo = $null
    foreach ($t in $newWs.ListObjects) { if ($t.Name -like 'tblJobs_*') { $newLo = $t } }
    Check ($null -ne $newLo) 'new room has a jobs table'
    foreach ($h in $allListed) {
        Check (IsHidden $newWs $newLo $h) "$h hidden on the new room (inherits Minimal)"
    }
    Check ((ViewSetting) -eq 'Minimal') 'creating a room did not reset the workbook view'
    Check ([int](FindDropDown $newWs 'pcb_ddViewMode').ListIndex -eq 3) "new room's drop-down shows Minimal"

    [void]$xl.Run('SetViewMode', 'All'); Settle
    foreach ($h in $allListed) {
        Check (-not (IsHidden $newWs $newLo $h)) "$h visible on the new room after switching to All"
    }

    Write-Host ''
    Write-Host '=== Add print room while All: new room starts with everything visible ==='
    $null = $xl.Run('CreatePrintRoom', 'View Edge Two', '', 'VEDG2')
    Settle
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    $ws2 = $wb.Worksheets('View Edge Two')
    $lo2 = $null
    foreach ($t in $ws2.ListObjects) { if ($t.Name -like 'tblJobs_*') { $lo2 = $t } }
    foreach ($h in $allListed) {
        Check (-not (IsHidden $ws2 $lo2 $h)) "$h visible on a room created in All"
    }
    Check ([int](FindDropDown $ws2 'pcb_ddViewMode').ListIndex -eq 1) "new room's drop-down shows All"

    # ------------------------------------------- Import into hidden columns
    Write-Host ''
    Write-Host '=== Import into a sheet whose columns are hidden ==='
    $xl.Run('btnExportAll')
    $mainCsv = Get-ChildItem $workDir -Filter 'PrintCosts-*-MAIN-*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Check ($null -ne $mainCsv) 'export produced a CSV for Example Print Room'

    [void]$xl.Run('SetViewMode', 'Minimal'); Settle
    $printerCol = Col $aloAnnex 'Printer'
    $idCol = Col $aloAnnex 'Job ID'
    # Annexe is a copy of Example Print Room, so its Job IDs already match the
    # export (import skips known IDs): remove two rows so there is something to bring back.
    $annex.Unprotect()
    $aloAnnex.ListRows(2).Delete()
    $aloAnnex.ListRows(1).Delete()
    $annex.Protect()
    Settle
    $before = $aloAnnex.ListRows.Count
    $rows = $xl.Run('ReadImportRows', $mainCsv.FullName)
    $xl.Run('ApplyImportConfirmed', $annex, $rows)
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Settle
    $after = $aloAnnex.ListRows.Count
    Check ($after -ge ($before + 2)) "the two deleted rows were imported back ($before -> $after)"
    $filled = 0; $ids = 0
    for ($i = 1; $i -le $aloAnnex.ListRows.Count; $i++) {
        $cells = $aloAnnex.ListRows($i).Range.Cells
        if ([string]$cells.Item(1, $printerCol).Value2 -ne '') { $filled++ }
        if ([string]$cells.Item(1, $idCol).Value2 -ne '') { $ids++ }
    }
    Check ($filled -ge ($after - $before)) "Printer (hidden) holds a value on imported rows ($filled filled)"
    Check ($ids -ge ($after - $before)) "Job ID (hidden) holds a value on imported rows ($ids filled)"
    foreach ($h in $allListed) {
        Check (IsHidden $annex $aloAnnex $h) "$h still hidden on Annexe after import"
    }
    Check ((ViewSetting) -eq 'Minimal') 'import did not change the workbook view'
    [void]$xl.Run('SetViewMode', 'All'); Settle

    # ------------------------------------- full setup pass in a hidden view
    Write-Host ''
    Write-Host '=== InitialiseWorkbook (setup) in Minimal keeps the columns hidden ==='
    [void]$xl.Run('SetViewMode', 'Minimal'); Settle
    $xl.Run('InitialiseWorkbook')
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Start-Sleep -Seconds 1
    foreach ($h in $allListed) {
        Check (IsHidden $main $lo $h) "$h still hidden on Example Print Room after setup"
    }
    [void]$xl.Run('SetViewMode', 'All'); Settle

    # ------------------------------------------------- bad Settings entries
    Write-Host ''
    Write-Host '=== A Settings list naming a nonexistent header fails soft ==='
    $origList = [string]$wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Text
    $origExtra = [string]$wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Text
    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = 'Status; No Such Column ;;Job ID;'
    $wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Value = 'Bogus Header;Printer'
    $ok = $true
    try { [void]$xl.Run('SetViewMode', 'Reduced') } catch { $ok = $false; Write-Host "  $_" }
    Settle
    Check $ok 'Reduced with a bad header in the list does not raise'
    Check ((IsHidden $main $lo 'Status') -and (IsHidden $main $lo 'Job ID')) 'the real headers in the list are still hidden (bad entry skipped, not fatal)'
    Check (-not (IsHidden $main $lo 'Area m2')) 'a column not in the edited list stays visible'
    $ok = $true
    try { [void]$xl.Run('SetViewMode', 'Minimal') } catch { $ok = $false; Write-Host "  $_" }
    Settle
    Check $ok 'Minimal with a bad header in the extra list does not raise'
    Check (IsHidden $main $lo 'Printer') 'Printer hidden in Minimal despite Bogus Header beside it'
    Check ((ViewSetting) -eq 'Minimal') 'setting reads Minimal'
    $xl.Run('btnRefreshLocations'); Settle
    Check ((IsHidden $main $lo 'Printer') -and (IsHidden $main $lo 'Status')) 'Refresh Locations with a bad list still applies the view'

    Write-Host ''
    Write-Host '=== Blank lists fall back to the defaults ==='
    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = ''
    $wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Value = ''
    [void]$xl.Run('SetViewMode', 'All'); Settle
    [void]$xl.Run('SetViewMode', 'Minimal'); Settle
    foreach ($h in $allListed) {
        Check (IsHidden $main $lo $h) "$h hidden in Minimal from the built-in default lists"
    }

    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = $origList
    $wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Value = $origExtra
    [void]$xl.Run('SetViewMode', 'All')
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
