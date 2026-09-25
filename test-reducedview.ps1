# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Snag 1e: workbook-wide reduced-clutter view toggle on location sheets.
#
#   - Toggling hides/shows Status, Job ID, Printer, Area m2, Disregard Paper,
#     Disregard Consumable on EVERY location sheet at once (one workbook-
#     wide setting, not per-sheet).
#   - The hidden-column list lives in a setting (SET_LOC_REDUCED_COLUMNS),
#     editable without a rebuild.
#   - A permanently-hidden column (H_Issues) is never touched by this - it
#     must stay hidden regardless of which state the toggle is in.
#   - The button caption flips between "Reduce clutter" and "Show all
#     columns" on every location sheet at once.
#
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself. Closes WITHOUT
# saving.

$ErrorActionPreference = 'Stop'
$deliverable = Join-Path $PSScriptRoot 'src\PrintCosts.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintCosts.xlsm'
Copy-Item $deliverable $f

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Main Print Room')
    $annex = $wb.Worksheets('Annexe')
    $lo = $main.ListObjects('tblJobs_MAIN')

    function Col($lo, $name) {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            if ($lo.ListColumns($i).Name -eq $name) { return $i }
        }
        return 0
    }
    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }
    function IsHidden($ws, $lo, $header) {
        $c = $lo.Range.Column + (Col $lo $header) - 1
        return [bool]$ws.Columns($c).Hidden
    }

    # -------------------------------------------------------------- setting
    Write-Host '=== Settings self-provisioned ==='
    $onOff = [string]$wb.Names.Item('SET_LOC_REDUCED_VIEW').RefersToRange.Text
    $list = [string]$wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Text
    Check ($onOff -eq 'No') "SET_LOC_REDUCED_VIEW defaults to No (got '$onOff')"
    Check ($list -like '*Status*Job ID*Printer*Area m2*Disregard Paper*Disregard Consumable*') "SET_LOC_REDUCED_COLUMNS holds the expected shortlist (got '$list')"

    # ---------------------------------------------------------- initial state
    Write-Host ''
    Write-Host '=== Full view initially: nothing in the shortlist is hidden ==='
    foreach ($h in 'Status', 'Job ID', 'Printer', 'Area m2', 'Disregard Paper', 'Disregard Consumable') {
        Check (-not (IsHidden $main $lo $h)) "$h is visible on Main Print Room"
    }
    $hIssuesHiddenBefore = IsHidden $main $lo 'H_Issues'
    Check $hIssuesHiddenBefore "H_Issues (permanently hidden, unrelated to this toggle) starts hidden"

    # -------------------------------------------------------------- toggle on
    Write-Host ''
    Write-Host '=== Toggling ON hides the shortlist on every location sheet ==='
    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    $aloAnnex = $annex.ListObjects('tblJobs_ANNEX')
    foreach ($h in 'Status', 'Job ID', 'Printer', 'Area m2', 'Disregard Paper', 'Disregard Consumable') {
        Check (IsHidden $main $lo $h) "$h is hidden on Main Print Room"
        Check (IsHidden $annex $aloAnnex $h) "$h is hidden on Annexe too (workbook-wide, not per-sheet)"
    }
    Check (-not (IsHidden $main $lo 'Chargeable Cost')) "Chargeable Cost (not in the shortlist) stays visible"
    Check (-not (IsHidden $main $lo 'Paid')) "Paid (not in the shortlist) stays visible"
    Check ((IsHidden $main $lo 'H_Issues') -eq $hIssuesHiddenBefore) "H_Issues' state is untouched by this toggle"

    $onOffAfter = [string]$wb.Names.Item('SET_LOC_REDUCED_VIEW').RefersToRange.Text
    Check ($onOffAfter -eq 'Yes') "SET_LOC_REDUCED_VIEW now reads Yes"

    $prefix = 'pcb_btnToggleReducedView'
    function FindButtonCaption($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [string]$b.Caption }
        }
        return $null
    }
    $btnCaption = FindButtonCaption $main $prefix
    Check ($btnCaption -eq 'Show all columns') "button caption flipped to 'Show all columns' (got '$btnCaption')"

    # ------------------------------------------------------------- toggle off
    Write-Host ''
    Write-Host '=== Toggling OFF shows the shortlist again ==='
    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    foreach ($h in 'Status', 'Job ID', 'Printer', 'Area m2', 'Disregard Paper', 'Disregard Consumable') {
        Check (-not (IsHidden $main $lo $h)) "$h is visible again on Main Print Room"
    }
    Check ((IsHidden $main $lo 'H_Issues') -eq $hIssuesHiddenBefore) "H_Issues' state is STILL untouched (stays permanently hidden)"
    $btnCaption2 = FindButtonCaption $main $prefix
    Check ($btnCaption2 -eq 'Reduce clutter') "button caption flipped back to 'Reduce clutter' (got '$btnCaption2')"

    # ------------------------------------------- button relocation (2026-09-25 fix)
    # A button anchored over a column the reduced view hides used to vanish
    # with it (Remove Row/Clear All sat on Printer/Disregard Consumable).
    # Fix: occasional-use buttons, including the toggle itself (moved here
    # same day - not used often enough to earn a spot near the table), now
    # live in a side panel (col 36, Import paired at col 48 - see
    # SIDE_PANEL_COL2's comment for why) immune by construction; the few
    # still inside the table (Add Print Job, Now, Clear defaults) self-
    # relocate off whatever column is currently hidden. REDUCED_COLUMNS_
    # DEFAULT alone never hides any of those three, so this needs its own
    # list to actually exercise the relocation path.
    Write-Host ''
    Write-Host '=== Buttons anchored inside the table self-relocate off a hidden column ==='

    function ButtonColumn($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") {
                for ($c = 1; $c -le 60; $c++) {
                    if ([Math]::Abs($ws.Cells(1, $c).Left - $b.Left) -lt 0.5) { return $c }
                }
                return -1
            }
        }
        return 0
    }

    $techCol = $lo.Range.Column + (Col $lo 'Technician') - 1

    foreach ($p in 'pcb_btnRemoveRow', 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnToggleReducedView') {
        Check ((ButtonColumn $main $p) -eq 36) "$p sits in the side panel (column 36), immune to column-hide"
    }
    Check ((ButtonColumn $main 'pcb_btnImportLocation') -eq 48) "pcb_btnImportLocation sits in the side panel (column 48, paired with Export), immune to column-hide"
    Check ((ButtonColumn $main 'pcb_btnClearDefaults') -eq $techCol) "btnClearDefaults starts anchored on Technician"

    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = 'Technician'
    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    Check (IsHidden $main $lo 'Technician') "Technician is hidden (test-only hide list)"
    $movedCol = ButtonColumn $main 'pcb_btnClearDefaults'
    Check ($movedCol -ne $techCol) "btnClearDefaults relocated off Technician (was col $techCol, now $movedCol)"
    Check ($movedCol -gt 0 -and -not [bool]$main.Columns($movedCol).Hidden) "btnClearDefaults' new column ($movedCol) is actually visible"
    foreach ($p in 'pcb_btnRemoveRow', 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnToggleReducedView') {
        Check ((ButtonColumn $main $p) -eq 36) "$p still in the side panel, unaffected by the test-only hide list"
    }
    Check ((ButtonColumn $main 'pcb_btnImportLocation') -eq 48) "pcb_btnImportLocation still in the side panel, unaffected by the test-only hide list"

    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    Check (-not (IsHidden $main $lo 'Technician')) "Technician is visible again"
    Check ((ButtonColumn $main 'pcb_btnClearDefaults') -eq $techCol) "btnClearDefaults moved back onto Technician now that it's visible"

    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = $list

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
