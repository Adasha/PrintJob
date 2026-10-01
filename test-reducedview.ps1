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
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT
# saving.

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

    function IsHidden($ws, $lo, $header) {
        return [bool](Invoke-ComRetry -Attempts 5 {
            $c = $lo.Range.Column + (Col $lo $header) - 1
            if ($c -lt 1) { throw "column lookup failed for '$header'" }
            $ws.Columns($c).Hidden
        })
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
        Check (-not (IsHidden $main $lo $h)) "$h is visible on Example Print Room"
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
        Check (IsHidden $main $lo $h) "$h is hidden on Example Print Room"
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
        Check (-not (IsHidden $main $lo $h)) "$h is visible again on Example Print Room"
    }
    Check ((IsHidden $main $lo 'H_Issues') -eq $hIssuesHiddenBefore) "H_Issues' state is STILL untouched (stays permanently hidden)"
    $btnCaption2 = FindButtonCaption $main $prefix
    Check ($btnCaption2 -eq 'Reduce clutter') "button caption flipped back to 'Reduce clutter' (got '$btnCaption2')"

    # ------------------------------------------- button relocation (2026-09-25 fix)
    # A button anchored over a column the reduced view hides used to vanish
    # with it (Remove Row/Clear All sat on Printer/Disregard Consumable).
    # Fix: occasional-use buttons, including the toggle itself (moved here
    # same day - not used often enough to earn a spot near the table), now
    # live in a side panel, all seven sharing column 36 (Import briefly
    # lived at a second column, paired with Export, until user-reported
    # feedback that it had visibly drifted away from the rest of the stack
    # - DrawOneAtTop packs all seven into one column at even pixel steps
    # instead) immune by construction; the few still inside the table (Add
    # Print Job, Now, Clear defaults) self-relocate off whatever column is
    # currently hidden. REDUCED_COLUMNS_DEFAULT alone never hides any of
    # those three, so this needs its own list to actually exercise the
    # relocation path.
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

    # Lowest-index match only, unlike ButtonColumn above: hiding a column
    # collapses it to zero width, so the FIRST visible column after it slides
    # left to share its exact pixel Left - the hidden column and the visible
    # one it now touches are indistinguishable by position alone. A button
    # relocated onto that visible neighbour is correctly placed, but
    # ButtonColumn would report the hidden column instead, since it always
    # returns the lowest column index at that Left. This variant returns the
    # first VISIBLE column at that Left instead, which is the one the
    # relocation logic actually cares about.
    function ButtonColumnVisible($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") {
                for ($c = 1; $c -le 60; $c++) {
                    if ([Math]::Abs($ws.Cells(1, $c).Left - $b.Left) -lt 0.5 -and -not [bool]$ws.Columns($c).Hidden) { return $c }
                }
                return -1
            }
        }
        return 0
    }

    $techCol = $lo.Range.Column + (Col $lo 'Technician') - 1

    foreach ($p in 'pcb_btnRemoveRow', 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnToggleReducedView', 'pcb_btnImportLocation') {
        Check ((ButtonColumn $main $p) -eq 36) "$p sits in the side panel (column 36), immune to column-hide"
    }
    Check ((ButtonColumn $main 'pcb_btnClearDefaults') -eq $techCol) "btnClearDefaults starts anchored on Technician"

    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = 'Technician'
    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    Check (IsHidden $main $lo 'Technician') "Technician is hidden (test-only hide list)"
    $movedCol = ButtonColumnVisible $main 'pcb_btnClearDefaults'
    Check ($movedCol -gt 0) "btnClearDefaults relocated to a visible column (found col $movedCol; was anchored on Technician, col $techCol, now hidden)"
    Check (-not [bool]$main.Columns($movedCol).Hidden) "btnClearDefaults' new column ($movedCol) is actually visible"
    foreach ($p in 'pcb_btnRemoveRow', 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnToggleReducedView', 'pcb_btnImportLocation') {
        Check ((ButtonColumn $main $p) -eq 36) "$p still in the side panel, unaffected by the test-only hide list"
    }

    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    Check (-not (IsHidden $main $lo 'Technician')) "Technician is visible again"
    Check ((ButtonColumn $main 'pcb_btnClearDefaults') -eq $techCol) "btnClearDefaults moved back onto Technician now that it's visible"

    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = $list

    # ------------------------------------- button width survives hide/reveal (2026-09-26 fix)
    # User-reported: after hiding then revealing columns once, the Now button
    # ended up straddling the D/E column border with only its "N" visible.
    # Root cause was Buttons.Add defaulting to Placement:=xlMoveAndSize, so
    # Excel silently shrank the button's own Width when Printer (its anchor
    # column, and one of REDUCED_COLUMNS_DEFAULT's own entries) got hidden -
    # RelocateButton only ever restored .Left, never .Width, so the shrunk
    # width stuck even once the button landed back on a visible column. The
    # position-only ButtonColumn check above would not have caught this
    # (Width can be wrong while Left still happens to land on a cell edge),
    # so this checks Width explicitly, across two full toggle cycles to
    # mirror the reported "hide then reveal" sequence.
    Write-Host ''
    Write-Host '=== Now/Add Print Job/Repeat Job keep their width across repeated hide/reveal cycles ==='

    function ButtonWidth($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [double]$b.Width }
        }
        return -1
    }
    function ButtonPlacement($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [int]$b.Placement }
        }
        return -1
    }

    $atRisk = 'pcb_btnAddPrintJob', 'pcb_btnRepeatJob', 'pcb_btnNow', 'pcb_btnClearDefaults'
    for ($cycle = 1; $cycle -le 2; $cycle++) {
        [void]$xl.Run('ToggleReducedView')
        Start-Sleep -Milliseconds 300
        [void]$xl.Run('ToggleReducedView')
        Start-Sleep -Milliseconds 300
        foreach ($p in $atRisk) {
            Check ((ButtonWidth $main $p) -eq 110) "$p is still 110pt wide after hide/reveal cycle $cycle"
            Check ((ButtonPlacement $main $p) -eq 3) "$p is xlFreeFloating (3) after hide/reveal cycle $cycle"
        }
    }
    Check ((ButtonColumn $main 'pcb_btnNow') -gt 0) "btnNow's Left still lands cleanly on a column edge, not straddling a border"

    # ---------------------------------- side panel tracks column 36 through toggles (2026-09-26 fix)
    # User-reported, same session: once the at-risk buttons above stopped
    # drifting, the side panel (Select printers.../Check this sheet/the
    # toggle itself/Remove Row/Clear All/Export.../Import...) turned out to
    # have its own version of the same bug - hiding the reduced-view
    # shortlist shrinks the total width of everything to the LEFT of
    # SIDE_PANEL_COL (column 36), which shifts that column's own pixel
    # position left. With Placement:=xlFreeFloating (this fix's own change)
    # the side panel no longer follows that shift automatically, so it ended
    # up sitting further right of column 36 than before - the opposite
    # direction from a naive guess, but exactly what "hiding clutter leaves
    # them even further right" describes. ApplyReducedView now calls
    # RepositionSidePanelButtons on every run (not just from setup) to keep
    # them pinned to column 36's current position.
    Write-Host ''
    Write-Host '=== Side panel stays pinned to column 36 through a hide/reveal cycle ==='

    function SidePanelLeft($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [double]$b.Left }
        }
        return -1
    }

    $sidePanel = 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnToggleReducedView', 'pcb_btnRemoveRow', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation'
    $col36Before = [double]$main.Cells(1, 36).Left
    foreach ($p in $sidePanel) {
        Check ((SidePanelLeft $main $p) -eq $col36Before) "$p sits on column 36 before hiding (Left=$col36Before)"
    }

    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    $col36Hidden = [double]$main.Cells(1, 36).Left
    Check ($col36Hidden -ne $col36Before) "hiding the shortlist actually moved column 36 (was $col36Before, now $col36Hidden) - otherwise this check proves nothing"
    foreach ($p in $sidePanel) {
        Check ((SidePanelLeft $main $p) -eq $col36Hidden) "$p followed column 36 to its new position while hidden (Left=$col36Hidden)"
    }

    [void]$xl.Run('ToggleReducedView')
    Start-Sleep -Milliseconds 300
    foreach ($p in $sidePanel) {
        Check ((SidePanelLeft $main $p) -eq $col36Before) "$p is back on column 36 after revealing (Left=$col36Before)"
    }

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
