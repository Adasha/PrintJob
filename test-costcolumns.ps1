# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Cost-columns toggle (docs/ARCHITECTURE.md §16.3, "button lag" writeup,
# 2026-09-27) - replaces the native Excel outline +/- control that used to
# sit over Paper Cost..Disregarded.
#
#   - A workbook-wide toggle (like reduced-clutter view, not per-sheet) hides
#     Paper Cost, Consumable Cost, Gross Cost and Disregarded on every
#     location sheet at once. Chargeable Cost and Paid are never hidden by
#     this - same guarantee the old outline group used to make.
#   - The button caption flips between "Hide cost detail" and "Show cost
#     detail" on every location sheet at once.
#   - THE ACTUAL BUG FIX: buttons anchored beside this column span reposition
#     in the SAME click that hides/shows the columns - no separate "next
#     click anywhere on the sheet" needed the way the old native outline
#     control required (RepositionLocationButtons only ran on
#     Workbook_SheetActivate/Workbook_SheetSelectionChange, proxies for an
#     event VBA doesn't have for an outline collapse/expand).
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
    $main = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.Run('SetQuiet', $true)
    $annex = $wb.Worksheets('Annexe')
    $lo = $main.ListObjects('tblJobs_MAIN')

    function IsHidden($ws, $lo, $header) {
        $c = $lo.Range.Column + (Col $lo $header) - 1
        return [bool]$ws.Columns($c).Hidden
    }
    function FindButtonCaption($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [string]$b.Caption }
        }
        return $null
    }
    function SidePanelLeft($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [double]$b.Left }
        }
        return -1
    }

    # -------------------------------------------------------------- setting
    Write-Host '=== Setting self-provisioned ==='
    $onOff = [string]$wb.Names.Item('SET_COST_COLS_HIDDEN').RefersToRange.Text
    Check ($onOff -eq 'No') "SET_COST_COLS_HIDDEN defaults to No (got '$onOff')"

    # ---------------------------------------------------------- initial state
    Write-Host ''
    Write-Host '=== Initially visible: nothing in the cost-detail span is hidden ==='
    foreach ($h in 'Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded') {
        Check (-not (IsHidden $main $lo $h)) "$h is visible on Example Print Room"
    }
    Check (-not (IsHidden $main $lo 'Chargeable Cost')) "Chargeable Cost is visible"
    Check (-not (IsHidden $main $lo 'Paid')) "Paid is visible"

    $col36Before = [double]$main.Cells(1, 36).Left
    $sidePanel = 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnToggleCostColumns', 'pcb_btnRemoveRow', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation'
    foreach ($p in $sidePanel) {
        Check ((SidePanelLeft $main $p) -eq $col36Before) "$p sits on column 36 before hiding (Left=$col36Before)"
    }

    # -------------------------------------------------------------- toggle on
    Write-Host ''
    Write-Host '=== Toggling ON hides the cost-detail span on every location sheet ==='
    $aloAnnex = $annex.ListObjects('tblJobs_ANNEX')
    [void]$xl.Run('btnToggleCostColumns')
    # Deliberately NO Start-Sleep before the checks below that matter most -
    # the whole point of this fix is that repositioning happens in the SAME
    # macro call, not on a subsequent click/sheet-activate. A short sleep
    # here would let a hypothetical regression back to the old lag pass
    # unnoticed by giving Workbook_SheetActivate/SelectionChange time to
    # paper over it - this test wants to catch exactly that gap reopening.
    #
    # A single throwaway probe read IS wrapped in Invoke-ComRetry, though -
    # not to wait for the fix (the button positions are already correct the
    # instant Run() returns, a synchronous call), but because toggling two
    # sheets' worth of columns/buttons is real work, and Excel can reject the
    # very next COM call outright while still settling from it (TestCommon.
    # ps1's own docstring). This retries OUR read of already-correct state,
    # not the macro call itself, so it does not mask a real lag regression.
    Invoke-ComRetry { $main.Columns(1).Hidden } | Out-Null

    foreach ($h in 'Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded') {
        Check (IsHidden $main $lo $h) "$h is hidden on Example Print Room"
        Check (IsHidden $annex $aloAnnex $h) "$h is hidden on Annexe too (workbook-wide, not per-sheet)"
    }
    Check (-not (IsHidden $main $lo 'Chargeable Cost')) "Chargeable Cost stays visible (never hidden by this toggle)"
    Check (-not (IsHidden $main $lo 'Paid')) "Paid stays visible (never hidden by this toggle)"

    $onOffAfter = [string]$wb.Names.Item('SET_COST_COLS_HIDDEN').RefersToRange.Text
    Check ($onOffAfter -eq 'Yes') "SET_COST_COLS_HIDDEN now reads Yes"

    $btnCaption = FindButtonCaption $main 'pcb_btnToggleCostColumns'
    Check ($btnCaption -eq 'Show cost detail') "button caption flipped to 'Show cost detail' (got '$btnCaption')"

    # ------------------------------------------ THE FIX: no lag, same click
    Write-Host ''
    Write-Host '=== No lag: side panel already repositioned in the SAME click that hid the columns ==='
    $col36Hidden = [double]$main.Cells(1, 36).Left
    Check ($col36Hidden -ne $col36Before) "hiding the span actually moved column 36 (was $col36Before, now $col36Hidden) - otherwise this check proves nothing"
    foreach ($p in $sidePanel) {
        Check ((SidePanelLeft $main $p) -eq $col36Hidden) "$p already sits on column 36's NEW position, no separate click needed (Left=$col36Hidden)"
    }

    # ------------------------------------------------------------- toggle off
    Write-Host ''
    Write-Host '=== Toggling OFF shows the span again, same click, no lag ==='
    [void]$xl.Run('btnToggleCostColumns')
    Invoke-ComRetry { $main.Columns(1).Hidden } | Out-Null
    foreach ($h in 'Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded') {
        Check (-not (IsHidden $main $lo $h)) "$h is visible again on Example Print Room"
    }
    $btnCaption2 = FindButtonCaption $main 'pcb_btnToggleCostColumns'
    Check ($btnCaption2 -eq 'Hide cost detail') "button caption flipped back to 'Hide cost detail' (got '$btnCaption2')"
    foreach ($p in $sidePanel) {
        Check ((SidePanelLeft $main $p) -eq $col36Before) "$p is back on column 36's original position, same click (Left=$col36Before)"
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
