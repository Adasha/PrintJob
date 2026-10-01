# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Column views on location sheets (snag 1e, reworked 2026-10-01): All /
# Reduced / Minimal buttons on row 2 (label in D2, buttons after it).
#
#   - One workbook-wide setting (SET_LOC_REDUCED_VIEW = All/Reduced/Minimal).
#   - Reduced hides SET_LOC_REDUCED_COLUMNS; Minimal hides those plus
#     SET_LOC_MINIMAL_COLUMNS. Both lists are editable settings.
#   - A permanently-hidden column (H_Issues) is never touched.
#   - The view buttons slide off any hidden column; the old toggle is gone from
#     the side panel and the remaining seven buttons stay on column 36.
#   - Add Print Job / Now / etc. keep their width and relocate off hidden columns.
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
    $aloAnnex = $annex.ListObjects('tblJobs_ANNEX')

    function IsHidden($ws, $lo, $header) {
        return [bool](Invoke-ComRetry -Attempts 5 {
            $c = $lo.Range.Column + (Col $lo $header) - 1
            if ($c -lt 1) { throw "column lookup failed for '$header'" }
            $ws.Columns($c).Hidden
        })
    }
    function FindButton($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return $b }
        }
        return $null
    }
    # First VISIBLE column whose Left matches the button's (a hidden column
    # shares its Left with the visible one after it).
    function ButtonColumnVisible($ws, $prefix) {
        $b = FindButton $ws $prefix
        if (-not $b) { return 0 }
        for ($c = 1; $c -le 60; $c++) {
            if ([Math]::Abs($ws.Cells(1, $c).Left - $b.Left) -lt 0.5 -and -not [bool]$ws.Columns($c).Hidden) { return $c }
        }
        return -1
    }
    function ButtonColumn($ws, $prefix) {
        $b = FindButton $ws $prefix
        if (-not $b) { return 0 }
        for ($c = 1; $c -le 60; $c++) {
            if ([Math]::Abs($ws.Cells(1, $c).Left - $b.Left) -lt 0.5) { return $c }
        }
        return -1
    }
    function ViewSetting { [string]$wb.Names.Item('SET_LOC_REDUCED_VIEW').RefersToRange.Text }

    $reducedCols = 'Status', 'Job ID', 'Area m2'
    $minimalExtra = 'Printer', 'Disregard Paper', 'Disregard Consumable', 'Print Width mm', 'Sheet size'

    # -------------------------------------------------------------- settings
    Write-Host '=== Settings self-provisioned ==='
    Check ((ViewSetting) -eq 'All') "SET_LOC_REDUCED_VIEW defaults to All (got '$(ViewSetting)')"
    $list = [string]$wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Text
    $extra = [string]$wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Text
    Check ($list -like '*Status*Job ID*Area m2*S_SchemaVer*') "SET_LOC_REDUCED_COLUMNS holds the reduced list (got '$list')"
    Check ($list -notlike '*Printer*') "reduced list no longer names Printer (got '$list')"
    Check ($extra -like '*Printer*Disregard Paper*Disregard Consumable*Print Width mm*Sheet size*') "SET_LOC_MINIMAL_COLUMNS holds the extra columns (got '$extra')"

    # ---------------------------------------------------------- initial state
    Write-Host ''
    Write-Host '=== All: nothing in either list is hidden ==='
    foreach ($h in ($reducedCols + $minimalExtra)) {
        Check (-not (IsHidden $main $lo $h)) "$h is visible on Example Print Room"
    }
    $hIssuesHiddenBefore = IsHidden $main $lo 'H_Issues'
    Check $hIssuesHiddenBefore 'H_Issues (permanently hidden) starts hidden'

    Write-Host ''
    Write-Host '=== View row: label in D2, buttons on E/F/G, side panel toggle gone ==='
    Check ([string]$main.Range('D2').Text -eq 'Show columns:') "D2 reads 'Show columns:' (got '$([string]$main.Range('D2').Text)')"
    Check ((ButtonColumn $main 'pcb_btnViewAll') -eq 5) 'All sits on column E'
    Check ((ButtonColumn $main 'pcb_btnViewReduced') -eq 6) 'Reduced sits on column F'
    Check ((ButtonColumn $main 'pcb_btnViewMinimal') -eq 7) 'Minimal sits on column G'
    Check ($null -eq (FindButton $main 'pcb_btnToggleReducedView')) 'old Reduce clutter button is gone'
    $sidePanel = 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnToggleCostColumns', 'pcb_btnRemoveRow', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation'
    foreach ($p in $sidePanel) {
        Check ((ButtonColumn $main $p) -eq 36) "$p sits in the side panel (column 36)"
    }
    # Stack is packed from the top: first button at the sheet top, no gap where the toggle was.
    $tops = $sidePanel | ForEach-Object { [double](FindButton $main $_).Top }
    $step = $tops[1] - $tops[0]
    for ($i = 1; $i -lt $tops.Count; $i++) {
        Check ([Math]::Abs(($tops[$i] - $tops[$i - 1]) - $step) -lt 0.5) "side panel step between button $i and the previous one is even (closed up)"
    }

    # ----------------------------------------------------------- Reduced view
    Write-Host ''
    Write-Host '=== Reduced hides only the reduced list, on every location sheet ==='
    [void]$xl.Run('btnViewReduced')
    Start-Sleep -Milliseconds 300
    foreach ($h in $reducedCols) {
        Check (IsHidden $main $lo $h) "$h is hidden on Example Print Room"
        Check (IsHidden $annex $aloAnnex $h) "$h is hidden on Annexe too (workbook-wide)"
    }
    foreach ($h in $minimalExtra) {
        Check (-not (IsHidden $main $lo $h)) "$h stays visible in Reduced"
    }
    Check (-not (IsHidden $main $lo 'Chargeable Cost')) 'Chargeable Cost stays visible'
    Check ((IsHidden $main $lo 'H_Issues') -eq $hIssuesHiddenBefore) "H_Issues' state is untouched"
    Check ((ViewSetting) -eq 'Reduced') 'setting reads Reduced'
    Check ([bool](FindButton $main 'pcb_btnViewReduced').Characters().Font.Bold) 'Reduced button is bold (active)'
    Check (-not [bool](FindButton $main 'pcb_btnViewAll').Characters().Font.Bold) 'All button is not bold'

    # ----------------------------------------------------------- Minimal view
    Write-Host ''
    Write-Host '=== Minimal hides reduced list AND the extras ==='
    [void]$xl.Run('btnViewMinimal')
    Start-Sleep -Milliseconds 300
    foreach ($h in ($reducedCols + $minimalExtra)) {
        Check (IsHidden $main $lo $h) "$h is hidden in Minimal"
        Check (IsHidden $annex $aloAnnex $h) "$h is hidden in Minimal on Annexe"
    }
    Check (-not (IsHidden $main $lo 'Chargeable Cost')) 'Chargeable Cost stays visible'
    Check ((IsHidden $main $lo 'H_Issues') -eq $hIssuesHiddenBefore) "H_Issues' state is untouched"
    Check ((ViewSetting) -eq 'Minimal') 'setting reads Minimal'

    # Printer (column E) is hidden now, so the buttons must have slid right.
    $lblCol = $null
    for ($c = 4; $c -le 12; $c++) { if ([string]$main.Cells(2, $c).Text -eq 'Show columns:') { $lblCol = $c; break } }
    Check ($lblCol -eq 4) "label still on D (Technician is visible; found col $lblCol)"
    $cAll = ButtonColumnVisible $main 'pcb_btnViewAll'
    $cRed = ButtonColumnVisible $main 'pcb_btnViewReduced'
    $cMin = ButtonColumnVisible $main 'pcb_btnViewMinimal'
    Check (($cAll -gt 0) -and -not [bool]$main.Columns($cAll).Hidden) "All is on a visible column (col $cAll)"
    Check (($cRed -gt $cAll) -and -not [bool]$main.Columns($cRed).Hidden) "Reduced is on a later visible column (col $cRed)"
    Check (($cMin -gt $cRed) -and -not [bool]$main.Columns($cMin).Hidden) "Minimal is on a later visible column (col $cMin)"
    $bA = FindButton $main 'pcb_btnViewAll'; $bR = FindButton $main 'pcb_btnViewReduced'; $bM = FindButton $main 'pcb_btnViewMinimal'
    Check (($bA.Left + $bA.Width) -le $bR.Left) 'All and Reduced do not overlap'
    Check (($bR.Left + $bR.Width) -le $bM.Left) 'Reduced and Minimal do not overlap'
    Check ([Math]::Abs($bA.Width - 70) -lt 0.5) 'view buttons keep their width'

    # ------------------------------------------------------------- back to All
    Write-Host ''
    Write-Host '=== All shows everything again; buttons back on E/F/G ==='
    [void]$xl.Run('btnViewAll')
    Start-Sleep -Milliseconds 300
    foreach ($h in ($reducedCols + $minimalExtra)) {
        Check (-not (IsHidden $main $lo $h)) "$h is visible again"
    }
    Check ((IsHidden $main $lo 'H_Issues') -eq $hIssuesHiddenBefore) "H_Issues' state is STILL untouched"
    Check ((ButtonColumn $main 'pcb_btnViewAll') -eq 5) 'All back on column E'
    Check ((ButtonColumn $main 'pcb_btnViewReduced') -eq 6) 'Reduced back on column F'
    Check ((ButtonColumn $main 'pcb_btnViewMinimal') -eq 7) 'Minimal back on column G'

    # --------------------------------------------- settings drive both views
    Write-Host ''
    Write-Host '=== Editing the settings changes what each view hides; Minimal still includes Reduced ==='
    $origList = [string]$wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Text
    $origExtra = [string]$wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Text
    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = 'Notes'
    $wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Value = 'Paid'
    [void]$xl.Run('btnViewReduced')
    Start-Sleep -Milliseconds 300
    Check (IsHidden $main $lo 'Notes') 'Reduced hides Notes (edited list)'
    Check (-not (IsHidden $main $lo 'Paid')) 'Reduced does not hide Paid (extra only)'
    [void]$xl.Run('btnViewMinimal')
    Start-Sleep -Milliseconds 300
    Check ((IsHidden $main $lo 'Notes') -and (IsHidden $main $lo 'Paid')) 'Minimal hides Notes (from Reduced) and Paid (extra)'
    [void]$xl.Run('btnViewAll')
    Start-Sleep -Milliseconds 300
    Check (-not (IsHidden $main $lo 'Notes') -and -not (IsHidden $main $lo 'Paid')) 'All shows both again'

    # Technician hidden by a test-only list: label and in-table buttons relocate.
    $techCol = $lo.Range.Column + (Col $lo 'Technician') - 1
    Check ((ButtonColumn $main 'pcb_btnClearDefaults') -eq $techCol) 'btnClearDefaults starts anchored on Technician'
    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = 'Technician'
    [void]$xl.Run('btnViewReduced')
    Start-Sleep -Milliseconds 300
    Check (IsHidden $main $lo 'Technician') 'Technician is hidden (test-only list)'
    $moved = ButtonColumnVisible $main 'pcb_btnClearDefaults'
    Check (($moved -gt 0) -and -not [bool]$main.Columns($moved).Hidden) "btnClearDefaults relocated to a visible column ($moved)"
    $lblCol = $null
    for ($c = 4; $c -le 12; $c++) { if ([string]$main.Cells(2, $c).Text -eq 'Show columns:') { $lblCol = $c; break } }
    Check (($lblCol -gt 4) -and -not [bool]$main.Columns($lblCol).Hidden) "label moved off hidden column D (now col $lblCol)"
    Check ([string]$main.Range('D2').Text -eq '') 'D2 no longer holds the label'
    [void]$xl.Run('btnViewAll')
    Start-Sleep -Milliseconds 300
    Check ((ButtonColumn $main 'pcb_btnClearDefaults') -eq $techCol) 'btnClearDefaults back on Technician'
    Check ([string]$main.Range('D2').Text -eq 'Show columns:') 'label back on D2'

    $wb.Names.Item('SET_LOC_REDUCED_COLUMNS').RefersToRange.Value = $origList
    $wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange.Value = $origExtra

    # --------------------------- button width survives repeated view changes
    Write-Host ''
    Write-Host '=== In-table buttons keep their width across repeated view changes ==='
    $atRisk = 'pcb_btnAddPrintJob', 'pcb_btnRepeatJob', 'pcb_btnNow', 'pcb_btnClearDefaults'
    for ($cycle = 1; $cycle -le 2; $cycle++) {
        [void]$xl.Run('btnViewMinimal'); Start-Sleep -Milliseconds 300
        [void]$xl.Run('btnViewAll'); Start-Sleep -Milliseconds 300
        foreach ($p in $atRisk) {
            $b = FindButton $main $p
            Check ([double]$b.Width -eq 110) "$p is still 110pt wide after cycle $cycle"
            Check ([int]$b.Placement -eq 3) "$p is xlFreeFloating after cycle $cycle"
        }
    }

    # ---------------------------------- side panel tracks column 36 through views
    Write-Host ''
    Write-Host '=== Side panel stays pinned to column 36 through a Minimal/All cycle ==='
    $col36Before = [double]$main.Cells(1, 36).Left
    foreach ($p in $sidePanel) {
        Check ([double](FindButton $main $p).Left -eq $col36Before) "$p sits on column 36 before hiding"
    }
    [void]$xl.Run('btnViewMinimal'); Start-Sleep -Milliseconds 300
    $col36Hidden = [double]$main.Cells(1, 36).Left
    Check ($col36Hidden -ne $col36Before) 'Minimal actually moved column 36 (otherwise this proves nothing)'
    foreach ($p in $sidePanel) {
        Check ([double](FindButton $main $p).Left -eq $col36Hidden) "$p followed column 36 while hidden"
    }
    [void]$xl.Run('btnViewAll'); Start-Sleep -Milliseconds 300
    foreach ($p in $sidePanel) {
        Check ([double](FindButton $main $p).Left -eq $col36Before) "$p back on column 36 after revealing"
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
