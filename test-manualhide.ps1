# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Manual column hide/unhide (NOT via the reduced-view toggle) while staying on
# the sheet. Guards the Workbook_SheetSelectionChange self-heal: the user
# hides a column by hand (or expands/collapses the cost-columns group) and
# never leaves the sheet, so Workbook_SheetActivate never fires. The next
# click on a cell must:
#   - move a table-anchored button (Clear defaults, on Technician) off a
#     column that is now hidden, and back when it is revealed;
#   - keep every side-panel button on column 47 as column 47 shifts left.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT
# saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
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

    function ButtonLeft($ws, $prefix) {
        $n = $ws.Buttons().Count
        for ($i = 1; $i -le $n; $i++) {
            $b = $ws.Buttons($i)
            if ([string]$b.Name -like "$prefix*") { return [double]$b.Left }
        }
        return -1
    }
    # First VISIBLE column whose Left matches the button (a hidden column
    # shares its Left with the visible one after it).
    function ButtonColumnVisible($ws, $prefix) {
        $left = ButtonLeft $ws $prefix
        if ($left -lt 0) { return 0 }
        for ($c = 1; $c -le 60; $c++) {
            if ([Math]::Abs($ws.Cells(1, $c).Left - $left) -lt 0.5 -and -not [bool]$ws.Columns($c).Hidden) { return $c }
        }
        return -1
    }
    # Fire Workbook_SheetSelectionChange by moving the selection.
    function Nudge($ws) {
        $ws.Range('B3').Select() | Out-Null
        Start-Sleep -Milliseconds 200
        $ws.Range('B4').Select() | Out-Null
        Start-Sleep -Milliseconds 300
    }

    $main.Activate()
    $techCol = $lo.Range.Column + (Col $lo 'Technician') - 1
    $sidePanel = 'pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnRemoveRow', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation'

    Write-Host '=== Baseline ==='
    Check ((ButtonColumnVisible $main 'pcb_btnClearDefaults') -eq $techCol) "btnClearDefaults starts on Technician (col $techCol)"
    $col36Before = [double]$main.Cells(1, 47).Left
    foreach ($p in $sidePanel) {
        Check ((ButtonLeft $main $p) -eq $col36Before) "$p on column 47 (Left=$col36Before)"
    }

    Write-Host ''
    Write-Host '=== Manually hide Technician, stay on the sheet, click a cell ==='
    $main.Columns($techCol).Hidden = $true
    Nudge $main
    $moved = ButtonColumnVisible $main 'pcb_btnClearDefaults'
    Check ($moved -gt 0 -and $moved -ne $techCol) "btnClearDefaults left the hidden Technician column (now col $moved)"
    Check ($moved -gt 0 -and -not [bool]$main.Columns($moved).Hidden) "its new column is visible"

    Write-Host ''
    Write-Host '=== Reveal Technician, click a cell ==='
    $main.Columns($techCol).Hidden = $false
    Nudge $main
    Check ((ButtonColumnVisible $main 'pcb_btnClearDefaults') -eq $techCol) "btnClearDefaults back on Technician"

    Write-Host ''
    Write-Host '=== Manually hide a column left of 47, click a cell: side panel follows ==='
    $hideCol = $lo.Range.Column + (Col $lo 'Printer') - 1
    $main.Columns($hideCol).Hidden = $true
    Nudge $main
    $col36Hidden = [double]$main.Cells(1, 47).Left
    Check ($col36Hidden -ne $col36Before) "hiding column $hideCol moved column 47 (was $col36Before, now $col36Hidden)"
    foreach ($p in $sidePanel) {
        Check ((ButtonLeft $main $p) -eq $col36Hidden) "$p followed column 47 (Left=$col36Hidden)"
    }
    $main.Columns($hideCol).Hidden = $false
    Nudge $main
    foreach ($p in $sidePanel) {
        Check ((ButtonLeft $main $p) -eq $col36Before) "$p back on column 47 after reveal"
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
