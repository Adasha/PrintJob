# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Snag 1c (Paid column) and 1d (column grouping fix).
#
#   - Paid column exists right after Chargeable Cost, Yes/No validation.
#   - New jobs default to "No"; pre-existing rows stay blank (not force-set).
#   - Summary and Reports totals split Chargeable by paid status, and the
#     two always reconcile exactly to the Chargeable total.
#   - The cost-column group runs Paper Cost..Disregarded only - Chargeable
#     Cost and Paid stay outside it (ungrouped, always visible).
#   - Export/Import carry Paid through.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT
# saving except where noted.

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


    # --------------------------------------------------------------- schema
    Write-Host '=== Schema version bumped ==='
    # 1.3 since the paper family removal dropped the S_Family snapshot column
    # (1.2 was the Sheet size job-row column; modUtils.SCHEMA_VER) - Paid itself is still what
    # bumped it from 1.0 to 1.1 originally, unaffected by this test.
    $schema = [string]$wb.Names.Item('SET_SCHEMA').RefersToRange.Text
    Check ($schema -eq '1.3') "SET_SCHEMA reads 1.3 (got '$schema')"

    # ----------------------------------------------------------- column exists
    Write-Host ''
    Write-Host '=== Paid column ==='
    $paidCol = Col $lo 'Paid'
    $chgCol = Col $lo 'Chargeable Cost'
    $notesCol = Col $lo 'Notes'
    Check ($paidCol -gt 0) "Paid column exists"
    Check ($paidCol -eq $chgCol + 1) "Paid sits immediately after Chargeable Cost (chg=$chgCol paid=$paidCol)"
    # Not "immediately after Paid" any more: 2026-09-22's ReorderJobColumns
    # (test-layout.ps1 covers this precisely) moved Status/Job ID to sit
    # between Paid and Notes, fixing a reduced-clutter-view header collision.
    Check ($notesCol -gt $paidCol) "Notes still sits after Paid (paid=$paidCol notes=$notesCol)"

    $sheetPaidCol = $lo.Range.Column + $paidCol - 1
    $firstDataRow = $lo.Range.Row + 1  # table header + 1 - not hardcoded, since 0.9.14's EnsureJobTableGap moved the table's own start row
    $dv = $main.Cells($firstDataRow, $sheetPaidCol).Validation
    Check ($dv.Type -eq 3) "Paid cells carry list (Yes/No) validation"

    # ---------------------------------------------------- existing rows blank
    Write-Host ''
    Write-Host '=== Pre-existing rows are left blank, not force-set ==='
    $row1Paid = [string]$main.Cells($firstDataRow, $sheetPaidCol).Text
    Check ([string]::IsNullOrEmpty($row1Paid)) "row 1 (pre-existing, predates the column) has a blank Paid (got '$row1Paid')"

    # ------------------------------------------------------------- new job
    Write-Host ''
    Write-Host '=== New jobs default to No ==='
    [void]$main.Activate()
    [void]$xl.Run('btnAddPrintJob')
    $newRow = $lo.ListRows($lo.ListRows.Count).Range.Row
    $newPaid = [string]$main.Cells($newRow, $sheetPaidCol).Text
    Check ($newPaid -eq 'No') "new job's Paid defaults to No (got '$newPaid')"

    # Mark row 1 (blank/unpaid) and the new row (No) - leave one Yes for the
    # totals check below.
    $sheetChgCol = $lo.Range.Column + $chgCol - 1
    $main.Cells($firstDataRow, $sheetPaidCol).Value2 = 'Yes'
    Start-Sleep -Milliseconds 200

    # ---------------------------------------------------------- Summary totals
    Write-Host ''
    Write-Host '=== Summary totals split by paid status and reconcile ==='
    $xl.CalculateFullRebuild()
    $sum = $wb.Worksheets('Summary')
    $chargeable = [double]$sum.Range('H6').Value2
    $paid = [double]$sum.Range('J6').Value2
    $unpaid = [double]$sum.Range('L6').Value2
    Check ($paid -gt 0) "Paid total is non-zero (got $paid)"
    Check ([Math]::Abs(($paid + $unpaid) - $chargeable) -lt 0.01) "Paid + Unpaid reconciles to Chargeable ($paid + $unpaid = $($paid+$unpaid), Chargeable=$chargeable)"

    # ---------------------------------------------------------- Reports totals
    Write-Host ''
    Write-Host '=== Reports "Matching" totals split by paid status and reconcile ==='
    $rep = $wb.Worksheets('Reports')
    $rChargeable = [double]$rep.Range('H13').Value2
    $rPaid = [double]$rep.Range('J13').Value2
    $rUnpaid = [double]$rep.Range('L13').Value2
    Check ([Math]::Abs(($rPaid + $rUnpaid) - $rChargeable) -lt 0.01) "Paid + Unpaid reconciles to Chargeable on Reports ($rPaid + $rUnpaid = $($rPaid+$rUnpaid), Chargeable=$rChargeable)"

    # ------------------------------------------------------- cost visibility
    # 2026-09-27: the outline group over Paper Cost..Disregarded was replaced
    # by the ToggleCostColumns button (docs/ARCHITECTURE.md §16.3, "button
    # lag" writeup) - test-costcolumns.ps1 covers that toggle directly. All
    # this checks here is that Chargeable Cost and Paid are never touched by
    # it, same guarantee the old outline group used to make.
    Write-Host ''
    Write-Host '=== Chargeable Cost and Paid are never hidden by the cost-columns toggle ==='
    Check (-not [bool]$main.Columns($sheetChgCol).Hidden) "Chargeable Cost is visible by default (got hidden=$([bool]$main.Columns($sheetChgCol).Hidden))"
    Check (-not [bool]$main.Columns($sheetPaidCol).Hidden) "Paid is visible by default (got hidden=$([bool]$main.Columns($sheetPaidCol).Hidden))"
    [void]$xl.Run('btnToggleCostColumns')
    Start-Sleep -Milliseconds 300
    Check (-not [bool]$main.Columns($sheetChgCol).Hidden) "Chargeable Cost stays visible once cost detail is hidden (got hidden=$([bool]$main.Columns($sheetChgCol).Hidden))"
    Check (-not [bool]$main.Columns($sheetPaidCol).Hidden) "Paid stays visible once cost detail is hidden (got hidden=$([bool]$main.Columns($sheetPaidCol).Hidden))"
    [void]$xl.Run('btnToggleCostColumns')
    Start-Sleep -Milliseconds 300

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
