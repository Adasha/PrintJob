# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Verifies the workbook's shipped column layout on a location sheet:
#   - Column order (2026-09-22): Status/Job ID sit after Paid, before the
#     snapshot block; Notes right after Job ID; S_SchemaVer last.
#   - Outline/hidden state (was test-groups.ps1): no outline group anywhere in
#     the table; Notes visible; H_Issues and the snapshot block hidden via
#     .Hidden. Guards Cut+Insert side effects (HISTORY: Sheet size column).
#   - The location header block (rows 1-9, columns A/B) and the view buttons
#     (All/Reduced/Minimal) survive a reduced-view toggle intact.
#   - _Data/Summary/Reports still reconcile.
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


    function SheetCol($header) { return $lo.Range.Column + (Col $lo $header) - 1 }

    # ------------------------------------------------------- outline / hidden
    Write-Host '=== No outline group; Notes visible; H_Issues/snapshot hidden ==='
    foreach ($h in 'Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded', 'Chargeable Cost', 'Paid') {
        $c = $main.Columns((SheetCol $h))
        Check ($c.OutlineLevel -eq 1) "$h is at outline level 1, no group (got $($c.OutlineLevel))"
    }
    $notesCol = $main.Columns((SheetCol 'Notes'))
    Check ($notesCol.OutlineLevel -eq 1) "Notes is at outline level 1 (got $($notesCol.OutlineLevel))"
    Check (-not [bool]$notesCol.Hidden) 'Notes is NOT hidden'
    foreach ($h in 'H_Issues', 'S_PrinterID', 'S_StampedBy') {
        $c = $main.Columns((SheetCol $h))
        Check ($c.OutlineLevel -eq 1) "$h is at outline level 1, no group (got $($c.OutlineLevel))"
        Check ([bool]$c.Hidden) "$h is hidden via .Hidden, not a group"
    }
    $maxLevel = 0
    for ($ci = 1; $ci -le $lo.ListColumns.Count; $ci++) {
        $lvl = $main.Columns($lo.Range.Column + $ci - 1).OutlineLevel
        if ($lvl -gt $maxLevel) { $maxLevel = $lvl }
    }
    Check ($maxLevel -eq 1) "max outline level across the whole table is 1 (got $maxLevel)"
    Write-Host ''
    # ------------------------------------------------------------ new order
    Write-Host '=== Column order ==='
    $names = @()
    for ($i = 1; $i -le $lo.ListColumns.Count; $i++) { $names += $lo.ListColumns($i).Name }
    Write-Host ("  {0}" -f ($names -join ', '))
    Check ($names[0] -eq 'Date/Time') "column 1 is Date/Time (got '$($names[0])')"
    $statusIdx = [array]::IndexOf($names, 'Status') + 1
    $jobIdIdx = [array]::IndexOf($names, 'Job ID') + 1
    $paidIdx = [array]::IndexOf($names, 'Paid') + 1
    $notesIdx = [array]::IndexOf($names, 'Notes') + 1
    $schemaIdx = [array]::IndexOf($names, 'S_SchemaVer') + 1
    Check ($statusIdx -eq $paidIdx + 1) "Status sits right after Paid (paid=$paidIdx status=$statusIdx)"
    Check ($jobIdIdx -eq $statusIdx + 1) "Job ID sits right after Status (status=$statusIdx jobId=$jobIdIdx)"
    Check ($notesIdx -eq $jobIdIdx + 1) "Notes sits right after Job ID (jobId=$jobIdIdx notes=$notesIdx)"
    Check ($schemaIdx -eq $names.Count) "S_SchemaVer is still the last column (idx=$schemaIdx of $($names.Count))"

    # ------------------------------------------------------------- header block
    Write-Host ''
    Write-Host '=== Header block survives before toggling ==='
    Check ([string]$main.Range('B1').Text -eq 'Example Print Room') "B1 (room name) reads correctly before toggle"
    Check ([string]$main.Range('B2').Text -eq 'Print Services') "B2 (department) reads correctly before toggle"

    Write-Host ''
    Write-Host '=== Reduce clutter ON: header block still visible, button still there ==='
    [void]$xl.Run('SetViewMode', 'Reduced')
    Start-Sleep -Milliseconds 300
    Check (-not [bool]$main.Columns('A').Hidden) "column A is NOT hidden (header block lives there)"
    Check (-not [bool]$main.Columns('B').Hidden) "column B is NOT hidden (header block lives there)"
    Check ([string]$main.Range('B1').Text -eq 'Example Print Room') "B1 (room name) still reads correctly with reduced view ON"
    Check ([string]$main.Range('B2').Text -eq 'Print Services') "B2 (department) still reads correctly with reduced view ON"

    $sheetStatusCol = $lo.Range.Column + $statusIdx - 1
    $sheetJobIdCol = $lo.Range.Column + $jobIdIdx - 1
    Check ([bool]$main.Columns($sheetStatusCol).Hidden) "Status's own (new) column IS hidden"
    Check ([bool]$main.Columns($sheetJobIdCol).Hidden) "Job ID's own (new) column IS hidden"

    $found = $false
    for ($i = 1; $i -le $main.DropDowns().Count; $i++) {
        if ([string]$main.DropDowns($i).Name -like 'pcb_ddViewMode*') { $found = $true }
    }
    Check $found "the view drop-down is still present after switching to Reduced"

    Write-Host ''
    Write-Host '=== Reduce clutter OFF again: everything restored ==='
    [void]$xl.Run('SetViewMode', 'All')
    Start-Sleep -Milliseconds 300
    Check (-not [bool]$main.Columns($sheetStatusCol).Hidden) "Status's column visible again"
    Check (-not [bool]$main.Columns($sheetJobIdCol).Hidden) "Job ID's column visible again"

    # ---------------------------------------------------------------- sanity
    Write-Host ''
    Write-Host '=== Downstream sanity: Summary/Reports still reconcile ==='
    $xl.CalculateFullRebuild()
    $sum = $wb.Worksheets('Summary')
    $chargeable = [double]$sum.Range("H$SUM_TOT_ROW").Value2
    $paid = [double]$sum.Range("J$SUM_TOT_ROW").Value2
    $unpaidV = [double]$sum.Range("L$SUM_TOT_ROW").Value2
    Check ([Math]::Abs(($paid + $unpaidV) - $chargeable) -lt 0.01) "Summary Paid+Unpaid still reconciles to Chargeable ($paid + $unpaidV = $($paid+$unpaidV), Chargeable=$chargeable)"
    $data = $wb.Worksheets('_Data')
    $spill = $data.Range('A10').SpillingToRange
    Check ($spill.Rows.Count -ge 5) "consolidated range still has rows (got $($spill.Rows.Count))"

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
