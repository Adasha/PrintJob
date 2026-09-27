# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# LOC_RollUnit order-dependency fix (docs/ARCHITECTURE.md §16.3, 2026-09-27).
#
# modValidation.OnQtyChanged's cm->m conversion used to only fire on Qty's
# own Change event. A roll length typed BEFORE Paper Stock was chosen was
# left as raw, un-converted centimetres forever - on a Centimetres location
# that silently stored a job 100x too large in metres. Fixed by extracting
# the conversion into ConvertQtyIfCentimetres, now also called from
# OnStockChanged the moment Paper Stock resolves to a Roll stock.
#
# Covers:
#   - natural order (Paper Stock, then Qty) still converts, unchanged
#   - reversed order (Qty, then Paper Stock) now ALSO converts - the bug
#   - an already-converted Qty is not divided by 100 again when Paper Stock
#     changes a second time (e.g. swapping between two Roll stocks)
#   - a Metres-location job is never touched regardless of order
#   - a Sheet-stock job is never touched regardless of order
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself - see
# test-validation.ps1's header comment for why. Closes WITHOUT saving.

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
    $stkCol = $lo.Range.Column + (Col $lo 'Paper Stock') - 1
    $qtyCol = $lo.Range.Column + (Col $lo 'Qty') - 1

    function New-Row {
        [void]$main.Activate()
        [void]$xl.Run('btnAddPrintJob')
        return $lo.ListRows($lo.ListRows.Count).Range.Row
    }
    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }
    function IsShaded($cell) {
        # Matches modValidation.QtyMarkedRewritten's own RGB(242,242,242).
        return ($cell.Interior.Color -eq (242 + 242 * 256 + 242 * 65536))
    }

    $rollUnit = $main.Range('LOC_RollUnit')

    # --------------------------------------------------------- natural order
    Write-Host '=== Centimetres location, natural order (Paper Stock then Qty) ==='
    $rollUnit.Value2 = 'Centimetres'
    $r1 = New-Row
    $main.Cells($r1, $stkCol).Value2 = 'Canvas 914mm roll'
    Start-Sleep -Milliseconds 300
    $main.Cells($r1, $qtyCol).Value2 = 250
    Start-Sleep -Milliseconds 300
    $got = [double]$main.Cells($r1, $qtyCol).Value2
    Check ($got -eq 2.5) "250 cm converts to 2.5 m (got $got)"
    Check (IsShaded $main.Cells($r1, $qtyCol)) 'Qty is shaded (converted)'

    # -------------------------------------------------- reversed order (bug)
    Write-Host ''
    Write-Host '=== Centimetres location, reversed order (Qty typed BEFORE Paper Stock) ==='
    $r2 = New-Row
    $main.Cells($r2, $qtyCol).Value2 = 150
    Start-Sleep -Milliseconds 300
    $beforeStock = [double]$main.Cells($r2, $qtyCol).Value2
    Check ($beforeStock -eq 150) "Qty sits unconverted while Paper Stock is still blank (got $beforeStock)"
    Check (-not (IsShaded $main.Cells($r2, $qtyCol))) 'Qty not yet shaded (nothing converted yet)'

    $main.Cells($r2, $stkCol).Value2 = 'Canvas 914mm roll'
    Start-Sleep -Milliseconds 300
    $got = [double]$main.Cells($r2, $qtyCol).Value2
    Check ($got -eq 1.5) "choosing Paper Stock afterwards now converts 150 cm to 1.5 m (got $got) - the fix"
    Check (IsShaded $main.Cells($r2, $qtyCol)) 'Qty is shaded (converted) after the fix fires'

    # ------------------------------------------ no double-conversion on swap
    Write-Host ''
    Write-Host '=== Swapping Paper Stock again does not re-convert an already-converted Qty ==='
    $main.Cells($r2, $stkCol).Value2 = 'Satin photo 610mm roll'
    Start-Sleep -Milliseconds 300
    $got = [double]$main.Cells($r2, $qtyCol).Value2
    Check ($got -eq 1.5) "Qty stays at 1.5 m, not divided by 100 again (got $got)"

    # ----------------------------------------------------- Metres unaffected
    Write-Host ''
    Write-Host '=== Metres location: reversed order never converts ==='
    $rollUnit.Value2 = 'Metres'
    $r3 = New-Row
    $main.Cells($r3, $qtyCol).Value2 = 150
    Start-Sleep -Milliseconds 300
    $main.Cells($r3, $stkCol).Value2 = 'Canvas 914mm roll'
    Start-Sleep -Milliseconds 300
    $got = [double]$main.Cells($r3, $qtyCol).Value2
    Check ($got -eq 150) "Qty stays 150 on a Metres location regardless of order (got $got)"
    Check (-not (IsShaded $main.Cells($r3, $qtyCol))) 'Qty never shaded on a Metres location'

    # ------------------------------------------------------ Sheet unaffected
    Write-Host ''
    Write-Host '=== Centimetres location, Sheet stock: Qty (a sheet count) is never touched ==='
    $rollUnit.Value2 = 'Centimetres'
    $r4 = New-Row
    $main.Cells($r4, $qtyCol).Value2 = 40
    Start-Sleep -Milliseconds 300
    $main.Cells($r4, $stkCol).Value2 = 'Gloss 200gsm SRA3 sheet'
    Start-Sleep -Milliseconds 300
    $got = [double]$main.Cells($r4, $qtyCol).Value2
    Check ($got -eq 40) "Qty stays 40 sheets, not treated as a length (got $got)"
    Check (-not (IsShaded $main.Cells($r4, $qtyCol))) 'Qty never shaded for a Sheet stock'

    $rollUnit.Value2 = 'Metres'
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
