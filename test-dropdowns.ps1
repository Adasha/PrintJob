# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Snag 1a: bidirectional Printer / Paper Stock dropdowns.
#
#   - Paper Stock chosen first still auto-fills Printer when only one
#     compatible option exists (and vice versa - Printer chosen first still
#     auto-fills Paper Stock, the pre-existing direction).
#   - An incompatible combination still gets caught, clearing the field that
#     was just changed and leaving the other one alone.
#
# 2026-09-23: the dropdown itself no longer narrows to compatible-only
# options (too restrictive per user feedback) - it always lists every active/
# permitted item, marking the ones the other field's current value rules out
# with the ' (unavailable)' suffix instead of hiding them. Auto-fill is
# unaffected (still driven by the narrowed/compatible count, just never
# shown to the dropdown itself). Picking a marked item is allowed and is
# exactly the AT-05 "incompatible combination" path, reached via the
# dropdown instead of by typing over an already-filled cell.
#
# Uses Example Print Room's real catalogue data (see Printers/Papers sheets).
# Compatibility is a numeric capacity fit (printer/paper compatibility
# rework), not the family-band membership this comment used to describe -
# migrated from the original "Supported families" data so the same
# combinations below still hold:
#   Xerox Versant 180     -> Max sheet size A1, no roll capacity (permitted here)
#   Epson SureColor P9500 -> Max roll width 1370mm              (permitted here)
#   HP DesignJet Z9+      -> Max roll width 610mm                (permitted here)
#   HP Latex 335          -> Max roll width 1370mm               (NOT permitted here)
#
#   Gloss 200gsm SRA3 sheet -> fits only Xerox's A1 max         (singleton)
#   Canvas 914mm roll       -> exceeds HP DesignJet's 610mm max, fits only Epson (singleton)
#   Satin photo 610mm roll  -> fits both Epson and HP DesignJet's roll max (2)
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself - see
# test-validation.ps1's header comment for why (Workbook_Open + AutoSave on a
# cloud-backed handle commits changes regardless of Close($false)).
#
# Closes WITHOUT saving.

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
    $work = $wb.Worksheets('_Work')

    # Absolute sheet columns, not table-relative - returning a live Range/
    # ListRow COM object from a PowerShell function is unreliable (PowerShell
    # can silently unroll an enumerable COM object through the pipeline), so
    # every cell below is addressed fresh off $main.Cells(row, col) using
    # plain integers only, never a COM object handed back from a function.
    $prnCol = $lo.Range.Column + (Col $lo 'Printer') - 1
    $stkCol = $lo.Range.Column + (Col $lo 'Paper Stock') - 1

    function New-Row {
        # Every statement's result that isn't voided joins this function's
        # output stream in PowerShell, turning a plain integer return into an
        # array (Activate()/Run() both produce COM return values) - [void]
        # each one so $lo.ListRows(...).Range.Row is the only thing returned.
        [void]$main.Activate()
        [void]$xl.Run('btnAddPrintJob')
        return $lo.ListRows($lo.ListRows.Count).Range.Row
    }

    function Staged([string]$tag) {
        for ($c = 1; $c -le 200; $c++) {
            if ([string]$work.Cells(1, $c).Text -eq $tag) {
                $items = @()
                for ($r = 2; $r -le 60; $r++) {
                    $v = [string]$work.Cells($r, $c).Text
                    if ($v -eq '') { break }
                    $items += $v
                }
                return $items
            }
        }
        return @()
    }


    # ------------------------------------------------------- stock -> printer
    Write-Host '=== Paper Stock chosen first narrows/auto-fills Printer ==='
    $r1 = New-Row
    $c = $main.Cells($r1, $stkCol)
    $c.Value2 = 'Gloss 200gsm SRA3 sheet'
    Start-Sleep -Milliseconds 300
    $got = [string]$main.Cells($r1, $prnCol).Text
    Check ($got -eq 'Xerox Versant 180') "Sheet stock auto-fills the one compatible printer (got '$got')"

    $r2 = New-Row
    $c = $main.Cells($r2, $stkCol)
    $c.Value2 = 'Canvas 914mm roll'
    Start-Sleep -Milliseconds 300
    $got = [string]$main.Cells($r2, $prnCol).Text
    Check ($got -eq 'Epson SureColor P9500') "Long Roll stock auto-fills the one compatible printer (got '$got')"

    $r3 = New-Row
    $c = $main.Cells($r3, $stkCol)
    $c.Value2 = 'Satin photo 610mm roll'
    Start-Sleep -Milliseconds 300
    $got = [string]$main.Cells($r3, $prnCol).Text
    Check ([string]::IsNullOrEmpty($got)) "Short Roll stock (2 compatible printers) leaves Printer blank, not auto-filled (got '$got')"
    $staged = Staged 'PRN|Example Print Room|Satin photo 610mm roll'
    Check ($staged.Count -eq 3) ("Printer list still shows all 3 permitted printers, not narrowed (got {0}: {1})" -f $staged.Count, ($staged -join ', '))
    Check ($staged -contains 'Xerox Versant 180 (unavailable)') "the incompatible printer (Sheet only) is listed but marked unavailable (got: $($staged -join ', '))"
    Check ($staged -contains 'Epson SureColor P9500') "the compatible Epson is listed, unmarked (got: $($staged -join ', '))"
    Check ($staged -contains 'HP DesignJet Z9+') "the compatible HP DesignJet is listed, unmarked (got: $($staged -join ', '))"

    # ------------------------------------------------- picking a marked item
    Write-Host ''
    Write-Host '=== Picking a marked "(unavailable)" dropdown entry re-filters, AT-05 style ==='
    $c = $main.Cells($r3, $prnCol)
    $c.Value2 = 'Xerox Versant 180 (unavailable)'
    Start-Sleep -Milliseconds 300
    $prnAfter = [string]$main.Cells($r3, $prnCol).Text
    $stkAfter = [string]$main.Cells($r3, $stkCol).Text
    Check ($prnAfter -eq 'Xerox Versant 180') "the marker is stripped off the Printer cell itself (got '$prnAfter')"
    Check ([string]::IsNullOrEmpty($stkAfter)) "the now-incompatible Paper Stock is cleared (got '$stkAfter')"
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*cannot be used on*') 'picking the marked entry still raised the AT-05 incompatibility warning'

    # ------------------------------------------------------- printer -> stock
    Write-Host ''
    Write-Host '=== Printer chosen first still narrows/auto-fills Paper Stock ==='
    # HP DesignJet Z9+ used to have exactly one compatible stock (Satin
    # photo 610mm roll, the only Short Roll family stock) and auto-filled
    # it. The printer/paper compatibility rework's "Supplied (Roll)" option
    # is compatible with every roll-capable printer regardless of width, so
    # this printer now has TWO compatible stocks and no longer auto-fills -
    # same as Xerox Versant 180 below, which already had more than one.
    $r4 = New-Row
    $c = $main.Cells($r4, $prnCol)
    $c.Value2 = 'HP DesignJet Z9+'
    Start-Sleep -Milliseconds 300
    $got = [string]$main.Cells($r4, $stkCol).Text
    Check ([string]::IsNullOrEmpty($got)) "HP DesignJet Z9+ (2 compatible stocks: Satin photo roll + Supplied (Roll)) leaves Paper Stock blank (got '$got')"

    $r5 = New-Row
    $c = $main.Cells($r5, $prnCol)
    $c.Value2 = 'Xerox Versant 180'
    Start-Sleep -Milliseconds 300
    $got = [string]$main.Cells($r5, $stkCol).Text
    Check ([string]::IsNullOrEmpty($got)) "Xerox Versant 180 (5 compatible sheet stocks) leaves Paper Stock blank (got '$got')"

    # --------------------------------------------------------- incompatible
    Write-Host ''
    Write-Host '=== Incompatible combination still caught (AT-05, both directions) ==='
    $r6 = New-Row
    $c = $main.Cells($r6, $prnCol)
    $c.Value2 = 'Xerox Versant 180'
    Start-Sleep -Milliseconds 300
    $c = $main.Cells($r6, $stkCol)
    $c.Value2 = 'Canvas 914mm roll'   # Long Roll - incompatible with Xerox (Sheet only)
    Start-Sleep -Milliseconds 300
    $stkAfter = [string]$main.Cells($r6, $stkCol).Text
    $prnAfter = [string]$main.Cells($r6, $prnCol).Text
    Check ([string]::IsNullOrEmpty($stkAfter)) "the just-typed incompatible Paper Stock is cleared (got '$stkAfter')"
    Check ($prnAfter -eq 'Xerox Versant 180') "the existing Printer is left alone (got '$prnAfter')"
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*cannot be used on*') 'a warning naming the incompatibility was raised'

    $r7 = New-Row
    $c = $main.Cells($r7, $stkCol)
    $c.Value2 = 'Canvas 914mm roll'
    Start-Sleep -Milliseconds 300
    $c = $main.Cells($r7, $prnCol)
    $c.Value2 = 'Xerox Versant 180'   # Sheet only - incompatible with Canvas (Long Roll)
    Start-Sleep -Milliseconds 300
    $stkAfter = [string]$main.Cells($r7, $stkCol).Text
    $prnAfter = [string]$main.Cells($r7, $prnCol).Text
    Check ([string]::IsNullOrEmpty($stkAfter)) "changing to an incompatible Printer clears the existing Paper Stock (got '$stkAfter')"
    Check ($prnAfter -eq 'Xerox Versant 180') "the just-typed Printer itself is kept (got '$prnAfter')"

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
