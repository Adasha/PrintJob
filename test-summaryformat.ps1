# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Summary-sheet formatting and the currency symbol (renamed from
# test-phase8.ps1, 0.10.25 - that name was the project phase that introduced
# these checks, not what they cover):
#
#   - Error state: red bold text on Summary's Type column whenever a job
#     references a paper stock no longer in tblPapers ("(not in Papers)").
#   - The Summary-sheet legend explaining the four everyday cell colours.
#   - SET_CURRENCY wired into the NumberFormat of the Summary and Reports
#     money cells (design doc 11 / architecture 16.2).
#
# The Status-column warning fill (the other half of the old script) moved to
# test-statusnotes.ps1, next to the Status hover notes it belongs with.
#
# Every check goes through Check, so run-tests.ps1 fails the script on a
# regression; the old script only printed "UNEXPECTED" and could not fail.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself - see
# test-validation.ps1's comment for why (AutoSave on a OneDrive-backed handle
# commits regardless of Close($false)). Closes WITHOUT saving.

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
$anyFail = $false
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    Add-AnnexeFixture $xl $wb | Out-Null

    # -------------------------------------------------- error: Summary ---
    Write-Host '=== Error state: Summary Type conditional formatting ==='
    $sum = $wb.Worksheets('Summary')
    $errRng = $sum.Range('D10:D2000')
    $n = $errRng.FormatConditions.Count
    Check ($n -ge 1) "Summary Type column has a conditional format (count $n)"
    if ($n -ge 1) {
        $fc = $errRng.FormatConditions.Item(1)
        Check ($fc.Formula1 -like '*(not in Papers)*') "rule tests for '(not in Papers)' (formula: $($fc.Formula1))"
        Check ($fc.Font.Color -eq 2097328) "rule font is red RGB(176,0,32) (got $($fc.Font.Color))"
    }

    # -------------------------------------------------------- legend ---
    Write-Host ''
    Write-Host '=== Summary legend (column P) ==='
    $expected = @('Type your own values here', 'Calculated automatically', 'Workbook configuration', 'Reference information (read-only)')
    for ($i = 0; $i -lt $expected.Count; $i++) {
        $r = 10 + $i
        $text = [string]$sum.Cells($r, 16).Text
        Check ($text -eq $expected[$i]) "P${r} reads '$($expected[$i])' (got '$text')"
    }

    # ------------------------------------------------------- currency ---
    Write-Host ''
    Write-Host '=== Currency wiring: SET_CURRENCY -> NumberFormat ==='
    $rep = $wb.Worksheets('Reports')
    $repAddr = 'F18'   # Matching: Chargeable total (modReports.REP_MATCH_VAL_ROW)
    $wb.Names.Item('SET_CURRENCY').RefersToRange.Value = '$'
    # InitialiseWorkbook ends with a summary dialog (Say) - quiet mode avoids
    # hanging on it with the Excel window invisible.
    $xl.Run('SetQuiet', $true)
    $xl.Run('InitialiseWorkbook')
    $xl.Run('SetQuiet', $false)
    $sum = $wb.Worksheets('Summary')   # re-fetch: BuildSummary rebuilds the sheet's ranges
    $rep = $wb.Worksheets('Reports')
    $sumFmt = $sum.Range('D6').NumberFormat
    $repFmt = $rep.Range($repAddr).NumberFormat
    # Exact match, not a loose "-like '*$*'": Excel's NumberFormat setter
    # silently canonicalises a BARE or quoted currency symbol back to the OS's
    # own regional currency when set from VBA, so modSettings.CurrencyFormatCode
    # wraps the symbol in Excel's [$symbol] bracket syntax and the format must
    # come back exactly as built.
    Check ($sumFmt -eq '[$$]#,##0.00') "Summary D6 follows SET_CURRENCY (got '$sumFmt')"
    Check ($repFmt -eq '[$$]#,##0.00') "Reports $repAddr follows SET_CURRENCY (got '$repFmt')"

    # Re-running InitialiseWorkbook must not have stacked duplicate rules.
    Write-Host ''
    Write-Host '=== Idempotency: re-run does not stack rules ==='
    $n2 = $sum.Range('D10:D2000').FormatConditions.Count
    Check ($n2 -eq 1) "Summary Type rule count after a second InitialiseWorkbook: $n2 (expected 1)"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($anyFail) { Write-Host 'FAIL: see above'; exit 1 }
Write-Host 'OK: summary formatting and currency wiring'
