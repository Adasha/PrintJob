# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Phase 8's original scope (design doc §11 / architecture §16.2), built 0.8.1:
#
#   - Warning state: amber fill via conditional formatting on every location
#     sheet's Status column, whenever a row reads anything other than "OK".
#   - Error state: red bold text via conditional formatting on Summary's
#     Type/Family columns, whenever a job references a paper stock no longer
#     in tblPapers ("(not in Papers)").
#   - The Summary-sheet legend explaining the four everyday cell colours.
#   - SET_CURRENCY wired into every NumberFormat/Format$ that used to
#     hardcode ChrW(163).
#
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself - see
# test-validation.ps1's comment for why (AutoSave on a OneDrive-backed handle
# commits regardless of Close($false)). Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
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
    Add-AnnexeFixture $xl $wb | Out-Null

    # ---------------------------------------------------- warning: Status ---
    Write-Host '=== Warning state: Status column conditional formatting ==='
    foreach ($sheetTable in @(@('Main Print Room', 'tblJobs_MAIN'), @('Annexe', 'tblJobs_ANNEX'))) {
        $ws = $wb.Worksheets($sheetTable[0])
        $lo = $ws.ListObjects($sheetTable[1])
        $rng = $lo.ListColumns('Status').DataBodyRange
        $n = $rng.FormatConditions.Count
        Write-Host ("  {0}: FormatConditions.Count={1}" -f $sheetTable[0], $n)
        if ($n -ge 1) {
            $fc = $rng.FormatConditions.Item(1)
            $okFormula = $fc.Formula1 -like '*<>*OK*'
            $okColor = $fc.Interior.Color -eq 49407  # RGB(255,192,0)
            Write-Host ("    formula: {0}  {1}" -f $fc.Formula1, $(if ($okFormula) { 'ok' } else { 'UNEXPECTED' }))
            Write-Host ("    fill:    {0}  {1}" -f $fc.Interior.Color, $(if ($okColor) { 'ok - amber RGB(255,192,0)' } else { 'UNEXPECTED' }))
        } else {
            Write-Host '    MISSING - BUG'
        }
    }

    # Sample row UNI-MAIN-00005 (SETUP.md:117) is deliberately invalid -
    # confirms the rule actually fires, not just that it exists.
    $main = $wb.Worksheets('Main Print Room')
    $mlo = $main.ListObjects('tblJobs_MAIN')
    $lastRow = $mlo.ListRows($mlo.ListRows.Count).Range
    $statusCell = $lastRow.Cells(1, 1)
    Write-Host ("  sample invalid row Status: '{0}'  DisplayFormat.Interior.Color={1}" -f `
        $statusCell.Text, $statusCell.DisplayFormat.Interior.Color)

    # -------------------------------------------------- error: Summary ---
    Write-Host ''
    Write-Host '=== Error state: Summary Type/Family conditional formatting ==='
    $sum = $wb.Worksheets('Summary')
    $errRng = $sum.Range('D10:E2000')
    $n = $errRng.FormatConditions.Count
    Write-Host ("  FormatConditions.Count={0}" -f $n)
    if ($n -ge 1) {
        $fc = $errRng.FormatConditions.Item(1)
        $okFormula = $fc.Formula1 -like '*(not in Papers)*'
        $okColor = $fc.Font.Color -eq 2097328  # RGB(176,0,32)
        Write-Host ("    formula: {0}  {1}" -f $fc.Formula1, $(if ($okFormula) { 'ok' } else { 'UNEXPECTED' }))
        Write-Host ("    font:    {0}  {1}" -f $fc.Font.Color, $(if ($okColor) { 'ok - red RGB(176,0,32)' } else { 'UNEXPECTED' }))
    } else {
        Write-Host '    MISSING - BUG'
    }

    # -------------------------------------------------------- legend ---
    Write-Host ''
    Write-Host '=== Summary legend (O9 down) ==='
    $expected = @('Type your own values here', 'Calculated automatically', 'Workbook configuration', 'Reference information (read-only)')
    for ($i = 0; $i -lt $expected.Count; $i++) {
        $r = 10 + $i
        $text = [string]$sum.Cells($r, 16).Text
        $ok = $text -eq $expected[$i]
        Write-Host ("  P{0}: '{1}'  {2}" -f $r, $text, $(if ($ok) { 'ok' } else { 'UNEXPECTED' }))
    }

    # ------------------------------------------------------- currency ---
    Write-Host ''
    Write-Host '=== Currency wiring: SET_CURRENCY -> NumberFormat ==='
    Write-Host ("  before change: Summary D6={0}  Reports H16={1}" -f `
        $sum.Range('D6').NumberFormat, $wb.Worksheets('Reports').Range('H16').NumberFormat)

    $wb.Names.Item('SET_CURRENCY').RefersToRange.Value = '$'
    # InitialiseWorkbook ends with a summary dialog (Say) - quiet mode avoids
    # hanging on it with the Excel window invisible (build.ps1's own comment
    # on this, and the reason for its SetQuiet/QuietLog pair).
    $xl.Run('SetQuiet', $true)
    $xl.Run('InitialiseWorkbook')
    $xl.Run('SetQuiet', $false)
    $sum = $wb.Worksheets('Summary')  # re-fetch: BuildSummary clears and rebuilds the sheet object's ranges
    $repFmt = $wb.Worksheets('Reports').Range('H16').NumberFormat
    $sumFmt = $sum.Range('D6').NumberFormat
    Write-Host ("  after `$ change: Summary D6={0}  Reports H16={1}" -f $sumFmt, $repFmt)
    # Exact match, not a loose "-like '*$*'": Excel's NumberFormat setter
    # silently canonicalises a BARE or quoted currency symbol back to the
    # OS's own regional currency when set from VBA (reproduced directly on
    # this machine - a plain "$#,##0.00" round-tripped as "£#,##0.00").
    # modSettings.CurrencyFormatCode wraps the symbol in Excel's [$symbol]
    # bracket syntax specifically to avoid that, so the format string must
    # come back exactly as built, brackets and all.
    $ok = ($sumFmt -eq '[$$]#,##0.00') -and ($repFmt -eq '[$$]#,##0.00')
    Write-Host ("  currency now driven by SET_CURRENCY: {0}" -f $(if ($ok) { 'ok' } else { 'UNEXPECTED - still hardcoded, or coerced back to OS currency?' }))

    # Re-running InitialiseWorkbook must not have stacked duplicate CF rules.
    Write-Host ''
    Write-Host '=== Idempotency: re-run does not stack rules ==='
    $mlo2 = $main.ListObjects('tblJobs_MAIN')
    $n2 = $mlo2.ListColumns('Status').DataBodyRange.FormatConditions.Count
    Write-Host ("  Status FormatConditions.Count after 2nd InitialiseWorkbook run: {0}  {1}" -f `
        $n2, $(if ($n2 -eq 1) { 'ok' } else { 'STACKED - BUG' }))
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
