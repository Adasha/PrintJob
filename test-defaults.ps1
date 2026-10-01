# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Snag 1b: batch default Technician/Printer/Paper selectors, one row each at
# A3:B5 (2026-09-25 layout fix - moved off their original single shared row
# to stop "Default: paper" landing on the Printer column and vanishing under
# reduced view; EnsureJobTableGap, which used to open the space for that
# single row, was removed along with it).
#
#   - Setting a default pre-fills every subsequently added job.
#   - Changing a default afterwards never alters jobs already added (AT-07/
#     AT-08's copy-not-reference principle, spec 9.2/10.10, applied here too).
#   - The defaults implement the same bidirectional filtering/autofill as the
#     table cells (spec 1a): choosing one narrows the other, singleton
#     narrows auto-fill.
#   - Clear defaults empties all three.
#
# Uses the same Example Print Room catalogue data as test-dropdowns.ps1.
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

    $techCol = $lo.Range.Column + (Col $lo 'Technician') - 1
    $prnCol  = $lo.Range.Column + (Col $lo 'Printer') - 1
    $stkCol  = $lo.Range.Column + (Col $lo 'Paper Stock') - 1

    function New-Row {
        [void]$main.Activate()
        [void]$xl.Run('btnAddPrintJob')
        return $lo.ListRows($lo.ListRows.Count).Range.Row
    }


    # ------------------------------------------------------------ named cells
    Write-Host '=== Named cells exist and are laid out one per row at A3:B5 ==='
    $expectedRows = @{ LOC_DefTech = 3; LOC_DefPrinter = 4; LOC_DefPaper = 5 }
    foreach ($nm in 'LOC_DefTech', 'LOC_DefPrinter', 'LOC_DefPaper') {
        $found = $false
        try { $r = $main.Names.Item($nm).RefersToRange; $found = ($r.Row -eq $expectedRows[$nm]) } catch {}
        Check $found "$nm exists and refers to row $($expectedRows[$nm])"
    }
    Check ([string]$main.Range('A3').Text -eq 'Default: technician') "A3 label reads correctly (got '$($main.Range('A3').Text)')"

    # Disregard-cost defaults sit under Default: paper (2026-09-29), row 8 is a blank spacer,
    # and the settings below moved down to rows 9-11.
    $layout = @{ LOC_DefDisPaper = 6; LOC_DefDisCons = 7; LOC_RollUnit = 9 }
    foreach ($nm in $layout.Keys) {
        $found = $false
        try { $r = $main.Names.Item($nm).RefersToRange; $found = ($r.Row -eq $layout[$nm] -and $r.Column -eq 2) } catch {}
        Check $found "$nm refers to B$($layout[$nm])"
    }
    Check ([string]$main.Range('A6').Text -eq 'Disregard paper') "A6 label reads correctly (got '$($main.Range('A6').Text)')"
    Check ([string]$main.Range('A7').Text -eq 'Disregard consumable') "A7 label reads correctly (got '$($main.Range('A7').Text)')"
    Check ([string]$main.Range('A8').Text -eq '' -and [string]$main.Range('B8').Text -eq '') 'Row 8 is a blank spacer'
    Check ([string]$main.Range('A9').Text -eq 'Roll length unit') "A9 label reads correctly (got '$($main.Range('A9').Text)')"
    Check ([string]$main.Range('A11').Text -eq 'Print jobs') "A11 label reads correctly (got '$($main.Range('A11').Text)')"

    # ------------------------------------------------------------- pre-fill
    Write-Host ''
    Write-Host '=== Setting defaults pre-fills subsequently added jobs ==='
    $tCell = $main.Names.Item('LOC_DefTech').RefersToRange
    $tCell.Value2 = 'A. Shailer'
    $pCell = $main.Names.Item('LOC_DefPrinter').RefersToRange
    $pCell.Value2 = 'Xerox Versant 180'
    Start-Sleep -Milliseconds 300
    # Xerox is Sheet-only -> Paper should auto-fill only if singleton; Main has
    # 5 sheet stocks, so it should stay blank, not auto-filled.
    $sCell = $main.Names.Item('LOC_DefPaper').RefersToRange
    Check ([string]::IsNullOrEmpty([string]$sCell.Text)) "Default Paper stays blank (5 compatible sheet stocks, not a singleton)"
    $sCell.Value2 = 'Gloss 200gsm SRA3 sheet'
    Start-Sleep -Milliseconds 300

    $r1 = New-Row
    $gotTech = [string]$main.Cells($r1, $techCol).Text
    $gotPrn  = [string]$main.Cells($r1, $prnCol).Text
    $gotStk  = [string]$main.Cells($r1, $stkCol).Text
    Check ($gotTech -eq 'A. Shailer') "new job's Technician pre-filled (got '$gotTech')"
    Check ($gotPrn -eq 'Xerox Versant 180') "new job's Printer pre-filled (got '$gotPrn')"
    Check ($gotStk -eq 'Gloss 200gsm SRA3 sheet') "new job's Paper Stock pre-filled (got '$gotStk')"

    $r2 = New-Row
    $gotTech2 = [string]$main.Cells($r2, $techCol).Text
    Check ($gotTech2 -eq 'A. Shailer') "a second job also pre-fills from the same defaults (got '$gotTech2')"

    # ---------------------------------------------- changing default is safe
    Write-Host ''
    Write-Host '=== Changing a default afterwards leaves earlier jobs alone (AT-07/AT-08) ==='
    $tCell.Value2 = 'J. Okonkwo'
    Start-Sleep -Milliseconds 300
    $stillA = [string]$main.Cells($r1, $techCol).Text
    Check ($stillA -eq 'A. Shailer') "row 1's Technician is untouched by the later default change (got '$stillA')"

    # ------------------------------------------------------- autofill/narrow
    Write-Host ''
    Write-Host '=== Defaults implement the same bidirectional filtering as table cells (1a) ==='
    [void]$xl.Run('ClearDefaults', $main)
    Start-Sleep -Milliseconds 200
    $sCell.Value2 = 'Canvas 914mm roll'   # Long Roll -> only Epson fits at this location
    Start-Sleep -Milliseconds 300
    $gotPrn3 = [string]$pCell.Text
    Check ($gotPrn3 -eq 'Epson SureColor P9500') "Long Roll default paper auto-fills the one compatible default printer (got '$gotPrn3')"

    # ------------------------------------------------------------------- clear
    Write-Host ''
    Write-Host '=== Clear defaults empties all three ==='
    [void]$xl.Run('ClearDefaults', $main)
    Start-Sleep -Milliseconds 300
    $allBlank = ([string]::IsNullOrEmpty([string]$tCell.Text)) -and ([string]::IsNullOrEmpty([string]$pCell.Text)) -and ([string]::IsNullOrEmpty([string]$sCell.Text))
    Check $allBlank "all three default cells are blank after Clear defaults"

    $r3 = New-Row
    $gotTech3 = [string]$main.Cells($r3, $techCol).Text
    Check ([string]::IsNullOrEmpty($gotTech3)) "a job added after Clear defaults gets no pre-fill (got '$gotTech3')"

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
