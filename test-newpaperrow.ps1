# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Regression test: a NEW row on the Papers sheet must reach the print rooms'
# Paper Stock dropdowns.
#
# After the printer/paper compatibility rework, Compatible() compares a
# stock's Width mm / Height mm with the printer's capacity. Nothing filled
# those in from the chosen Std. size, and a new row's Active was blank (only a
# literal "Yes" counts as active), so a new sheet stock fitted no printer and
# was never listed - while editing an existing row still worked, because it
# already carried its numbers. Fixed by modCatalog.OnPaperEdited (fills
# Width/Height from Std. size; defaults Active) and modCatalog.AddCatalogRow
# (defaults Active).
#
# Uses Example Print Room's catalogue: Xerox Versant 180 takes sheets up to A1.
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
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $papers = $wb.Worksheets('Papers')
    $plo = $papers.ListObjects('tblPapers')
    $main = $wb.Worksheets('Example Print Room')
    $jlo = $main.ListObjects('tblJobs_MAIN')

    function Col($lo, $name) {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            if ($lo.ListColumns($i).Name -eq $name) { return $i }
        }
        throw "column '$name' not found"
    }
    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }
    # The items currently offered by a Paper Stock cell's dropdown, read back
    # through its validation source on the _Work staging sheet.
    function Stock-Items($cell) {
        $src = [string]$cell.Validation.Formula1
        if ($src -notmatch "^='?_Work'?!(.+)$") { throw "unexpected validation source '$src'" }
        $rng = $wb.Worksheets('_Work').Range(($Matches[1] -replace '\$', ''))
        $items = @()
        foreach ($c in $rng.Cells) { $items += [string]$c.Value2 }
        return $items
    }
    function New-Paper([string]$desc, [string]$family, [string]$mode, [string]$std, $w, $h) {
        [void]$papers.Activate()
        [void]$xl.Run('AddCatalogRow', 'tblPapers')
        $row = $plo.ListRows.Count
        $r = $plo.ListRows($row).Range
        $r.Cells(1, (Col $plo 'StockID')).Value2 = 'STK-TEST-' + $row
        $r.Cells(1, (Col $plo 'Description')).Value2 = $desc
        $r.Cells(1, (Col $plo 'Paper type')).Value2 = 'Matte'
        $r.Cells(1, (Col $plo 'Family')).Value2 = $family
        $r.Cells(1, (Col $plo 'Size mode')).Value2 = $mode
        if ($std) { $r.Cells(1, (Col $plo 'Std. size')).Value2 = $std }
        if ($null -ne $w) { $r.Cells(1, (Col $plo 'Width mm')).Value2 = $w }
        if ($null -ne $h) { $r.Cells(1, (Col $plo 'Height mm')).Value2 = $h }
        $r.Cells(1, (Col $plo 'Cost')).Value2 = 0.5
        Start-Sleep -Milliseconds 300
        return $row
    }
    function Job-StockCell {
        [void]$main.Activate()
        [void]$xl.Run('btnAddPrintJob')
        $jr = $jlo.ListRows($jlo.ListRows.Count).Range.Row
        return $main.Cells($jr, (Col $jlo 'Paper Stock'))
    }

    # ------------------------------------------------ standard-size sheet stock
    Write-Host '=== New sheet stock, Size mode Standard, Std. size A3, no width/height typed ==='
    $desc = 'TEST Matte 150gsm A3 sheet'
    $row = New-Paper $desc 'Sheet' 'Standard' 'A3' $null $null
    $r = $plo.ListRows($row).Range
    $w = [double]$r.Cells(1, (Col $plo 'Width mm')).Value2
    $h = [double]$r.Cells(1, (Col $plo 'Height mm')).Value2
    $active = [string]$r.Cells(1, (Col $plo 'Active')).Value2
    Check ($w -eq 297 -and $h -eq 420) "Width/Height filled from A3 (got $w x $h)"
    Check ($active -eq 'Yes') "Active defaulted to Yes (got '$active')"
    Check ([bool]$xl.Run('Compatible', 'Xerox Versant 180', $desc)) 'compatible with Xerox Versant 180 (max sheet A1)'
    Check (-not [bool]$xl.Run('Compatible', 'HP DesignJet Z9+', $desc)) 'not compatible with HP DesignJet Z9+ (roll only)'

    Write-Host ''
    Write-Host '=== It reaches the print room dropdown after leaving Papers ==='
    $cell = Job-StockCell        # activating the print room rebinds every dropdown
    $items = Stock-Items $cell
    Check ($items -contains $desc) 'new stock is listed in Paper Stock'

    # -------------------------------------------- changing the standard size
    Write-Host ''
    Write-Host '=== Changing Std. size re-fills Width/Height ==='
    [void]$papers.Activate()
    $r.Cells(1, (Col $plo 'Std. size')).Value2 = 'A2'
    Start-Sleep -Milliseconds 300
    $w2 = [double]$r.Cells(1, (Col $plo 'Width mm')).Value2
    $h2 = [double]$r.Cells(1, (Col $plo 'Height mm')).Value2
    Check ($w2 -eq 420 -and $h2 -eq 594) "Width/Height follow the new size (got $w2 x $h2)"

    # ---------------------------------- a hand-typed custom size is not touched
    Write-Host ''
    Write-Host '=== Custom sheet: typed Width/Height are kept ==='
    $desc2 = 'TEST Custom 200x300 sheet'
    $row2 = New-Paper $desc2 'Sheet' 'Custom sheet' $null 200 300
    $r2 = $plo.ListRows($row2).Range
    Check ([double]$r2.Cells(1, (Col $plo 'Width mm')).Value2 -eq 200 -and [double]$r2.Cells(1, (Col $plo 'Height mm')).Value2 -eq 300) 'typed size unchanged'
    Check ([bool]$xl.Run('Compatible', 'Xerox Versant 180', $desc2)) 'compatible with Xerox Versant 180'

    # ------------------------------------------------ explicit Active = No kept
    Write-Host ''
    Write-Host '=== An explicit Active = No is never overwritten ==='
    $r2.Cells(1, (Col $plo 'Active')).Value2 = 'No'
    Start-Sleep -Milliseconds 300
    $r2.Cells(1, (Col $plo 'Description')).Value2 = $desc2 + ' x'
    Start-Sleep -Milliseconds 300
    Check ([string]$r2.Cells(1, (Col $plo 'Active')).Value2 -eq 'No') 'Active stays No after a further edit'

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
