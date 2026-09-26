# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Snag list items 2b (export header block) and 2c (single-value promotion),
# package 6b of the 2026-09-22 post-phase-8 snag list.
#
# Runs Export report twice: once filtered to a single printer (so Printer -
# and, on this workbook, Location too - should be promoted into the header
# and dropped from the table), and once unfiltered (so nothing promotes and
# every column stays). Drives a COPY in %TEMP%, never src\PrintCosts.xlsm.

$ErrorActionPreference = 'Stop'
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
    $xl.Run('SetQuiet', $true)
    $rep = $wb.Worksheets('Reports')

    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
        if (-not $cond) { $script:anyFail = $true }
    }
    $anyFail = $false

    function OpenAndReadHeader($path) {
        $xl2 = New-Object -ComObject Excel.Application
        $xl2.Visible = $false
        $xl2.DisplayAlerts = $false
        $wbCheck = $xl2.Workbooks.Open($path)
        $wsCheck = $wbCheck.Worksheets(1)
        $used = $wsCheck.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count
        # Stop at the table header ('Date/Time'), not at the first blank row -
        # 2026-09-26's layout gaps (after the title, before the Chargeable/
        # Still owed summary) put deliberate blank rows INSIDE the header
        # block, so a blank row no longer means "header is over". Blank rows
        # are skipped rather than added to Lines - callers check for presence
        # of specific labels, not position, and gaps are asserted separately
        # below by row number.
        $lines = New-Object System.Collections.Generic.List[string]
        for ($r = 1; $r -le $rows; $r++) {
            $a = [string]$wsCheck.Cells($r, 1).Value2
            if ($a -eq 'Date/Time') { break }
            if ($a -ne '') { $lines.Add($a) }
            if ($lines.Count -gt 30) { break }
        }
        $result = [PSCustomObject]@{
            Lines = $lines
            Sheet = $wsCheck
            Rows  = $rows
            Cols  = $cols
        }
        return @($result, $xl2, $wbCheck)
    }

    Write-Host '=== Filtered to one printer: Printer (and Location, single-room workbook) promote ==='
    $rep.Range('F5').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    $rep.Activate()
    $xl.Run('btnExportReport')
    Write-Host $xl.Run('QuietLog')
    Start-Sleep -Milliseconds 300
    $xlsx1 = Get-ChildItem $workDir -Filter 'PrintCosts-Report-*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Check ($null -ne $xlsx1) "an .xlsx was produced"

    $res = OpenAndReadHeader $xlsx1.FullName
    $info = $res[0]; $xl2 = $res[1]; $wbCheck = $res[2]
    try {
        Write-Host ("  header lines:`n    " + ($info.Lines -join "`n    "))
        Check ($info.Lines[0] -eq 'Print job report') "A1 is the report title"
        Check ($info.Lines -contains 'Site ID') "Site ID line present"
        Check ($info.Lines -contains 'Site name') "Site name line present"
        Check ($info.Lines -contains 'Date range') "Date range line present"
        Check ($info.Lines -contains 'Rows') "Rows line present"
        Check ($info.Lines -contains 'Total chargeable') "Total chargeable line present"
        Check ($info.Lines -contains 'Still owed') "Still owed line present"
        Check (-not ($info.Lines -contains 'Schema version')) "Schema version NOT in the header block (moved below the table)"
        Check (-not ($info.Lines -contains 'Generated')) "Generated NOT in the header block (moved below the table)"
        $printerLine = $info.Lines | Where-Object { $_ -like 'Printer:*' }
        Check ($null -ne $printerLine) "Printer promoted into the header ('$printerLine')"
        if ($printerLine) { Check ($printerLine -eq 'Printer: Epson SureColor P9500') "promoted value matches the filter ('$printerLine')" }

        # 2026-09-26 layout gaps: a blank row below the title, and another
        # above the Total chargeable/Still owed summary.
        Check (([string]$info.Sheet.Cells(2, 1).Value2) -eq '' -and ([string]$info.Sheet.Cells(2, 2).Value2) -eq '') "row 2 is blank (gap below the title)"
        $totalRow = 0
        for ($r = 1; $r -le $info.Rows; $r++) {
            if ([string]$info.Sheet.Cells($r, 1).Value2 -eq 'Total chargeable') { $totalRow = $r; break }
        }
        Check ($totalRow -gt 0) "Total chargeable row found (row $totalRow)"
        if ($totalRow -gt 0) {
            Check (([string]$info.Sheet.Cells($totalRow - 1, 1).Value2) -eq '' -and ([string]$info.Sheet.Cells($totalRow - 1, 2).Value2) -eq '') "row above Total chargeable is blank"
        }

        # Find the table header row (first row containing 'Date/Time').
        $tableRow = 0
        for ($r = 1; $r -le $info.Rows; $r++) {
            if ([string]$info.Sheet.Cells($r, 1).Value2 -eq 'Date/Time') { $tableRow = $r; break }
        }
        Check ($tableRow -gt 0) "table header row found (row $tableRow)"
        $tableHeaders = @()
        for ($c = 1; $c -le $info.Cols; $c++) {
            $h = [string]$info.Sheet.Cells($tableRow, $c).Value2
            if ($h -eq '') { break }
            $tableHeaders += $h
        }
        Write-Host ("  table columns: " + ($tableHeaders -join ', '))
        Check (-not ($tableHeaders -contains 'Printer')) "Printer column dropped from the table"
        Check ($tableHeaders -contains 'Technician') "Technician column still present (was being silently dropped by the old .Hidden-bounded LastVisibleColumn)"
        Check ($tableHeaders -contains 'Notes') "Notes column still present (same old bug)"
        Check ($tableHeaders -contains 'Chargeable') "Chargeable column still present"
        Check (-not ($tableHeaders -contains 'Area m2')) "Area m2 excluded from the export by default"

        # 2026-09-26: Schema version / Generated moved below the table, with a
        # blank gap, in low-contrast grey text (RGB(110,110,110) - same value
        # in both directions since R=G=B, so no BGR/RGB ordering confusion).
        $schemaRow = 0
        for ($r = $tableRow; $r -le $info.Rows; $r++) {
            if ([string]$info.Sheet.Cells($r, 1).Value2 -eq 'Schema version') { $schemaRow = $r; break }
        }
        Check ($schemaRow -gt 0) "Schema version found below the table (row $schemaRow)"
        if ($schemaRow -gt 0) {
            Check (([string]$info.Sheet.Cells($schemaRow - 1, 1).Value2) -eq '' -and ([string]$info.Sheet.Cells($schemaRow - 1, 2).Value2) -eq '') "row above Schema version is blank"
            Check (([string]$info.Sheet.Cells($schemaRow + 1, 1).Value2) -eq 'Generated') "Generated immediately follows Schema version"
            $expectedGrey = 110 + 110 * 256 + 110 * 65536
            Check ($info.Sheet.Cells($schemaRow, 1).Font.Color -eq $expectedGrey) "Schema version is low-contrast grey"
            Check ($info.Sheet.Cells($schemaRow + 1, 1).Font.Color -eq $expectedGrey) "Generated is low-contrast grey"
        }
    }
    finally {
        $wbCheck.Close($false)
        $xl2.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
    }

    Write-Host ''
    Write-Host '=== Unfiltered: nothing promoted, every column present ==='
    # ReportSnapshotPath names files to the minute, not the second - deleting
    # the first export before running the second avoids both landing on the
    # identical path within the same clock-minute, which would otherwise
    # trigger Excel's native "file already exists" overwrite prompt (a
    # DisplayAlerts=False-defeating dialog seen while developing this test,
    # not a product bug - Export report is never run twice a second in
    # practice).
    Remove-Item $xlsx1.FullName -Force -ErrorAction SilentlyContinue
    $rep.Range('F5').Value2 = ''
    $xl.CalculateFullRebuild()
    $rep.Activate()
    $xl.Run('btnExportReport')
    Write-Host $xl.Run('QuietLog')
    Start-Sleep -Milliseconds 300
    $xlsx2 = Get-ChildItem $workDir -Filter 'PrintCosts-Report-*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Check ($null -ne $xlsx2) "a second .xlsx was produced"

    $res2 = OpenAndReadHeader $xlsx2.FullName
    $info2 = $res2[0]; $xl3 = $res2[1]; $wbCheck2 = $res2[2]
    try {
        $printerLine2 = $info2.Lines | Where-Object { $_ -like 'Printer:*' }
        Check ($null -eq $printerLine2) "Printer NOT promoted when unfiltered (multiple printers in the results)"

        $tableRow2 = 0
        for ($r = 1; $r -le $info2.Rows; $r++) {
            if ([string]$info2.Sheet.Cells($r, 1).Value2 -eq 'Date/Time') { $tableRow2 = $r; break }
        }
        $tableHeaders2 = @()
        for ($c = 1; $c -le $info2.Cols; $c++) {
            $h = [string]$info2.Sheet.Cells($tableRow2, $c).Value2
            if ($h -eq '') { break }
            $tableHeaders2 += $h
        }
        Write-Host ("  table columns: " + ($tableHeaders2 -join ', '))
        Check ($tableHeaders2 -contains 'Printer') "Printer column present when unfiltered"
        # 16, not 17: Area m2 is excluded from the export by default
        # (2026-09-26, modExport.RemoveExcludedColumns) - derivable from Qty
        # and the paper's own dimensions, not something read directly off an
        # export. Still present on the live Reports sheet itself.
        Check (-not ($tableHeaders2 -contains 'Area m2')) "Area m2 excluded from the export by default"
        Check ($tableHeaders2.Count -eq 16) "all 16 exported result columns present (got $($tableHeaders2.Count))"
    }
    finally {
        $wbCheck2.Close($false)
        $xl3.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl3)
    }

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Write-Host ''
if ($anyFail) {
    Write-Host 'FAIL'
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
} else {
    Write-Host 'PASS'
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
