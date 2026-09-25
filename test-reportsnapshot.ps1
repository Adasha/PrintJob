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
        $lines = New-Object System.Collections.Generic.List[string]
        for ($r = 1; $r -le $rows; $r++) {
            $a = [string]$wsCheck.Cells($r, 1).Value2
            if ($a -eq '' -and [string]$wsCheck.Cells($r, 2).Value2 -eq '') { break }
            $lines.Add($a)
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
    $rep.Range('F6').Value2 = 'Epson SureColor P9500'
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
        Check ($info.Lines -contains 'Schema version') "Schema version line present"
        Check ($info.Lines -contains 'Site ID') "Site ID line present"
        Check ($info.Lines -contains 'Site name') "Site name line present"
        Check ($info.Lines -contains 'Date range') "Date range line present"
        Check ($info.Lines -contains 'Generated') "Generated line present"
        Check ($info.Lines -contains 'Rows') "Rows line present"
        $printerLine = $info.Lines | Where-Object { $_ -like 'Printer:*' }
        Check ($null -ne $printerLine) "Printer promoted into the header ('$printerLine')"
        if ($printerLine) { Check ($printerLine -eq 'Printer: Epson SureColor P9500') "promoted value matches the filter ('$printerLine')" }

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
    $rep.Range('F6').Value2 = ''
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
        Check ($tableHeaders2.Count -eq 17) "all 17 result columns present (got $($tableHeaders2.Count))"
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
