# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Summary (regrouped to a Location x Printer x Paper Stock key) and Reports
# (renamed from Cost Calculations, extended with Technician/Printer/Paper
# Stock/Quantity filters, sort-by-column, and a hidden Job ID column).
#
# AT-10  a student's total across every print room in this workbook
# AT-12  the four blank-criteria cases
#
# Drives the criteria cells, then closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself - see verify.ps1's
# comment for why (AutoSave on a OneDrive-backed handle commits regardless of
# Close($false)).
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
    # Excel can still be settling from Workbook_Open's own macro work when
    # this lands - retried the same way build.ps1 retries its SaveAs.
    $s = Invoke-ComRetry { $wb.Worksheets('Summary') }
    $c = $wb.Worksheets('Reports')

    # Reports' results columns moved on 2026-09-22 (Student Name/No and Paid
    # added, snag list items 2a/2d) - found by header (row 15) rather than
    # hardcoded, so this test doesn't go stale again the next time a column
    # is added.
    function ReportsCol($header) {
        for ($col = 1; $col -le 30; $col++) {
            if ([string]$c.Cells(15, $col).Text -eq $header) { return $col }
        }
        return 0
    }

    Write-Host '=== Reports: hidden Job ID correlation column ==='
    $jobIdCol = ReportsCol 'Job ID'
    if ($jobIdCol -eq 0) { Write-Host 'FAIL: no Job ID header found on row 15'; exit 1 }
    # [char] on its own is a .NET Char, and Excel's COM Columns(...) indexer
    # silently takes THAT as a numeric column index (its ordinal value, e.g.
    # 82 for 'R') rather than the single-letter column reference it looks
    # like - .ToString() forces it to a real string first. Found by
    # comparing Columns($jobIdColLetter) against the literal Columns('R')
    # side by side: same column, different (wrong) answer without this cast.
    $jobIdColLetter = ([char](64 + $jobIdCol)).ToString()
    if (-not $c.Columns($jobIdColLetter).Hidden) { Write-Host "FAIL: column $jobIdColLetter (Job ID) should be hidden"; exit 1 }
    $sp = $c.Range('A16').SpillingToRange
    Write-Host ("  results spill: {0} rows x {1} cols" -f $sp.Rows.Count, $sp.Columns.Count)
    $jobIdSample = [string]$sp.Cells(1, $jobIdCol).Value2
    Write-Host ("  {0}16 (Job ID) sample value: '{1}'" -f $jobIdColLetter, $jobIdSample)
    if ($jobIdSample -notmatch '-MAIN-|-ANNEX-') { Write-Host 'FAIL: hidden Job ID column does not look like a Job ID'; exit 1 }
    Write-Host "  OK: Job ID is present as a hidden column ($jobIdColLetter)"

    Write-Host ''
    Write-Host '=== Summary (Location x Printer x Paper Stock) ==='
    Write-Host ("  A10 formula length: {0} chars" -f $s.Range('A10').Formula2.Length)
    try {
        $sp = $s.Range('A10').SpillingToRange
        Write-Host ("  spill: {0} rows x {1} cols" -f $sp.Rows.Count, $sp.Columns.Count)
    } catch { Write-Host ('  NO SPILL: ' + $s.Range('A10').Text) }
    Write-Host ("  totals: jobs={0} gross={1} disregarded={2} chargeable={3}" -f `
        $s.Range('B6').Text, $s.Range('D6').Text, $s.Range('F6').Text, $s.Range('H6').Text)
    $hdrRow = @()
    for ($col = 1; $col -le 14; $col++) { $hdrRow += [string]$s.Cells(9, $col).Text }
    Write-Host ('  headers: ' + ($hdrRow -join ' | '))
    if ($hdrRow[1] -ne 'Printer') {
        Write-Host "FAIL: Summary column B header should be 'Printer', got '$($hdrRow[1])'"
        exit 1
    }
    Write-Host "  OK: Printer is now part of the Summary key (Location, Printer, Paper stock, ...)"
    for ($r = 10; $r -le 15; $r++) {
        $v = @()
        # Location, Printer, Paper stock, Unit, Jobs, Qty, Gross, Chargeable
        foreach ($col in 1, 2, 3, 6, 7, 8, 12, 14) { $v += [string]$s.Cells($r, $col).Text }
        if ($v[0] -ne '') { Write-Host ('   ' + ($v -join ' | ')) }
    }

    Write-Host ''
    Write-Host '=== Reports: AT-12 blank-criteria cases (student name/number/date range) ==='
    function Try-Criteria($name, $num, $from, $to, $label) {
        function Set-Crit($cell, $v) {
            if ($null -eq $v -or "$v" -eq '') { $cell.ClearContents() | Out-Null; return }
            if ($v -is [datetime]) {
                $cell.Formula = ('=DATE({0},{1},{2})' -f $v.Year, $v.Month, $v.Day)
                return
            }
            $cell.Value2 = [string]$v
        }
        Set-Crit $c.Range('B4') $name
        Set-Crit $c.Range('B5') $num
        Set-Crit $c.Range('B7') $from
        Set-Crit $c.Range('B8') $to
        $xl.CalculateFullRebuild()
        $jobs = $c.Range('B13').Text
        # "Matching" totals moved from a label-left-of-value row (B/D/F/H/J/L)
        # to a label-above-value layout (row 12/13, same column) on
        # 2026-09-22 - the six metrics only ever sit on columns the
        # minimum-columns view (snag 2d) never hides, so Chargeable is F13
        # now, not H13.
        $charge = $c.Range('F13').Text
        $warn = [string]$c.Range('A9').Text
        $line = "  {0,-34} jobs={1,-4} chargeable={2,-10}" -f $label, $jobs, $charge
        if ($warn -ne '') { $line += " WARN: $warn" }
        Write-Host $line
    }

    $d15 = Get-Date '2026-09-15'
    $d16 = Get-Date '2026-09-16'
    $d17 = Get-Date '2026-09-17'

    Try-Criteria '' '' '' ''                      'all blank (everything)'
    Try-Criteria 'smith' '' '' ''                 'name only, partial + lowercase'
    Try-Criteria '' '2203110' '' ''               'number only'
    Try-Criteria '' '' $d16 ''                    'from date only'
    Try-Criteria '' '' '' $d16                    'to date only'
    Try-Criteria '' '' $d16 $d16                  'single-day range (incl. times)'
    Try-Criteria 'Patel' '' $d15 $d17             'name + both dates'
    Try-Criteria 'Smith' '2203110' '' ''          'mismatched name/number (14.1)'
    Try-Criteria 'Nobody' '' '' ''                'no matches'
    Try-Criteria '' '' '16/09/2026' ''            'from date as TEXT (robustness)'

    # B6 is a deliberate blank gap (locked, unlike the input cells around it)
    # between the Student/Department pair and the date pair - not included
    # here, since ClearContents on a range spanning a locked cell fails
    # outright on a protected sheet.
    $c.Range('B4:B5').ClearContents() | Out-Null
    $c.Range('B7:B8').ClearContents() | Out-Null

    Write-Host ''
    Write-Host '=== Reports: new filters (Technician / Printer / Paper Stock / Quantity) ==='
    Write-Host '  (F5/F6/F7 are now dropdowns - Technician exact-matches recorded jobs, Printer/Paper stock list the whole active catalogue)'
    Write-Host '  (2026-09-27: Technician/Printer/Paper Stock/Quantity shifted down one row, F4:F7 -> F5:F8, to make room for Location at F4)'
    $c.Range('F5').Value2 = 'J. Okonkwo'
    $xl.CalculateFullRebuild()
    Write-Host ("  Technician = 'J. Okonkwo': jobs={0}" -f $c.Range('B13').Text)
    if ([int]$c.Range('B13').Text -eq 0) { Write-Host 'FAIL: expected at least one match on Technician filter'; exit 1 }
    $c.Range('F5').ClearContents() | Out-Null

    $c.Range('F6').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    Write-Host ("  Printer = 'Epson SureColor P9500': jobs={0}" -f $c.Range('B13').Text)
    if ([int]$c.Range('B13').Text -eq 0) { Write-Host 'FAIL: expected at least one match on Printer filter'; exit 1 }
    $c.Range('F6').ClearContents() | Out-Null

    $c.Range('F8').Value2 = 12
    $xl.CalculateFullRebuild()
    Write-Host ("  Quantity = 12: jobs={0}" -f $c.Range('B13').Text)
    if ([int]$c.Range('B13').Text -eq 0) { Write-Host 'FAIL: expected at least one match on Quantity filter'; exit 1 }
    $c.Range('F8').ClearContents() | Out-Null
    $xl.CalculateFullRebuild()

    Write-Host ''
    Write-Host '=== Reports: dropdown/date-picker UI controls ==='
    foreach ($addr in 'F5', 'F6', 'F7', 'B10', 'F10') {
        $t = $c.Range($addr).Validation.Type
        Write-Host ("  {0} validation type: {1} (3 = list/dropdown)" -f $addr, $t)
        if ($t -ne 3) { Write-Host "FAIL: $addr should be a list dropdown"; exit 1 }
    }
    foreach ($addr in 'B7', 'B8') {
        $t = $c.Range($addr).Validation.Type
        Write-Host ("  {0} validation type: {1} (4 = date; the native calendar picker is web-only and does not appear on desktop Excel)" -f $addr, $t)
        if ($t -ne 4) { Write-Host "FAIL: $addr should be Date-type validation"; exit 1 }
    }
    Write-Host '  OK: Technician/Printer/Paper Stock/Sort by/Sort direction are dropdowns; From/To date use Date validation'

    Write-Host ''
    Write-Host '=== Reports: buttons within the first screenful ==='
    foreach ($b in $c.Buttons()) {
        Write-Host ("  {0,-34} left={1}px" -f $b.Name, [int]$b.Left)
        if ([int]$b.Left -gt 900) { Write-Host "FAIL: $($b.Name) is positioned too far right ($([int]$b.Left)px)"; exit 1 }
    }

    Write-Host ''
    Write-Host '=== Reports: sort by column ==='
    $c.Range('B10').Value2 = 'Qty'
    $c.Range('F10').Value2 = 'Descending'
    $xl.CalculateFullRebuild()
    $sp = $c.Range('A16').SpillingToRange
    $qtyCol = ReportsCol 'Qty'
    $qtys = @()
    for ($r = 1; $r -le $sp.Rows.Count; $r++) { $qtys += [double]$sp.Cells($r, $qtyCol).Value2 }
    Write-Host ('  Qty column, sorted Descending: ' + ($qtys -join ', '))
    $sortedDesc = $qtys | Sort-Object -Descending
    $matches = $true
    for ($i = 0; $i -lt $qtys.Count; $i++) { if ($qtys[$i] -ne $sortedDesc[$i]) { $matches = $false } }
    if (-not $matches) { Write-Host 'FAIL: results were not sorted by Qty Descending'; exit 1 }
    Write-Host '  OK: SORTBY reordered the results as requested'
    $c.Range('B10').ClearContents() | Out-Null
    $c.Range('F10').ClearContents() | Out-Null
    $xl.CalculateFullRebuild()

    Write-Host ''
    Write-Host '=== breakdowns (no criteria) ==='
    # Shifted from Q16/U16 to T16/X16 on 2026-09-22 - the results table grew
    # by 3 columns (Student Name/No, Paid), which pushed the old Q15 start
    # into the table itself.
    foreach ($addr in 'T16', 'X16') {
        Write-Host ("  {0}:" -f $addr)
        try {
            $sp = $c.Range($addr).SpillingToRange
            for ($r = 1; $r -le [Math]::Min($sp.Rows.Count, 6); $r++) {
                $v = @()
                for ($k = 1; $k -le $sp.Columns.Count; $k++) { $v += [string]$sp.Cells($r, $k).Text }
                Write-Host ('     ' + ($v -join ' | '))
            }
        } catch { Write-Host ('     NO SPILL: ' + $c.Range($addr).Text) }
    }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'PASS'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
