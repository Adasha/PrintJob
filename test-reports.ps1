# Summary (regrouped to a Location x Printer x Paper Stock key) and Reports
# (renamed from Cost Calculations, extended with Technician/Printer/Paper
# Stock/Quantity filters, sort-by-column, and a hidden Job ID column).
#
# AT-10  a student's total across every print room in this workbook
# AT-12  the four blank-criteria cases
#
# Drives the criteria cells, then closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself - see verify.ps1's
# comment for why (AutoSave on a OneDrive-backed handle commits regardless of
# Close($false)).
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
    $s = $wb.Worksheets('Summary')
    $c = $wb.Worksheets('Reports')

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
        # Location, Printer, Paper stock, Unit, Jobs, Quantity, Gross, Chargeable
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
        Set-Crit $c.Range('B5') $name
        Set-Crit $c.Range('B6') $num
        Set-Crit $c.Range('B7') $from
        Set-Crit $c.Range('B8') $to
        $xl.CalculateFullRebuild()
        $jobs = $c.Range('B12').Text
        $charge = $c.Range('H12').Text
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

    $c.Range('B5:B8').ClearContents() | Out-Null

    Write-Host ''
    Write-Host '=== Reports: new filters (Technician / Printer / Paper Stock / Quantity) ==='
    Write-Host '  (F5/F6/F7 are now dropdowns - Technician exact-matches recorded jobs, Printer/Paper stock list the whole active catalogue)'
    $c.Range('F5').Value2 = 'J. Okonkwo'
    $xl.CalculateFullRebuild()
    Write-Host ("  Technician = 'J. Okonkwo': jobs={0}" -f $c.Range('B12').Text)
    if ([int]$c.Range('B12').Text -eq 0) { Write-Host 'FAIL: expected at least one match on Technician filter'; exit 1 }
    $c.Range('F5').ClearContents() | Out-Null

    $c.Range('F6').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    Write-Host ("  Printer = 'Epson SureColor P9500': jobs={0}" -f $c.Range('B12').Text)
    if ([int]$c.Range('B12').Text -eq 0) { Write-Host 'FAIL: expected at least one match on Printer filter'; exit 1 }
    $c.Range('F6').ClearContents() | Out-Null

    $c.Range('F8').Value2 = 12
    $xl.CalculateFullRebuild()
    Write-Host ("  Quantity = 12: jobs={0}" -f $c.Range('B12').Text)
    if ([int]$c.Range('B12').Text -eq 0) { Write-Host 'FAIL: expected at least one match on Quantity filter'; exit 1 }
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
    Write-Host '=== Reports: hidden Job ID correlation column ==='
    if (-not $c.Columns('O').Hidden) { Write-Host 'FAIL: column O (Job ID) should be hidden'; exit 1 }
    $sp = $c.Range('A15').SpillingToRange
    Write-Host ("  results spill: {0} rows x {1} cols" -f $sp.Rows.Count, $sp.Columns.Count)
    $jobIdSample = [string]$sp.Cells(1, 15).Value2
    Write-Host ("  O15 (Job ID) sample value: '{0}'" -f $jobIdSample)
    if ($jobIdSample -notmatch '-MAIN-|-ANNEX-') { Write-Host 'FAIL: hidden Job ID column does not look like a Job ID'; exit 1 }
    Write-Host '  OK: Job ID is present as a hidden 15th column'

    Write-Host ''
    Write-Host '=== Reports: sort by column ==='
    $c.Range('B10').Value2 = 'Quantity'
    $c.Range('F10').Value2 = 'Descending'
    $xl.CalculateFullRebuild()
    $sp = $c.Range('A15').SpillingToRange
    $qtys = @()
    for ($r = 1; $r -le $sp.Rows.Count; $r++) { $qtys += [double]$sp.Cells($r, 5).Value2 }
    Write-Host ('  Quantity column, sorted Descending: ' + ($qtys -join ', '))
    $sortedDesc = $qtys | Sort-Object -Descending
    $matches = $true
    for ($i = 0; $i -lt $qtys.Count; $i++) { if ($qtys[$i] -ne $sortedDesc[$i]) { $matches = $false } }
    if (-not $matches) { Write-Host 'FAIL: results were not sorted by Quantity Descending'; exit 1 }
    Write-Host '  OK: SORTBY reordered the results as requested'
    $c.Range('B10').ClearContents() | Out-Null
    $c.Range('F10').ClearContents() | Out-Null
    $xl.CalculateFullRebuild()

    Write-Host ''
    Write-Host '=== breakdowns (no criteria) ==='
    foreach ($addr in 'Q15', 'U15') {
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
