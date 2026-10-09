# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Summary and Reports, end to end (merges the former test-reports,
# test-reports2, test-reportsnapshot and test-summaryfixes):
#
#   1. Reports: hidden Job ID column; Summary key (Location x Printer x Paper
#      Stock); AT-10 / AT-12 criteria cases; Technician/Printer/Paper
#      Stock filters; dropdown and date validation; buttons within the
#      first screenful; sort-by-column; breakdowns.
#   2. Export-names prompt (live view always shows names), the Reports
#      minimum-columns view, and the Paid column.
#   3. Summary sheet: no "Go to Settings" button; ToggleConfigSheets stays on
#      Summary.
#   4. Export report: header block, single-value promotion (Printer promoted
#      when filtered, nothing promoted when not), exported name/no blank under
#      redacted to hyphens when the prompt is answered No, populated under Yes. Needs a SECOND location
#      sheet (test-fixture-annexe.ps1): PrintCosts.xlsx ships only "Example
#      Print Room", so with one room every row's Location is trivially uniform
#      and PromoteUniformColumns correctly promotes and drops it even when
#      unfiltered. The fixture is added up front for the whole script.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself (AutoSave on a
# OneDrive-backed handle commits regardless of Close($false)). Closes WITHOUT
# saving.

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
    $wb = $xl.Workbooks.Open($f)
    Add-AnnexeFixture $xl $wb | Out-Null
    # The fixture leaves Excel recalculating; without this the first criteria
    # write below is rejected (RPC_E_CALL_REJECTED).
    $xl.CalculateFullRebuild()
    Start-Sleep -Seconds 3
    $xl.Run('SetQuiet', $true)
    # Excel can still be settling from Workbook_Open's own macro work when
    # this lands - retried the same way build.ps1 retries its SaveAs.
    $s = Invoke-ComRetry { $wb.Worksheets('Summary') }
    $c = $wb.Worksheets('Reports')
    $rep = $c

    # Reports' results columns move whenever one is added - found by header
    # (row 10, from column H; returns the position within the table) rather than hardcoded.
    function ReportsCol($header) {
        for ($col = 8; $col -le 38; $col++) {
            if ([string]$c.Cells(10, $col).Text -eq $header) { return $col - 7 }
        }
        return 0
    }
    Write-Host '=== Reports: hidden Job ID correlation column ==='
    $jobIdCol = ReportsCol 'Job ID'
    Check (-not ($jobIdCol -eq 0)) 'no Job ID header found on row 10'
    # [char] on its own is a .NET Char, and Excel's COM Columns(...) indexer
    # silently takes THAT as a numeric column index (its ordinal value, e.g.
    # 82 for 'R') rather than the single-letter column reference it looks
    # like - .ToString() forces it to a real string first. Found by
    # comparing Columns($jobIdColLetter) against the literal Columns('R')
    # side by side: same column, different (wrong) answer without this cast.
    $jobIdColLetter = $c.Cells(1, 7 + $jobIdCol).Address($false, $false) -replace '\d+', ''
    Check ([bool]$c.Cells(1, 7 + $jobIdCol).EntireColumn.Hidden) "column $jobIdColLetter (Job ID) should be hidden"
    $sp = $c.Range('H11').SpillingToRange
    Write-Host ("  results spill: {0} rows x {1} cols" -f $sp.Rows.Count, $sp.Columns.Count)
    $jobIdSample = [string]$sp.Cells(1, $jobIdCol).Value2
    Write-Host ("  {0}16 (Job ID) sample value: '{1}'" -f $jobIdColLetter, $jobIdSample)
    Check (-not ($jobIdSample -notmatch '-MAIN-|-ANNEX-')) 'hidden Job ID column does not look like a Job ID'
    Write-Host "  OK: Job ID is present as a hidden column ($jobIdColLetter)"

    Write-Host ''
    Write-Host '=== Summary (Location x Printer x Paper Stock) ==='
    Write-Host ("  Summary table formula length: {0} chars" -f $s.Range("A$SUM_FIRST_ROW").Formula2.Length)
    try {
        $sp = $s.Range("A$SUM_FIRST_ROW").SpillingToRange
        Write-Host ("  spill: {0} rows x {1} cols" -f $sp.Rows.Count, $sp.Columns.Count)
    } catch { Write-Host ('  NO SPILL: ' + $s.Range("A$SUM_FIRST_ROW").Text) }
    Write-Host ("  totals: jobs={0} gross={1} disregarded={2} chargeable={3}" -f `
        $s.Range("B$SUM_TOT_ROW").Text, $s.Range("D$SUM_TOT_ROW").Text, $s.Range("F$SUM_TOT_ROW").Text, $s.Range("H$SUM_TOT_ROW").Text)
    $hdrRow = @()
    for ($col = 1; $col -le 13; $col++) { $hdrRow += [string]$s.Cells($SUM_HDR_ROW, $col).Text }
    Write-Host ('  headers: ' + ($hdrRow -join ' | '))
    Check ($hdrRow[1] -eq 'Printer') "Summary column B header is 'Printer' (got '$($hdrRow[1])')"
    Write-Host "  OK: Printer is now part of the Summary key (Location, Printer, Paper stock, ...)"
    for ($r = 10; $r -le 15; $r++) {
        $v = @()
        # Location, Printer, Paper stock, Unit, Jobs, Qty, Gross, Chargeable
        foreach ($col in 1, 2, 3, 5, 6, 7, 11, 13) { $v += [string]$s.Cells($r, $col).Text }
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
        $jobs = $c.Range('I9').Text
        # "Matching" totals moved from a label-left-of-value row (B/D/F/H/J/L)
        # to a label-above-value layout (row 14/15, same column) on
        # 2026-09-22 - the six metrics only ever sit on columns the
        # minimum-columns view (snag 2d) never hides, so Chargeable is F13
        # now, not H13.
        $charge = $c.Range('N9').Text
        $warn = [string]$c.Range('A13').Text
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
    Write-Host '=== Reports: new filters (Technician / Printer / Paper Stock) ==='
    Write-Host '  (F5/F6/F7 are now dropdowns - Technician exact-matches recorded jobs, Printer/Paper stock list the whole active catalogue)'
    Write-Host '  (2026-09-27: Technician/Printer/Paper Stock shifted down one row, F4:F6 -> F5:F7, to make room for Location at F4)'
    $c.Range('F5').Value2 = 'J. Okonkwo'
    $xl.CalculateFullRebuild()
    Write-Host ("  Technician = 'J. Okonkwo': jobs={0}" -f $c.Range('I9').Text)
    Check (-not ([int]$c.Range('I9').Text -eq 0)) 'expected at least one match on Technician filter'
    $c.Range('F5').ClearContents() | Out-Null

    $c.Range('F7').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    Write-Host ("  Printer = 'Epson SureColor P9500': jobs={0}" -f $c.Range('I9').Text)
    Check (-not ([int]$c.Range('I9').Text -eq 0)) 'expected at least one match on Printer filter'
    $c.Range('F7').ClearContents() | Out-Null

    Write-Host ''
    Write-Host '=== Reports: filter block layout (2026-10-02) ==='
    $expect = [ordered]@{
        A4 = 'Student/dept. name'; A5 = 'Student number'; A6 = 'Department'; A7 = 'From date'; A8 = 'To date'; A10 = 'Min cost'; A11 = 'Max cost'; A12 = 'Paid'
        D4 = 'Location (print room)'; D5 = 'Technician'; D7 = 'Printer'; D8 = 'Paper stock'; D9 = 'Paper type'; D10 = 'Student-supplied paper'
        A14 = 'Has a problem'; D14 = 'Has notes'; A15 = 'Disregarded'
        H7 = 'Sort by'; J7 = 'Sort direction'
    }
    foreach ($k in $expect.Keys) {
        Check ([string]$c.Range($k).Text -eq $expect[$k]) "$k reads '$($expect[$k])' (got '$($c.Range($k).Text)')"
    }
    foreach ($k in 'D6', 'A9', 'B9', 'A16', 'D16') {
        Check ([string]$c.Range($k).Text -eq '') "$k is a blank gap row (got '$($c.Range($k).Text)')"
    }

    Check ([string]$c.Range('A12').Text -eq 'Paid' -and [string]$c.Range('A11').Text -eq 'Max cost') 'Min cost, Max cost and Paid are consecutive rows (no gap above Paid, 0.10.26)'

    Write-Host ''
    Write-Host '=== Reports: Paid and Paper type filters ==='
    $xl.CalculateFullRebuild()
    $all = [int]$c.Range('I9').Text
    $c.Range('B12').Value2 = 'Yes'; $xl.CalculateFullRebuild(); $yes = [int]$c.Range('I9').Text
    $c.Range('B12').Value2 = 'No';  $xl.CalculateFullRebuild(); $no = [int]$c.Range('I9').Text
    $c.Range('B12').ClearContents() | Out-Null
    Write-Host ("  all={0} paid=Yes: {1}  paid=No: {2}" -f $all, $yes, $no)
    Check (($yes + $no) -eq $all) 'Paid Yes + No should account for every job (blank counts as No)'
    $c.Range('F9').Value2 = 'Roll';  $xl.CalculateFullRebuild(); $roll = [int]$c.Range('I9').Text
    $c.Range('F9').Value2 = 'Sheet'; $xl.CalculateFullRebuild(); $sheet = [int]$c.Range('I9').Text
    $c.Range('F9').ClearContents() | Out-Null
    Write-Host ("  paper type Roll: {0}  Sheet: {1}" -f $roll, $sheet)
    Check (($roll + $sheet) -eq $all) 'Roll + Sheet should account for every job'
    Check ($roll -gt 0 -and $sheet -gt 0) 'fixture should contain both roll and sheet jobs'

    $xl.CalculateFullRebuild()

    Write-Host ''
    Write-Host '=== Reports: dropdown/date-picker UI controls ==='
    foreach ($addr in 'F5', 'F7', 'F8', 'I7', 'K7', 'B12', 'F9') {
        $t = $c.Range($addr).Validation.Type
        Write-Host ("  {0} validation type: {1} (3 = list/dropdown)" -f $addr, $t)
        Check (-not ($t -ne 3)) "$addr should be a list dropdown"
    }
    foreach ($addr in 'B7', 'B8') {
        $t = $c.Range($addr).Validation.Type
        Write-Host ("  {0} validation type: {1} (4 = date; the native calendar picker is web-only and does not appear on desktop Excel)" -f $addr, $t)
        Check (-not ($t -ne 4)) "$addr should be Date-type validation"
    }
    Write-Host '  OK: Technician/Printer/Paper Stock/Paper type/Paid/Sort by/Sort direction are dropdowns; From/To date use Date validation'

    Write-Host ''
    Write-Host '=== Reports: buttons within the first screenful ==='
    foreach ($b in $c.Buttons()) {
        Write-Host ("  {0,-34} left={1}px" -f $b.Name, [int]$b.Left)
        Check (-not ([int]$b.Left -gt 900)) "$($b.Name) is positioned too far right ($([int]$b.Left)px)"
    }

    Write-Host ''
    Write-Host '=== Reports: sort by column ==='
    $c.Range('I7').Value2 = 'Qty'
    $c.Range('K7').Value2 = 'Descending'
    $xl.CalculateFullRebuild()
    $sp = $c.Range('H11').SpillingToRange
    $qtyCol = ReportsCol 'Qty'
    $qtys = @()
    for ($r = 1; $r -le $sp.Rows.Count; $r++) { $qtys += [double]$sp.Cells($r, $qtyCol).Value2 }
    Write-Host ('  Qty column, sorted Descending: ' + ($qtys -join ', '))
    $sortedDesc = $qtys | Sort-Object -Descending
    $matches = $true
    for ($i = 0; $i -lt $qtys.Count; $i++) { if ($qtys[$i] -ne $sortedDesc[$i]) { $matches = $false } }
    Check (-not (-not $matches)) 'results were not sorted by Qty Descending'
    Write-Host '  OK: SORTBY reordered the results as requested'
    $c.Range('I7').ClearContents() | Out-Null
    $c.Range('K7').ClearContents() | Out-Null
    $xl.CalculateFullRebuild()

    Write-Host ''
    Write-Host '=== breakdowns (no criteria) ==='
    # Shifted from Q16/U16 to T16/X16 on 2026-09-22 - the results table grew
    # by 3 columns (Student Name/No, Paid), which pushed the old Q15 start
    # into the table itself.
    foreach ($addr in 'AA11', 'AE11') {
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

    # ============================================== Export names / min columns
    # -------------------------------------------------- 2a: default state
    Write-Host '=== Export names toggle is gone (2026-10-02); live view always shows names ==='
    Check ([string]$c.Range('AQ12').Text -eq '' -and [string]$c.Range('AR12').Text -eq '') 'no Export names label or dropdown at AQ12:AR12'
    $xl.CalculateFullRebuild()
    $nameCol = ReportsCol 'Student name'
    $noCol = ReportsCol 'Student no'
    Check ($nameCol -gt 0) "Student name column exists in the results table (col $nameCol)"
    Check ($noCol -gt 0) "Student no column exists in the results table (col $noCol)"
    $sp = $c.Range('H11').SpillingToRange
    $anyNameNo = $false
    for ($r = 1; $r -le $sp.Rows.Count; $r++) {
        if ([string]$sp.Cells($r, $nameCol).Value2 -ne '') { $anyNameNo = $true }
    }
    Check $anyNameNo "the live results still show Student names in the live view"
    $zeroNo = 0
    for ($r = 1; $r -le $sp.Rows.Count; $r++) { if ([string]$sp.Cells($r, $noCol).Text -eq '0') { $zeroNo++ } }
    Check ($zeroNo -eq 0) "a job with no student number shows blank, not '0' ($zeroNo rows show 0)"

    # ------------------------------------------------------- 2a: toggled on
    Write-Host ''
    Write-Host '=== Live view shows names (again, after a rebuild) ==='
    $xl.CalculateFullRebuild()
    $sp = $c.Range('H11').SpillingToRange
    $anyName = $false
    for ($r = 1; $r -le $sp.Rows.Count; $r++) {
        if ([string]$sp.Cells($r, $nameCol).Value2 -ne '') { $anyName = $true }
    }
    Check $anyName "at least one row shows a Student name in the live view"
    $sampleName = [string]$sp.Cells(1, $nameCol).Value2
    Write-Host "  sample: '$sampleName'"

    $xl.CalculateFullRebuild()

    # ------------------------------------------------------------- 2d
    Write-Host ''
    Write-Host '=== Minimum-columns view: only the documented set stays visible ==='
    $shouldShow = 'Date/Time', 'Location', 'Student name', 'Student no', 'Paper stock', 'Chargeable', 'Paid'
    $shouldHide = 'Printer', 'Qty', 'Unit', 'Area m2', 'Paper cost', 'Consumable cost', 'Gross', 'Disregarded', 'Technician', 'Notes'
    foreach ($h in $shouldShow) {
        $col = ReportsCol $h
        $letter = ([char](64 + $col + 7)).ToString()
        Check (-not $c.Columns($letter).Hidden) "$h ($letter) is visible"
    }
    foreach ($h in $shouldHide) {
        $col = ReportsCol $h
        $letter = ([char](64 + $col + 7)).ToString()
        Check ([bool]$c.Columns($letter).Hidden) "$h ($letter) is hidden by default"
    }

    # --------------------------------------------------------- Paid column
    Write-Host ''
    Write-Host '=== Paid appears in the results table ==='
    $paidCol = ReportsCol 'Paid'
    Check ($paidCol -gt 0) "Paid column exists in the results table (col $paidCol)"
    $sp = $c.Range('H11').SpillingToRange
    $sample = [string]$sp.Cells(1, $paidCol).Value2
    Write-Host "  row 1 Paid = '$sample'"
    # This sample workbook's rows predate the Paid column (blank = unpaid,
    # §5) - the display must show blank, not the literal "0" INDEX returns
    # for a genuinely-empty cell without &"" coercion in the formula.
    Check ($sample -ne '0') "blank Paid displays as blank, not the literal '0' (got '$sample')"


    # ======================================================= Summary sheet
    Write-Host '=== 3b: Go to Settings button is gone ==='
    $summary = $wb.Worksheets('Summary')
    $found = $false
    $foundToggle = $false
    for ($i = 1; $i -le $summary.Buttons().Count; $i++) {
        $cap = $summary.Buttons($i).Caption
        if ($cap -eq 'Go to Settings') { $found = $true }
        if ($cap -like '*settings sheets*') { $foundToggle = $true }
    }
    Check (-not $found) "no button captioned 'Go to Settings' on Summary"
    Check $foundToggle "the Hide/Show settings sheets toggle is still present"

    Write-Host ''
    Write-Host '=== 3a: ToggleConfigSheets stays on (or returns to) Summary ==='
    $main = $wb.Worksheets('Example Print Room')
    $main.Activate()
    Check ($wb.ActiveSheet.Name -eq 'Example Print Room') "active sheet is Example Print Room before toggling (got '$($wb.ActiveSheet.Name)')"
    $xl.Run('ToggleConfigSheets')
    Check ($wb.ActiveSheet.Name -eq 'Summary') "active sheet is Summary after ToggleConfigSheets (got '$($wb.ActiveSheet.Name)')"

    $settingsVisible = $wb.Worksheets('Settings').Visible
    Write-Host ("  Settings sheet Visible = $settingsVisible (expect hidden, since sheets started shown)")
    Check ($settingsVisible -ne -1) "the four configuration sheets were actually hidden by the toggle"

    # Toggle back (show again), from Summary itself this time - should still
    # end on Summary (the ordinary, already-passing case).
    $xl.Run('ToggleConfigSheets')
    Check ($wb.ActiveSheet.Name -eq 'Summary') "still on Summary after toggling back (got '$($wb.ActiveSheet.Name)')"
    Check ($wb.Worksheets('Settings').Visible -eq -1) "the four configuration sheets are visible again"


    # ============================================================ Export report
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

    Write-Host '=== Filtered to one printer: Printer promotes ==='
    $xl.Run('SetExportNames', $false)
    # F6, not F5 (2026-09-27: Location took F4, Technician/Printer/Paper
    # Stock shifted down one row).
    $rep.Range('F7').Value2 = 'Epson SureColor P9500'
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
        Check ($info.Lines -contains 'Paid') "Paid line present in the summary"
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
        for ($ci = 1; $ci -le $info.Cols; $ci++) {
            $h = [string]$info.Sheet.Cells($tableRow, $ci).Value2
            if ($h -eq '') { break }
            $tableHeaders += $h
        }
        Write-Host ("  table columns: " + ($tableHeaders -join ', '))
        Check (-not ($tableHeaders -contains 'Printer')) "Printer column dropped from the table"
        Check ($tableHeaders -contains 'Technician') "Technician column still present (was being silently dropped by the old .Hidden-bounded LastVisibleColumn)"
        Check (-not ($tableHeaders -contains 'Notes')) "Notes column excluded from the printed report (2026-10-02)"

        # Printable layout (2026-10-02): A4 landscape, one page wide, header
        # row repeating, page numbers; the Paid summary sits between Total
        # chargeable and Still owed.
        $ps = $info.Sheet.PageSetup
        Check ($ps.Orientation -eq 2) "page is landscape (got $($ps.Orientation))"
        Check ($ps.PaperSize -eq 9) "paper is A4 (got $($ps.PaperSize))"
        Check (($ps.Zoom -eq $false) -and ($ps.FitToPagesWide -eq 1)) 'fits one page wide'
        Check ($ps.PrintTitleRows -eq ('$' + $tableRow + ':$' + $tableRow)) "table header row repeats on every page (got '$($ps.PrintTitleRows)')"
        Check ($ps.RightFooter -like '*Page &P of &N*') 'page numbers in the footer'
        Check ([string]$info.Sheet.Cells($totalRow + 1, 1).Value2 -eq 'Paid') 'Paid row follows Total chargeable'
        Check ([string]$info.Sheet.Cells($totalRow + 2, 1).Value2 -eq 'Still owed') 'Still owed row follows Paid'
        $hdrFill = [int]$info.Sheet.Cells($tableRow, 1).Interior.Color
        Check ($hdrFill -ne 16777215) 'table header row is tinted'
        # Greyscale safety: the darkest tint must stay pale (luma >= 200/255).
        $rr = $hdrFill -band 255; $gg = ($hdrFill -shr 8) -band 255; $bb = ($hdrFill -shr 16) -band 255
        $luma = 0.299 * $rr + 0.587 * $gg + 0.114 * $bb
        Check ($luma -ge 200) "header tint stays light enough to photocopy (luma $([int]$luma))"
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
    $rep.Range('F7').Value2 = ''
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
        for ($ci = 1; $ci -le $info2.Cols; $ci++) {
            $h = [string]$info2.Sheet.Cells($tableRow2, $ci).Value2
            if ($h -eq '') { break }
            $tableHeaders2 += $h
        }
        Write-Host ("  table columns: " + ($tableHeaders2 -join ', '))
        Check ($tableHeaders2 -contains 'Printer') "Printer column present when unfiltered"
        # Location must survive unfiltered too - it is only genuinely uniform
        # here because the Annexe fixture above gives the workbook a second
        # room to spread across. Without that fixture this is the exact bug
        # this test exists to catch: a single-location workbook makes every
        # row's Location trivially uniform, so PromoteUniformColumns silently
        # (and, for a single-location workbook, correctly) drops it - which
        # looked identical to a real export defect until traced here.
        $locationLine2 = $info2.Lines | Where-Object { $_ -like 'Location:*' }
        Check ($null -eq $locationLine2) "Location NOT promoted when unfiltered (two rooms in the results)"
        Check ($tableHeaders2 -contains 'Location') "Location column present when unfiltered"
        # 16, not 17: Area m2 is excluded from the export by default
        # (2026-09-26, modExport.RemoveExcludedColumns) - derivable from Qty
        # and the paper's own dimensions, not something read directly off an
        # export. Still present on the live Reports sheet itself.
        Check (-not ($tableHeaders2 -contains 'Area m2')) "Area m2 excluded from the export by default"
        Check ($tableHeaders2.Count -eq 17) "all 17 exported result columns present - Area m2 and Notes excluded (got $($tableHeaders2.Count))"

        # Export prompt (2026-10-02): answered No for this export (SetExportNames),
        # so Student name VALUES are redacted to a line of 10 hyphens - not blank,
        # not dropped - while Student no is left alone (test-reports2.ps1 covers
        # the live side).
        $nameIdx = [array]::IndexOf($tableHeaders2, 'Student name') + 1
        $noIdx = [array]::IndexOf($tableHeaders2, 'Student no') + 1
        Check ($nameIdx -gt 0 -and $noIdx -gt 0) "Student name/no columns still present in the export"
        $nameVals = 0; $nameHyphens = 0; $noVals = 0
        for ($r = $tableRow2 + 1; $r -lt $info2.Rows; $r++) {   # stops before the Total row
            $nv = [string]$info2.Sheet.Cells($r, $nameIdx).Value2
            if ($nv -ne '') { $nameVals++; if ($nv -eq '----------') { $nameHyphens++ } }
            if ([string]$info2.Sheet.Cells($r, $noIdx).Value2 -ne '') { $noVals++ }
        }
        Check ($nameVals -gt 0 -and $nameVals -eq $nameHyphens) "every exported Student name is the 10-hyphen placeholder ($nameHyphens of $nameVals)"
        Check ($noVals -gt 0) "Student no is NOT redacted ($noVals values kept)"
    }
    finally {
        $wbCheck2.Close($false)
        $xl3.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl3)
    }

    Write-Host ''
    Write-Host '=== Export prompt answered Yes: the export includes names ==='
    Remove-Item $xlsx2.FullName -Force -ErrorAction SilentlyContinue
    $xl.Run('SetExportNames', $true)
    $hSet = $wb.Names.Item('SET_REPORT_HEADING').RefersToRange
    Check ([string]$hSet.Value2 -eq 'Print job report') "Report heading setting exists with the default (got '$($hSet.Value2)')"
    $hSet.Value2 = 'Dept of Design - Print charges'
    $xl.CalculateFullRebuild()
    $rep.Activate()
    $xl.Run('btnExportReport')
    Write-Host $xl.Run('QuietLog')
    Start-Sleep -Milliseconds 300
    $xlsx3 = Get-ChildItem $workDir -Filter 'PrintCosts-Report-*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Check ($null -ne $xlsx3) "a third .xlsx was produced"
    $res3 = OpenAndReadHeader $xlsx3.FullName
    $info3 = $res3[0]; $xl4 = $res3[1]; $wbCheck3 = $res3[2]
    try {
        $tableRow3 = 0
        for ($r = 1; $r -le $info3.Rows; $r++) {
            if ([string]$info3.Sheet.Cells($r, 1).Value2 -eq 'Date/Time') { $tableRow3 = $r; break }
        }
        $tableHeaders3 = @()
        for ($ci = 1; $ci -le $info3.Cols; $ci++) {
            $h = [string]$info3.Sheet.Cells($tableRow3, $ci).Value2
            if ($h -eq '') { break }
            $tableHeaders3 += $h
        }
        $nameIdx3 = [array]::IndexOf($tableHeaders3, 'Student name') + 1
        $anyName3 = $false
        if ($nameIdx3 -gt 0) {
            for ($r = $tableRow3 + 1; $r -lt $info3.Rows; $r++) {
                if ([string]$info3.Sheet.Cells($r, $nameIdx3).Value2 -ne '') { $anyName3 = $true }
            }
        }
        # A single shared name is promoted to a header line instead of a column.
        $promotedName3 = @($info3.Lines | Where-Object { $_ -like 'Student name:*' }).Count -gt 0
        Check ($anyName3 -or $promotedName3) "exported file contains student names when the prompt is answered Yes"
        Check ($info3.Lines[0] -eq 'Dept of Design - Print charges') "the heading comes from the Report heading setting (got '$($info3.Lines[0])')"
    }
    finally {
        $wbCheck3.Close($false)
        $xl4.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl4)
    }
    $xl.Run('SetExportNames', $false)

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Write-Host ''
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
if ($anyFail) { Write-Host 'FAIL'; exit 1 } else { Write-Host 'PASS' }
