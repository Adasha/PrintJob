# Phase 5: Summary and Cost Calculations.
#
# AT-10  a student's total across every print room in this workbook
# AT-12  the four blank-criteria cases
#
# Drives the criteria cells, then closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself.
#
# Neither Close($false) nor AutoSaveOn is the protection it looks like.
# Workbook_Open does real work on every open - ProtectAll unprotects and
# reprotects all fourteen sheets, plus Invalidate and HealButtons - so the
# workbook is dirty the instant it opens. Excel holds this file through its
# cloud-backed handle (Workbook.Path returns a d.docs.live.net URL), so
# AutoSave is on and COMMITS that immediately: $wb.Saved reads True on the
# very next line after Open, despite every sheet having just been modified.
# The bytes land when the handle closes - measured, the timestamp and the
# SHA-256 both move at Close/Quit, not at Open.
#
# Which is exactly why Close($false) does not help. It means "discard unsaved
# changes", and by then there are none: AutoSave has accepted them, so what
# is written at close is committed state. Setting AutoSaveOn = $false after
# Open is too late for the same reason, and Open offers no earlier hook - so
# no guard is attempted here, because none can work. (build.ps1 sets it, but
# only ever on a %TEMP% copy, which is not cloud-backed; its own comment notes
# it is usually a no-op there.) Tested: AutoSaveOn False, Saved forced True,
# nothing touched at all - the file still changes.
#
# None of this is the sync client. It reproduces with syncing paused and all
# sync operations settled. It is Excel-side only.
#
# verify.ps1 and probe.ps1 take the other way out: read-only, so Workbook_Open
# cannot save anything. These scripts need read-write to drive mutating
# macros, so a throwaway copy is what is left. Discarded at the end.
#
# Left unfixed, this rewrites the artefact under test: the phase 5-7 runs did,
# and a later script read an earlier one's edits back as though they were the
# build.
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
    $c = $wb.Worksheets('Cost Calculations')

    Write-Host '=== Summary ==='
    Write-Host ("  A10 formula length: {0} chars" -f $s.Range('A10').Formula2.Length)
    try {
        $sp = $s.Range('A10').SpillingToRange
        Write-Host ("  spill: {0} rows x {1} cols" -f $sp.Rows.Count, $sp.Columns.Count)
    } catch { Write-Host ('  NO SPILL: ' + $s.Range('A10').Text) }
    Write-Host ("  totals: jobs={0} gross={1} disregarded={2} chargeable={3}" -f `
        $s.Range('B6').Text, $s.Range('D6').Text, $s.Range('F6').Text, $s.Range('H6').Text)
    for ($r = 9; $r -le 14; $r++) {
        $v = @()
        foreach ($col in 1, 2, 5, 6, 7, 11, 13) { $v += [string]$s.Cells($r, $col).Text }
        if ($v[0] -ne '') { Write-Host ('   ' + ($v -join ' | ')) }
    }

    Write-Host ''
    Write-Host '=== Cost Calculations: AT-12 blank-criteria cases ==='
    function Try-Criteria($name, $num, $from, $to, $label) {
        # Two COM traps here. .Value2 is the serial-number variant and rejects
        # a DateTime outright; and once PowerShell has bound that setter with
        # a String it refuses a Double on later calls. So: dates go in as a
        # DATE() formula - a real date value, and locale-independent, unlike
        # typing "16/09/2026" - and everything else as a string.
        function Set-Crit($cell, $v) {
            if ($null -eq $v -or "$v" -eq '') { $cell.ClearContents() | Out-Null; return }
            if ($v -is [datetime]) {
                $cell.Formula = ('=DATE({0},{1},{2})' -f $v.Year, $v.Month, $v.Day)
                return
            }
            $cell.Value2 = [string]$v
        }
        Set-Crit $c.Range('C5') $name
        Set-Crit $c.Range('C6') $num
        Set-Crit $c.Range('C7') $from
        Set-Crit $c.Range('C8') $to
        $xl.CalculateFullRebuild()
        $jobs = $c.Range('B12').Text
        $charge = $c.Range('H12').Text
        $warn = [string]$c.Range('A9').Text
        $line = "  {0,-34} jobs={1,-4} chargeable={2,-10}" -f $label, $jobs, $charge
        if ($warn -ne '') { $line += " WARN: $warn" }
        Write-Host $line
    }

    # Real DateTime values, as a user typing into a date-formatted cell gets.
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
    # A date left as text, which is what an unformatted cell can produce.
    Try-Criteria '' '' '16/09/2026' ''            'from date as TEXT (robustness)'

    # Parentheses matter: $range.ClearContents without them returns the method
    # definition and clears nothing, so the breakdowns below were silently
    # being read with the previous test's criteria still applied.
    $c.Range('C5:C8').ClearContents() | Out-Null
    $xl.CalculateFullRebuild()
    Write-Host ("  (criteria cleared: C5..C8 = '{0}','{1}','{2}','{3}')" -f `
        $c.Range('C5').Text, $c.Range('C6').Text, $c.Range('C7').Text, $c.Range('C8').Text)

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
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
