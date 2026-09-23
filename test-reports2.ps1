# Snag list items 2a (student name/no visibility toggle) and 2d (Reports
# minimum-columns view). Package 6a of the 2026-09-22 post-phase-8 snag list.
#
# Drives the toggle cell and inspects the results table, then closes WITHOUT
# saving.

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
    $c = $wb.Worksheets('Reports')

    function ReportsCol($header) {
        for ($col = 1; $col -le 30; $col++) {
            if ([string]$c.Cells(15, $col).Text -eq $header) { return $col }
        }
        return 0
    }
    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }

    # -------------------------------------------------- 2a: default state
    Write-Host '=== Student name/no toggle: defaults to No (data protection) ==='
    Check ([string]$c.Range('O10').Text -eq 'No') "O10 defaults to No (got '$($c.Range('O10').Text)')"
    $xl.CalculateFullRebuild()
    $nameCol = ReportsCol 'Student name'
    $noCol = ReportsCol 'Student no'
    Check ($nameCol -gt 0) "Student name column exists in the results table (col $nameCol)"
    Check ($noCol -gt 0) "Student no column exists in the results table (col $noCol)"
    $sp = $c.Range('A16').SpillingToRange
    $blank = $true
    for ($r = 1; $r -le $sp.Rows.Count; $r++) {
        if ([string]$sp.Cells($r, $nameCol).Value2 -ne '' -or [string]$sp.Cells($r, $noCol).Value2 -ne '') { $blank = $false }
    }
    Check $blank "every row's Student name/no is blank while the toggle is No"

    # ------------------------------------------------------- 2a: toggled on
    Write-Host ''
    Write-Host '=== Student name/no toggle: Yes reveals them ==='
    $c.Range('O10').Value2 = 'Yes'
    $xl.CalculateFullRebuild()
    $sp = $c.Range('A16').SpillingToRange
    $anyName = $false
    for ($r = 1; $r -le $sp.Rows.Count; $r++) {
        if ([string]$sp.Cells($r, $nameCol).Value2 -ne '') { $anyName = $true }
    }
    Check $anyName "at least one row shows a Student name once the toggle is Yes"
    $sampleName = [string]$sp.Cells(1, $nameCol).Value2
    Write-Host "  sample: '$sampleName'"

    $c.Range('O10').Value2 = 'No'
    $xl.CalculateFullRebuild()

    # ------------------------------------------------------------- 2d
    Write-Host ''
    Write-Host '=== Minimum-columns view: only the documented set stays visible ==='
    $shouldShow = 'Date/Time', 'Location', 'Student name', 'Student no', 'Paper stock', 'Chargeable', 'Paid'
    $shouldHide = 'Printer', 'Qty', 'Unit', 'Area m2', 'Paper cost', 'Consumable cost', 'Gross', 'Disregarded', 'Technician', 'Notes'
    foreach ($h in $shouldShow) {
        $col = ReportsCol $h
        $letter = ([char](64 + $col)).ToString()
        Check (-not $c.Columns($letter).Hidden) "$h ($letter) is visible"
    }
    foreach ($h in $shouldHide) {
        $col = ReportsCol $h
        $letter = ([char](64 + $col)).ToString()
        Check ([bool]$c.Columns($letter).Hidden) "$h ($letter) is hidden by default"
    }

    # --------------------------------------------------------- Paid column
    Write-Host ''
    Write-Host '=== Paid appears in the results table ==='
    $paidCol = ReportsCol 'Paid'
    Check ($paidCol -gt 0) "Paid column exists in the results table (col $paidCol)"
    $sp = $c.Range('A16').SpillingToRange
    $sample = [string]$sp.Cells(1, $paidCol).Value2
    Write-Host "  row 1 Paid = '$sample'"
    # This sample workbook's rows predate the Paid column (blank = unpaid,
    # §5) - the display must show blank, not the literal "0" INDEX returns
    # for a genuinely-empty cell without &"" coercion in the formula.
    Check ($sample -ne '0') "blank Paid displays as blank, not the literal '0' (got '$sample')"

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
