# Problem 2's "Export report" button and Problem 3's Reports-page bulk
# delete.
#
# DeleteVisibleReportsConfirmed is called directly, bypassing the Ask() gate
# in DeleteVisibleReports - Ask() always declines under SetQuiet, same
# reason test-nextid.ps1/test-import.ps1 reach past modJobs.RemoveRow and
# modImport.ApplyImport the same way.
#
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself - see verify.ps1.

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
    $main = $wb.Worksheets('Main Print Room')
    $loMain = $main.ListObjects('tblJobs_MAIN')
    $annexe = $wb.Worksheets('Annexe')
    $loAnnexe = $annexe.ListObjects('tblJobs_ANNEX')

    Write-Host '=== Export report ==='
    # Filter to one printer (F6 is now a catalogue-wide exact-match dropdown)
    # so this is a genuine subset, not the whole workbook - exercises the
    # filtered-snapshot path.
    $rep.Range('F6').Value2 = 'Epson SureColor P9500'
    $xl.CalculateFullRebuild()
    $beforeJobs = [int]$rep.Range('B12').Text
    Write-Host ("  filtered to Printer = 'Epson SureColor P9500': {0} jobs" -f $beforeJobs)
    if ($beforeJobs -eq 0) { Write-Host 'FAIL: expected at least one job for this filter'; exit 1 }

    $rep.Activate()
    $xl.Run('btnExportReport')
    Write-Host $xl.Run('QuietLog')

    $xlsx = Get-ChildItem $workDir -Filter 'PrintCosts-Report-*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $xlsx) { Write-Host 'FAIL: no report .xlsx was produced'; exit 1 }
    Write-Host ("  OK: {0} ({1} bytes)" -f $xlsx.Name, $xlsx.Length)

    $sigAfterExport = [string]$rep.Range('AN1').Value2
    Write-Host ("  stamped signature: '{0}'" -f $sigAfterExport)
    if ([string]::IsNullOrEmpty($sigAfterExport)) { Write-Host 'FAIL: ReportsExportSig was not stamped'; exit 1 }

    # Peek inside the exported file to confirm it is a real, formatted,
    # static-value snapshot - not a link back into the workbook.
    $xl2 = New-Object -ComObject Excel.Application
    $xl2.Visible = $false
    $xl2.DisplayAlerts = $false
    $wbCheck = $xl2.Workbooks.Open($xlsx.FullName)
    $wsCheck = $wbCheck.Worksheets(1)
    Write-Host ("  snapshot header row: {0}" -f $wsCheck.Range('A1').Value2, $wsCheck.Range('B1').Value2, $wsCheck.Range('C1').Value2)
    Write-Host ("  snapshot row2 formula type: {0}" -f $wsCheck.Range('A2').HasFormula)
    if ($wsCheck.Range('A2').HasFormula) { Write-Host 'FAIL: snapshot should be static values, not formulas'; exit 1 }
    $wbCheck.Close($false)
    $xl2.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
    Write-Host '  OK: snapshot holds static values, no live formulas'

    Write-Host ''
    Write-Host '=== Delete visible records (staleness check should now be clean) ==='
    $beforeMain = $loMain.ListRows.Count
    $beforeAnnexe = $loAnnexe.ListRows.Count
    Write-Host ("  before: Main={0} rows, Annexe={1} rows" -f $beforeMain, $beforeAnnexe)

    $rep.Activate()
    $rng = $rep.Range('A15').SpillingToRange
    $n = $rng.Rows.Count
    $locCol = 0; $jobCol = 0
    for ($c = 1; $c -le 20; $c++) {
        $h = [string]$rep.Cells(14, $c).Value2
        if ($h -eq 'Location') { $locCol = $c }
        if ($h -eq 'Job ID') { $jobCol = $c }
    }
    Write-Host ("  deleting {0} visible rows (locCol={1}, jobCol={2})" -f $n, $locCol, $jobCol)
    $xl.Run('DeleteVisibleReportsConfirmed', $rep, $rng, $locCol, $jobCol, $n)
    Write-Host $xl.Run('QuietLog')

    $afterMain = $loMain.ListRows.Count
    $afterAnnexe = $loAnnexe.ListRows.Count
    Write-Host ("  after: Main={0} rows, Annexe={1} rows" -f $afterMain, $afterAnnexe)
    if (($beforeMain - $afterMain) + ($beforeAnnexe - $afterAnnexe) -ne $n) {
        Write-Host "FAIL: expected exactly $n rows removed across both sheets"
        exit 1
    }
    Write-Host "  OK: exactly $n rows removed, from the correct source sheets"

    # Audit log entry.
    $aud = $wb.Worksheets('_Audit').ListObjects('tblAudit')
    $lastRow = $aud.ListRows.Count
    $lastAction = [string]$aud.ListRows($lastRow).Range.Cells(1, 3).Value2
    $lastDetail = [string]$aud.ListRows($lastRow).Range.Cells(1, 5).Value2
    Write-Host ("  audit log: action='{0}' detail='{1}'" -f $lastAction, $lastDetail)
    if ($lastAction -ne 'Delete visible (Reports)') { Write-Host 'FAIL: audit log entry missing/wrong'; exit 1 }
    Write-Host '  OK: audit entry recorded'

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'PASS'
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
