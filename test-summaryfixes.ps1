# Snag list items 3a (ToggleConfigSheets stays on Summary) and 3b (Go to
# Settings button removed), package 7 of the 2026-09-22 post-phase-8 snag
# list. Drives a COPY in %TEMP%, never src\PrintCosts.xlsm.

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
    $anyFail = $false
    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
        if (-not $cond) { $script:anyFail = $true }
    }

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
    $main = $wb.Worksheets('Main Print Room')
    $main.Activate()
    Check ($wb.ActiveSheet.Name -eq 'Main Print Room') "active sheet is Main Print Room before toggling (got '$($wb.ActiveSheet.Name)')"
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
