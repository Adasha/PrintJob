# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# The Select printers... picker sheet (_Picker) placed its checkboxes and
# OK/Cancel from the cells' positions BEFORE setting its column widths, and
# the buttons are free-floating, so they stayed wherever the sheet's previous
# widths had put them (found in the 2026-10-30 Mac screenshot session; fixed
# in 0.11.5, modPicker.Show). Here the sheet starts with deliberately wrong
# widths, then every control must sit where the final layout puts it.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT
# saving. mac/test-picker.py is its twin.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f
$FIRST_ROW = 6   # modPicker.FIRST_ROW

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Example Print Room')
    $pick = $wb.Worksheets('_Picker')

    # Stale widths, as a previous use of the sheet could leave them.
    $xl.Run('UnlockSheet', $pick)
    foreach ($c in 1..4) { $pick.Columns($c).ColumnWidth = 20 }
    $xl.Run('RelockSheet', $pick)

    Write-Host '=== Select printers... places its controls after its widths ==='
    $main.Activate()
    $xl.Run('PickPrinters', $main)
    $bLeft = [double]$pick.Cells(1, 2).Left
    $cLeft = [double]$pick.Cells(1, 3).Left
    $cRight = [double]$pick.Cells(1, 4).Left

    $n = 0
    while ($true) {
        $r = $FIRST_ROW + $n
        $cb = $null
        try { $cb = $pick.CheckBoxes("pick_$r") } catch { }
        if (-not $cb) { break }
        $left = [double]$cb.Left; $width = [double]$cb.Width
        Check ($cLeft -le $left -and ($left + $width) -le $cRight) ("the row $r tick sits inside the Include column ({0:N1}-{1:N1} in {2:N1}-{3:N1})" -f $left, ($left + $width), $cLeft, $cRight)
        $n++
    }
    Check ($n -gt 0) "the picker lists printers (got $n)"

    $ok = $null; $cancel = $null
    foreach ($b in $pick.Buttons()) {
        if ($b.OnAction -like '*btnPickerOK') { $ok = $b }
        if ($b.OnAction -like '*btnPickerCancel') { $cancel = $b }
    }
    Check ($ok -and $cancel) 'OK and Cancel are drawn'
    if ($ok -and $cancel) {
        $top = [double]$pick.Rows($FIRST_ROW + $n + 1).Top
        Check ([Math]::Abs($ok.Top - $top) -lt 1 -and [Math]::Abs($cancel.Top - $top) -lt 1) "OK and Cancel sit one row below the list (tops $($ok.Top), $($cancel.Top); row $top)"
        Check ([Math]::Abs($ok.Left - $bLeft) -lt 1) ("OK starts at the Option column (left $($ok.Left), column {0:N1})" -f $bLeft)
        Check (($ok.Left + $ok.Width) -le $cancel.Left) ("Cancel is to the right of OK, not over it (OK ends {0:N1}, Cancel starts $($cancel.Left))" -f ($ok.Left + $ok.Width))
        $right = $cancel.Left + $cancel.Width
        Check ([Math]::Abs($right - $cRight) -lt 1) ("Cancel ends at the Include column's right edge ({0:N1}, column {1:N1})" -f $right, $cRight)
    }

    $xl.Run('btnPickerCancel')
    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'DONE'
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
