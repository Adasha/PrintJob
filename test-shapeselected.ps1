# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# 0.12.2: a button (or other shape) left selected.
#
# A copied sheet arrives with every shape selected, and Selection is then a
# DrawingObjects, not a Range. Selection.Cells raised 438 in SelectedRow (Repeat,
# Remove, Now) and RemoveCatalogRow. They now go through modUtils.SelectedCell,
# which answers Nothing, so each command says "no row is selected" instead.
#
# Mac port: mac/test-shapeselected.py. Drives a COPY in %TEMP%, never
# src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
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
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $ws = $wb.Worksheets('Example Print Room')
    $lo = $ws.ListObjects('tblJobs_EXAMPLE')
    $cat = $wb.Worksheets('Consumables').ListObjects('tblColours')
    $ws.Activate()
    $btn = $ws.Buttons(1)

    function SelectButton { $ws.Activate(); [void]$btn.Select() }

    Write-Host ''
    Write-Host '=== A button is selected ==='
    SelectButton
    Check ([Microsoft.VisualBasic.Information]::TypeName($xl.Selection) -ne 'Range') 'the button really is the selection'

    $before = $lo.ListRows.Count
    $catBefore = $cat.ListRows.Count
    $stamp = ($lo.ListColumns('Date/Time').DataBodyRange.Value2 | ForEach-Object { "$_" }) -join '|'

    SelectButton
    Check ([int]$xl.Run('SelectedRow', $ws, $lo) -eq 0) 'SelectedRow answers 0 instead of raising 438'
    Check ([int]$xl.Run('SelectedTableRow', $ws, $lo) -eq 0) 'SelectedTableRow answers 0'

    $calls = @(
        @{ Name = 'StampNow';         Args = @($ws) },
        @{ Name = 'RemoveRow';        Args = @($ws) },
        @{ Name = 'RepeatJob';        Args = @($ws) },
        @{ Name = 'AddPass';          Args = @($ws) },
        @{ Name = 'RemovePass';       Args = @($ws) },
        @{ Name = 'RemoveCatalogRow'; Args = @('tblColours') }
    )
    foreach ($c in $calls) {
        SelectButton
        $ok = $true
        try { [void]$xl.Run($c.Name, $c.Args[0]) } catch { $ok = $false; Write-Host "        $($_.Exception.Message)" }
        Check $ok ("{0} with a button selected raises no error" -f $c.Name)
    }

    Check ($lo.ListRows.Count -eq $before) 'no print job row was added or removed'
    Check ((($lo.ListColumns('Date/Time').DataBodyRange.Value2 | ForEach-Object { "$_" }) -join '|') -eq $stamp) 'no Date/Time was stamped'
    Check ($cat.ListRows.Count -eq $catBefore) 'no consumable was removed'

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'closed without saving'
if ($script:anyFail) { exit 1 }
