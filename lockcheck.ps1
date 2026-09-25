# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Continue'
$f = Join-Path $PSScriptRoot 'src\PrintCosts.xlsm'

Write-Host 'EXCEL processes:'
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) { $p | Select-Object Id, StartTime, MainWindowTitle | Format-Table -AutoSize | Out-String | Write-Host }
else    { Write-Host '  none' }

Write-Host 'Owner lock files in src:'
$locks = Get-ChildItem (Split-Path $f) -Force | Where-Object { $_.Name -like '~*' }
if ($locks) { $locks | ForEach-Object { Write-Host ('  ' + $_.Name) } } else { Write-Host '  none' }

Write-Host 'Exclusive write test:'
try {
    $s = [IO.File]::Open($f, 'Open', 'ReadWrite', 'None')
    $s.Close()
    Write-Host '  NOT LOCKED - opened for exclusive write'
} catch {
    Write-Host ('  LOCKED: ' + $_.Exception.Message)
}

Write-Host 'File:'
Get-Item $f | Select-Object Length, LastWriteTime, Attributes | Format-List | Out-String | Write-Host
