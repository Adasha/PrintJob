# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Settings table row order and editable/locked styling.
#
#   - LOC_MINIMAL_COLUMNS sits directly below LOC_REDUCED_COLUMNS, in a fresh
#     build and after a second setup run (idempotent). An existing row is never
#     moved: upgrading means a fresh workbook plus an import.
#   - SET_LOC_MINIMAL_COLUMNS still resolves to the Value cell of that row.
#   - Every Value cell is styled to match its lock state: locked ("Read-only"
#     notes) are grey/italic with no border, editable are white with a border.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

function RowOf($lo, $key) {
    $kc = Col $lo 'Key'
    for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
        if ([string]$lo.DataBodyRange.Cells($i, $kc).Value2 -eq $key) { return $i }
    }
    return 0
}

function CheckOrder($wb, $tag) {
    $lo = $wb.Worksheets('Settings').ListObjects('tblSettings')
    $red = RowOf $lo 'LOC_REDUCED_COLUMNS'
    $min = RowOf $lo 'LOC_MINIMAL_COLUMNS'
    Check ($red -gt 0 -and $min -gt 0) "$tag`: both rows exist (reduced=$red minimal=$min)"
    Check ($min -eq $red + 1) "$tag`: LOC_MINIMAL_COLUMNS is directly below LOC_REDUCED_COLUMNS (rows $red, $min)"
    $n = $wb.Names.Item('SET_LOC_MINIMAL_COLUMNS').RefersToRange
    Check ($n.Row -eq $lo.DataBodyRange.Cells($min, 3).Row) "$tag`: SET_LOC_MINIMAL_COLUMNS points at the moved row"
    $keys = @(); for ($i = 1; $i -le $lo.ListRows.Count; $i++) { $keys += [string]$lo.DataBodyRange.Cells($i, 1).Value2 }
    Check (@($keys | Group-Object | Where-Object Count -gt 1).Count -eq 0) "$tag`: no duplicate keys"
}

function CheckStyle($wb, $tag) {
    $lo = $wb.Worksheets('Settings').ListObjects('tblSettings')
    $nc = Col $lo 'Notes'; $vc = Col $lo 'Value'
    $locked = 0; $open = 0
    for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
        $v = $lo.DataBodyRange.Cells($i, $vc)
        $isRo = ([string]$lo.DataBodyRange.Cells($i, $nc).Value2).Trim().StartsWith('Read-only')
        $key = [string]$lo.DataBodyRange.Cells($i, 1).Value2
        Check ([bool]$v.Locked -eq $isRo) "$tag`: $key lock state follows its Notes"
        if ($isRo) {
            $locked++
            Check ($v.Interior.Color -eq 14606046 -and $v.Font.Italic -eq $true) "$tag`: $key (locked) is grey italic"
            Check ($v.Borders(7).LineStyle -eq -4142) "$tag`: $key (locked) has no border"
        } else {
            $open++
            Check ($v.Interior.Color -eq 16777215 -and $v.Font.Italic -eq $false) "$tag`: $key (editable) is white, upright"
            Check ($v.Borders(7).LineStyle -eq 1) "$tag`: $key (editable) has a border"
        }
    }
    Check ($locked -gt 0 -and $open -gt 0) "$tag`: table has both kinds ($locked locked, $open editable)"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = Invoke-ComRetry -Attempts 5 { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    CheckOrder $wb 'built'
    CheckStyle $wb 'built'

    $xl.Run('InitialiseWorkbook')
    Start-Sleep -Milliseconds 500
    CheckOrder $wb 'second setup'
    CheckStyle $wb 'second setup'
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($script:anyFail) { exit 1 }
