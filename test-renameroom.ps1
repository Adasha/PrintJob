# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Renaming a print room (modRegistry.RenameRoomSheet, 2026-10-22): typing a new
# name into the room's name cell renames its sheet tab to match.
#
#   - The name cell is unlocked on the protected sheet, so it can be typed into.
#   - The tab takes the new name; the registry row is re-keyed to it, keeping
#     the room's code and Job ID high-water mark; the rename is audited.
#   - A name another sheet holds gets a " 2" suffix, as Add print room does.
#   - A blank name leaves the tab alone.
#
# The name is written over COM, which fires Worksheet_Change as typing does.
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$first = 'Rename Me'; $renamed = 'Renamed Room'; $taken = 'Example Print Room'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $ws = $xl.Run('CreatePrintRoom', $first, 'Facilities', 'RENAME')
    Check ($null -ne $ws) "CreatePrintRoom made '$first'"
    $reg = $null
    foreach ($s in $wb.Worksheets) { foreach ($lo in $s.ListObjects) { if ($lo.Name -eq 'tblLocations') { $reg = $lo } } }

    function SheetNames { return @($wb.Worksheets | ForEach-Object { [string]$_.Name }) }
    function RegRow([string]$sheetName) {
        for ($i = 1; $i -le $reg.ListRows.Count; $i++) {
            if ([string]$reg.ListRows($i).Range.Cells(1, (Col $reg 'SheetName')).Text -eq $sheetName) {
                $h = @{}
                foreach ($c in $reg.ListColumns) { $h[[string]$c.Name] = [string]$reg.ListRows($i).Range.Cells(1, $c.Index).Text }
                return $h
            }
        }
        return $null
    }
    function AuditLast {
        $a = $wb.Worksheets('_Audit').ListObjects('tblAudit')
        $r = $a.ListRows($a.ListRows.Count).Range
        return @(for ($c = 1; $c -le $a.ListColumns.Count; $c++) { [string]$r.Cells(1, $c).Text })
    }

    # ---------------------------------------------------------------- rename
    Write-Host '=== Typing a new room name renames the tab ==='
    # Two job IDs handed out, so the room has a Job ID high-water mark to keep.
    for ($k = 0; $k -lt 2; $k++) { $xl.Run('NextJobId', $ws, $ws.ListObjects('tblJobs_RENAME')) | Out-Null }
    $hwm = (RegRow $first)['Job ID HWM']
    Check ($hwm -eq '2') "the room has a Job ID high-water mark of 2 (got '$hwm')"
    $nameCell = $ws.Names.Item('LOC_Name').RefersToRange
    Check ([bool]$ws.ProtectContents) 'the sheet is protected'
    Check (-not [bool]$nameCell.Locked) 'the room name cell is not locked'
    $nameCell.Value2 = $renamed
    $names = SheetNames
    Check (($names -contains $renamed) -and -not ($names -contains $first)) "the tab now reads '$renamed' (sheets: $($names -join ', '))"
    Check ([bool]$wb.Worksheets($renamed).ProtectContents) 'the sheet is still protected'
    $row = RegRow $renamed
    Check ($null -ne $row -and $row['Name'] -eq $renamed) 'the registry row follows the new tab name'
    Check ($null -ne $row -and $row['Code'] -eq 'RENAME') 'the room keeps its code'
    Check ($null -ne $row -and $row['Job ID HWM'] -eq '2') 'the Job ID high-water mark came across'
    Check ($null -eq (RegRow $first)) 'no registry row is left under the old tab name'
    $last = AuditLast
    Check (($last -contains 'Rename print room') -and (@($last | Where-Object { $_ -like "*$first*" }).Count -gt 0)) "the rename is in the audit log ($($last -join ' | '))"

    # ------------------------------------------------------------ name taken
    Write-Host ''
    Write-Host '=== A name another sheet holds gets a suffix ==='
    $nameCell.Value2 = $taken
    $names = SheetNames
    Check (($names -contains "$taken 2") -and -not ($names -contains $renamed)) "the tab reads '$taken 2' (sheets: $($names -join ', '))"
    Check ([string]$wb.Worksheets($taken).Names.Item('LOC_Name').RefersToRange.Text -eq $taken) "'$taken' itself is untouched"

    # ----------------------------------------------------------------- blank
    Write-Host ''
    Write-Host '=== A blank name leaves the tab alone ==='
    $nameCell.Value2 = ''
    Check ((SheetNames) -contains "$taken 2") "the tab still reads '$taken 2'"

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
