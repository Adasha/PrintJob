# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Settings-sheet button arrangement (ARCHITECTURE.md §16.3 punch list). The ten
# action buttons sit in a band ABOVE the tables, in three labelled groups, packed
# close together:
#
#   Print rooms - Add print room, Remove print room, Refresh Locations
#   Data        - Export All, Import, Backup, Restore
#   Workbook    - Check workbook, Re-stamp prices, About
#
# Checks: every button exists once and sits above the tables' subtitle row; each
# group is on its own row, in order, with the right label; buttons in a row are
# close together (small gap, no overlap) and inside the width of columns A:D;
# the +/- buttons sit just above the lookup tables; the stale "Commands" note is
# gone; and running setup again changes nothing (idempotent - rows are inserted
# once, not once per run).
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$groups = [ordered]@{
    'Print rooms' = 'btnAddPrintRoom', 'btnRemovePrintRoom', 'btnRefreshLocations'
    'Data'        = 'btnExportAll', 'btnImportGlobal', 'btnBackupWorkbook', 'btnRestoreWorkbook'
    'Workbook'    = 'btnCheckWorkbook', 'btnReStamp', 'btnAbout'
}
$small = 'btnAddRowPaperTypes', 'btnRemoveRowPaperTypes', 'btnAddRowSize', 'btnRemoveRowSize', 'btnAddRowConsumable', 'btnRemoveRowConsumable'

function SettingsButtons($ws) {
    $list = @()
    $n = $ws.Buttons().Count
    for ($i = 1; $i -le $n; $i++) {
        $b = $ws.Buttons($i)
        $list += [pscustomobject]@{ Name = [string]$b.Name; Left = [double]$b.Left; Top = [double]$b.Top; W = [double]$b.Width; H = [double]$b.Height }
    }
    return $list
}
function Find($buttons, $macro) { @($buttons | Where-Object { $_.Name -like "pcb_${macro}_*" }) }

function CheckLayout($wb, $tag) {
    $ws = $wb.Worksheets('Settings')
    $buttons = SettingsButtons $ws
    $hdr = $ws.ListObjects('tblSettings').Range.Row
    Check ($hdr -eq 8) "$tag`: tblSettings header is on row 8 (was $hdr)"
    foreach ($t in 'tblPaperTypes', 'tblStandardSizes', 'tblConsumables') {
        Check ($ws.ListObjects($t).Range.Row -eq 8) "$tag`: $t header is on row 8 too (tables stay aligned)"
    }
    Check ([string]$ws.Range('P2').Value2 -ne 'Commands') "$tag`: the stale 'Commands' heading is gone"
    Check (-not ([string]$ws.Range('P6').Value2 -like 'Buttons appear*')) "$tag`: the stale placeholder note is gone"

    $limit = $ws.Cells(1, 5).Left          # right edge of column D
    $subtitleTop = $ws.Rows(7).Top
    $row = 3
    foreach ($g in $groups.Keys) {
        Check ([string]$ws.Cells($row, 1).Value2 -eq $g) "$tag`: A$row is the '$g' label"
        $rowTop = $ws.Rows($row).Top; $rowBottom = $rowTop + $ws.Rows($row).Height
        $placed = @()
        foreach ($m in $groups[$g]) {
            $hit = @(Find $buttons $m)
            Check ($hit.Count -eq 1) "$tag`: exactly one $m"
            if ($hit.Count -eq 1) {
                $b = $hit[0]
                Check ($b.Top -ge $rowTop -and ($b.Top + $b.H) -le $rowBottom + 0.5) "$tag`: $m sits inside its group row ($row)"
                Check (($b.Top + $b.H) -le $subtitleTop) "$tag`: $m is above the table titles"
                $placed += $b
            }
        }
        $placed = @($placed | Sort-Object Left)
        for ($i = 1; $i -lt $placed.Count; $i++) {
            $gap = $placed[$i].Left - ($placed[$i - 1].Left + $placed[$i - 1].W)
            Check ($gap -ge 0 -and $gap -le 10) ("$tag`: $g buttons {0}/{1} are clustered (gap {2:N1}pt)" -f $i, ($i + 1), $gap)
        }
        if ($placed.Count -gt 0) {
            $last = $placed[-1]
            Check (($last.Left + $last.W) -le $limit) "$tag`: the $g row ends inside columns A:D"
        }
        $row++
    }
    foreach ($m in $small) {
        $hit = @(Find $buttons $m)
        Check ($hit.Count -eq 1) "$tag`: exactly one $m"
        if ($hit.Count -eq 1) {
            Check ($hit[0].Top -ge $ws.Rows(6).Top -and ($hit[0].Top + $hit[0].H) -le $subtitleTop + 2) "$tag`: $m sits on row 6, just above its table's title"
        }
    }
    $all = @($groups.Values | ForEach-Object { $_ }) + $small
    $extra = @($buttons | Where-Object { $_.Name -like 'pcb_*' }).Count - $all.Count
    Check ($extra -eq 0) "$tag`: no stray buttons on Settings (extra: $extra)"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = Invoke-ComRetry -Attempts 5 { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)

    CheckLayout $wb 'built'

    # Idempotence: a second setup run must not insert another band.
    $xl.Run('InitialiseWorkbook')
    Start-Sleep -Milliseconds 500
    CheckLayout $wb 'after second setup'
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($script:anyFail) { exit 1 }
