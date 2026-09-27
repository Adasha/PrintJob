# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Reports the structure of a workbook, read-only.
#
#   probe.ps1                        whole workbook: sheets, tables, columns, names
#   probe.ps1 -Sheet Settings        one sheet, cell by cell, for placing new content
#   probe.ps1 -File src\PrintJob.xlsm -Sheet _Data
#
# Opens read-only and closes without saving, so it is safe to run against the
# deliverable.

[CmdletBinding()]
param(
    [string] $File  = 'src\PrintCosts.xlsx',
    [string] $Sheet,
    [int]    $Rows  = 40,
    [int]    $Cols  = 8
)

$ErrorActionPreference = 'Stop'

$path = if ([IO.Path]::IsPathRooted($File)) { $File } else { Join-Path $PSScriptRoot $File }
if (-not (Test-Path $path)) { throw "Not found: $path" }

function Visibility($v) {
    switch ($v) {
        -1      { 'visible' }
         0      { 'hidden' }
         2      { 'very hidden' }
        default { "?$v" }
    }
}

function ColLetter([int] $n) {
    $s = ''
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = [char](65 + $m) + $s
        $n = [int](($n - $m - 1) / 26)
    }
    $s
}

# Read the names the file actually stores, before Excel opens it and adds its
# own. Used below to tell the two apart.
$stored = @()
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($path)
try {
    $e = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' }
    if ($e) {
        $sr = New-Object IO.StreamReader($e.Open())
        $wbXml = $sr.ReadToEnd()
        $sr.Close()
        $stored = [regex]::Matches($wbXml, '<definedName[^>]*name="([^"]+)"') |
                  ForEach-Object { $_.Groups[1].Value }
    }
}
finally { $zip.Dispose() }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($path, $false, $true)   # UpdateLinks=false, ReadOnly=true
    Write-Host ("FILE  {0}" -f $path)
    Write-Host ''

    if ($Sheet) {
        # ---- one sheet, cell by cell -------------------------------------
        $ws = $wb.Worksheets($Sheet)
        Write-Host ("SHEET {0}  ({1})" -f $ws.Name, (Visibility $ws.Visible))
        Write-Host ("used range: {0}" -f $ws.UsedRange.Address())
        foreach ($lo in $ws.ListObjects) {
            Write-Host ("  table {0}: {1}" -f $lo.Name, $lo.Range.Address())
        }
        $shapes = @()
        foreach ($b in $ws.Buttons()) { $shapes += ("{0} -> {1}" -f $b.Name, $b.OnAction) }
        if ($shapes.Count -gt 0) {
            Write-Host '  buttons:'
            $shapes | ForEach-Object { Write-Host ("    {0}" -f $_) }
        }
        Write-Host ''
        for ($r = 1; $r -le $Rows; $r++) {
            $cells = @()
            for ($c = 1; $c -le $Cols; $c++) {
                $t = [string]$ws.Cells($r, $c).Text
                if ($t -ne '') { $cells += ("{0}{1}={2}" -f (ColLetter $c), $r, $t) }
            }
            if ($cells.Count -gt 0) { Write-Host ('  ' + ($cells -join '   ')) }
        }
    }
    else {
        # ---- whole workbook ----------------------------------------------
        foreach ($ws in $wb.Worksheets) {
            Write-Host ("SHEET {0}  ({1})" -f $ws.Name, (Visibility $ws.Visible))
            foreach ($lo in $ws.ListObjects) {
                # Not $cols: PowerShell variable names are case-insensitive, so
                # that is the same variable as the [int] $Cols parameter, and
                # assigning an array to it fails parameter transformation.
                $colNames = @()
                foreach ($c in $lo.ListColumns) { $colNames += $c.Name }
                Write-Host ("   TABLE {0} at {1}: {2}" -f $lo.Name, $lo.Range.Address(), ($colNames -join ' | '))
            }
        }
        # Excel's Names collection includes hidden placeholders it invents at
        # runtime and never saves - _xlfn.IFERROR is the usual one. Reported
        # plainly they look like content of the file, which is how they end up
        # logged as defects. Anything not in xl/workbook.xml is marked.
        Write-Host ''
        Write-Host 'DEFINED NAMES (sheet-scoped ones are shown as Sheet!Name):'
        foreach ($n in $wb.Names) {
            $bare = ($n.Name -split '!')[-1]
            $tag = if ($stored -contains $bare) { '' } else { '   <- runtime only, NOT stored in the file' }
            Write-Host ("   {0} -> {1}{2}" -f $n.Name, $n.RefersTo, $tag)
        }
    }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
