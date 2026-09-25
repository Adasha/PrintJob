# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

$ErrorActionPreference = 'Stop'
$f = Join-Path $PSScriptRoot 'src\PrintCosts.xlsm'
Write-Host '=== version ==='

# Document properties are read out of the OOXML package, not through Excel.
# $wb.BuiltinDocumentProperties('Title') does not resolve as a parameterized
# COM property from PowerShell: it yields a blank rather than an error, which
# reads exactly like a workbook whose properties were never set. Done before
# Excel opens the file, because Excel takes a lock.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($f)
try {
    $core = $zip.Entries | Where-Object { $_.FullName -eq 'docProps/core.xml' }
    if ($core) {
        $sr = New-Object IO.StreamReader($core.Open())
        $xml = [xml]$sr.ReadToEnd()
        $sr.Close()
        $cp = $xml.coreProperties
        Write-Host ("  Title     {0}" -f $cp.title)
        Write-Host ("  Subject   {0}" -f $cp.subject)
        Write-Host ("  Author    {0}" -f $cp.creator)
        Write-Host ("  Comments  {0}" -f $cp.description)
    } else {
        Write-Host '  docProps/core.xml MISSING'
    }
}
finally { $zip.Dispose() }

# ReadOnly, and it matters. Workbook_Open does real work every time the file
# is opened - ProtectAll, HealButtons, the 1904 check - which marks the
# workbook dirty. Opened read-write, a script that only reports on the file
# ends up rewriting it, moving its timestamp and setting OneDrive syncing
# again on every run. Opened read-only those changes simply cannot be saved.
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f, $false, $true)   # UpdateLinks=false, ReadOnly=true

    foreach ($n in @('SET_APP_VER', 'SET_SCHEMA', 'SET_BUILT', 'SET_BUILT_BY')) {
        try { Write-Host ("  {0,-14} {1}" -f $n, $wb.Names($n).RefersToRange.Text) }
        catch { Write-Host ("  {0,-14} MISSING" -f $n) }
    }

    Write-Host ''
    Write-Host '=== job tables ==='
    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            if ($lo.Name -like 'tblJobs_*') {
                Write-Host ("{0,-18} on {1,-18} rows={2}" -f $lo.Name, $ws.Name, $lo.ListRows.Count)
            }
        }
    }

    Write-Host ''
    Write-Host '=== tblLocations ==='
    $lo = $wb.Worksheets('_Registry').ListObjects('tblLocations')
    Write-Host ("rows: {0}" -f $lo.ListRows.Count)
    for ($r = 1; $r -le $lo.ListRows.Count; $r++) {
        $vals = @()
        for ($c = 1; $c -le $lo.ListColumns.Count; $c++) {
            $vals += [string]$lo.ListRows($r).Range.Cells(1, $c).Text
        }
        Write-Host ('  ' + ($vals -join ' | '))
    }

    Write-Host ''
    Write-Host '=== _Data ==='
    $d = $wb.Worksheets('_Data')
    Write-Host ('A10 formula: ' + $d.Range('A10').Formula2)
    try {
        $spill = $d.Range('A10').SpillingToRange
        Write-Host ("spill: {0}  ({1} rows x {2} cols)" -f $spill.Address(), $spill.Rows.Count, $spill.Columns.Count)
    } catch {
        Write-Host ('NO SPILL. A10 shows: ' + $d.Range('A10').Text)
    }
    $hdr = @()
    for ($c = 1; $c -le 21; $c++) { $hdr += [string]$d.Cells(9, $c).Text }
    Write-Host ('headers: ' + (($hdr | Where-Object { $_ -ne '' }) -join ' | '))
    for ($r = 10; $r -le 13; $r++) {
        $vals = @()
        for ($c = 1; $c -le 10; $c++) { $vals += [string]$d.Cells($r, $c).Text }
        Write-Host ('  ' + ($vals -join ' | '))
    }

    Write-Host ''
    Write-Host '=== buttons on Main Print Room ==='
    $ws = $wb.Worksheets('Main Print Room')
    foreach ($b in $ws.Buttons()) { Write-Host ("  {0,-34} -> {1}" -f $b.Name, $b.OnAction) }
}
finally {
    # Runs even when something above throws. It does NOT run if the process is
    # killed outright - including by piping this script into Select-Object
    # -First, which tears down the pipeline early and orphans an Excel holding
    # the workbook open. Don't do that; add a switch to the script instead.
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
