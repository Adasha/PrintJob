# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# An apostrophe in a print room's name (0.12.1): the LOC_ names are rebuilt from
# the sheet name on every refresh, and an unescaped apostrophe made that raise
# 1004. CreatePrintRoom and ChangeRoomName are called directly, so there is no
# InputBox to answer.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$room = "Dean's Room"; $renamed = "Bob's Lab"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    function SheetNames { return @($wb.Worksheets | ForEach-Object { [string]$_.Name }) }
    function Clean([string]$log, [string]$what) { Check (-not $log.Contains('went wrong')) "$what raised no error" }
    function NameAddr([string]$sheet, [string]$name) { return [string]$wb.Worksheets($sheet).Names.Item($name).RefersToRange.Address($false, $false) }

    Write-Host '=== Adding a room with an apostrophe ==='
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('CreatePrintRoom', $room, '', 'DEANS')
    Clean ([string]$xl.Run('QuietLog')) 'Add print room'
    Check ((SheetNames) -contains $room) "the tab is called $room"
    $got = NameAddr $room 'LOC_Columns'
    Check ($got -eq 'AX11') "LOC_Columns is on the room (got '$got')"
    $got = NameAddr $room 'LOC_DefTech'
    Check ($got -eq 'B4') "LOC_DefTech is on the room (got '$got')"
    $got = NameAddr $room 'LOC_Export'
    Check ($got -eq 'AX8') "LOC_Export is on the room (got '$got')"
    $got = [string]$wb.Worksheets($room).Range('B12').Text
    Check ($got.StartsWith('0 printers')) "Printers at this location reads (got '$got')"

    Write-Host ''
    Write-Host '=== Refreshing ==='
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('RefreshLocations')
    $log = [string]$xl.Run('QuietLog')
    Clean $log 'Refresh Locations'
    Check ($log.Contains('2 print rooms registered')) 'both rooms are registered'
    Check ([string]$wb.Worksheets('_Data').Range('A10').Formula -ne '') '_Data has its formula'

    Write-Host ''
    Write-Host '=== Renaming ==='
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('ChangeRoomName', $wb.Worksheets($room), $renamed)
    Clean ([string]$xl.Run('QuietLog')) 'Rename print room'
    Check ((SheetNames) -contains $renamed) "the tab follows to $renamed"
    $got = NameAddr $renamed 'LOC_DefTech'
    Check ($got -eq 'B4') "LOC_DefTech is still on the room (got '$got')"

    Write-Host ''
    Write-Host '=== An apostrophe at either end ==='
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('CreatePrintRoom', "'Quoted'", '', 'QUOTED')
    Clean ([string]$xl.Run('QuietLog')) 'Add print room'
    Check ((SheetNames) -contains 'Quoted') 'the tab drops them'

    Write-Host ''
    Write-Host '=== Removing ==='
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('DeletePrintRoom', $wb.Worksheets($renamed))
    Clean ([string]$xl.Run('QuietLog')) 'Remove print room'
    Check ((SheetNames) -notcontains $renamed) 'the room has gone'

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
