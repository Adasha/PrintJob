# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# A print room code is used once (0.11.4). Two rooms on one code hand out the
# same job IDs, so:
#
#   - Add print room's code prompt refuses a code another room has
#     (modRegistry.RoomUsingCode) and suggests a free one (FreeCode). The
#     prompt itself is an InputBox, so only those two are tested here.
#   - Refresh Locations lets a room already in the registry keep its code
#     when a new room repeats it. A new room is inserted right after the
#     first room, so before 0.11.4 a newcomer repeating a LATER room's code
#     took it, and that room was renumbered.
#
# CreatePrintRoom is called directly, so there is no InputBox to answer.
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$first = 'Photography'; $clash = 'Photo Lab'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
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
    function TableNames([string]$sheetName) { return @($wb.Worksheets($sheetName).ListObjects | ForEach-Object { [string]$_.Name }) }

    # ---------------------------------------------------------------- prompt
    Write-Host '=== The prompt knows which codes are taken ==='
    $got = [string]$xl.Run('RoomUsingCode', 'example')
    Check ($got -eq 'Example Print Room') "'example' is the sample room's code (got '$got')"
    $got = [string]$xl.Run('RoomUsingCode', 'NOPE')
    Check ($got -eq '') "'NOPE' is nobody's code (got '$got')"
    $got = [string]$xl.Run('FreeCode', 'nope')
    Check ($got -eq 'NOPE') "an unused code is suggested as typed (got '$got')"

    $ws = $xl.Run('CreatePrintRoom', $first, 'Media Services', 'PHOTO')
    Check ($null -ne $ws) "CreatePrintRoom made '$first'"
    $ids = @(for ($k = 0; $k -lt 2; $k++) { [string]$xl.Run('NextJobId', $ws, $ws.ListObjects('tblJobs_PHOTO')) })
    Check (($ids -join ',') -eq 'UNI-PHOTO-00001,UNI-PHOTO-00002') "'$first' handed out two job IDs ($($ids -join ', '))"
    $got = [string]$xl.Run('RoomUsingCode', 'photo')
    Check ($got -eq $first) "'photo' is now '$first''s code (got '$got')"
    $got = [string]$xl.Run('FreeCode', 'PHOTO')
    Check ($got -eq 'PHOTO2') "the suggestion after a clash is PHOTO2 (got '$got')"

    # --------------------------------------------------------------- refresh
    Write-Host ''
    Write-Host '=== A new room repeating a code does not take it from the room that has it ==='
    $ws2 = $xl.Run('CreatePrintRoom', $clash, '', 'PHOTO')
    Check ($null -ne $ws2) "CreatePrintRoom made '$clash'"
    $names = SheetNames
    Check (($names -contains $clash) -and ($names -contains $first) -and ([array]::IndexOf($names, $clash) -lt [array]::IndexOf($names, $first))) "'$clash' sits in front of '$first', the order that went wrong (sheets: $($names -join ', '))"
    $got = [string]$xl.Run('LocValue', $ws, 'LOC_Code')
    Check ($got -eq 'PHOTO') "'$first' keeps PHOTO (got '$got')"
    $got = [string]$xl.Run('LocValue', $ws2, 'LOC_Code')
    Check ($got -eq 'PHOTO2') "'$clash' is given PHOTO2 (got '$got')"
    $row = RegRow $first
    Check ($null -ne $row -and $row['Code'] -eq 'PHOTO' -and $row['Job ID HWM'] -eq '2') "the registry still has '$first' on PHOTO with its high-water mark of 2"
    $row = RegRow $clash
    Check ($null -ne $row -and $row['Code'] -eq 'PHOTO2') "the registry has '$clash' on PHOTO2"
    Check (((TableNames $first) -contains 'tblJobs_PHOTO') -and ((TableNames $clash) -contains 'tblJobs_PHOTO2')) "each room's job table is named after its own code"
    $got = [string]$xl.Run('NextJobId', $ws2, $ws2.ListObjects('tblJobs_PHOTO2'))
    Check ($got -eq 'UNI-PHOTO2-00001') "'$clash''s first job ID is UNI-PHOTO2-00001 (got '$got')"
    $got = [string]$xl.Run('NextJobId', $ws, $ws.ListObjects('tblJobs_PHOTO'))
    Check ($got -eq 'UNI-PHOTO-00003') "'$first' carries on from UNI-PHOTO-00003 (got '$got')"

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
