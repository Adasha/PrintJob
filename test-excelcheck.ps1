# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# An Excel missing LET and the other newer functions (0.12.1). Writing the
# workbook's formulas into one stores them broken, so the commands that write
# them refuse first (modUtils.ExcelCanRun). No Excel with the fault is to hand:
# the real check is tried against an unknown function, and the refusals are
# reached through PretendMissingExcelFunctions.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$room = 'Example Print Room'
$tbl = 'tblJobs_EXAMPLE'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    function SheetNames { return @($wb.Worksheets | ForEach-Object { [string]$_.Name }) }

    Write-Host '=== The check ==='
    $got = [string]$xl.Run('MissingExcelFunctions')
    Check ($got -eq '') "this Excel has every function (missing: '$got')"
    $got = [string]$xl.Run('ExcelKnowsFormula', '=LET(x,1,x)')
    Check ($got -eq 'True') "a known function is known (got '$got')"
    $got = [string]$xl.Run('ExcelKnowsFormula', '=NOSUCHFUNCTION(1)')
    Check ($got -eq 'False') "an unknown function is not (got '$got')"
    [void]$xl.Run('WarnIfExcelFunctionsMissing')
    $got = [string]$xl.Run('QuietLog')
    Check ($got -eq '') "no warning on opening (got '$got')"

    $extra = $xl.Run('CreatePrintRoom', 'Second Room', '', 'SECOND')
    Check ($null -ne $extra) 'a second room to work on'
    $dataBefore = [string]$wb.Worksheets('_Data').Range('A10').Formula
    $sheetsBefore = (SheetNames) -join ','
    $xl.Run('SetQuiet', $true)

    Write-Host ''
    Write-Host '=== Pretending LET and VSTACK are missing ==='
    [void]$xl.Run('PretendMissingExcelFunctions', 'LET, VSTACK')
    $got = [string]$xl.Run('MissingExcelFunctions')
    Check ($got -eq 'LET, VSTACK') "the check reports them (got '$got')"
    [void]$xl.Run('WarnIfExcelFunctionsMissing')
    $got = [string]$xl.Run('QuietLog')
    Check ($got.Contains('LET, VSTACK') -and $got.Contains('will not work properly')) 'the workbook warns on opening'

    function Refused([string]$what, [scriptblock]$act) {
        $xl.Run('SetQuiet', $true)
        [void](& $act)
        $log = [string]$xl.Run('QuietLog')
        Check ($log.Contains("$what did not run") -and $log.Contains('LET, VSTACK')) "$what refuses, naming them"
    }
    $second = $wb.Worksheets('Second Room')
    Refused 'Refresh Locations' { $xl.Run('RefreshLocations') }
    Check ([string]$wb.Worksheets('_Data').Range('A10').Formula -eq $dataBefore) "_Data's formula is untouched"
    Check (@($wb.Worksheets($room).ListObjects | ForEach-Object { $_.Name }) -contains $tbl) 'job tables were not parked'
    Refused 'Add print room' { $xl.Run('CreatePrintRoom', 'Refused Room', '', 'NOPE') }
    Refused 'Rename print room' { $xl.Run('ChangeRoomName', $second, 'Renamed Room') }
    Refused 'Remove print room' { $xl.Run('DeletePrintRoom', $second) }
    Refused 'Restore workbook' { $xl.Run('RestoreFromFileConfirmed', 'C:\nonexistent.xlsm') }
    $got = (SheetNames) -join ','
    Check ($got -eq $sheetsBefore) "no sheet came or went, or changed name ($got)"

    Write-Host ''
    Write-Host '=== Formulas are not rewritten while it is missing ==='
    $ws = $wb.Worksheets($room)
    $ws.Unprotect()
    $xl.EnableEvents = $false
    [void]$ws.Range('B12').ClearContents()
    [void]$xl.Run('EnsurePrintersDisplay', $ws)
    Check ([string]$ws.Range('B12').Formula -eq '') 'the room header is left alone'

    Write-Host ''
    Write-Host '=== With the functions back ==='
    [void]$xl.Run('PretendMissingExcelFunctions', '')
    $got = [string]$xl.Run('MissingExcelFunctions')
    Check ($got -eq '') "the check reports nothing (got '$got')"
    [void]$xl.Run('EnsurePrintersDisplay', $ws)
    $got = [string]$ws.Range('B12').Text
    Check ($got.StartsWith('3 printers')) "the room header is written (got '$got')"
    $xl.EnableEvents = $true
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('RefreshLocations')
    $got = [string]$xl.Run('QuietLog')
    Check ($got.Contains('2 print rooms registered')) 'Refresh Locations runs'

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
