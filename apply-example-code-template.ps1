# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Re-runnable template change: the shipped sample room's code goes from MAIN to
# EXAMPLE. MAIN is what a site would naturally call its own first room, so a
# real room coded MAIN collided with the sample (Add print room turned it into
# MAIN2, and a restored MAIN room found the sample sheet). PrintCosts.xlsx is
# binary, so the change is kept as this script rather than as a diff.
#
# Works on a COPY in %TEMP% (never opens the real file - ARCHITECTURE 13.1) and
# only replaces src\PrintCosts.xlsx when every step has succeeded. Idempotent.
#
#   powershell -ExecutionPolicy Bypass -File apply-example-code-template.ps1
#
# What it changes on the "Example Print Room" sheet:
#   - LOC_Code (AX2) MAIN -> EXAMPLE
#   - table tblJobs_MAIN -> tblJobs_EXAMPLE (the ROWS() count in B13 follows it)
#   - the twelve sample Job IDs UNI-MAIN-000nn -> UNI-EXAMPLE-000nn
#   - _Registry's code cell for the room (rewritten by Refresh Locations anyway)

[CmdletBinding()]
param([string] $Template = '')
$ErrorActionPreference = 'Stop'
if (-not $Template) { $Template = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'src\PrintCosts.xlsx' }
if (-not (Test-Path $Template)) { throw "Not found: $Template" }

$oldCode = 'MAIN'; $newCode = 'EXAMPLE'
$work = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsExampleCode-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item $Template $work
$stampBefore = (Get-Item $Template).LastWriteTimeUtc

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($work)
    try { $wb.AutoSaveOn = $false } catch {}
    $ws = $wb.Worksheets.Item('Example Print Room')

    $prot = @()
    foreach ($s in $wb.Worksheets) { if ($s.ProtectContents) { $s.Unprotect(); $prot += $s.Name } }

    $codeCell = $ws.Names.Item('LOC_Code').RefersToRange
    if ([string]$codeCell.Value2 -eq $oldCode) { $codeCell.Value2 = $newCode; Write-Host "LOC_Code -> $newCode" }

    $lo = $ws.ListObjects.Item(1)
    if ($lo.Name -eq "tblJobs_$oldCode") { $lo.Name = "tblJobs_$newCode"; Write-Host "table -> $($lo.Name)" }

    $ids = $lo.ListColumns.Item('Job ID').DataBodyRange
    $n = 0
    foreach ($c in $ids.Cells) {
        $v = [string]$c.Value2
        if ($v -like "*-$oldCode-*") { $c.Value2 = $v.Replace("-$oldCode-", "-$newCode-"); $n++ }
    }
    Write-Host "$n Job ID(s) rewritten"

    $reg = $wb.Worksheets.Item('_Registry')
    $f = $reg.UsedRange.Find($oldCode, [Type]::Missing, -4163, 1)   # xlValues, xlWhole
    if ($f -and $f.Row -gt 1) { $f.Value2 = $newCode; Write-Host "_Registry $($f.Address($false,$false)) -> $newCode" }

    foreach ($name in $prot) { $wb.Worksheets.Item($name).Protect() }

    # Nothing about the old code may be left on the sheet.
    $left = $ws.UsedRange.Find("-$oldCode-", [Type]::Missing, -4123, 2)
    if ($left) { throw "'-$oldCode-' still present at $($left.Address($false,$false))" }
    if ($ws.UsedRange.Find("tblJobs_$oldCode", [Type]::Missing, -4123, 2)) { throw "tblJobs_$oldCode still referenced" }

    $wb.Save()
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

if ((Get-Item $Template).LastWriteTimeUtc -ne $stampBefore) { throw "$Template changed while the script ran; not overwriting it." }
Copy-Item $work $Template -Force
Remove-Item $work -Force
Write-Host "template updated: $Template"
