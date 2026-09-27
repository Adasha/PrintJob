# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Prerequisite: persisted Job ID high-water mark (modRegistry.NextJobId).
#
# Confirms the bug the snag list names is actually fixed: deleting the
# highest-numbered job must not let the next Add Print Job reissue its ID.
# Also confirms the high-water mark survives a RefreshLocations rebuild,
# since WriteRegistry deletes and re-adds every registry row on every run.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself - see verify.ps1's
# comment for why (AutoSave on a OneDrive-backed handle commits on Close/Quit
# regardless of the Close($false) argument).

$ErrorActionPreference = 'Stop'
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $ws = $wb.Worksheets('Example Print Room')
    $lo = $ws.ListObjects('tblJobs_MAIN')

    function Get-JobIdCol {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            if ($lo.ListColumns($i).Name -eq 'Job ID') { return $i }
        }
        throw 'Job ID column not found'
    }
    $idCol = Get-JobIdCol

    $ws.Activate()
    $xl.Run('btnAddPrintJob')
    $n1 = $lo.ListRows.Count
    $id1 = [string]$lo.ListRows($n1).Range.Cells(1, $idCol).Value2
    Write-Host ("added row {0}: {1}" -f $n1, $id1)
    if ([string]::IsNullOrEmpty($id1)) { throw 'first Add Print Job produced no Job ID' }

    # Delete that row directly via COM - RemoveRow's own Ask() always declines
    # under SetQuiet, by design (an unattended run must never confirm a
    # destructive op on the user's behalf), so the button path cannot be
    # driven here. This is the exact scenario the bug fix targets: the
    # highest-numbered row is gone before the next ID is allocated.
    $ws.Unprotect()
    $lo.ListRows($n1).Delete()
    $ws.Protect()
    Write-Host ("deleted row {0} ({1})" -f $n1, $id1)

    $xl.Run('btnAddPrintJob')
    $n2 = $lo.ListRows.Count
    $id2 = [string]$lo.ListRows($n2).Range.Cells(1, $idCol).Value2
    Write-Host ("added row {0}: {1}" -f $n2, $id2)

    if ($id2 -eq $id1) {
        Write-Host ''
        Write-Host "FAIL: '$id2' reissued the deleted ID '$id1'"
        exit 1
    }
    Write-Host ("OK: new ID '{0}' does not reuse deleted ID '{1}'" -f $id2, $id1)

    # The HWM must survive WriteRegistry's delete-and-rebuild of every
    # registry row, which RefreshLocations performs on every call.
    $xl.Run('btnRefreshLocations')
    $xl.Run('btnAddPrintJob')
    $n3 = $lo.ListRows.Count
    $id3 = [string]$lo.ListRows($n3).Range.Cells(1, $idCol).Value2
    Write-Host ("after RefreshLocations, added row {0}: {1}" -f $n3, $id3)
    if ($id3 -eq $id1 -or $id3 -eq $id2) {
        Write-Host ''
        Write-Host "FAIL: '$id3' collides with an earlier ID after RefreshLocations"
        exit 1
    }
    Write-Host "OK: high-water mark survived RefreshLocations"

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'PASS'
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
