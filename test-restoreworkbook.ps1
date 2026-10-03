# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Restore workbook from an OLDER COPY OF THE WORKBOOK instead of a CSV set:
# the upgrade path is a fresh workbook plus a restore.
#
# Copy A is populated (Annexe room with a department, default technician,
# permitted printers, Centimetres roll unit) and saved as the "old" workbook
# under the SAME file name as the fresh copy B - which Excel cannot open
# alongside it, so this also proves the temporary-copy route. A probe line is
# injected into A's Workbook_Open before it is saved; it must NOT run when B
# reads the old file (macros off, events off), and DOES run when the same file
# is opened normally (the positive control, so the check cannot pass
# vacuously). B must end up with the Annexe room created with A's settings and
# the same job records, the old file must be untouched, and no temporary copy
# may be left behind. A workbook that is not a Print Cost workbook is refused.
# Drives COPIES in %TEMP%, never src\PrintJob.xlsm.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$probe = 'PROBE-OPEN-RAN'

function New-WorkDir {
    $d = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $d | Out-Null
    return $d
}

function Sum-Qty($wb, $sheet, $table) {
    $lo = $wb.Worksheets($sheet).ListObjects($table)
    $c = Col $lo 'Qty'
    $sum = 0.0
    for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
        $v = $lo.ListRows($i).Range.Cells(1, $c).Value2
        if ($null -ne $v -and "$v" -ne '') { $sum += [double]$v }
    }
    return [math]::Round($sum, 6)
}

$workDirA = New-WorkDir
$fA = Join-Path $workDirA 'PrintJob.xlsm'
Copy-Item $deliverable $fA
$workDirOld = New-WorkDir
$fOld = Join-Path $workDirOld 'PrintJob.xlsm'     # same name as the fresh copy on purpose
$workDirB = New-WorkDir
$fB = Join-Path $workDirB 'PrintJob.xlsm'
Copy-Item $deliverable $fB

$anyFail = $false
$tempBefore = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'PrintCosts-restore-*' -ErrorAction SilentlyContinue).Count

Write-Host '=== Build the "old" workbook from copy A ==='
$xlA = New-Object -ComObject Excel.Application
$xlA.Visible = $false
$xlA.DisplayAlerts = $false
$wbA = $null
$expect = @{}
try {
    $wbA = $xlA.Workbooks.Open($fA)
    Add-AnnexeFixture $xlA $wbA | Out-Null

    $annexeWs = $wbA.Worksheets('Annexe')
    $xlA.Run('UnlockSheet', $annexeWs)
    $annexeWs.Names.Item('LOC_Dept').RefersToRange.Value2 = 'Fine Art'
    $annexeWs.Names.Item('LOC_DefTech').RefersToRange.Value2 = 'Test Tech'
    $loP = $wbA.Worksheets('Printers').ListObjects('tblPrinters')
    $mc = Col $loP 'Model'
    $models = @(); for ($i = 1; $i -le [math]::Min(2, $loP.ListRows.Count); $i++) { $models += [string]$loP.ListRows($i).Range.Cells(1, $mc).Value2 }
    $expect.Printers = ($models -join ';')
    $annexeWs.Names.Item('LOC_Printers').RefersToRange.Value2 = $expect.Printers
    $xlA.Run('RelockSheet', $annexeWs)

    # The real user path: events on, so the change handler rescales the roll rows.
    $xlA.Run('UnlockSheet', $annexeWs)
    $annexeWs.Names.Item('LOC_RollUnit').RefersToRange.Value2 = 'Centimetres'
    $xlA.Run('RelockSheet', $annexeWs)
    Start-Sleep -Milliseconds 500

    $expect.Dept = 'Fine Art'; $expect.Tech = 'Test Tech'; $expect.Unit = 'Centimetres'
    $expect.Jobs = $wbA.Worksheets('Annexe').ListObjects('tblJobs_ANNEX').ListRows.Count
    $expect.Qty = Sum-Qty $wbA 'Annexe' 'tblJobs_ANNEX'
    $expect.Printers2 = $wbA.Worksheets('Printers').ListObjects('tblPrinters').ListRows.Count
    Write-Host "  Annexe: $($expect.Jobs) job row(s), Qty total $($expect.Qty) (cm), printers '$($expect.Printers)'"

    # Probe: Workbook_Open of the saved copy sets the status bar.
    $cm = $wbA.VBProject.VBComponents('ThisWorkbook').CodeModule
    $line = $cm.ProcBodyLine('Workbook_Open', 0)
    $cm.InsertLines($line + 1, "    Application.StatusBar = `"$probe`"")

    $wbA.SaveCopyAs($fOld)
}
finally {
    if ($wbA) { try { $wbA.Close($false) } catch {} }
    $xlA.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlA)
}
Check ($expect.Qty -gt 0) "the old workbook has roll Qty to compare (total $($expect.Qty))"
$hashBefore = (Get-FileHash $fOld).Hash

Write-Host ''
Write-Host '=== Positive control: opened normally, the old workbook runs its Workbook_Open ==='
$xlC = New-Object -ComObject Excel.Application
$xlC.Visible = $false
$xlC.DisplayAlerts = $false
$wbC = $null
try {
    $wbC = $xlC.Workbooks.Open($fOld, 0, $true)
    Start-Sleep -Milliseconds 500
    Check ([string]$xlC.StatusBar -eq $probe) "probe ran on a normal open (status bar '$($xlC.StatusBar)')"
}
finally {
    if ($wbC) { try { $wbC.Close($false) } catch {} }
    $xlC.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlC)
}

Write-Host ''
Write-Host '=== Restore fresh copy B from the old workbook ==='
$xlB = New-Object -ComObject Excel.Application
$xlB.Visible = $false
$xlB.DisplayAlerts = $false
$wbB = $null
try {
    $wbB = $xlB.Workbooks.Open($fB)
    $xlB.Run('SetQuiet', $true)
    $xlB.StatusBar = $false

    $names = @(); foreach ($ws in $wbB.Worksheets) { $names += $ws.Name }
    Check ($names -notcontains 'Annexe') 'fresh copy has no Annexe before the restore'

    $xlB.Run('RestoreFromFileConfirmed', $fOld)
    Write-Host $xlB.Run('QuietLog')
    Start-Sleep -Milliseconds 300

    Check ([string]$xlB.StatusBar -ne $probe) "old workbook's Workbook_Open did not run while it was read (status bar '$($xlB.StatusBar)')"

    $names = @(); foreach ($ws in $wbB.Worksheets) { $names += $ws.Name }
    Check ($names -contains 'Annexe') 'Annexe room was created from the old workbook'
    $a = $wbB.Worksheets('Annexe')
    Check ([string]$xlB.Run('LocValue', $a, 'LOC_Code') -eq 'ANNEX') 'created room has code ANNEX'
    Check ([string]$xlB.Run('LocValue', $a, 'LOC_Dept') -eq $expect.Dept) "department carried over ('$($xlB.Run('LocValue', $a, 'LOC_Dept'))')"
    Check ([string]$xlB.Run('LocValue', $a, 'LOC_DefTech') -eq $expect.Tech) "default technician carried over ('$($xlB.Run('LocValue', $a, 'LOC_DefTech'))')"
    Check ([string]$xlB.Run('LocValue', $a, 'LOC_Printers') -eq $expect.Printers) "permitted printers carried over ('$($xlB.Run('LocValue', $a, 'LOC_Printers'))')"
    Check ([string]$xlB.Run('LocValue', $a, 'LOC_RollUnit') -eq $expect.Unit) "roll length unit carried over ('$($xlB.Run('LocValue', $a, 'LOC_RollUnit'))')"
    $lo = $a.ListObjects('tblJobs_ANNEX')
    Check ($lo.ListRows.Count -eq $expect.Jobs) "Annexe holds $($expect.Jobs) job row(s) (got $($lo.ListRows.Count))"
    $q = Sum-Qty $wbB 'Annexe' 'tblJobs_ANNEX'
    Check ($q -eq $expect.Qty) "Annexe Qty total matches the old workbook, so no unit conversion crept in ($q vs $($expect.Qty))"
    $pc = $wbB.Worksheets('Printers').ListObjects('tblPrinters').ListRows.Count
    Check ($pc -eq $expect.Printers2) "tblPrinters has the old workbook's $($expect.Printers2) rows (got $pc)"

    Write-Host ''
    Write-Host '=== Restoring again creates no second room and no duplicates ==='
    $xlB.Run('RestoreFromFileConfirmed', $fOld)
    Start-Sleep -Milliseconds 300
    $count = 0; foreach ($ws in $wbB.Worksheets) { if ($ws.Name -like 'Annexe*') { $count++ } }
    Check ($count -eq 1) "still one Annexe sheet (got $count)"
    $lo = $wbB.Worksheets('Annexe').ListObjects('tblJobs_ANNEX')
    Check ($lo.ListRows.Count -eq $expect.Jobs) "Annexe still $($expect.Jobs) row(s) (got $($lo.ListRows.Count))"

    Write-Host ''
    Write-Host '=== A workbook that is not a Print Cost workbook is refused ==='
    $fOther = Join-Path $workDirB 'Other.xlsx'
    $wbO = $xlB.Workbooks.Add()
    $wbO.SaveAs($fOther, 51)
    $wbO.Close($false)
    $xlB.Run('SetQuiet', $true)
    $rowsBefore = $wbB.Worksheets('Annexe').ListObjects('tblJobs_ANNEX').ListRows.Count
    $xlB.Run('RestoreFromFileConfirmed', $fOther)
    $log = [string]$xlB.Run('QuietLog')
    Check ($log -like '*does not look like a Print Cost Management workbook*') 'refused with an explanation'
    Check ($wbB.Worksheets('Annexe').ListObjects('tblJobs_ANNEX').ListRows.Count -eq $rowsBefore) 'nothing changed'

    $xlB.Run('SetQuiet', $false)
}
finally {
    if ($wbB) { try { $wbB.Close($false) } catch {} }
    $xlB.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xlB)
}

Check ((Get-FileHash $fOld).Hash -eq $hashBefore) 'the old workbook file was not modified'
$tempAfter = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'PrintCosts-restore-*' -ErrorAction SilentlyContinue).Count
Check ($tempAfter -eq $tempBefore) "no temporary copy left behind ($tempAfter vs $tempBefore before)"

Write-Host ''
if ($anyFail) { Write-Host 'FAIL' } else { Write-Host 'PASS' }
foreach ($d in $workDirA, $workDirOld, $workDirB) { Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue }
if ($anyFail) { exit 1 }
