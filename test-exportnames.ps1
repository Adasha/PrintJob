# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Student names in the job-level CSV (docs/ARCHITECTURE.md §10.4, O7).
#
# Export asks "include student names?" (Yes / No / Cancel). A quiet run cannot
# answer, so it includes names unless SetExportNames says otherwise; that is
# how this drives both branches.
#   - names included: Student Name and Student No both in the file, location
#     stamped as exported
#   - names omitted: header block says so, Student Name blank, Student No kept,
#     every other column present, location NOT stamped as exported
#   - import of a no-names file keeps the names already on overwritten rows and
#     appends new rows with a blank name
#   - import of a with-names file still overwrites the name
#
# Exports within one minute share a filename, so each file is copied aside
# straight after it is written. Drives a COPY in %TEMP%, never
# src\PrintJob.xlsm itself - see test-validation.ps1's header. Closes WITHOUT
# saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
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
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Example Print Room')
    $annexe = $wb.Worksheets('Annexe')
    $loM = $main.ListObjects('tblJobs_MAIN')
    $loA = $annexe.ListObjects('tblJobs_ANNEX')

    function Set-Cell($ws, [int]$row, [int]$col, $v) { Invoke-ComRetry { $c = $ws.Cells($row, $col); if ($v -is [string]) { $c.Value2 = [string]$v } else { $c.Value2 = [double]$v } } | Out-Null; Start-Sleep -Milliseconds 300 }
    function Cc($lo, $name) { return $lo.Range.Column + (Col $lo $name) - 1 }
    function Find-Row($ws, $lo, $jobId) {
        $idc = Cc $lo 'Job ID'
        for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
            $r = $lo.ListRows($i).Range.Row
            if ([string]$ws.Cells($r, $idc).Value2 -eq $jobId) { return $r }
        }
        return 0
    }
    # Export Main now, copy the file aside under $name, return the copy.
    function Export-Main($name) {
        [void]$main.Activate()
        [void]$xl.Run('btnExport')
        $csv = Get-ChildItem $workDir -Filter 'PrintCosts-*-MAIN-*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $copy = Join-Path $workDir $name
        Copy-Item $csv.FullName $copy -Force
        Remove-Item $csv.FullName -Force
        return $copy
    }
    function Import-Csv-Into($ws, $path) {
        $rows = $xl.Run('ReadImportRows', $path)
        $xl.Run('ApplyImportConfirmed', $ws, $rows)
        Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
        Start-Sleep -Milliseconds 300
    }

    # ------------------------------------------------ a job with a name
    Write-Host '=== Setup: one Main job with a student name and number ==='
    [void]$main.Activate()
    Invoke-ComRetry { [void]$xl.Run('btnAddPrintJob') } | Out-Null
    $r = Invoke-ComRetry { $loM.ListRows($loM.ListRows.Count).Range.Row }
    $mId = Cc $loM 'Job ID'; $mName = Cc $loM 'Student Name'; $mNo = Cc $loM 'Student No'
    Set-Cell $main $r (Cc $loM 'Paper Stock') 'Canvas 914mm roll'
    Set-Cell $main $r (Cc $loM 'Qty') 2
    Set-Cell $main $r $mName 'Ada Lovelace'
    Set-Cell $main $r $mNo 'S1234'
    $jobId = [string]$main.Cells($r, $mId).Value2
    Check ($jobId.Length -gt 0) "job $jobId created"

    # ---------------------------------------------- default: names in
    Write-Host ''
    Write-Host '=== Quiet default: names included ==='
    $withNames = Export-Main 'with-names.csv'
    $lines = Get-Content $withNames -Encoding UTF8
    Check ($lines[7] -notmatch 'Student names') 'header block carries no omitted marker'
    $recs = $lines[8..($lines.Count - 1)] | ConvertFrom-Csv
    $rec = $recs | Where-Object { $_.'Job ID' -eq $jobId }
    Check ($rec.'Student Name' -eq 'Ada Lovelace') "Student Name in file ($($rec.'Student Name'))"
    Check ($rec.'Student No' -eq 'S1234') 'Student No in file'
    Check ($main.Range('LOC_Export').Text -match '^Exported') "location stamped as exported ($($main.Range('LOC_Export').Text))"

    # ------------------------------------------------- names omitted
    Write-Host ''
    Write-Host '=== Names omitted ==='
    # Change the data so the fingerprint differs from the stamped export.
    Set-Cell $main $r (Cc $loM 'Qty') 3
    [void]$xl.Run('RefreshExportStatus', $main)
    $before = $main.Range('LOC_Export').Text
    Check ($before -match '^CHANGED') "status shows CHANGED after an edit ($before)"
    $xl.Run('SetExportNames', $false)
    $noNames = Export-Main 'no-names.csv'
    $xl.Run('SetExportNames', $true)
    $lines = Get-Content $noNames -Encoding UTF8
    Check ($lines[7] -match 'Student names,Omitted') "header block marks names omitted ($($lines[7]))"
    Check ($lines[8] -match '^Job ID,') 'header row still on line 9'
    $recs = $lines[8..($lines.Count - 1)] | ConvertFrom-Csv
    $rec = $recs | Where-Object { $_.'Job ID' -eq $jobId }
    Check ([string]::IsNullOrEmpty($rec.'Student Name')) 'Student Name blank'
    Check ($rec.'Student No' -eq 'S1234') 'Student No kept'
    foreach ($need in 'S_UnitCost', 'S_ConsRate', 'S_SchemaVer', 'Chargeable Cost', 'Notes') {
        Check ($null -ne $rec.PSObject.Properties[$need]) "column $need still present"
    }
    Check ($main.Range('LOC_Export').Text -eq $before) 'no-names export did not change the export status'

    # ------------------------------------------------------- imports
    Write-Host ''
    Write-Host '=== Importing the no-names file ==='
    Set-Cell $main $r $mName 'Changed Name'
    Import-Csv-Into $main $noNames
    $r = Find-Row $main $loM $jobId
    Check ([string]$main.Cells($r, $mName).Value2 -eq 'Changed Name') 'overwrite kept the name already on the sheet'
    Check ([string]$main.Cells($r, $mNo).Value2 -eq 'S1234') 'Student No still S1234'

    Import-Csv-Into $annexe $noNames
    $aId = Cc $loA 'Job ID'
    $ar = Find-Row $annexe $loA $jobId
    Check ($ar -gt 0) 'job appended to Annexe'
    if ($ar -gt 0) {
        Check ([string]::IsNullOrEmpty([string]$annexe.Cells($ar, (Cc $loA 'Student Name')).Value2)) 'appended row has a blank name'
        Check ([string]$annexe.Cells($ar, (Cc $loA 'Student No')).Value2 -eq 'S1234') 'appended row has the Student No'
    }

    Write-Host ''
    Write-Host '=== Importing the with-names file still overwrites the name ==='
    Import-Csv-Into $main $withNames
    $r = Find-Row $main $loM $jobId
    Check ([string]$main.Cells($r, $mName).Value2 -eq 'Ada Lovelace') 'name restored from the with-names file'

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'closed without saving'

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
if ($script:anyFail) { exit 1 }
