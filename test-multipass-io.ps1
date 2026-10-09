# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Prerequisite: multi-pass costing (0.11.0). Export, Import, Re-stamp and Backup
# of colour passes (docs/multipass-costing-design.md decisions 18, 20, 21, 25, 30,
# 31, 33, 34):
#
#   1. Export writes passes as extra rows of the same jobs file, in the new
#      column order (Row Type, Parent, Pass ... Passes, Colour ... Set-up Cost).
#   2. Delete a job, import the file: the job AND its passes are back, in order,
#      at the same costs, at the right outline levels.
#   3. Importing over an existing job replaces its pass group as one block (a
#      file with fewer passes leaves fewer; never doubled); a pass whose job is
#      not in the file is skipped.
#   4. A pre-0.11.0 file (no Row Type) imports its jobs and leaves passes alone.
#   5. An unknown colour imports, keeps its stamped cost, and shows the notice.
#   6. Re-stamp: set-up cost and ink rate follow the Printers / Consumables sheets;
#      a pass whose colour cannot be found is left untouched.
#   7. A pass can be added to a job that is NOT the last row (calculated columns
#      fill into a mid-table insert).
#   8. The colours table is in the backup (one more catalogue CSV).
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$script:anyFail = $false

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $site = [string]$xl.Run('SettingText', 'SITE_ID', 'SITE')
    $main = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    $lo = $main.ListObjects('tblJobs_MAIN')
    function J([int]$rowNo, [string]$h) { return (Invoke-ComRetry -RetryOnNull { $lo.ListRows($rowNo).Range.Cells(1, (Col $lo $h)) }) }
    function W([int]$rowNo, [string]$h, $v) { [void]$xl.Run('UnlockSheet', $main); if ($v -is [int] -or $v -is [double]) { $v = ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }; $cell = J $rowNo $h; Invoke-ComRetry { $cell.Value2 = $v } | Out-Null }
    function Pick([int]$rowNo) { [void]$main.Activate(); [void]((J $rowNo 'Date/Time').Select()) }
    function Level([int]$rowNo) { return [int]$lo.ListRows($rowNo).Range.EntireRow.OutlineLevel }
    function Near([double]$a, [double]$b) { return ([Math]::Abs($a - $b) -lt 0.0001) }
    function Latest-Csv { return (Get-ChildItem $workDir -Filter 'PrintCosts-*-MAIN-*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1) }

    # ---- setup: colours, a RISO printer, one job with three passes ----------------
    $cws = $wb.Worksheets('Consumables'); $clo = $cws.ListObjects('tblColours')
    [void]$cws.Activate()
    foreach ($row in @(@('Black', 0.30), @('Blue', 0.40), @('Fluoro Orange', 0.55))) {
        [void]$xl.Run('AddCatalogRow', 'tblColours'); $n = $clo.ListRows.Count
        [void]$xl.Run('UnlockSheet', $cws)
        $r = $clo.ListRows($n).Range
        $r.Cells(1, (Col $clo 'Colour')).Value2 = $row[0]
        $r.Cells(1, (Col $clo 'Consumable type')).Value2 = 'Risograph'
        $r.Cells(1, (Col $clo 'Cost per m2')).Value2 = [double]$row[1]
    }
    $pws = $wb.Worksheets('Printers'); $plo = $pws.ListObjects('tblPrinters'); [void]$pws.Activate()
    [void]$xl.Run('AddCatalogRow', 'tblPrinters'); $pn = $plo.ListRows.Count; $pr = $plo.ListRows($pn).Range
    [void]$xl.Run('UnlockSheet', $pws)
    $pr.Cells(1, (Col $plo 'Model')).Value2 = 'Test RISO'
    $pr.Cells(1, (Col $plo 'Consumable type')).Value2 = 'Risograph'
    $pr.Cells(1, (Col $plo 'Colour mode')).Value2 = 'multi-pass'
    $pr.Cells(1, (Col $plo 'Template cost')).Value2 = [double]3
    $pr.Cells(1, (Col $plo 'Max sheet size')).Value2 = 'A1'
    $pr.Cells(1, (Col $plo 'Active')).Value2 = 'Yes'
    [void]$xl.Run('UnlockSheet', $main)
    $loc = $main.Names.Item('LOC_Printers').RefersToRange; $loc.Value2 = ([string]$loc.Value2 + ';Test RISO')
    $sample = $lo.ListRows.Count

    function New-Riso-Job([string]$student, [string[]]$colours) {
        [void]$main.Activate(); [void]$xl.Run('AddPrintJob', $main)
        $j = $lo.ListRows.Count
        W $j 'Student Name' $student
        W $j 'Technician' ([string](J 1 'Technician').Value2)
        W $j 'Printer' 'Test RISO'
        W $j 'Paper Stock' 'Gloss 200gsm SRA3 sheet'
        W $j 'Qty' 100
        foreach ($c in $colours) {
            Pick $j; [void]$xl.Run('AddPass', $main)
            W $lo.ListRows.Count 'Colour' $c
        }
        return $j
    }
    $job = New-Riso-Job 'Round Trip' @('Black', 'Blue', 'Fluoro Orange')
    $jobId = [string](J $job 'Job ID').Value2
    $gross0 = [double](J $job 'Gross Cost').Value2
    $cons0 = [double](J $job 'Consumable Cost').Value2
    $setup0 = [double](J $job 'Set-up Cost').Value2
    Check ((Near $cons0 18.0) -and (Near $setup0 9.0) -and (Near $gross0 58.0)) "setup job costs 18 + 9 + 31 = 58 (got $cons0 / $setup0 / $gross0)"

    Write-Host '=== A pass can be added to a job that is not the last row ==='
    $job2 = New-Riso-Job 'Second' @('Black')
    Pick $job
    [void]$xl.Run('AddPass', $main)
    $newPass = $job + 4
    Check ([string](J $newPass 'Row Type').Value2 -eq 'Pass' -and [string](J $newPass 'Parent').Value2 -eq $jobId) 'the 4th pass lands directly under the first job'
    Check ((Level $newPass) -eq 2) '...at level 2'
    Check ([string](J ($job + 5) 'Row Type').Value2 -eq 'Job') '...and the next job moved down intact'
    Check ([string](J $newPass 'Consumable Cost').Value2 -eq '0' -or [string](J $newPass 'Consumable Cost').Value2 -eq '') 'its calculated columns filled (no colour yet: ink 0)'
    Check ([string](J $newPass 'Status').Value2 -like '*Colour required*') '...Status asks for its colour (the Status formula filled in)'
    # take it back out so the rest of the test works on the job as set up
    $detail = 'test'
    [void]$xl.Run('DeletePass', $main, $lo, $job, $newPass, $detail)
    Check (([int](J $job 'Passes').Value2) -eq 3) 'removed again: Passes is 3'
    [void]$xl.Run('DeleteJobGroupAt', $main, $lo, ($job + 4), 'test')     # the second job and its pass
    Check ($lo.ListRows.Count -eq $job + 3) 'second job removed; back to one group'

    Write-Host ''
    Write-Host '=== Export ==='
    [void]$main.Activate()
    [void]$xl.Run('btnExport')
    $csv = Latest-Csv
    Check ($null -ne $csv) 'a CSV was produced'
    $lines = Get-Content $csv.FullName -Encoding UTF8
    $hdr = ($lines[8] -split ',') | ForEach-Object { $_.Trim('"') }
    Write-Host ("  header: " + ($hdr -join ' | '))
    $iJob = $hdr.IndexOf('Job ID')
    Check ($hdr[0] -eq 'Job ID' -and $hdr[1] -eq 'Row Type' -and $hdr[2] -eq 'Parent' -and $hdr[3] -eq 'Pass') 'Row Type, Parent and Pass follow Job ID'
    Check (($hdr.IndexOf('Passes') -eq $hdr.IndexOf('Printer') + 1) -and ($hdr.IndexOf('Set-up Cost') -eq $hdr.IndexOf('Consumable Cost') + 1)) 'Passes follows Printer; Set-up Cost follows Consumable Cost'
    Check (($hdr -contains 'S_SetupCost') -and ($hdr -contains 'S_ColourID') -and ($hdr -contains 'Colour')) 'Colour, S_SetupCost and S_ColourID are in the file'
    $data = $lines[8..($lines.Count - 1)] | ConvertFrom-Csv
    $passRows = @($data | Where-Object { $_.'Row Type' -eq 'Pass' })
    Check ($passRows.Count -eq 3) "the file holds the job's 3 pass rows (got $($passRows.Count))"
    Check ((@($passRows | Where-Object { $_.Parent -eq $jobId }).Count) -eq 3) '...each naming its job in Parent'
    Check ((@($passRows | Where-Object { $_.'Job ID' -ne '' }).Count) -eq 0) '...with no Job ID of their own'
    Check ((@($passRows | Where-Object { $_.'Student Name' -ne '' -or $_.Qty -ne '' }).Count) -eq 0) '...and none of the job fields filled in'
    Check ((@($data | Where-Object { $_.'Row Type' -eq 'Job' }).Count) -eq $sample + 1) 'every other row is a Job row'

    Write-Host ''
    Write-Host '=== Delete the job, import the file: job and passes come back ==='
    [void]$xl.Run('DeleteJobGroupAt', $main, $lo, $job, 'test')
    Check ($lo.ListRows.Count -eq $sample) 'job and its passes deleted'
    $rows = $xl.Run('ReadImportRows', $csv.FullName)
    Check ([int]$rows.Count() -eq $sample + 1 + 3) "ReadImportRows returns every row, passes included ($([int]$rows.Count()))"
    [void]$xl.Run('ApplyImportConfirmed', $main, $rows)
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    $back = $lo.ListRows.Count
    Check ($back -eq $sample + 4) "four rows came back (now $back)"
    $jr = 0; for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if ([string](J $i 'Job ID').Value2 -eq $jobId) { $jr = $i } }
    Check ($jr -gt 0) 'the job is back'
    Check (([string](J ($jr + 1) 'Row Type').Value2 -eq 'Pass') -and ([int](J ($jr + 1) 'Pass').Value2 -eq 1) -and ([string](J ($jr + 1) 'Colour').Value2 -eq 'Black')) 'pass 1 is Black'
    Check (([int](J ($jr + 2) 'Pass').Value2 -eq 2) -and ([string](J ($jr + 2) 'Colour').Value2 -eq 'Blue') -and ([string](J ($jr + 3) 'Colour').Value2 -eq 'Fluoro Orange')) 'passes 2 and 3 are Blue and Fluoro Orange'
    Check ((Level $jr) -eq 1 -and (Level ($jr + 1)) -eq 2 -and (Level ($jr + 3)) -eq 2) 'outline levels are right'
    Check ((Near ([double](J $jr 'Consumable Cost').Value2) $cons0) -and (Near ([double](J $jr 'Set-up Cost').Value2) $setup0) -and (Near ([double](J $jr 'Gross Cost').Value2) $gross0)) 'costs match the original exactly (from the stamped rates, not recosted)'
    Check ([int](J $jr 'Passes').Value2 -eq 3) 'Passes is 3'

    Write-Host ''
    Write-Host '=== Importing over an existing job replaces its passes as one block ==='
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $csv.FullName))
    Check ($lo.ListRows.Count -eq $sample + 4) 'the same file again: still four rows, passes not doubled'
    # A file with only two of the three passes: the job ends with two.
    $dropLine = $null; $kept = @()
    foreach ($ln in $lines[0..8]) { $kept += $ln }
    $seenPass = 0
    foreach ($ln in $lines[9..($lines.Count - 1)]) {
        if ($ln -match '^,Pass,') { $seenPass++; if ($seenPass -eq 3) { continue } }
        $kept += $ln
    }
    $two = Join-Path $workDir 'two-passes.csv'
    Set-Content $two $kept -Encoding UTF8
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $two))
    Check ($lo.ListRows.Count -eq $sample + 3) 'a file with two passes leaves the job with two'
    Check (([int](J $jr 'Passes').Value2) -eq 2) '...Passes 2'
    # An orphan pass (its job is not in the file) is skipped.
    $orph = @(); foreach ($ln in $lines[0..8]) { $orph += $ln }
    $orph += ',Pass,UNI-MAIN-99999,1,,,,,,,Black,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,'
    $orphFile = Join-Path $workDir 'orphan.csv'
    Set-Content $orphFile $orph -Encoding UTF8
    $before = $lo.ListRows.Count
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $orphFile))
    Check ($lo.ListRows.Count -eq $before) 'a pass whose job is not in the file is skipped'
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $csv.FullName))   # restore the 3-pass state

    Write-Host ''
    Write-Host '=== A file from before 0.11.0 (no Row Type) leaves passes alone ==='
    $oldCols = @('Job ID', 'Date/Time', 'Student Name', 'Student No', 'Technician', 'Printer', 'Paper Stock', 'Unit', 'Qty', 'Print Width mm', 'Sheet size', 'Disregard Paper', 'Disregard Consumable', 'Paid', 'Notes', 'S_PrinterID', 'S_StockID', 'S_TechID', 'S_Measure', 'S_UnitCost', 'S_StockWidth_mm', 'S_SheetHeight_mm', 'S_ConsRate', 'S_StampedAt', 'S_StampedBy', 'S_SchemaVer')
    $oldRows = @($data | Where-Object { $_.'Job ID' -eq $jobId } | ForEach-Object { $_ | Select-Object $oldCols })
    $oldRows[0].Notes = 'edited in old file'
    $oldCsv = ($oldRows | ConvertTo-Csv -NoTypeInformation)
    $oldFile = Join-Path $workDir 'old-format.csv'
    Set-Content $oldFile (@($lines[0..7]) + $oldCsv) -Encoding UTF8
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $oldFile))
    $jr = 0; for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if ([string](J $i 'Job ID').Value2 -eq $jobId) { $jr = $i } }
    Check ([string](J $jr 'Notes').Value2 -eq 'edited in old file') 'the old-format job row was imported'
    Check (([int](J $jr 'Passes').Value2) -eq 3 -and ([string](J ($jr + 1) 'Row Type').Value2 -eq 'Pass')) 'and its passes were not touched'

    Write-Host ''
    Write-Host '=== An unknown colour imports, keeps its cost, shows a notice ==='
    $ghost = @(); foreach ($ln in $lines[0..8]) { $ghost += $ln }
    foreach ($ln in $lines[9..($lines.Count - 1)]) { $ghost += ($ln -replace ',Black,', ',Ghost Pink,') }
    $ghostFile = Join-Path $workDir 'ghost.csv'
    Set-Content $ghostFile $ghost -Encoding UTF8
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $ghostFile))
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    $jr = 0; for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if ([string](J $i 'Job ID').Value2 -eq $jobId) { $jr = $i } }
    Check ([string](J ($jr + 1) 'Colour').Value2 -eq 'Ghost Pink') 'the colour came in as written'
    Check ([string](J ($jr + 1) 'Status').Value2 -like 'Notice:*') "it shows the notice: '$([string](J ($jr + 1) 'Status').Value2)'"
    Check ((Near ([double](J $jr 'Consumable Cost').Value2) $cons0)) 'and still counts at the stamped rate'

    Write-Host ''
    Write-Host '=== Re-stamp prices ==='
    [void]$xl.Run('ApplyImportConfirmed', $main, $xl.Run('ReadImportRows', $csv.FullName))   # back to known colours
    $jr = 0; for ($i = 1; $i -le $lo.ListRows.Count; $i++) { if ([string](J $i 'Job ID').Value2 -eq $jobId) { $jr = $i } }
    [void]$xl.Run('UnlockSheet', $cws)
    $clo.ListRows(1).Range.Cells(1, (Col $clo 'Cost per m2')).Value2 = [double]0.50          # Black 0.30 -> 0.50
    [void]$xl.Run('UnlockSheet', $pws)
    $pr.Cells(1, (Col $plo 'Template cost')).Value2 = [double]4                              # set-up 3 -> 4
    Check ((Near ([double](J ($jr + 1) 'S_ConsRate').Value2) 0.30)) 'before: the pass still has its stamped 0.30'
    [void]$xl.Run('ReStampAllConfirmed', 1)
    Invoke-ComRetry { $xl.CalculateFullRebuild() } | Out-Null
    Check ((Near ([double](J ($jr + 1) 'S_ConsRate').Value2) 0.50)) 'after: Black re-stamped to 0.50'
    Check ((Near ([double](J ($jr + 2) 'S_ConsRate').Value2) 0.40)) 'Blue unchanged at 0.40'
    Check ((Near ([double](J ($jr + 1) 'S_SetupCost').Value2) 4)) 'set-up cost re-stamped from the printer: 4'
    Check ((Near ([double](J $jr 'Set-up Cost').Value2) 12.0)) 'the job rolls up 3 x 4 = 12.00'
    # An unknown colour is left alone (decision 31).
    W ($jr + 3) 'Colour' 'Not Defined Here'
    $rateBefore = [double](J ($jr + 3) 'S_ConsRate').Value2
    [void]$xl.Run('ReStampAllConfirmed', 1)
    Check ((Near ([double](J ($jr + 3) 'S_ConsRate').Value2) $rateBefore)) 're-stamp leaves a pass with an unknown colour at its stamped rate'

    Write-Host ''
    Write-Host '=== Backup includes the colours table ==='
    [void]$xl.Run('btnBackupWorkbook')
    $cat = @(Get-ChildItem $workDir -Filter 'PrintCosts-*-CATALOG-tblColours-*.csv')
    Check ($cat.Count -eq 1) 'the colours table has its own backup CSV'
    $printersCsv = @(Get-ChildItem $workDir -Filter 'PrintCosts-*-CATALOG-tblPrinters-*.csv')
    $ptext = Get-Content $printersCsv[0].FullName -Raw
    Check ($ptext -match 'Colour mode' -and $ptext -match 'Template cost') 'and the printers backup carries Colour mode and Template cost'
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($script:anyFail) { Write-Host 'FAIL'; exit 1 }
Write-Host 'multi-pass import/export checks complete'
