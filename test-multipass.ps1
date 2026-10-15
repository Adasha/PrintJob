# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Prerequisite: multi-pass (per-colour) costing (0.11.0, docs/ARCHITECTURE.md
# section 17, docs/multipass-costing-design.md, modPasses).
#
#   1. The Consumables sheet and a multi-pass printer: IDs, defaults, settings.
#   2. A job on the multi-pass printer: no flat rate, "needs a pass" until it has one.
#   3. Add pass: rows beneath the job, Parent / Pass / Row Type, outline level 2,
#      Template cost stamped; the job's Passes and Set-up Cost roll up.
#   4. Colours stamp their rate; ink is area x rate per pass; the job sums its passes;
#      Gross = Paper + Consumable + Set-up; Disregard Consumable covers ink AND set-up.
#   5. A colour of another consumable type is a problem; an unknown colour is a
#      notice (yellow level), not a problem, and keeps counting.
#   6. A job with passes keeps its printer; a single-pass job refuses Add pass.
#   7. Removing a pass renumbers; collapse / expand; Repeat job copies the group;
#      deleting a job takes its passes; "Print jobs" counts jobs, not passes.
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

    Write-Host '=== Consumables sheet and a multi-pass printer ==='
    $cws = Invoke-ComRetry { $wb.Worksheets('Consumables') }
    $clo = $cws.ListObjects('tblColours')
    foreach ($h in 'ColourID', 'Consumable type', 'Colour', 'Cost per m2', 'Active') { Check ((Col $clo $h) -gt 0) "colours column $h exists" }
    $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }
    Check (($names.IndexOf('Consumables') -gt $names.IndexOf('Papers')) -and ($names.IndexOf('Consumables') -lt $names.IndexOf('Settings'))) 'tab order: Papers, Consumables ... Settings'
    $hwm = $wb.Names.Item('SET_COLOUR_ID_HWM').RefersToRange
    Check ([int]$hwm.Value2 -eq 0 -and $hwm.Locked) 'COLOUR_ID_HWM starts at 0 and is read-only'

    function Set-Colour([int]$rowNo, [string]$name, [string]$type, [double]$rate) {
        $r = $clo.ListRows($rowNo).Range
        [void]$xl.Run('UnlockSheet', $cws)
        $r.Cells(1, (Col $clo 'Colour')).Value2 = $name
        $r.Cells(1, (Col $clo 'Consumable type')).Value2 = $type
        $r.Cells(1, (Col $clo 'Cost per m2')).Value2 = $rate
    }
    [void]$cws.Activate()
    foreach ($row in @(@('Black', 'Risograph', 0.30), @('Blue', 'Risograph', 0.40), @('Fluoro Orange', 'Risograph', 0.55), @('Cyan', 'Ink', 0.10))) {
        [void]$xl.Run('AddCatalogRow', 'tblColours')
        $n = $clo.ListRows.Count
        Set-Colour $n $row[0] $row[1] $row[2]
    }
    Check ([string]$clo.ListRows(1).Range.Cells(1, (Col $clo 'ColourID')).Value2 -eq "$site-CLR-0001") "first colour ID is $site-CLR-0001"
    Check ([string]$clo.ListRows(1).Range.Cells(1, (Col $clo 'Active')).Value2 -eq 'Yes') 'Active defaults to Yes'
    Check ([int]$hwm.Value2 -eq 4) 'counter advanced to 4'

    $pws = Invoke-ComRetry { $wb.Worksheets('Printers') }
    $plo = $pws.ListObjects('tblPrinters')
    [void]$pws.Activate()
    [void]$xl.Run('AddCatalogRow', 'tblPrinters')
    $pn = $plo.ListRows.Count
    $pr = $plo.ListRows($pn).Range
    [void]$xl.Run('UnlockSheet', $pws)
    $pr.Cells(1, (Col $plo 'Model')).Value2 = 'Test RISO'
    $pr.Cells(1, (Col $plo 'Consumable type')).Value2 = 'Risograph'
    $pr.Cells(1, (Col $plo 'Cost per m2')).Value2 = 0.99
    $pr.Cells(1, (Col $plo 'Colour mode')).Value2 = 'multi-pass'
    $pr.Cells(1, (Col $plo 'Template cost')).Value2 = 3
    $pr.Cells(1, (Col $plo 'Max sheet size')).Value2 = 'A1'
    $pr.Cells(1, (Col $plo 'Active')).Value2 = 'Yes'

    $main = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    $lo = $main.ListObjects('tblJobs_MAIN')
    $loc = $main.Names.Item('LOC_Printers').RefersToRange
    [void]$xl.Run('UnlockSheet', $main)
    $loc.Value2 = ([string]$loc.Value2 + ';Test RISO')
    function J([int]$rowNo, [string]$h) { return (Invoke-ComRetry -RetryOnNull { $lo.ListRows($rowNo).Range.Cells(1, (Col $lo $h)) }) }
    # A write from here: the sheet is unprotected first (the workbook's own code re-protects it), and
    # the Worksheet_Change handling runs exactly as it would for a person typing.
    function W([int]$rowNo, [string]$h, $v) { [void]$xl.Run('UnlockSheet', $main); if ($v -is [int] -or $v -is [double]) { $v = ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }; $cell = J $rowNo $h; Invoke-ComRetry { $cell.Value2 = $v } | Out-Null }
    function Pick([int]$rowNo) { [void]$main.Activate(); [void]((J $rowNo 'Date/Time').Select()) }
    function Level([int]$rowNo) { return [int]$lo.ListRows($rowNo).Range.EntireRow.OutlineLevel }
    function Hidden([int]$rowNo) { return [bool]$lo.ListRows($rowNo).Range.EntireRow.Hidden }
    $sampleRows = $lo.ListRows.Count

    Write-Host ''
    Write-Host '=== A job on the multi-pass printer ==='
    [void]$main.Activate()
    [void]$xl.Run('AddPrintJob', $main)
    $job = $lo.ListRows.Count
    $jobId = [string](J $job 'Job ID').Value2
    Check ([string](J $job 'Row Type').Value2 -eq 'Job') 'a new job is a Job row'
    Check ((Level $job) -eq 1) 'at outline level 1'
    W $job 'Student Name' 'Pass Tester'
    W $job 'Technician' ([string](J 1 'Technician').Value2)
    W $job 'Printer' 'Test RISO'
    W $job 'Paper Stock' 'Gloss 200gsm SRA3 sheet'
    W $job 'Qty' 100
    Check ([string](J $job 'S_PrinterID').Value2 -ne '') 'the job is stamped with the printer'
    Check ([string](J $job 'S_ConsRate').Value2 -eq '') 'a multi-pass job carries no flat consumable rate'
    Check ([string](J $job 'Status').Value2 -like '*At least one colour pass is required*') "Status asks for a pass: '$([string](J $job 'Status').Value2)'"
    Check ([string](J $job 'Passes').Value2 -eq '') 'Passes is blank with none'
    $paper = [double](J $job 'Paper Cost').Value2
    $area = [double](J $job 'Area m2').Value2
    Check ([Math]::Abs($area - 14.4) -lt 0.0001) "area is 14.4 m2 (got $area)"
    Check ([Math]::Abs($paper - 31.0) -lt 0.0001) "paper cost is 31.00 (got $paper)"
    Check ([double](J $job 'Consumable Cost').Value2 -eq 0) 'consumable cost is 0 with no passes'

    Write-Host ''
    Write-Host '=== Add pass ==='
    Pick $job
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('AddPass', $main)
    [void]$xl.Run('AddPass', $main)
    [void]$xl.Run('AddPass', $main)
    Check ($lo.ListRows.Count -eq $job + 3) "three pass rows added beneath the job (rows now $($lo.ListRows.Count))"
    foreach ($i in 1..3) {
        $r = $job + $i
        Check ([string](J $r 'Row Type').Value2 -eq 'Pass') "row $r is a Pass row"
        Check ([string](J $r 'Parent').Value2 -eq $jobId) "pass $i has Parent $jobId"
        Check ([int](J $r 'Pass').Value2 -eq $i) "pass number $i"
        Check ((Level $r) -eq 2) "pass $i is at outline level 2"
        Check ([double](J $r 'S_SetupCost').Value2 -eq 3) "pass $i carries the printer's Template cost"
        Check ([string](J $r 'Job ID').Value2 -eq '') "pass $i has no Job ID of its own"
    }
    Check ((Level $job) -eq 1) 'the job stays at level 1'
    Check ([int](J $job 'Passes').Value2 -eq 3) 'Passes counts 3'
    Check ([double](J $job 'Set-up Cost').Value2 -eq 9) 'Set-up Cost rolls up to 9.00'
    Check ([string](J ($job + 1) 'Status').Value2 -like '*Colour required*') "an unfilled pass asks for a colour: '$([string](J ($job + 1) 'Status').Value2)'"

    Write-Host ''
    Write-Host '=== Colours and costing ==='
    W ($job + 1) 'Colour' 'Black'
    W ($job + 2) 'Colour' 'Blue'
    W ($job + 3) 'Colour' 'Fluoro Orange'
    Check ([double](J ($job + 1) 'S_ConsRate').Value2 -eq 0.30) 'Black stamps 0.30'
    Check ([string](J ($job + 1) 'S_ColourID').Value2 -eq "$site-CLR-0001") 'and its ColourID'
    Check ([Math]::Abs([double](J ($job + 1) 'Consumable Cost').Value2 - 4.32) -lt 0.0001) 'pass 1 ink = 14.4 x 0.30 = 4.32'
    Check ([Math]::Abs([double](J ($job + 2) 'Consumable Cost').Value2 - 5.76) -lt 0.0001) 'pass 2 ink = 5.76'
    Check ([Math]::Abs([double](J ($job + 3) 'Consumable Cost').Value2 - 7.92) -lt 0.0001) 'pass 3 ink = 7.92'
    Check ([Math]::Abs([double](J $job 'Consumable Cost').Value2 - 18.0) -lt 0.0001) 'the job sums its ink: 18.00'
    Check ([Math]::Abs([double](J $job 'Set-up Cost').Value2 - 9.0) -lt 0.0001) 'and its set-up: 9.00'
    Check ([Math]::Abs([double](J $job 'Gross Cost').Value2 - 58.0) -lt 0.0001) 'Gross = 31 + 18 + 9 = 58.00'
    Check ([Math]::Abs([double](J $job 'Chargeable Cost').Value2 - 58.0) -lt 0.0001) 'Chargeable is the full 58.00'
    Check ([string](J ($job + 1) 'Gross Cost').Value2 -eq '') 'a pass row has no gross of its own'
    Check ([string](J $job 'Status').Value2 -eq 'OK') "the job is OK: '$([string](J $job 'Status').Value2)'"
    Check ([string](J ($job + 2) 'Status').Value2 -eq 'OK') 'a coloured pass is OK'

    W $job 'Disregard Consumable' 'Yes'
    Check ([Math]::Abs([double](J $job 'Chargeable Cost').Value2 - 31.0) -lt 0.0001) 'Disregard Consumable drops ink AND set-up: chargeable 31.00'
    Check ([Math]::Abs([double](J $job 'Disregarded').Value2 - 27.0) -lt 0.0001) 'Disregarded = 27.00'
    W $job 'Disregard Consumable' 'No'

    # A later price change never reaches the stamped rate (snapshot rule, decision 23).
    [void]$cws.Activate()
    [void]$xl.Run('UnlockSheet', $cws)
    $clo.ListRows(1).Range.Cells(1, (Col $clo 'Cost per m2')).Value2 = 0.90
    Check ([Math]::Abs([double](J ($job + 1) 'Consumable Cost').Value2 - 4.32) -lt 0.0001) 'editing the colour price leaves the pass at its stamped rate'
    [void]$xl.Run('UnlockSheet', $cws)
    $clo.ListRows(1).Range.Cells(1, (Col $clo 'Cost per m2')).Value2 = 0.30

    Write-Host ''
    Write-Host '=== Problems versus notices ==='
    W ($job + 2) 'Colour' 'Mystery'
    $st = [string](J ($job + 2) 'Status').Value2
    Check ($st -like 'Notice:*') "an unknown colour shows a notice: '$st'"
    Check ([string](J ($job + 2) 'H_Issues').Value2 -eq '') '...and is not a problem'
    Check ([Math]::Abs([double](J ($job + 2) 'Consumable Cost').Value2 - 5.76) -lt 0.0001) '...and still counts at its stamped rate'
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('CheckSheet', $main)
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*Notices (not counted as problems)*') 'Check sheet reports notices apart'
    $problemPart = ($log -split 'Notices \(not counted')[0]
    Check ($problemPart -notlike '*Pass *') '...and none of the passes is listed as a problem'
    W ($job + 2) 'Colour' 'Cyan'
    Check ([string](J ($job + 2) 'Status').Value2 -like '*different consumable type*') "a colour of another type is a problem: '$([string](J ($job + 2) 'Status').Value2)'"
    W ($job + 2) 'Colour' 'Blue'
    Check ([string](J ($job + 2) 'Status').Value2 -eq 'OK') 'a Blue pass is OK again'
    # The Pass column is locked for a person; here it is written with events off to prove the check.
    $xl.EnableEvents = $false
    W ($job + 3) 'Pass' 7
    $xl.EnableEvents = $true
    Check ([string](J ($job + 3) 'H_Issues').Value2 -like '*out of sequence*') 'a pass number out of sequence is a problem'
    $xl.EnableEvents = $false
    W ($job + 3) 'Pass' 3
    $xl.EnableEvents = $true
    Check ([string](J ($job + 3) 'Status').Value2 -eq 'OK') "put back, the pass is OK ('$([string](J ($job + 3) 'Status').Value2)')"

    Write-Host ''
    Write-Host '=== The job keeps its printer; single-pass jobs refuse a pass ==='
    $xl.Run('SetQuiet', $true)
    W $job 'Printer' 'Xerox Versant 180'
    Check ([string](J $job 'Printer').Value2 -eq 'Test RISO') 'changing the printer of a job with passes is reverted'
    Check (([string]$xl.Run('QuietLog')) -like '*cannot be changed*') '...with a message'
    Pick 1
    $xl.Run('SetQuiet', $true)
    [void]$xl.Run('AddPass', $main)
    Check (([string]$xl.Run('QuietLog')) -like '*Select a multi-pass job, or one of its passes, first.*') 'Add pass on a single-pass job shows the short message'
    Check ($lo.ListRows.Count -eq $job + 3) '...and adds nothing'

    Write-Host ''
    Write-Host '=== Collapse and expand ==='
    Pick ($job + 2)                     # a pass selected: acts on its job
    [void]$xl.Run('TogglePasses', $main)
    Check ((Hidden ($job + 1)) -and (Hidden ($job + 2)) -and (Hidden ($job + 3))) 'Toggle passes collapses the job from one of its passes'
    Check (-not (Hidden $job)) '...the job row stays visible'
    [void]$xl.Run('TogglePasses', $main)
    Check (-not ((Hidden ($job + 1)) -or (Hidden ($job + 3)))) 'a second click expands'
    $lo.ListRows($job + 1).Range.EntireRow.Hidden = $true     # a mix
    [void]$xl.Run('ToggleAllPasses', $main)
    Check (-not (Hidden ($job + 1))) 'Toggle all with a mix expands everything first'
    [void]$xl.Run('ToggleAllPasses', $main)
    Check ((Hidden ($job + 1)) -and (Hidden ($job + 3))) 'all expanded -> collapse all'
    [void]$xl.Run('ToggleAllPasses', $main)
    Check (-not (Hidden ($job + 2))) 'all collapsed -> expand all'

    Write-Host ''
    Write-Host '=== Remove a pass: renumbering ==='
    $detail = 'test'
    [void]$xl.Run('DeletePass', $main, $lo, $job, ($job + 2), $detail)
    Check ($lo.ListRows.Count -eq $job + 2) 'one row fewer'
    Check (([int](J ($job + 1) 'Pass').Value2 -eq 1) -and ([int](J ($job + 2) 'Pass').Value2 -eq 2)) 'the rest are numbered 1, 2'
    Check ([string](J ($job + 2) 'Colour').Value2 -eq 'Fluoro Orange') '...and it was the middle one that went'
    Check ([int](J $job 'Passes').Value2 -eq 2) 'Passes now 2'
    Check ([Math]::Abs([double](J $job 'Set-up Cost').Value2 - 6.0) -lt 0.0001) 'Set-up Cost now 6.00'
    Check ((Level ($job + 1)) -eq 2 -and (Level ($job + 2)) -eq 2 -and (Level $job) -eq 1) 'outline levels still line up'

    Write-Host ''
    Write-Host '=== Repeat job copies the group ==='
    Pick ($job + 1)
    [void]$xl.Run('RepeatJob', $main)
    $job2 = $job + 3
    Check ($lo.ListRows.Count -eq $job + 5) "repeat added a job and two passes (rows $($lo.ListRows.Count))"
    $id2 = [string](J $job2 'Job ID').Value2
    Check ($id2 -ne '' -and $id2 -ne $jobId) "the copy has its own Job ID ($id2)"
    Check (([string](J ($job2 + 1) 'Parent').Value2 -eq $id2) -and ([string](J ($job2 + 2) 'Parent').Value2 -eq $id2)) 'its passes belong to it'
    Check (([string](J ($job2 + 1) 'Colour').Value2 -eq 'Black') -and ([string](J ($job2 + 2) 'Colour').Value2 -eq 'Fluoro Orange')) 'with the same colours'
    Check ((Level $job2) -eq 1 -and (Level ($job2 + 1)) -eq 2 -and (Level ($job2 + 2)) -eq 2) 'and the right outline levels'
    Check ([int](J $job2 'Passes').Value2 -eq 2) 'Passes 2 on the copy'

    Write-Host ''
    Write-Host '=== Counts and consolidation ==='
    $jobs = ($sampleRows + 2)
    Check ([int]$main.Range('B13').Value2 -eq $jobs) "Print jobs counts jobs, not passes ($([int]$main.Range('B13').Value2) of $jobs)"
    $dws = $wb.Worksheets('_Data')
    $spill = $dws.Range('A10').SpillingToRange
    Check ([int]$spill.Rows.Count -eq $jobs) "the consolidated range holds jobs only ($([int]$spill.Rows.Count) of $jobs)"

    Write-Host ''
    Write-Host '=== Delete a job takes its passes ==='
    $before = $lo.ListRows.Count
    [void]$xl.Run('DeleteJobGroupAt', $main, $lo, $job2, 'test')
    Check ($lo.ListRows.Count -eq $before - 3) 'all three rows removed'
    Check ((Level ($job + 1)) -eq 2 -and (Level $job) -eq 1) 'the first group is undisturbed'
    [void]$xl.Run('DeleteJobGroupAt', $main, $lo, $job, 'test')
    Check ($lo.ListRows.Count -eq $sampleRows) 'back to the sample rows'
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($script:anyFail) { Write-Host 'FAIL'; exit 1 }
Write-Host 'multi-pass checks complete'
