# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Prerequisite: the Departments sheet (0.10.30, modDepartments / modCatalog
# "department lookup").
#
#   1. The sheet, table, hidden Match list, Settings counter and tab order exist.
#   2. Add row gives a site-prefixed DeptID with Active = Yes, Free = No; a row
#      typed under the table gets its ID and defaults; Aliases are tidied; a
#      name/alias used twice is warned about, not blocked.
#   3. Typing a FREE, ACTIVE department's name (or alias, any case, padded) as a
#      job's Student Name sets both disregards to Yes - the job is still costed
#      (Gross unchanged, Disregarded = Gross, Chargeable 0). Free = No, inactive
#      and non-department names change nothing; a hand-set value is never
#      re-applied by an unrelated edit; moving to a non-department leaves them.
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

    Write-Host '=== The sheet ==='
    $dws = Invoke-ComRetry { $wb.Worksheets('Departments') }
    $dlo = $dws.ListObjects('tblDepartments')
    foreach ($h in 'DeptID', 'Name', 'Aliases', 'Free', 'Active', 'Notes', 'Match list', 'Dis Paper %', 'Dis Cons %', 'Allowance', 'Allowance period') {
        Check ((Col $dlo $h) -gt 0) "column $h exists"
    }
    Check ($dlo.ListColumns('Match list').Range.EntireColumn.Hidden -eq $true) 'Match list is hidden'
    Check ([int]$dws.Visible -eq -1) 'sheet is visible'
    $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }
    $iDep = $names.IndexOf('Departments'); $iSet = $names.IndexOf('Settings'); $iPap = $names.IndexOf('Papers')
    Check ($iDep -gt $iPap -and $iDep -lt $iSet) 'tab order: Papers, Departments, Settings'
    $hwm = $wb.Names.Item('SET_DEPT_ID_HWM').RefersToRange
    Check ($hwm.Locked -eq $true) 'DEPT_ID_HWM value cell is locked (Read-only)'
    Check ([int]$hwm.Value2 -eq 0) 'DEPT_ID_HWM starts at 0'

    function Set-Dept([int]$rowNo, [string]$name, [string]$aliases, [string]$free, [string]$active) {
        $r = $dlo.ListRows($rowNo).Range
        [void]$xl.Run('UnlockSheet', $dws)
        $r.Cells(1, (Col $dlo 'Name')).Value2 = $name
        if ($aliases -ne $null) { $r.Cells(1, (Col $dlo 'Aliases')).Value2 = $aliases }
        if ($free)   { $r.Cells(1, (Col $dlo 'Free')).Value2 = $free }
        if ($active) { $r.Cells(1, (Col $dlo 'Active')).Value2 = $active }
    }
    function Cell-Of([int]$rowNo, [string]$h) { return $dlo.ListRows($rowNo).Range.Cells(1, (Col $dlo $h)) }

    Write-Host ''
    Write-Host '=== Add row / typed rows ==='
    [void]$dws.Activate()
    [void]$xl.Run('AddCatalogRow', 'tblDepartments')
    $id1 = [string](Cell-Of 1 'DeptID').Value2
    Check ($id1 -match "^$([regex]::Escape($site))-DEP-0001$") "first ID is $id1"
    Check ([string](Cell-Of 1 'Active').Value2 -eq 'Yes') 'Active defaults to Yes'
    Check ([string](Cell-Of 1 'Free').Value2 -eq 'No') 'Free defaults to No'
    Set-Dept 1 'Fine Art' ' FA ; Art School;;' 'Yes' $null
    Check ([string](Cell-Of 1 'Aliases').Value2 -eq 'FA;Art School') "aliases tidied to '$([string](Cell-Of 1 'Aliases').Value2)'"
    Check ([string](Cell-Of 1 'Match list').Value2 -eq ';fine art;fa;art school;') "match list is '$([string](Cell-Of 1 'Match list').Value2)'"

    [void]$xl.Run('AddCatalogRow', 'tblDepartments')
    Set-Dept 2 'Textiles' '' $null $null
    [void]$xl.Run('AddCatalogRow', 'tblDepartments')
    Set-Dept 3 'Old Dept' '' 'Yes' 'No'
    # A row typed straight under the table (no Add row) still gets ID + defaults.
    [void]$xl.Run('AddCatalogRow', 'tblDepartments')
    [void]$xl.Run('UnlockSheet', $dws)
    (Cell-Of 4 'Name').Value2 = 'Typed Dept'
    Check ([string](Cell-Of 4 'DeptID').Value2 -match "^$([regex]::Escape($site))-DEP-0004$") 'typed row gets DEP-0004'
    Check ([string](Cell-Of 4 'Active').Value2 -eq 'Yes' -and [string](Cell-Of 4 'Free').Value2 -eq 'No') 'typed row gets Active=Yes, Free=No'
    Check ([int]$hwm.Value2 -eq 4) 'counter advanced to 4'

    $xl.Run('SetQuiet', $true)
    (Cell-Of 4 'Name').Value2 = 'fa'
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*used more than once*') 'a name that is another department''s alias is warned about'
    Check ([string](Cell-Of 4 'Name').Value2 -eq 'fa') '...and not blocked'
    (Cell-Of 4 'Name').Value2 = 'Typed Dept'

    Write-Host ''
    Write-Host '=== Job entry ==='
    $main = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    $lo = $main.ListObjects('tblJobs_MAIN')
    [void]$xl.Run('UnlockSheet', $main)
    function J([int]$rowNo, [string]$h) { return $lo.ListRows($rowNo).Range.Cells(1, (Col $lo $h)) }
    function Reset-Row([int]$rowNo) {
        (J $rowNo 'Disregard Paper').Value2 = 'No'
        (J $rowNo 'Disregard Consumable').Value2 = 'No'
    }
    function Dis([int]$rowNo) { return "$([string](J $rowNo 'Disregard Paper').Value2)/$([string](J $rowNo 'Disregard Consumable').Value2)" }

    Reset-Row 1
    $gross = [double](J 1 'Gross Cost').Value2
    Check ($gross -gt 0) "row 1 has a cost to disregard ($gross)"
    Check ([double](J 1 'Chargeable Cost').Value2 -eq $gross) 'charged in full to begin with'

    (J 1 'Student Name').Value2 = ' fine ART '
    Check ((Dis 1) -eq 'Yes/Yes') "name in odd case and padded: $(Dis 1)"
    Check ([double](J 1 'Gross Cost').Value2 -eq $gross) 'Gross unchanged: still costed'
    Check ([double](J 1 'Chargeable Cost').Value2 -eq 0) 'Chargeable is 0'
    Check ([Math]::Abs([double](J 1 'Disregarded').Value2 - $gross) -lt 0.0001) 'Disregarded equals Gross'

    Reset-Row 2
    (J 2 'Student Name').Value2 = 'FA'
    Check ((Dis 2) -eq 'Yes/Yes') 'an alias matches'
    Reset-Row 3
    (J 3 'Student Name').Value2 = 'Textiles'
    Check ((Dis 3) -eq 'No/No') 'Free = No department changes nothing'
    Reset-Row 4
    (J 4 'Student Name').Value2 = 'Old Dept'
    Check ((Dis 4) -eq 'No/No') 'inactive department changes nothing'
    Reset-Row 5
    (J 5 'Student Name').Value2 = 'Alex Student'
    Check ((Dis 5) -eq 'No/No') 'an ordinary student changes nothing'

    # Hand override is never re-applied by an unrelated edit.
    (J 1 'Disregard Paper').Value2 = 'No'
    (J 1 'Notes').Value2 = 'edited'
    Check ((Dis 1) -eq 'No/Yes') "a hand-set value survives other edits: $(Dis 1)"
    (J 1 'Student No').Value2 = 'X123'
    Check ((Dis 1) -eq 'No/Yes') 'and editing Student No does not apply the rule'
    # ...but re-entering the department re-stamps.
    (J 1 'Student Name').Value2 = 'Textiles'
    Check ((Dis 1) -eq 'No/Yes') 'switching to a non-free department leaves them'
    (J 1 'Student Name').Value2 = 'Fine Art'
    Check ((Dis 1) -eq 'Yes/Yes') 'entering a free department again re-stamps'
    (J 1 'Student Name').Value2 = 'Alex Student'
    Check ((Dis 1) -eq 'Yes/Yes') 'moving to a student leaves the values'

    # Renaming the department on the sheet: the old spelling no longer matches.
    Reset-Row 6
    (Cell-Of 1 'Name').Value2 = 'Fine Arts'
    (J 6 'Student Name').Value2 = 'Fine Art'
    Check ((Dis 6) -eq 'No/No') 'after renaming the department, the old name no longer matches'
    (J 6 'Student Name').Value2 = 'Fine Arts'
    Check ((Dis 6) -eq 'Yes/Yes') 'the new name does'

    Write-Host ''
    Write-Host '=== Reporting ==='
    # State now: the department is called Fine Arts. Make three department jobs
    # (two Fine Arts - one by alias, odd case - and one Textiles) and leave
    # the rest as they are.
    (J 1 'Student Name').Value2 = 'fine arts'
    (J 2 'Student Name').Value2 = 'FA'
    (J 3 'Student Name').Value2 = 'Textiles'
    (J 4 'Student Name').Value2 = 'Old Dept'
    (J 6 'Student Name').Value2 = 'Zed Student'     # was left on the department name by the section above
    $xl.CalculateFullRebuild()
    $data = $wb.Worksheets('_Data')
    $hdrs = @{}
    for ($c = 1; $c -le 60; $c++) { $h = [string]$data.Cells(9, $c).Value2; if ($h) { $hdrs[$h] = $c } }
    Check ($hdrs.ContainsKey('Department')) '_Data has a Department column'
    $spill = $data.Range('A10').SpillingToRange
    $rows = $spill.Rows.Count
    $dc = $hdrs['Department']; $sc = $hdrs['Student Name']
    $v = $data.Range($data.Cells(10, 1), $data.Cells(9 + $rows, $dc)).Value2
    $fa = 0; $tx = 0; $od = 0; $blank = 0; $badNonBlank = 0
    for ($r = 1; $r -le $rows; $r++) {
        $d = [string]$v[$r, $dc]; $s = [string]$v[$r, $sc]
        switch ($d) { 'Fine Arts' { $fa++ } 'Textiles' { $tx++ } 'Old Dept' { $od++ } '' { $blank++ } default { $badNonBlank++ } }
    }
    Check ($fa -eq 2) "two jobs are classed Fine Arts (the name and the alias) - got $fa"
    Check ($tx -eq 1) "one is Textiles - got $tx"
    Check ($od -eq 1) "an INACTIVE department still classifies its jobs - got $od"
    Check ($badNonBlank -eq 0) 'no other department name appears'
    Check ($blank -eq ($rows - 4)) "every other row is blank (got $blank of $rows)"

    $rep = Invoke-ComRetry { $wb.Worksheets('Reports') }
    [void]$rep.Activate()
    function Jobs { $xl.CalculateFullRebuild(); [int]$rep.Range('I9').Text }
    $all = Jobs
    $rep.Range('F11').Value2 = 'Fine Arts'
    Check ((Jobs) -eq 2) 'Reports Department filter = Fine Arts shows 2 jobs'
    $rep.Range('F11').Value2 = 'Textiles'
    Check ((Jobs) -eq 1) '...Textiles shows 1'
    $rep.Range('F11').Value2 = 'Old Dept'
    Check ((Jobs) -eq 1) '...an inactive department can still be filtered'
    Check ([int]$rep.Range('F11').Validation.Type -eq 3) 'F11 is a dropdown'
    $rep.Range('F11').Value2 = ''
    Check ((Jobs) -eq $all) 'cleared, back to all jobs'
    $rep.Range('F11').Value2 = 'Fine Arts'
    [void]$xl.Run('ClearReportFilters', $rep)
    Check ([string]$rep.Range('F11').Text -eq '') 'Clear all filters clears the Department box'
    $rep.Range('F11').Value2 = 'Fine Arts'
    Check ([bool]$xl.Run('HasActiveFilter', $rep)) 'a set Department filter counts as an active filter'
    $rep.Range('F11').Value2 = ''

    # The Reports breakdown (column AI) lists departments only.
    $bc = $rep.Cells(10, 35).Address($false, $false) -replace '\d+', ''
    Check ([string]$rep.Cells(10, 35).Text -eq 'By department') 'Reports has a By department block'
    $names = @(); $jobsCol = @{}
    for ($r = 12; $r -le 20; $r++) { $n = [string]$rep.Cells($r, 35).Text; if ($n) { $names += $n; $jobsCol[$n] = [int]$rep.Cells($r, 36).Value2 } }
    Check (($names -join ',') -eq 'Fine Arts,Old Dept,Textiles') "breakdown keys: $($names -join ',')"
    Check ($jobsCol['Fine Arts'] -eq 2 -and $jobsCol['Textiles'] -eq 1) 'breakdown job counts are right'

    # The Summary box reconciles to the Summary totals.
    $sum = Invoke-ComRetry { $wb.Worksheets('Summary') }
    [void]$sum.Activate()
    $xl.CalculateFullRebuild()
    Check ([string]$sum.Range('I5').Text -eq 'By department') 'Summary has the By department box'
    $dn = @(); $jobsT = 0; $gT = 0.0; $dsT = 0.0; $chT = 0.0
    for ($r = 7; $r -le 21; $r++) {
        $n = [string]$sum.Cells($r, 9).Text
        if ($r -le 20 -and $n) { $dn += $n }
        if ($r -eq 21 -or $n) {
            $jobsT += [double]$sum.Cells($r, 10).Value2; $gT += [double]$sum.Cells($r, 11).Value2
            $dsT += [double]$sum.Cells($r, 12).Value2; $chT += [double]$sum.Cells($r, 13).Value2
        }
    }
    Check (($dn -join ',') -eq 'Fine Arts,Old Dept,Textiles') "box lists departments with jobs, sorted: $($dn -join ',')"
    Check ([string]$sum.Range('I21').Text -eq 'Everyone else') 'last line is Everyone else'
    Check ([Math]::Abs($jobsT - [double]$sum.Range('B24').Value2) -lt 0.001) "jobs reconcile to the total ($jobsT)"
    Check ([Math]::Abs($gT - [double]$sum.Range('D24').Value2) -lt 0.005) "gross reconciles ($gT)"
    Check ([Math]::Abs($dsT - [double]$sum.Range('F24').Value2) -lt 0.005) "disregarded reconciles ($dsT)"
    Check ([Math]::Abs($chT - [double]$sum.Range('H24').Value2) -lt 0.005) "chargeable reconciles ($chT)"
    Check ([double]$sum.Range('L7').Value2 -gt 0) 'Fine Arts shows a disregarded amount (free department)'

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
if ($script:anyFail) { Write-Host 'FAIL' } else { Write-Host 'PASS' }
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
if ($script:anyFail) { exit 1 }
