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
