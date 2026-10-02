# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Job-table Data Validation is bound to columns by header NAME, so it cannot
# drift when columns move (docs/CHANGELOG.md 0.9.15; EnsureJobColumnValidation).
#
#   1. every rule sits on the right column, first and last body row
#   2. no other column carries validation
#   3. nothing is left on the cells below the table (the shipped template once
#      had a stale Yes/No list on what had become Sheet size, rows 28-2010)
#   4. a row added to the table gets the same rules
#   5. columns moved with the old Cut + Insert Shift:=xlToRight (the mechanism
#      that spread rules across columns) are repaired by one BindColumns
#
# Drives a COPY of src\PrintJob.xlsm in %TEMP% (see test-validation.ps1 for why),
# closes without saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')

$xlValidateDecimal = 2; $xlValidateList = 3; $xlValidateDate = 4

# Header name -> expected validation type, and the list text where it is fixed.
$rules = [ordered]@{
    'Date/Time'            = @{ Type = $xlValidateDate }
    'Technician'           = @{ Type = $xlValidateList }
    'Printer'              = @{ Type = $xlValidateList }
    'Paper Stock'          = @{ Type = $xlValidateList }
    'Qty'                  = @{ Type = $xlValidateDecimal }
    'Print Width mm'       = @{ Type = $xlValidateDecimal }
    'Sheet size'           = @{ Type = $xlValidateList; F1 = '=RNG_STD_SIZES' }
    'Disregard Paper'      = @{ Type = $xlValidateList; F1 = 'Yes,No' }
    'Disregard Consumable' = @{ Type = $xlValidateList; F1 = 'Yes,No' }
    'Paid'                 = @{ Type = $xlValidateList; F1 = 'Yes,No' }
}

# -1 = no validation. Excel hands back $null for that (no throw), so both are covered.
function VType($cell) {
    try { $t = $cell.Validation.Type } catch { return -1 }
    if ($null -eq $t) { return -1 }
    return [int]$t
}
function VF1($cell)   { try { return [string]$cell.Validation.Formula1 } catch { return '' } }

# Checks one row of the table against $rules, looking every column up by name.
function Check-Row($lo, [int]$rowNo, [string]$label) {
    foreach ($ci in 1..$lo.ListColumns.Count) {
        $name = $lo.ListColumns($ci).Name
        $cell = $lo.ListRows($rowNo).Range.Cells(1, $ci)
        $t = VType $cell
        if ($rules.Contains($name)) {
            $want = $rules[$name]
            $ok = ($t -eq $want.Type)
            if ($ok -and $want.F1) { $ok = ((VF1 $cell) -eq $want.F1) }
            Check $ok ("{0}: '{1}' rule ok (type {2}, '{3}')" -f $label, $name, $t, (VF1 $cell))
        } else {
            Check ($t -eq -1) ("{0}: '{1}' has no validation" -f $label, $name)
        }
    }
}

function Check-Below($ws, $lo, [string]$label) {
    $first = $lo.Range.Row + $lo.Range.Rows.Count
    $bad = @()
    foreach ($r in @($first, ($first + 4), 2010)) {
        foreach ($ci in 1..$lo.ListColumns.Count) {
            $c = $ws.Cells($r, $lo.Range.Column + $ci - 1)
            if ((VType $c) -ne -1) { $bad += $c.Address($false, $false) }
        }
    }
    Check ($bad.Count -eq 0) ("{0}: no validation below the table{1}" -f $label, $(if ($bad) { ' - stray on ' + ($bad -join ', ') } else { '' }))
}

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
    $ws = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    $lo = $ws.ListObjects('tblJobs_MAIN')

    Write-Host '=== as shipped, after setup ==='
    Check-Row $lo 1 'first row'
    Check-Row $lo $lo.ListRows.Count 'last row'
    Check-Below $ws $lo 'shipped'

    Write-Host '=== a row added to the table ==='
    $xl.Run('SetQuiet', $true)
    $ws.Activate()
    $before = $lo.ListRows.Count
    [void]$xl.Run('btnAddPrintJob')
    Check ($lo.ListRows.Count -eq $before + 1) ("row added ({0} -> {1})" -f $before, $lo.ListRows.Count)
    Check-Row $lo $lo.ListRows.Count 'added row'
    Check-Below $ws $lo 'after add'

    Write-Host '=== columns moved with Cut + Insert, then one BindColumns ==='
    $ws.Unprotect()
    $xlShiftToRight = -4161
    # Paid in front of Qty, Disregard Paper in front of Student Name - one move
    # past rule columns, one across plain ones, both the old reorder's pattern.
    foreach ($mv in @(@('Paid', 'Qty'), @('Disregard Paper', 'Student Name'))) {
        [void]$lo.ListColumns($mv[0]).Range.Cut()
        [void]$lo.ListColumns($mv[1]).Range.Insert($xlShiftToRight)
    }
    Check ((Col $lo 'Paid') -lt (Col $lo 'Qty')) 'Paid now sits before Qty'
    $xl.Run('BindColumns', $ws)
    Check-Row $lo 1 'reordered, first row'
    Check-Row $lo $lo.ListRows.Count 'reordered, last row'
    Check-Below $ws $lo 'reordered'

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
