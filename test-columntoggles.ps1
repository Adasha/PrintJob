# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# The Show: checkboxes on a print room and the workbook-wide internal columns
# setting (0.12.0, docs/column-toggles-design.md). A test cannot click, so the
# choices go through modColumns.SetColumnGroup, which chkShowColumn also calls.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$EPSON = 'Epson SureColor P9500'
$GROUPS = [ordered]@{
    passes    = @('Passes', 'Colour')
    cost      = @('Paper Cost', 'Consumable Cost', 'Set-up Cost', 'Gross Cost', 'Disregarded', 'Area m2')
    disregard = @('Disregard Paper', 'Disregard Consumable')
    size      = @('Print Width mm', 'Sheet size')
    tech      = @('Technician')
    printer   = @('Printer')
}
$INTERNAL = 'Job ID', 'Row Type', 'Parent', 'Pass', 'S_SchemaVer'
$ALWAYS = 'Date/Time', 'Student Name', 'Student No', 'Paper Stock', 'Unit', 'Qty', 'Chargeable Cost', 'Paid', 'Status', 'Notes'
$WORKING = 'H_Issues', 'H_Ink', 'H_Setup', 'H_Notices'

$script:anyFail = $false
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $main = Invoke-ComRetry { $wb.Worksheets('Example Print Room') }
    [void]$main.Activate()
    $lo = $main.ListObjects('tblJobs_EXAMPLE')
    $pws = $wb.Worksheets('Printers'); $plo = $pws.ListObjects('tblPrinters')

    function Hid([string]$h, $table = $lo) { return [bool]$table.ListColumns($h).Range.EntireColumn.Hidden }
    function AllHid($hs) { foreach ($h in $hs) { if (-not (Hid $h)) { return $false } }; return $true }
    function NoneHid($hs) { foreach ($h in $hs) { if (Hid $h) { return $false } }; return $true }
    function Ticked([string]$g) { return ([int]$main.CheckBoxes("pcb_show_$g").Value -eq 1) }
    function Choose([string]$g, [bool]$shown) { [void]$xl.Run('SetColumnGroup', $main, $g, $shown) }
    function Loc([string]$n) { return $main.Names.Item($n).RefersToRange }
    function Stored() { return [string](Loc 'LOC_Columns').Text }
    function SetLoc([string]$n, [string]$text) {
        [void]$xl.Run('UnlockSheet', $main)
        (Loc $n).Value2 = $text
        [void]$xl.Run('RelockSheet', $main)
    }
    function SetColourMode([string]$model, [string]$mode) {
        $row = 0
        for ($i = 1; $i -le $plo.ListRows.Count; $i++) { if (([string]$plo.ListRows($i).Range.Cells(1, (Col $plo 'Model')).Value2).Trim() -eq $model) { $row = $i } }
        [void]$xl.Run('UnlockSheet', $pws)
        $plo.ListRows($row).Range.Cells(1, (Col $plo 'Colour mode')).Value2 = $mode
        [void]$xl.Run('RelockSheet', $pws)
        Check ((Col $plo 'Model') -gt 0) "$model set to $(if ($mode) { $mode } else { 'single pass' })"
    }
    # Leave and come back: activation re-applies the room.
    function Revisit() {
        [void]$wb.Worksheets('Summary').Activate()
        Start-Sleep -Milliseconds 200
        [void]$main.Activate()
        Start-Sleep -Milliseconds 200
    }
    function LastJob() { return [int]$lo.ListRows.Count }
    function JobCell([string]$h, [int]$n) { return $lo.ListRows($n).Range.Cells(1, (Col $lo $h)) }
    function JobText([string]$h, [int]$n) { return [string](JobCell $h $n).Text }
    function SelectName([int]$n) { [void]$main.Activate(); [void](JobCell 'Student Name' $n).Select() }
    function ColumnTexts($table, [string]$h) { $out = @(); for ($i = 1; $i -le $table.ListRows.Count; $i++) { $out += ([string]$table.ListRows($i).Range.Cells(1, (Col $table $h)).Value2).Trim() }; return $out }
    function WbName([string]$n) { try { return $wb.Names.Item($n) } catch { return $null } }
    function JobsTableOf($ws) { foreach ($t in $ws.ListObjects) { if ($t.Name -like 'tblJobs_*') { return $t } } }

    Write-Host '=== Built state ==='
    $oldDrop = $false; foreach ($d in $main.DropDowns()) { if ($d.Name -like 'pcb_*') { $oldDrop = $true } }
    Check (-not $oldDrop) 'the old view drop-down is gone'
    $oldBtn = $false; foreach ($b in $main.Buttons()) { if ($b.Name -like 'pcb_btnToggleCost*' -or $b.Name -like 'pcb_btnToggleMultiPass*') { $oldBtn = $true } }
    Check (-not $oldBtn) 'the old toggle buttons are gone'
    Check ([string]$main.Range('D2').Text -eq '') 'row 2 no longer holds the old label'
    $setLo = $wb.Worksheets('Settings').ListObjects('tblSettings')
    $keys = ColumnTexts $setLo 'Key'
    foreach ($k in 'LOC_REDUCED_VIEW', 'LOC_REDUCED_COLUMNS', 'LOC_MINIMAL_COLUMNS', 'COST_COLS_HIDDEN', 'MULTIPASS_COLS_HIDDEN') {
        Check (($keys -notcontains $k) -and ($null -eq (WbName "SET_$k"))) "$k is retired"
    }
    Check ((Stored) -eq '') 'no clicks are stored on a fresh build'
    Check (NoneHid $ALWAYS) 'the always-visible columns show'
    Check (AllHid $INTERNAL) 'the internal columns are hidden'
    Check (AllHid $WORKING) 'the working columns stay hidden'
    foreach ($g in 'cost', 'disregard', 'size', 'tech') {
        Check ((Ticked $g) -and (NoneHid $GROUPS[$g])) "$g defaults to shown and ticked"
    }

    Write-Host ''
    Write-Host '=== Each box moves exactly its own columns ==='
    foreach ($g in $GROUPS.Keys) { Choose $g $true }
    foreach ($g in @($GROUPS.Keys)) {
        $others = @(); foreach ($k in $GROUPS.Keys) { if ($k -ne $g) { $others += $GROUPS[$k] } }
        Choose $g $false
        Check ((AllHid $GROUPS[$g]) -and -not (Ticked $g)) "$g off: its columns hide and the box clears"
        Check (NoneHid $others) "$g off: the other groups are untouched"
        Check (NoneHid $ALWAYS) "$g off: the always-visible columns stay"
        Choose $g $true
        Check ((NoneHid $GROUPS[$g]) -and (Ticked $g)) "$g on: its columns return and the box ticks"
    }

    Write-Host ''
    Write-Host '=== The strip stays clear with every group off ==='
    $everything = @(); foreach ($k in $GROUPS.Keys) { $everything += $GROUPS[$k] }
    foreach ($g in $GROUPS.Keys) { Choose $g $false }
    Check (NoneHid $ALWAYS) 'every group off: the always-visible columns stay'
    $start = [double]$main.Cells(1, 4).Left
    $lbl = $main.Shapes.Item('pcb_lblShow')
    $x = [double]$lbl.Left
    Check ([Math]::Abs($x - $start) -lt 1) "the label starts at column D ($x, column $start)"
    $prev = $x + [double]$lbl.Width
    foreach ($g in $GROUPS.Keys) {
        $s = $main.Shapes.Item("pcb_show_$g")
        Check ([double]$s.Left -ge $prev - 0.5) "$g sits after the control before it ($($s.Left) >= $prev)"
        $prev = [double]$s.Left + [double]$s.Width
    }
    $panel = [double]$main.Cells(1, 47).Left
    Check ($prev -le $panel) "the strip ends before the side panel ($prev <= $panel)"
    [void]$xl.Run('btnRefreshLocations')
    Start-Sleep -Milliseconds 500
    Check (AllHid $everything) 'Refresh Locations keeps the choices'
    [void]$xl.Run('InitialiseWorkbook')
    Start-Sleep -Seconds 1
    Check (AllHid $everything) 'a second setup keeps the choices'
    Check (-not (Ticked 'cost')) '...and the redrawn boxes show them'
    foreach ($g in $GROUPS.Keys) { Choose $g $true }

    Write-Host ''
    Write-Host '=== Choices are per room ==='
    $room = $xl.Run('CreatePrintRoom', 'Toggle Room', 'Facilities')
    [void]$main.Activate()
    $rtbl = JobsTableOf $room
    Choose 'cost' $false
    Check (AllHid $GROUPS['cost']) 'cost off on this room'
    Check (-not (Hid 'Paper Cost' $rtbl)) 'the other room keeps its cost columns'
    $room2 = $xl.Run('CreatePrintRoom', 'Toggle Room Two', 'Facilities')
    [void]$main.Activate()
    $rtbl2 = JobsTableOf $room2
    Check (-not (Hid 'Paper Cost' $rtbl2)) 'a new room starts from the defaults'
    Check ([string]$room2.Names.Item('LOC_Columns').RefersToRange.Text -eq '') 'a new room has no stored clicks'
    Choose 'cost' $true

    Write-Host ''
    Write-Host '=== The Printer rule ==='
    [void](Loc 'LOC_Columns').ClearContents()
    $allowed = [string](Loc 'LOC_Printers').Text
    [void]$xl.Run('ApplyRoomColumns', $main)
    Check ((-not (Hid 'Printer')) -and (Ticked 'printer')) 'several printers: Printer shows'
    SetLoc 'LOC_Printers' $EPSON
    Revisit
    Check ((Hid 'Printer') -and -not (Ticked 'printer')) 'one printer: Printer hides'
    Choose 'printer' $true
    Check (-not (Hid 'Printer')) 'a click shows it'
    Revisit
    Check (-not (Hid 'Printer')) '...and the click sticks'
    SetLoc 'LOC_Printers' $allowed
    Revisit
    Check (-not (Hid 'Printer')) 'back to several: shown'
    Check ((Stored) -notlike '*printer*') "the rule's answer changed, so the click is dropped"
    Choose 'printer' $false
    Check (Hid 'Printer') 'several printers, click off: hidden'
    SetLoc 'LOC_Printers' $EPSON
    Revisit
    Check ((Hid 'Printer') -and ((Stored) -notlike '*printer*')) 'one printer: the rule agrees, and the old click is dropped'
    SetLoc 'LOC_Printers' $allowed
    Revisit
    Check (-not (Hid 'Printer')) 'several again: the rule shows it'

    Write-Host ''
    Write-Host '=== The Colour passes rule ==='
    Check ((ColumnTexts $plo 'Colour mode') -notcontains 'multi-pass') 'the built workbook has no multi-pass printer'
    Check ((AllHid $GROUPS['passes']) -and -not (Ticked 'passes')) 'no multi-pass printer here: the pass columns hide'
    SetColourMode $EPSON 'multi-pass'
    Revisit
    Check ((NoneHid $GROUPS['passes']) -and (Ticked 'passes')) 'a permitted multi-pass printer: they show'
    Choose 'passes' $false
    Revisit
    Check (AllHid $GROUPS['passes']) 'a click hides them, and sticks'
    SetColourMode $EPSON ''
    Revisit
    SetColourMode $EPSON 'multi-pass'
    Revisit
    Check (NoneHid $GROUPS['passes']) "the rule's answer changed and back: the click was dropped, so they show"

    Write-Host ''
    Write-Host '=== Showing a column when it is needed ==='
    SetLoc 'LOC_Printers' $EPSON
    Revisit
    SetLoc 'LOC_DefTech' ''
    Choose 'tech' $false
    [void]$main.Activate()
    [void]$xl.Run('btnAddPrintJob')
    Start-Sleep -Milliseconds 300
    $job = LastJob
    Check ((JobText 'Printer' $job) -eq $EPSON) "one permitted printer: New job fills it in (got '$(JobText 'Printer' $job)')"
    Check (Hid 'Printer') '...so the Printer column stays hidden'
    Check ((-not (Hid 'Technician')) -and (Ticked 'tech')) 'no default technician: New job brings Technician back and ticks it'
    $tech = (ColumnTexts $wb.Worksheets('Print Technicians').ListObjects('tblTechnicians') 'Name')[0]
    SetLoc 'LOC_DefTech' $tech
    Choose 'tech' $false
    [void]$xl.Run('btnAddPrintJob')
    Start-Sleep -Milliseconds 300
    Check ((JobText 'Technician' (LastJob)) -eq $tech) 'a default technician fills the new job'
    Check (Hid 'Technician') '...and Technician stays hidden'
    SetLoc 'LOC_DefTech' ''
    SelectName (LastJob)
    Choose 'tech' $false
    [void]$xl.Run('btnRepeatJob')
    Start-Sleep -Milliseconds 300
    Check ((JobText 'Technician' (LastJob)) -eq '') 'Repeat with no default technician leaves it blank'
    Check ((-not (Hid 'Technician')) -and (Ticked 'tech')) '...and brings Technician back and ticks it'

    # The printer is multi-pass here, which takes sheet stock only, so a supplied roll would be refused.
    SetColourMode $EPSON ''
    Choose 'size' $false
    (JobCell 'Paper Stock' (LastJob)).Value2 = 'Supplied (Roll)'
    Start-Sleep -Milliseconds 300
    Check ((JobText 'Paper Stock' (LastJob)) -eq 'Supplied (Roll)') 'the supplied stock was accepted'
    Check ((NoneHid $GROUPS['size']) -and (Ticked 'size')) 'student-supplied stock brings Print size back'
    SetColourMode $EPSON 'multi-pass'

    $job = LastJob
    SelectName $job
    Choose 'passes' $false
    [void]$xl.Run('btnAddPass')
    Start-Sleep -Milliseconds 300
    Check ((LastJob) -eq $job + 1) 'Add pass added a row'
    Check ((NoneHid $GROUPS['passes']) -and (Ticked 'passes')) 'Add pass brings Colour passes back'
    Check (Hid 'Pass') 'the Pass number stays with the internal columns'
    [void]$xl.Run('ApplyRoomColumns', $main)
    SetColourMode $EPSON ''
    Revisit
    [void](Loc 'LOC_Columns').ClearContents()
    Revisit
    Check (NoneHid $GROUPS['passes']) 'pass rows but no multi-pass printer: the pass columns still show'

    Write-Host ''
    Write-Host '=== The internal columns setting ==='
    $setting = WbName 'SET_SHOW_INTERNAL_COLS'
    Check (($null -ne $setting) -and ([string]$setting.RefersToRange.Text -eq 'No')) 'the setting exists and starts at No'
    [void]$xl.Run('UnlockSheet', $wb.Worksheets('Settings'))
    $setting.RefersToRange.Value2 = 'Yes'
    Start-Sleep -Milliseconds 500
    Check (NoneHid $INTERNAL) 'Yes: the internal columns show here'
    Check (-not (Hid 'Job ID' $rtbl)) '...and on the other rooms'
    Check (AllHid $WORKING) '...and the working columns stay hidden'
    $setting.RefersToRange.Value2 = 'No'
    Start-Sleep -Milliseconds 500
    Check (AllHid $INTERNAL) 'No: they hide again'
    Check (Hid 'Job ID' $rtbl) '...on the other rooms too'

    $xl.Run('SetQuiet', $false)
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($script:anyFail) { Write-Host 'FAIL'; exit 1 }
Write-Host 'column toggles checks complete'
