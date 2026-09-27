# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Phase 7: the validation sweep.
#
#   AT-11  student name/number conflict WARNS and never blocks
#   AT-14  Clear All names the location, the count and the date range
#   AT-16  inactive records leave the dropdowns but stay in the reports
#
# Quiet mode does the heavy lifting for AT-14: Ask() returns False and logs
# its prompt, so the confirmation text can be read without deleting anything.
# The test is non-destructive by construction rather than by care.
#
# Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself.
#
# Neither Close($false) nor AutoSaveOn is the protection it looks like.
# Workbook_Open does real work on every open - ProtectAll unprotects and
# reprotects all fourteen sheets, plus Invalidate and HealButtons - so the
# workbook is dirty the instant it opens. Excel holds this file through its
# cloud-backed handle (Workbook.Path returns a d.docs.live.net URL), so
# AutoSave is on and COMMITS that immediately: $wb.Saved reads True on the
# very next line after Open, despite every sheet having just been modified.
# The bytes land when the handle closes - measured, the timestamp and the
# SHA-256 both move at Close/Quit, not at Open.
#
# Which is exactly why Close($false) does not help. It means "discard unsaved
# changes", and by then there are none: AutoSave has accepted them, so what
# is written at close is committed state. Setting AutoSaveOn = $false after
# Open is too late for the same reason, and Open offers no earlier hook - so
# no guard is attempted here, because none can work. (build.ps1 sets it, but
# only ever on a %TEMP% copy, which is not cloud-backed; its own comment notes
# it is usually a no-op there.) Tested: AutoSaveOn False, Saved forced True,
# nothing touched at all - the file still changes.
#
# None of this is the sync client. It reproduces with syncing paused and all
# sync operations settled. It is Excel-side only.
#
# verify.ps1 and probe.ps1 take the other way out: read-only, so Workbook_Open
# cannot save anything. These scripts need read-write to drive mutating
# macros, so a throwaway copy is what is left. Discarded at the end.
#
# Left unfixed, this rewrites the artefact under test: the phase 5-7 runs did,
# and a later script read an earlier one's edits back as though they were the
# build.
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
    $main = $wb.Worksheets('Main Print Room')
    $lo = $main.ListObjects('tblJobs_MAIN')
    $before = $lo.ListRows.Count

    function Col($lo, $name) {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            if ($lo.ListColumns($i).Name -eq $name) { return $i }
        }
        return 0
    }

    # ---------------------------------------------------------------- AT-14
    Write-Host '=== AT-14: Clear All confirmation ==='
    $xl.Run('SetQuiet', $true)
    $main.Activate()
    $xl.Run('btnClearAll')
    $log = [string]$xl.Run('QuietLog')
    Write-Host $log
    $after = $lo.ListRows.Count
    Write-Host ("  rows before={0} after={1}  {2}" -f $before, $after,
        $(if ($before -eq $after) { 'NOT DELETED - Ask returned False as designed' } else { 'DELETED - BUG' }))
    foreach ($need in 'Main Print Room', '5 print job', '15/09/2026', '17/09/2026', 'cannot be undone') {
        Write-Host ("  contains '{0,-18}' {1}" -f $need, $(if ($log -like "*$need*") { 'yes' } else { 'NO' }))
    }

    # ---------------------------------------------------------------- AT-11
    Write-Host ''
    Write-Host '=== AT-11: student name/number conflict ==='
    $xl.Run('SetQuiet', $true)
    $main.Unprotect()
    $nameCol = Col $lo 'Student Name'
    $noCol   = Col $lo 'Student No'
    Write-Host ("  row1: {0} / {1}" -f $lo.ListRows(1).Range.Cells(1, $nameCol).Text, $lo.ListRows(1).Range.Cells(1, $noCol).Text)
    Write-Host ("  row2: {0} / {1}" -f $lo.ListRows(2).Range.Cells(1, $nameCol).Text, $lo.ListRows(2).Range.Cells(1, $noCol).Text)

    $cell = $lo.ListRows(2).Range.Cells(1, $nameCol)
    $wasName = [string]$cell.Text
    $cell.Value2 = 'Jane Smyth'
    Start-Sleep -Milliseconds 300
    $log = [string]$xl.Run('QuietLog')
    Write-Host '  --- warning raised ---'
    Write-Host $log
    Write-Host ("  value kept as typed: '{0}'  {1}" -f $cell.Text,
        $(if ($cell.Text -eq 'Jane Smyth') { 'WARNED not blocked - correct' } else { 'BLOCKED - BUG' }))

    # ---------------------------------------------------------------- AT-16
    Write-Host ''
    Write-Host '=== AT-16: inactive records ==='
    $tech = $wb.Worksheets('Print Technicians')
    $tech.Unprotect()
    $tlo = $tech.ListObjects('tblTechnicians')
    $nCol = Col $tlo 'Name'
    $aCol = Col $tlo 'Active'
    $victim = [string]$tlo.ListRows(1).Range.Cells(1, $nCol).Text
    Write-Host ("  deactivating technician: {0}" -f $victim)
    $tlo.ListRows(1).Range.Cells(1, $aCol).Value2 = 'No'
    $xl.Run('Invalidate')
    $xl.Run('BindColumns', $main)

    # The dropdown lists live in staging columns on _Work, tagged in row 1.
    $work = $wb.Worksheets('_Work')
    $listed = @()
    for ($c = 1; $c -le 200; $c++) {
        if ([string]$work.Cells(1, $c).Text -eq 'TEC') {
            for ($r = 2; $r -le 60; $r++) {
                $v = [string]$work.Cells($r, $c).Text
                if ($v -eq '') { break }
                $listed += $v
            }
            break
        }
    }
    Write-Host ("  technician dropdown now: {0}" -f ($listed -join ', '))
    Write-Host ("  '{0}' excluded from dropdown: {1}" -f $victim,
        $(if ($listed -notcontains $victim) { 'yes - correct' } else { 'NO - BUG' }))

    # ...but their historical rows must remain in the reports.
    $data = $wb.Worksheets('_Data')
    $xl.CalculateFullRebuild()
    $spill = $data.Range('A10').SpillingToRange
    # Technician's column in _Data - found by header (row 9), not a hardcoded
    # index: 2026-09-22's ReorderJobColumns moved Status/Job ID to the back
    # of the job table, which shifted every column after them, _Data
    # included.
    $techCol = 0
    for ($c = 1; $c -le 30; $c++) {
        if ([string]$data.Cells(9, $c).Text -eq 'Technician') { $techCol = $c; break }
    }
    $stillThere = 0
    for ($r = 1; $r -le $spill.Rows.Count; $r++) {
        if ([string]$spill.Cells($r, $techCol).Text -eq $victim) { $stillThere++ }
    }
    Write-Host ("  rows in the consolidated range still crediting '{0}': {1}  {2}" -f $victim, $stillThere,
        $(if ($stillThere -gt 0) { 'retained - correct' } else { 'LOST - BUG' }))

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
