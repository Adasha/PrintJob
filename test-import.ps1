# Problem 1: Export All Locations, and Import (restore into origin, and
# import into a different room).
#
# PickImportFile/PickTargetLocation open real OS dialogs and cannot run
# unattended, so this drives the pipeline beneath them directly:
# modExport.ExportOne (via btnExport, quiet-mode safe) to produce a file,
# then modImport.ReadImportRows + ApplyImportConfirmed to read it back in -
# the same split test-nextid.ps1 uses to reach past modJobs.RemoveRow's
# Ask() gate.
#
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself - see verify.ps1.

$ErrorActionPreference = 'Stop'
$deliverable = Join-Path $PSScriptRoot 'src\PrintCosts.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintCosts.xlsm'
Copy-Item $deliverable $f

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $xl.Run('SetQuiet', $true)
    $main = $wb.Worksheets('Main Print Room')
    $annexe = $wb.Worksheets('Annexe')
    $loMain = $main.ListObjects('tblJobs_MAIN')

    function Col($lo, $name) {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            if ($lo.ListColumns($i).Name -eq $name) { return $i }
        }
        throw "column '$name' not found"
    }
    $idCol = Col $loMain 'Job ID'
    $chgCol = Col $loMain 'Chargeable Cost'
    $unitCostCol = Col $loMain 'S_UnitCost'

    # --- Export All Locations -------------------------------------------
    Write-Host '=== Export All Locations ==='
    $xl.Run('btnExportAll')
    Write-Host $xl.Run('QuietLog')
    $mainCsv = Get-ChildItem $workDir -Filter 'PrintCosts-*-MAIN-*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $mainCsv) { Write-Host 'FAIL: no CSV produced for Main Print Room'; exit 1 }
    Write-Host ("  {0}" -f $mainCsv.Name)

    # --- capture the row we are about to delete, then delete it ----------
    $victimRow = 1
    $victimId = [string]$loMain.ListRows($victimRow).Range.Cells(1, $idCol).Value2
    $victimChg = [double]$loMain.ListRows($victimRow).Range.Cells(1, $chgCol).Value2
    $victimUnitCost = [double]$loMain.ListRows($victimRow).Range.Cells(1, $unitCostCol).Value2
    Write-Host ''
    Write-Host ("victim row: {0}  chargeable={1}  S_UnitCost={2}" -f $victimId, $victimChg, $victimUnitCost)

    $main.Unprotect()
    $loMain.ListRows($victimRow).Delete()
    $main.Protect()
    $stillThere = $false
    for ($i = 1; $i -le $loMain.ListRows.Count; $i++) {
        if ([string]$loMain.ListRows($i).Range.Cells(1, $idCol).Value2 -eq $victimId) { $stillThere = $true }
    }
    if ($stillThere) { throw 'victim row was not actually deleted' }
    Write-Host 'deleted from Main Print Room'

    # --- restore it via import into the SAME room -------------------------
    Write-Host ''
    Write-Host '=== Import: restore into origin ==='
    $rows = $xl.Run('ReadImportRows', $mainCsv.FullName)
    $xl.Run('ApplyImportConfirmed', $main, $rows)
    Write-Host $xl.Run('QuietLog')

    $foundRow = 0
    for ($i = 1; $i -le $loMain.ListRows.Count; $i++) {
        if ([string]$loMain.ListRows($i).Range.Cells(1, $idCol).Value2 -eq $victimId) { $foundRow = $i }
    }
    if ($foundRow -eq 0) { Write-Host "FAIL: $victimId was not restored"; exit 1 }
    $xl.CalculateFullRebuild()
    $restoredChg = [double]$loMain.ListRows($foundRow).Range.Cells(1, $chgCol).Value2
    $restoredUnitCost = [double]$loMain.ListRows($foundRow).Range.Cells(1, $unitCostCol).Value2
    Write-Host ("restored row {0}: {1}  chargeable={2}  S_UnitCost={3}" -f $foundRow, $victimId, $restoredChg, $restoredUnitCost)

    if ([Math]::Abs($restoredChg - $victimChg) -gt 0.001) {
        Write-Host "FAIL: chargeable cost changed on restore ($victimChg -> $restoredChg) - it was recosted instead of using the snapshot"
        exit 1
    }
    if ([Math]::Abs($restoredUnitCost - $victimUnitCost) -gt 0.001) {
        Write-Host "FAIL: S_UnitCost changed on restore ($victimUnitCost -> $restoredUnitCost)"
        exit 1
    }
    Write-Host 'OK: restored row matches the original snapshot exactly (not recosted)'

    # --- import the SAME export into a DIFFERENT room ----------------------
    Write-Host ''
    Write-Host '=== Import: into a different room ==='
    $loAnnexe = $annexe.ListObjects('tblJobs_ANNEX')
    $beforeAnnexeCount = $loAnnexe.ListRows.Count
    $rows2 = $xl.Run('ReadImportRows', $mainCsv.FullName)
    $xl.Run('ApplyImportConfirmed', $annexe, $rows2)
    Write-Host $xl.Run('QuietLog')

    Write-Host ("DIAG: annexe.Name={0} loAnnexe.Name={1} ListColumns.Count={2} ListRows.Count={3}" -f `
        $annexe.Name, $loAnnexe.Name, $loAnnexe.ListColumns.Count, $loAnnexe.ListRows.Count)
    for ($i = 1; $i -le $loAnnexe.ListColumns.Count; $i++) {
        Write-Host ("  col {0}: '{1}'" -f $i, $loAnnexe.ListColumns($i).Name)
    }
    $idColAnnexe = Col $loAnnexe 'Job ID'
    $movedIn = $false
    for ($i = 1; $i -le $loAnnexe.ListRows.Count; $i++) {
        if ([string]$loAnnexe.ListRows($i).Range.Cells(1, $idColAnnexe).Value2 -eq $victimId) { $movedIn = $true }
    }
    if (-not $movedIn) { Write-Host "FAIL: $victimId was not imported into Annexe Print Room"; exit 1 }
    if ($victimId -notmatch '-MAIN-') { Write-Host "FAIL: test assumption broken, victim ID does not carry a MAIN prefix"; exit 1 }
    Write-Host ("OK: '{0}' landed in Annexe's table with its ORIGINAL Job ID (prefix untouched)" -f $victimId)

    # --- confirm it now reports under Annexe's location in _Data ----------
    $xl.Run('btnRefreshLocations')
    $xl.CalculateFullRebuild()
    $d = $wb.Worksheets('_Data')
    $spill = $d.Range('A10').SpillingToRange
    # Job ID's column in _Data - found by header (row 9), not a hardcoded
    # index: 2026-09-22's ReorderJobColumns moved Status/Job ID to the back
    # of the job table (fixing a reduced-clutter-view bug), which shifted
    # every column after them, _Data included.
    $jobIdCol = 0
    for ($c = 1; $c -le 30; $c++) {
        if ([string]$d.Cells(9, $c).Text -eq 'Job ID') { $jobIdCol = $c; break }
    }
    $foundLoc = $null
    for ($r = 1; $r -le $spill.Rows.Count; $r++) {
        if ([string]$spill.Cells($r, $jobIdCol).Value2 -eq $victimId) { $foundLoc = [string]$spill.Cells($r, 1).Value2 }
    }
    Write-Host ("_Data reports {0} under Location = '{1}'" -f $victimId, $foundLoc)
    if ($foundLoc -ne 'ANNEX') {
        Write-Host "FAIL: expected Location 'ANNEX', got '$foundLoc'"
        exit 1
    }
    Write-Host "OK: consolidated range reports the imported job under its NEW sheet's location"

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
