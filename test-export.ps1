# Phase 6: per-location export, the derived fingerprint, and the status line.
#
# Exports a print room, reads the CSV back, then edits a row and confirms the
# fingerprint notices. Closes the workbook WITHOUT saving; the CSV it writes
# is left for inspection and removed at the end.

$ErrorActionPreference = 'Stop'
# Drives a COPY in %TEMP%, never src\PrintCosts.xlsm itself.
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
#
# $src is the copy's folder, so the export lands there and is read back
# from there. It also means this script no longer deletes the real
# exports in src\ to find its own - which it used to do on every run.
$deliverable = Join-Path $PSScriptRoot 'src\PrintCosts.xlsm'
$src = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $src | Out-Null
$f = Join-Path $src 'PrintCosts.xlsm'
Copy-Item $deliverable $f
$workDir = $src

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $ws = $wb.Worksheets('Main Print Room')
    $xl.Run('SetQuiet', $true)

    Write-Host '=== status before export ==='
    Write-Host ('  B8: ' + $ws.Range('B8').Text)

    $ws.Activate()
    $xl.Run('btnExport')
    Write-Host ''
    Write-Host '=== what the export reported ==='
    Write-Host $xl.Run('QuietLog')

    Write-Host '=== status after export ==='
    Write-Host ('  B8: ' + $ws.Range('B8').Text)

    # Edit a quantity. The fingerprint must notice: no event fires for this
    # when macros are off, which is the whole reason it is derived.
    $lo = $ws.ListObjects('tblJobs_MAIN')
    $qtyCol = 0
    for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
        if ($lo.ListColumns($i).Name -eq 'Quantity') { $qtyCol = $i }
    }
    $cell = $lo.ListRows(1).Range.Cells(1, $qtyCol)
    $was = $cell.Value2
    $ws.Unprotect('printlog')
    $cell.Value2 = [double]$was + 7
    $xl.CalculateFullRebuild()
    $xl.Run('RefreshExportStatus', $ws)
    Write-Host ''
    Write-Host ("=== after changing a quantity ({0} -> {1}) ===" -f $was, $cell.Value2)
    Write-Host ('  B8: ' + $ws.Range('B8').Text)

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Write-Host ''
Write-Host '=== the CSV ==='
$csv = Get-ChildItem $src -Filter 'PrintCosts-*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $csv) { Write-Host '  NO CSV WRITTEN'; exit 1 }
Write-Host ("  {0}  ({1} bytes)" -f $csv.Name, $csv.Length)
Write-Host ''
$lines = Get-Content $csv.FullName
Write-Host '  --- header block ---'
$lines[0..7] | ForEach-Object { Write-Host ('  ' + $_) }
Write-Host '  --- column headers ---'
Write-Host ('  ' + $lines[8])
Write-Host '  --- first record ---'
Write-Host ('  ' + $lines[9])
Write-Host ''
Write-Host ("  total lines: {0}" -f $lines.Count)
$hdrs = $lines[8] -split ','
Write-Host ("  columns: {0}" -f $hdrs.Count)
foreach ($need in 'S_UnitCost', 'S_ConsRate', 'S_StockWidth_mm', 'S_SheetHeight_mm', 'S_SchemaVer') {
    $present = $hdrs -contains $need
    Write-Host ("  snapshot column {0,-18} {1}" -f $need, $(if ($present) { 'present' } else { 'MISSING' }))
}

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
