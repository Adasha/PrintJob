# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Prerequisite: site-prefixed catalogue IDs (modCatalog "catalogue IDs").
#
# Technician / printer / paper IDs (TechID, PrinterID, StockID) are allocated
# like Job IDs - SITE-CODE-00001 from a persisted high-water mark - so that
# combining another workbook's configuration adds rows instead of overwriting.
#
#   1. Every named row already has a unique, well-formed ID (Setup back-fill).
#   2. The Add row path allocates the next ID, and never reissues a deleted one.
#   3. A row typed under the table gets its ID once it has a name.
#   4. Restoring another site's catalogue (same names, different site prefix)
#      ADDS rows, renaming any whose name is already taken to "Name (SITE)";
#      it overwrites nothing, is idempotent, and does not wind the counters back.
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
function Check([bool]$cond, [string]$msg) {
    Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    if (-not $cond) { $script:anyFail = $true }
}

$specs = @(
    @{ Sheet = 'Print Technicians'; Table = 'tblTechnicians'; Code = 'TCH'; Id = 'TechID';    Name = 'Name';        Hwm = 'TECH_ID_HWM' },
    @{ Sheet = 'Printers';          Table = 'tblPrinters';    Code = 'PRN'; Id = 'PrinterID'; Name = 'Model';       Hwm = 'PRINTER_ID_HWM' },
    @{ Sheet = 'Papers';            Table = 'tblPapers';      Code = 'STK'; Id = 'StockID';   Name = 'Description'; Hwm = 'STOCK_ID_HWM' }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)
    $site = [string]$xl.Run('SettingText', 'SITE_ID', 'SITE')
    Write-Host "site code: $site"

    function ColIdx($lo, [string]$name) {
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) { if ($lo.ListColumns($i).Name -eq $name) { return $i } }
        throw "column '$name' not found in $($lo.Name)"
    }
    # Named rows only: a table's blank templated row is not a record.
    function Get-Named($s) {
        $lo = $wb.Worksheets($s.Sheet).ListObjects($s.Table)
        $idC = ColIdx $lo $s.Id; $nmC = ColIdx $lo $s.Name
        $out = @()
        for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
            $nm = [string]$lo.ListRows($i).Range.Cells(1, $nmC).Value2
            if ($nm.Trim().Length -gt 0) {
                $out += [pscustomobject]@{ Id = [string]$lo.ListRows($i).Range.Cells(1, $idC).Value2; Name = $nm }
            }
        }
        return , $out
    }
    function Num([string]$id) { return [int](($id -split '-')[-1]) }
    function Remove-RowsWithId($s, [string[]]$ids) {
        $ws = $wb.Worksheets($s.Sheet); $lo = $ws.ListObjects($s.Table); $idC = ColIdx $lo $s.Id
        [void]$xl.Run('UnlockSheet', $ws)
        for ($i = $lo.ListRows.Count; $i -ge 1; $i--) {
            if ($ids -contains [string]$lo.ListRows($i).Range.Cells(1, $idC).Value2) { $lo.ListRows($i).Delete() }
        }
        [void]$xl.Run('RelockSheet', $ws)
    }

    $tag = [regex]::Escape($site)

    foreach ($s in $specs) {
        Write-Host ''
        Write-Host ("=== {0} ===" -f $s.Table)
        $ws = $wb.Worksheets($s.Sheet); $lo = $ws.ListObjects($s.Table)
        $idC = ColIdx $lo $s.Id; $nmC = ColIdx $lo $s.Name
        $pattern = "^$tag-$($s.Code)-\d{5}$"

        $rows = Get-Named $s
        Check ($rows.Count -gt 0) "starts with named rows (got $($rows.Count))"
        Check (@($rows | Where-Object { [string]::IsNullOrWhiteSpace($_.Id) }).Count -eq 0) 'every named row has an ID'
        Check (@($rows | Where-Object { $_.Id -notmatch $pattern }).Count -eq 0) "every ID matches $site-$($s.Code)-nnnnn"
        Check ((@($rows | Select-Object -ExpandProperty Id -Unique)).Count -eq $rows.Count) 'IDs are unique'
        $existing = @($rows | Select-Object -ExpandProperty Id)

        # Add row: next ID, then delete it and add again - never reissued.
        [void]$ws.Activate()
        [void]$xl.Run('AddCatalogRow', $s.Table)
        $n = $lo.ListRows.Count
        $id1 = [string]$lo.ListRows($n).Range.Cells(1, $idC).Value2
        Check ($id1 -match $pattern) "Add row allocates a well-formed ID ($id1)"
        Check ($existing -notcontains $id1) 'the new ID is not one already in use'

        [void]$xl.Run('UnlockSheet', $ws)
        $lo.ListRows($n).Delete()
        [void]$xl.Run('RelockSheet', $ws)
        [void]$xl.Run('AddCatalogRow', $s.Table)
        $n = $lo.ListRows.Count
        $id2 = [string]$lo.ListRows($n).Range.Cells(1, $idC).Value2
        Check ($id2 -ne $id1 -and (Num $id2) -gt (Num $id1)) "deleting the newest row does not let its ID be reissued ($id1 deleted, next is $id2)"
        $hwm = [int][double]$xl.Run('SettingNum', $s.Hwm, 0)
        Check ($hwm -eq (Num $id2)) "counter $($s.Hwm) = $hwm matches the last ID issued"

        # A row typed under the table: no ID until it has a name.
        [void]$xl.Run('UnlockSheet', $ws)
        [void]$lo.ListRows.Add()
        [void]$xl.Run('RelockSheet', $ws)
        $tn = $lo.ListRows.Count
        $lo.ListRows($tn).Range.Cells(1, $nmC).Value2 = "Typed $($s.Code) test"
        Start-Sleep -Milliseconds 300
        $id3 = [string]$lo.ListRows($tn).Range.Cells(1, $idC).Value2
        Check ($id3 -match $pattern) "a typed row gets an ID once named ($id3)"
        Check ((Num $id3) -eq ((Num $id2) + 1)) 'and it is the next number in sequence'

        Remove-RowsWithId $s @($id2, $id3)
    }

    Write-Host ''
    Write-Host '=== Combining another site''s catalogue ==='
    $before = @{}; $hwmBefore = @{}
    foreach ($s in $specs) { $before[$s.Table] = Get-Named $s }

    [void]$xl.Run('BackupAll')
    $csvs = Get-ChildItem $workDir -Filter 'PrintCosts-*-CATALOG-*.csv'
    Check ($csvs.Count -eq 7) "backup wrote 7 catalogue CSVs (got $($csvs.Count))"

    # Turn the backup into "another site's": same names, different site prefix.
    $latin1 = [Text.Encoding]::GetEncoding(28591)
    foreach ($s in $specs) {
        $csv = $csvs | Where-Object { $_.Name -like "*-CATALOG-$($s.Table)-*" } | Select-Object -First 1
        $text = $latin1.GetString([IO.File]::ReadAllBytes($csv.FullName))
        $text = $text.Replace("$site-$($s.Code)-", "OTHER-$($s.Code)-")
        [IO.File]::WriteAllBytes($csv.FullName, $latin1.GetBytes($text))
    }
    $anyCsv = $csvs | Where-Object { $_.Name -like '*-CATALOG-tblPrinters-*' } | Select-Object -First 1

    # Advance one counter after the backup: a restore must not wind it back.
    $techSpec = $specs[0]
    [void]$wb.Worksheets($techSpec.Sheet).Activate()
    [void]$xl.Run('AddCatalogRow', $techSpec.Table)
    $tlo = $wb.Worksheets($techSpec.Sheet).ListObjects($techSpec.Table)
    $bumpId = [string]$tlo.ListRows($tlo.ListRows.Count).Range.Cells(1, (ColIdx $tlo $techSpec.Id)).Value2
    Remove-RowsWithId $techSpec @($bumpId)
    $techHwmBefore = [int][double]$xl.Run('SettingNum', $techSpec.Hwm, 0)
    Check ($techHwmBefore -eq (Num $bumpId)) "counter advanced past the backup ($bumpId)"

    [void]$xl.Run('RestoreFromFileConfirmed', $anyCsv.FullName)
    Start-Sleep -Milliseconds 300

    foreach ($s in $specs) {
        $after = Get-Named $s
        $b = $before[$s.Table]
        Write-Host ("  -- {0}: {1} rows before, {2} after" -f $s.Table, $b.Count, $after.Count)
        Check ($after.Count -eq 2 * $b.Count) "the other site's rows were ADDED (expected $(2 * $b.Count))"
        $missing = @($b | Where-Object { $o = $_; -not @($after | Where-Object { $_.Id -eq $o.Id -and $_.Name -eq $o.Name }).Count })
        Check ($missing.Count -eq 0) 'every original row is untouched (same ID, same name)'
        $unrenamed = @($b | Where-Object { $o = $_; -not @($after | Where-Object { $_.Id -like 'OTHER-*' -and $_.Name -eq "$($o.Name) (OTHER)" }).Count })
        Check ($unrenamed.Count -eq 0) 'each clashing name was kept under a "Name (OTHER)" suffix'
        Check ((@($after | Select-Object -ExpandProperty Name | ForEach-Object { $_.ToLower() } | Select-Object -Unique)).Count -eq $after.Count) 'no two rows share a name'
    }
    $techHwmAfter = [int][double]$xl.Run('SettingNum', $techSpec.Hwm, 0)
    Check ($techHwmAfter -eq $techHwmBefore) "restore did not wind the counter back ($techHwmBefore -> $techHwmAfter)"

    Write-Host ''
    Write-Host '=== Restoring the same backup again is idempotent ==='
    [void]$xl.Run('RestoreFromFileConfirmed', $anyCsv.FullName)
    Start-Sleep -Milliseconds 300
    foreach ($s in $specs) {
        $again = Get-Named $s
        Check ($again.Count -eq 2 * $before[$s.Table].Count) "$($s.Table) still $(2 * $before[$s.Table].Count) rows"
    }

    Write-Host ''
    Write-Host '=== Next ID after the merge does not collide ==='
    [void]$wb.Worksheets($techSpec.Sheet).Activate()
    [void]$xl.Run('AddCatalogRow', $techSpec.Table)
    $tlo = $wb.Worksheets($techSpec.Sheet).ListObjects($techSpec.Table)
    $idNew = [string]$tlo.ListRows($tlo.ListRows.Count).Range.Cells(1, (ColIdx $tlo $techSpec.Id)).Value2
    $named = Get-Named $techSpec
    $all = @($named | Select-Object -ExpandProperty Id)
    Check ($idNew -match "^$tag-TCH-\d{5}$" -and $all -notcontains $idNew) "new technician gets a fresh local-site ID ($idNew)"

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
