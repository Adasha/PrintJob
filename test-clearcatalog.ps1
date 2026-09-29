# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Regression test: "Clear table" on Print Technicians, Printers and Papers
# (modCatalog.ClearCatalogTable / DoClearCatalogTable).
#
# Checks, for each of the three sheets:
#   - the button exists on that sheet
#   - the confirmed clear empties THAT table only (the other two are untouched)
#   - one blank templated row is left, RowCount reports 0, and AddCatalogRow
#     can reuse it (Active defaults to Yes)
# Plus: the prompting entry point declines in quiet mode and deletes nothing
# (Ask always declines when quiet), and clearing an already-empty table is a
# harmless no-op.
#
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT
# saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
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
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    $xl.Run('SetQuiet', $true)

    function Check([bool]$cond, [string]$msg) {
        Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    }
    function Rows([string]$sheet, [string]$table) {
        return [int]$xl.Run('RowCount', $wb.Worksheets($sheet).ListObjects($table))
    }
    function Has-Button([string]$sheet, [string]$macro) {
        foreach ($b in $wb.Worksheets($sheet).Buttons()) {
            if ([string]$b.OnAction -like "*$macro") { return $true }
        }
        return $false
    }

    $cases = @(
        @{ Sheet = 'Print Technicians'; Table = 'tblTechnicians'; Macro = 'btnClearTechnicians' },
        @{ Sheet = 'Printers';          Table = 'tblPrinters';    Macro = 'btnClearPrinters' },
        @{ Sheet = 'Papers';            Table = 'tblPapers';      Macro = 'btnClearPapers' }
    )

    foreach ($c in $cases) {
        Write-Host ("=== {0} ===" -f $c.Sheet)
        Check (Has-Button $c.Sheet $c.Macro) ("'Clear table' button ($($c.Macro)) is on the sheet")
    }

    foreach ($c in $cases) {
        Write-Host ''
        Write-Host ("=== Clear {0} only ===" -f $c.Table)
        $before = @{}
        foreach ($o in $cases) { $before[$o.Table] = Rows $o.Sheet $o.Table }
        Check ($before[$c.Table] -gt 0) ("$($c.Table) starts with rows (got $($before[$c.Table]))")

        # Quiet mode: Ask declines, so the prompting path must delete nothing.
        [void]$xl.Run('ClearCatalogTable', $c.Table)
        Check ((Rows $c.Sheet $c.Table) -eq $before[$c.Table]) 'declined confirmation deletes nothing'

        [void]$xl.Run('DoClearCatalogTable', $c.Table)
        $lo = $wb.Worksheets($c.Sheet).ListObjects($c.Table)
        Check ((Rows $c.Sheet $c.Table) -eq 0) 'table reports zero rows after clear'
        Check ($lo.ListRows.Count -eq 1) 'one blank templated row is left'
        foreach ($o in $cases) {
            if ($o.Table -ne $c.Table) {
                Check ((Rows $o.Sheet $o.Table) -eq $before[$o.Table]) ("$($o.Table) untouched ($($before[$o.Table]) rows)")
            }
        }

        # An empty table clears as a no-op.
        [void]$xl.Run('DoClearCatalogTable', $c.Table)
        Check ($lo.ListRows.Count -eq 1) 'clearing an empty table is a no-op'

        # The blank row is reusable and gets its default Active.
        [void]$wb.Worksheets($c.Sheet).Activate()
        [void]$xl.Run('AddCatalogRow', $c.Table)
        Check ($lo.ListRows.Count -eq 1) 'AddCatalogRow reuses the blank row rather than adding a second'
        $hasActive = $false
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) { if ($lo.ListColumns($i).Name -eq 'Active') { $hasActive = $i } }
        if ($hasActive) {
            Check ([string]$lo.ListRows(1).Range.Cells(1, $hasActive).Value2 -eq 'Yes') 'Active defaults to Yes on the reused row'
        }
    }

    $xl.Run('SetQuiet', $false)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Host ''
Write-Host 'closed without saving'
