# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Regression test: removing a table's ONLY remaining row, then adding a row,
# must not dress the new row like the table header.
#
# ListRows(1).Delete on a one-row table takes it to zero rows. With no data row
# left for Excel to copy from, the next ListRows.Add copies the HEADER row's
# formatting (navy fill, white bold wrapped text). Seen on the Consumables
# sheet (tblColours, 2026-10-15) and the same on Print Technicians, Printers,
# Papers, Departments and the job tables. modUtils.BlankSoleRow now blanks the
# last row instead (as Clear table / Clear All do).
#
# For each table: delete down to one row, remove that row, then add two rows
# and compare their look with the original first row. Job tables go through
# DeleteJobGroupAt.
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

    # How a cell looks, as one comparable string.
    function Look($cell) {
        return ('fill={0}/{1} bold={2} font={3} wrap={4}' -f $cell.Interior.Color, $cell.Interior.Pattern, $cell.Font.Bold, $cell.Font.Color, $cell.WrapText)
    }
    # Looks of the first two columns of table row $n.
    function RowLook($lo, [int]$n) {
        $r = $lo.ListRows($n).Range
        return (Look $r.Cells(1, 1)) + ' | ' + (Look $r.Cells(1, 2))
    }

    $cases = @(
        @{ Sheet = 'Print Technicians'; Table = 'tblTechnicians'; Jobs = $false },
        @{ Sheet = 'Printers';          Table = 'tblPrinters';    Jobs = $false },
        @{ Sheet = 'Papers';            Table = 'tblPapers';      Jobs = $false },
        @{ Sheet = 'Departments';       Table = 'tblDepartments'; Jobs = $false },
        @{ Sheet = 'Consumables';       Table = 'tblColours';     Jobs = $false },
        @{ Sheet = 'Example Print Room'; Table = 'tblJobs_EXAMPLE';  Jobs = $true }
    )

    foreach ($c in $cases) {
        Write-Host ''
        Write-Host ("=== {0} / {1} ===" -f $c.Sheet, $c.Table)
        $ws = $wb.Worksheets($c.Sheet)
        $lo = $ws.ListObjects($c.Table)
        [void]$xl.Run('UnlockSheet', $ws)
        $hdrLook = Look $lo.HeaderRowRange.Cells(1, 1)

        # Down to one row (bypassing the sole-row guard, which only the last delete hits).
        while ($lo.ListRows.Count -gt 1) { $lo.ListRows($lo.ListRows.Count).Range.EntireRow.Delete() | Out-Null }
        $want = RowLook $lo 1
        Check ($want -notlike "*$hdrLook*") 'the remaining row does not look like the header to begin with'

        if ($c.Jobs) {
            [void]$xl.Run('DeleteJobGroupAt', $ws, $lo, 1, 'test')
        } else {
            [void]$xl.Run('DoRemoveCatalogRow', $lo, 1)
        }
        Check ($lo.ListRows.Count -eq 1) 'removing the last row leaves one blank row, not zero'
        Check (($null -ne $lo.DataBodyRange)) 'the table still has a data body'
        Check ([int]$xl.Run('RowCount', $lo) -eq 0) 'RowCount reports zero rows'
        Check ((RowLook $lo 1) -eq $want) 'the blank row keeps its own formatting'

        # Add two rows the old way (ListRows.Add) - the second copies the first.
        [void]$xl.Run('UnlockSheet', $ws)
        $r = $lo.ListRows.Add()
        Check ((RowLook $lo $r.Index) -eq $want) 'a row added afterwards is not dressed like the header'
        Check ((RowLook $lo $r.Index) -notlike "*$hdrLook*") 'new row does not carry the header fill'
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
if ($script:anyFail) { exit 1 }
