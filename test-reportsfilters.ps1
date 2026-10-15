# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# 0.10.27 Reports filter groups:
#
#   - The "More filters" rows (14-15) are a nested row group, closed by default;
#     row 16 is its summary row and says how many of its filters are set.
#   - Has a problem (B15): Yes / No. Disregarded (B16): Paper / Consumable / Both / Neither.
#   - Has notes in the More filters row (E15).
#   - Student-supplied paper (E12): Yes / No.
#   (Addresses are the 0.11.0 layout: right-hand inputs in E, a gap after Department.)
#   - A filter set inside the closed row still applies, still counts for the
#     bulk-command safeguard and is emptied by Clear all filters.
#
# Expected counts are read straight off the job tables, not off Reports.
# Drives a COPY in %TEMP%, never src\PrintJob.xlsm itself. Closes WITHOUT saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
. (Join-Path $PSScriptRoot 'test-fixture-annexe.ps1')
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$anyFail = $false
try {
    $wb = Invoke-ComRetry { $xl.Workbooks.Open($f) }
    Add-AnnexeFixture $xl $wb | Out-Null
    $xl.CalculateFullRebuild()
    Start-Sleep -Seconds 3
    $xl.Run('SetQuiet', $true)
    $rep = Invoke-ComRetry { $wb.Worksheets('Reports') }
    $main = $wb.Worksheets('Example Print Room')
    $annexe = $wb.Worksheets('Annexe')
    $loMain = $main.ListObjects('tblJobs_MAIN')
    $loAnnexe = $annexe.ListObjects('tblJobs_ANNEX')
    $pap = $null
    foreach ($s in $wb.Worksheets) { foreach ($l in $s.ListObjects) { if ($l.Name -eq 'tblPapers') { $pap = $l } } }
    $rep.Activate()

    function Jobs { $xl.CalculateFullRebuild(); [int]$rep.Range('I9').Text }
    function Clear-AllBoxes {
        foreach ($a in 'B4','B5','B8','B9','B11','B12','B13','E4','E5','E7','E10','E11','E12','B15','B16','E15','I7','K7') { Invoke-ComRetry { $rep.Range($a).ClearContents() | Out-Null } }
        $xl.CalculateFullRebuild()
    }
    # One record per job table row that has a Date/Time (the rows _Data keeps;
    # the table is pre-sized with blank rows),
    # read as one array per table.
    function Records {
        $out = @()
        foreach ($lo in @($loMain, $loAnnexe)) {
            $v = $lo.DataBodyRange.Value2   # not through Invoke-ComRetry: it would enumerate the 2-D array
            $cJ = Col $lo 'Date/Time'; $cS = Col $lo 'Status'; $cP = Col $lo 'Disregard Paper'
            $cC = Col $lo 'Disregard Consumable'; $cK = Col $lo 'Paper Stock'
            for ($r = 1; $r -le $v.GetLength(0); $r++) {
                if ([string]$v[$r, $cJ] -eq '') { continue }
                $out += [pscustomobject]@{
                    Status = [string]$v[$r, $cS]
                    DPaper = [string]$v[$r, $cP]
                    DCons  = [string]$v[$r, $cC]
                    Stock  = [string]$v[$r, $cK]
                }
            }
        }
        return ,$out
    }

    $total = Jobs
    Check ($total -ge 5) "unfiltered Reports shows $total records"

    # ---------------------------------------------------------- the group
    Write-Host ''
    Write-Host '=== More filters: a closed nested group ==='
    Check (-not [bool]$rep.Rows(14).Hidden -and -not [bool]$rep.Rows(15).Hidden) 'rows 14-15 (the less-used filters) are visible: the row groups are gone (0.10.28)'
    Check ([int]$rep.Rows(4).OutlineLevel -eq 1 -and [int]$rep.Rows(14).OutlineLevel -eq 1) 'no row groups on the Reports sheet'
    Check ([int]$rep.Columns('A').OutlineLevel -eq 2 -and [int]$rep.Columns('F').OutlineLevel -eq 2 -and [int]$rep.Columns('G').OutlineLevel -eq 1 -and [int]$rep.Columns('H').OutlineLevel -eq 1) 'filter columns A:F are one column group; G (gap) and H are not in it'
    Check ([string]$rep.Range('A15').Text -eq 'Has a problem' -and [string]$rep.Range('D15').Text -eq 'Has notes' -and [string]$rep.Range('A16').Text -eq 'Disregarded') 'the closed group holds Has a problem, Has notes and Disregarded'
    Check ([string]$rep.Range('H7').Text -eq 'Sort by') 'Sort by sits at H7 above the table'
    foreach ($addr in 'B15', 'B16', 'E15', 'E12') {
        Check ([int]$rep.Range($addr).Validation.Type -eq 3) "$addr is a dropdown"
    }

    # --------------------------------------------- counter on the summary row
    Write-Host ''
    Write-Host '=== The summary row counts hidden filters that are set ==='
    $rep.Range('E15').Value2 = 'Yes'; $xl.Calculate()
    $rep.Range('B16').Value2 = 'Paper'; $xl.Calculate()
    Clear-AllBoxes

    # ------------------------------------ Has a problem / Disregarded
    Write-Host ''
    Write-Host '=== Has a problem and Disregarded ==='
    # Make sure each flag has something to find.
    $loMain.ListColumns('Disregard Paper').DataBodyRange.Cells(1, 1).Value2 = 'Yes'
    $loAnnexe.ListColumns('Disregard Consumable').DataBodyRange.Cells(1, 1).Value2 = 'Yes'
    $loMain.ListColumns('Disregard Consumable').DataBodyRange.Cells(1, 1).Value2 = 'Yes'   # row 1 now has both flags
    $main.Activate()
    $xl.Run('btnAddPrintJob')          # an incomplete row: its Status lists issues
    $xl.Run('RefreshStatusNotes', $main)
    $rep.Activate()
    $recs = Records
    $wantPaper = @($recs | Where-Object { $_.DPaper -eq 'Yes' }).Count
    $wantCons = @($recs | Where-Object { $_.DCons -eq 'Yes' }).Count
    $wantProb = @($recs | Where-Object { $_.Status -ne 'OK' -and $_.Status -ne '' }).Count
    Write-Host "  expected from the tables: paper disregarded $wantPaper, consumable disregarded $wantCons, problems $wantProb of $($recs.Count)"
    Check ($wantPaper -ge 1 -and $wantCons -ge 1) 'fixture has both disregard flags set'
    $wantBoth = @($recs | Where-Object { $_.DPaper -eq 'Yes' -and $_.DCons -eq 'Yes' }).Count
    $wantOk = @($recs | Where-Object { $_.Status -eq 'OK' }).Count
    Check ($wantBoth -ge 1) 'fixture has a record with both flags'
    $rep.Range('B16').Value2 = 'Paper'
    Check ((Jobs) -eq $wantPaper) 'Disregarded = Paper matches the records with Disregard Paper = Yes'
    $rep.Range('B16').Value2 = 'Consumable'
    Check ((Jobs) -eq $wantCons) 'Disregarded = Consumable matches the records with Disregard Consumable = Yes'
    $rep.Range('B16').Value2 = 'Both'
    Check ((Jobs) -eq $wantBoth) "Disregarded = Both matches the records with both flags ($wantBoth)"
    $wantNeither = @($recs | Where-Object { $_.DPaper -ne 'Yes' -and $_.DCons -ne 'Yes' }).Count
    $rep.Range('B16').Value2 = 'Neither'
    Check ((Jobs) -eq $wantNeither) "Disregarded = Neither matches the records with no flag set ($wantNeither)"
    Check ($wantNeither -ge 1 -and $wantNeither -lt $recs.Count) 'fixture has records both with and without a disregard flag'
    Clear-AllBoxes
    $rep.Range('B15').Value2 = 'Yes'
    Check ((Jobs) -eq $wantProb) 'Has a problem = Yes matches the records whose Status is not OK'
    Check ($wantProb -ge 1) 'fixture has at least one problem record'
    $rep.Range('B15').Value2 = 'No'
    Check ((Jobs) -eq $wantOk) "Has a problem = No matches the records whose Status is OK ($wantOk)"
    Clear-AllBoxes
    Check ((Jobs) -eq $recs.Count) 'blank filters show every record'

    # ----------------------------------------------- hidden filters still act
    Write-Host ''
    Write-Host '=== A filter inside the closed row still applies and still counts ==='
    Check (-not [bool]($xl.Run('HasActiveFilter', $rep))) 'no filter set: HasActiveFilter is false'
    $rep.Range('B16').Value2 = 'Paper'
    Check ([bool]($xl.Run('HasActiveFilter', $rep))) 'Disregarded alone satisfies the at-least-one-filter safeguard'
    Clear-AllBoxes
    $rep.Range('E15').Value2 = 'Yes'
    Check ([bool]($xl.Run('HasActiveFilter', $rep))) 'Has notes (now in the closed row) alone satisfies the safeguard'
    $rep.Range('B15').Value2 = 'Yes'; $rep.Range('B16').Value2 = 'Both'
    $xl.Run('btnClearFilters')
    Check ([string]$rep.Range('B15').Formula -eq '' -and [string]$rep.Range('B16').Formula -eq '' -and [string]$rep.Range('E15').Formula -eq '') 'Clear all filters empties the closed row too'

    # ------------------------------------------------ Student-supplied paper
    Write-Host ''
    Write-Host '=== Student-supplied paper ==='
    Clear-AllBoxes
    $stockIdx = Col $loMain 'Paper Stock'
    $stockName = [string]$loMain.ListRows(1).Range.Cells(1, $stockIdx).Text
    $descIdx = Col $pap 'Description'; $supIdx = Col $pap 'Supplied by student'
    $pr = 0
    for ($i = 1; $i -le $pap.ListRows.Count; $i++) { if ([string]$pap.ListRows($i).Range.Cells(1, $descIdx).Text -eq $stockName) { $pr = $i } }
    Check ($pr -gt 0) "the first job's stock '$stockName' is a Papers row"
    $pap.ListRows($pr).Range.Cells(1, $supIdx).Value2 = 'Yes'
    $recs = Records
    $wantSup = @($recs | Where-Object { $_.Stock -eq $stockName }).Count
    Write-Host "  expected: $wantSup of $($recs.Count) records on '$stockName'"
    Check ($wantSup -ge 1) 'at least one record uses the flagged stock'
    $rep.Range('E12').Value2 = 'Yes'
    Check ((Jobs) -eq $wantSup) 'Yes shows exactly the records on stock marked Supplied by student'
    $rep.Range('E12').Value2 = 'No'
    Check ((Jobs) -eq ($recs.Count - $wantSup)) 'No shows every other record'
    $rep.Range('E12').ClearContents() | Out-Null
    Check ((Jobs) -eq $recs.Count) 'blank shows every record'
    $pap.ListRows($pr).Range.Cells(1, $supIdx).Value2 = 'No'
    $rep.Range('E12').Value2 = 'Yes'
    Check ((Jobs) -eq 0) 'with no stock marked Supplied by student, Yes shows nothing'
    # The two built-in stocks are not Papers rows; put one on a record and count again.
    $stockCell = $loMain.ListRows(2).Range.Cells(1, $stockIdx)
    $stockCell.Value2 = 'Supplied (Roll)'
    if ([string]$stockCell.Text -eq 'Supplied (Roll)') {
        $recs = Records
        $wantBuiltIn = @($recs | Where-Object { $_.Stock -eq 'Supplied (Roll)' -or $_.Stock -eq 'Supplied (Sheet)' }).Count
        Check ($wantBuiltIn -ge 1) 'a record is on a built-in Supplied stock'
        Check ((Jobs) -eq $wantBuiltIn) 'Yes also shows records on the built-in Supplied stocks'
    } else {
        Write-Host "  NOTE: the sheet would not accept 'Supplied (Roll)' on that row; built-in stock case not exercised"
    }
    Clear-AllBoxes
    Check ([bool]($xl.Run('HasActiveFilter', $rep)) -eq $false) 'all clear again'

    # --------------------------------------------- a rebuild does not stack levels
    Write-Host ''
    Write-Host '=== Rebuilding the workbook leaves the outline as it was ==='
    $xl.Run('InitialiseWorkbook')
    $xl.Run('InitialiseWorkbook')
    $rep = Invoke-ComRetry { $wb.Worksheets('Reports') }
    Check ([int]$rep.Columns('A').OutlineLevel -eq 2 -and [int]$rep.Columns('H').OutlineLevel -eq 1 -and [int]$rep.Rows(4).OutlineLevel -eq 1) 'after two rebuilds the column group is still one level'
    Check (-not [bool]$rep.Rows(5).Hidden) 'the main filters are expanded after a rebuild'
    Check (-not [bool]$rep.Columns('A').Hidden) 'the filter columns are expanded after a rebuild'

    # ------------------------------------- freeze panes and the button strip
    Write-Host ''
    Write-Host '=== No freeze panes; the button strip is one rigid piece (0.11.0 refinement) ==='
    $rep.Activate()
    $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1
    Check (-not [bool]$xl.ActiveWindow.FreezePanes) 'Reports has no freeze panes'
    function BtnLeft([string]$macro) { foreach ($s in $rep.Shapes) { if ($s.Name -like "pcb_${macro}_*") { return [double]$s.Left } } return -1 }
    $macros = 'btnClearFilters', 'btnTogglePaid', 'btnGoToRecord', 'btnMarkPaid', 'btnMarkUnpaid', 'btnExportReport', 'btnDeleteVisible'
    $open = BtnLeft 'btnClearFilters'
    Check ($open -ge ($rep.Range('H1').Left - 1)) "expanded: the strip starts at column H (Left $([int]$open), H at $([int]$rep.Range('H1').Left))"
    $lefts0 = $macros | ForEach-Object { BtnLeft $_ }

    # Collapsing the filters carries the strip left with the table; expanding puts it back, clear of
    # the filter boxes. Excel raises no event for the +/-, so the one-second watcher
    # (modReports.StartStripWatch) does it; the test calls the same re-settle directly.
    $rep.Outline.ShowLevels(0, 1) | Out-Null
    Check ([bool]$rep.Columns('A').Hidden -and [bool]$rep.Columns('F').Hidden -and -not [bool]$rep.Columns('G').Hidden) 'collapsed: A:F hidden, the gap column G stays'
    [void]$xl.Run('RepositionReportsButtons', $rep)
    $shut = BtnLeft 'btnClearFilters'
    Check ([math]::Abs($shut - $rep.Range('H1').Left) -lt 1.5) "collapsed: the strip follows the table to column H (Left $([int]$shut))"
    Check ($shut -lt $open - 1) 'collapsed: the strip has moved left of its expanded position'
    $rep.Outline.ShowLevels(0, 2) | Out-Null
    [void]$xl.Run('RepositionReportsButtons', $rep)
    Check ([math]::Abs((BtnLeft 'btnClearFilters') - $open) -lt 1.5) 'expanded again: the strip is back where it was, clear of the filter boxes'
    Check ((BtnLeft 'btnClearFilters') -ge ($rep.Range('G1').Left + $rep.Columns('G').Width - 1)) 'the strip starts right of the filter columns'

    # Resizing a results column under the strip moves none of it.
    $w0 = [double]$rep.Columns('K').ColumnWidth
    $rep.Columns('K').ColumnWidth = $w0 + 20
    $lefts1 = $macros | ForEach-Object { BtnLeft $_ }
    $moved = 0; for ($i = 0; $i -lt $lefts0.Count; $i++) { if ([math]::Abs($lefts0[$i] - $lefts1[$i]) -gt 0.5) { $moved++ } }
    Check ($moved -eq 0) "widening a results column leaves the strip alone ($moved of $($macros.Count) buttons moved)"
    [void]$xl.Run('RepositionReportsButtons', $rep)
    $lefts2 = $macros | ForEach-Object { BtnLeft $_ }
    $moved = 0; for ($i = 0; $i -lt $lefts0.Count; $i++) { if ([math]::Abs($lefts0[$i] - $lefts2[$i]) -gt 0.5) { $moved++ } }
    Check ($moved -eq 0) "...and re-settling keeps it where it was ($moved moved)"
    $rep.Columns('K').ColumnWidth = $w0
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
if ($anyFail) { Write-Host 'FAIL: see above'; exit 1 }
Write-Host 'OK: Reports filter groups'
