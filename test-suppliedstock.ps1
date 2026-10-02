# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# "Supplied by student" paper stock (printer/paper compatibility rework,
# requirement 2): a student brings their own paper. Two built-in
# stocks, "Supplied (Roll)" / "Supplied (Sheet)", Cost = 0 - built into
# modCatalog.LoadCatalog rather than stored as tblPapers rows, so they cannot
# be deleted or renamed. Paper Cost computes to zero with no change to any
# formula (S_UnitCost = 0 is read exactly like any other stock's). A
# tblPapers row with Supplied by student = Yes is also charged zero for paper. Ink/consumable cost is charged
# normally from the printer's own rate, still waivable per-row via the
# existing Disregard Consumable flag. The real size is entered per job
# rather than read from a catalogue row: Print Width mm becomes required
# for 'Supplied (Roll)' (validated against the chosen printer's Max roll
# width mm), and the new job-row "Sheet size" column plays the same role
# for 'Supplied (Sheet)' (validated against Max sheet size).
#
# Uses Example Print Room's migrated catalogue data:
#   Epson SureColor P9500 -> Max roll width 1370mm
#   Xerox Versant 180     -> Max sheet size A1 (841x594mm)
#   HP DesignJet Z9+      -> Max roll width 610mm, no sheet capacity at all
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
    $main = $wb.Worksheets('Example Print Room')
    $lo = $main.ListObjects('tblJobs_MAIN')

    function Col($lo, $name) {
        for ($try = 1; $try -le 5; $try++) {
            try {
                for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
                    if ($lo.ListColumns($i).Name -eq $name) { return $i }
                }
            } catch {}
            Start-Sleep -Milliseconds 500
        }
        $have = @(); for ($k = 1; $k -le $lo.ListColumns.Count; $k++) { $have += $lo.ListColumns($k).Name }; throw "column '$name' not found in table '$($lo.Name)' (columns: $($have -join ', '))"
    }
    $prnCol = Col $lo 'Printer'
    $stkCol = Col $lo 'Paper Stock'
    $qtyCol = Col $lo 'Qty'
    $widthCol = Col $lo 'Print Width mm'
    $sizeCol = Col $lo 'Sheet size'
    $areaCol = Col $lo 'Area m2'
    $paperCostCol = Col $lo 'Paper Cost'
    $consCostCol = Col $lo 'Consumable Cost'
    $statusCol = Col $lo 'Status'

    function New-Row {
        [void]$main.Activate()
        [void]$xl.Run('btnAddPrintJob')
        # Right after the VBA add-row Excel can hand back $null for ListRows(n); retry until a real row number comes back.
        return Invoke-ComRetry -Attempts 5 {
            $row = [int]$lo.ListRows($lo.ListRows.Count).Range.Row
            if ($row -lt 1) { throw 'new row not readable yet' }
            $row
        }
    }

    # ---------------------------------------------------- roll, zero paper cost
    Write-Host '=== Supplied (Roll): zero paper cost, normal consumable cost ==='
    $r1 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r1, $prnCol).Value2 = 'Epson SureColor P9500' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r1, $stkCol).Value2 = 'Supplied (Roll)' }
    Start-Sleep -Milliseconds 300
    Invoke-ComRetry -Attempts 5 { $main.Cells($r1, $widthCol).Value2 = 900 }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r1, $qtyCol).Value2 = 5 }
    Start-Sleep -Milliseconds 300

    $paperCost = [double]$main.Cells($r1, $paperCostCol).Value2
    $consCost = [double]$main.Cells($r1, $consCostCol).Value2
    $area = [double]$main.Cells($r1, $areaCol).Value2
    Check ($paperCost -eq 0) "Paper Cost is 0 (got $paperCost)"
    Check ($consCost -gt 0) "Consumable Cost is charged normally (got $consCost)"
    Check ([Math]::Abs($area - (0.9 * 5)) -lt 0.001) "Area m2 uses the entered Print Width, not a catalogue width (got $area)"

    # ---------------------------------------------------------- roll, required
    Write-Host ''
    Write-Host '=== Supplied (Roll): Print Width mm required ==='
    $r2 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r2, $prnCol).Value2 = 'Epson SureColor P9500' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r2, $stkCol).Value2 = 'Supplied (Roll)' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r2, $qtyCol).Value2 = 3 }
    Start-Sleep -Milliseconds 300
    $status = [string]$main.Cells($r2, $statusCol).Value2
    Check ($status -like '*Print width required for student-supplied roll stock*') "Status flags the missing width (got '$status')"

    # ----------------------------------------------- roll, exceeds printer max
    Write-Host ''
    Write-Host '=== Supplied (Roll): width exceeding the printer max is rejected ==='
    $r3 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r3, $prnCol).Value2 = 'HP DesignJet Z9+' }   # Max roll width 610mm
    Invoke-ComRetry -Attempts 5 { $main.Cells($r3, $stkCol).Value2 = 'Supplied (Roll)' }
    Start-Sleep -Milliseconds 300
    Invoke-ComRetry -Attempts 5 { $main.Cells($r3, $widthCol).Value2 = 900 }
    Start-Sleep -Milliseconds 300
    $widthAfter = [string]$main.Cells($r3, $widthCol).Value2
    Check ([string]::IsNullOrEmpty($widthAfter)) "900mm on a 610mm-max printer is cleared, not accepted (got '$widthAfter')"
    $log = [string]$xl.Run('QuietLog')
    Check ($log -like '*maximum roll width*') 'a warning naming the printer''s maximum roll width was raised'

    # --------------------------------------------------- sheet, zero paper cost
    Write-Host ''
    Write-Host '=== Supplied (Sheet): zero paper cost, normal consumable cost ==='
    $r4 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r4, $prnCol).Value2 = 'Xerox Versant 180' }   # Max sheet size A1
    Invoke-ComRetry -Attempts 5 { $main.Cells($r4, $stkCol).Value2 = 'Supplied (Sheet)' }
    Start-Sleep -Milliseconds 300
    Invoke-ComRetry -Attempts 5 { $main.Cells($r4, $sizeCol).Value2 = 'A3' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r4, $qtyCol).Value2 = 10 }
    Start-Sleep -Milliseconds 300

    $paperCost4 = [double]$main.Cells($r4, $paperCostCol).Value2
    $consCost4 = [double]$main.Cells($r4, $consCostCol).Value2
    Check ($paperCost4 -eq 0) "Paper Cost is 0 (got $paperCost4)"
    Check ($consCost4 -gt 0) "Consumable Cost is charged normally (got $consCost4)"

    # -------------------------------------------------------- sheet, required
    Write-Host ''
    Write-Host '=== Supplied (Sheet): Sheet size required ==='
    $r5 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r5, $prnCol).Value2 = 'Xerox Versant 180' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r5, $stkCol).Value2 = 'Supplied (Sheet)' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r5, $qtyCol).Value2 = 4 }
    Start-Sleep -Milliseconds 300
    $status5 = [string]$main.Cells($r5, $statusCol).Value2
    Check ($status5 -like '*Sheet size required for student-supplied sheet stock*') "Status flags the missing size (got '$status5')"

    # -------------------------------------- sheet, exceeds a smaller printer max
    Write-Host ''
    Write-Host '=== Supplied (Sheet): size exceeding a smaller printer max is rejected ==='
    $printers = $wb.Worksheets('Printers').ListObjects('tblPrinters')
    $maxSizeCol = 0
    for ($i = 1; $i -le $printers.ListColumns.Count; $i++) {
        if ($printers.ListColumns($i).Name -eq 'Max sheet size') { $maxSizeCol = $i; break }
    }
    $xeroxRow = $null
    for ($i = 1; $i -le $printers.ListRows.Count; $i++) {
        if ([string]$printers.ListRows($i).Range.Cells(1, 2).Value2 -eq 'Xerox Versant 180') { $xeroxRow = $i; break }
    }
    $wb.Worksheets('Printers').Unprotect()
    $printers.ListRows($xeroxRow).Range.Cells(1, $maxSizeCol).Value2 = 'SRA3'
    $wb.Worksheets('Printers').Protect()
    $xl.Run('Invalidate') | Out-Null

    $r6 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r6, $prnCol).Value2 = 'Xerox Versant 180' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r6, $stkCol).Value2 = 'Supplied (Sheet)' }
    Start-Sleep -Milliseconds 300
    Invoke-ComRetry -Attempts 5 { $main.Cells($r6, $sizeCol).Value2 = 'A1' }   # bigger than the now-reduced SRA3 max
    Start-Sleep -Milliseconds 300
    $sizeAfter = [string]$main.Cells($r6, $sizeCol).Value2
    Check ([string]::IsNullOrEmpty($sizeAfter)) "A1 on an SRA3-max printer is cleared, not accepted (got '$sizeAfter')"
    $log6 = [string]$xl.Run('QuietLog')
    Check ($log6 -like '*maximum sheet size*') 'a warning naming the printer''s maximum sheet size was raised'

    # ------------------------------------------------------- Disregard Consumable still works
    Write-Host ''
    Write-Host '=== Disregard Consumable still waives ink cost on a supplied-stock row ==='
    $disregardCol = Col $lo 'Disregard Consumable'
    $chargeableCol = Col $lo 'Chargeable Cost'
    Invoke-ComRetry -Attempts 5 { $main.Cells($r1, $disregardCol).Value2 = 'Yes' }
    Start-Sleep -Milliseconds 300
    $chargeable1 = [double]$main.Cells($r1, $chargeableCol).Value2
    Check ($chargeable1 -eq 0) "Chargeable Cost drops to 0 once Disregard Consumable is set (got $chargeable1, Paper Cost was already 0)"

    # ------------------------------------- built in, not rows of tblPapers
    Write-Host ''
    Write-Host '=== Supplied (Roll)/(Sheet) are built in, not tblPapers rows ==='
    $pap = $null
    foreach ($s in $wb.Worksheets) { foreach ($l in $s.ListObjects) { if ($l.Name -eq 'tblPapers') { $pap = $l } } }
    $papWs = $pap.Parent
    function DescCount($pap, $desc) {
        $dc = Col $pap 'Description'; $n = 0
        for ($i = 1; $i -le $pap.ListRows.Count; $i++) { if ([string]$pap.DataBodyRange.Cells($i, $dc).Value2 -eq $desc) { $n++ } }
        return $n
    }
    function SetPap($row, $name, $val) { $c = Col $pap $name; if ($val -is [int]) { $val = [double]$val }; try { $cell = $pap.ListRows($row).Range.Cells(1, $c); if ($val -is [double]) { $cell.Formula = $val.ToString([Globalization.CultureInfo]::InvariantCulture) } else { $cell.Value2 = $val } } catch { Write-Host "SetPap FAILED row=$row col=$name($c) val=$val type=$($val.GetType().Name): $($_.Exception.Message)"; throw } }
    function GetPap($row, $name) { $c = Col $pap $name; return $pap.ListRows($row).Range.Cells(1, $c).Value2 }

    # A workbook built before this change still carries the two rows until
    # Setup runs; the migration is what Setup calls.
    $xl.Run('RemoveLegacySuppliedRows')
    Check ((DescCount $pap 'Supplied (Roll)') -eq 0) 'no Supplied (Roll) row in tblPapers'
    Check ((DescCount $pap 'Supplied (Sheet)') -eq 0) 'no Supplied (Sheet) row in tblPapers'
    $xl.Run('RemoveLegacySuppliedRows')
    Check ((DescCount $pap 'Supplied (Roll)') -eq 0) 'running the migration again is harmless'
    Check ([bool]$xl.Run('Compatible', 'Epson SureColor P9500', 'Supplied (Roll)')) 'Supplied (Roll) is still usable on a roll printer'
    Check (-not [bool]$xl.Run('Compatible', 'HP DesignJet Z9+', 'Supplied (Sheet)')) 'Supplied (Sheet) is refused by a printer with no sheet capacity'

    Write-Host ''
    Write-Host '=== A legacy Supplied row is removed by the migration ==='
    $papWs.Unprotect()
    $xl.EnableEvents = $false
    [void]$pap.ListRows.Add()
    $ln = Invoke-ComRetry -Attempts 5 { $c = [int]$pap.ListRows.Count; if ($c -lt 1) { throw "row count not readable yet" }; $c }
    SetPap $ln 'StockID' 'STK-SUP-ROLL'
    SetPap $ln 'Description' 'Supplied (Roll)'
    SetPap $ln 'Measure' 'Roll'
    SetPap $ln 'Cost' 0
    SetPap $ln 'Active' 'Yes'
    SetPap $ln 'Supplied by student' 'Yes'
    $xl.EnableEvents = $true
    Check ((DescCount $pap 'Supplied (Roll)') -eq 1) 'legacy row planted for the test'
    $xl.Run('RemoveLegacySuppliedRows')
    Check ((DescCount $pap 'Supplied (Roll)') -eq 0) 'the migration removes it'

    Write-Host ''
    Write-Host '=== Typing a reserved name into the Papers table is rejected ==='
    $papWs.Activate()
    $xl.Run('AddCatalogRow', 'tblPapers')
    $n = Invoke-ComRetry -Attempts 5 { $c = [int]$pap.ListRows.Count; if ($c -lt 1) { throw "row count not readable yet" }; $c }
    Check ([string](GetPap $n 'Active') -eq 'Yes') 'new row defaults Active to Yes'
    Check ([string](GetPap $n 'Supplied by student') -eq 'No') "Add Row defaults Supplied by student to No (got '$([string](GetPap $n 'Supplied by student'))')"
    $papWs.Unprotect()
    SetPap $n 'Description' 'Supplied (Sheet)'
    Start-Sleep -Milliseconds 300
    Check ([string]::IsNullOrEmpty([string](GetPap $n 'Description'))) 'the reserved name is cleared, not accepted'
    Check (([string]$xl.Run('QuietLog')) -like '*built-in paper stock*') 'a warning explained why'

    # ------------------------------------------- Supplied by student = Yes rows
    Write-Host ''
    Write-Host '=== A new stock defaults to Supplied by student = No ==='
    $papWs.Unprotect()
    SetPap $n 'Description' 'Bulk test roll'
    Start-Sleep -Milliseconds 300
    Check ([string](GetPap $n 'Supplied by student') -eq 'No') "Supplied by student defaults to No (got '$([string](GetPap $n 'Supplied by student'))')"
    $papWs.Unprotect()
    SetPap $n 'Measure' 'Roll'
    SetPap $n 'Size mode' 'Roll'
    SetPap $n 'Width mm' 610
    SetPap $n 'Cost' 7
    Start-Sleep -Milliseconds 300

    Write-Host ''
    Write-Host '=== Supplied by student = No: paper is charged ==='
    $r7 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r7, $prnCol).Value2 = 'Epson SureColor P9500' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r7, $stkCol).Value2 = 'Bulk test roll' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r7, $qtyCol).Value2 = 4 }
    Start-Sleep -Milliseconds 300
    $paper7 = [double]$main.Cells($r7, $paperCostCol).Value2
    Check ($paper7 -gt 0) "Paper Cost is charged for an ordinary stock (got $paper7)"

    Write-Host ''
    Write-Host '=== Supplied by student = Yes: cost forced to zero ==='
    $papWs.Unprotect()
    SetPap $n 'Supplied by student' 'Yes'
    Start-Sleep -Milliseconds 300
    Check ([double](GetPap $n 'Cost') -eq 0) "the Cost cell is reset to 0 (got $([string](GetPap $n 'Cost')))"
    Check (([string]$xl.Run('QuietLog')) -like '*not charged for*') 'a warning explained why'
    $r8 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r8, $prnCol).Value2 = 'Epson SureColor P9500' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r8, $stkCol).Value2 = 'Bulk test roll' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r8, $qtyCol).Value2 = 4 }
    Start-Sleep -Milliseconds 300
    $paper8 = [double]$main.Cells($r8, $paperCostCol).Value2
    $cons8 = [double]$main.Cells($r8, $consCostCol).Value2
    Check ($paper8 -eq 0) "Paper Cost is 0 (got $paper8)"
    Check ($cons8 -gt 0) "Consumable Cost is charged normally (got $cons8)"

    Write-Host ''
    Write-Host '=== The flag decides, not the Cost cell (a cost written behind the events) ==='
    $papWs.Unprotect()
    $xl.EnableEvents = $false
    SetPap $n 'Cost' 7
    $xl.EnableEvents = $true
    $xl.Run('Invalidate')
    $r9 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r9, $prnCol).Value2 = 'Epson SureColor P9500' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r9, $stkCol).Value2 = 'Bulk test roll' }
    Invoke-ComRetry -Attempts 5 { $main.Cells($r9, $qtyCol).Value2 = 4 }
    Start-Sleep -Milliseconds 300
    $paper9 = [double]$main.Cells($r9, $paperCostCol).Value2
    Check ($paper9 -eq 0) "Paper Cost is still 0 with Cost = 7 and Supplied by student = Yes (got $paper9)"

    Write-Host ''
    Write-Host '=== Such a stock keeps its catalogue size (no per-job size) ==='
    $r10 = New-Row
    Invoke-ComRetry -Attempts 5 { $main.Cells($r10, $prnCol).Value2 = 'HP DesignJet Z9+' }   # Max roll width 610mm - fits exactly
    Invoke-ComRetry -Attempts 5 { $main.Cells($r10, $stkCol).Value2 = 'Bulk test roll' }
    Start-Sleep -Milliseconds 300
    Check ([string]$main.Cells($r10, $stkCol).Value2 -eq 'Bulk test roll') 'a 610mm bulk roll is accepted on a 610mm-max printer'
    $status10 = [string]$main.Cells($r10, $statusCol).Value2
    Check ($status10 -notlike '*Print width required*') "Print width is NOT demanded for it (status '$status10')"
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
