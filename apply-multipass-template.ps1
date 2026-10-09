# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# One-off, re-runnable template change for multi-pass costing (0.11.0, schema
# 1.4; docs/ARCHITECTURE.md section 17). PrintCosts.xlsx is binary, so the change
# is kept as this script rather than as a diff.
#
# It works on a COPY in %TEMP% (never opens the real file - see ARCHITECTURE
# 13.1), and only replaces src\PrintCosts.xlsx when every step has succeeded.
# Idempotent: a step whose result is already in place is skipped, so running it
# twice changes nothing.
#
#   powershell -ExecutionPolicy Bypass -File apply-multipass-template.ps1
#
# What it does:
#   1. Moves the location sheet's right-hand blocks (the LOC_* settings block
#      AL:AM and, in code, the side panel) right of the wider job table.
#   2. Job table: adds Passes, Colour, Set-up Cost, Row Type, Parent, Pass,
#      H_Ink, H_Setup, H_Notices, S_SetupCost, S_ColourID; rewrites Consumable
#      Cost, Gross Cost, Chargeable Cost and Status for pass rows.
#   3. Printers: adds Colour mode and Template cost.
#   4. Settings: Risograph consumable type, MULTIPASS_COLS_HIDDEN, and the
#      Reduced-view default hides the multi-pass columns.
#   5. Writes summaryBelow="0" into the location sheet's XML (Outline.SummaryRow
#      cannot be changed once a table exists).

[CmdletBinding()]
param(
    [string] $Template = (Join-Path $PSScriptRoot 'src\PrintCosts.xlsx'),
    [string] $DebugCheckpoints = ''   # a folder: SaveCopyAs after each step, to bisect a bad save
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Template)) { throw "Not found: $Template" }

$work = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsMultipass-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item $Template $work
$stampBefore = (Get-Item $Template).LastWriteTimeUtc

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
$locName = 'Example Print Room'

function Checkpoint([string]$tag) { if ($DebugCheckpoints) { $wb.SaveCopyAs((Join-Path $DebugCheckpoints ($tag + '.xlsx'))) } }
function Col($lo, $name) { foreach ($c in $lo.ListColumns) { if ($c.Name -eq $name) { return $c } } return $null }

# Adds a column at the 1-based position (or leaves an existing one), copying the look
# of a model column. Returns the ListColumn.
function Ensure-Col($lo, [string]$name, [int]$pos, [string]$model) {
    $c = Col $lo $name
    if ($c) { return $c }
    $c = $lo.ListColumns.Add($pos)
    $c.Name = $name
    $m = Col $lo $model
    if ($m) {
        $m.Range.Copy() | Out-Null
        $c.Range.PasteSpecial(-4122) | Out-Null     # xlPasteFormats
        try { $xl.CutCopyMode = $false } catch {}
        $c.Range.Cells.Item(1,1).Value2 = $name      # the paste can take the model's header text
    }
    return $c
}

try {
    $wb = $xl.Workbooks.Open($work)
    try { $wb.AutoSaveOn = $false } catch {}    # only settable on a cloud-backed file
    $ws =$wb.Worksheets.Item($locName)
    $lo = $ws.ListObjects.Item(1)

    # The shipped sheets are protected (no password); build setup re-applies the real
    # protection, so each is unprotected here and given a plain Protect() at the end.
    $prot = @()
    foreach ($s in $wb.Worksheets) { if ($s.ProtectContents) { $s.Unprotect(); $prot += $s.Name } }

    # ---- 1. move the LOC_* block AL1:AM10 to AW1:AX10 (names follow a Cut) -------------
    if ($ws.Range('AL1').Value2 -eq 'Location settings') {
        $ws.Range('AL1:AM10').Cut($ws.Range('AW1')) | Out-Null
        $ws.Columns('AW').ColumnWidth = $ws.Columns('AL').ColumnWidth
        $ws.Columns('AX').ColumnWidth = $ws.Columns('AM').ColumnWidth
        Write-Host 'moved LOC_* block AL:AM -> AW:AX'
    }
    Checkpoint 'a-moved'

    # Column widths and Hidden belong to the sheet COLUMN, not the table column, so an
    # insert leaves them behind (ARCHITECTURE section 5). Capture them by header name now
    # and put them back after every insert.
    function Save-Widths($t) {
        $d = @{}
        foreach ($c in $t.ListColumns) { $e = $c.Range.EntireColumn; $d[$c.Name] = @{ W = $e.ColumnWidth; H = [bool]$e.Hidden } }
        return $d
    }
    function Restore-Widths($t, $d) {
        foreach ($c in $t.ListColumns) {
            if (-not $d.ContainsKey($c.Name)) { continue }
            $e = $c.Range.EntireColumn
            $e.Hidden = $false
            if (-not $d[$c.Name].H) { $e.ColumnWidth = [double]$d[$c.Name].W } else { $e.ColumnWidth = [double]0; $e.Hidden = $true }
        }
    }
    $jobW = Save-Widths $lo
    $prnW = $null

    # ---- 2. job table columns ------------------------------------------------------
    $null = Ensure-Col $lo 'Passes'     ((Col $lo 'Printer').Index + 1)  'Qty'
    $null = Ensure-Col $lo 'Colour'      ((Col $lo 'Passes').Index + 1)   'Student No'
    $null = Ensure-Col $lo 'Set-up Cost' ((Col $lo 'Consumable Cost').Index + 1) 'Consumable Cost'
    $null = Ensure-Col $lo 'Row Type'    ((Col $lo 'Job ID').Index + 1)   'Job ID'
    $null = Ensure-Col $lo 'Parent'      ((Col $lo 'Row Type').Index + 1) 'Job ID'
    $null = Ensure-Col $lo 'Pass'        ((Col $lo 'Parent').Index + 1)   'Qty'
    $null = Ensure-Col $lo 'H_Ink'       ((Col $lo 'H_Issues').Index + 1) 'H_Issues'
    $null = Ensure-Col $lo 'H_Setup'     ((Col $lo 'H_Ink').Index + 1)    'H_Issues'
    $null = Ensure-Col $lo 'H_Notices'   ((Col $lo 'H_Setup').Index + 1)  'H_Issues'
    $null = Ensure-Col $lo 'S_SetupCost' ((Col $lo 'S_ConsRate').Index + 1) 'S_ConsRate'
    $null = Ensure-Col $lo 'S_ColourID'  ((Col $lo 'S_SetupCost').Index + 1) 'S_StockID'
    Restore-Widths $lo $jobW

    $n = $lo.ListRows.Count
    $f = @{
        'Passes'          = '=IF([@[Row Type]]="Pass","",IF([@[Job ID]]="","",IF(COUNTIFS([Parent],[@[Job ID]])=0,"",COUNTIFS([Parent],[@[Job ID]]))))'
        'Consumable Cost' = '=IF([@[Row Type]]="Pass",[@[H_Ink]],IF([@[Area m2]]="","",ROUND([@[Area m2]]*[@[S_ConsRate]],SET_ROUND_DP)+IF([@[Job ID]]="",0,SUMIFS([H_Ink],[Parent],[@[Job ID]]))))'
        'Set-up Cost'     = '=IF([@[Row Type]]="Pass",[@[H_Setup]],IF([@[Area m2]]="","",IF([@[Job ID]]="",0,SUMIFS([H_Setup],[Parent],[@[Job ID]]))))'
        'Gross Cost'      = '=IF([@[Paper Cost]]="","",[@[Paper Cost]]+[@[Consumable Cost]]+N([@[Set-up Cost]]))'
        'Chargeable Cost' = '=IF([@[Paper Cost]]="","",IF([@[Disregard Paper]]="Yes",0,[@[Paper Cost]])+IF([@[Disregard Consumable]]="Yes",0,[@[Consumable Cost]]+N([@[Set-up Cost]])))'
        'Status'          = '=IF(AND([@[Job ID]]="",[@[Row Type]]<>"Pass"),"",IF([@[H_Issues]]<>"",MID([@[H_Issues]],3,255),IF([@[H_Notices]]<>"",MID([@[H_Notices]],3,255),"OK")))'
        'H_Ink'           = '=IF([@[Row Type]]<>"Pass","",IFERROR(ROUND(XLOOKUP([@Parent],[Job ID],[Area m2])*N([@[S_ConsRate]]),SET_ROUND_DP),0))'
        'H_Setup'         = '=IF([@[Row Type]]<>"Pass","",N([@[S_SetupCost]]))'
    }
    foreach ($k in $f.Keys) {
        try { (Col $lo $k).DataBodyRange.Formula2 = $f[$k] } catch { throw "formula for column '$k' rejected: $($f[$k])" }
    }

    Checkpoint 'c-formulas'
    # Sample rows are Job rows (a blank Row Type is read as Job as well).
    (Col $lo 'Row Type').DataBodyRange.Value2 = 'Job'

    # Formats, widths, visibility.
    (Col $lo 'Passes').DataBodyRange.NumberFormat = '0'
    (Col $lo 'Passes').DataBodyRange.HorizontalAlignment = -4108
    (Col $lo 'Pass').DataBodyRange.NumberFormat = '0'
    (Col $lo 'Pass').DataBodyRange.HorizontalAlignment = -4108
    foreach ($h in 'Set-up Cost') { (Col $lo $h).DataBodyRange.NumberFormat = (Col $lo 'Consumable Cost').DataBodyRange.Cells.Item(1,1).NumberFormat }
    (Col $lo 'S_SetupCost').DataBodyRange.NumberFormat = (Col $lo 'S_ConsRate').DataBodyRange.Cells.Item(1,1).NumberFormat
    $widths = @{ 'Passes'=7.5; 'Colour'=16; 'Set-up Cost'=12.1; 'Row Type'=8.5; 'Parent'=14; 'Pass'=6 }
    foreach ($k in $widths.Keys) { (Col $lo $k).Range.EntireColumn.ColumnWidth = [double]$widths[$k] }
    foreach ($h in 'H_Ink','H_Setup','H_Notices','S_SetupCost','S_ColourID') {
        $r = (Col $lo $h).Range.EntireColumn
        $r.ColumnWidth = [double]0; $r.Hidden = $true
    }
    Write-Host ("job table now {0} columns" -f $lo.ListColumns.Count)

    Checkpoint 'd-formats'
    # ---- 3. Printers ---------------------------------------------------------------
    $pws = $wb.Worksheets.Item('Printers')
    $plo = $pws.ListObjects.Item(1)
    $prnW = Save-Widths $plo
    $null = Ensure-Col $plo 'Colour mode'   ((Col $plo 'Cost per m2').Index + 1) 'Consumable type'
    $null = Ensure-Col $plo 'Template cost' ((Col $plo 'Colour mode').Index + 1) 'Cost per m2'
    Restore-Widths $plo $prnW
    foreach ($cell in (Col $plo 'Colour mode').DataBodyRange.Cells) { if (-not $cell.Value2) { $cell.Value2 = 'single pass' } }
    $v = (Col $plo 'Colour mode').DataBodyRange.Validation
    $v.Delete()
    $v.Add(3, 1, 1, 'single pass,multi-pass')       # xlValidateList, xlValidAlertStop, xlBetween
    $v.InputTitle = 'Colour mode'
    $v.InputMessage = 'single pass: one flat Cost per m2, as always. multi-pass (a RISO duplicator, say): each job is printed in several colour passes, costed from the Consumables sheet, and each pass costs the Template cost to set up.'
    $v.ErrorTitle = 'Colour mode'
    $v.ErrorMessage = 'Choose single pass or multi-pass.'
    $v = (Col $plo 'Template cost').DataBodyRange.Validation
    $v.Delete()
    $v.Add(2, 1, 7, '0')                            # xlValidateDecimal, xlValidAlertStop, xlGreaterEqual
    $v.InputTitle = 'Template cost'
    $v.InputMessage = 'Multi-pass printers only: what it costs to set up one colour pass (making a RISO master, for example). Charged once per pass, however many copies are run. Leave blank for none.'
    $v.ErrorTitle = 'Template cost'
    $v.ErrorMessage = 'Enter a number, zero or more.'
    (Col $plo 'Colour mode').Range.EntireColumn.ColumnWidth = [double]14
    (Col $plo 'Template cost').Range.EntireColumn.ColumnWidth = [double]14

    Checkpoint 'e-printers'
    # ---- 4. Settings ---------------------------------------------------------------
    $sws = $wb.Worksheets.Item('Settings')
    $clo = $sws.ListObjects.Item('tblConsumables')
    $have = $false
    foreach ($r in $clo.ListRows) { if ($r.Range.Cells.Item(1,1).Value2 -eq 'Risograph') { $have = $true } }
    if (-not $have) {
        $last = $clo.ListRows.Item($clo.ListRows.Count)
        if ($last.Range.Cells.Item(1,1).Value2) { $last = $clo.ListRows.Add() }
        $last.Range.Cells.Item(1,1).Value2 = 'Risograph'
        $last.Range.Cells.Item(1,2).Value2 = 'Yes'
    }
    $slo = $sws.ListObjects.Item('tblSettings')
    $keyCol = (Col $slo 'Key').Index
    $rowOf = @{}
    for ($i = 1; $i -le $slo.ListRows.Count; $i++) { $rowOf[[string]$slo.ListRows.Item($i).Range.Cells.Item(1,$keyCol).Value2] = $i }
    if (-not $rowOf.ContainsKey('MULTIPASS_COLS_HIDDEN')) {
        $new = $slo.ListRows.Add($rowOf['COST_COLS_HIDDEN'] + 1)
        $new.Range.Cells.Item(1,1).Value2 = 'MULTIPASS_COLS_HIDDEN'
        $new.Range.Cells.Item(1,2).Value2 = 'Hide multi-pass columns'
        $new.Range.Cells.Item(1,3).Value2 = 'No'
        $new.Range.Cells.Item(1,4).Value2 = 'Yes hides the multi-pass columns (Passes, Colour, Row Type, Parent, Pass) on every location sheet (toggled by the button on each one). No shows them. The Set-up Cost column belongs to the cost detail setting instead.'
        $null = $wb.Names.Add('SET_MULTIPASS_COLS_HIDDEN', ('=Settings!$C$' + ($new.Range.Row)))
    }
    # The Reduced view does NOT hide the multi-pass columns: only their own toggle does. (An
    # interim build named them here; put the list back to its original value if it has them.)
    for ($i = 1; $i -le $slo.ListRows.Count; $i++) {
        $rr = $slo.ListRows.Item($i).Range
        if ($rr.Cells.Item(1,1).Value2 -eq 'LOC_REDUCED_COLUMNS' -and $rr.Cells.Item(1,3).Value2 -eq 'Status;Job ID;Area m2;S_SchemaVer;Passes;Colour;Row Type;Parent;Pass') {
            $rr.Cells.Item(1,3).Value2 = 'Status;Job ID;Area m2;S_SchemaVer'
        }
        if ($rr.Cells.Item(1,1).Value2 -eq 'SCHEMA') { $rr.Cells.Item(1,3).Value2 = '1.4' }
    }

    Checkpoint 'f-settings'
    foreach ($n in $prot) { $wb.Worksheets.Item($n).Protect() }
    Checkpoint 'g-protected'
    $wb.Save()
    $wb.Close($false)
    $wb = $null
}
catch {
    Write-Host ("FAILED at script line {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim())
    throw
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

if ($DebugCheckpoints) { Copy-Item $work (Join-Path $DebugCheckpoints 'h-saved.xlsx') -Force }
# ---- 5. summaryBelow="0" in the location sheet XML ---------------------------------
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($work, 'Update')
try {
    $wbXml = New-Object System.Xml.XmlDocument
    $e = $zip.GetEntry('xl/workbook.xml'); $sr = New-Object IO.StreamReader($e.Open()); $wbXml.LoadXml($sr.ReadToEnd()); $sr.Close()
    $relsXml = New-Object System.Xml.XmlDocument
    $e = $zip.GetEntry('xl/_rels/workbook.xml.rels'); $sr = New-Object IO.StreamReader($e.Open()); $relsXml.LoadXml($sr.ReadToEnd()); $sr.Close()
    $ns = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
    $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $sheet = $wbXml.SelectSingleNode("//m:sheet[@name='$locName']", $ns)
    $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $target = $null
    foreach ($rel in $relsXml.DocumentElement.ChildNodes) { if ($rel.Id -eq $rid) { $target = $rel.Target } }
    $part = 'xl/' + ($target -replace '^/?(xl/)?', '')
    $e = $zip.GetEntry($part)
    $sr = New-Object IO.StreamReader($e.Open()); $xml = $sr.ReadToEnd(); $sr.Close()
    if ($xml -notmatch 'summaryBelow="0"') {
        # '<sheetPr' must be followed by a space, '/' or '>': a bare prefix match also hits
        # <sheetProtection/> and corrupts the sheet. Instance Replace(..., 1) replaces the first match only.
        $self  = [regex]::new('<sheetPr(?<a>(?:\s[^>]*?)?)/>')
        $open  = [regex]::new('<sheetPr(?:\s[^>]*)?>')
        $tab   = [regex]::new('<tabColor[^>]*/>')
        $outl  = [regex]::new('<outlinePr')
        $root  = [regex]::new('(<worksheet[^>]*>)')
        if ($self.IsMatch($xml)) {
            $xml = $self.Replace($xml, '<sheetPr${a}><outlinePr summaryBelow="0"/></sheetPr>', 1)
        } elseif ($outl.IsMatch($xml)) {
            $xml = $outl.Replace($xml, '<outlinePr summaryBelow="0"', 1)
        } elseif ($open.IsMatch($xml)) {
            if ($tab.IsMatch($xml)) { $xml = $tab.Replace($xml, { param($m) $m.Value + '<outlinePr summaryBelow="0"/>' }, 1) }
            else { $xml = $open.Replace($xml, { param($m) $m.Value + '<outlinePr summaryBelow="0"/>' }, 1) }
        } else {
            $xml = $root.Replace($xml, '$1<sheetPr><outlinePr summaryBelow="0"/></sheetPr>', 1)
        }
        $e.Delete()
        $ne = $zip.CreateEntry($part)
        $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false)))
        $sw.Write($xml); $sw.Close()
        Write-Host "summaryBelow=0 written to $part"
    }
}
finally { $zip.Dispose() }

if ($DebugCheckpoints) { Copy-Item $work (Join-Path $DebugCheckpoints 'i-patched.xlsx') -Force }
if ((Get-Item $Template).LastWriteTimeUtc -ne $stampBefore) { throw "$Template changed while the script ran; not overwriting it." }
Copy-Item $work $Template -Force
Remove-Item $work -Force
Write-Host "template updated: $Template"

