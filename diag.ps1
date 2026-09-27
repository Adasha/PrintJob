# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Strips the IFERROR off the breakdown formula so the real error shows.
$ErrorActionPreference = 'Stop'
# Works on a COPY in %TEMP%. This script deliberately mutates the workbook
# (it strips an IFERROR to expose the underlying error), and Workbook_Open
# plus Excel's cloud-backed AutoSave mean edits reach the real file whether or
# not it is ever saved - see the note at the top of the test-*.ps1 scripts.
$deliverable = Join-Path $PSScriptRoot 'src\PrintJob.xlsm'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsDiag-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item $deliverable $f
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $c = $wb.Worksheets('Cost Calculations')
    $c.Unprotect()

    $orig = $c.Range('Q15').Formula2
    Write-Host '--- Q15 as written ---'
    Write-Host $orig
    Write-Host ''

    # Peel: =IFERROR( <expr> ,"")   ->   = <expr>
    $inner = $orig
    if ($inner.StartsWith('=IFERROR(')) {
        $inner = $inner.Substring('=IFERROR('.Length)
        $inner = $inner.Substring(0, $inner.LastIndexOf(',""'))
        $inner = '=' + $inner
    }
    Write-Host '--- evaluating without IFERROR in AA40 ---'
    $c.Range('AA40').Formula2 = $inner
    $xl.CalculateFullRebuild()
    Write-Host ('AA40 -> ' + $c.Range('AA40').Text)

    # Build it up piece by piece to find where it breaks.
    $ok = '(IF($C$5="",TRUE,ISNUMBER(SEARCH($C$5,INDEX(_Data!$A$10#,,MATCH("Student Name",_Data!$A$9:$AZ$9,0))))))'
    $tests = @{
        'AA42' = '=ROWS(INDEX(_Data!$A$10#,,MATCH("Location",_Data!$A$9:$AZ$9,0)))'
        'AA43' = '=ROWS(FILTER(INDEX(_Data!$A$10#,,MATCH("Location",_Data!$A$9:$AZ$9,0)),' + $ok + '))'
        'AA44' = '=ROWS(SORT(UNIQUE(FILTER(INDEX(_Data!$A$10#,,MATCH("Location",_Data!$A$9:$AZ$9,0)),' + $ok + '))))'
        'AA45' = '=LET(k,FILTER(INDEX(_Data!$A$10#,,MATCH("Location",_Data!$A$9:$AZ$9,0)),' + $ok + '),u,SORT(UNIQUE(k)),TEXTJOIN("/",TRUE,BYROW(u,LAMBDA(x,SUM(--(k=x))))))'
        'AA46' = '=LET(k,FILTER(INDEX(_Data!$A$10#,,MATCH("Location",_Data!$A$9:$AZ$9,0)),' + $ok + '),g,FILTER(INDEX(_Data!$A$10#,,MATCH("Gross Cost",_Data!$A$9:$AZ$9,0)),' + $ok + '),u,SORT(UNIQUE(k)),TEXTJOIN("/",TRUE,BYROW(u,LAMBDA(x,SUM(FILTER(g,k=x,0))))))'
        'AA47' = '=LET(u,SORT(UNIQUE(FILTER(INDEX(_Data!$A$10#,,MATCH("Location",_Data!$A$9:$AZ$9,0)),' + $ok + '))),ROWS(VSTACK(HSTACK("a","b"),HSTACK(u,u))))'
    }
    foreach ($k in ($tests.Keys | Sort-Object)) {
        $c.Range($k).Formula2 = $tests[$k]
    }
    $xl.CalculateFullRebuild()
    foreach ($k in ($tests.Keys | Sort-Object)) {
        Write-Host ("{0} -> {1}" -f $k, $c.Range($k).Text)
    }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
