# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Blanking a job row (a new print room, which resets its copied template row)
# must not wipe the calculated columns - Unit, Area m2, the costs and Status. They
# used to go with ClearContents, so every job added afterwards had no Unit and
# no costs. Drives a COPY in %TEMP%; closes without saving.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestCommon.ps1')
$workDir = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsTest-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$f = Join-Path $workDir 'PrintJob.xlsm'
Copy-Item (Join-Path $PSScriptRoot 'src\PrintJob.xlsm') $f

$vba = @'
Public Function MakeRoom() As String
    Dim ws As Worksheet
    SetQuiet True
    Set ws = CreatePrintRoom("Formula Room", "Facilities")
    SetQuiet True
    MakeRoom = ws.Name
End Function
'@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($f)
    $m = $null
    for ($i = 1; $i -le 6 -and -not $m; $i++) { try { $m = $wb.VBProject.VBComponents.Add(1) } catch { Start-Sleep -Seconds $i } }
    $m.Name = 'modClearFormulasTest'
    $m.CodeModule.AddFromString($vba)
    $name = $xl.Run('MakeRoom')
    $ws = $wb.Worksheets($name)
    $lo = $ws.ListObjects(1)
    $row = $lo.ListRows(1).Range.Row
    foreach ($h in 'Unit', 'Area m2', 'Paper Cost', 'Consumable Cost', 'Status') {
        $c = $ws.Cells($row, $lo.Range.Column + (Col $lo $h) - 1)
        Check ($c.HasFormula) "new room: '$h' keeps its formula"
    }
    Check ($lo.ListRows.Count -eq 1) 'new room keeps one blank row'
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
if ($script:anyFail) { exit 1 }
