$ErrorActionPreference = 'Continue'
$dir = Join-Path $PSScriptRoot 'src'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Peek($p) {
    $name = Split-Path $p -Leaf
    $fi = Get-Item $p
    try {
        $z = [IO.Compression.ZipFile]::OpenRead($p)
        $vba    = ($z.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' }) -ne $null
        $sheets = ($z.Entries | Where-Object { $_.FullName -like 'xl/worksheets/*.xml' }).Count
        $type = '?'
        $ct = $z.Entries | Where-Object { $_.FullName -eq '[Content_Types].xml' }
        if ($ct) {
            $sr = New-Object IO.StreamReader($ct.Open()); $x = $sr.ReadToEnd(); $sr.Close()
            if ($x -match 'ms-excel\.sheet\.macroEnabled') { $type = 'macroEnabled' }
            elseif ($x -match 'spreadsheetml\.sheet\.main') { $type = 'plain xlsx' }
        }
        $title = ''; $desc = ''
        $core = $z.Entries | Where-Object { $_.FullName -eq 'docProps/core.xml' }
        if ($core) {
            $sr = New-Object IO.StreamReader($core.Open()); $cx = [xml]$sr.ReadToEnd(); $sr.Close()
            $title = $cx.coreProperties.title
            $desc  = $cx.coreProperties.description
        }
        $z.Dispose()
        Write-Host ("{0,-42} {1,8}b  {2:dd/MM HH:mm:ss}  vba={3,-5} sheets={4,-3} {5}" -f `
            $name, $fi.Length, $fi.LastWriteTime, $vba, $sheets, $type)
        if ($desc) { Write-Host ("    -> {0}" -f $desc) }
        elseif ($title) { Write-Host ("    -> {0}" -f $title) }
    } catch {
        Write-Host ("{0,-42} {1,8}b  UNREADABLE: {2}" -f $name, $fi.Length, $_.Exception.Message)
    }
}

Write-Host '=== source and current build ==='
foreach ($n in 'PrintCosts.xlsx', 'PrintCosts.xlsm') {
    $p = Join-Path $dir $n
    if (Test-Path $p) { Peek $p } else { Write-Host "$n MISSING" }
}
Write-Host ''
Write-Host '=== backups, newest first ==='
Get-ChildItem $dir -Filter '*.bak.xlsm' | Sort-Object Name -Descending | ForEach-Object { Peek $_.FullName }
