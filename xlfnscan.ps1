# Finds every _xlfn / _xlws occurrence in a workbook's XML.
#
# Reads the OOXML package directly, so it catches things the grid does not
# show: hidden and very-hidden sheets, conditional formatting rules, data
# validation formulas, and defined names.
param([string] $File = 'src\PrintCosts.xlsm')

$ErrorActionPreference = 'Stop'
$path = if ([IO.Path]::IsPathRooted($File)) { $File } else { Join-Path $PSScriptRoot $File }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($path)
try {
    Write-Host ("FILE {0}" -f $path)
    $found = $false
    foreach ($e in $zip.Entries) {
        if ($e.FullName -notlike '*.xml') { continue }
        $sr = New-Object IO.StreamReader($e.Open())
        $text = $sr.ReadToEnd()
        $sr.Close()
        if ($text -notmatch '_xlfn|_xlws') { continue }
        $found = $true
        Write-Host ''
        Write-Host ("--- {0} ---" -f $e.FullName)
        # Report each distinct prefixed function name and how many times it appears.
        $hits = [regex]::Matches($text, '_xlfn\.(_xlws\.)?[A-Za-z0-9_.]+|_xlws\.[A-Za-z0-9_.]+')
        $hits | ForEach-Object { $_.Value } | Group-Object | Sort-Object Name |
            ForEach-Object { Write-Host ("   {0,-28} x{1}" -f $_.Name, $_.Count) }
    }
    if (-not $found) { Write-Host '  no _xlfn / _xlws anywhere' }
}
finally { $zip.Dispose() }
