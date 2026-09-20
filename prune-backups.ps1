# Keeps the most recent N build backups and removes the rest.
#
# The backups are build outputs, regenerable from src\PrintCosts.xlsx plus the
# VBA source, so a deep history of them earns nothing and syncs a few hundred
# megabytes over time. build.ps1 calls this after every successful build.
param([int] $Keep = 5)

$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot 'src'

$all = Get-ChildItem $dir -Filter '*.bak.xlsm' -ErrorAction SilentlyContinue | Sort-Object Name -Descending
if (-not $all -or $all.Count -le $Keep) {
    Write-Host ("backups: {0}, keeping all (limit {1})" -f @($all).Count, $Keep)
    return
}

$drop  = $all | Select-Object -Skip $Keep
$freed = ($drop | Measure-Object Length -Sum).Sum
$drop | Remove-Item -Force

Write-Host ("backups: {0} -> {1}, removed {2} ({3} MB)" -f `
    $all.Count, $Keep, $drop.Count, [math]::Round($freed / 1MB, 2))
