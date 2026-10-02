# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Runs every test-*.ps1 in turn (each drives its own COPY of src\PrintJob.xlsm)
# and reports one line per script. A script FAILS if it exits non-zero OR prints
# a line containing the word FAIL - several scripts only print FAIL rather than
# setting an exit code, so output is the real signal.
#
#   run-tests.ps1                   all tests
#   run-tests.ps1 -Only groups,paid name fragments; runs matching scripts
#   run-tests.ps1 -Label baseline   log folder suffix (logs land in %TEMP%\testrun-<Label>)
#   run-tests.ps1 -KeepBackups 3    backups to keep after a green full run (default 1)
#
# After a FULL run (no -Only) in which every script passes, old src\*.bak.xlsm
# are pruned: the build has been tested, so only the newest backup is kept as
# a way back. A partial or failing run prunes nothing.
#
# Close Excel workbooks first; do not run while src\PrintJob.xlsm is open.

param([string[]]$Only, [string]$Label = (Get-Date -Format 'HHmmss'), [int]$KeepBackups = 1)

$logDir = Join-Path ([IO.Path]::GetTempPath()) "testrun-$Label"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$tests = Get-ChildItem $PSScriptRoot -Filter 'test-*.ps1' | Where-Object { $_.Name -ne 'test-fixture-annexe.ps1' } | Sort-Object Name
if ($Only) { $tests = $tests | Where-Object { $n = $_.Name; $Only | Where-Object { $n -like "*$_*" } } }

$results = @()
foreach ($t in $tests) {
    $log = Join-Path $logDir ($t.BaseName + '.log')
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Start-Sleep -Seconds 3   # let the previous test's Excel finish exiting; back-to-back launches flake
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $t.FullName 2>&1
    $code = $LASTEXITCODE
    $out | Out-File -FilePath $log -Encoding utf8
    $fails = @($out | Select-String -Pattern '\bFAIL\b|Exception|cannot be loaded').Count
    $status = if ($code -ne 0) { "EXIT $code" } elseif ($fails -gt 0) { "FAIL x$fails" } else { 'pass' }
    $retried = $false
    if ($status -ne 'pass') {   # COM timing flakes pass on a clean second attempt; real failures fail twice
        Start-Sleep -Seconds 8
        $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $t.FullName 2>&1
        $code = $LASTEXITCODE
        $out | Out-File -FilePath $log -Encoding utf8
        $fails = @($out | Select-String -Pattern '\bFAIL\b|Exception|cannot be loaded').Count
        $status = if ($code -ne 0) { "EXIT $code" } elseif ($fails -gt 0) { "FAIL x$fails" } else { 'pass (retry)' }
    }
    $results += [pscustomobject]@{ Test = $t.Name; Status = $status; Sec = [int]$sw.Elapsed.TotalSeconds }
    Write-Host ('{0,-32} {1,-10} {2,4}s' -f $t.Name, $status, [int]$sw.Elapsed.TotalSeconds)
}
$bad = @($results | Where-Object { $_.Status -notlike 'pass*' }).Count
Write-Host ("`n{0} of {1} passed. Logs: {2}" -f ($results.Count - $bad), $results.Count, $logDir)
if ($bad -eq 0 -and -not $Only -and $results.Count -gt 0) {
    & (Join-Path $PSScriptRoot 'prune-backups.ps1') -Keep $KeepBackups
}
exit $bad
