# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Shared by the test-*.ps1 regression scripts (dot-sourced): Check and Col at the bottom; first, a retry-with-backoff wrapper
# for a COM call that immediately follows a heavy VBA operation on the
# macro-enabled deliverable - most commonly Workbooks.Open itself
# (Workbook_Open runs real work on every open: ProtectAll, Invalidate,
# HealButtons, docs/ARCHITECTURE.md §13.1/§13.2), but the same shape recurs
# anywhere a Run() call leaves Excel still settling a large recalculation
# when it returns (e.g. ApplyImportConfirmed's own Check-sheet sweep,
# test-import.ps1). Excel can still be busy with that when the test
# script's own next COM call lands, rejecting it outright: "Call was
# rejected by callee. (Exception from HRESULT: 0x80010001
# (RPC_E_CALL_REJECTED))", or handing back "You cannot call a method on a
# null-valued expression" instead of the object it should have returned.
# Same root cause as build.ps1's SaveAs retry loop (~line 179); this is
# that pattern reused for the read side.
function Invoke-ComRetry {
    param(
        [Parameter(Mandatory)] [scriptblock]$Action,
        [int]$Attempts = 3
    )
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return & $Action
        } catch {
            if ($attempt -eq $Attempts) { throw }
            Write-Host "  COM call rejected (attempt $attempt), waiting..."
            Start-Sleep -Seconds ($attempt * 2)
        }
    }
}

# One assertion line. Prints "  OK    msg" or "  FAIL  msg" (run-tests.ps1
# fails a script on any FAIL line) and records the failure in $script:anyFail
# for scripts that also want an explicit non-zero exit at the end.
function Check([bool]$cond, [string]$msg) {
    Write-Host ("  {0}  {1}" -f $(if ($cond) { 'OK  ' } else { 'FAIL' }), $msg)
    if (-not $cond) { $script:anyFail = $true }
}

# 1-based position of the ListObject column called $name, or 0 if absent.
function Col($lo, $name) {
    for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
        if ($lo.ListColumns($i).Name -eq $name) { return $i }
    }
    return 0
}