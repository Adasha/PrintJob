# Builds src\PrintCosts.xlsm from src\PrintCosts.xlsx + src\PrintCosts-VBA\*.
#
# Replaces steps 2 to 5 of PrintCosts-VBA\SETUP.md. Requires Excel, and
# "Trust access to the VBA project object model" ticked in the Trust Center.
#
#   powershell -ExecutionPolicy Bypass -File build.ps1

$ErrorActionPreference = 'Stop'

$root = Join-Path $PSScriptRoot 'src'
$vba  = Join-Path $root 'PrintCosts-VBA'
$src  = Join-Path $root 'PrintCosts.xlsx'
$dst  = Join-Path $root 'PrintCosts.xlsm'

if (-not (Test-Path $src)) { throw "Not found: $src" }

# Checked before anything else happens. Excel holds a write lock on an open
# workbook, and without this the build imports every module, runs setup and
# stamps the version before failing on the very last step - and leaves a
# pointless backup behind.
function Test-Locked($path) {
    if (-not (Test-Path $path)) { return $false }
    try {
        $s = [IO.File]::Open($path, 'Open', 'ReadWrite', 'None')
        $s.Close()
        return $false
    } catch { return $true }
}
foreach ($p in @($src, $dst)) {
    if (Test-Locked $p) {
        throw "$(Split-Path $p -Leaf) is open in Excel (or otherwise locked). Close it and run this again."
    }
}

# Is the source actually a source? A plain .xlsx holds no VBA project. If it
# does, it is a macro-enabled workbook wearing the wrong extension - which is
# what the AutoSave bug below produced, and Excel refuses to open it at all.
# Better to say so plainly than to fail later with "the file format or file
# extension is not valid".
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zsrc = [IO.Compression.ZipFile]::OpenRead($src)
try {
    if (($zsrc.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' })) {
        throw "$(Split-Path $src -Leaf) contains a VBA project, so it is not a clean source. Restore it from OneDrive version history."
    }
}
finally { $zsrc.Dispose() }

if (Test-Path $dst) {
    $bak = Join-Path $root ('PrintCosts.' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.bak.xlsm')
    Copy-Item $dst $bak
    Write-Host "backed up existing xlsm -> $(Split-Path $bak -Leaf)"
}

# The build works on a COPY of the source, in the temp folder, and never opens
# the real .xlsx at all.
#
# This is not tidiness. The source lives in OneDrive, where AutoSave is on by
# default, and the build necessarily leaves the open workbook heavily modified
# - fifteen imported modules, setup run, version stamped - for the whole time
# between Open and SaveAs. AutoSave wrote all of that back over the source,
# and DisplayAlerts=$false silently swallowed the "you can't save macros in a
# macro-free workbook" prompt that would have stopped it. The result was a
# 410KB .xlsx containing a VBA project, which Excel then refused to open at
# all. Recovered from OneDrive version history.
#
# A build must never mutate its own input. Working on a copy outside the
# synced folder is what makes that true rather than merely intended.
$work = Join-Path ([IO.Path]::GetTempPath()) ('PrintCostsBuild-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item $src $work
$srcStampBefore = (Get-Item $src).LastWriteTimeUtc

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null

try {
    $wb = $xl.Workbooks.Open($work)
    # Belt and braces. AutoSave only applies to cloud-backed files, so on a
    # temp copy this is usually a no-op or unavailable - hence the try.
    try { $wb.AutoSaveOn = $false } catch { }

    # --- Trust Center check -------------------------------------------------
    try { $null = $wb.VBProject.VBComponents.Count }
    catch {
        throw "Cannot reach the VBA project. Tick File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model'."
    }

    # --- clear any previously imported components ---------------------------
    # Type 1 = standard module, 2 = class module, 3 = UserForm. Type 100 is a
    # document module (ThisWorkbook, sheets) and cannot be removed.
    $doomed = @()
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Type -eq 1 -or $c.Type -eq 2 -or $c.Type -eq 3) { $doomed += $c }
    }
    foreach ($c in $doomed) {
        Write-Host "removed  $($c.Name)"
        $wb.VBProject.VBComponents.Remove($c)
    }

    # --- import ------------------------------------------------------------
    $files = Get-ChildItem (Join-Path $vba '*') -Include *.bas, *.cls -File |
             Where-Object { $_.Name -ne 'ThisWorkbook.cls' } |
             Sort-Object Name
    if ($files.Count -eq 0) { throw "No .bas/.cls files found in $vba" }
    foreach ($f in $files) {
        $null = $wb.VBProject.VBComponents.Import($f.FullName)
        Write-Host "imported $($f.Name)"
    }

    # --- ThisWorkbook ------------------------------------------------------
    # It already exists in every workbook, so its code is pasted in rather
    # than imported. Everything from 'Option Explicit' down is the code; the
    # lines above it are the .cls header Excel writes on export.
    $twPath = Join-Path $vba 'ThisWorkbook.cls'
    if (-not (Test-Path $twPath)) { throw "Not found: $twPath" }
    $lines = Get-Content $twPath
    $start = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Option Explicit') { $start = $i; break }
    }
    if ($null -eq $start) { throw "No 'Option Explicit' line in ThisWorkbook.cls" }
    $code = ($lines[$start..($lines.Count - 1)]) -join "`r`n"

    $tw = $wb.VBProject.VBComponents.Item('ThisWorkbook')
    if ($tw.CodeModule.CountOfLines -gt 0) {
        $tw.CodeModule.DeleteLines(1, $tw.CodeModule.CountOfLines)
    }
    $tw.CodeModule.AddFromString($code)
    Write-Host "pasted   ThisWorkbook.cls ($($lines.Count - $start) lines)"

    # --- run setup ---------------------------------------------------------
    # Running a macro forces the project to compile, so a compile error
    # surfaces here rather than on the user's first click. Quiet mode stops
    # InitialiseWorkbook's completion dialog hanging an unattended run.
    $xl.Run('SetQuiet', $true)
    $xl.Run('InitialiseWorkbook')
    # Separate from setup on purpose: "Built" must mean when this file was
    # produced, and users re-run InitialiseWorkbook every time they add a
    # print room.
    $xl.Run('StampBuild')
    $log = $xl.Run('QuietLog')
    $xl.Run('SetQuiet', $false)
    $stamped = $xl.Run('VersionString')

    Write-Host ''
    Write-Host '--- InitialiseWorkbook said ---'
    Write-Host $log

    # --- save as .xlsm -----------------------------------------------------
    # 52 = xlOpenXMLWorkbookMacroEnabled. Saving as .xlsx discards the code.
    #
    # Retried, because Excel rejects COM calls outright while it is busy and
    # the report sheets give it plenty to be busy with - a full recalculation
    # of several dynamic-array formulas over every job row. A rejected call
    # surfaces as "Unable to get the SaveAs property of the Workbook class",
    # which reads like the method does not exist rather than like a server
    # that is simply not listening yet. Calculation is forced to manual first
    # so the save is not racing a recalculation it triggered itself.
    # Excel saves to a TEMP path, and the result is copied into place
    # afterwards. Saving a 400KB workbook straight into an actively-syncing
    # OneDrive folder is refused outright - every retry failed the same way -
    # and the error, "Unable to get the SaveAs property of the Workbook
    # class", suggests a missing method rather than a contested destination.
    # A plain file copy has none of that trouble.
    #
    # Calculation is forced to manual and CalculateBeforeSave off, so the save
    # is not racing a recalculation of the report formulas that it triggered
    # itself.
    $xl.Calculation = -4135      # xlCalculationManual
    try { $xl.CalculateBeforeSave = $false } catch { }

    $workDst = [IO.Path]::ChangeExtension($work, '.xlsm')
    $saved = $false
    for ($attempt = 1; $attempt -le 3 -and -not $saved; $attempt++) {
        try {
            $wb.SaveAs($workDst, 52)
            $saved = $true
        } catch {
            if ($attempt -eq 3) { throw }
            Write-Host "  save attempt $attempt refused, waiting..."
            Start-Sleep -Seconds ($attempt * 2)
        }
    }
    $wb.Close($false)
    $wb = $null
    Copy-Item $workDst $dst -Force
    Remove-Item $workDst -Force -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host "wrote $dst"
    Write-Host "       $stamped"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    if (Test-Path $work) { Remove-Item $work -Force -ErrorAction SilentlyContinue }
}

# Backups are build outputs and regenerable, so a deep history of them earns
# nothing and syncs a few hundred megabytes over time.
& (Join-Path $PSScriptRoot 'prune-backups.ps1') -Keep 5

# A regression guard for the bug that destroyed the source once already. The
# build must leave the .xlsx exactly as it found it; if that ever stops being
# true, this says so at the time rather than days later when Excel refuses to
# open it.
$srcStampAfter = (Get-Item $src).LastWriteTimeUtc
if ($srcStampAfter -ne $srcStampBefore) {
    Write-Host ''
    Write-Warning ("THE SOURCE .xlsx WAS MODIFIED BY THIS BUILD ({0:u} -> {1:u}). " -f $srcStampBefore, $srcStampAfter)
    Write-Warning 'That must never happen. Restore it from OneDrive version history before building again.'
}

Write-Host 'done'
