# Setting up the workbook

## The short way (Windows, with Excel installed)

From the `PrintJob` folder:

```
powershell -ExecutionPolicy Bypass -File build.ps1
```

`build.ps1` does steps 2 to 5 below in one go: it backs up any existing
`.xlsm`, imports every module into a fresh copy of the `.xlsx`, pastes in the
`ThisWorkbook` code, runs `InitialiseWorkbook`, and saves `src\PrintCosts.xlsm`.
Step 1 is still needed once, and Excel must be closed on that file.

It prints whatever `InitialiseWorkbook` would have shown in a dialog, because
it puts the workbook in quiet mode first (`modUtils.SetQuiet`) - an unattended
run has nobody to click OK. Quiet mode also makes `Ask` answer No, so a script
can never confirm a destructive operation on your behalf.

The manual route below still works, and is the only route on Mac.

---

## The long way

The `.xlsx` carries everything a file can carry — sheets, tables, formulas,
validation, formatting, named ranges. What it cannot carry is code and drawn
controls, because a `.xlsm`'s VBA project is a binary structure that can only
be written by Excel itself. These steps take about five minutes and are needed
once.

## 1. Allow macros to be imported (Windows only)

**File → Options → Trust Center → Trust Center Settings → Macro Settings**, and
tick **Trust access to the VBA project object model**. Untick it afterwards if
you prefer; it is only needed for the import.

On Mac there is no equivalent setting and nothing to change.

## 2. Import the modules

Open `PrintCosts.xlsx`, then open the VBA editor:

- Windows: **Alt + F11**
- Mac: **Tools → Macro → Visual Basic Editor**

In the editor, **File → Import File…**, and import these eighteen files. Order
does not matter.

| File | What it is |
|---|---|
| `clsDict.cls` | Keyed store — replaces `Scripting.Dictionary`, which does not exist on Mac |
| `clsStock.cls` | One paper stock |
| `clsPrinterDef.cls` | One printer |
| `modUtils.bas` | Application state, messaging, table and name access |
| `modSettings.bas` | Typed access to the global settings |
| `modProtect.bas` | Sheet protection |
| `modCatalog.bas` | Configuration loaded once per operation |
| `modSnapshot.bas` | Freezes the rates each record is costed at |
| `modJobs.bas` | Add, stamp, remove, clear |
| `modValidation.bas` | Entry-time checks and the workbook sweep |
| `modLists.bas` | Dependent dropdowns |
| `modPicker.bas` | The multi-select picker |
| `modInit.bas` | One-time setup |
| `modMain.bas` | The button entry points |
| `modRegistry.bas` | Finds the print rooms, renames their tables, writes the reporting range |
| `modReports.bas` | Builds the Summary and Cost Calculations sheets |
| `modExport.bas` | Per-location CSV, and what has not been exported |
| `modVersion.bas` | Version identity, the About block, file properties |

The last four were added in build phases 4 to 7 and were missing from this
list until 0.7.1. If you are on a Mac this list is the whole build, so a
missing module is not a documentation slip - it is a project that will not
compile at step 3. Check you imported eighteen.

`ThisWorkbook.cls` is the exception: it **cannot be imported**, because every
workbook already has a `ThisWorkbook` object. Open `ThisWorkbook.cls` in a text
editor, copy everything from `Option Explicit` downwards, then in the VBA
editor double-click **ThisWorkbook** in the project tree and paste it in.

If any module is already in the project from a previous attempt, remove it
first: right-click it in the project tree, **Remove**, and answer **No** when
asked whether to export. Otherwise Excel appends a `1` to the name and you end
up with both.

## 3. Compile

**Debug → Compile VBAProject.** It should complete silently. If it stops on
anything, send me the message and the line — better to know now than after the
workbook is in use.

## 4. Run the setup macro

Back in Excel: **Alt + F8** (Mac: **Tools → Macro → Macros…**), choose
**InitialiseWorkbook**, and run it.

It draws the buttons over the placeholder labels, protects every sheet, forces
the 1900 date system, and tells you how many print rooms it found.

## 5. Save as .xlsm

**File → Save As**, and choose **Excel Macro-Enabled Workbook (.xlsm)**. Saving
as `.xlsx` silently discards all the code.

---

## Afterwards

**Run `InitialiseWorkbook` again whenever you add a print room.** Duplicating a
sheet copies its buttons but not always their macro links, and the new sheet
needs protecting. Re-running is safe at any time — it removes the buttons it
created before redrawing them.

**Clear the sample data** before live use. Each print room has a **Clear All**
button. Row `UNI-MAIN-00005` is deliberately invalid, to show the Status column
working.

**The sheet password is `printlog`**, set in `modProtect`. Change it there if
you want; it stops accidental edits, not determined ones.

## Adding a print room

1. Right-click a print room tab, **Move or Copy**, tick **Create a copy**.
2. Rename the new tab.
3. Click **Clear All** on it to discard the copied records.
4. Enter the room name, department and defaults; click **Select printers...**.
5. Click **Refresh Locations** on the Settings sheet.

Step 5 is the one that matters. Excel renames the copied table to something
like `tblJobs_ANNEX14` and gives the copy the original's location code;
Refresh renames the table deterministically, issues the copy a fresh code, and
rewrites the reporting range. The original always keeps its code, because its
job IDs already carry it.

## Scripts in the project folder

| Script | What it does |
|---|---|
| `build.ps1` | Builds the `.xlsm`. See the top of this file |
| `probe.ps1` | Structure of a workbook, read-only. No arguments: every sheet, table, column and defined name. `-Sheet Settings`: that sheet cell by cell, with its tables and buttons — use before placing anything new. `-File` to point it at the `.xlsm` |
| `verify.ps1` | Opens the built `.xlsm` **read-only** and reports version, document properties, tables, registry, the consolidated range and button bindings |
| `test-duplicate.ps1` | AT-13. Duplicates a print room in VBA, refreshes, reports, and closes without saving |
| `lockcheck.ps1` | Who is holding the `.xlsm`: Excel processes, owner lock files, and an exclusive-write test |
| `xlfnscan.ps1` | Every `_xlfn.` / `_xlws.` occurrence in the workbook's XML, including hidden sheets, conditional formatting and data validation |
| `test-reports.ps1` | Phase 5. Drives the Cost Calculations criteria through every AT-12 case and reports Summary and breakdown figures. Closes without saving |
| `filecheck.ps1` | Integrity of the `.xlsx`, `.xlsm` and every backup: size, VBA project present, sheet count, declared type, version stamp |
| `test-export.ps1` | Phase 6. Exports a print room, reads the CSV back, edits a row and confirms the fingerprint notices |
| `test-validation.ps1` | Phase 7. AT-11, AT-14 and AT-16 — the conflict warning, the Clear All confirmation, and inactive-record behaviour |
| `prune-backups.ps1` | Keeps the N most recent `*.bak.xlsm` and removes the rest. `build.ps1` calls it with `-Keep 5` |

**The tests are non-destructive by construction, not by care.** `test-validation.ps1`
reads the Clear All confirmation by putting the workbook in quiet mode first:
`Ask` then returns **False** and logs its prompt, so the exact text can be
checked while the deletion it guards never happens.

**Every test script drives a copy in `%TEMP%`, not `src\PrintCosts.xlsm`.** This
is the part that makes the sentence above true, and it was added in 0.7.1 after
the original mechanism turned out not to work.

Closing without saving is *not* a guarantee here, and neither is turning
AutoSave off. `Workbook_Open` does real work on every open — `ProtectAll`
unprotects and reprotects all fourteen sheets, plus `Invalidate` and
`HealButtons` — so the workbook is dirty the instant it opens. Excel holds this
file through its **cloud-backed handle** (`Workbook.Path` returns a
`d.docs.live.net` URL), so AutoSave is on and **commits that immediately**:
`$wb.Saved` reads **True** on the very next line after `Open`, despite every
sheet having just been modified. The bytes land when the handle closes — the
timestamp and the SHA-256 both move at `Close`/`Quit`, not at `Open`.

Which is exactly why `Close($false)` does not help. It means *discard unsaved
changes*, and by then there are none: AutoSave has already accepted them, so
what gets written at close is committed state. Setting `AutoSaveOn = $false`
after `Open` is too late for the same reason, and `Open` offers no earlier
hook — so the guard cannot work here by construction, and none is attempted.
`build.ps1` does set it, but only ever on a `%TEMP%` copy, which is not
cloud-backed — as its own comment notes, there it is usually a no-op.

**None of this is the sync client.** It happens with syncing paused and all
sync operations settled. It is Excel-side only, and `verify.ps1` has documented
the first half of it since phase 5, which is why that script and `probe.ps1`
open the workbook **read-only** — the other way to be safe. The four `test-*`
scripts need read-write to drive mutating macros, so a throwaway copy is what
is left. `diag.ps1` uses one too, for the same reason.

The consequence was quiet and worth knowing about: the phase 5–7 test runs were
rewriting the artefact they validated, so a later run read the earlier run's
edits back as though they were the build. Two scripts run back to back could
disagree about the figures with nothing wrong in the workbook at all.

If you are checking that for yourself, hash the file either side of a run:

```
$h = (Get-FileHash src\PrintCosts.xlsm -Algorithm SHA256).Hash
powershell -ExecutionPolicy Bypass -File test-reports.ps1
(Get-FileHash src\PrintCosts.xlsm -Algorithm SHA256).Hash -eq $h
```

It should print `True`.

## Exporting a print room

Each print room sheet has an **Export…** button. It writes
`PrintCosts-<SITE>-<LOC>-yyyymmdd-hhmm.csv` beside the workbook, holding
**every column including the frozen prices** each job was costed at — so the
file is a complete record, not just a readable summary.

Cell **B8** on each sheet says whether there are unexported changes, and
*Check workbook* and *Refresh Locations* list every print room that has any.

**Why this matters.** Excel raises no event when a worksheet is deleted, so
nothing can intercept someone right-clicking a tab and choosing Delete. The
records have to already be somewhere else. The status on B8 is a sign on the
cliff, not a fence — but it is where the user is looking when they reach for
that tab.

The "has anything changed" test is **derived, not tracked**: a fingerprint of
row count, cell count, quantity and chargeable totals, and the latest date,
compared against the one stored at the last export. A dirty flag would need
setting by every mutator, and would quietly lie in two cases nobody thinks to
test — a user typing into the sheet with macros disabled, and an old copy of a
sheet being dropped back in.

## The source .xlsx and OneDrive

**`build.ps1` never opens `src\PrintCosts.xlsx`.** It copies it to `%TEMP%`,
works on the copy, and copies the finished `.xlsm` back into `src\`.

That is not fastidiousness. The project lives in OneDrive, where AutoSave is
on by default, and a build necessarily leaves the open workbook heavily
modified — fifteen imported modules, setup run, version stamped — between
`Open` and `SaveAs`. AutoSave wrote all of that back over the source, and
`DisplayAlerts = $false` swallowed the "you can't save macros in a macro-free
workbook" prompt that would have stopped it. The result was a 410KB `.xlsx`
containing a VBA project, which Excel then refused to open at all. It was
recovered from OneDrive version history.

Three guards now make that failure loud rather than silent:

- the build refuses to start if the `.xlsx` contains a VBA project;
- it compares the source's timestamp before and after, and warns if the build
  changed it;
- `$wb.AutoSaveOn = $false` on the working copy.

Excel also **saves to `%TEMP%` and the result is copied into place**. Saving a
400KB workbook directly into an actively-syncing OneDrive folder is refused
outright, and the error — *"Unable to get the SaveAs property of the Workbook
class"* — looks like a missing method rather than a contested destination.

**On `_xlfn.` prefixes — do not strip them.** Functions newer than the file
format's baseline are *stored* prefixed (`_xlfn.LET`, `_xlfn.VSTACK`,
`_xlfn.HSTACK`, `_xlfn.SEQUENCE`, `_xlfn._xlws.FILTER`) and Excel hides the
prefix when it displays the formula. Removing a prefix from the XML makes Excel
stop recognising the function, which produces the `#NAME?` it looks like it
should prevent. `IFERROR` predates the baseline and is correctly stored without
one.

Excel also invents a hidden defined name, `_xlfn.<FUNC>` resolving to `#NAME?`,
for each prefixed token it meets. Those are runtime only and never saved —
`probe.ps1` marks them as such. They are not a defect and there is nothing to
remove.

`test-duplicate.ps1` injects its test code into the open workbook and never
saves, so nothing test-related reaches the deliverable.

`verify.ps1` opens read-only deliberately. `Workbook_Open` does real work every
time the file is opened — re-applies protection, heals the button bindings,
checks the date system — which marks the workbook dirty. Opened read-write, a
script that only *reports* on the file ends up rewriting it, moving its
timestamp and setting OneDrive syncing on every run.

**Never pipe these scripts into `Select-Object -First`.** It tears down the
pipeline as soon as it has enough lines, which kills the script before its
cleanup runs and leaves an Excel process holding the workbook open — OneDrive
then sits on "sync pending" while insisting it is up to date. `lockcheck.ps1`
is how you find that, and killing the orphaned process is how you clear it.

## What is not built yet

Phases 8 and 9 of the design document: visual polish and user notes, and the
full acceptance test run on Windows and Mac with its report.

The validation sweep (phase 7) is done. AT-11, AT-14 and AT-16 pass, and
nothing in the workbook needed changing to make them — the mechanisms were
built in earlier phases and this confirmed them.

Summary and Cost Calculations are built. Both are **live formulas** over the
consolidated range on `_Data` — `modReports` writes their layout and formulas
once and does no reporting at run time, so there is nothing to refresh and
nothing that can go stale.

Per-location export is built (`modExport`). Import is deliberately left to a
later revision; the format carries everything an import would need.
