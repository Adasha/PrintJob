# Setting up the workbook

## The short way (Windows, with Excel installed)

From the `PrintJob` folder:

```
powershell -ExecutionPolicy Bypass -File build.ps1
```

`build.ps1` does steps 2 to 5 below in one go: it backs up any existing
`.xlsm`, imports every module into a fresh copy of the `.xlsx`, pastes in the
`ThisWorkbook` code, runs `InitialiseWorkbook`, and saves `src\PrintJob.xlsm`.
It also saves a copy named `src\PrintJob-v<version>-<yyyymmdd-HHmm>-g<sha>.xlsm`
(a trailing `+` on the sha means uncommitted changes); that is the file to copy to
another machine, so the name says which build it is. The newest five are kept.
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

In the editor, **File → Import File…**, and import these twenty-one files. Order
does not matter.

| File | What it is |
|---|---|
| `clsDict.cls` | Keyed store — replaces `Scripting.Dictionary`, which does not exist on Mac |
| `clsStock.cls` | One paper stock |
| `clsPrinterDef.cls` | One printer |
| `clsRestoreRoom.cls` | One print room's share of a restore |
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
| `modReports.bas` | Builds the Summary and Reports sheets; Reports-page delete |
| `modExport.bas` | Per-location CSV and Export All Locations; the Export report snapshot; what has not been exported |
| `modImport.bas` | Restores or merges an exported file into a print room's job table |
| `modBackup.bas` | Backup workbook / Restore workbook (the catalogue and settings CSVs); Restore creates any print room the backup has and this workbook lacks, and can read an older copy of the workbook instead of the CSVs |
| `modVersion.bas` | Version identity, the About popup, file properties |

If you are on a Mac this list is the whole build, so a missing module is not a
documentation slip - it is a project that will not compile at step 3. Check you
imported twenty-one.

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

**Use the Settings sheet's Add print room... and Remove print room... buttons
to change the set of rooms** (see "Adding a print room" below). If you add a
room by hand instead, click **Refresh Locations** afterwards.

**To upgrade to a new version, open the new workbook, click Restore
workbook... on the Settings sheet and choose your old workbook.** It brings
across the catalogue tables and Settings values, creates each of your print
rooms with its settings (default technician, printer and paper, permitted
printers, roll length unit) and imports their job records. Nothing needs
backing up first and the old file is not changed. Any one file from a CSV
backup works too. The new workbook's sample print
room is left as it is; remove it with Remove print room... when you no longer
want it. Running `InitialiseWorkbook` again is safe at any time: it re-draws
the buttons and re-applies protection.

**Each print room has a column view and a cost-detail toggle on row 2**, above
the job table. The drop-down offers All, Reduced and Minimal; Reduced hides the
columns listed in `SET_LOC_REDUCED_COLUMNS` and Minimal hides those plus
`SET_LOC_MINIMAL_COLUMNS` (both editable on Settings). Changing it changes
every room together. The **Hide/Show cost detail** button beside it toggles the
Paper Cost to Disregarded columns, again for every room.

**Roll length unit is a per-room setting** (`Roll length unit` in the room's
header block, Metres or Centimetres). Qty is typed in that unit and the Unit
column says which; costs, reports and the Summary always run in metres.
Changing the setting converts the roll lengths already entered. One room cannot
mix the two units job by job.

**The Paid column on the Reports sheet can be edited directly**, one cell at a
time. **Mark all as Paid / Unpaid** changes every record the filters currently
show, and it, like Delete visible records, refuses to run until at least one
filter is set.

**Clear the sample data** before live use. Each print room has a **Clear All**
button. Row `UNI-MAIN-00012` is deliberately invalid, to show the Status column
working.

**Sheets are protected without a password.** That stops accidental edits to
locked cells, not determined ones. Set a password by hand in Excel (Review >
Protect Sheet) on a particular workbook if you need one - nothing in
`modProtect` generates one by default any more.

**Exports write beside the workbook by default.** Set the "Export folder"
value on the Settings sheet to send CSV exports and report snapshots
somewhere else instead - it only takes effect if that folder already exists
on the computer running Excel.

**The Summary sheet's "Hide settings sheets" button** hides Print
Technicians, Printers, Papers and Settings from the tab bar (not full Excel
protection - Unhide still reaches them). Click it again, now labelled "Show
settings sheets", to bring them back.

**Add row/Remove row/Clear table buttons** sit above the Printers, Papers and Print
Technicians tables. Removing a row never affects a print job already
recorded against it - every job snapshots the price it was costed at.
**Clear table** empties that one sheet's table (after a warning naming what
will be deleted) so a bureau can start that list afresh; the other sheets are
not touched and it cannot be undone, so use Backup workbook first if unsure. The
smaller lookup tables on the Settings sheet (paper types,
standard sizes, consumables) don't have their own buttons; their cells are
unlocked, so a row added the normal Excel Table way (Tab at the last cell,
or right-click > Insert > Table Rows) keeps its formatting and validation.

**Review Printers' Max roll width mm / Max sheet size after upgrading from
before the printer/paper compatibility rework (`docs/HISTORY.md`
§16.5).** The first `InitialiseWorkbook` run after upgrading auto-derives
these two fields from whatever the printer's old `Supported families`
setting already allowed - safe in that it preserves the workbook's existing
compatibility exactly, but the derived numbers reflect what the catalogue
already contained, not necessarily the printer's real physical specs.
Check them against each printer's actual maximum roll width and largest
supported sheet size before relying on the workbook to reject a genuinely
oversized job.

## Adding a print room

Click **Add print room...** at the bottom of the Settings sheet. It asks for a
room name, an optional department and a short code for the room's job IDs
(for example `PHOTO` gives `UNI-PHOTO-00001`), then makes a clean copy of an
existing room and refreshes the registry. The new room starts empty, in
Metres, with no printers selected, so pick its printers with **Select
printers...** afterwards.

**Remove print room...**, next to it, deletes a room on purpose. It refuses to
remove the last room, tells you how many records and which dates would be lost
and whether they have been exported, and only goes ahead once you type the
room's name back exactly. Export the room first if you may want the records.

The manual route still works and is the only one that does not depend on the
Settings buttons:

1. Right-click a print room tab, **Move or Copy**, tick **Create a copy**.
2. Rename the new tab.
3. Click **Clear All** on it to discard the copied records, and **Clear
   defaults** to discard the copied batch defaults.
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
| `test-reports.ps1` | Summary and Reports end to end: every AT-12 criteria case, the Technician/Printer/Paper Stock filters and sort-by-column, export-names prompt, minimum-columns view, Paid column, Summary buttons, and Export report (header block, single-value promotion, name blanking). Adds the Annexe fixture as a second location. Closes without saving |
| `test-layout.ps1` | Shipped column order, no outline groups, Notes visible, H_Issues/snapshot hidden, header block survives a reduced-view toggle, Summary/Reports still reconcile |
| `test-manualhide.ps1` | Columns hidden by hand (not via the toggle): table-anchored buttons leave a hidden column and side-panel buttons follow column 36 on the next click |
| `TestCommon.ps1` | Not a test: shared `Invoke-ComRetry`, `Check` (assertion line) and `Col` (column lookup), dot-sourced by the tests |
| `test-export.ps1` | Phase 6. Exports a print room, reads the CSV back, edits a row and confirms the fingerprint notices |
| `test-validation.ps1` | Phase 7. AT-11, AT-14 and AT-16 — the conflict warning, the Clear All confirmation, and inactive-record behaviour |
| `test-nextid.ps1` | The persisted Job ID high-water mark: deleting the top row must not reissue its ID, and the mark must survive RefreshLocations |
| `test-catalogids.ps1` | Site-prefixed technician/printer/paper IDs: allocated on Add row and on typed rows, never reissued after a delete, and a restore of another site's catalogue adds rows (renaming clashing names to `Name (SITE)`) instead of overwriting |
| `test-import.ps1` | Export All Locations, and Import restoring into origin and into a different room |
| `test-deletereports.ps1` | Export report (a static-value `.xlsx` snapshot) and the Reports-page bulk delete, including the audit log entry |
| `test-suppliedstock.ps1` | Printer/paper compatibility rework: "Supplied by student" paper stock — zero Paper Cost with normal Consumable Cost, Print Width mm/Sheet size required-field enforcement, and rejection when an entered width/size exceeds the chosen printer's capacity |
| `run-tests.ps1` | Runs every `test-*.ps1` in turn and prints one line per script; `-Only` takes name fragments. Logs land in `%TEMP%\testrun-<Label>` |
| `test-addroom.ps1`, `test-removeroom.ps1` | The Add print room and Remove print room commands |
| `test-reducedview.ps1`, `test-costcolumns.ps1` | The All / Reduced / Minimal column views and the Hide/Show cost detail toggle |
| `test-viewedge.ps1` | Column views meeting Refresh Locations, setup, Add print room, Import and bad Settings lists |
| `test-rollunit.ps1` | The per-room roll length unit (cm or metres) |
| `test-importunits.ps1` | Import converting roll Qty between a cm file and a metres room, and the reverse |
| `test-jobvalidation.ps1` | Job-table validation is bound to the right columns by name, nothing is left below the table, and a column move is repaired |
| `test-reportsnav.ps1` | Reports filters added in 0.10.25 (Has notes, Min/Max cost), Clear all filters, Go to record into both rooms, header buttons staying in place when a column is hidden, button names agreeing with their macros (and HealButtons after a rename), healing a Paid edit whose Change event was lost, and editing Paid under a Paid = No view |
| `test-reportsfilters.ps1` | The closed More filters group on Reports (outline, counter), Has a problem, Disregarded, Has notes in the closed group, Student-supplied paper including the built-in stocks, and hidden filters still counting for Clear all filters and the safeguard |
| `test-summaryformat.ps1` | Summary sheet formatting: the red "(not in Papers)" rule, the colour legend, and the currency symbol following `SET_CURRENCY` (renamed from `test-phase8.ps1`) |
| `test-statusnotes.ps1` | Status hover notes on problem rows, the narrow Status column, and the amber warning fill on every location sheet |
| `test-togglepaid.ps1`, `test-markpaid.ps1` | The Toggle Paid button on the Reports sheet, and Mark all as Paid / Unpaid with its filter safeguard |
| `test-backup.ps1`, `test-clearcatalog.ps1`, `test-newpaperrow.ps1` | Backup/Restore workbook, Clear table on the catalogue sheets, and Add row on Papers |
| `test-restorerooms.ps1` | Restore creating a print room the backup has and the workbook lacks (code, name, department, job records), and a second restore not creating another |
| `test-restoreworkbook.ps1` | Restore workbook from an older copy of the workbook: room, settings and records carried over, the older file read without its macros running, left untouched, no temporary copy left behind, a non-Print-Cost workbook refused |
| `prune-backups.ps1` | Keeps the N most recent `*.bak.xlsm` and removes the rest. `run-tests.ps1` calls it with `-Keep 1` after a full, all-green run (`build.ps1` no longer prunes) |

**The tests are non-destructive by construction, not by care.** `test-validation.ps1`
reads the Clear All confirmation by putting the workbook in quiet mode first:
`Ask` then returns **False** and logs its prompt, so the exact text can be
checked while the deletion it guards never happens.

**Every test script drives a copy in `%TEMP%`, not `src\PrintJob.xlsm`.** This
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
open the workbook **read-only** — the other way to be safe. The `test-*`
scripts need read-write to drive mutating macros, so a throwaway copy is what
is left.

The consequence was quiet and worth knowing about: the phase 5–7 test runs were
rewriting the artefact they validated, so a later run read the earlier run's
edits back as though they were the build. Two scripts run back to back could
disagree about the figures with nothing wrong in the workbook at all.

If you are checking that for yourself, hash the file either side of a run:

```
$h = (Get-FileHash src\PrintJob.xlsm -Algorithm SHA256).Hash
powershell -ExecutionPolicy Bypass -File test-reports.ps1
(Get-FileHash src\PrintJob.xlsm -Algorithm SHA256).Hash -eq $h
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
then sits on "sync pending" while insisting it is up to date. Listing Excel's processes (`Get-Process excel`)
is how you find that, and killing the orphaned process is how you clear it.

## What is not built yet

The phase 10 manual run: the acceptance tests that depend on a person typing
(AT-01 to AT-09), and AT-15, which needs a Mac, have never been run by hand.
The scripted ones pass. Phase 8 (visual polish — the Summary legend,
conditional formatting for warning/error states, and currency wired to the
global setting) is built; see `docs/HISTORY.md` §16.2. Also not built:
aggregation for an automated master workbook (`docs/ARCHITECTURE.md` §12.2).
Export All Locations plus Import covers the manual route, and Export asks
whether to include student names.

The validation sweep (phase 7) is done. AT-11, AT-14 and AT-16 pass, and
nothing in the workbook needed changing to make them — the mechanisms were
built in earlier phases and this confirmed them.

Summary and Reports are built. Both are **live formulas** over the
consolidated range on `_Data` — `modReports` writes their layout and formulas
once and does no reporting at run time, so there is nothing to refresh and
nothing that can go stale. Summary is keyed on Location, Printer and Paper
Stock. Reports (renamed from Cost Calculations) adds Technician/Printer/Paper
Stock filters, sort-by-any-column, a hidden Job ID correlation
column, an "Export report" static-value `.xlsx` snapshot, and a bulk
"Delete visible records" and a "Mark all as Paid/Unpaid" pair, both confined to
whatever the active filters show and both refusing to run with no filter set.

Per-location export, Export All Locations, and Import are built (`modExport`,
`modImport`). Import restores a backup into the location it came from, moves
one room's exported jobs into another, or pulls several rooms' exports into
one copy for reporting - conflicts are resolved by Job ID (overwrite), and
only the input and snapshot columns are written so an imported row costs
correctly at the prices it was originally stamped with rather than being
re-derived from the target workbook's current catalogue.

Job IDs are allocated from a persisted high-water mark (`modRegistry.NextJobId`),
not a scan of the sheet's current rows, so deleting the highest-numbered job
cannot cause its ID to be reissued.

Technician, printer and paper IDs (`TechID`, `PrinterID`, `StockID`) are
allocated the same way, with the site code in front (`<SITE>-PRN-0001`), so
Restore workbook can combine another workbook's configuration with this one
without overwriting it (`modCatalog.NextCatalogId`). Nobody types them (the ID
columns are locked), and the number is four digits; the
counters are the `Read-only` rows `TECH_ID_HWM`, `PRINTER_ID_HWM` and
`STOCK_ID_HWM` on Settings. A restored row whose name is already taken by a
different row is kept and renamed `Name (SITE)`.
