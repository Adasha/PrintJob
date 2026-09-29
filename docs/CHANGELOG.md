# Changelog

Moved out of `modVersion.bas` (where it was ~600 lines of comment) so the module stays cheap to read. Newest first. Text is verbatim from the old header comment.

```
0.10.8 - Location sheets: sensible column widths, direct user request
("several columns are excessively wide").
  - modInit.ApplyJobColumnWidths now sets an explicit width on every column
    that was still using the PrintCosts.xlsx leftover: Technician 120px,
    Printer 160px, Paper Stock 170px (was 220), Sheet size 90 (was 115),
    Paper Cost 90 (was 115), Consumable Cost 100, Gross Cost 90,
    Disregarded 100 (was 115), Chargeable Cost 100 (was 220).
  - Untouched on purpose: Status (needs room for issue text), Notes (free
    text), Student Name/No, and the widths fixed in 0.9.14.
  - Re-applied on every setup/Refresh Locations run like the existing ones,
    so duplicated and older sheets pick it up. ARCHITECTURE.md column widths
    paragraph updated.
  - No schema change, no migration (workbook not released).

0.10.7 - Reports: Paid is directly editable, direct user request. The Paid
cells of the results table now take the same Yes/No dropdown as the room
sheets, and a choice is written straight to the job record it belongs to.
  - The results are one spilled formula: typing into a cell of it blocks the
    spill and blanks the whole row, so the record's Job ID/Location cannot
    be read once Change fires. modReports.RememberReportsPaidCell therefore
    notes which job the selected Paid cell shows (on selection and on sheet
    arrival, keeping the previous selection's note as well - Excel can move
    the selection before or after raising Change), and
    OnReportsPaidEdited (Workbook_SheetChange) applies the edit to that job
    and clears the typed constant so the spill returns.
  - Refused, with the constant cleared and nothing changed: a cell that is
    not a record's Paid cell, a value that is not Yes/No, a multi-cell
    paste/fill, a record no longer found, and a record whose Paid value has
    changed since the cell was selected (stale note). Re-picking the value
    already shown is a no-op.
  - Logged to the audit trail as "Paid edited (Reports)". Costs: the write
    clears Undo; one cell at a time (Mark all as... is the bulk route).
  - Paid cells (rows 16-2000) unlocked with Yes/No validation; the rest of
    the results table stays locked. ResnapReportsSelection is also called
    after Mark all as... and Delete visible.
  - test-paidedit.ps1 (new). ARCHITECTURE.md section 10.10.
  - No schema change, no migration (workbook not released).

0.10.6 - Reports "Mark all as..." Paid/Unpaid, direct user request, plus an
at-least-one-filter safeguard on the bulk commands.
  - New cluster in the Reports header, 1 column x 3 rows at O1:O3: label
    "Mark all as..." (O1), Paid (row 2), Unpaid (row 3). Paid sets the Paid
    column to Yes on every record currently shown, Unpaid to No. Records the
    filters hide are not affected, and the filters themselves are never
    touched. Asks for confirmation first (count, per-room breakdown, how many
    are already at the requested value).
  - Safeguard: neither Mark all as... nor Delete visible records will run
    unless at least one filter is set (modReports.HasActiveFilter /
    RequireActiveFilter). Mirrors Criteria(): spaces-only text, and a date or
    quantity that will not coerce (which Criteria ignores), do not count; Sort
    by/direction and Export names are not filters. Checked before any prompt,
    and again inside DeleteVisibleReportsConfirmed and
    MarkVisibleReportsConfirmed.
  - Job IDs are indexed once per room instead of rescanning per record.
  - Buttons: modInit (DrawOneAtTop, sized to their row), handlers
    btnMarkPaid/btnMarkUnpaid in modMain. Audit: "Mark visible Paid (Reports)"
    / "Mark visible Unpaid (Reports)".
  - test-markpaid.ps1 (new). ARCHITECTURE.md section 10.9; sections 8.5 and
    10.6 updated (the Reports buttons are in column T, not F, since 2026-09-26).
  - No schema change, no migration (workbook not released).

0.10.5 - "Export names", direct user request: the Reports toggle at O10
("Show names") now affects the exported report only, not the live view.
  - Live results: Student name/no always shown. The results formula in A16
    no longer references O10 (modReports).
  - Export report: modExport.BlankNameColumns blanks the Student name/no
    values unless O10 is Yes (default No, unchanged); the blank columns stay
    in the file. Applied before PromoteUniformColumns/RemoveExcludedColumns.
  - Label at N10 reads "Export names"; the dropdown's input message now
    says the results shown on the sheet always include names.
  - test-reports2.ps1: live view shows names under No and Yes; label check.
    test-reportsnapshot.ps1: export blank under No, populated under Yes.
  - No schema change, no migration (workbook not released).

0.10.4 - site-prefixed catalogue IDs, direct user request: TechID, PrinterID
and StockID are now allocated automatically, so configuration from another
workbook can be combined without overwriting anything. Before this nothing
filled them in (a new row's ID stayed blank) and the template's sample rows
shipped fixed IDs (TEC-001, PRN-001, STK-001), identical in every workbook.
  - IDs follow the Job ID pattern with the site prepended: <SITE>-TCH-0001,
    <SITE>-PRN-0001, <SITE>-STK-0001 (SET_SITE_ID, table code, four digits).
    modCatalog.NextCatalogId allocates from one persisted high-water mark per
    table (TECH_ID_HWM, PRINTER_ID_HWM, STOCK_ID_HWM in tblSettings, Notes
    "Read-only" so a restore skips them); the row scan is only a floor, so a
    deleted ID is never reissued. modRegistry.ScanMaxSuffix is now Public.
  - Filled by AddCatalogRow, by Workbook_SheetChange for a row typed under
    the table once it has a name (modCatalog.OnPaperEdited / OnCatalogEdited),
    and by Setup (EnsureCatalogIdSettings, EnsureCatalogIds) for any named row
    without one. An existing ID is never changed. The ID column is locked
    (modInit.UnlockTableBody) so it cannot be typed over; VBA still writes it.
    The number is four digits (Job IDs stay at five).
  - Template: the ID cells of the 18 sample rows in src\PrintCosts.xlsx (3
    technicians, 4 printers, 11 papers incl. the two legacy STK-SUP-* rows
    Setup deletes) are blank, so each build issues its own IDs. Edited in the
    sheet XML directly; no other cell changed.
  - Restore workbook (modBackup.ApplyCatalogRows): another site's rows now
    have different keys, so they are appended, not overwritten. A row whose
    NAME is already used by a different row (every lookup is by name, and
    clsDict.Add replaces on a duplicate) is kept under its own ID with the
    site added to the name - "HP T730 (SITE2)" - via RenameIfNameTaken; the
    result message counts them. Restoring the same backup twice is
    idempotent. SyncCatalogHwm then raises the counter to cover the restored
    IDs. No legacy-version migration: the workbook is not released yet.
  - New test-catalogids.ps1.

2026-09-29 - number-format fix on Printers and Papers, direct user request
(formatting only, no VBA change, no version bump). The Cost per m2 currency
format (GBP, 4 dp) had smeared onto Max roll width mm and Max sheet size.
  - Printers: Cost per m2 now GBP 2 dp; Max roll width mm #,##0; Max sheet
    size General. Fixed in src\PrintCosts.xlsx (cells AND the tblPrinters
    column dxfs, so new rows inherit the right formats) and in the built
    src\PrintJob.xlsm.
  - Papers: template was already correct. In the built PrintJob.xlsm, Width mm
    and Height mm were mixed 0 / #,##0 and the calculated Cost unit column was
    Text (@, which turns a re-typed formula into literal text); now #,##0 and
    General, matching the template.
  - The smear lived in the template, so every rebuild reintroduced it.
  - Printers: the Yes/No dropdown was on F11:F2001 (Max sheet size) instead
    of G (Active) below row 10; now G7:G2001.
  - Column widths tidied: Printers E/F 14.7 (E was 34.7, a leftover from the
    old Supported families column) and G 9.7 set explicitly; Papers StockID
    15.7 (STK-SUP-SHEET was clipped), Paper type 17.7, Measure/Std. size 11.7,
    Supplied by student 14.7.

0.10.3 - paper family removed, direct user request: on Papers the Family
column and the calculated Measure column always showed the same thing, since
the 0.10.0 rework left tblPaperFamilies with two rows (Sheet -> Sheet, Roll ->
Roll) and nothing but Measure was ever read. Measure is now the one input.
  - Papers: Family column and the calculated Measure column replaced by a
    single Measure dropdown (Sheet/Roll). Cost unit still keys off Measure.
  - Settings: tblPaperFamilies, its lstFamilies name and its +/- buttons
    (btnAddRowFamily/btnRemoveRowFamily) removed; the standard sizes and
    consumables tables moved left to close the gap. Backup All writes seven
    catalogue CSVs, not eight.
  - Job rows: S_Family snapshot column removed (S_Measure already records it);
    modUtils.SCHEMA_VER 1.2 -> 1.3. Exports from 1.2 still import - columns
    are read by header name, so the extra S_Family column is ignored.
  - Summary: Family column removed (Type, Unit, Jobs ... shift left one).
  - modCatalog reads tblPapers[Measure] directly (no family lookup); clsStock.
    Family and modCatalog.Basis removed.
  - No migration: no workbooks were in the wild. Rebuild from PrintCosts.xlsx.
  - Tests updated (test-newpaperrow, test-suppliedstock, test-phase8,
    test-reports, test-paid, test-backup).

0.10.2 - Clear table on Print Technicians, Printers and Papers, direct user
request: lets a bureau empty ONE of these lists to start it afresh without
touching the others.
  - modCatalog.ClearCatalogTable asks first (default No): names the table,
    row count and the first five rows, and warns it cannot be undone
    (suggests Backup workbook). Printers/Papers add a table-specific note.
  - modCatalog.DoClearCatalogTable is the unprompted core: audit entry,
    deletes rows 2..n, clears row 1 (formula cells skipped), Invalidate.
    Leaves one blank row, as modJobs.ClearAll does; AddCatalogRow reuses it.
  - modMain btnClearTechnicians/Printers/Papers; buttons drawn by
    modInit.InitialiseWorkbook next to Add row/Remove row.
  - Recorded jobs are unaffected (frozen S_* prices). Settings lookup tables
    do not get the button.
  - New test-clearcatalog.ps1.

0.10.1 - student-supplied stock made built in, direct
user request: the "Supplied (Roll)"/"Supplied (Sheet)" rows could be deleted
from the Papers sheet and were only restored by Setup, and nothing stopped
them being renamed or re-costed - yet the job-row status formulas match
their exact names. They are ever-present regardless of printer or location,
so they are no longer tblPapers rows:
  - modCatalog.LoadCatalog adds them in code (AddBuiltInStocks; PerJobSize
    = True on clsStock marks them), last in the dropdown. Their names are
    reserved (SUPPLIED_ROLL/SUPPLIED_SHEET, IsBuiltInStock); typing either
    into the Papers table is cleared with a message (OnPaperEdited).
  - modCatalog.RemoveLegacySuppliedRows (Setup) deletes the old rows from
    existing workbooks and writes an audit note. Recorded jobs are
    unaffected - they carry their own S_* snapshot. EnsureSuppliedStockRows
    and RestoreSuppliedStockRow are gone.
  - modReports.BuildSummary's Type/Family/Unit XLOOKUPs against tblPapers
    fall back to the fixed values for the two names instead of showing
    "(not in Papers)".
  - "Supplied by student" stays on tblPapers, now for user stocks such as a
    bulk delivery of paper a student supplies: it defaults to No (new rows,
    typed rows, and Setup fills blanks - NormaliseSuppliedFlags), and Yes
    means paper cost is always 0. Before, zero cost held only because the
    two shipped rows happened to have Cost = 0; LoadCatalog now forces
    Cost = 0 for any Yes row and OnPaperEdited resets the visible cell.
  - Behaviour change: the "size entered per job" rules (Print Width mm
    required, Sheet size required, compatible with any printer that takes
    that kind of stock) now apply only to the two built-in stocks, not to
    every Yes row - a Yes row uses its own catalogue size like any stock.
  - modCatalog.EnsureSuppliedColumnValidation re-applies the column's Yes/No
    list with corrected help text (the template's said Yes was only for the
    two shipped rows). PrintCosts.xlsx itself is unchanged and still ships
    the two rows; Setup removes them on build.
  - test-suppliedstock.ps1's "setup restores a deleted row" section is
    replaced with checks for the above.

0.10.0 - printer/paper compatibility rework, direct user request: testing
showed the family-based compatibility model was too coarse for real
printer/stock combinations - tblPaperFamilies was standing in for roll-
width bands ("Short Roll" <=24in, "Long Roll" >24in), and a printer opting
into "Sheet" accepted any sheet size at all, with no real numeric check
anywhere. Real new functionality, not just fixes, so the phase digit
moves.

Compatibility is now a genuine size fit rather than family membership:
  - tblPrinters gains Max roll width mm and Max sheet size (picked from
    tblStandardSizes) in place of the old Supported families multi-select
    - a printer takes roll stock iff the first is set, sheet stock iff
    the second is set, and both can be set on the same printer.
    modCatalog.Compatible checks the stock's own size against whichever
    applies (FitsWithinMaxSheet allows either orientation for sheet
    stock). modPicker.PickFamilies and its "Select families..." button
    are removed - clsPrinterDef.Families is gone with them.
  - modCatalog.MigratePrinterCapacities derives each printer's new fields
    from whatever it was compatible with under the OLD family-list rule,
    one-time and idempotent, before removing the legacy column - existing
    compatibility is preserved by construction rather than guessed at,
    though real printer specs are worth reviewing afterward (SETUP.md).
  - modCatalog.MigrateRollFamilies merges "Short Roll"/"Long Roll" into a
    single "Roll" family, since families no longer need one row per width
    band - only Sheet vs Roll matters to Compatible() now. Family/Measure
    stay as two separate columns regardless (Measure remains the only
    thing formulas key off, per the existing I2 rule); the duplication
    concern this was raised alongside is resolved by Printers no longer
    referencing family at all, not by removing the label.

A new "Supplied by student" option on Paper Stock, for when a student
brings their own paper: paper cost is always zero (real tblPapers catalogue
rows, Cost = 0 - "Supplied (Roll)"/"Supplied (Sheet)", modCatalog.
EnsureSuppliedByStudentColumn), ink/consumable cost is charged normally
from the printer's rate exactly as before (still waivable per-row via the
existing Disregard Consumable flag - no new mechanism needed there). The
size is entered per job rather than read from a catalogue row: Print Width
mm becomes required (not optional) for 'Supplied (Roll)', and a new job-row
column, Sheet size (a Standard Sizes pick), plays the same role for
'Supplied (Sheet)' - both validated against the chosen printer's own
capacity (modValidation.OnWidthChanged/OnSheetSizeChanged/
RevalidateSuppliedSize), and both required-if-blank checks are added to
the H_Issues calculated column so Check sheet/Check workbook catch a
forgotten one the same way they already catch every other required field
(modInit.EnsureJobIssuesFormula, new - asserted in VBA rather than a
static xlsx formula edit, the same self-healing reasoning
EnsureJobColumnValidation already applies to this table's validation).

Sheet size is a genuine new job-row column, so SCHEMA_VER (modUtils) moves
1.1 -> 1.2 - the second bump since inception, same shape as the Paid
column's own 1.0 -> 1.1 (§16.1/§3.5).

0.9.20 - Import button reunited with the rest of the side panel, direct
user report (2026-09-25): "the import button seems to have wandered over
to the right. it should be with the other right-hand-side buttons." It had
been paired horizontally with Export at a second column (0.9.17) because
every-other-row spacing only left six safe slots above the table header
for seven buttons - packing them tighter risked visual overlap, since row
height is a whole-row property shared with the main A/B block, so the
side panel couldn't just use taller rows without inflating that block too.

modInit.DrawOneAtTop (new) sidesteps the row grid entirely: takes an
explicit pixel Top instead of a row number, so all seven buttons now share
one column, spaced evenly across whatever room is actually available above
the table's current header row (computed at runtime, not a hardcoded
pixel figure) with an 8pt safety margin before it - not tied to any
particular row height, so it keeps working if the header moves again.
SIDE_PANEL_COL2, the now-unused second column, is removed.

0.9.19 - Field swap between the main block and the side panel, user
request (2026-09-25): Sheet status and Export move to the side panel
(AL/AM rows 5-6); the permitted-printers list moves the other way, onto
the main block (A7/B7), reformatted for people rather than for code - "N
printers (comma, separated, list)", singular handled. LOC_Printers itself
- the raw semicolon list the Select printers... picker writes and
modCatalog's compatibility checks read - stays exactly as it always was
functionally, just relocated to the side panel (AM7) under its own label
("Permitted printers (raw list)") so the picker's plain-value write is
unaffected; the friendly A7/B7 display is a separate, read-only LET
formula deriving from it (modInit.EnsurePrintersDisplay, new). Export's
target cell (modExport.EXPORT_CELL) moves from $B$8 to $AM$6 - a one-line
change, since RefreshExportStatus already writes its own "Export" label
via Offset(0, -1) rather than a hardcoded address. Print jobs (added
0.9.17) moves from A9/B9 up to A8/B8, the row Export vacated.

0.9.18 - Three snags from user testing of 0.9.17's layout rework, all
fixed same day (2026-09-25):

1. No breathing room between the configuration block (ends row 9) and the
   toolbar, which sat flush against "Print jobs". modInit.EnsureToolbarGap
   (new) inserts one blank row above the toolbar, called once per sheet
   from InitialiseWorkbook - narrower in scope than the EnsureJobTableGap
   0.9.17 removed (one row, not two), since only the toolbar needs
   separating from the block now. Toolbar moves from row 10 to row 11; the
   table's header follows naturally, from row 12 to row 13.

2 and 3, same root cause: "I can't access any of the drop down menus
(arrow appears but can't be clicked)" and "the buttons on the right
overlap the settings content". A Buttons.Add shape's own Width is
independent of the column it's anchored to - at SIDE_PANEL_COL's previous
(narrow, default) width, every 140pt-wide side-panel button sprawled
across two or three columns rightward, straight over the settings block
relocated to AL/AM in 0.9.17 - blocking its Yes/No dropdown arrows from
being clicked (the shape sits above the cell in z-order) as well as
simply looking wrong. Fixed by widening SIDE_PANEL_COL to comfortably
exceed the button's own width, so every button's footprint stays
contained within its one column - nothing placed to its right is ever at
risk again, not just AL/AM today.

0.9.17 - Reduced-clutter-view visibility fixes, direct user report
(2026-09-25): toolbar buttons and header-block content both vanished when
the columns beneath them were hidden by reduced view - the same
"Excel's column-hide is whole-column" constraint §4.1 had already flagged
as an unsolved general risk (2026-09-22), now recurring in practice.

Buttons: occasional-use location buttons (Remove Row, Select printers,
Check this sheet, Clear All, Export, Import, the reduced-view toggle
itself) move to a side panel (modInit.SIDE_PANEL_COL/SIDE_PANEL_COL2, past
the table's own columns) immune by construction. The few still anchored
inside the table's column span (Add Print Job, Now, Clear defaults)
self-relocate off whatever column is currently hidden via the new
modInit.RelocateAtRiskButtons, called every time column visibility can
change - general protection against SET_LOC_REDUCED_COLUMNS being edited
to name a column one of them happens to sit on.

Header-block content: LOC_RollUnit's dropdown and the "Default: paper"
label both sat on column E (table column 5, Printer - one of the default
hidden columns), vanishing out of the box with no custom editing needed.
Cell content can't float to a different column the way a button can, so
the fix is structural: everything that must stay visible - room name,
department, the three per-job selectors, roll unit, sheet status, export
status, and a new live job count - now lives only in columns A/B, which
never appear in any hide list. Location code, the two disregard-cost
defaults and the permitted-printers list (not used per-job) moved to the
side panel to make room. modInit.EnsureJobTableGap is removed entirely -
the selectors no longer need a dedicated row, so there is nothing left to
open space for, and the table header reverts to its originally-shipped
row 12.

A third bug surfaced by that same reversion: modInit.ReorderJobColumns'
Range.Cut/Range.Insert Shift:=xlToRight operates on the table's whole
header+data row span, not just its columns - with the header back at row
12, row 13 became the table's first data row, exactly where the side
panel's Import button had landed, and the next reorder silently dragged
it sideways. Fixed by keeping every side-panel row at 11 or above.

0.9.16 - Location-sheet batch defaults, direct user report (2026-09-25 -
reproduced on Mac, not Windows): picking an incompatible printer/paper
default pair off the dropdown left the raw "(unavailable)"-suffixed value
sitting in the cell, and Add Print Job copied it straight into the new
row uncaught. modValidation.OnDefaultCellChanged already runs the AT-05
compatibility check the instant a default cell changes, but that depends
on Worksheet_Change firing - which Excel for Mac does not reliably do for
an in-cell dropdown selection. modValidation.EnsureDefaultsClean is a new
defensive re-check, called from modJobs.AddPrintJob right before it copies
the defaults into the row, so a stale/incompatible default is caught and
cleaned regardless of whether the change-time event fired.

0.9.15 - Job-table Data Validation corruption fix, direct user report
(2026-09-25): every column from Unit through Chargeable Cost was showing a
dropdown to pick a technician's name instead of its own content, and
picking one overwrote a formula cell, breaking that row.

Root cause, confirmed against the built .xlsm over COM: modInit.
ReorderJobColumns (0.9.4) moves columns via repeated Range.Cut +
Range.Insert Shift:=xlToRight inside the table. Excel extends a validated
column's rule onto cells newly shifted in beside it on each Insert, and
the ~30 inserts one full reorder performs compound that into a wide,
wrong span - Technician's own list, meant for one column, ended up the
Formula1 on eleven; the Date/Time rule similarly ended up on Student Name
and Student No. The corruption happened once, when 0.9.4 first reordered
an existing workbook's columns, and then sat there undisturbed - the
reorder's own "already in the right position" check meant it never ran
the Cut+Insert again to have a chance of un-corrupting it.

modInit.EnsureJobColumnValidation (new) clears the whole table body's
validation and restores exactly what PrintCosts.xlsx ships for Date/Time,
Qty, Print Width mm, Disregard Paper, Disregard Consumable and Paid.
Called from modLists.BindColumns - the existing "rebuild dependent
dropdowns, self-heal, never trust what a previous run left behind" entry
point that every setup run, Refresh Locations, Check workbook/sheet and
the picker already go through - so the fix reaches an already-corrupted
workbook on the next ordinary refresh, not only a full rebuild.

0.9.14 - Five direct user-feedback items, layout and naming only, no schema
change:

  - "Quantity" renamed to "Qty" everywhere it is a table/column header (the
    job table on every location sheet, the Reports results table and its
    Sort-by list, Summary's breakdown table, the export column order/
    header row) - narrower headers were costing width the already-tight
    job-row columns needed elsewhere. modInit.EnsureQtyColumnName renames
    the live ListColumn (checked-first, same shape as EnsurePaidColumn)
    rather than hand-editing PrintCosts.xlsx; Excel updates every
    structured reference (Area m2/Paper Cost's own formulas, _Data's
    consolidated header copy) as part of the rename itself. Every VBA call
    site that looked the column up by its old name - modReports.C/SumBy,
    modExport.ExportColumns/Agg, modImport.WriteImportedRow - now asks for
    "Qty". Not a schema bump (§3.5: a rename doesn't change what a column
    MEANS, only its label) - a file exported before the rename degrades
    the same graceful way a pre-Paid export already does (missing key
    reads back blank, not an error). The Reports page's OWN "Quantity"
    filter-box label (F8) is deliberately left alone - a criteria label,
    not a table header, with no width pressure on it.
  - "Standard size" (tblPapers) renamed to "Std. size" the same way -
    modCatalog.EnsureStdSizeColumnName - but with no call sites to fix,
    since nothing in this codebase looked that one up by name.
  - Location-sheet job-row columns narrowed to fit more of the table in
    one screen: Unit and Qty capped at 50px, Area m2 at 60px, Paid at
    40px; Job ID and the Reports page's own Date/Time column fixed at
    exactly 100px. modUtils.ColWidthForPx converts a pixel target into
    Range.ColumnWidth's own "characters of the Normal-style font" unit
    using Calibri 11's well-known metric (7px/char + 5px padding, the same
    constant openpyxl/xlsxwriter use) - there is no Range.WidthInPixels to
    set directly. Always rounds down, so a cap can never render one pixel
    over.
  - Summary's own totals-breakdown header (row 9) now shares the Reports
    page's warm totals tint (RGB(244, 232, 222), added in 0.9.13 for
    Reports' own "Matching" row) instead of the blue it used to share with
    Reports' results-table header - this table is itself a totals
    breakdown (by location/printer/paper stock), the same category as
    "Matching", not a per-job record list.
  - A blank row now sits both above and below the location-sheet batch-
    defaults row (modInit.EnsureJobTableGap, the same "insert a row above
    a table, checked first" idiom EnsureTableGap already used for the
    catalogue tables) - it used to run straight into the main toolbar
    below with no gap at all, and read as crowded against the config
    block above despite row 8 already being blank there. Shifts the
    defaults from row 9 to row 10 and the toolbar from row 10 to row 12;
    every hardcoded row reference in modInit (EnsureJobDefaults,
    DrawLocationButtons, the reduced-view toggle/Clear defaults buttons)
    moved with it. Named ranges (LOC_DefTech/LOC_DefPrinter/LOC_DefPaper)
    needed no separate migration - Excel's own row-insert already carries
    a name's RefersTo along with whatever physically moved.

0.9.13 - Two small Reports-page fixes, direct user feedback:

  - Matching's totals header (row 12) now gets its own subtle tint,
    RGB(244, 232, 222) - the results header's own RGB(222, 232, 244) with
    red and blue swapped, so the two bands read as related but distinct
    rather than the totals row looking unstyled next to it.
  - "Show student/department name/no" (the toggle at what was E9/F9) had
    no visible label of its own reason to be there - it is a display
    option, not a filter criterion, and shared row 9 with the unrelated
    name/number mismatch warning purely because that row happened to be
    free. Moved to a third group on the sort-controls row (I10/J10),
    after Sort by/Sort direction, with its own label restored. Every
    formula and comment that referenced $F$9 now reads $J$10 instead;
    nothing outside modReports.bas touched that cell.

0.9.12 - Settings-page layout: the About section (modVersion.WriteAbout,
removed) is gone from the bottom of the Settings sheet, per direct user
feedback that it duplicated the About popup for no benefit while eating
vertical space. The eight action buttons that used to run down column T -
out of the way, but also out of sight on a normal-width window - now sit in
the freed space instead, directly below whichever table on the sheet runs
deepest (modInit.TablesBottom, the same "bottom" WriteAbout used to compute
before clearing its own text). Two logical groups of four, one row each:
Refresh Locations/Check workbook/Re-stamp prices/About, then Export All
Locations/Import/Backup workbook/Restore workbook. modInit.InitialiseWorkbook
also clears the old About block's text area unconditionally on every setup
run, not just a from-scratch rebuild, so a workbook built before this change
loses the stale text the next time setup runs on it.

ShowAbout (the popup itself, unchanged in structure) now carries the brief
description that used to live only in the About block's prose, so the
popup alone covers project name, version, author and what the workbook
does - nothing is lost by removing the on-sheet block.

0.9.11 - Printer/Paper Stock dropdowns no longer hide incompatible options,
per direct user feedback that the bidirectional narrowing added in 0.9.0
was "too restrictive." Every active/permitted item is now always listed;
whichever ones the other field's current value rules out are suffixed
" (unavailable)" (modLists.UNAVAILABLE_SUFFIX) instead of being removed.
Picking one is allowed - modValidation.Clean() strips the suffix and hands
the clean name to the EXISTING AT-05 "incompatible combination" handling
(OnPrinterChanged/OnStockChanged/OnDefaultCellChanged), which already
clears the other field and explains why, unchanged from 0.9.0. Auto-fill
when exactly one compatible option remains is UNCHANGED - it now runs off
a narrowed list computed only for that decision, never shown to the
dropdown itself, so 0.9.0's tested auto-fill behaviour (test-dropdowns.ps1)
still holds.

Known compromise, not a full fix - flagged for revisiting rather than
presented as settled: Excel's native in-cell dropdown (Data Validation)
cannot style individual list entries - no italics, colour or shading on
one item within the list, which is what was actually asked for. A text
suffix is what is achievable inside a real Excel dropdown, and is judged
good enough for now - the workbook stays with native Excel styling
throughout rather than building a bespoke picker UI just for this. See
docs/ARCHITECTURE.md §7.2 / §16.3.

0.9.10 - Department/"charge to" terminology (snag 4b, D19): free text,
no new catalogue table, per the user's own decision on review. The
Student Name/No columns already accept anything typed into them (D1 -
no authoritative student list, cross-checked only for consistency) - a
department name behaves identically with no code change needed. Only the
user-facing WORDING that names the field got a light "Student/
Department" pass: the Reports filter labels and autocomplete titles, the
"Show student/department name/no" toggle, the Remove-row confirmation
("Student/Department: "), the name/number consistency warning, and the
About blurb. The underlying column headers stay "Student Name"/"Student
No" deliberately - renaming those would break every header-name lookup
that already reads them (§5, §8.3, §10.4 of the architecture doc) for no
real capability gained.

0.9.9 - Full workbook backup/restore (snag 4a). New modBackup.bas: Backup
workbook writes one CSV per catalogue table (Technicians, Printers,
Papers, the four Settings-page lookup tables, and Settings itself)
alongside the existing per-location job exports, all sharing one
timestamp; Restore workbook reads any one file from a backup back in,
finds every sibling sharing the same timestamp, and applies the whole
set - overwrite-by-stable-key for catalogue rows (mirroring Import's own
Job ID rule), and modImport.ApplyImportConfirmed unchanged for job
records. Two new Settings buttons: "Backup workbook..." / "Restore
workbook...".

One genuine, previously-latent bug found and fixed at the root: VBA's And
does not short-circuit, so `Count = 1 And IsBlankRow(lo, 1)` - a pattern
already used in four other modules - still called IsBlankRow (indexing
ListRows(1)) even when Count had already evaluated to 0, raising
"Subscript out of range". Dormant until now because nothing in ordinary
use ever left a table at genuinely zero rows; restoring into a table
emptied by hand is the first path to actually hit it. modUtils.IsBlankRow
now returns False for an out-of-range row instead of indexing blindly,
fixing every call site at once.

0.9.8 - Summary sheet fixes (snags 3a, 3b). ToggleConfigSheets (the
Hide/Show settings sheets button) now captures the Summary worksheet
explicitly and re-activates it unconditionally after hiding/showing the
four configuration sheets, rather than relying on Summary having stayed
active throughout - hiding whichever sheet happens to be active is what
forces Excel to jump to the next one in tab order, so this is correct
regardless of how ToggleConfigSheets was actually invoked. The "Go to
Settings" button is removed outright rather than made target-aware: with
four configuration sheets and no way to tell which one a task needs, it
could only ever jump to one of them, and Hide/Show settings sheets plus
Excel's own tabs already reach all four once visible - the snag list's
own "easiest fix."

0.9.7 - Export report's header block and single-value promotion (snags
2b, 2c). modExport.ExportReportSnapshot's .xlsx archive now opens with a
metadata block - report title, schema version, site ID/name, the date
range covered (the From/To filter boxes when set, else the actual spread
of the exported rows' own Date/Time column), a generated-at timestamp and
the row count - then a blank row, then the results table. Live Reports
sheet untouched throughout, per the user's own review answer: this is
export-time-only.

Single-value promotion: any of Student name, Student no, Location,
Printer, Paper stock, Technician that holds the SAME non-blank value on
every exported row is lifted into the header block as a "Header: Value"
line and dropped from the table, so filtering a report down to one
printer (say) states that once at the top instead of repeating it down an
entire column. A candidate column with no non-blank values at all (e.g.
Student name/no when the 2a toggle is off) is left alone untouched -
there is nothing true to promote.

One real bug found building this: LastVisibleColumn (renamed
LastHeaderColumn) filtered its column bound by .Hidden, which was
harmless when only the trailing Job ID column was ever hidden - but 2d's
minimum-columns view (0.9.6) now hides Printer and Technician among
others, so the OLD function silently truncated every export at the last
VISIBLE column, dropping Technician and Notes from the file entirely with
no error. The promotion logic needs those hidden columns' values to check
for uniformity, so the bound is now by header name (stop before the
trailing "Job ID") rather than by Hidden state.

0.9.6 - Reports student name/no toggle and minimum-columns view (snags
2a, 2d): a Yes/No toggle at F9 (default No) blanks Student Name/No in the
results table and any export, satisfying data protection; Student Name/
No and Paid added to the results table; only seven columns (Date/Time,
Location, Student Name, Student No, Paper stock, Chargeable, Paid) stay
visible by default, the rest hidden via modReports.ApplyReportsMinimum
Columns (not user-facing yet - documented future enhancement).

Four real bugs found and fixed while testing edge cases (blank-everything,
nothing-paid-yet), not just the happy path:
  - The relaid-out "Matching" totals row (MatchTotal, label above value
    rather than to its left) initially still shared columns with newly-
    hidden results columns, so several totals displayed blank - same
    class of bug as the location-sheet header-block collision, now fixed
    by only ever using columns the minimum-columns view never hides.
  - SUM(FILTER(...))-SUM(FILTER(...)) for the Unpaid total returned 0
    whenever nothing was marked Paid yet - FILTER with zero matches
    raises #CALC!, which poisons the whole subtraction and gets masked
    by the outer IFERROR into a false 0. Fixed with FILTER's own
    [if_empty] third argument.
  - IF($F$9="Yes", <column array>, "") broke every results row past the
    first once the toggle was off: the condition is a scalar (one toggle
    cell), so the IF collapses to the bare scalar "" rather than a
    column of blanks - HSTACK does not broadcast a scalar against a
    column, the exact trap _Data's own consolidated-range formula (§7.1)
    already had to work around, met again in a new place. Fixed by
    nesting the toggle inside an outer IF whose own condition is already
    a genuine per-row array.
  - Paid displayed a literal "0" for any row that predates the column
    (blank Paid counts as unpaid, design 5) - INDEX on a truly blank
    cell returns the number 0, and &"" does not fix an already-numeric
    0 (0&"" is still the text "0"). Fixed by testing for that specific
    value, safe here since Paid never legitimately holds a real 0.

Also found (and fixed) a TEST-SCRIPT bug that looked exactly like
intermittent COM automation flakiness: a PowerShell [char] handed to
Excel's COM Columns(...) indexer is read by its ORDINAL VALUE, not as a
one-letter column reference - .ToString() first fixes it. Cost real
debugging time because the symptom (a genuinely hidden column reading
back as visible) is indistinguishable from a real settling/timing issue
until compared side by side against the literal column letter.

0.9.5 - spotted while visually checking 0.9.4's fix: two extra column
groups next to the cost group, one of them hiding a genuine bug.
GroupJobColumns used to ALSO group S_PrinterID->S_SchemaVer, on top of
the grouping PrintCosts.xlsx already ships for that exact span - nesting
a second outline level over the first rather than reusing it. Visibly:
two unwanted extra groups cluttering the outline pane next to the one
group that matters day to day. Substantively: Notes ended up hidden - a
genuine input column, never meant to be, silently swept into the
snapshot block's hidden state as a side effect of EnsurePaidColumn (0.9.2)
inserting Paid directly next to it. modInit.FlattenOutline (new) resets
Notes->S_SchemaVer to outline level 1 outright rather than re-grouping
it, and explicitly un-hides Notes; H_Issues/the snapshot columns keep
their own .Hidden state exactly as before.

0.9.4 - two bugs found testing 0.9.3's reduced-clutter view, both fixed
the same day:

  - Hiding a column hides the WHOLE column, every row, not just the
    table's - Status and Job ID used to be the table's first two columns
    (sheet A/B), the SAME columns the location header block above the
    table (room name, department, code, defaults) occupies in rows 1-9.
    Turning reduced view on blanked the header along with the table
    cells. modInit.ReorderJobColumns moves Status and Job ID to just
    after Paid, ahead of Notes/H_Issues/the snapshot block, so columns
    A/B are always Date/Time/Student Name - never anything reduced view
    can hide. Column order isn't load-bearing anywhere (design 5.1), with
    one matching fix: modRegistry's consolidated-range span was bounded
    by column NAME from "Job ID" to "Notes" - with Job ID no longer the
    leftmost column, that bound moved to "Date/Time" so the span still
    covers everything (Status/Job ID included, now inside the span
    rather than starting it).
  - The toggle button itself was anchored to column 13, which happened to
    be one of its own hidden columns (Disregard Consumable) - turning
    reduced view on could take out the one button that turns it back
    off. Moved to column 7, deliberately independent of the reorder
    above.

Neither is a general fix for the underlying constraint (Excel's
column-hide is whole-column-only, and the hidden-column list is
user-editable, so a future edit could collide with something else on the
sheet) - revisit properly later. Also flagged: the toggle's only feedback
is its own button caption; a clearer visual state indicator is worth
adding, deferred alongside the rest of the visual-polish work.

0.9.3 - reduced-clutter view toggle (snag 1e): a workbook-wide (not
per-sheet) toggle hiding Status, Job ID, Printer, Area m2, Disregard Paper
and Disregard Consumable on every location sheet at once, driven by two
self-provisioned settings (SET_LOC_REDUCED_VIEW, SET_LOC_REDUCED_COLUMNS)
rather than a hard-coded list, so the shortlist can change without a
rebuild. modInit.ApplyColumnVisibility is the general "hide/show exactly
these headers, touch nothing else" primitive underneath - deliberately not
a blanket show-everything-then-hide-the-list reset, since H_Issues is
hidden permanently and must never be touched by this. Same primitive gets
reused by the Reports page's own fixed minimum-columns view once that
lands (snag 2d).

0.9.2 - Paid job-row column and the column-grouping fix (snags 1c, 1d):

  - A new Paid (Yes/No) column, right after Chargeable Cost. The first
    genuine job-row column change since inception, so SCHEMA_VER (modUtils)
    moves 1.0 -> 1.1 - which surfaced that SET_SCHEMA had shipped as a
    static value in PrintCosts.xlsx with no code keeping it in sync with
    the constant, the same drift that caught APP_VERSION out below. Fixed:
    modVersion.EnsureSchemaSetting now stamps SET_SCHEMA from SCHEMA_VER on
    every setup run, same as EnsureVersionSettings already does for
    APP_VER. Existing rows are left blank rather than force-set to "No" -
    blank already reads as unpaid everywhere (Summary/Reports totals,
    export, import).
  - modInit.GroupJobColumns's cost-column group ran one column too far
    (Paper Cost through Chargeable Cost itself), so collapsing it hid the
    one figure a collapsed view most needs to keep showing. Now runs
    Paper Cost through Disregarded only; Chargeable Cost and the new Paid
    column stay outside it, always visible.
  - Summary and Reports both split their Chargeable total into Paid/
    Unpaid, Unpaid computed as "filtered Chargeable minus Paid" rather
    than a second criteria expression, so the two figures can never drift
    apart from the total they came from.

0.9.1 - batch default Technician/Printer/Paper selectors (snag 1b): three
cells above the toolbar (row 9) pre-fill Technician/Printer/Paper Stock on
every subsequently added job until Clear defaults empties them, using the
same copy-not-reference principle as the existing disregard-cost defaults
(AT-07/AT-08) and the same bidirectional filtering/autofill as 0.9.0's
table-cell dropdowns. modInit.EnsureJobDefaults self-provisions the three
named cells (LOC_DefTech/LOC_DefPrinter/LOC_DefPaper), same pattern as
LOC_Export; modLists.BindDefaultCells and modValidation.OnDefaultCellChanged
reuse 0.9.0's binding/compatibility-check machinery rather than duplicating
it.

0.9.0 - phase 9 begins: the 2026-09-22 post-phase-8 snag list. Real new
functionality, not just fixes, so the phase digit moves again:

  - Printer and Paper Stock on a location sheet now filter each other in
    BOTH directions, auto-filling either one when only a single compatible
    option remains. Previously Paper Stock was locked until a printer was
    chosen; both cells now start unlocked and fully populated.
    modCatalog.PrintersForStock is new (the reverse of the existing
    StocksFor); modLists.BindPrinterRange/BindStockRange/AutoFillIfSingle
    replace the old one-directional BindStockRange; modValidation's
    OnPrinterChanged/OnStockChanged are now symmetric, each rebinding both
    cells' lists rather than assuming only one drives the other. See
    docs/ARCHITECTURE.md §7.2 for the design and a note on the one real
    bug this surfaced (Range.Value on a multi-cell Union is an array, not
    a scalar - AutoFillIfSingle has to walk target.Cells, not read
    target.Value directly).

0.8.1 - phase 8's original scope, finally built (see the 0.8.0 note below,
which explicitly deferred this):

  - modSettings.CurrencySymbol (already reading SET_CURRENCY, previously
    unreferenced) is now read everywhere a NumberFormat or Format$ used to
    hardcode ChrW(163) - modReports (Summary and Reports), modExport's
    export-report snapshot, and modJobs.RemoveRow's delete-confirmation
    dialog. The NumberFormat sites go through the new
    modSettings.CurrencyFormatCode rather than CurrencySymbol() directly -
    verified directly against this workbook over COM: setting a Range's
    NumberFormat from VBA to a BARE or quoted currency symbol (e.g.
    "$#,##0.00") is silently canonicalised by Excel back to the OS's own
    regional currency symbol, no matter what SET_CURRENCY holds - invisible
    on a GBP-locale machine, where it happened to always match, but it
    would have silently defeated the setting everywhere else. Excel's
    "[$symbol]" bracket syntax round-trips exactly as given and was
    confirmed to render identically for the unaffected GBP case, so
    CurrencyFormatCode is a strict fix, not a behaviour change. Format$
    (RemoveRow's dialog) does not have this problem - confirmed separately
    - and keeps using CurrencySymbol() plain. Changing SET_CURRENCY still
    only repaints Summary/Reports on the next InitialiseWorkbook run, same
    as every other piece of their formatting.
  - Conditional formatting exists in the workbook for the first time.
    modInit.ApplyStatusFormat gives every location sheet's Status column an
    amber fill whenever a row reads anything other than "OK" - the
    "warning state" §5/§11 already called out as "conditionally
    formatted". modReports.FormatSummaryErrors gives Summary's Type/Family
    columns red bold text when a job references a paper stock that no
    longer exists in tblPapers ("(not in Papers)") - the "error state",
    matched to the one place the workbook's own formulas already flag a
    genuine data-integrity break rather than a routine per-row warning.
  - modReports.DrawLegend adds the Summary-sheet legend §11 planned and
    never built: the four everyday cell-role colours (input, calculated,
    configuration, read-only reference), at O9 down, in the colours
    already used for those roles elsewhere on Summary/Reports.

0.8.0 - NextId high-water mark, Export/Import, Reports rework, UI snag
list. Real new functionality, not just fixes, so the phase digit moves:

  - modRegistry.NextJobId now allocates from a persisted high-water mark
    (a new "Job ID HWM" registry column) instead of a scan of rows on the
    sheet, so deleting the highest-numbered job can no longer reissue its
    ID. Fixed first because Export/Import both depend on the "globally
    unique, never re-issued" invariant actually holding.
  - modExport gained Export All Locations (every room in one pass) and an
    Export report snapshot - a static-value .xlsx of the Reports sheet's
    current filtered/sorted view, with its own signature stamped on
    Reports for the delete command below to check against. Exports now
    resolve to a selectable folder (SET_EXPORT_FOLDER, falling back to the
    workbook's own location if unset or not present on this machine), and
    the OneDrive path fix now also searches
    ~/Library/CloudStorage/OneDrive* on Mac, where the Windows-only
    OneDrive environment variables never existed in the first place.
  - modImport is new: restores or merges an exported file into a
    location's job table. Rows match by Job ID - append if new, overwrite
    if not - and only input/snapshot columns are written, so an imported
    row costs correctly from its own frozen rates regardless of the
    target workbook's current catalogue.
  - Cost Calculations renamed to Reports: Technician/Printer/Paper
    Stock/Quantity filters, sort-by-any-column, a hidden Job ID
    correlation column, and a bulk "Delete visible records" command
    confined to whatever the active filters show. Summary regrouped to a
    Location x Printer x Paper stock key.
  - The 2026-09-21 snag list: no default sheet-protection password,
    unlocked configuration table inputs, freeze panes moved off the
    config sheets and onto Reports, tab order fixed, a hide/show toggle
    for the config sheets, column and row groups, Settings notes
    wrapped, and catalogue row add/remove buttons.

Not part of this bump: phase 8's original scope (the Summary legend,
conditional formatting for warning/error states, and wiring SET_CURRENCY
into the still-hardcoded GBP symbol) remains outstanding.

0.7.1 - review pass before phase 8. Seven fixes, no new features:

  - modReports.SheetNamed moved a sheet LATER with Move Before:, which is a
    no-op, so every build shipped Cost Calculations on tab 1 and Summary on
    tab 2 - the reverse of what the code asks for.
  - Val() on numeric cells replaced with modUtils.NumOf. Val round-trips
    through a locale-formatted string, so on any comma-decimal machine every
    paper cost, consumable rate and stock dimension read as zero. Design 8.2
    states this rule for dates; the cause is not specific to them.
  - modSnapshot.LogAudit and modExport.ExportSig no longer wrap a batch of
    independent operations in one On Error Resume Next (design 8.2).
  - modRegistry.WriteRegistry no longer raises on a workbook with no print
    rooms, which is the design 3.3 case it most needs to survive.
  - modSettings.Money made locale-safe and given CurrencySymbol.

A patch increment, not a phase: the phase digit still reads 7 because phase
8 has not been built. The revision digit is what "0.<phase>.<revision>"
exists for.
```
