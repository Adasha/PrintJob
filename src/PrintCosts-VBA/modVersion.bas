Attribute VB_Name = "modVersion"
Option Explicit

' Version identity, and the About popup on the Settings sheet.
'
' There are TWO version numbers and they mean different things:
'
'   APP_VERSION   this build of the workbook - code, layout, formulas.
'                 Changes every time something is rebuilt.
'
'   SCHEMA_VER    the shape of the stored data, in modUtils. Changes only
'                 when the job table's columns change, because that is what
'                 decides whether one workbook's records can be read by
'                 another (spec 11 collation, and the S_SchemaVer stamped on
'                 every job row).
'
' A build can go 0.4.0 -> 0.5.0 with the schema staying at 1.0, and that is
' the normal case. Conflating the two would force a schema bump on every
' cosmetic change and make the collation check meaningless.
'
' APP_VERSION is 0.<phase>.<revision>, tracking the build sequence in the
' design document: 0.7.x is "phases 1-7 built", and it reaches 1.0.0 when
' acceptance passes at the end of the sequence.

Public Const APP_NAME As String = "Print Cost Management"
' 0.9.16 - Location-sheet batch defaults, direct user report (2026-09-25 -
' reproduced on Mac, not Windows): picking an incompatible printer/paper
' default pair off the dropdown left the raw "(unavailable)"-suffixed value
' sitting in the cell, and Add Print Job copied it straight into the new
' row uncaught. modValidation.OnDefaultCellChanged already runs the AT-05
' compatibility check the instant a default cell changes, but that depends
' on Worksheet_Change firing - which Excel for Mac does not reliably do for
' an in-cell dropdown selection. modValidation.EnsureDefaultsClean is a new
' defensive re-check, called from modJobs.AddPrintJob right before it copies
' the defaults into the row, so a stale/incompatible default is caught and
' cleaned regardless of whether the change-time event fired.
'
' 0.9.15 - Job-table Data Validation corruption fix, direct user report
' (2026-09-25): every column from Unit through Chargeable Cost was showing a
' dropdown to pick a technician's name instead of its own content, and
' picking one overwrote a formula cell, breaking that row.
'
' Root cause, confirmed against the built .xlsm over COM: modInit.
' ReorderJobColumns (0.9.4) moves columns via repeated Range.Cut +
' Range.Insert Shift:=xlToRight inside the table. Excel extends a validated
' column's rule onto cells newly shifted in beside it on each Insert, and
' the ~30 inserts one full reorder performs compound that into a wide,
' wrong span - Technician's own list, meant for one column, ended up the
' Formula1 on eleven; the Date/Time rule similarly ended up on Student Name
' and Student No. The corruption happened once, when 0.9.4 first reordered
' an existing workbook's columns, and then sat there undisturbed - the
' reorder's own "already in the right position" check meant it never ran
' the Cut+Insert again to have a chance of un-corrupting it.
'
' modInit.EnsureJobColumnValidation (new) clears the whole table body's
' validation and restores exactly what PrintCosts.xlsx ships for Date/Time,
' Qty, Print Width mm, Disregard Paper, Disregard Consumable and Paid.
' Called from modLists.BindColumns - the existing "rebuild dependent
' dropdowns, self-heal, never trust what a previous run left behind" entry
' point that every setup run, Refresh Locations, Check workbook/sheet and
' the picker already go through - so the fix reaches an already-corrupted
' workbook on the next ordinary refresh, not only a full rebuild.
'
' 0.9.14 - Five direct user-feedback items, layout and naming only, no schema
' change:
'
'   - "Quantity" renamed to "Qty" everywhere it is a table/column header (the
'     job table on every location sheet, the Reports results table and its
'     Sort-by list, Summary's breakdown table, the export column order/
'     header row) - narrower headers were costing width the already-tight
'     job-row columns needed elsewhere. modInit.EnsureQtyColumnName renames
'     the live ListColumn (checked-first, same shape as EnsurePaidColumn)
'     rather than hand-editing PrintCosts.xlsx; Excel updates every
'     structured reference (Area m2/Paper Cost's own formulas, _Data's
'     consolidated header copy) as part of the rename itself. Every VBA call
'     site that looked the column up by its old name - modReports.C/SumBy,
'     modExport.ExportColumns/Agg, modImport.WriteImportedRow - now asks for
'     "Qty". Not a schema bump (§3.5: a rename doesn't change what a column
'     MEANS, only its label) - a file exported before the rename degrades
'     the same graceful way a pre-Paid export already does (missing key
'     reads back blank, not an error). The Reports page's OWN "Quantity"
'     filter-box label (F8) is deliberately left alone - a criteria label,
'     not a table header, with no width pressure on it.
'   - "Standard size" (tblPapers) renamed to "Std. size" the same way -
'     modCatalog.EnsureStdSizeColumnName - but with no call sites to fix,
'     since nothing in this codebase looked that one up by name.
'   - Location-sheet job-row columns narrowed to fit more of the table in
'     one screen: Unit and Qty capped at 50px, Area m2 at 60px, Paid at
'     40px; Job ID and the Reports page's own Date/Time column fixed at
'     exactly 100px. modUtils.ColWidthForPx converts a pixel target into
'     Range.ColumnWidth's own "characters of the Normal-style font" unit
'     using Calibri 11's well-known metric (7px/char + 5px padding, the same
'     constant openpyxl/xlsxwriter use) - there is no Range.WidthInPixels to
'     set directly. Always rounds down, so a cap can never render one pixel
'     over.
'   - Summary's own totals-breakdown header (row 9) now shares the Reports
'     page's warm totals tint (RGB(244, 232, 222), added in 0.9.13 for
'     Reports' own "Matching" row) instead of the blue it used to share with
'     Reports' results-table header - this table is itself a totals
'     breakdown (by location/printer/paper stock), the same category as
'     "Matching", not a per-job record list.
'   - A blank row now sits both above and below the location-sheet batch-
'     defaults row (modInit.EnsureJobTableGap, the same "insert a row above
'     a table, checked first" idiom EnsureTableGap already used for the
'     catalogue tables) - it used to run straight into the main toolbar
'     below with no gap at all, and read as crowded against the config
'     block above despite row 8 already being blank there. Shifts the
'     defaults from row 9 to row 10 and the toolbar from row 10 to row 12;
'     every hardcoded row reference in modInit (EnsureJobDefaults,
'     DrawLocationButtons, the reduced-view toggle/Clear defaults buttons)
'     moved with it. Named ranges (LOC_DefTech/LOC_DefPrinter/LOC_DefPaper)
'     needed no separate migration - Excel's own row-insert already carries
'     a name's RefersTo along with whatever physically moved.
'
' 0.9.13 - Two small Reports-page fixes, direct user feedback:
'
'   - Matching's totals header (row 12) now gets its own subtle tint,
'     RGB(244, 232, 222) - the results header's own RGB(222, 232, 244) with
'     red and blue swapped, so the two bands read as related but distinct
'     rather than the totals row looking unstyled next to it.
'   - "Show student/department name/no" (the toggle at what was E9/F9) had
'     no visible label of its own reason to be there - it is a display
'     option, not a filter criterion, and shared row 9 with the unrelated
'     name/number mismatch warning purely because that row happened to be
'     free. Moved to a third group on the sort-controls row (I10/J10),
'     after Sort by/Sort direction, with its own label restored. Every
'     formula and comment that referenced $F$9 now reads $J$10 instead;
'     nothing outside modReports.bas touched that cell.
'
' 0.9.12 - Settings-page layout: the About section (modVersion.WriteAbout,
' removed) is gone from the bottom of the Settings sheet, per direct user
' feedback that it duplicated the About popup for no benefit while eating
' vertical space. The eight action buttons that used to run down column T -
' out of the way, but also out of sight on a normal-width window - now sit in
' the freed space instead, directly below whichever table on the sheet runs
' deepest (modInit.TablesBottom, the same "bottom" WriteAbout used to compute
' before clearing its own text). Two logical groups of four, one row each:
' Refresh Locations/Check workbook/Re-stamp prices/About, then Export All
' Locations/Import/Backup workbook/Restore workbook. modInit.InitialiseWorkbook
' also clears the old About block's text area unconditionally on every setup
' run, not just a from-scratch rebuild, so a workbook built before this change
' loses the stale text the next time setup runs on it.
'
' ShowAbout (the popup itself, unchanged in structure) now carries the brief
' description that used to live only in the About block's prose, so the
' popup alone covers project name, version, author and what the workbook
' does - nothing is lost by removing the on-sheet block.
'
' 0.9.11 - Printer/Paper Stock dropdowns no longer hide incompatible options,
' per direct user feedback that the bidirectional narrowing added in 0.9.0
' was "too restrictive." Every active/permitted item is now always listed;
' whichever ones the other field's current value rules out are suffixed
' " (unavailable)" (modLists.UNAVAILABLE_SUFFIX) instead of being removed.
' Picking one is allowed - modValidation.Clean() strips the suffix and hands
' the clean name to the EXISTING AT-05 "incompatible combination" handling
' (OnPrinterChanged/OnStockChanged/OnDefaultCellChanged), which already
' clears the other field and explains why, unchanged from 0.9.0. Auto-fill
' when exactly one compatible option remains is UNCHANGED - it now runs off
' a narrowed list computed only for that decision, never shown to the
' dropdown itself, so 0.9.0's tested auto-fill behaviour (test-dropdowns.ps1)
' still holds.
'
' Known compromise, not a full fix - flagged for revisiting rather than
' presented as settled: Excel's native in-cell dropdown (Data Validation)
' cannot style individual list entries - no italics, colour or shading on
' one item within the list, which is what was actually asked for. A text
' suffix is what is achievable inside a real Excel dropdown, and is judged
' good enough for now - the workbook stays with native Excel styling
' throughout rather than building a bespoke picker UI just for this. See
' docs/ARCHITECTURE.md §7.2 / §16.3.
'
' 0.9.10 - Department/"charge to" terminology (snag 4b, D19): free text,
' no new catalogue table, per the user's own decision on review. The
' Student Name/No columns already accept anything typed into them (D1 -
' no authoritative student list, cross-checked only for consistency) - a
' department name behaves identically with no code change needed. Only the
' user-facing WORDING that names the field got a light "Student/
' Department" pass: the Reports filter labels and autocomplete titles, the
' "Show student/department name/no" toggle, the Remove-row confirmation
' ("Student/Department: "), the name/number consistency warning, and the
' About blurb. The underlying column headers stay "Student Name"/"Student
' No" deliberately - renaming those would break every header-name lookup
' that already reads them (§5, §8.3, §10.4 of the architecture doc) for no
' real capability gained.
'
' 0.9.9 - Full workbook backup/restore (snag 4a). New modBackup.bas: Backup
' workbook writes one CSV per catalogue table (Technicians, Printers,
' Papers, the four Settings-page lookup tables, and Settings itself)
' alongside the existing per-location job exports, all sharing one
' timestamp; Restore workbook reads any one file from a backup back in,
' finds every sibling sharing the same timestamp, and applies the whole
' set - overwrite-by-stable-key for catalogue rows (mirroring Import's own
' Job ID rule), and modImport.ApplyImportConfirmed unchanged for job
' records. Two new Settings buttons: "Backup workbook..." / "Restore
' workbook...".
'
' One genuine, previously-latent bug found and fixed at the root: VBA's And
' does not short-circuit, so `Count = 1 And IsBlankRow(lo, 1)` - a pattern
' already used in four other modules - still called IsBlankRow (indexing
' ListRows(1)) even when Count had already evaluated to 0, raising
' "Subscript out of range". Dormant until now because nothing in ordinary
' use ever left a table at genuinely zero rows; restoring into a table
' emptied by hand is the first path to actually hit it. modUtils.IsBlankRow
' now returns False for an out-of-range row instead of indexing blindly,
' fixing every call site at once.
'
' 0.9.8 - Summary sheet fixes (snags 3a, 3b). ToggleConfigSheets (the
' Hide/Show settings sheets button) now captures the Summary worksheet
' explicitly and re-activates it unconditionally after hiding/showing the
' four configuration sheets, rather than relying on Summary having stayed
' active throughout - hiding whichever sheet happens to be active is what
' forces Excel to jump to the next one in tab order, so this is correct
' regardless of how ToggleConfigSheets was actually invoked. The "Go to
' Settings" button is removed outright rather than made target-aware: with
' four configuration sheets and no way to tell which one a task needs, it
' could only ever jump to one of them, and Hide/Show settings sheets plus
' Excel's own tabs already reach all four once visible - the snag list's
' own "easiest fix."
'
' 0.9.7 - Export report's header block and single-value promotion (snags
' 2b, 2c). modExport.ExportReportSnapshot's .xlsx archive now opens with a
' metadata block - report title, schema version, site ID/name, the date
' range covered (the From/To filter boxes when set, else the actual spread
' of the exported rows' own Date/Time column), a generated-at timestamp and
' the row count - then a blank row, then the results table. Live Reports
' sheet untouched throughout, per the user's own review answer: this is
' export-time-only.
'
' Single-value promotion: any of Student name, Student no, Location,
' Printer, Paper stock, Technician that holds the SAME non-blank value on
' every exported row is lifted into the header block as a "Header: Value"
' line and dropped from the table, so filtering a report down to one
' printer (say) states that once at the top instead of repeating it down an
' entire column. A candidate column with no non-blank values at all (e.g.
' Student name/no when the 2a toggle is off) is left alone untouched -
' there is nothing true to promote.
'
' One real bug found building this: LastVisibleColumn (renamed
' LastHeaderColumn) filtered its column bound by .Hidden, which was
' harmless when only the trailing Job ID column was ever hidden - but 2d's
' minimum-columns view (0.9.6) now hides Printer and Technician among
' others, so the OLD function silently truncated every export at the last
' VISIBLE column, dropping Technician and Notes from the file entirely with
' no error. The promotion logic needs those hidden columns' values to check
' for uniformity, so the bound is now by header name (stop before the
' trailing "Job ID") rather than by Hidden state.
'
' 0.9.6 - Reports student name/no toggle and minimum-columns view (snags
' 2a, 2d): a Yes/No toggle at F9 (default No) blanks Student Name/No in the
' results table and any export, satisfying data protection; Student Name/
' No and Paid added to the results table; only seven columns (Date/Time,
' Location, Student Name, Student No, Paper stock, Chargeable, Paid) stay
' visible by default, the rest hidden via modReports.ApplyReportsMinimum
' Columns (not user-facing yet - documented future enhancement).
'
' Four real bugs found and fixed while testing edge cases (blank-everything,
' nothing-paid-yet), not just the happy path:
'   - The relaid-out "Matching" totals row (MatchTotal, label above value
'     rather than to its left) initially still shared columns with newly-
'     hidden results columns, so several totals displayed blank - same
'     class of bug as the location-sheet header-block collision, now fixed
'     by only ever using columns the minimum-columns view never hides.
'   - SUM(FILTER(...))-SUM(FILTER(...)) for the Unpaid total returned 0
'     whenever nothing was marked Paid yet - FILTER with zero matches
'     raises #CALC!, which poisons the whole subtraction and gets masked
'     by the outer IFERROR into a false 0. Fixed with FILTER's own
'     [if_empty] third argument.
'   - IF($F$9="Yes", <column array>, "") broke every results row past the
'     first once the toggle was off: the condition is a scalar (one toggle
'     cell), so the IF collapses to the bare scalar "" rather than a
'     column of blanks - HSTACK does not broadcast a scalar against a
'     column, the exact trap _Data's own consolidated-range formula (§7.1)
'     already had to work around, met again in a new place. Fixed by
'     nesting the toggle inside an outer IF whose own condition is already
'     a genuine per-row array.
'   - Paid displayed a literal "0" for any row that predates the column
'     (blank Paid counts as unpaid, design 5) - INDEX on a truly blank
'     cell returns the number 0, and &"" does not fix an already-numeric
'     0 (0&"" is still the text "0"). Fixed by testing for that specific
'     value, safe here since Paid never legitimately holds a real 0.
'
' Also found (and fixed) a TEST-SCRIPT bug that looked exactly like
' intermittent COM automation flakiness: a PowerShell [char] handed to
' Excel's COM Columns(...) indexer is read by its ORDINAL VALUE, not as a
' one-letter column reference - .ToString() first fixes it. Cost real
' debugging time because the symptom (a genuinely hidden column reading
' back as visible) is indistinguishable from a real settling/timing issue
' until compared side by side against the literal column letter.
'
' 0.9.5 - spotted while visually checking 0.9.4's fix: two extra column
' groups next to the cost group, one of them hiding a genuine bug.
' GroupJobColumns used to ALSO group S_PrinterID->S_SchemaVer, on top of
' the grouping PrintCosts.xlsx already ships for that exact span - nesting
' a second outline level over the first rather than reusing it. Visibly:
' two unwanted extra groups cluttering the outline pane next to the one
' group that matters day to day. Substantively: Notes ended up hidden - a
' genuine input column, never meant to be, silently swept into the
' snapshot block's hidden state as a side effect of EnsurePaidColumn (0.9.2)
' inserting Paid directly next to it. modInit.FlattenOutline (new) resets
' Notes->S_SchemaVer to outline level 1 outright rather than re-grouping
' it, and explicitly un-hides Notes; H_Issues/the snapshot columns keep
' their own .Hidden state exactly as before.
'
' 0.9.4 - two bugs found testing 0.9.3's reduced-clutter view, both fixed
' the same day:
'
'   - Hiding a column hides the WHOLE column, every row, not just the
'     table's - Status and Job ID used to be the table's first two columns
'     (sheet A/B), the SAME columns the location header block above the
'     table (room name, department, code, defaults) occupies in rows 1-9.
'     Turning reduced view on blanked the header along with the table
'     cells. modInit.ReorderJobColumns moves Status and Job ID to just
'     after Paid, ahead of Notes/H_Issues/the snapshot block, so columns
'     A/B are always Date/Time/Student Name - never anything reduced view
'     can hide. Column order isn't load-bearing anywhere (design 5.1), with
'     one matching fix: modRegistry's consolidated-range span was bounded
'     by column NAME from "Job ID" to "Notes" - with Job ID no longer the
'     leftmost column, that bound moved to "Date/Time" so the span still
'     covers everything (Status/Job ID included, now inside the span
'     rather than starting it).
'   - The toggle button itself was anchored to column 13, which happened to
'     be one of its own hidden columns (Disregard Consumable) - turning
'     reduced view on could take out the one button that turns it back
'     off. Moved to column 7, deliberately independent of the reorder
'     above.
'
' Neither is a general fix for the underlying constraint (Excel's
' column-hide is whole-column-only, and the hidden-column list is
' user-editable, so a future edit could collide with something else on the
' sheet) - revisit properly later. Also flagged: the toggle's only feedback
' is its own button caption; a clearer visual state indicator is worth
' adding, deferred alongside the rest of the visual-polish work.
'
' 0.9.3 - reduced-clutter view toggle (snag 1e): a workbook-wide (not
' per-sheet) toggle hiding Status, Job ID, Printer, Area m2, Disregard Paper
' and Disregard Consumable on every location sheet at once, driven by two
' self-provisioned settings (SET_LOC_REDUCED_VIEW, SET_LOC_REDUCED_COLUMNS)
' rather than a hard-coded list, so the shortlist can change without a
' rebuild. modInit.ApplyColumnVisibility is the general "hide/show exactly
' these headers, touch nothing else" primitive underneath - deliberately not
' a blanket show-everything-then-hide-the-list reset, since H_Issues is
' hidden permanently and must never be touched by this. Same primitive gets
' reused by the Reports page's own fixed minimum-columns view once that
' lands (snag 2d).
'
' 0.9.2 - Paid job-row column and the column-grouping fix (snags 1c, 1d):
'
'   - A new Paid (Yes/No) column, right after Chargeable Cost. The first
'     genuine job-row column change since inception, so SCHEMA_VER (modUtils)
'     moves 1.0 -> 1.1 - which surfaced that SET_SCHEMA had shipped as a
'     static value in PrintCosts.xlsx with no code keeping it in sync with
'     the constant, the same drift that caught APP_VERSION out below. Fixed:
'     modVersion.EnsureSchemaSetting now stamps SET_SCHEMA from SCHEMA_VER on
'     every setup run, same as EnsureVersionSettings already does for
'     APP_VER. Existing rows are left blank rather than force-set to "No" -
'     blank already reads as unpaid everywhere (Summary/Reports totals,
'     export, import).
'   - modInit.GroupJobColumns's cost-column group ran one column too far
'     (Paper Cost through Chargeable Cost itself), so collapsing it hid the
'     one figure a collapsed view most needs to keep showing. Now runs
'     Paper Cost through Disregarded only; Chargeable Cost and the new Paid
'     column stay outside it, always visible.
'   - Summary and Reports both split their Chargeable total into Paid/
'     Unpaid, Unpaid computed as "filtered Chargeable minus Paid" rather
'     than a second criteria expression, so the two figures can never drift
'     apart from the total they came from.
'
' 0.9.1 - batch default Technician/Printer/Paper selectors (snag 1b): three
' cells above the toolbar (row 9) pre-fill Technician/Printer/Paper Stock on
' every subsequently added job until Clear defaults empties them, using the
' same copy-not-reference principle as the existing disregard-cost defaults
' (AT-07/AT-08) and the same bidirectional filtering/autofill as 0.9.0's
' table-cell dropdowns. modInit.EnsureJobDefaults self-provisions the three
' named cells (LOC_DefTech/LOC_DefPrinter/LOC_DefPaper), same pattern as
' LOC_Export; modLists.BindDefaultCells and modValidation.OnDefaultCellChanged
' reuse 0.9.0's binding/compatibility-check machinery rather than duplicating
' it.
'
' 0.9.0 - phase 9 begins: the 2026-09-22 post-phase-8 snag list. Real new
' functionality, not just fixes, so the phase digit moves again:
'
'   - Printer and Paper Stock on a location sheet now filter each other in
'     BOTH directions, auto-filling either one when only a single compatible
'     option remains. Previously Paper Stock was locked until a printer was
'     chosen; both cells now start unlocked and fully populated.
'     modCatalog.PrintersForStock is new (the reverse of the existing
'     StocksFor); modLists.BindPrinterRange/BindStockRange/AutoFillIfSingle
'     replace the old one-directional BindStockRange; modValidation's
'     OnPrinterChanged/OnStockChanged are now symmetric, each rebinding both
'     cells' lists rather than assuming only one drives the other. See
'     docs/ARCHITECTURE.md §7.2 for the design and a note on the one real
'     bug this surfaced (Range.Value on a multi-cell Union is an array, not
'     a scalar - AutoFillIfSingle has to walk target.Cells, not read
'     target.Value directly).
'
' 0.8.1 - phase 8's original scope, finally built (see the 0.8.0 note below,
' which explicitly deferred this):
'
'   - modSettings.CurrencySymbol (already reading SET_CURRENCY, previously
'     unreferenced) is now read everywhere a NumberFormat or Format$ used to
'     hardcode ChrW(163) - modReports (Summary and Reports), modExport's
'     export-report snapshot, and modJobs.RemoveRow's delete-confirmation
'     dialog. The NumberFormat sites go through the new
'     modSettings.CurrencyFormatCode rather than CurrencySymbol() directly -
'     verified directly against this workbook over COM: setting a Range's
'     NumberFormat from VBA to a BARE or quoted currency symbol (e.g.
'     "$#,##0.00") is silently canonicalised by Excel back to the OS's own
'     regional currency symbol, no matter what SET_CURRENCY holds - invisible
'     on a GBP-locale machine, where it happened to always match, but it
'     would have silently defeated the setting everywhere else. Excel's
'     "[$symbol]" bracket syntax round-trips exactly as given and was
'     confirmed to render identically for the unaffected GBP case, so
'     CurrencyFormatCode is a strict fix, not a behaviour change. Format$
'     (RemoveRow's dialog) does not have this problem - confirmed separately
'     - and keeps using CurrencySymbol() plain. Changing SET_CURRENCY still
'     only repaints Summary/Reports on the next InitialiseWorkbook run, same
'     as every other piece of their formatting.
'   - Conditional formatting exists in the workbook for the first time.
'     modInit.ApplyStatusFormat gives every location sheet's Status column an
'     amber fill whenever a row reads anything other than "OK" - the
'     "warning state" §5/§11 already called out as "conditionally
'     formatted". modReports.FormatSummaryErrors gives Summary's Type/Family
'     columns red bold text when a job references a paper stock that no
'     longer exists in tblPapers ("(not in Papers)") - the "error state",
'     matched to the one place the workbook's own formulas already flag a
'     genuine data-integrity break rather than a routine per-row warning.
'   - modReports.DrawLegend adds the Summary-sheet legend §11 planned and
'     never built: the four everyday cell-role colours (input, calculated,
'     configuration, read-only reference), at O9 down, in the colours
'     already used for those roles elsewhere on Summary/Reports.
'
' 0.8.0 - NextId high-water mark, Export/Import, Reports rework, UI snag
' list. Real new functionality, not just fixes, so the phase digit moves:
'
'   - modRegistry.NextJobId now allocates from a persisted high-water mark
'     (a new "Job ID HWM" registry column) instead of a scan of rows on the
'     sheet, so deleting the highest-numbered job can no longer reissue its
'     ID. Fixed first because Export/Import both depend on the "globally
'     unique, never re-issued" invariant actually holding.
'   - modExport gained Export All Locations (every room in one pass) and an
'     Export report snapshot - a static-value .xlsx of the Reports sheet's
'     current filtered/sorted view, with its own signature stamped on
'     Reports for the delete command below to check against. Exports now
'     resolve to a selectable folder (SET_EXPORT_FOLDER, falling back to the
'     workbook's own location if unset or not present on this machine), and
'     the OneDrive path fix now also searches
'     ~/Library/CloudStorage/OneDrive* on Mac, where the Windows-only
'     OneDrive environment variables never existed in the first place.
'   - modImport is new: restores or merges an exported file into a
'     location's job table. Rows match by Job ID - append if new, overwrite
'     if not - and only input/snapshot columns are written, so an imported
'     row costs correctly from its own frozen rates regardless of the
'     target workbook's current catalogue.
'   - Cost Calculations renamed to Reports: Technician/Printer/Paper
'     Stock/Quantity filters, sort-by-any-column, a hidden Job ID
'     correlation column, and a bulk "Delete visible records" command
'     confined to whatever the active filters show. Summary regrouped to a
'     Location x Printer x Paper stock key.
'   - The 2026-09-21 snag list: no default sheet-protection password,
'     unlocked configuration table inputs, freeze panes moved off the
'     config sheets and onto Reports, tab order fixed, a hide/show toggle
'     for the config sheets, column and row groups, Settings notes
'     wrapped, and catalogue row add/remove buttons.
'
' Not part of this bump: phase 8's original scope (the Summary legend,
' conditional formatting for warning/error states, and wiring SET_CURRENCY
' into the still-hardcoded GBP symbol) remains outstanding.
'
' 0.7.1 - review pass before phase 8. Seven fixes, no new features:
'
'   - modReports.SheetNamed moved a sheet LATER with Move Before:, which is a
'     no-op, so every build shipped Cost Calculations on tab 1 and Summary on
'     tab 2 - the reverse of what the code asks for.
'   - Val() on numeric cells replaced with modUtils.NumOf. Val round-trips
'     through a locale-formatted string, so on any comma-decimal machine every
'     paper cost, consumable rate and stock dimension read as zero. Design 8.2
'     states this rule for dates; the cause is not specific to them.
'   - modSnapshot.LogAudit and modExport.ExportSig no longer wrap a batch of
'     independent operations in one On Error Resume Next (design 8.2).
'   - modRegistry.WriteRegistry no longer raises on a workbook with no print
'     rooms, which is the design 3.3 case it most needs to survive.
'   - modSettings.Money made locale-safe and given CurrencySymbol.
'
' A patch increment, not a phase: the phase digit still reads 7 because phase
' 8 has not been built. The revision digit is what "0.<phase>.<revision>"
' exists for.
Public Const APP_VERSION As String = "0.9.16"
Public Const APP_AUTHOR As String = "Adam Shailer"

Public Function VersionString() As String
    VersionString = APP_NAME & " v" & APP_VERSION & " (data schema " & SCHEMA_VER & ")"
End Function

' --------------------------------------------------------------- stamping ---
' Called by build.ps1 AFTER InitialiseWorkbook, not by it.
'
' "Built" has to mean when this file was produced, and InitialiseWorkbook is
' re-run by users every time they add a print room. Stamping from there would
' quietly reset the build date to whenever somebody last duplicated a sheet.
Public Sub StampBuild()
    On Error GoTo Fail
    AppOff
    EnsureVersionSettings
    SetSetting "BUILT", Now
    SetSetting "BUILT_BY", CurrentUser
    StampProperties
    AppOn
    Say "Stamped as " & VersionString & ".", "Built " & Format$(Now, "dd/mm/yyyy hh:mm") & " by " & CurrentUser & "."
    Exit Sub
Fail:
    AppReset
    ReportError "StampBuild"
End Sub

' Adds the version rows to tblSettings if they are not there, and gives each a
' workbook-scoped SET_ name so modSettings can reach it like any other setting.
'
' APP_VER is written from the constant every time, because the code in the
' workbook IS its version - importing a newer module set makes it a newer
' workbook, and the cell should say so without waiting for a rebuild. BUILT
' and BUILT_BY are left to StampBuild.
Public Sub EnsureVersionSettings()
    EnsureSetting "APP_VER", "Workbook version", "Read-only. Matches the code in this workbook."
    EnsureSetting "BUILT", "Built", "Read-only. When this file was produced."
    EnsureSetting "BUILT_BY", "Built by", "Read-only."
    SetSetting "APP_VER", APP_VERSION
End Sub

' Stamps SET_SCHEMA from modUtils.SCHEMA_VER on every setup run, the same
' reasoning as APP_VER above: the code in the workbook IS its schema, so a
' value that only ever shipped statically in PrintCosts.xlsx has no way to
' stay in sync with it otherwise - exactly the drift that caught out
' APP_VERSION for thirteen commits before this file's own §16.1. SCHEMA had
' never needed this before 2026-09-22 (design 3.3): no job-row column had
' ever changed since inception, so the shipped static value was never wrong.
' The Paid column (snag list item 1c) is the first real bump, and is what
' surfaced the gap - fixed here rather than hand-editing the .xlsx, for the
' same reproducibility reason every other self-provisioned setting exists.
Public Sub EnsureSchemaSetting()
    EnsureSetting "SCHEMA", "Data schema version", "Read-only. The shape of the stored job-row columns; changes only when one is added, removed or renamed."
    SetSetting "SCHEMA", SCHEMA_VER
End Sub

' Public: modExport.EnsureExportSettings reuses this for the same reason
' EnsureVersionSettings' own rows go through it - a setting is not real to
' modSettings.SettingText until it has a SET_<KEY> name, and this is the one
' place that adds a row to tblSettings and names it in the same step.
Public Function EnsureSetting(ByVal Key As String, ByVal Label As String, ByVal Notes As String) As Range
    Dim lo As ListObject, i As Long, r As ListRow, c As Range

    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Key").Value)), Key, vbTextCompare) = 0 Then
            Set c = CellIn(lo, i, "Value")
            Exit For
        End If
    Next i

    If c Is Nothing Then
        UnlockSheet lo.Parent
        Set r = lo.ListRows.Add
        r.Range.Cells(1, 1).Value = Key
        r.Range.Cells(1, 2).Value = Label
        r.Range.Cells(1, 4).Value = Notes
        Set c = r.Range.Cells(1, 3)
        RelockSheet lo.Parent
    End If

    ' modSettings addresses every setting as SET_<KEY>, so the name is what
    ' makes a new row a real setting rather than just a row of text.
    On Error Resume Next
    ThisWorkbook.Names("SET_" & Key).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="SET_" & Key, _
        RefersTo:="='" & c.Parent.Name & "'!" & c.Address(True, True, xlA1)

    Set EnsureSetting = c
End Function

' --------------------------------------------------- file-level identity ---
' So the version is visible without opening the file: Explorer's details pane,
' Finder's Get Info, and File > Info in Excel all read these.
' Each property is set individually and failures are reported rather than
' swallowed. An earlier version wrapped the lot in On Error Resume Next and
' silently wrote nothing at all - the properties came back empty and the only
' way to find out was to go and look.
Private Sub StampProperties()
    Dim bad As String
    bad = SetProp("Title", APP_NAME & " v" & APP_VERSION)
    bad = bad & SetProp("Subject", "Print output and cost logging, data schema " & SCHEMA_VER)
    bad = bad & SetProp("Author", APP_AUTHOR)
    bad = bad & SetProp("Comments", VersionString & ", built " & Format$(Now, "dd/mm/yyyy hh:mm"))
    If Len(bad) > 0 Then
        Say "Some file properties could not be set.", bad, _
            "The version is still on the Settings sheet; only the Windows file details are affected."
    End If
End Sub

Private Function SetProp(ByVal Nm As String, ByVal Value As String) As String
    On Error GoTo Fail
    ThisWorkbook.BuiltinDocumentProperties(Nm).Value = Value
    Exit Function
Fail:
    SetProp = "- " & Nm & ": " & Err.Number & " " & Err.Description & vbCrLf
End Function

' ------------------------------------------------------------------ about ---
' Bound to the About button on the Settings sheet. The only place any of this
' is shown now that the on-sheet About block (WriteAbout) is gone, so it
' carries the brief description that used to live only in that block's prose
' as well as the identity/version details it always showed.
Public Sub ShowAbout()
    Dim built As String, who As String, ver As String

    ver = SettingText("APP_VER", "(not stamped)")
    built = SettingText("BUILT")
    who = SettingText("BUILT_BY")

    Say APP_NAME & " v" & ver, _
        "Logs print output and cost per student or department across the university's print bureaux. " & _
        "Each record freezes the prices it was costed at, so changing a paper or consumable price " & _
        "never alters what has already been charged." & vbCrLf & vbCrLf & _
        "Data schema " & SCHEMA_VER & IIf(Len(built) > 0, vbCrLf & "Built " & built, "") & _
        IIf(Len(who) > 0, " by " & who, ""), _
        "Specified by " & APP_AUTHOR & ". Built by Claude (Anthropic) working with " & APP_AUTHOR & ", September 2026.", _
        "About"
End Sub
