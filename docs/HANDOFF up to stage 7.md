# Handoff — University Printing Cost Management Workbook

**Source:** `docs/stage1-design-architecture.md` v1.8 (20 Sep 2026). The PDF of the same document has been removed as redundant; its text layer was never extractable anyway (Type3 subsetted fonts, no ToUnicode map), so anything reading it had to render it as images.
**Workbook version:** 0.7.1, data schema 1.0
**State:** Phases 1–7 built and verified, plus a 0.7.1 review pass. **Phase 8 (visual polish, legend, user notes) is next. Phase 9 (acceptance on Windows and Mac, test report) follows.**

Unlike the previous revision of this handoff, this one is written with the repo in hand. The claims below were cross-checked against the built `.xlsm` by reading the OOXML package directly, not inferred from the design document — see §10 for exactly what was confirmed and what was not.

---

## 1. What this is

An Excel `.xlsm` (Microsoft 365 only, Windows and Mac) for logging print jobs per print room and costing them. One location sheet per room, plus Summary and Cost Calculations report sheets. VBA handles entry, validation, snapshotting, registry and export. Reports are live dynamic-array formulas, not VBA output.

## 2. Immediate next steps

**Phase 8 has a pinned scope.** §10 of the design doc lists the visual model; most of it is already built. What is genuinely missing is three things and no more:

1. **The Summary legend** explaining the four everyday cell states.
2. **Conditional formatting for the warning (amber fill) and error (red text) states.** There is currently **no conditional formatting anywhere in the workbook** — confirmed by scanning every sheet's XML. §4 of the design doc also calls for the Status column to be conditionally formatted. This is the bulk of the phase.
3. **Currency from the global setting.** §10 says "GBP formatting driven by the global setting"; in fact `£` is hardcoded as `ChrW(163)` in twelve places across `modReports` and `modJobs`. `modSettings.CurrencySymbol()` exists as of 0.7.1, reads `SET_CURRENCY`, and is wired to nothing. `SET_FOOTER` and `SET_FY_START` are likewise defined names with no consumer — decide whether to use them or drop them from §2.3.

Already done, so do not redo: fills, borders, italics, the collapsed snapshot group, freeze panes (`xSplit=2 ySplit=12`) and filter buttons on all twelve tables.

**Carried into phase 8:** `modExport.ExportLocation` never sets `Application.DisplayAlerts` around its `SaveAs`, and every test script sets it `$false` on the Application before driving the workbook — so Excel's CSV prompt cannot appear under test and will appear for a user. Untidy rather than dangerous since 0.7.1, because the file-exists check now turns a cancelled prompt into an honest failure instead of a false success.

**Phase 9:** acceptance testing. AT-01, 02, 04–07 need testing as a person typing, not just by script. AT-15 needs a Mac (picker render check). Also add the end-to-end export case: export, delete the sheet, confirm records are recoverable from the CSV. Two further manual checks now earn their place — the OneDrive-URL export path (§5), which no script exercises any more, and the SETUP.md manual import route on a Mac, which was broken until 0.7.1.

Bump the workbook version to `0.8.0` when phase 8 is built.

Verification status: AT-10, 11, 12, 13, 14, 16 verified by script. AT-03, 08, 09 hold by construction. AT-01, 02, 04–07, 15 wait for phase 9.

## 3. Key decisions (do not re-litigate without reason)

| # | Decision |
|---|---|
| D1 / I5 | No Students register. Name/number consistency is checked against all existing rows and **warns, never blocks** |
| D2 | **Snapshot** resolved rates onto each job row at creation. Visible costs are live formulas over the snapshot values |
| D3 | Global settings: currency and rounding, org/report header, site identity. No VAT, no overhead |
| D4 | Per-location toolbar of Form Control buttons, plus double-click to stamp date/time |
| D5 | Location sheets found by a **marker cell** (`AZ1` = `PRINTLOC/v1`). *Refresh Locations* rescans and rebuilds a hidden registry |
| D6 | One `Cost` column on Papers, with a calculated `Cost unit` label (£/sheet or £/metre) |
| D7 | Print width in mm |
| D8 | Gross, Disregarded and Chargeable are separate columns. Consumption is unaffected by disregard flags |
| D9 | Standard A-sizes picker, or custom width × height in mm |
| D10 | Multi-select via a **worksheet** picker (`_Picker`), writing a delimited string to one cell |
| D12 | **Microsoft 365 only**. Dynamic arrays may be relied on |
| D13 | Reports are formulas over a consolidated spill range on `_Data` |
| D15 | `.xlsm` built by PowerShell driving Excel over COM. Mac uses the manual import route |
| D16 | **Two independent version numbers**: workbook version (build) and data schema version (record shape) |
| D17 | **Multi-workbook aggregation is out of scope for v1** (planned feature). O7 deferred on GDPR grounds |
| D18 | Deleted location sheets: **export-based mitigation** (per-location CSV, visible status), not a shadow copy. Excel has no `BeforeDeleteSheet` event, so this is "a sign on the cliff, not a fence" |

**Interpretations confirmed:**
- I1: sheet-stock consumable area = sheet area × number of sheets.
- I2: paper families carry an explicit `Measurement basis` (`Sheet`/`Roll`). Behaviour keys off that, never the family name.
- I3: sheet quantities are whole numbers; roll quantities are decimal metres.
- I4: rounding is per cost component at row level.

## 4. Architecture in brief

**Data model.** Config tables: `tblSettings`, `tblPaperTypes`, `tblStandardSizes`, `tblPaperFamilies`, `tblConsumables`, `tblTechnicians`, `tblPrinters`, `tblPapers`. Each setting has a workbook-scoped name `SET_<KEY>`; `modSettings` resolves settings by that name.

**IDs.**
- Config IDs are `TEC-001`, `PRN-001`, `STK-001`. They are never re-issued and come from the highest ever issued.
- Job IDs are `<SITE>-<LOC>-00001`.
- Location codes are **never reassigned**, because they are baked into job IDs.
- Config rows are deactivated, not deleted.

**Hidden sheets** (`xlSheetVeryHidden`):
- `_Registry`: `tblLocations`, including `Last export` and `Export sig`, keyed by sheet name deliberately so a duplicated sheet reads as unexported.
- `_Work`: validation staging only.
- `_Data`: the consolidated range.
- `_Audit`, `_Export` (reserved), `_Picker`.

**Sheet order.** Summary is tab 1, Cost Calculations tab 2. Until 0.7.1 every build shipped these reversed — see §5.

**Location sheet layout.**
- Rows 1–8 hold the config block with sheet-scoped names (`LOC_Name`, `LOC_Dept`, `LOC_Code`, `LOC_DefDisPaper`, `LOC_DefDisCons`, `LOC_Printers`, `LOC_Status`, `LOC_Export`).
- Row 10 is the toolbar, row 12 the header, row 13+ the table `tblJobs_<CODE>`.
- Columns 1–21 are the visible table A–U (through `H_Issues`); 22–33 are the snapshot block V–AG. `H_Issues` and the snapshot block are hidden and share one collapsed outline group.
- Column order is **not load-bearing**: everything resolves by header name, so reordering is free and does not bump the schema version.

**Job-row formulas** read only snapshot columns and the row's own inputs. No job formula touches `tblPapers` or `tblPrinters`. That is what satisfies historical integrity (AT-09). Print width appears in the Area formula and never in Paper Cost (AT-02/03). The Area formula ships as a branching `IF` rather than the design doc's `LET`, but expands to exactly the documented result on both branches.

**Refresh Locations sequence:** park all job tables on temporary names → assign codes where blank or duplicate (original keeps its code because it is reached first) → rename tables to `tblJobs_<code>` → re-point button `OnAction` → rebuild dependent dropdowns and protection → rewrite the consolidated formula on `_Data!A10` → recompute export status → stamp `SET_LASTREF`.

**Reports.**
- Summary is one row per Location × Paper stock, with totals **above** the spill.
- Cost Calculations is a live `FILTER` over four criteria cells, with breakdowns **beside** the spill (using `BYROW`/`LAMBDA`).
- `modReports` only builds layout and formulas once; it computes nothing at run time.

**Validation, three layers:**
1. Data Validation (static lists, numeric and date constraints).
2. Dependent dropdowns on **per-list staging columns in `_Work`**, tagged in row 1 (`PRN|<sheet>`, `TEC`, `STK|<model>`). Rebuilt on change, never on selection. Paper Stock is locked while Printer is blank.
3. `Worksheet_Change` → `OnCellChanged` for cross-field rules, with `Status` persisting the verdict. *Check workbook* sweeps everything. Export state is counted separately from data problems.

**Export.**
- CSV via `SaveAs xlCSVUTF8`, containing **all columns including the snapshot block** (the one irreversible decision). `H_Issues` is the only exclusion.
- Header block, then a header row; dates as `yyyy-mm-dd hh:nn:ss`.
- Named `PrintCosts-<SITE>-<LOC>-yyyymmdd-hhmm.csv`, written beside the workbook — see §5 on what "beside" resolves to.
- Export state is a **derived fingerprint** (row count + CountA + Sum(Quantity) + Sum(Chargeable) + Max(Date/Time)) compared to the last export. Import is a later revision.

**Protection.** Sheets `UserInterfaceOnly:=True` (re-applied on open, since it is not persisted). Workbook structure is unprotected so users can duplicate sheets. Sheet password `printlog`. VBA project locked for viewing.

**Modules:** `modMain`, `modJobs`, `modSnapshot`, `modValidation`, `modLists`, `modRegistry`, `modReports`, `modExport`, `modVersion`, `modPicker`, `modInit`, `modCatalog`, `modSettings`, `modProtect`, `modUtils`. Classes: `clsDict`, `clsStock`, `clsPrinterDef`. That is 15 modules + 3 classes + `ThisWorkbook` — **eighteen importable files**, which matters on Mac (§7).

## 5. Conventions and gotchas (each one came from a real defect)

- `Option Explicit` everywhere. Public entry points use depth-counted `AppOff`/`AppOn`, `On Error GoTo Fail`, and restore state on exit.
- **`EnableEvents = False` during our own writes**, always. Anything that needs to know data changed must **derive** it, not listen for it.
- **Derive state, do not track it.** This is applied to the marker cell, the registry, the export fingerprint and the report figures.
- Write formulas from VBA with **`.Formula2`**, never `.Formula`.
- Reports address `_Data` columns **by header name**, not position.
- Criteria expressions in `FILTER` must be **arrays**. A bare `TRUE` gives `#CALC!`, so seed with `Job ID <> ""`.
- Coerce both sides of date comparisons with `*1` (text dates silently match nothing).
- **Never `Val()` on a cell — date *or* number.** `Val` takes a `String`, so the value round-trips through the machine's decimal separator and is parsed back expecting a US point. On a comma-decimal locale every paper cost, consumable rate and stock dimension reads as **zero**, silently. Use `modUtils.DateSerialOf` for dates and `modUtils.NumOf` for numbers; both read `.Value2`. The design doc stated this for dates only because that is where it was first caught. AT-15 runs on a Mac, which is where it would have surfaced.
- **`ws.Move Before:=` only moves a sheet EARLIER.** Moving one that already sits before the target is a no-op — removing it shifts everything down by one and it lands back where it started. Moving later needs `After:`. This shipped the report tabs reversed in every build up to 0.7.0.
- **`ThisWorkbook.Path` is not reliably a filesystem path.** For a OneDrive-backed workbook Excel may report the service URL (`https://d.docs.live.net/...`), and this project lives in OneDrive. Concatenating that with `Application.PathSeparator` produced a hybrid, `SaveAs` did not raise, and the export announced success having written nothing — then stamped the location as exported, so the §3.3 warning standing between a sheet deletion and the records said the opposite of the truth. `modExport` now maps the URL back to the local OneDrive root via `Environ$("OneDrive")`, refuses with an explanation when it cannot, and **verifies the file exists before stamping anything**. That last check is the durable part.
- A staging sheet used for CSV export must be formatted as **Text** (`NumberFormat = "@"`), or dates and `1.0` get re-parsed.
- `On Error Resume Next` around a block only, never around a batch of independent writes.
- **Read `Err.Number`/`Err.Description` first thing in a handler.** Every form of `On Error` resets the `Err` object, and `RelockSheet` runs two of them — so reading `Err` after reprotecting reports `0:` and loses the only description of what went wrong.
- Quiet mode (`modUtils.SetQuiet`): `Say` collects messages instead of `MsgBox`, and `Ask` returns **False**. This is what lets a test read a destructive command's confirmation text while guaranteeing it aborts. It is *not* what makes the tests non-destructive overall — see §7.
- `modUtils.Say` takes what/why/what-to-do as three separate arguments (spec §15).
- Excel auto-renames duplicated `ListObject`s (`tblJobs_ANNEX` → `tblJobs_ANNEX14`), which is why refresh parks tables first.
- `_xlfn.` prefixes in stored formulas are **required**. Excel's `Names` collection shows phantom `_xlfn.*` placeholders it never saves (this caused a false defect, O10).

## 6. Versioning

- Workbook version is `modVersion.APP_VERSION` → `SET_APP_VER`; numbering is `0.<phase>.<revision>`, reaching `1.0.0` when acceptance passes. `0.7.1` is a revision of phase 7, not a new phase — the phase digit moves only when a phase is built.
- Schema version is `modUtils.SCHEMA_VER` → `SET_SCHEMA`, currently `1.0`, stamped on every job row as `S_SchemaVer`. It bumps only when a column is added, removed or renamed, not on reorder. **0.7.1 did not touch it.**
- `StampBuild` (build script only, after `InitialiseWorkbook`) writes `BUILT` and `BUILT_BY`.

## 7. Build and test tooling (Windows only)

**`build.ps1`** validates the source, backs up the existing `.xlsm`, copies the source `.xlsx` to `%TEMP%`, imports all modules, pastes `ThisWorkbook.cls`, runs `InitialiseWorkbook` then `StampBuild` in quiet mode, saves to `%TEMP%`, copies into `src\`, and prunes backups to five. Running a macro over COM also forces the whole project to compile, so a build failing is a compile check.

Layout: `src\PrintCosts.xlsx` (everything a file can carry), `src\PrintCosts-VBA\*.bas|*.cls` (authoritative VBA source). Requires "Trust access to the VBA project object model" once.

Supporting scripts: `probe.ps1`, `verify.ps1`, `test-duplicate.ps1` (AT-13), `test-reports.ps1` (AT-10/12), `test-export.ps1`, `test-validation.ps1` (AT-11/14/16), `filecheck.ps1`, `lockcheck.ps1`, `xlfnscan.ps1`, `prune-backups.ps1`, `diag.ps1`.

**Backups** are `PrintCosts.<yyyyMMdd-HHmmss>.bak.xlsm` in `src\`, made by `build.ps1`, pruned to five by `prune-backups.ps1` — which sorts by **Name**, so the datetime in the name is load-bearing. The working file is always `PrintCosts.xlsm`. Do not introduce any other naming; anything else in `src\` matching neither pattern is not a backup.

**Hard-won cautions:**

- **OneDrive AutoSave nearly destroyed the source `.xlsx`.** A build must never open or mutate its own input. Four guards now exist (works on a `%TEMP%` copy; refuses if the `.xlsx` contains VBA; compares source timestamp before and after; `AutoSaveOn = $false`). **Do not weaken these.**
- **The test scripts must work on a `%TEMP%` copy too, and now do.** Closing without saving is not a guarantee for this workbook, and neither is turning AutoSave off. `Workbook_Open` does real work on every open — `ProtectAll` unprotects and reprotects all fourteen sheets, plus `Invalidate` and `HealButtons` — so it is dirty the instant it opens. Excel holds the file through its **cloud-backed handle**, so AutoSave is on and commits that immediately: `$wb.Saved` reads **True** on the very next line after `Open`, despite every sheet having just been modified. The bytes land at `Close`/`Quit`. That is precisely why `Close($false)` cannot help — it discards *unsaved* changes and by then there are none — and why setting `AutoSaveOn = $false` after `Open` is too late, with no earlier hook available. It fails *by construction*, so do not re-add that guard to the test scripts; `build.ps1` keeps it only because it opens a copy that is not cloud-backed. **None of this is the sync client**: it reproduces with syncing paused and settled. Left unfixed it rewrote the artefact under test, and phases 5–7 were run that way.
- **Do not open `src\PrintCosts.xlsm` read-write for anything, including diagnostics.** `verify.ps1` and `probe.ps1` open **read-only**, which is the other way to be safe.
- **Do not kill Excel while it holds the workbook.** Observed once and not yet fully explained: after an orphaned Excel process was killed, the next open wrote a **stale cached session** over a freshly rebuilt file, reverting it by several builds. Four `PrintCosts (N).xlsm` files stamped **0.4.0** also appeared in `src\` on the same day, which is the same behaviour in another form. Suspected mechanism is the Office document cache (`%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache`); it has not been touched or confirmed. Let scripts quit Excel themselves.
- **Verifying "the tests did not mutate the deliverable" needs a hash, and a hash is not enough on its own.** Compare SHA-256 either side of a run *and* check content invariants, because an unchanged hash only proves that run was clean, not that the file was correct going in.
- **`BuiltinDocumentProperties` does not resolve over COM from PowerShell**; it returns blank without raising. Verification scripts read the OOXML package directly. **When a check and the thing it checks disagree, suspect the check.** This applies to PowerShell too: `$null.Anything` returns `$null` rather than raising, so a probe written that way can "prove" the opposite of the truth. Confirm VBA behaviour by running VBA.
- Excel for Mac has no ActiveX, COM, `Scripting.Dictionary`, `FileSystemObject`, `ADODB.Stream` or reliable `FileDialog`. Mac is handled via `clsDict`, Form Controls only, a worksheet picker, and manual module import. **`SETUP.md`'s import list is the whole build on Mac** — it listed fourteen of eighteen modules until 0.7.1, omitting `modRegistry`, `modReports`, `modExport` and `modVersion`, which produced a project that would not compile. This blocked AT-15.

## 8. Open and deferred items

- **O7 (deferred):** aggregation granularity for the future master workbook. Options: job-level, summary-level (loses cross-site student totals), or job-level with student number but no names. Decide on data-protection grounds.
- **Post-v1:** multi-workbook aggregation (§11). It reads exported CSVs, not source workbooks. v1 must preserve these invariants: globally unique job IDs, real `SET_SITE_ID`, `S_SchemaVer` on every row, consolidated `_Data`, never-reassigned location codes, complete per-location export.
- **Data protection:** export multiplies where student names and numbers live outside workbook protection. Worth flagging in whatever operating procedure is written.
- **Unrecorded in the design doc:** the stale-cached-session behaviour in §7. One clear observation and a suspected mechanism is not enough to write in as fact; if it recurs, confirm it and add it to §12.1.
- **Unused settings:** `SET_CURRENCY`, `SET_FOOTER`, `SET_FY_START` are defined names no code reads. Phase 8 decides.
- Deployment constraint, not a defect: `_Data` shows `#NAME?` on Excel without dynamic arrays (out of scope per D12).

## 9. Working notes for Claude Code

- Treat the `.xlsx` plus `PrintCosts-VBA\*` as source of truth; the `.xlsm` is a build output.
- Rebuild and run all four `test-*.ps1` after any VBA change, **one at a time**. They are regression tests as of phase 7. Running them in a tight loop lets Excel share one process between scripts, and a later script then reads an earlier one's in-memory state as though it were the build.
- Check the deliverable's content after a run, not just that the run left it unchanged.
- After changing any column's name or meaning, bump `SCHEMA_VER`; a same-name meaning change is a discipline problem no mechanism will catch.
- Never guess at Excel behaviour over COM; confirm by reading the file package or running a test — in the language whose behaviour is in question.
- The VBA source files are not uniform: `modExport`, `modRegistry`, `modReports` and `modVersion` are LF-only, the rest CRLF. Preserve each file's own convention when editing.

## 10. What 0.7.1 changed, and what was actually verified

0.7.1 is a review pass over the whole source against the design document, before building on it. **No schema change and no new behaviour.** Nine code fixes: the reversed report tabs; `Val()` on numeric cells; two blanket `On Error Resume Next` handlers over batches of writes (`LogAudit`, `ExportSig`); a crash in `WriteRegistry` when no print rooms exist; a locale-unsafe `Money()`; the export writing nothing while reporting success; and the test scripts rewriting the workbook they test. Three documentation fixes: SETUP.md's module list, `SET_RoundDP` → `SET_ROUND_DP` in §4.1, and the design doc's own account of the AutoSave mechanism, which named the wrong cause. Full detail in the 1.8 changelog entry at the top of the design doc.

**Confirmed by reading the built file, not assumed:** the §4.1 job-row formulas; the §7.1 consolidated formula including the `IF(SEQUENCE(ROWS(…))>0,"CODE")` broadcast guard and outer `IFERROR`; the 21 + 12 column layout and the marker cell; `H_Issues` implementing all seven §6.3 cross-field rules; freeze panes and filter buttons; the absence of any conditional formatting; and that the source `.xlsx` contains no VBA project.

**Regression run at 0.7.1:** all four scripts pass, figures identical to the phase 5–7 baselines (totals £494.56 gross / £479.06 chargeable; CSV 2197 bytes, 32 columns), each verified by SHA-256 to have left `src\PrintCosts.xlsm` byte-identical.

**Not verified, and worth knowing:** the OneDrive-URL export path is no longer exercised by any script, since the tests now run from `%TEMP%` where the path is already local — it needs a manual click of **Export…**. The `Environ$("OneDrive")` mapping is Windows-shaped; on Mac `ThisWorkbook.Path` should already be local so the branch should not trigger, but that is reasoning, not evidence.
