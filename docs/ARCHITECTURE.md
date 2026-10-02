# University Printing Cost Management Workbook — Architecture & Implementation Reference

**This is the single source of truth for this project's design and implementation.** It supersedes and replaces `stage1-design-architecture.md`, `HANDOFF up to stage 7.md`, `snaglist-stage7handoff.txt` and `snaglist_2026-09-21.rtf`, which described the same system at different, now-stale points in time and had started to disagree with each other and with the code. Their content is folded in below; nothing that still mattered was dropped, but the historical changelog they were built from has been compressed into §16 rather than repeated inline.

**Source of truth for the code itself:** `src\PrintCosts.xlsx` (everything a file can carry except VBA) plus `src\PrintCosts-VBA\*.bas` / `*.cls` (the authoritative VBA source). `src\PrintJob.xlsm` is a **build output** — never hand-edit it. `src\PrintCosts-VBA\SETUP.md` is the separate, actively-maintained *operational* guide (how to build, test, and set up the workbook) and is not duplicated here; this document is design and architecture, SETUP.md is procedure.

**Naming inconsistency, flagged not fixed (2026-09-27).** The project folder is `PrintJob`, but the source workbook is `PrintCosts.xlsx` and the VBA source folder is `PrintCosts-VBA` — a mismatch that predates this note. `build.ps1`'s output was renamed to `PrintJob.xlsm` (2026-09-27, direct user request) specifically because that rename is free: `HealButtons` (§4.2) already re-qualifies every button's `OnAction` on open, so the workbook self-heals under any filename. Renaming the `.xlsx` source, `PrintCosts-VBA` folder, or the `PrintCosts-*` export/backup file prefixes (`modExport.bas`, `modBackup.bas`) would be a much larger, higher-risk pass with no equivalent self-healing mechanism, and was deliberately left alone. Worth a proper look in the future if the inconsistency keeps bothering people.

**Current state, as verified against the actual VBA source on 2026-09-28:**

- Workbook version reported in-code: `0.10.7` (`modVersion.APP_VERSION`) — `0.8.0` reflected the NextId fix, Export/Import, the Reports rework and the 2026-09-21 snag list; `0.8.1` closed out phase 8's own original scope (§16.1); `0.9.0`–`0.9.20` are phase 9's packages and post-package addenda — see `docs/CHANGELOG.md` for the fixes beyond `0.9.11`; `0.10.0` is the printer/paper compatibility rework (§16.5): capacity-based printer/stock compatibility in place of the old paper-family width bands, and student-supplied paper stock; `0.10.1` makes the two student-supplied stocks built in (no longer `tblPapers` rows) and makes `Supplied by student = Yes` force a zero paper cost; `0.10.2` adds Clear table on the Technicians/Printers/Papers sheets; `0.10.3` removes the paper family concept (Papers' `Family` column, `tblPaperFamilies`, the job-row `S_Family` snapshot) — Papers' `Measure` (`Sheet`/`Roll`) is now the one input (see `docs/CHANGELOG.md`); `0.10.4` gives technicians, printers and papers site-prefixed, auto-allocated IDs (§3.2); `0.10.5` makes the Reports name toggle export-only ("Export names", §8); `0.10.6` adds the Reports "Mark all as…" Paid/Unpaid cluster and an at-least-one-filter safeguard shared with Delete visible records (§10.9); `0.10.7` makes the Reports Paid column directly editable with the Yes/No dropdown (§10.10).
- Data schema version: `1.3` (`modUtils.SCHEMA_VER`) — `1.0` since inception; bumped to `1.1` by phase 9's Paid column (snag 1c, §5), which also surfaced and fixed a gap in how `SET_SCHEMA` stayed in sync with the constant (§3.3, §3.5); bumped to `1.2` by the printer/paper compatibility rework's new job-row "Sheet size" column (§5, §16.5); bumped to `1.3` by dropping the `S_Family` snapshot column when the paper family concept was removed (§5, `docs/CHANGELOG.md`).
- Everything in the original design document's phases 1–7 is built and verified.
- Phase 8's original scope (visual polish) is **fully built** — see §16.2.
- The "larger changes" plan (`snaglist-stage7handoff.txt`: NextId fix, Import/Export rework, Reports rework, bulk delete) is **fully built**.
- The 2026-09-21 seventeen-item snag list (`snaglist_2026-09-21.rtf`) is **fully built**.
- **Phase 9 is the 2026-09-22 post-phase-8 snag list — in progress, see §16.4.** Renumbered from the original plan: this document previously reserved "Phase 9" for formal acceptance testing, but a real, larger batch of functional work arrived first, so acceptance testing is now **Phase 10** (every "Phase 9" reference to acceptance testing below has moved accordingly).
- Phase 10 (formal acceptance testing on Windows and Mac, test report) has **not** run.

---

## 1. What this is

An Excel `.xlsm` (Microsoft 365 only, Windows and Mac) for logging print jobs per print room and costing them. One location sheet per room, plus **Summary** and **Reports** report sheets. VBA handles entry, validation, snapshotting, the location registry, and export/import. Reports are live dynamic-array formulas over a consolidated range, not VBA-generated output — the whole reporting layer can only ever be a wrong *sheet*, visible immediately, never a wrong *figure* on a sheet that looks right.

---

## 2. Decisions taken at clarification

| # | Question | Decision |
|---|---|---|
| D1 | Student name/number consistency with no master list | No Students register. On entry, cross-check the pair against all existing rows in all locations; warn (never block) on conflict |
| D2 | Mechanism for historical cost integrity | Snapshot the resolved rates onto each row at creation; visible costs remain live formulas over those snapshot values |
| D3 | Contents of "global settings" | Currency format + rounding rule; organisation/report header metadata; site identity; export folder. No VAT, no overhead markup |
| D4 | Per-row *Now* and *Remove* controls | Toolbar of Form Control buttons above each table, acting on the selected row; plus double-click to stamp date/time |
| D5 | How reports discover location sheets after duplication | Marker cell on each location sheet; a **Refresh Locations** command rescans and rebuilds a hidden registry |
| D6 | Papers "Cost" means £/sheet or £/metre | One `Cost` column, with an adjacent calculated `Cost unit` label driven by the stock's `Measure` |
| D7 | Units for transaction print width | Millimetres, matching stock width |
| D8 | Presentation of disregarded costs | Gross, Disregarded and Chargeable as separate columns; consumption unaffected throughout |
| D9 | Custom sheet sizes | Standard A-sizes remain a picker; explicit custom width × height in mm also permitted |
| D10 | Multi-select for printer families and location printers | Worksheet-based picker with a multi-select list, writing a delimited string to a single cell |
| D11 | Deployment / file-sharing model | Single workbook is the product. Multi-workbook aggregation is a planned feature — D17, §12 |
| D12 | Minimum Excel version | **Microsoft 365 builds only**, Windows and Mac. Dynamic array formulas may be relied upon |
| D13 | Report construction | Reports are live formulas over a consolidated spill range, not VBA-generated output (§8) |
| D14 | Multi-site collation mechanism | Superseded by D17/D18: collation, when built, reads the **exported CSV files** (§12.2) |
| D15 | `.xlsm` delivery | A PowerShell build script drives Excel over COM. Manual import route retained for Mac (SETUP.md) |
| D16 | Version identity | **Two independent version numbers** — a workbook version for the build, and a data schema version for the shape of the records (§2.6) |
| D17 | Scope of v1 | **Multi-workbook aggregation is a planned feature, not in v1.** v1 delivers a single workbook complete in itself, preserving the invariants an aggregator needs (§12.1). O7 deferred with it, though see §12.4 — the built Export/Import mechanism now covers most of what a job-level aggregator needs, manually |
| D18 | Deleted location sheets | **Export-based mitigation** rather than a shadow copy: per-location CSV, a visible per-sheet status, surfaced by *Check workbook* and *Refresh Locations* (§10.4, §3.3) |
| D19 | Department/"charge to" terminology (2026-09-22 snag 4b) | **Deliberately free-text, no new catalogue table.** The Student Name/No columns already accept anything typed into them (D1 — no authoritative student list, cross-checked only for consistency, never validated against a register) — a department name behaves identically with no code change. User-facing wording that names the field ("Student name" filter labels, the Remove-row confirmation, the name/number consistency warning, the About blurb) was given a light "Student/Department" pass so it reads sensibly either way; the underlying column headers stay `Student Name`/`Student No` deliberately, since renaming them would break every header-name lookup that already reads them (§5, §8.3, §10.4) for no real capability gained. **Fallback if revisited:** a `Departments` catalogue table, mirroring `tblTechnicians`, if the free-text approach turns out not to be enough post-phase-9. |

### 2.1 Interpretations — confirmed at sign-off

- **I1 — Sheet-stock consumable area scales with quantity.** Area is `sheet area × number of sheets`.
- **I2 — Paper stock carries a measure.** Each stock has an explicit **Measure** of `Sheet` or `Roll` (a dropdown on Papers), and behaviour keys off that attribute. There is no separate "family" concept any more (removed 0.10.3, `docs/CHANGELOG.md`).
- **I3 — Sheet-stock quantities are whole numbers.** Roll quantities are decimal metres; sheet quantities integers ≥ 1.
- **I4 — Rounding is applied per cost component at row level**, so a printed report's rows add up to its totals.
- **I5 — AT-11 is a warning, not a block.** With no authoritative student list the first entry of any student is unverifiable, and legitimate name changes would trigger it. Verified in phase 7.

---

## 3. Data model

### 3.1 Entities and relationships

```
Settings (global key/value)
PaperType ──┐
            ├──< PaperStock (Measure: Sheet/Roll)
StandardSize┘         │
Technician ───────────┼──>PrintJob
                      │
                      └───────────────────── Printer ──── ConsumableType
                                  │
                              Location (one sheet each)
                                  │
                              Site (one workbook each)
```

- A **PaperStock** belongs to exactly one **PaperType** and has a **Measure** (`Sheet` or `Roll`, chosen from a dropdown on Papers) that decides whether it is counted in sheets or metres. There used to be a separate **PaperFamily** entity carrying a *Measurement basis*; once the width-band families merged into one `Roll` family it only ever mapped `Sheet`→`Sheet` and `Roll`→`Roll`, so it was removed (0.10.3, `docs/CHANGELOG.md`).
- A **PrintJob** references one Technician, one Printer and one PaperStock, and belongs to one Location.
- A stock is valid for a printer purely by size, not by any grouping of stocks: a Roll stock's width against the printer's **Max roll width mm**, a Sheet stock's dimensions (either orientation) against the printer's **Max sheet size** (a Standard Size). A **Printer** does not reference paper groupings at all — the old `Supported families` multi-select is gone (§3.3, §16.5).
- The two built-in **"Supplied (Roll)"/"Supplied (Sheet)"** stocks (student-supplied paper) skip the size comparison entirely — compatible with any printer that has the matching capacity field set at all, since the real size is entered per job rather than read from a catalogue row (§5, §16.5). They are added in code by `modCatalog.LoadCatalog` (`AddBuiltInStocks`), are **not** `tblPapers` rows, and their names are reserved (`IsBuiltInStock`; `OnPaperEdited` rejects them). A `tblPapers` row with **Supplied by student = Yes** (e.g. a bulk delivery of paper for many jobs) is an ordinary stock with its own catalogue size, compatibility and validation — the flag only means its paper cost is always 0 (`LoadCatalog` forces `Cost = 0`; `OnPaperEdited` also resets the visible Cost cell).
- A **Location** permits a subset of printers.
- A **Site** is one workbook file. In v1 there is one site per workbook and nothing collates them automatically (§12) — though see §12.4 for the manual route that now exists.

### 3.2 Identity and referential integrity

Every configuration record carries a **stable ID** — `<SITE>-TCH-0001`, `<SITE>-PRN-0001`, `<SITE>-STK-0001` — assigned once and **never re-issued**, even after deletion. Job rows store the ID in their snapshot block, so renaming a printer or technician cannot orphan or silently re-point a historical record. IDs come from the highest ever issued, never the row count.

The visible cell on a job row holds the **display name**, because that is what a technician recognises in a dropdown. Configuration names are therefore unique within their table, enforced on entry.

**Job IDs are `<SITE>-<LOC>-00001`.** The site component exists so records from several workbook copies can never be conflated. It costs nothing in v1 and is one of the invariants aggregation depends on (§12.1).

Because a location's code is baked into its job IDs, **an existing code is never reassigned** (§4.2).

> **Integrity rule:** configuration rows must be deactivated, not deleted (§9). Deleting a row history references is blocked with a message naming the number of dependent records.

**Job ID allocation is a persisted high-water mark, not a row scan (fixed 2026-09-21).** Before this fix, `NextJobId` allocated "highest ID currently present on the sheet, plus one" — so deleting the highest-numbered job row let the next new job silently reissue that same Job ID, breaking the "globally unique, never re-issued" invariant that export, import and any future aggregation all depend on. This was fixed **before** Import was built, because Import needed the invariant to actually hold.

Current mechanism (`modRegistry.NextJobId`, `modRegistry.bas:570`):

- A **`Job ID HWM`** column lives on `_Registry.tblLocations`, one value per location, added via `EnsureColumn` like `Last export`/`Export sig` (§3.4) and preserved across `RefreshLocations` rebuilds the same way (captured into a keyed collection before the registry is wiped, restored per sheet name afterward).
- `NextJobId` reads the persisted value (`JobIdHWM`), takes `Max(persisted, ScanMaxSuffix(...))` as a **floor only** — the row scan can never move the effective value backwards, since deleting a row only lowers what the scan finds, never the persisted figure — increments by one, writes the new value back, and returns the formatted ID.
- The scan still runs, but only earns its keep in two cases the persisted value alone can't cover: seeding the HWM the first time the column exists on an already-populated sheet, and raising the mark after **Import** has just written rows under the same prefix with higher numbers than anything allocated locally so far.
- Verified by `test-nextid.ps1`: deleting the top row must not reissue its ID, and the mark must survive `RefreshLocations`.

**Catalogue IDs (technicians, printers, papers) follow the same rule (2026-09-29).** `TechID`, `PrinterID` and `StockID` are `<SITE>-TCH-0001`, `<SITE>-PRN-0001` and `<SITE>-STK-0001`: the site (`SET_SITE_ID`), a table code and a four-digit number, allocated by `modCatalog.NextCatalogId`. The site is in the ID for the same reason it is in a Job ID: two workbooks that each numbered from 1 would issue the same ID to different things, and Restore workbook (§10.7) matches catalogue rows by that ID, so it would overwrite one site's printer with another's. With the site in the ID, combining another workbook's configuration appends rather than overwrites.

- **Nobody types an ID.** `AddCatalogRow` fills it as the row is created; a row typed under the table gets one once it has a name (`ThisWorkbook.Workbook_SheetChange` -> `modCatalog.OnPaperEdited` / `OnCatalogEdited`); Setup (`EnsureCatalogIds`) fills any named row that has none. An existing ID is never changed, and the ID column is locked (`modInit.UnlockTableBody`, re-applied on every Setup run) so it cannot be edited by hand either; VBA still writes it, since every write is bracketed by `UnlockSheet`/`RelockSheet` and the sheets use `UserInterfaceOnly` protection.
- **Counters.** One persisted high-water mark per table - `TECH_ID_HWM`, `PRINTER_ID_HWM`, `STOCK_ID_HWM` - as rows in `tblSettings` (there is no per-location part, so not `_Registry`). Same floor rule as `NextJobId`: `Max(counter, ScanMaxSuffix)`, so deleting the newest row never reissues its ID. Their Notes start `Read-only`, which locks the cell and makes Restore workbook skip them, so a backup can never wind a counter back.
- **The template's sample rows ship with blank IDs**, so every build issues its own site-prefixed ones. Fixed IDs such as `TEC-001` in the template would have been identical in every workbook built from it, which is the collision this design exists to prevent.
- **Names are the lookup key everywhere** (`clsDict.Add` replaces on a duplicate), so two rows with one name would shadow each other. When a restore brings in a row whose name is already used by a *different* row, it is kept under its own ID and its site is added to the name: `HP T730` becomes `HP T730 (SITE2)` (a number is added if even that is taken). See §10.7.
- Verified by `test-catalogids.ps1`.

### 3.3 Configuration tables

| Table | Columns |
|---|---|
| `tblSettings` | Key, Setting, Value, Notes |
| `tblPaperTypes` | Paper type, Active |
| `tblStandardSizes` | Size name, Width mm, Height mm |
| `tblConsumables` | Consumable type, Active |

Each setting is exposed as a workbook-scoped defined name so formulas and VBA reference meaning rather than cell addresses: `SET_SITE_ID`, `SET_SITE_NAME`, `SET_ORG`, `SET_DEPT`, `SET_CURRENCY`, `SET_ROUND_DP`, `SET_EXPORT_FOLDER`, plus read-only `SET_SCHEMA`, `SET_LASTREF`, `SET_APP_VER`, `SET_BUILT`, `SET_BUILT_BY`, and the three catalogue ID counters `SET_TECH_ID_HWM`, `SET_PRINTER_ID_HWM`, `SET_STOCK_ID_HWM` (§3.2).

`modSettings` addresses every setting as `SET_<KEY>`, so the defined name — not the row — is what makes an entry a real setting. Settings that don't ship in the `.xlsx` are self-provisioned by VBA via `modVersion.EnsureSetting`: `APP_VER`/`BUILT`/`BUILT_BY` (`modVersion.EnsureVersionSettings`), `SCHEMA` (`modVersion.EnsureSchemaSetting`, new 2026-09-22 — see §5's Paid column note for why) and the three catalogue ID counters (`modCatalog.EnsureCatalogIdSettings`). The rest — `SITE_ID`, `SITE_NAME`, `ORG`, `DEPT`, `ROUND_DP`, `CURRENCY`, `EXPORT_FOLDER`, `LOC_REDUCED_VIEW`, `LOC_REDUCED_COLUMNS`, `COST_COLS_HIDDEN` — ship directly in `PrintCosts.xlsx`, rows and `SET_` names both; no setup routine adds them.

**Print Technicians** — `tblTechnicians`: TechID, Name, Department, Active.
**Printers** — `tblPrinters`: PrinterID, Model, Consumable type, **Cost per m2**, **Max roll width mm**, **Max sheet size**, Active. The last two (printer/paper compatibility rework, §16.5) replace the old `Supported families` multi-select: a printer takes roll stock iff the first is set, sheet stock iff the second is set (a Standard Size name), and both can be set on the same printer. `modCatalog.EnsurePrinterCapacityColumns` adds them; `MigratePrinterCapacities` derives their initial values from whatever the printer was compatible with under the old family-list rule, then removes the legacy column.
**Papers** — `tblPapers`: StockID, Description, Paper type, **Measure** (`Sheet`/`Roll` dropdown), Size mode, Std. size (renamed from "Standard size" 0.9.14 — `modCatalog.EnsureStdSizeColumnName`), **Width mm**, **Height mm**, **Cost**, *Cost unit* (calc), **Supplied by student** (Yes/No, default No; Yes = paper cost always 0, §16.5), Active. The two student-supplied stocks `Supplied (Roll)`/`Supplied (Sheet)` are **not** rows of this table — they are built in (`modCatalog.AddBuiltInStocks`, `SUPPLIED_ROLL`/`SUPPLIED_SHEET`), reserved names, `Cost = 0`, no catalogue size. Earlier builds shipped them as real rows; `modCatalog.RemoveLegacySuppliedRows` (called from Setup) deletes those, and jobs already recorded keep their frozen `S_*` snapshot. The Summary sheet's Type/Unit lookups fall back to fixed values for the two names (`modReports.BuildSummary`).

All dimensions are stored in millimetres; metres appear only as a transaction quantity for roll stock.

**Adding/removing configuration rows (snag item 17, resolved).** `modCatalog.AddCatalogRow` / `RemoveCatalogRow` back Add-row/Remove-row buttons on Printers, Papers and Print Technicians (full-size buttons) and compact `+`/`-` button pairs above the three Settings lookup tables — Paper types, Size, Consumable (button names kept short deliberately: `Button.Name` silently truncates at 32 characters and raises error 1004 at 33+). `AddCatalogRow` reuses an existing single blank row if present, else `ListRows.Add`, which keeps formatting and validation automatically. `RemoveCatalogRow` operates on the row under the current selection, shows every column's value in the confirmation, and states explicitly that jobs already recorded against a removed row keep their frozen snapshot prices — deactivation-not-deletion (§3.2) is enforced by the confirmation text, not by disabling the button.

The smaller lookup tables on Settings (paper types, standard sizes, consumables) deliberately don't get their own buttons: their cells are unlocked (§9.4), so a row added the normal Excel Table way (Tab at the last cell, or right-click → Insert → Table Rows) keeps formatting and validation without VBA's help.

### 3.4 Hidden working sheets

`xlSheetVeryHidden`, so they cannot be unhidden from the tab context menu:

| Sheet | Purpose |
|---|---|
| `_Registry` | `tblLocations`: SheetName, Code, Name, Department, Rows, First date, Last date, State, **Last export**, **Export sig**, **Job ID HWM** |
| `_Work` | **Validation staging only** — §6.2, and reused for Reports filter dropdowns (§8) |
| `_Data` | The consolidated job range (§7.1) |
| `_Audit` | `tblAudit`: When, User, Action, Location, Detail (§10.3) |
| `_Export` | Reserved. The export writes to disk (§10.4) |
| `_Picker` | Backing sheet for the multi-select picker (§8.3 of the VBA architecture section, §9.3 below) |

**Why `_Data` is separate from `_Work`.** `modLists` claims `_Work` one column at a time, tagging each in row 1 and scanning up to 200 for a free slot, and will overwrite anything else parked there.

**Why the export stamp lives in the registry.** `Last export`, `Export sig` and `Job ID HWM` are keyed by sheet name, not stored in the location's config block. A sheet-scoped cell would copy with a duplicated sheet, and the copy would inherit "already exported" (or a Job ID counter) while holding records never exported anywhere, or none at all. A registry row means a duplicate gets a fresh row and correctly reads as unexported with a fresh counter; a renamed sheet likewise. Both are the conservative direction.

The shipped `.xlsx` predates the export and NextId-fix features, so these columns are added to the existing registry table by `EnsureColumn` rather than being part of its creation.

### 3.5 Version identity (D16)

| | Meaning | Changes when | Held in |
|---|---|---|---|
| **Workbook version** | This build — code, layout, formulas | Every rebuild worth distinguishing | `modVersion.APP_VERSION` → `SET_APP_VER` |
| **Data schema version** | The shape of the stored records | Only on a genuine shape change | `modUtils.SCHEMA_VER` → `SET_SCHEMA`, stamped on every job row as `S_SchemaVer` |

The schema version is load-bearing: it is what an importer or aggregator checks before reading records, and what a job row carries so its origin is recoverable. Conflating the two would force a schema bump on every cosmetic change and make that check meaningless.

**What bumps the schema version.** Adding, removing or renaming a column bumps it. **Reordering columns does not** — exports carry a header row and are read by column *name*. The case the number cannot protect against is a column whose *meaning* changes while its name stays the same; that is a discipline, not a mechanism. The `Job ID HWM` registry column and the `EXPORT_FOLDER` setting are new, but neither is a job-row column, so neither touched `SCHEMA_VER`. **The Paid column (2026-09-22, §5) is the first one that did** — `1.0` → `1.1` — and it surfaced that `SET_SCHEMA` had no code path keeping it in sync with the constant (fixed, `modVersion.EnsureSchemaSetting`, §3.3).

**Numbering is `0.<phase>.<revision>`.** It reaches `1.0.0` when acceptance passes. Bumped to `0.8.0` on 2026-09-21 (§16.1) to reflect the NextId fix, Export All Locations, Import, the Reports rework and the full snag list — a phase-digit move, not a revision, because that batch is real new functionality rather than fixes with no behaviour change.

**Where the version is visible**, in ascending order of effort: the file's **document properties** (Explorer, Finder — without opening the file); `tblSettings`' own `Workbook version` row on the Settings sheet; and the **About button**, whose popup (`modVersion.ShowAbout`) is now the only place the workbook's name, version, author and a brief description are shown together — the on-sheet About block this button used to point to was removed in 0.9.12 (§16.4's addendum) once the popup grew that same prose. All now report `0.8.0` once the workbook is rebuilt from source (`build.ps1` — the version lives in the VBA source until a build stamps it into the `.xlsm`).

**What stamps what.** `EnsureVersionSettings` writes `APP_VER` from the constant on every setup run, because the code in the workbook *is* its version. `BUILT` and `BUILT_BY` are written only by `StampBuild`, called by the build script **after** `InitialiseWorkbook` — users re-run setup whenever they add a print room, and stamping from there would reset the build date to whenever someone last duplicated a sheet.

---

## 4. Location-sheet architecture

### 4.1 Layout

```
Rows 1–11   Configuration block, columns A/B only (see the 2026-09-25
            rework below): room name, department, the three per-job
            default selectors, the two disregard-cost defaults (rows 6–7,
            2026-09-29), a blank spacer (row 8), roll-unit setting,
            permitted-printers display, live job count (rows 9–11)
Row 12      Blank gap (2026-09-25 — user-reported: no
            breathing room otherwise)
Row 13      Toolbar (Form Control buttons)
Row 15      Table header
Row 16+     tblJobs_<CODE> body
```

| Cell | Name | Content |
|---|---|---|
| B1 | `LOC_Name` | Room name |
| B2 | `LOC_Dept` | Department |
| B3 | `LOC_DefTech` | Batch default Technician — pre-fills each new job until cleared |
| B4 | `LOC_DefPrinter` | Batch default Printer — same bidirectional filtering as the table cells (§7.2) |
| B5 | `LOC_DefPaper` | Batch default Paper Stock — same bidirectional filtering as the table cells (§7.2) |
| B6 | `LOC_DefDisPaper` | Default disregard paper cost (Yes/No) — moved here from the side panel 2026-09-29; seeds each new job, not cleared by Clear defaults |
| B7 | `LOC_DefDisCons` | Default disregard consumable cost (Yes/No) — same |
| Row 8 | — | Blank spacer, sets the two disregard defaults apart from the settings below |
| B9 | `LOC_RollUnit` | Roll-stock `Qty` entry unit, Metres/Centimetres (default Metres) — always converted to and stored as metres, see below (was B6 until 2026-09-29) |
| B10 | — | Permitted-printers display, `"N printers (comma, separated, list)"` — a read-only `LET` formula over `LOC_Printers` (below), not the named range itself (2026-09-25, swapped with Sheet status/Export, see below) |
| B11 | — | Live count of jobs on the sheet, `=ROWS(tblJobs_<CODE>)` — a genuine formula, re-written whenever the table is renamed so it never goes stale |
| AM2 | `LOC_Code` | Location code (auto-assigned, read-only) — side panel |
| AM3:AM4 | — | Empty since 2026-09-29 (the disregard defaults moved to B6/B7) |
| AM5 | `LOC_Status` | Validation summary for the sheet — side panel (2026-09-25, was B7) |
| AM6 | `LOC_Export` | Export status — side panel (2026-09-25, was B8) |
| AM7 | `LOC_Printers` | Permitted printers, raw semicolon-delimited list — side panel (2026-09-25, was AM5). What `modPicker`'s Select printers… dialog actually writes and `modCatalog`'s compatibility checks actually read; `B7`'s friendly display derives from this, never the other way round |
| AZ1 | — | Marker cell, `PRINTLOC/v1`. Hidden column, never edited |

`LOC_Export` is deliberately separate from `LOC_Status`: validation and export state are independent, and a sheet can easily be valid and unexported at once. Both are derived and rewritten, never typed. `LOC_Export`'s name is created by `modExport` on every refresh rather than shipped in the `.xlsx`, so a duplicated or renamed sheet gets a correct one without hand surgery.

**Field swap (2026-09-25, user request) — Sheet status/Export move to the side panel; the permitted-printers list moves the other way, reformatted for people rather than code.** `modInit.EnsurePrintersDisplay` (new) writes `B7`'s `LET` formula fresh on every run (same self-healing reasoning as the job count), reading `LOC_Printers` wherever it currently is rather than assuming a cell — count-then-list, correct singular ("1 printer") vs plural. `LOC_Printers` itself is untouched functionally: still the exact semicolon-delimited string `modPicker.bas`'s multi-select dialog writes with a plain `.Value =` and `modCatalog.bas`'s two compatibility-check call sites read, just relocated to the side panel under its own label so the picker's write target moves with it automatically (name-based access throughout, never a hardcoded cell). `modExport.EXPORT_CELL` moved from `$B$8` to `$AM$6` — a one-line constant change, since `RefreshExportStatus` already writes its own "Export" label via `Offset(0, -1)` rather than a second hardcoded address, so the label followed the value without a separate edit. The live job count (§ above, added 0.9.17 at `A9`/`B9`) moves up to `A8`/`B8`, the row Export vacated.

**Disregard-cost defaults back on the left (2026-09-29, user request; v0.10.10).** `LOC_DefDisPaper`/`LOC_DefDisCons` left the side panel for `A6:B7`, under Default: paper, with a blank spacer row 8 separating them from the settings below; roll unit, printers display and job count moved from rows 6–8 to 9–11. `PrintCosts.xlsx` ships the table header on row 15 (toolbar row 13, named `TOOLBAR_ROW`), with the roll unit, printers display and job count already at `A9:B11`, the disregard values in `B6:B7` and their names pointing there (two setup-time patchers, `EnsureToolbarGap` and `EnsureConfigLayout`, did this on every build at first and were removed once the file shipped the result); `EnsureJobDefaults` supplies labels/styling and `ConfigValidation` the Yes/No list. Column A/B are never in a reduced-view hide list, so both stay visible. Everything else reaches these cells by name. Nothing migrates an older build: rebuild with `build.ps1` (Refresh Locations alone will not rearrange one).

**Toolbar gap (2026-09-25, user-reported, same day as the field swap above).** With `EnsureJobTableGap` removed (above) the toolbar sat flush against row 9 with no breathing room. `modInit.EnsureToolbarGap` (new, replaces it in miniature) inserts a single blank row above the toolbar — one row, not two, since only the toolbar needs separating from the block now, not a whole extra defaults row. Toolbar moves from row 10 to row 11; the table header follows naturally, from row 12 to row 13. (Superseded: the current layout is in the 0.10.10 entry above, and `EnsureToolbarGap` has since been removed — the `.xlsx` ships the result.)

**Layout rework (2026-09-25) — columns A/B are now permanently off-limits to reduced view; everything that doesn't need daily visibility moved to a side panel.** Direct report: `LOC_RollUnit`'s dropdown (then at `E4`) and the "Default: paper" label (then at `E10`) both sat on column E — table column 5, **Printer**, one of `REDUCED_COLUMNS_DEFAULT`'s own seven hidden columns — so both vanished the moment reduced view was switched on, out of the box, no custom hide-list editing required. Same root cause as the two 2026-09-22 bugs below: Excel's column-hide is whole-column, so anything sharing a column with a hidden table column disappears with it, and `SET_LOC_REDUCED_COLUMNS` is user-editable, so no single column choice can be guaranteed safe forever.

The fix generalises rather than relocates-and-hopes: **columns A and B never appear in any hide list, by construction.** Everything that must stay visible regardless of the reduced-view setting lives only there, stacked one row each rather than spread across column-pairs — the three per-job selectors (moved off a single shared row spread A–F) and the roll-unit setting (moved off D4/E4) now occupy `A3:B6`, where Location code/the disregard defaults/the permitted-printers list used to be. None of those three are used per-job the way the selectors are, so they moved out to the side panel (`AL`/`AM`, columns 38/39 — hand-edited directly into `PrintCosts.xlsx` for both shipped location sheets, named ranges repointed, so a future duplicate inherits the new layout the same way it always inherited the old one) to make room. Room name and Department never moved at all — already on A/B, nothing to fix (Sheet status/Export did move, but later, as part of a separate field swap — see below). `modInit.EnsureJobTableGap` is **removed entirely**: the selectors no longer need a dedicated row of their own, so there's nothing left to open space for. (The table header's row position, and the live job count's cell, both moved again almost immediately after — see the toolbar-gap and field-swap notes below; treat this paragraph as describing the shape of the fix, not the current row numbers.)

**A third bug surfaced by that same reversion, same day.** `modInit.ReorderJobColumns`'s `Range.Cut`/`Range.Insert Shift:=xlToRight` operates on the table's whole header+data row span, not just its columns. With the header briefly back at its originally-shipped row 12 (before the toolbar-gap fix below pushed it to 13), row 13 briefly became the table's first *data* row — exactly where the side panel's Import button had landed (one-per-row stacking: 1, 3, 5, 7, 9, 11, 13). The next column reorder silently dragged it sideways, since a floating shape anchored inside that row range moves with whatever the row-scoped insert does. Fixed for one release by keeping every side-panel row at 11 or above and pairing Import horizontally with Export at a second column (`SIDE_PANEL_COL2`) instead of adding an 8th stacked row — **superseded 2026-09-25, same day as the toolbar-gap fix below, once the second column was itself reported as looking wrong** ("the import button seems to have wandered over to the right"). The real fix (`modInit.DrawOneAtTop`, below) sidesteps row-based positioning entirely rather than trading one row-grid workaround for a column-grid one.

**Batch defaults (snag 1b, 2026-09-22; relocated 2026-09-25 — see above).** `modInit.EnsureJobDefaults` self-provisions `LOC_DefTech`/`LOC_DefPrinter`/`LOC_DefPaper` the same way `modExport.EnsureExportName` self-provisions `LOC_Export` — created (delete-then-add) on every `InitialiseWorkbook` and every `RefreshLocations`, not shipped in the `.xlsx`, so an older location sheet or a freshly duplicated one always ends up with correctly-scoped names. `modJobs.AddPrintJob` copies each non-blank default into the new row (`CopyDefault`) using the same "copy, don't reference" principle as the existing disregard-cost defaults (AT-07/AT-08) — a default changed later never touches a job already added from it. A **Clear defaults** button (`modJobs.ClearDefaults`) empties all three for the next batch — sits at `D4`, roughly centred against the three-row block it clears; column D is inside the table's own column span, so it's covered by `modInit.RelocateAtRiskButtons` (below) rather than a fixed "safe" anchor. `modLists.BindDefaultCells` binds all three cells' dropdowns using the same `BindStockRange`/`BindPrinterRange`/`ApplyTo` functions the table columns use — a single conceptual "row" rather than a whole column — and `modValidation.OnDefaultCellChanged` mirrors `OnPrinterChanged`/`OnStockChanged` for live edits to the two cells directly (same AT-05 compatibility check, same bidirectional rebind).

**Roll length entry unit (2026-09-25; relocated same day, see above).** Some print rooms prefer to type a roll job's length in centimetres rather than metres, but `Qty` is what `Area m2`/`Paper Cost` read directly (§5.1) and must always hold metres. `modInit.EnsureRollUnitSetting` self-provisions `LOC_RollUnit` at `B6` (Metres/Centimetres dropdown, defaults to Metres) with its label at `A6` — grouped with the three selectors above as the "regularly used, must stay visible" set, per direct user feedback, rather than sitting apart from them. Same self-healing treatment as `LOC_DefTech`/`LOC_Export`: re-provisioned on every `InitialiseWorkbook` **and** `RefreshLocations` run, so a duplicated or pre-existing sheet always ends up with a correctly-scoped name. `modValidation.OnQtyChanged` is the only thing that reads it: for a Centimetres-preferring location, whatever gets typed into a **Roll**-stock row's `Qty` is divided by 100 and the cell rewritten in place, the same live-rewrite idiom `OnWidthChanged` uses for its own cell — Sheet stock is untouched (`Qty` there counts sheets, not a length) and no formula, the `Unit` column, or any other code needs to know this setting exists. `modInit.EnsureJobColumnValidation`'s `Qty` input tooltip reads it too, purely for wording ("centimetres for roll stock" vs. "metres for roll stock").

**Superseded 2026-10-01: `LOC_RollUnit` is now a display unit, not an entry-time conversion.** `Qty` holds roll lengths in the sheet's own unit (centimetres on a Centimetres sheet) and nothing is rewritten on entry, so the 2026-09-27 order-dependency problem (`ConvertQtyIfCentimetres`, now removed) no longer exists. The Unit column reads `cm`/`metres`/`sheets`; `Unit`, `Area m2` and `Paper Cost` are restated by `modInit.EnsureRollUnitFormulas` (a `cm` row's `Qty` is divided by 100 inside the formulas, so area and money are identical either way). `modRegistry.MetresBlock` converts each location's `cm` rows back to metres/`"metres"` as `_Data` is built, so reports, Summary and the report snapshot export never see centimetres, and other locations follow their own settings. Changing `LOC_RollUnit` rescales the existing roll rows (`modValidation.ApplyRollUnitChange`; sheet-scoped constant name `LOC_RollUnitApplied` records which unit the stored values are in, so re-picking the same unit is a no-op). Location exports are written as displayed (the `Unit` column says which unit `Qty` is in); `modImport` converts a roll `Qty` if the file's unit differs from the receiving room's, and no longer overwrites the calculated `Unit` cell. A sheet from before this change that is set to Centimetres holds metres in `Qty`; `EnsureRollUnitFormulas` scales its roll rows x100 once, detected by its old `Unit` formula not naming `LOC_RollUnit`.

Duplicating a location sheet (§4.4) copies whatever values are currently sitting in the defaults, same as it copies `LOC_Name`/`LOC_DefDisPaper` — the documented add-location procedure now clears them alongside **Clear All**.

**Side panel (2026-09-25) — Location code, the two disregard-cost defaults and the permitted-printers list, columns AL/AM (38/39).** *(The two disregard-cost defaults moved back to `A6:B7` on 2026-09-29 — see above.)* Hand-edited directly into `PrintCosts.xlsx` (both shipped location sheets) rather than self-provisioned — these are the same shipped named ranges they always were, just repointed to their new cells. None of the four are used per-job, unlike the selectors above, so relocating them cost nothing in day-to-day use. The explanatory prose that used to sit at `D3` (about the **Select printers…** button) moved to `AL6` alongside the relocated printers list; `D1`/`D2` (heading, and the "defaults seed each new print job" note) stayed put, since they still describe what's now directly below them.

**Occasional-use toolbar buttons — Remove Row, Select printers…, Check this sheet, Clear All, Export…, Import…, all six sharing column `SIDE_PANEL_COL` (36) (2026-09-25; the reduced-view toggle and the cost-columns toggle that joined it on 2026-09-27 have since moved to the view row on row 2, below — see §16.3's "button lag" writeup for the cost-columns toggle's history).** Same reasoning as the cell-content move above, applied to buttons: past the table's own columns entirely, so `ApplyColumnVisibility` can never reach them regardless of what `SET_LOC_REDUCED_COLUMNS` names. Only Add Print Job and Now stay on the toolbar proper (row 11, next to the table — moved from row 10 by the toolbar-gap fix below) as the two used on every single job entry. `SIDE_PANEL_COL` was also widened (`ColWidthForPx(210)`, user-reported): a `Buttons.Add` shape's own width is independent of its anchor column's width, so at the column's previous narrow width every 140pt button sprawled two-three columns rightward, straight over the settings block at `AL`/`AM` — blocking its dropdown arrows from being clicked as well as visibly overlapping it. Widening the column keeps every button's footprint contained within its one column.

**All eight packed into one column via pixel-based stacking, not the row grid (2026-09-25, direct user report — see the row-scope bug above for why a second column existed first, and why it didn't last).** Every-other-row spacing only had six safe slots above the table header for seven buttons (eight at the peak, six now), and row height is a whole-row property shared with the main A/B block — the side panel couldn't use taller rows without inflating that block too. `modInit.DrawOneAtTop` (new) takes an explicit pixel `Top` instead of a row number, so the buttons are spaced evenly across whatever vertical room is actually available above the table's *current* header row (`ws.Cells(lo.Range.Row, 1).Top`, read at runtime, not a hardcoded figure) with an 8pt safety margin before it — the first attempt packed the last button flush against the header with none at all. Not tied to any particular row height, so it keeps working if the header moves again, unlike the row-based approach it replaces.

**Update 2026-10-01 (v0.10.12, reshaped in v0.10.14): the single toggle became three column views — All / Reduced / Minimal** — first as three buttons, from 0.10.14 as one drop-down (a forms control, `pcb_ddViewMode`, OnAction `modMain.ddViewMode` → `modInit.ViewDropDownChanged`). The **view row** is row 2 from D: label `Show columns:` in D2, the drop-down where the All button used to sit, then the Hide/Show cost detail button (moved out of the side panel in 0.10.14, so the view row now shows both independent column settings). `modInit.RepositionViewButtons` lays all three out on visible columns, re-reads the drop-down's selection from the setting and re-captions the cost button every time. `SET_LOC_REDUCED_VIEW` holds the mode, `SET_LOC_REDUCED_COLUMNS` the Reduced list (Status, Job ID, Area m2, S_SchemaVer) and the new `SET_LOC_MINIMAL_COLUMNS` the *extra* columns Minimal hides on top of it (Printer, both Disregard columns, Print Width mm, Sheet size). The Reduce-clutter side-panel button is gone; the paragraph below describes the original design.

**Reduced-clutter view (snag 1e, 2026-09-22).** A **workbook-wide** toggle — not per-sheet, so every location sheet always agrees with every other rather than risking drift — that hides Status, Job ID, Printer, Area m2, Disregard Paper, Disregard Consumable and `S_SchemaVer` on every location sheet at once. Driven by two settings, `SET_LOC_REDUCED_VIEW` (Yes/No) and `SET_LOC_REDUCED_COLUMNS` (the semicolon-delimited header list) — the list lives in a setting rather than a VBA constant specifically so it can be edited without a rebuild if the shortlist needs to change, matching the snag list's own "keep it flexible" instruction. `modInit.ApplyColumnVisibility(lo, Headers, Hide)` is the general primitive underneath (hides or shows *exactly* the named columns, by header, and touches nothing else) — it does **not** do a blanket "show everything, then hide the list" reset, because some columns (`H_Issues`) are hidden permanently, shipped that way in the `.xlsx` and never toggled by any code, and a blanket reset would have unhidden it the first time this ran. `modInit.ApplyReducedView` applies the current setting to one sheet's table and is called from both `InitialiseWorkbook` and `RefreshLocations`, same self-healing/duplicated-sheet reasoning as `EnsureJobDefaults` — and now also calls `modInit.RelocateAtRiskButtons` (2026-09-25) on every run, re-anchoring Add Print Job/Now/Clear defaults off whichever column is currently hidden. `modInit.ApplyColumnVisibility` is `Public` and reused as-is by the Reports page's own fixed minimum-columns view (§8.3, snag 2d) — same primitive, no toggle there, just applied once at build time.

**Three bugs found in testing, all the same "Excel's column-hide is whole-column" constraint recurring in different forms — the first fixed 2026-09-22, the other two 2026-09-25:**

1. **Hiding a column hides the WHOLE column, every row — not just the table's.** Status and Job ID were originally the job table's first two columns (sheet A/B), which are the *same* columns the location header block above the table occupied at the time. Toggling reduced view on was blanking the header along with the table cells. Fixed by `modInit.ReorderJobColumns`: Status and Job ID move to just after Paid, ahead of Notes/`H_Issues`/the snapshot block, so columns A/B are always "Date/Time"/"Student Name" — never anything the reduced view can hide. Column order isn't load-bearing anywhere (§5.1), with one matching fix needed: `modRegistry`'s consolidated-range span was bounded by column *name* from `"Job ID"` to `"Notes"` (`FIRST_JOB_COL`/`LAST_JOB_COL`, §8.1) — with Job ID no longer the leftmost real column, that bound moved to `"Date/Time"` instead, so the span still covers every real column (Status and Job ID included, since they now sit *inside* it rather than starting it).
2. **A toolbar button anchored to a hidden column vanishes with it.** Caught once 2026-09-22 (the toggle button at column 13, coinciding with Disregard Consumable — patched by moving it to column 7) and again 2026-09-25 once reported directly (Remove Row on Printer, Clear All on Disregard Consumable). Fixed generally this time rather than patched column-by-column — see "Occasional-use toolbar buttons" and `RelocateAtRiskButtons` above.
3. **Cell content sharing a column with a hidden table column vanishes too, same as a button — but can't self-relocate the way a button can.** `LOC_RollUnit` and "Default: paper" on column E (Printer) — see the layout rework above. The fix here is structural (confine anything that must stay visible to columns A/B, which never appear in any hide list) rather than dynamic, since a cell's value can't float to a different column at runtime.

**The button case is now general; the two 2026-09-25 structural fixes (columns A/B, the side panel) close out this section's earlier "not a general solution — revisit later" note** for both the cell-content and button classes of this bug. `RelocateAtRiskButtons` is still a per-button anchor list rather than a blanket guarantee, so a future button added inside the table's column span needs to be added to it explicitly.

**Also flagged, not yet addressed:** the reduced-view toggle's only feedback that it's engaged is the button's own caption text ("Reduce clutter" ↔ "Show all columns"). A clearer visual indicator of the current state (sheet-level, not just the button) would be nice, but the broader "Package 10" visual-polish rework this was originally filed against has since been descoped (§16.3) as more than this workbook needs — native Excel styling is considered sufficient for now, and this specific idea is not currently planned.

**Why sheet-scoped names matter:** Excel copies them with the worksheet. Workbook-scoped names would collide on copy and silently become `LOC_Name1`, `LOC_Name2`.

**The marker is the exception, and is a plain cell.** It copies with a duplicated sheet exactly as a scoped name does, Excel cannot quietly rename it, and it is one moving part instead of two.

**Freeze panes (unchanged by the 2026-09-21 snag list; that work targeted Settings/Printers/Papers/Technicians and Reports instead — §9.4, §11).** Location sheets keep whatever freeze panes the `.xlsx` ships.

**Column grouping is entirely `modInit.NormalizeJobColumnOutlines`' concern, and as of 2026-09-27 groups nothing at all.** From 2026-09-22 it grouped exactly one range, `Paper Cost`→`Disregarded` (the snag list's original "add column groups to related table columns including costs", item 14), behind Excel's native `+`/`-` outline control. That native group is what caused the button-lag bug recorded in §16.3 — a manual outline toggle is not a macro call, so nothing could reposition the side-panel buttons anchored beside it until the next click anywhere on the sheet. Replaced 2026-09-27 by a workbook-wide macro toggle, same shape as the reduced-clutter view: `modInit.ApplyCostColumnsVisibility`/`ToggleCostColumns` hide/show that same four-column span via a "Hide/Show cost detail" button (in the side panel until 0.10.14, then on the view row, row 2), repositioning every affected button in the same click. `NormalizeJobColumnOutlines` now only ungroups — including its own former group, migrating a workbook built under the previous scheme — so `modProtect.ProtectSheet`'s `EnableOutlining = True` is vestigial for this span (kept regardless: harmless, and still correct in principle if grouping is ever reintroduced elsewhere).

**Everything from Notes onward is deliberately left at outline level 1 — not grouped — even though `H_Issues` and the whole snapshot block ship hidden in `PrintCosts.xlsx` (§3.4, §5).** An earlier version of this code *also* called `GroupColumnRange` on `S_PrinterID`→`S_SchemaVer`, on top of the grouping the `.xlsx` already ships for that same span — nesting a second outline level over it rather than reusing the first. The visible symptom: two extra, unwanted groups next to the cost group (Notes+`H_Issues` stuck together at one level, most of the snapshot block one level deeper again), and — the real bug hiding inside that mess — **Notes itself ended up hidden**, a genuine input column silently swept into the snapshot block's hidden state as a side effect of `EnsurePaidColumn` inserting Paid directly next to it. `modInit.FlattenOutline` fixes both: resets outline level to 1 across the whole `Notes`→`S_SchemaVer` span (rather than re-grouping it) and explicitly un-hides Notes. `H_Issues`/the snapshot columns stay invisible exactly as before, via their own `.Hidden` state — which never needed an outline group to hold it in the first place. Verified by `test-layout.ps1`.

### 4.2 Detection and registration (D5)

A sheet is a location if and only if `AZ1` reads `PRINTLOC/v1`. **Refresh Locations** walks every worksheet, tests the marker, and rebuilds `_Registry.tblLocations`. In order:

1. **Parks every job table on a temporary name.** Excel auto-renames a duplicated `ListObject` — in testing, `tblJobs_ANNEX` became `tblJobs_ANNEX14`. Renaming straight to the final name fails whenever that name is still held by a table the pass has not reached, which is exactly the situation after a duplication.
2. **Assigns a location code** where `LOC_Code` is blank or duplicated. Sheets are visited in workbook order and Excel inserts a duplicate immediately after its original, so the original is reached first and keeps its code — required, because its job IDs already carry it.
3. **Renames each job table** to `tblJobs_` + its code (AT-13).
4. **Re-points Form Control buttons** whose `OnAction` acquired a workbook-file prefix on copy.
5. **Rebuilds the dependent dropdowns** and **re-applies protection**.
6. **Rewrites the consolidated range formula** (§7.1) and **refreshes the Reports filter dropdowns** (§8, since `_Data` just changed).
7. **Recomputes each location's export status** and writes `LOC_Export` — after the registry is written, because the status is read out of it (§10.4).
8. **Stamps `SET_LASTREF`.**

Any location with unexported changes is named in the refresh message. `Last export`, `Export sig` and `Job ID HWM` are captured before the registry rows are wiped and restored per sheet name afterward, so none of them resets on a routine refresh.

**Button bindings need healing on open, too.** Excel re-qualifies every button's `OnAction` with the workbook's file name when it saves, so the shipped file reads `PrintJob.xlsm!btnAddPrintJob` — which breaks if the file is renamed. `Workbook_Open` calls `HealButtons`, which rewrites only on a mismatch, so merely opening an untouched workbook never marks it dirty and prompts a save. This is also what makes renaming the built `.xlsm` itself (§1's naming-inconsistency note) safe, whether done by `build.ps1` or by hand.

### 4.3 Deleted locations (D18, was O9)

**Excel provides no `BeforeDeleteSheet` event.** Deleting a location sheet cannot be intercepted, undone or audited. What changes is whether the records were somewhere else first.

- **A recovery path.** Each location exports to CSV holding every column, including the snapshot block (§10.4), and — new — an entire workbook's locations can be exported in one action (§10.4), and a CSV export can be **imported back** (§10.5).
- **A visible warning where the user is.** `LOC_Export` states how many records have changed since the last export, and *Check workbook* and *Refresh Locations* surface the same line per location.
- **Self-healing afterwards.** The consolidated formula is wrapped in `IFERROR`, so a deleted sheet degrades to an empty result rather than cascading `#REF!`. Registry entries are compared against the sheets present on each refresh, and anything vanished is named.

**This is a sign on the cliff, not a fence**, and the design says so rather than implying otherwise. The rejected alternative was a mirrored shadow copy of every location's rows — a permanent cost on every write, for an edge case a working export covers better and which has independent value.

**A guarded in-app remove command — Settings sheet, "Remove print room..." button (2026-09-27, direct user request), `modRegistry.RemovePrintRoom`/`DeletePrintRoom`.** Sits right next to "Add print room..." (§4.4), and answers the question §4.4's addendum and §16.3 had both left open: what a deliberate, in-app *remove* command should do. The manual route (right-click the tab → *Delete*) is untouched and still fully supported — this is an additional guarded route, not a replacement, built on top of the same self-healing this section already documents rather than replacing any of it.

Because Excel raises no delete event, the safety has to be entirely up front, before anything happens, not after:
1. **Pick the room** — a list of every registered room's code and name, typed back (same "type its code" shape as `modImport.PickTargetLocation`, the Import — choose a room picker).
2. **Refuse the last room outright**, not just warn about it. *Add print room* works by duplicating an existing room (`TemplateLocationSheet`), so removing the last one would break the button sitting right next to this one — a real structural dependency, not a soft preference, so it is a hard stop rather than an `Ask()`.
3. **Name what will be lost** — room name and code, record count, date span, and export status (`modExport.ExportStatusText`, the same line *Check workbook*/*Refresh Locations* already surface) — before any confirmation is asked.
4. **A stronger warning when there are unexported changes**, since those records would be lost for good rather than just removed from this sheet — an extra `Ask()` step naming that explicitly, on top of the confirmation below, not instead of it.
5. **Type the room's name back exactly**, rather than a plain Yes/No — deliberate, since "Yes" is one keystroke away from an unrecoverable loss and a typed name cannot be dismissed by reflex the way a default button can.

Only then does `DeletePrintRoom` run: log the removal to `tblAudit` (name, code, record count, date span, export status — captured *before* the sheet goes, since nothing can be read off a deleted sheet afterward), delete the sheet with `Application.DisplayAlerts` suppressed (so Excel's own native "data may exist" prompt doesn't stack a second, unlabelled confirmation on top of the one already given), then run the same **Refresh Locations** pass every other structural change in this module ends with — `MissingSheets` (§4.2) already compares the registry against what is actually on the workbook and names anything gone, so the "self-healing afterwards" story above covers this command's own deletion for free, the same way it covers a manual one.

Split the same way `AddPrintRoom`/`CreatePrintRoom` and `modImport.ApplyImport`/`ApplyImportConfirmed` are: a prompting wrapper (`RemovePrintRoom`, bound to the Settings button) and a parameter-driven core (`DeletePrintRoom`) a test script can call directly. Unlike those two, the split here is not optional for testing — `RemovePrintRoom`'s picker and typed-name confirmation both go through `Application.InputBox`, which (unlike `Ask`/`Say`) has no `SetQuiet` bypass and would hang an unattended COM run waiting for an answer nobody can give, so `test-removeroom.ps1` calls `DeletePrintRoom` directly, the same way `test-addroom.ps1` calls `CreatePrintRoom` rather than `AddPrintRoom`.

### 4.4 Documented procedure for adding a location

1. Right-click any location tab → *Move or Copy* → tick *Create a copy*.
2. Rename the new tab.
3. Click **Clear All** to discard the copied records, and **Clear defaults** to discard the copied batch-default selectors (§4.1).
4. Enter room name, department, defaults; click *Select printers…*.
5. Click **Refresh Locations**.

**Automated equivalent — Settings sheet, "Add print room..." button (2026-09-27, direct user request), `modRegistry.AddPrintRoom`/`CreatePrintRoom`.** Prompts for a room name (becomes the new sheet's tab, sanitised and de-duplicated against every existing sheet name, not just other locations), an optional department, and a short job-ID code, duplicates the first registered location sheet, then does the equivalent of steps 3-5 above in one pass — **and goes further than the manual procedure ever did**: as well as the job table and the three batch defaults (Clear All/Clear defaults' own scope), it also blanks the two disregard-cost defaults, the permitted-printers list, and the roll-unit preference (reset to the Metres default rather than left as whatever the template had) before Refresh Locations runs. The manual procedure above is unaffected and stays fully supported — this is an additional route, not a replacement. **Removing a room automatically is now built too** — see §4.3's addendum for `RemovePrintRoom`/`DeletePrintRoom`. Split into a prompting entry point and a parameter-driven core (`CreatePrintRoom`) a test script can call directly without a dialog to answer, same shape as `modImport.ApplyImport`/`ApplyImportConfirmed` and this module's own `RefreshLocations`; verified by `test-addroom.ps1`, which deliberately dirties the template's defaults first so the reset is actually exercised, not trivially true. Ends by opening the "Select printers…" picker for the new room, since Refresh Locations already gives its own registration confirmation and a duplicate dialog would be redundant.

**Job-ID code prompt (2026-09-27, direct user request, same day).** A third dialog captures the short code the room's job IDs are built from (§3.2's `<SITE>-<LOC>-00001`) — e.g. `PHOTO` for `UNI-PHOTO-00001` — rather than only ever deriving it silently from the room name the way a manually-copied-and-renamed sheet still does (`AssignCode`'s existing fallback, unchanged). Pre-filled with a suggestion derived from the room name (so accepting the default is usually enough) but always editable, since a good job-ID code and a good room name often want to be different shapes — "Photography" suggests `PHOTOG`, but the user can still type `PHOTO`. Whatever is typed is normalised the same way either path already was: uppercased, non-alphanumeric characters stripped, capped at six characters (`CleanCode`, generalised to take a `MaxLen` rather than hard-coding `AssignCode`'s own eight-character ceiling — a deliberately-chosen code reads better shorter than the derive-from-name fallback's limit). Left blank (a test script calling `CreatePrintRoom` without a `Code` argument), the new sheet's `LOC_Code` is blanked exactly as before and `AssignCode` falls back to deriving one from the sheet name, so existing callers are unaffected. A collision with an already-registered code (unlikely for a deliberately chosen short code, but possible) is still caught and de-duplicated by `AssignCode`'s own `Uniquify` the moment Refresh Locations runs, the same as any other code. Verified by `test-addroom.ps1`'s second room, created with a deliberately messy code (`"photo lab 2026!!"`) to prove the normalisation rather than just the happy path — comes out `PHOTOL`.

---

## 5. Transaction table

**Column order changed 2026-09-22** (§4.1's reduced-clutter-view fix, `modInit.ReorderJobColumns`) — Status and Job ID used to be columns 1-2; they now sit just after Paid. The `#` column below is the *current* physical position; none of it is load-bearing (see the note below the table).

**`PrintCosts.xlsx` itself reordered to match, 2026-09-26.** The shipped `.xlsx` had never been updated to reflect the move above — it still had Status/Job ID as columns 1-2, so opening it directly to edit a design element (a column's width, its input styling, a validation rule) meant mentally translating through `ReorderJobColumns` to know where that column would actually end up in the built `.xlsm`. Fixed by physically cutting Status and Job ID to just after Chargeable Cost in both `tblJobs_MAIN` and `tblJobs_ANNEX` (via Excel COM automation, not hand-edited XML — see the note below on why). `Paid` and the `Quantity`→`Qty` rename stay VBA-only exactly as before (§3.3/§3.5's "build it in VBA, not by hand" reasoning) — this touched order only, so `EnsurePaidColumn` still inserts Paid right after Chargeable Cost (landing exactly between it and Status, matching the table above) and `EnsureQtyColumnName`/`ReorderJobColumns` both remain no-ops against a freshly-setup copy of this file.

**Cut/insert reordering corrupts two column-indexed properties that don't travel with the moved cells — found and fixed doing the above, worth knowing before ever reordering table columns by hand again.** `Range.Cut` + `Range.Insert Shift:=xlToRight`, run repeatedly to walk ~19 columns into place, moves cell content and per-cell formatting correctly (confirmed: calculated-column formulas, the `Status`/date dxf styling, and every data value all followed their column intact), but two things are keyed by column *index* rather than by the cell content that moved:
- **Data validation ranges.** Each single-column rule (Technician, Printer, Paper Stock, Qty, Print Width, the two Yes/No lists, Date/Time) got smeared across the *entire* current data-row span on every column the repeated inserts passed through — e.g. Technician's rule ended up on `D13:Q17` instead of `D13:D2000`. This is the exact `EnsureJobColumnValidation` bug already documented in `modInit.bas` for the `.xlsm` case; it just has no VBA to self-heal it in the `.xlsx`, so it was fixed by hand (delete, then re-add each rule as a clean single-column range at its new column letter).
- **Column width.** `<cols>` width entries are addressed by column index and don't move with a cut at all — after the reorder, column A (now Date/Time) still carried Status's old 30-character width, and so on down the line. Fixed by remapping each column's width to follow its content to its new position.

Anyone reordering job-table columns directly in `PrintCosts.xlsx` again (as opposed to via VBA, where `EnsureJobColumnValidation`/`ApplyJobColumnWidths` already re-assert both of these on every setup run) needs to redo both fixes by hand afterward.

| # | Column | Type | Notes |
|---|---|---|---|
| 1 | Date/Time | input | Required. Toolbar *Now*, or double-click |
| 2 | Student Name | input | At least one of name/number required |
| 3 | Student No | input | " |
| 4 | Technician | input | Dropdown, active only |
| 5 | Printer | input | Dropdown, active ∩ permitted here |
| 6 | Paper Stock | input | Dropdown, active ∩ compatible with printer |
| 7 | Unit | calc | `sheets` or `metres` |
| 8 | Qty (renamed from "Quantity" 0.9.14 — `modInit.EnsureQtyColumnName`) | input | > 0; whole number for sheet stock (I3) |
| 9 | Print Width mm | input | Roll only; blank = full stock width; ≤ stock width. Required, not optional, when Paper Stock is `Supplied (Roll)` — see below |
| 10 | Sheet size | input | New, printer/paper compatibility rework (§16.5). Meaningful only when Paper Stock is `Supplied (Sheet)`, where it's required — a Standard Sizes pick, the nearest size to the sheet the student brought. Ignored (and cleared) on every other row |
| 11 | Disregard Paper | input | Yes/No, seeded from location default |
| 12 | Disregard Consumable | input | Yes/No, seeded from location default |
| 13 | Area m2 | calc | Printed area |
| 14–18 | Paper Cost, Consumable Cost, Gross Cost, Disregarded, Chargeable Cost | calc | §5.1 |
| 19 | Paid | input | Yes/No, blank on rows that predate it (2026-09-22 snag list item 1c) — see below |
| 20 | Status | calc | `OK` or a warning; conditionally formatted |
| 21 | Job ID | VBA | `<SITE>-<LOC>-00001`, never re-used (§3.2) |
| 22 | Notes | input | Optional |
| 23 | H_Issues | calc | Hidden working column behind Status |

**Snapshot block (24–34)** — locked, grey, collapsed group headed *Historical record — do not edit*: `S_PrinterID`, `S_StockID`, `S_TechID`, `S_Measure`, `S_UnitCost`, `S_StockWidth_mm`, `S_SheetHeight_mm`, `S_ConsRate`, `S_StampedAt`, `S_StampedBy`, `S_SchemaVer`.

**Sheet size (printer/paper compatibility rework, §16.5) — the second genuine job-row column since inception, `modUtils.SCHEMA_VER` 1.1 → 1.2.** Added via `modInit.EnsureSheetSizeJobColumn`, same idempotent `ListColumns.Add` shape as `EnsurePaidColumn`. Plays the same role for `Supplied (Sheet)` that `Print Width mm` already played for roll stock: there's no catalogue size to fall back on for student-supplied paper, so the real size is entered per job and stamped into `S_StockWidth_mm`/`S_SheetHeight_mm` by `modSnapshot.StampRow` (resolved via a `tblStandardSizes` lookup, `modCatalog.StdSizeDims`) exactly as if it had come from a catalogue row — **no change to any cost formula was needed** (§5.1). Both `Print Width mm` (now required, not optional, for `Supplied (Roll)`) and `Sheet size` are validated against the *printer's* capacity, not a stock's nominal size (`modValidation.OnWidthChanged`/`OnSheetSizeChanged`/`RevalidateSuppliedSize`) — and a blank value on a row that needs one is caught by `H_Issues` (`modInit.EnsureJobIssuesFormula`, asserted in VBA rather than a static `.xlsx` formula edit, same self-healing reasoning as `EnsureJobColumnValidation`) the same way every other required field already is, so Check sheet/Check workbook catch a forgotten one.

Column *order* is not load-bearing anywhere: formulas use structured references, VBA resolves columns by header name, reports resolve `_Data` by header name (§6.2), imports write by header name (§10.5), and exports carry a header row read by name (§10.4). Reordering is free and does not bump the schema version — proven in practice by the 2026-09-22 move above, which touched nothing but `modInit` (the reorder itself) and `modRegistry`'s `FIRST_JOB_COL` constant (the one place a column's *name*, not position, was baked in as a span boundary — see §8.1).

**Paid (snag 1c, 2026-09-22) — the first real schema bump since inception.** Whether a chargeable cost has been settled — Yes/No, no location default (unlike Disregard Paper/Consumable, every new job just starts unpaid). Added via `modInit.EnsurePaidColumn` (`ListColumns.Add` at a fixed position right after Chargeable Cost, checked-first so it's idempotent), not shipped in the `.xlsx` — the same "build it in VBA, not by hand" reasoning as `_Registry`/`_Audit`/`_Data` (§3.4). **Existing rows are left blank, deliberately not force-set to "No"**: blank means "not recorded either way" for a job that predates the column, and every consumer (Summary/Reports totals, export, import) already treats blank the same as "No" rather than requiring a value, so no backfill or migration step was needed.

Because this adds a genuine job-row column, `modUtils.SCHEMA_VER` moves from `1.0` to `1.1` — the first change since inception (§3.5's "adding a column bumps the schema version" rule, never previously exercised). This surfaced a real gap: `SET_SCHEMA` had shipped as a **static value directly in `PrintCosts.xlsx`** (§3.3) with no code path keeping it in sync with the constant, unlike `SET_APP_VER` which `modVersion.EnsureVersionSettings` stamps from `APP_VERSION` on every setup run — exactly the kind of drift §16.1 caught `APP_VERSION` itself in for thirteen commits. Fixed the same way: `modVersion.EnsureSchemaSetting` (new, called from `InitialiseWorkbook` alongside `EnsureVersionSettings`) now stamps `SET_SCHEMA` from `SCHEMA_VER` every run, so `SET_SCHEMA` can never again silently disagree with the code that actually enforces the shape.

Both **Summary** and **Reports** split their Chargeable total by paid status (Summary `J6`/`L6`, Reports "Matching" `J13`/`L13`) — Paid is `SUM` of Chargeable Cost where `Paid="Yes"`; Unpaid is the filtered Chargeable total *minus* that Paid figure, rather than a separate `<>"Yes"` criteria, so the two always reconcile exactly to Chargeable by construction and a blank Paid falls into Unpaid either way.

**Quantity renamed to Qty (0.9.14) — a pure rename, not a schema bump.** Direct user feedback: "Quantity" cost the narrow job-row columns more width than it needed to. `modInit.EnsureQtyColumnName` renames the live `ListColumn` (checked-first, idempotent, same shape as `EnsurePaidColumn`) rather than hand-editing `PrintCosts.xlsx` — Excel updates every structured reference to the column (the Area m2/Paper Cost formulas above, `_Data`'s consolidated copy of the header) as part of the rename itself, the same as a person renaming it through the ribbon would get. Every VBA call site that looks the column up by name (`modReports.C`/`SumBy`, `modExport.ExportColumns`/`Agg`, `modImport.WriteImportedRow`) was updated to ask for `"Qty"` instead. Not schema-bearing (§3.5 — only a genuine add/remove/rename of what a job row's columns *mean* counts, and this one doesn't change what's stored, only its label), so `SCHEMA_VER` stays at `1.1`. A file exported before the rename has a column literally called "Quantity"; re-importing it degrades the same way a pre-Paid export already does — the missing key reads back blank rather than raising (`modImport`'s own comment).

**Column widths (0.9.14).** `modInit.ApplyJobColumnWidths`, called for every location sheet alongside the other per-sheet fixups, narrows Unit/Qty (≤50px), Area m2 (≤60px) and Paid (≤40px), and fixes Job ID and Date/Time at exactly 100px — all via `modUtils.ColWidthForPx`, which converts a pixel target into `Range.ColumnWidth`'s own "characters of the Normal-style font" unit using Calibri 11's well-known 7px-per-character/5px-padding metric (the same constant openpyxl/xlsxwriter use), since there is no `Range.WidthInPixels` to set directly. Approximate by construction — always rounds down, so a cap can never render one pixel over, and an exact target (Job ID/Date-Time) accepts the same sub-pixel tolerance. **0.10.8:** the same routine also sets explicit widths on the remaining sized columns, which previously kept the oversized `PrintCosts.xlsx` values — Technician 120px, Printer 160px, Paper Stock 170px, Sheet size 90px, Paper Cost 90px, Consumable Cost 100px, Gross Cost 90px, Disregarded 100px, Chargeable Cost 100px. Headers wrap, so a column needs only its longest word plus the filter button. Notes is deliberately left as shipped. **0.10.9:** Status is 120px; the full message is carried in a hover note on each problem row's Status cell (none on OK/blank rows), synced by `modInit.RefreshStatusNotes` — diff-only, called from `ApplyStatusFormat`, `Workbook_SheetChange` and `Workbook_SheetSelectionChange`. A note is used because Excel has no dynamic cell tooltip and a note is the only hover text; being static, it is refreshed on the next edit or click rather than instantly on a catalogue-driven change.

### 5.1 Formulas

```
Unit             =IF([@S_Measure]="Sheet","sheets","metres")

Area m2          =IF([@Qty]="","",
                   LET(w, IF([@S_Measure]="Sheet",
                             [@S_StockWidth_mm],
                             IF([@[Print Width mm]]="", [@S_StockWidth_mm], [@[Print Width mm]])),
                       h, IF([@S_Measure]="Sheet", [@S_SheetHeight_mm]/1000, [@Qty]),
                       n, IF([@S_Measure]="Sheet", [@Qty], 1),
                       (w/1000) * h * n))

Paper Cost       =IF([@Qty]="","",ROUND([@Qty]*[@S_UnitCost],SET_ROUND_DP))
Consumable Cost  =IF([@[Area m2]]="","",ROUND([@[Area m2]]*[@S_ConsRate],SET_ROUND_DP))
Gross Cost       =IF([@[Paper Cost]]="","",[@[Paper Cost]]+[@[Consumable Cost]])
Chargeable Cost  =IF([@[Paper Cost]]="","",
                    IF([@[Disregard Paper]]="Yes",0,[@[Paper Cost]])
                  + IF([@[Disregard Consumable]]="Yes",0,[@[Consumable Cost]]))
Disregarded      =IF([@[Gross Cost]]="","",[@[Gross Cost]]-[@[Chargeable Cost]])
```

The formulas read **only snapshot columns and the row's own inputs**. No formula on a job row reaches into `tblPapers` or `tblPrinters` — the property that satisfies AT-09 and makes historical costs immune to config changes, and is also exactly why **imported rows cost correctly without any catalog reconciliation** (§10.5): an imported row already carries the rates it needs.

`Print Width` appears in the area formula and nowhere in paper cost: AT-02 and AT-03 made structural rather than procedural.

`Qty` is always metres for roll stock by the time any of the above runs, regardless of what a person typed — a Centimetres-preferring location's raw entry is converted at the cell itself before these formulas ever see it (`LOC_RollUnit`, §4.1).

---

## 6. Historical-data strategy (D2)

### 6.1 Stamping

When **Add Print Job** creates a row, and whenever printer or paper stock changes, `modSnapshot.StampRow` resolves the current configuration and writes the twelve snapshot values. Costs are computed by the row's own formulas from those values — *live formulas over frozen rates*, so the arithmetic stays auditable and a mis-typed quantity recalculates immediately.

### 6.2 Re-stamping

A **Re-stamp prices** command exists for the genuine correction case. Deliberately not on the location toolbar: it lives on Settings, confirms while naming the affected rows, and writes an `_Audit` entry.

### 6.3 What survives configuration change

| Change | Effect on existing rows |
|---|---|
| Paper cost edited | None — from `S_UnitCost` (AT-09) |
| Consumable £/m² edited | None — from `S_ConsRate` |
| Printer renamed | None — row holds `S_PrinterID` |
| Printer capacity (Max roll width mm / Max sheet size) changed | None — compatibility validated at entry (AT-05); §16.5 |
| Stock dimensions changed | None — area from the stamped dimensions |
| **Record set Inactive** | **None — remains in reports; excluded only from new dropdowns (AT-16, verified phase 7)** |
| Location defaults changed | None — flags copied into the row at creation (AT-07, AT-08) |
| **Global rounding changed** | **Recalculates existing rows — accepted (O1)**, and labelled as such on Settings |
| **Config row removed via the row buttons (§3.3)** | None — jobs already recorded keep their snapshot prices; the confirmation text says so explicitly |

---

## 7. Data-validation strategy

### 7.1 Layer 1 — Data Validation

Static lists, numeric constraints, date type constraints. Enforced before VBA sees anything.

### 7.2 Layer 2 — Dependent lists on staged ranges

Version 1.1 specified dropdowns on spill references. Two hard limits ruled that out:

- **A Data Validation rule is uniform down a table column**, but each row's candidate stocks depend on *that row's* printer. One spill cannot serve every row, and the workaround — a `SelectionChange` handler writing the active row's printer into a context cell — makes the dropdown depend on the selection. Replacing a cell's validation closes an open dropdown, so the arrow visibly flashed and vanished.
- **A validation list supplied as a literal string is capped at 255 characters**, which a real stock list exceeds.

`modLists` gives **every distinct list its own staging column** on `_Work`, claimed by a tag in row 1 — `PRN|<all>|<sheet>`, `PRN|<sheet>|<stock>`, `TEC`, `STK|<all>|<sheet>`, `STK|<model>` — and reused for the Reports page's own dropdowns (§8). Consequences:

- Lists rebuild when something **changes**, never on selection, so nothing flashes.
- Rows are grouped by the *other* field's value before binding (Printer rows grouped by Paper Stock, and vice versa), so each distinct list is written once, not once per row.
- A per-list column was necessary: one shared column per list *type* meant the last sheet bound overwrote the others, and one print room showed another's printers.

**Printer and Paper Stock filter each other, in both directions (2026-09-22 snag list item 1a).** Earlier, Paper Stock was locked until a printer was chosen, so a printer always had to be picked first. Both cells now start unlocked and fully populated — Printer with every active printer permitted at the location, Paper Stock with every active stock compatible with *some* printer permitted there. `modCatalog.PrintersForStock` (the reverse of the existing `StocksFor`) computes what the OTHER field's current value makes compatible; if that narrowed set collapses to exactly one option, it's auto-filled rather than making the user pick the only choice (`modLists.AutoFillIfSingle`) — but never overwriting a value already there, and never mistaking the "nothing available" placeholder text for a real singleton. `modValidation.OnPrinterChanged`/`OnStockChanged` are now symmetric: either field can strand the other's value (AT-05 applies both ways), and each change rebinds both cells' lists so a field that gets cleared widens the other back out immediately.

**The dropdown itself no longer narrows to compatible-only options (0.9.11, direct user feedback that 0.9.0's narrowing was "too restrictive").** Every active/permitted item is always listed on both fields; whichever ones the OTHER field's current value rules out are suffixed `" (unavailable)"` (`modLists.UNAVAILABLE_SUFFIX`) rather than removed. Picking a marked item is allowed — `modValidation.Clean()` strips the suffix off the cell that was just edited (both table rows and the batch-default cells) before anything else runs, so the existing AT-05 handling above sees a clean catalogue name and clears/explains exactly as it already did; the marker never reaches a formula, a snapshot, or `Compatible()`. Auto-fill is unaffected: `AutoFillIfSingle` is still driven by the narrowed/compatible collection computed above, which is never itself shown to the dropdown — only used to decide whether to fill.

**This is a known compromise, not a full fix.** What was actually asked for was per-item visual styling (italics, a lighter colour, shading) on the unavailable entries. Native Excel Data Validation cannot do that: the in-cell dropdown list is rendered by Excel itself with no per-item formatting hook VBA can reach. A worksheet-based picker — in the style of `modPicker.bas`'s existing multi-select dialog (§9.3 below), which is a real worksheet and so CAN render muted/greyed/italic cells per item — would get closer to what was asked for, but was judged overcomplicated for what this actually needs (2026-09-23) and is **not planned** — see §16.3's note on descoping the broader visual-polish rework this would have been part of. The text suffix stands as the answer for now; the workbook keeps native Excel styling throughout rather than building a bespoke picker UI for this one field pair.

A `Range.Value` read on a target spanning more than one cell — the `Union` a row-group can produce when the matching rows aren't contiguous — returns an **array**, not a scalar; `AutoFillIfSingle` walks `target.Cells` individually rather than reading `target.Value` directly, which is exactly the bug this surfaced as during development (`CStr()` on that array raises a Type Mismatch that only a non-contiguous group triggers, so it's easy to miss in a small test dataset — worth remembering for any future code that reads `.Value` off a grouped/Union range).

Verified by `test-dropdowns.ps1`: both directions narrow and auto-fill correctly, a still-ambiguous narrowed list is left blank with the right options staged, and an incompatible combination is still caught and cleared regardless of which field was typed second.

### 7.3 Layer 3 — Row and workbook validation

`Worksheet_Change` runs `OnCellChanged` for cross-field rules validation lists cannot express: print width > stock width (AT-04); print width against sheet stock; non-integer sheet quantity (I3); neither student name nor number; name/number conflict against history (D1, AT-11); printer/stock combination made incompatible by a later edit (AT-05).

Because a row can be invalidated by an edit elsewhere, *Status* persists the verdict, and **Check workbook** sweeps every location and reports all outstanding problems in one list.

**Check workbook also reports export state, counted separately.** Unexported records are not a defect in the data — they are a risk *to* it, because sheet deletion cannot be intercepted (§4.3). Folding them into the problem count would make a perfectly valid workbook report problems, and a count that cries wolf gets ignored.

Messages state what is wrong, why, and what to do — enforced by `modUtils.Say` taking the three parts as separate arguments so a caller cannot quietly omit the third.

**Imported rows are written with events disabled** (§10.5), so they bypass `Worksheet_Change` entirely — Status/H_Issues do not populate until an explicit **Check workbook** / **Check this sheet** sweep runs on them.

---

## 8. Reports (D13)

Two report sheets: **Summary** (at-a-glance totals) and **Reports** (filterable record-level detail — renamed from "Cost Calculations" as part of the 2026-09-21 rework, "Problem 2" in the historical plan). Both are **built once** by `modReports` — layout and formulas, computed nothing at run time — and are live over the consolidated range on `_Data`, so neither can go stale and neither needs refreshing by hand except when `RefreshLocations` rewrites `_Data` itself.

### 8.1 The consolidated range

`Refresh Locations` writes one formula to `_Data!A10`:

```
=LET(raw,
  VSTACK(
    HSTACK(IF(SEQUENCE(ROWS(tblJobs_MAIN[Date/Time]))>0,"MAIN"),  tblJobs_MAIN[[Date/Time]:[Notes]]),
    HSTACK(IF(SEQUENCE(ROWS(tblJobs_ANNEX[Date/Time]))>0,"ANNEX"), tblJobs_ANNEX[[Date/Time]:[Notes]])),
  IFERROR(FILTER(raw, INDEX(raw,,2)<>""), ""))
```

The span's boundary column names (`FIRST_JOB_COL`/`LAST_JOB_COL`, `modRegistry`) changed from `"Job ID"`/`"Notes"` to `"Date/Time"`/`"Notes"` on 2026-09-22, when `modInit.ReorderJobColumns` moved Status and Job ID away from the front of the table (§4.1, §5) — Job ID stopped being the leftmost real column, so the span's start had to move with it. Status and Job ID are still fully included in the consolidated range either way, since they now sit *inside* the `Date/Time`→`Notes` span rather than starting it.

- **`IF(SEQUENCE(ROWS(…))>0,"CODE")` rather than a bare `"CODE"`**: `HSTACK` does not broadcast a scalar against a column.
- **`FILTER(raw, INDEX(raw,,2)<>"")` drops blank table rows.**
- **`IFERROR` on the outside** is the §4.3 safety net.

Shape: `Location`, then `Date/Time` through `Notes` — 22 columns (21 job-row columns, up from 20 once Paid was added, §5). Headers are written to row 9 from the first location's actual header row, so they cannot drift.

VBA's entire role in reporting is rewriting that one formula (plus refreshing the Reports filter dropdowns, §8.4). Everything downstream is a live worksheet formula.

### 8.2 Summary sheet — built, reworked 2026-09-21

One row per **Location × Printer × Paper stock** — a three-column key, extended from the original two-column Location × Paper-stock key (historical "Problem 2": two printers sharing a stock get separate rows, because cost-per-print differs by printer's consumable rate even for the same paper). Key pairs derived with `SORT(UNIQUE(HSTACK(...)))` and aggregated with `COUNTIFS`/`SUMIFS` taking the key columns as **array criteria**, which makes the results spill alongside the keys instead of needing one formula per row.

| Location | Printer | Paper stock | Type | Unit | Jobs | Qty | Area m² | Paper cost | Consumable cost | Gross | Disregarded | Chargeable |

Type and Unit are resolved from `tblPapers` by stock description, wrapped in `IFNA` so a stock renamed or removed since shows `(not in Papers)` rather than an error. These are labels only — no cost figure is ever looked up live (§5.1).

Consumption columns count **every** record regardless of disregard flags; only money columns split (D8).

**Totals sit above the table, not beneath it.** The detail spills to an unpredictable height, so anything below it is overwritten the moment a job is added.

**A cell-colour legend and conditional formatting (phase 8 carry-over, resolved 0.8.1, see §16.2) sit at `O9` down**, clear of the report table and the buttons drawn at column O rows 1/3/5 (row 7 freed up 2026-09-22 when "Go to Settings" was removed, §16.4 Package 7): `modReports.DrawLegend` explains the four everyday cell colours, and `modReports.FormatSummaryErrors` gives the Type column red bold text for a paper stock no longer in `tblPapers`.

**"Hide settings sheets" / "Show settings sheets" toggle (snag item 13, resolved).** A button on Summary (`modInit.ToggleConfigSheets`) hides Print Technicians, Printers, Papers and Settings using `xlSheetHidden` (not `xlSheetVeryHidden`, so **Unhide** still reaches them — this is UI tidiness, not a security boundary) and relabels itself between the two states by reading which of the four sheets it can currently find. **Stays on Summary regardless (snag 3a, 2026-09-22):** hiding the *currently active* sheet forces Excel to activate whatever is next in tab order — harmless when Summary (where the button lives) is already active, but `ToggleConfigSheets` now captures the `Summary` worksheet explicitly and re-activates it unconditionally at the end, rather than relying on it never having moved.

### 8.3 Reports sheet — rebuilt 2026-09-21 (was "Cost Calculations")

A live `FILTER`+`SORTBY` driven by criteria cells; results update as criteria are typed, and can now be sorted by any result column.

**Filters** (constant `SHEET_REPORTS = "Reports"`):

| Cell | Filter | Behaviour |
|---|---|---|
| B5 | Student name | Fragment search, case-insensitive |
| B6 | Student number | Exact match after trimming |
| B7 / B8 | Date From / To | `*1`-coerced (tolerant of text dates); To-date test is `< end + 1` so a job logged at 16:30 on the closing date is not excluded |
| F4 | Location (print room) | Dropdown, exact match — every registered room, not just those with a job logged (2026-09-25; moved here from O4 2026-09-27, see below) |
| F5 | Technician | Dropdown, exact match |
| F6 | Printer | Dropdown, exact match |
| F7 | Paper Stock | Dropdown, exact match |
| F8 | Quantity | Dropdown, exact match |
| O10 | **Export names** (label at N10) | **Yes/No, defaults to No — governs the exported report only; the live results always show names. Snag 2a, see below** |

**Filter block reordered (2026-09-27, direct user request).** Technician/Printer/Paper Stock/Quantity moved down one row each (`D4:D7`/`F4:F7` → `D5:D8`/`F5:F8`), and Location (print room) moved into the row they vacated (`D4`/`F4`) — previously at `N4`/`O4`, parked there only to dodge the columns `ApplyReportsMinimumColumns` hides by default (below). The whole block now reads top-to-bottom as one group: Location, Technician, Printer, Paper stock, Quantity.

Layout convention (snag item 2, resolved): **label → input → hint → gap**, input immediately to the right of its label, with an unused trailing column separating the two filter groups — replacing the earlier NAME | gap | INPUT | hint arrangement.

**Export names (snag 2a, 2026-09-22, moved to `O10` on 2026-09-23; made export-only and renamed from "Show names" 2026-09-29, direct user request) — data protection.** Toggles whether an **exported** report (`modExport.ExportReportSnapshot`) contains real Student name/Student no values, defaulting to **No** (opt in to reveal, not opt out). The **live Reports results always show names** — the toggle has no reference in the results formula any more. `modExport.BlankNameColumns` blanks the **values** (not the columns) of the two columns in the snapshot block unless `O10` is exactly `Yes` (anything else, including an empty cell, counts as No), before `RemoveExcludedColumns`/`PromoteUniformColumns`; the blank columns stay in the file, and an all-blank column is never promoted. The label at `N10` reads "Export names" and its input message says the results shown on the sheet always include names. Before 2026-09-29 the toggle blanked the values in the `A16` formula itself, which hid names from the live view too and reached exports only because the snapshot copied the live values. Tests: `test-reports.ps1` (live view shows names under No and Yes; label; export blank under No, populated under Yes).

Originally sat at `E9`/`F9` (row 9, the same row as the disjoint-criteria warning at `A9`), moved to `N10`/`O10` (labelled "Show student/department name/no", alongside Sort by/Sort direction) once the missing-label bug below was traced to its real cause: `modReports.ApplyReportsMinimumColumns` hides entire **columns** by results-header name — reaching every row on the sheet, not just the results table — and `E` (`"Printer"`) was one of the hidden ones while `F` (`"Paper stock"`) wasn't, so the label sat in a column that vanished under the default view while its dropdown, one column over, stayed visible with nothing beside it. `N`/`O` (`"Chargeable"`/`"Paid"`) are both permanently kept, so label and dropdown now survive the default view together. `E10`/`G10` (Sort direction's own label/hint) have this identical latent bug and were deliberately left as found — flagged for a general fix rather than patched column-by-column, since the underlying tension (row-level chrome sharing columns with a results-only hide) is still being decided.

**A real bug this surfaced, worth remembering for any future per-row toggle like this one: `IF(<scalar cell>="Yes", <column array>, "")` is not an elementwise blank.** The condition here is a single toggle cell — a scalar — so with the toggle off the whole `IF` collapses to the bare scalar `""` rather than a column of blanks matching every other `HSTACK` argument's height. This is the exact "HSTACK does not broadcast a scalar against a column" trap §7.1's own `_Data` formula already documents, just met again in a new place. The fix nests the toggle inside an **outer** `IF` whose own condition (`Job ID <> ""`) is already a genuine per-row array — once the outer `IF` is evaluating elementwise, the inner one is too. The visible symptom before the fix: every row past the first silently showed `FILTER`'s own "no data" fallback text instead of real job data, because `HSTACK` couldn't reconcile a 1-tall scalar against 17-tall arrays.

`A9` carries the disjoint-criteria warning (name and number both given, never appearing together on the same row) — red bold text, only fires when both are non-blank.

**Row 11 is deliberately blank** (snag item 1, resolved) as a gap before the "Matching" totals row.

**Sort controls** (`A10`/`B10` sort column — a dropdown of the result headers, Job ID excluded; `E10`/`F10` sort direction — Ascending/Descending) sit **below the filters and above the totals row** (snag item 3, resolved), not beside the Technician/Printer/Paper/Quantity group.

**Filter and sort rows are grouped** (snag item 15, resolved): rows 5:10 — both filter groups, the name/number warning, and the sort controls — collapse together as one outline block.

**Totals** ("Matching", rows 12–13) mirror the current filtered set: Jobs, Gross, Disregarded, Chargeable, and (2026-09-22, snag 1c) Paid/Unpaid. **Laid out differently from Summary's totals since 2026-09-22**: each metric's label sits directly **above** its value (row 12/row 13, same column: `MatchTotal`) rather than to its left, and both only ever land on columns `B`, `C`, `D`, `F`, `N`, `O` — the same columns the minimum-columns view below never hides. This exists because the results table grew past 14 columns (Student Name/No, Paid) and several of its columns are hidden by default (2d) — a totals cell sharing a column with a newly-hidden results column would display blank, the identical class of bug §4.1's location-sheet fix found. When the columns between `B`/`C`/`D`/`F`/`N`/`O` collapse (the default state), the six end up rendering adjacent anyway, so nothing looks gapped day to day. The header row itself (`A12:O12`) is tinted `RGB(244, 232, 222)` (0.9.13) — the results header's own `RGB(222, 232, 244)` with red and blue swapped, same lightness/saturation but warm rather than cool, so the two bands read as related but distinct; Summary's own totals-breakdown header (§8.2, row 9) uses the same warm tint rather than the blue it used to share with this sheet's results header, for the same "totals table, not a record list" reasoning.

**A second real bug found alongside the first: `SUM(FILTER(array, ok)) - SUM(FILTER(array, paidOk))` errors instead of returning the full total when `paidOk` matches zero rows** (the ordinary "nothing marked Paid yet" case). `FILTER` with an all-`FALSE` criteria raises `#CALC!` rather than returning an empty array unless given a third `[if_empty]` argument — and an error on *either* side of a subtraction makes the whole expression an error, which the outer `IFERROR` then silently replaces with `0` instead of the correct, perfectly valid `SUM(FILTER(array, ok))` term. Fixed by passing `0` as `FILTER`'s third argument on both terms (`FILTER(array, criteria, 0)`), so a no-match branch degrades to `0` locally rather than poisoning the whole calculation. Found by testing the everyday "nothing paid yet" case specifically, not just the "one row marked Yes" case the original test already covered — worth remembering for any future `SUM(FILTER(...))−SUM(FILTER(...))` pattern.

**Results table** header at row 15; records spill from `A16` via one `FILTER`+`SORTBY` `LET` formula. Columns (2026-09-22: Student Name/No and Paid added, snag 2a/2d): Date/Time, Location, Student Name, Student No, Printer, Paper stock, Qty, Unit, Area m², Paper cost, Consumable cost, Gross, Disregarded, Chargeable, Paid, Technician, Notes — then a **hidden Job ID column** appended after Notes, added specifically as the correlation key **Delete visible records** (§10.6) needs to map a visible row back to its source location sheet and row. Date/Time (`A`) is fixed at 100px (0.9.14, `modUtils.ColWidthForPx`) rather than the sheet's general 14-character default, wide enough for a full `dd/mm/yyyy hh:mm` stamp.

**Paid displays a phantom `0` for a row that predates the column without a fix.** A cell `INDEX` references that has never held any value reads back as the number `0`, not an empty string — the classic "reference to a genuinely blank cell returns 0" Excel behaviour — and `&""` does **not** fix it (`0&""` is still the text `"0"`, merely stringified after the fact). Since Paid only ever legitimately holds `"Yes"`, `"No"` or blank, never a real `0`, the fix tests for that specific value: `IF(<Paid>=0,"",<Paid>)` — safe here in a way it wouldn't be for a field that could legitimately be zero.

They cannot use `SUMIFS` for the breakdowns below — its arguments must be ranges, and after `FILTER` these are arrays — so per-key aggregation is `BYROW` with a `LAMBDA`. **Breakdowns sit to the right of the results** (`T15` by print room, `X15` by paper stock — `key | Jobs | Gross | Chargeable`, shifted from `Q15`/`U15` on 2026-09-22 when the results table grew by 3 columns), for the same reason totals sit above the Summary table: the list spills to an unknown height.

**Minimum columns (snag 2d) — the results table keeps every column, but only seven stay visible by default: Date/Time, Location, Student Name, Student No, Paper stock, Chargeable, Paid.** The rest (Printer, Qty, Unit, Area m², Paper cost, Consumable cost, Gross, Disregarded, Technician, Notes) are hidden, never removed — `modReports.ApplyReportsMinimumColumns`, applied once at build time, not a user-facing toggle (customising the set is a documented future enhancement, matching the snag list's own note). This reuses the exact same "hide by header name, touch nothing else" technique as `modInit.ApplyColumnVisibility` (the location sheets' reduced-clutter view, §4.1), reimplemented locally because the results table is a spilled array with a header row, not an Excel Table/ListObject `ApplyColumnVisibility` could be called against directly.

**Freeze panes** sit at `A15` (§9.4), so the filters, totals and header row stay visible while scrolling the results (snag item 9, resolved) — this replaces the freeze panes removed from the four config sheets (§9.4).

**This is the whole of AT-10 for v1.** A student's total across every print *room* in this workbook is in scope and served here; only the cross-*workbook* case is deferred (D17, §12).

**Verified** by `test-reports.ps1`: all criteria combinations including Technician/Printer/Paper Stock/Quantity and sort-by-any-column, single-day ranges including times, partial and case-insensitive names, the disjoint-criteria warning, and a date left as text. Breakdowns reconcile exactly to the Summary totals.

### 8.4 Reports filter dropdowns

`RefreshReportFilterLists` (called from both `BuildReports` and `RefreshLocations`, since the latter changes `_Data`):

- **Student name / Student number**: non-strict autocomplete (`ApplyTo ..., Strict:=False`) over previously-recorded values — offered for convenience, but free text (or an unlogged number) is still accepted, since there is no authoritative student list (D1).
- **Technician**: recorded values, strict list.
- **Printer / Paper Stock**: **every active catalogue entry**, not just ones actually used yet (snag item 4, resolved) — deliberately independent of each other rather than cross-filtered by compatibility, so an incompatible combination on Reports simply returns an empty result rather than being prevented at the filter stage (unlike the location-sheet entry dropdowns, §7.2, which do enforce compatibility because an incompatible *job* would be a real error).

**No calendar-icon date picker** was added to the From/To date cells (snag item 5). Checked and confirmed: the pop-up calendar icon on a date-formatted cell is an Excel-for-the-web feature and never ships on desktop Excel, Windows or Mac — this is a platform gap, not a bug, and is recorded as such rather than worked around. A custom worksheet-based picker (in the style of `modPicker`'s no-ActiveX multi-select) was considered and explicitly deferred rather than built.

### 8.5 Where the commands live

| Sheet | Buttons |
|---|---|
| Summary | Refresh Locations, Check workbook, **Hide/Show settings sheets** (Go to Settings removed, 2026-09-22 — see §16.4) |
| Settings | Two rows of four, below `tblSettings` (0.9.12, §16.4 addendum): Refresh Locations, Check workbook, Re-stamp prices…, About / Export All Locations…, Import (choose room)…, **Backup workbook…**, **Restore workbook…** (§10.7) |
| Each location | Add Print Job, Now (toolbar, row 13); Remove Row, Select printers…, Check this sheet, Clear All, Export…, Import…, **Reduce clutter / Show all columns** (side panel, §4.1, 2026-09-25); **Clear defaults** (row 4, over Technician) |
| Reports | **Export report…**, **Delete visible records…** (column T, rows 1 and 3); **Mark all as…** cluster — label `O1`, **Paid** (row 2) and **Unpaid** (row 3) buttons (§10.9) |

Settings keeps its own copies deliberately: it is where someone lands when configuring, and *Re-stamp prices* and *About* belong nowhere else. Export report / Delete visible records sit at the top of the Reports sheet (rows 1 and 3, column T since 2026-09-26; the *Mark all as…* cluster just left of them at column O), inside the first screenful, above the filter rows.

---

## 9. VBA architecture

### 9.1 Modules

| Module | Responsibility | State |
|---|---|---|
| `modMain` | Public entry points bound to buttons. Thin wrappers; names are a stable contract with the Form Controls | Built |
| `modJobs` | AddPrintJob, StampNow, RemoveRow, ClearAll | Built |
| `modSnapshot` | StampRow, ReStampAll, LogAudit | Built |
| `modValidation` | OnCellChanged, CheckSheet, CheckWorkbook, student consistency scan | Built |
| `modLists` | Dependent dropdowns and their staging columns; reused for Reports filter autocomplete | Built |
| `modRegistry` | RefreshLocations, EnsureSystemSheets, code assignment, table renaming, button healing, consolidated-range formula, **Job ID high-water mark** | Built |
| `modReports` | **Builds** the Summary and Reports sheets — layout and formulas, once; Reports-page bulk delete and bulk Paid/Unpaid; the shared at-least-one-filter safeguard; filter-list refresh | Built |
| `modExport` | Per-location CSV, **Export All Locations**, **Export report snapshot**, export-folder resolution (incl. the Mac OneDrive fix), the fingerprint, the `LOC_Export` status line | Built |
| `modImport` | **Restores or merges an exported CSV into a location's job table** | Built |
| `modVersion` | Version constants, settings rows, the About popup, document properties | Built |
| `modPicker` | The multi-select picker | Built |
| `modInit` | One-time/re-runnable setup: draws the Form Controls, applies protection, unlocks config inputs, sets/clears freeze panes, reorders tabs, groups columns, wraps Settings notes, hands over to RefreshLocations | Built |
| `modCatalog` | Configuration loaded once per operation; **catalogue row Add/Remove; Clear table (§10.8)** | Built |
| `modSettings` | Typed accessors for settings and named ranges | Built |
| `modProtect` | Protect/unprotect wrappers, re-applied on open, **no default password** | Built |
| `modUtils` | Application state, messaging, quiet mode, table and name access, ID generation, schema version constant | Built |

| Class | Responsibility | State |
|---|---|---|
| `clsDict` | Keyed collection, built on `Collection` — §9.3 | Built |
| `clsStock` / `clsPrinterDef` | One paper stock / one printer | Built |

**`modReports` exists, but computes nothing at run time.** It writes layout and formulas once. A defect in `modReports` can only produce a wrong *sheet*, visible immediately, not a wrong *figure* on a sheet that looks right.

`clsCatalog` and `clsLocation` were not needed. `modCatalog` holds configuration in module-level `clsDict` instances with an `Invalidate` flag; location sheets are addressed through `modUtils` helpers.

### 9.2 Conventions

- `Option Explicit` in every module.
- Every public entry point: disable events/screen updating → `On Error GoTo Fail` → work → restore state → exit. `AppOff`/`AppOn` are depth-counted so nesting is safe.
- **`EnableEvents = False` during our own writes is not negotiable.** Handlers write cells, and writing a cell re-enters `Worksheet_Change`. `AddPrintJob` writes five cells on a row that does not exist yet; with events live, validation would judge a half-built row, emit spurious Status warnings, and could clear a legitimate Print Width. `ClearAll`, `ReStampAll` and `ApplyImport` would each run a validation pass per cell write. Anything needing to know that data changed must **derive** it, not listen for it.
- **Derive state; do not track it.** Applied throughout: `IsLocation` reads a marker cell (§4.1); the registry rescans (§4.2); the export fingerprint is computed on demand (§10.4); the Reports export signature likewise (§10.4); report figures are formulas rather than stored results (§8); the Job ID high-water mark is the one deliberate exception — it *is* tracked state, because a derived value (a row scan) is exactly the bug it fixes (§3.2). Derived state cannot rot, survives a user editing with macros disabled, and survives an old copy of a sheet being dropped back in — all three defeat a tracked flag silently, which is why the Job ID mark is still read as a **floor** under a scan rather than trusted blindly.
- **Reports address `_Data` columns by header name**: `INDEX(_Data!$A$10#,,MATCH("Qty",_Data!$A$9:$AZ$9,0))`. Wordier than `INDEX(...,,10)`, and the reason the job table can be reordered without touching a report formula.
- **A criteria expression must be an *array*, not a scalar.** `IF($C$5="",TRUE,…)` returns a bare `TRUE` when the box is empty, and `FILTER(column, TRUE)` is `#CALC!`. With all criteria blank the product is the scalar `1`, so Reports failed in its most ordinary state. Every criteria product is seeded with a column-shaped term (`Job ID <> ""`) that fixes its height. This still applies with the extra Technician/Printer/Paper Stock/Quantity criteria added in the 2026-09-21 rework — same pattern, one more multiplied `IF(...)` term each.
- **Coerce both sides of a date comparison** with `*1`. A user typing `16/09/2026` into an unformatted cell leaves text behind, and `number >= text` is FALSE for every row without raising anything — a filter that silently returns nothing.
- **A staging sheet used to reach Excel's CSV/xlsx writer must be formatted as Text.** The export formats dates as ISO and the schema version as `1.0`, then writes them into cells; without `NumberFormat = "@"` Excel re-parses both — the date becomes a date value again and is written out in the machine's locale (`9/15/2026` here), and `1.0` becomes the number `1`, so an importer checking for `1.0` sees `1`. Applies equally to `modImport`'s read side and to the Export report `.xlsx` snapshot.
- **Totals go above a spill; secondary tables go beside it.** A spill's height is unknowable, so anything below it is displaced the moment the data grows. Applied to Summary's totals, Reports' totals, and Reports' breakdowns.
- Formulas written from VBA use `.Formula2`, never `.Formula`. `.Formula2` is array-aware; `.Formula` applies implicit intersection and silently stores a single value where a spill was intended.
- No hard-coded row or column numbers: columns are resolved by header name.
- **A module-level `Const`/`Dim` must sit in the declarations section at the top of its module, clustered with the module's other module-level declarations — never mid-file between two procedures.** Every module in this project already followed this by convention; a `Private Const` declared after several `Sub`s had already appeared (2026-09-22, `modInit`'s reduced-clutter view) compiled visibly, imported without error, and still failed the whole module with "Compile error: Variable not defined" pointing at a *later, correctly-spelled, correctly-declared* use of it — a genuinely confusing failure mode, since the declaration itself looks completely valid in isolation and the error message doesn't mention the real cause. `build.ps1`'s own quiet-mode wrapping cannot catch this either: a compile error is a blocking VBE dialog raised before any of `InitialiseWorkbook`'s own `On Error` handling ever runs, so it hangs unattended automation rather than reporting cleanly — diagnosed here by reading the dialog's Win32 window text directly, since the COM call blocks and returns nothing over the wire.
- **Never `Val()` on a date or numeric cell.** `Val` takes a `String`, so a date is coerced to text and its leading digits read (`15/09/2026 10:24` comes back as `15`, a date in January 1900), and a number round-trips through the machine's decimal separator (comma-decimal locales silently read every paper cost, consumable rate and stock dimension as zero). `modUtils.DateSerialOf` and `modUtils.NumOf` read `.Value2` and are the only sanctioned routes. The Job ID suffix scan (`ScanMaxSuffix`) is the one place `Val()` is still used deliberately — a digits-only suffix has no fractional part, so the locale trap does not apply there.
- **`On Error Resume Next` around a block, never around a batch of independent writes.** When several independent operations are wrapped in one blanket handler, a failure is invisible and untraceable to which operation caused it.
- **Quiet mode.** `modUtils.SetQuiet` switches `Say` from `MsgBox` to collecting messages for `QuietLog`, and makes `Ask` return **False**. A script driving the workbook over COM has nobody to dismiss a dialog, and one `MsgBox` hangs the run indefinitely. `Ask` returning False is deliberate: an unattended run must never confirm a destructive operation on the user's behalf — and it has a second use, since it lets a test read a destructive command's confirmation text while guaranteeing the command aborts.
- **Show what's about to be lost, before losing it.** `RemoveRow`, `ClearAll`, config-row removal (§3.3), Import's overwrite count (§10.5) and Reports' bulk delete (§10.6) all confirm with the specific record count (and, where relevant, a per-location breakdown) rather than a bare "Are you sure?".
- **Read `Err.Number`/`Err.Description` first thing in a handler.** Every form of `On Error` resets the `Err` object, so reading it after any cleanup step (e.g. reprotecting a sheet) loses the description of what actually went wrong.
- **VBA's `And`/`Or` do not short-circuit - both operands are always evaluated.** `If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then` (a pattern used in five modules - `modJobs`, `modImport`, `modCatalog`, `modSnapshot`, and 2026-09-22's `modBackup`) still calls `IsBlankRow(lo, 1)`, which indexes `ListRows(1)`, even when `Count = 1` has already evaluated to `False` - there is no early exit the way there would be in most other languages. Harmless while `Count >= 1` (index 1 is always in range then), but "Subscript out of range" the moment `Count = 0`, a state nothing in ordinary use ever produced until restoring a backup into a table emptied down to zero rows (§10.7) reached it for the first time. Fixed at the root: `modUtils.IsBlankRow` now returns `False` for a row number outside the table's current range instead of indexing blindly, protecting every call site rather than restructuring each one's `If` around the non-short-circuiting operator.
- **`Any` is a reserved VBA token even outside a `Declare` statement.** `Dim any As Boolean` (`modExport.PromoteUniformColumns`, 2026-09-22) compiled invisibly as valid-looking source and imported without error, then broke the whole module with a bare "Compile error: Syntax error" and no line number — the same class of confusing, hard-to-spot compile failure as the module-level `Const`-placement bug above, just a different trigger. Diagnosed the same way: the VBE's blocking dialog read via Win32 window text, since the COM call hangs rather than raising. Renamed to `hasValue`; the fix generalises to avoid `Any` (and, on the same reasoning, other `Declare`-only contextual keywords) as an ordinary identifier anywhere in this project, not just here.

### 9.3 Cross-platform strategy

**Platform is the binding constraint**, not Excel version (D12): Excel for Mac has no ActiveX and no COM automation.

| Unavailable on Mac | Approach taken |
|---|---|
| `Scripting.Dictionary` | `clsDict`, built on `Collection` with error-trapped lookup |
| `Scripting.FileSystemObject` | Not used |
| `ADODB.Stream` | Not used — CSV/xlsx written via Excel's own `SaveAs` |
| ActiveX controls | Form Controls only |
| `Application.FileDialog` (save side) | Export writes to a resolved folder path (§10.4), not a chooser |
| UserForm rendering differences | The picker is a **worksheet** pretending to be a dialog (`_Picker`) |
| COM automation | The build script is Windows-only; Mac uses the manual import route (SETUP.md) |
| `Environ$("OneDrive")` / `OneDriveConsumer` / `OneDriveCommercial` | These env vars are Windows-only. On Mac, local OneDrive roots are found by scanning `~/Library/CloudStorage/OneDrive*` instead (§10.4) |

- **Protection must be re-applied on open.** `UserInterfaceOnly:=True` is not persisted; a workbook saved in that state reopens fully protected and VBA can no longer write to its own sheets.
- **1900 date system enforced** on open, since an older Mac workbook may carry 1904 and shift every date by four years.

### 9.4 Protection model

| Element | State |
|---|---|
| Job-row input cells | Unlocked |
| Job-row formula, snapshot and header cells | Locked |
| **Configuration table data-body cells** (Technicians, Printers, Papers, and the four Settings lookup tables) | **Unlocked** (snag item 6, resolved) |
| **`tblSettings` Value cells not marked "Read-only" in Notes** | **Unlocked**; read-only rows (schema version, build stamp, last-refresh) stay locked |
| Sheets | `UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True, DrawingObjects:=False, EnableOutlining:=True` |
| Hidden sheets | `xlSheetVeryHidden` |
| Workbook structure | **Unprotected** — required so users can duplicate location sheets (§4.4) |
| VBA project | Locked for viewing |
| **Sheet protection password** | **None by default** (removed 2026-09-21, snag item 7). Set one by hand in Excel (Review → Protect Sheet) on a specific workbook if required — nothing in `modProtect` generates one |

Structural operations route through `modProtect` wrappers that unprotect, act and reprotect within a single error-guarded call, because adding rows to a `ListObject` on a protected sheet is unreliable even with `UserInterfaceOnly`.

**Freeze panes, current layout (snag items 8 and 9, resolved):**

| Sheet | Freeze panes |
|---|---|
| Settings, Print Technicians, Printers, Papers | **None** — removed as unneeded visual clutter, since these are configuration tables typically viewed in full |
| Reports | **`A15`** — added, freezing the filters/totals/header above the results table |
| Location sheets | Unchanged (whatever the `.xlsx` ships) |
| Summary | Unchanged |

### 9.5 Sheet tab order

`modInit.ReorderSheetTabs`, run at the end of every `InitialiseWorkbook`, produces (snag item 12, resolved):

**Summary → Reports → [location sheets, in workbook order] → Print Technicians → Printers → Papers → Settings**

Hidden system sheets (`_Data`, `_Registry`, `_Audit`, `_Work`, `_Picker`) are `xlSheetVeryHidden` and excluded from tab-order logic. `modReports.SheetNamed` places `Reports` at tab index 1 and `Summary` at index 2 on a first-run build.

**A `Move Before:` gotcha, worth remembering for any future tab-order code:** `Worksheet.Move Before:=` only moves a sheet *earlier*. Moving one that already sits before the target is a no-op — removing it from its old position shifts everything down by one and it lands right back where it started. Moving a sheet *later* needs `Move After:=`. This shipped the Summary/Reports tabs reversed in every build up to 0.7.0, because the code asked for a later move using `Before:`.

---

## 10. Destructive operations, export and import

### 10.1 Remove row

Acts on the selected row. Confirms with the job's identifying detail — ID, date, student, printer, cost — never a bare "Are you sure?". Writes an `_Audit` entry, then deletes the row.

### 10.2 Clear All — verified

Scoped to the current location sheet. States the location name, record count and date range:

> *Clear All — Example Print Room*
> *This will permanently delete 5 print jobs dated 15/09/2026 to 17/09/2026.*
> *Records on other print room sheets are not affected.*
> *This cannot be undone. Continue?*

Where the location has unexported changes, the confirmation says so.

### 10.3 Audit note

VBA operations clear Excel's undo stack. `_Audit` records what was removed, when and by whom. Sheet deletion is the one destructive act that cannot be caught (§4.3).

### 10.4 Export — built, extended 2026-09-21

**Purpose**, several at once: a recovery path for a deleted location sheet (§4.3), a way to get records out for finance, the transport for manual multi-workbook aggregation (§12.4), and — new — a point-in-time archive of a filtered Reports view for physical/human records.

**Per-location export** (`modExport.ExportLocation`) and **Export All Locations** (`modExport.ExportAllLocations`, new) — the latter runs every location through one `AppOff`/`AppOn` bracket and shows a single summary dialog of done/skipped/failed per room, rather than requiring one click per location.

**Content — the one irreversible decision.** Every column, **including the snapshot block**. An export without `S_UnitCost`, `S_ConsRate` and the rest is permanently unimportable: re-importing would recost every row at today's prices, destroying exactly the integrity AT-09 protects. `H_Issues` is the single exclusion — an internal working cell behind the Status message, meaningless outside the workbook and recomputed on import (well, on the next Check workbook sweep — §7.3).

**Format.** CSV, written by copying to a temporary workbook and using `SaveAs FileFormat:=xlCSVUTF8`. Excel's own CSV writer, so quoting and comma escaping are its problem, and `xlCSVUTF8` gets `£` and any non-ASCII student name right. The staging sheet is formatted as Text throughout.

A header block — schema version, site ID, location code and name, generated-at, row count — then a header row, then records. Dates are written `yyyy-mm-dd hh:nn:ss`: a serial would be unreadable and a locale-formatted date ambiguous, since `06/07` is two different days depending on who opens it.

**Column order is not part of the contract.** The header row names the columns and any importer maps by name, so the job table may be reordered freely without invalidating older exports or bumping the schema version. The file uses a canonical order fixed in `modExport`.

**Naming.** `PrintCosts-<SITE>-<LOC>-yyyymmdd-hhmm.csv`, written to the export folder (below).

**Selectable export folder (snag item 10, resolved).** A `SET_EXPORT_FOLDER` setting (`EXPORT_FOLDER` internally), editable on the Settings sheet, read by `modExport.ExportFolder()`. If blank, or if the folder named there does not currently exist **on the machine running Excel**, the export falls back to the workbook's own resolved location — the setting is checked, not trusted, so a folder that only exists on one machine cannot silently swallow another user's exports. The row ships in `PrintCosts.xlsx`.

**Resolving "beside the workbook" — including the OneDrive fix (snag item 1, resolved).** `ThisWorkbook.Path` is not reliably a filesystem path: for a OneDrive-backed workbook, Excel may report the service URL (`https://d.docs.live.net/...`) instead. `ExportFolder()`:

1. Uses the export-folder setting if valid (above).
2. Otherwise, if `ThisWorkbook.Path` is not itself a URL, uses it directly — the ordinary case on both Windows and Mac.
3. Otherwise (a OneDrive service URL), strips the scheme and host and tries progressively shorter path tails against every candidate local OneDrive root, returning the first that resolves to a real folder.
4. Candidate roots (`CandidateOneDriveRoots`): on Windows, the `OneDrive`, `OneDriveConsumer` and `OneDriveCommercial` environment variables. **On Mac these variables don't exist**, so the function additionally scans `~/Library/CloudStorage/OneDrive*` — this Mac-specific scan is exactly what the 2026-09-21 fix added; without it, a fully-synced Mac OneDrive folder still produced *"The workbook's own folder could not be resolved to a location on this computer"*, purely because the only roots ever tried were Windows environment variables.
5. If nothing resolves, the export **refuses with an explanation** rather than guessing.

**The durable check, regardless of platform: the export verifies the file exists on disk before stamping anything as exported.** An earlier defect had the export announce success having written nothing (because the resolved path was wrong), then stamp the location as exported — which meant the §4.3 warning between a sheet deletion and its records said the opposite of the truth. That check is what makes the whole export status trustworthy.

**Import is a later revision — see §10.5, now built.**

**Knowing what is unexported — derived, not tracked.** A per-location fingerprint computed on demand:

```
row count + CountA(data range) + Sum(Qty) + Sum(Chargeable Cost) + Max(Date/Time)
```

compared against `tblLocations.Export sig` from the last export. Nothing hooks any event and no mutator must remember anything.

*The limitation, plainly:* a collision needs an edit leaving count, cell count, both sums and the latest date unchanged — two compensating changes in one location between exports. Constructible deliberately, vanishingly unlikely by accident.

**Export report — a separate signature, for a separate artefact (new, "Problem 2/3").** `ExportReportSnapshot` writes a **static-value `.xlsx`** (not CSV) of the Reports sheet's current filtered-and-sorted results, chosen over CSV/PDF specifically because it (a) carries enough formatting to be readable, (b) opens directly and consistently on most computers, (c) stays modifiable with ordinary software, and (d) is machine-readable text/metadata with no OCR needed — a PDF whose text layer turned out not to be extractable at all is exactly what motivated ruling PDF out here (§16, the removed design-doc PDF). It formats money/quantity columns, then stamps its **own** signature — same shape as the export fingerprint (row count, cell count, chargeable-cost sum, max date) but of the *filtered Reports view*, not a location sheet — into hidden cells on Reports (`AN1`/`AN2`). This is what **Delete visible records** (§10.6) compares against.

**Header block and single-value promotion (snags 2b/2c, 2026-09-22).** The `.xlsx` opens with a metadata block above the table — report title, schema version, site ID/name, the **date range covered** (the From/To filter boxes, `B7`/`B8`, when either is set; otherwise the actual `MIN`/`MAX` of the exported rows' own Date/Time column), a generated-at timestamp, and the row count — then a blank row, then the results table. Unlike the per-location CSV's header block, there is no single "Location" line here by default, since Reports spans every location in the workbook; if a filtered report happens to cover only one, that is exactly what single-value promotion (next) surfaces.

Any of **Student name, Student no, Location, Printer, Paper stock, Technician** that holds the **same non-blank value on every exported row** is lifted out of the table and written as a `"Header: Value"` line in the metadata block instead — filtering a report down to one printer, say, states that fact once at the top rather than repeating it down an entire column of identical text. A candidate column with no non-blank values at all (Student name/no when the 2a toggle is off, most commonly) is left exactly as it is: there is nothing true to promote, and removing an all-blank column would read as data loss rather than tidying. This is `modExport.PromoteUniformColumns`, called only from `ExportReportSnapshot` — the **live Reports sheet is never touched** by this, per the resolved review answer: promotion is an export-time-only transformation, not a second "reduced" view of the sheet itself.

**A real bug found building this.** The column bound previously used to decide how much of the results table to export (`LastVisibleColumn`) filtered by `.Hidden` — harmless while only the trailing, always-hidden `Job ID` column was ever hidden, but §8.3's minimum-columns view (0.9.6) now hides `Printer` and `Technician` among others by default. The old function silently truncated every export at the last **visible** column, dropping `Technician` and `Notes` from the file entirely with no error raised anywhere — the kind of defect nothing short of opening the file and counting columns would catch. Renamed `LastHeaderColumn` and rewritten to bound by header **name** (stop before the trailing `Job ID`) rather than by `Hidden` state, since the promotion logic specifically needs those hidden columns' values to check for uniformity.

**Data protection.** Export multiplies the places student names and numbers live, in files outside the workbook's protection. Not an objection, but the reason exports land in a resolved local folder rather than an arbitrary chooser, and a point for whoever writes the operating procedure. It connects directly to the deferred O7 (§12.3) and its BONUS resolution path (§12.4).

**Verified** by `test-export.ps1` (per-location export and fingerprint) and `test-import.ps1` (Export All Locations).

### 10.5 Import — built ("Problem 1")

**Scope.** Always imports into **one** location's job table — there is deliberately no "import into every location" counterpart to Export All Locations. Two entry points: import into whichever location sheet is currently active, or a global command that first prompts for a destination location by code.

**What it covers, in one mechanism:**
- Restoring a backup into the location it came from (after data loss, or reverting a bad edit).
- Importing one location's exported jobs into a different room.
- Pulling several locations' exports into one copy of the workbook, for reporting or handover — see the BONUS note in §12.4.

**No cost reconciliation needed.** Job-row formulas only ever read the row's own snapshot columns and inputs, never `tblPapers`/`tblPrinters` (§5.1), so an imported row already carries the rates it needs and costs correctly regardless of what the target workbook's configuration tables contain.

**No catalog reconciliation either.** Import never creates, matches or edits rows in the target's `tblPrinters`/`tblTechnicians`/`tblPapers`/`tblConsumables`. An imported row's Technician/Printer/Paper Stock values are carried as plain text even when absent from the target's current lists. New jobs entered in that workbook afterwards continue to be driven only by that workbook's own catalog tables — import never pollutes them.

**What is written.** Only the input and snapshot columns per row: Job ID, Date/Time, Student Name, Student No, Technician, Printer, Paper Stock, Unit, Qty, Print Width mm, Disregard Paper, Disregard Consumable, Notes, and every `S_*` snapshot column. The calculated columns (Area m², Paper Cost, Consumable Cost, Gross Cost, Disregarded, Chargeable Cost) are left to the job table's existing per-row formulas, which reproduce the historical figures exactly from the snapshot — precisely as a locally entered job does.

**Conflict handling is by Job ID** (globally unique, never reassigned — §3.2). No existing row with that ID: append. An existing row with that ID: **overwrite** it with the imported version. Overwrite, not skip, is deliberate: it is what makes "restore a backup over stale/edited data" and "re-aggregate overlapping exports" both work correctly.

**Before an overwrite is committed**, a confirmation names how many existing records will be appended vs. replaced, following the same "show what you're about to lose" pattern as `RemoveRow`/`ClearAll`.

**Imported rows bypass `Worksheet_Change`** (written via VBA with `EnableEvents` False, same as every other write the app makes to its own sheets), so Status/H_Issues do not populate automatically — a `CheckSheet` sweep runs immediately after import to fill them in, rather than leaving rows silently unchecked.

**File format read is exactly what Export writes** — same header set, same ISO date convention (`yyyy-mm-dd hh:nn:ss`) that avoids locale ambiguity on the way back in.

**Verified** by `test-import.ps1`: Export All Locations, then Import restoring into the origin location and into a different room.

### 10.6 Delete visible records — built ("Problem 3")

**Scope.** Confined to the Reports page, operating on **whatever the current filters show** — deliberately the higher-friction, encourages-a-backup-first path for bulk pruning, as opposed to per-location `RemoveRow`/`ClearAll` which remain unchanged and are sufficient for their own scope.

**Row mapping.** For each visible row: `Location` → `SheetForCode` finds the target worksheet; the hidden `Job ID` column (§8.3) → `FindReportRow` locates the actual table row on that sheet, which is then deleted. Values are copied into memory *before* any deletion begins, because the live spill range shrinks as rows are removed elsewhere in the same run.

**Confirmation content:**
- The total record count and a **per-room breakdown**.
- A staleness warning comparing the current filtered set's signature (§10.4) against the signature stamped by the last **Export report** run — either *"Export report has never been run for a filtered set like this one"* or *"The filters or underlying data have changed since the last Export report"*. This replaces the original plan's "warn if not exported to CSV" wording, since Reports-page deletion is keyed off the Export-report signature rather than the per-location CSV export status (they are different artefacts answering different questions).
- *"This cannot be undone."*

**Filter safeguard (2026-09-29, direct user request).** Refuses to run unless at least one filter is set — see §10.9. Checked first, before anything is counted, and again inside `DeleteVisibleReportsConfirmed` (the routine that actually deletes, and the one tests reach directly).

**Audit.** Logged to `tblAudit` as `Delete visible (Reports)`.

**Verified** by `test-deletereports.ps1`, alongside the Export report snapshot itself; the filter safeguard by `test-markpaid.ps1`.

### 10.7 Full workbook backup / restore — built (snag list item 4a)

The ad-hoc "get me back to where I was" path, distinct from the two artefacts above: not a per-location CSV (§10.4), not a filtered Reports snapshot (§10.4's Export report), but every catalogue/configuration table plus every location's job records, in one pass, sharing one timestamp. New module `modBackup.bas`. Buttons on Settings: **Backup workbook...** / **Restore workbook...**, alongside the existing Export All Locations / Import commands — the four now share the second of the two button rows below `tblSettings` (0.9.12, §16.4 addendum; previously stacked in column T).

**Backup All** (`modBackup.BackupAll`): runs the existing `ExportAllLocations` unchanged (own summary dialog) for job records, then writes one CSV per catalogue table — `tblTechnicians`, `tblPrinters`, `tblPapers`, the three Settings-page lookup tables (`tblPaperTypes`, `tblStandardSizes`, `tblConsumables`), and `tblSettings` itself (one of the seven, not a separate mechanism — `modUtils.Tbl` finds it on the Settings sheet the same way it finds the other three lookup tables there). Named `PrintCosts-<SITE>-CATALOG-<TableName>-yyyymmdd-hhmm.csv`, written to the same resolved `ExportFolder()` (now `Public`, reused unchanged rather than re-deriving the OneDrive-URL resolution logic — §10.4's own env-var gotcha lives there, and duplicating it would risk drifting out of sync).

**Deliberate reuse over a second CSV format.** A catalogue CSV's header block is padded to the same eight rows `modExport.BuildBlock` uses for a per-location job export (title/schema/site/generated/rows, then a deliberately blank row 8), so the table header always lands on row 9 and data on row 10 — exactly where `modImport.ReadImportRows`'s own hardcoded row numbers already look. Restore therefore reads catalogue rows with the **same, unmodified, already-tested function** that reads job rows; no second CSV parser exists in this workbook.

**Restore Workbook** (`modBackup.RestoreWorkbook`): the user picks **any one file** from a backup via `Application.GetOpenFilename` (the same Mac-safe picker `modImport.PickImportFile` already uses — never `Application.FileDialog`, §9.3). Every sibling file sharing the same trailing `-yyyymmdd-hhmm.csv` in the same folder is found and classified by filename — safe here specifically because this module wrote every filename it will ever read back, unlike an arbitrary file. A preview (row counts per table, per location) is shown before confirming; **"This cannot be undone."**

**Conflict handling — the same rule as Import, generalised.** Each catalogue table has a stable key column used for overwrite-vs-append matching, exactly mirroring Job ID's role for job rows:

| Table | Key column |
|---|---|
| `tblTechnicians` | TechID |
| `tblPrinters` | PrinterID |
| `tblPapers` | StockID |
| `tblPaperTypes` | Paper type |
| `tblStandardSizes` | Size name |
| `tblConsumables` | Consumable type |
| `tblSettings` | Key |

The four Settings-page lookup tables have no synthetic ID (§3.3) — their natural-key text column is already what every lookup in this workbook treats as their identity (`modCatalog`'s own dictionaries are keyed the same way), so reusing it here assumes nothing new. A column currently holding a **formula** (`tblPapers`' Measure/Cost unit) is left alone on restore — checked by `.HasFormula`, not a hardcoded per-table column list, so a future calculated column added to any catalogue table is protected automatically without a matching code change. A value that parses as a number is written as one (locale-aware `CDbl`, not text that merely looks numeric), so cost/width/height columns stay usable by downstream formulas.

**Same name, different ID (2026-09-29).** Because catalogue IDs carry their site (§3.2), a backup from another workbook has different keys and is *appended*, never overwritten. Its names may still clash with existing rows, and every lookup is by name, so `RenameIfNameTaken` keeps the incoming row under its own ID and adds the site read off that ID to its name - `HP T730 (SITE2)`, then `HP T730 (SITE2 2)` if that is taken too; an ID in some other shape gives `(imported)`. A row is never a clash with the row it is overwriting, so restoring the same backup twice is idempotent. The result message says how many rows were renamed. After the rows are written `modCatalog.SyncCatalogHwm` raises the table's ID counter to cover any IDs under this site's prefix (never lowers it).

**Settings rows marked "Read-only" are skipped on restore.** `APP_VER`/`SCHEMA`/`BUILT`/`BUILT_BY`/`LASTREF` describe the *current* build (the same "Read-only" convention `modInit.UnlockSettingsValues` already reads from each row's own Notes column) — rewinding them to a backup's old values would make the workbook misreport its own version and schema, which nothing else about a restore should touch.

**Location job records reuse `modImport.ApplyImportConfirmed` unchanged**, one call per location file found — same "existing Job ID: overwrite, new Job ID: append" rule, same audit trail, same post-import `CheckSheet` sweep. The location a file belongs to is read from its own header block (`modExport.BuildBlock`'s "Location code" line, row 5) rather than parsed from the filename, so a restore does not depend on assumptions about what characters a site ID or location code might contain.

**A genuine, previously-latent bug found building this — worth its own note in §9.2: VBA's `And` does not short-circuit.** `If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then` — copied from the existing pattern already used by `modJobs`, `modImport`, `modCatalog` and `modSnapshot` — evaluates **both** operands regardless of the first, so `IsBlankRow(lo, 1)` still ran, and raised "Subscript out of range" calling `ListRows(1)` on a table with **zero** rows, even though the left operand (`Count = 1`) had already failed and should have made the right operand irrelevant. Dormant for the whole life of this project so far: nothing in ordinary use ever leaves a table at genuinely zero rows (`ClearAll` and friends always leave the one blank templated row) — restoring a backup into a catalogue table someone had emptied by hand is the first code path to actually reach that state. Fixed at the root rather than at each of the five call sites: `modUtils.IsBlankRow` now returns `False` (not an error) for a row number outside the table's current range, which is what "is this specific row blank" should mean for a row that does not exist.

**Verified** by new `test-backup.ps1`: backs up a populated copy, empties a second copy's Printers table and one location's job table entirely (a genuinely destructive corruption, not a partial one — this is exactly what surfaced the bug above), restores from the first copy's backup, and confirms row counts match the original, an untouched table (Papers) is unaffected, the audit log records the restore, and running the same restore a second time is idempotent (no duplicate rows).

---


### 10.8 Clear table — Print Technicians, Printers, Papers (direct user request, 2026-09-29)

A **Clear table** button (row 4, column 5, beside Add row/Remove row) on each of the three catalogue sheets empties **that sheet's table only**, so a bureau can start one list afresh without disturbing the others (e.g. new technicians, same printers). The four Settings-page lookup tables and `tblSettings` deliberately have no such button.

`modCatalog.ClearCatalogTable` confirms first (default button **No**), naming the table, the row count and the first five rows, and stating the consequences: rows vanish from every print room's dropdowns; recorded jobs keep their frozen prices; the other configuration sheets are untouched; **it cannot be undone — use Backup workbook first**. Printers adds that rooms' "Select printers" choices refer to printers by name; Papers notes that the built-in `Supplied (Roll)`/`Supplied (Sheet)` stocks are not table rows and stay available.

`DoClearCatalogTable` is the unprompted core (split out because `Ask` always declines in quiet mode, so tests could not otherwise reach it). It writes an `_Audit` entry ("Clear table"), deletes rows 2..n, `ClearContents` row 1 (skipping any cell holding a formula), then `Invalidate`s the catalogue cache so dropdowns rebind. **One blank templated row is deliberately left**, for the same reason as `modJobs.ClearAll` (§10.2): with no data row left, Excel has nothing to copy formatting/validation from. `AddCatalogRow` reuses that blank row (`Count = 1 And IsBlankRow`) and `RowCount` reports it as zero rows.

Buttons: `btnClearTechnicians` / `btnClearPrinters` / `btnClearPapers` (`modMain`), drawn by `modInit.InitialiseWorkbook`. Test: `test-clearcatalog.ps1`.

### 10.9 Mark all as Paid / Unpaid, and the filter safeguard (direct user request, 2026-09-29)

**What it is.** A 1-column × 3-row cluster (`O1:O3`) in the Reports header strip, just left of Export report / Delete visible records: label **Mark all as…** at `O1` (cell text, written by `BuildReports`), a **Paid** button in row 2 and an **Unpaid** button in row 3 (drawn by `InitialiseWorkbook`, sized to their row rather than `DrawOne`'s 22pt so neither spills into the row below). Paid sets the `Paid` column to `Yes`, Unpaid to `No`. Handlers: `modMain.btnMarkPaid` / `btnMarkUnpaid` → `modReports.MarkVisibleReports`.

**Scope.** Exactly the records the Reports sheet currently shows — the spilled results range, read the same way §10.6 reads it. Records the filters exclude are never written. **The filters are not touched**: the routine reads the results and writes only to the source room tables, so every filter box and the sort settings are as the user left them (checked by `test-markpaid.ps1`, which compares all of them before and after).

**Confirmation.** `Ask()` first, in the shape of §10.6's prompt: *"Mark N visible records as Paid?"*, a per-room breakdown, how many are already at the requested value (left alone, not rewritten), and a reminder that hidden records and the filters are unaffected. Default button is No.

**The write** (`MarkVisibleReportsConfirmed`, Public so tests can bypass `Ask()`): Job IDs and locations are copied into memory first (writing `Paid` recalculates the live spill); each room's `Job ID` column is read **once** into a Job ID → row map, rather than rescanning the table per record as `FindReportRow` does; each affected room sheet is unlocked once up front and relocked together at the end and in the error path. Logged to `tblAudit` as `Mark visible Paid (Reports)` / `Mark visible Unpaid (Reports)`.

**Filter safeguard — shared with Delete visible records (§10.6).** `modReports.HasActiveFilter` / `RequireActiveFilter`: both bulk commands refuse to run unless at least one filter is set, so neither can be pointed at the whole job list by a stray click. It **mirrors `Criteria()`** rather than testing "is a box non-empty": a box counts only if it actually narrows the results. The free-text and dropdown boxes (name, number, room, technician, printer, paper stock) count unless empty or spaces-only; From date, To date and Quantity count only if `*1` coercion succeeds (`Criteria` ignores them otherwise). Sort by, Sort direction and Export names are not filters. The check runs first in the interactive entry point — before any prompt or count — and again inside each `…Confirmed` routine, which is the code that actually writes. **If a filter box moves, `HasActiveFilter`'s address list must change with `Criteria()`** (comment in both).

**Why column O.** Tried at V first: `test-reports.ps1` requires every Reports button within the first screenful (Left ≤ 900pt) and V sits at ~1020pt. A2's instruction text runs to ~689pt, so N or earlier would overlap its tail; O (~720pt) is clear of it, and is the always-visible `Paid` column of the results table, so a hidden column can never swallow the cluster.

**Not covered.** A filter that happens to match every record (a Location filter on a one-room workbook, say) passes the check — the safeguard is against *no* filter, not against a broad one.

**Verified** by `test-markpaid.ps1`: the cluster's caption, macro and geometry; `HasActiveFilter` against blank / spaces / uncoercible text / sort-only / real values; both commands refusing with no filter and changing nothing; only visible rows changing under a filter; filters unchanged afterwards; room sheets re-protected; the interactive path asking first; audit entries; Delete visible still working with a filter set.

### 10.10 Editing Paid directly on the Reports page (direct user request, 2026-09-29)

The **Paid** cells of the results table take the same Yes/No dropdown as the room sheets (`modReports.MakePaidEditable`: rows 16-2000 of the Paid column unlocked, list validation `Yes,No` with input/error messages). Everything else in the results table stays locked. Choosing a value writes it to the job record the row belongs to, in its own room, straight away.

**Why it is not just "unlock the column".** The results are **one spilled formula**. Typing into a cell of a spill range puts a constant there, turns the anchor `A16` into `#SPILL!` and so blanks *every* cell of the table — including the row's `Job ID` and `Location`, the only things that identify the record. By the time `Worksheet_Change` fires they cannot be read from the sheet. So the design is two-part:

- **`RememberReportsPaidCell`** (from `Workbook_SheetSelectionChange`, and `Workbook_SheetActivate` on arrival) notes which job the selected Paid cell shows: address, Job ID, Location, and the Paid value displayed. It keeps the **previous** selection's note as well, because Excel can move the selection after Enter either before or after it raises Change; the edit is matched to whichever note has the edited cell's address. `ResnapReportsSelection` re-notes the active cell after anything that changes the results under a selection that has not moved (a Paid edit, Mark all as…, Delete visible) and drops the stale previous note.
- **`OnReportsPaidEdited`** (from `Workbook_SheetChange`; returns at once for any edit outside the results' Paid column, so the filter boxes cost nothing) applies the edit to the noted job — `SheetForCode`, `FindReportRow`, the `Paid` cell, sheet unlocked and relocked — logs `Paid edited (Reports)` to the audit table (`<job>: <old> -> <new>`), then **clears the typed constant** so the spill returns.

**What is refused, and how.** Every refusal clears the constant first (an orphan constant left in the results area would block the spill the next time it grew that far), says why, and changes nothing: a cell that is not a record's Paid cell (below the results, or the "no jobs match" message row); a value that is not Yes/No (paste bypasses the dropdown's own check); a multi-cell write (paste/fill; *Mark all as…*, §10.9, is the bulk route); a record no longer found; and a **stale note** — the record's own Paid value must still equal what the row showed when it was selected, otherwise the results changed underneath the selection and the write would land on the wrong job. Re-picking the value already shown just puts the spill back.

**Costs, inherent to editing inside a spill.** The write clears Excel's Undo history (Ctrl+Z will not undo a Paid change; change it back instead), and edits are one cell at a time.

**Verified** by `test-paidedit.ps1`, which drives it the way a user does (a real `Select()` then a real cell write, so both events fire): validation and locking, an edit landing on the right record in both rooms with nothing else changed and the spill intact afterwards, the no-op re-pick, each refusal above, and the previous-selection ordering.

## 11. Visual design

| Cell role | Treatment |
|---|---|
| User input | White fill, blue left border, unlocked |
| Calculated | Light grey fill, italic, locked |
| Configuration | Pale blue fill, unlocked |
| Read-only reference | Light grey fill, grey text, locked |
| Snapshot / historical | Darker grey, collapsed group, locked |
| Warning state | Amber fill via conditional formatting, on every location sheet's Status column (`modInit.ApplyStatusFormat`) whenever a row reads anything other than `OK` |
| Error state | Red bold text via conditional formatting, on Summary's Type column (`modReports.FormatSummaryErrors`) whenever a job references a paper stock no longer in `tblPapers` (`"(not in Papers)"`) |

Transaction tables use freeze panes below the header, filter buttons, and banded rows. GBP formatting is driven by the global `SET_CURRENCY` setting (`modSettings.CurrencySymbol`) everywhere a money `NumberFormat` or `Format$` is built — see §16.2.

**A legend explaining the four everyday cell states lives on the Summary sheet**, at `O9` down (`modReports.DrawLegend`) — clear of the report table and the buttons `InitialiseWorkbook` draws at column O rows 1/3/5/7.

**Text wrapping** on the global settings table's Notes column is built (`modInit.FormatSettingsNotes`, snag item 16).

**Column groups** exist on every location sheet's cost and snapshot column ranges, and on the Reports page's filter/sort row block (§4.1, §8.3, snag items 14/15).

---

## 12. Multi-workbook aggregation — planned feature (D17)

**Automated collation across workbooks is still not built.** What has changed since the original design is that a **manual** route now exists and covers most of the same ground — see §12.4.

### 12.1 What v1 must preserve

| Invariant | Why it already holds |
|---|---|
| Job IDs globally unique — `UNI-MAIN-00001` | §3.2 |
| Site ID is a real setting | `SET_SITE_ID` |
| Every job row stamps `S_SchemaVer` | §5 |
| All records consolidated into one range | `_Data`, built for the local reports anyway |
| Location codes deterministic, never reassigned | §4.2 |
| A complete per-location export exists | §10.4 |
| **Job IDs come from a value that only ever increases**, never a row scan | §3.2, fixed 2026-09-21 — this specifically protects the "globally unique, never re-issued" half of the first invariant, which the original row-scan allocator could silently violate |

v1 does not have to *do* anything automated for aggregation; it has to avoid breaking these. That makes aggregator-readiness testable at acceptance rather than aspirational.

### 12.2 Deferred decision — O7

**Granularity**: does an automated master read job-level records, or each site's Summary?

- **Summary-level** aggregates by Location × Printer × Paper stock, discarding the student. A master built from summaries could never produce a cross-site student total — the entire purpose of Reports and AT-10.
- **Job-level** preserves it, and costs nothing extra since the export already exists.
- **Job-level with student number but no names** keeps cross-site totals working while narrowing what leaves each site.

**Still deferred deliberately on data-protection grounds** for any *automated* collation tool — a decision about lawful basis and data handling rather than about software. Recorded here so it is re-decided from these three options rather than re-derived. See §12.4 for why the built manual mechanism has, in effect, already made a job-level choice for the manual case without that choice needing to be finalised for an automated one.

### 12.3 The transport, when automated collation is built

Collation would read **the exported CSV files** (§10.4), not the source workbooks — no Excel automation of other people's files, no sandbox handling on Mac, and no reading a spill range out of a closed workbook, which does not reliably work at all.

### 12.4 The manual route that already exists ("BONUS" from the Problem 1 plan)

Export All Locations + Import (§10.4, §10.5) together let **any copy of this workbook serve as an ad-hoc master**: a blank copy imports the exports from every site's locations, and Summary/Reports then run over the merged data exactly as they would for a single site's own records, since both sheets are built on `_Data` regardless of how the rows arrived there.

This is import-by-hand of full job rows including student name and number — it is **not** an implementation of D17/O7's automated master workbook, and it does not resolve the O7 granularity decision for that future tool. It is recorded here because it is real, already-shipped functionality that happens to serve the same underlying need (GDPR-safer aggregation, since files move over whatever secure channel the organisation already uses rather than an automated process reaching into other people's workbooks), aggregation is expected to happen only a few times a year so the manual steps are acceptable, and any future automated master should be designed with this manual path already in mind rather than as an unrelated feature.

---

## 13. Build sequence and tooling

### 13.1 Delivery

A `.xlsm` cannot be assembled outside Excel: the VBA project is a binary structure no library on the build machine can author. Two artefacts and a script:

- `src\PrintCosts.xlsx` — everything a file can carry.
- `src\PrintCosts-VBA\*.bas`, `*.cls` — the complete VBA source, the authoritative copy.
- `build.ps1` — validates the source, backs up the existing `.xlsm`, copies the source to `%TEMP%`, imports every module into the copy, pastes `ThisWorkbook.cls` into the existing document module, runs `InitialiseWorkbook` then `StampBuild` in quiet mode, saves to a temp `.xlsm`, copies that into `src\`, and prunes old backups to the five most recent.

Requires **Trust access to the VBA project object model**, once. Windows only. The manual import route in `src\PrintCosts-VBA\SETUP.md` is the only route on Mac and must list all eighteen files — it was missing four (`modRegistry`, `modReports`, `modExport`, `modVersion`) until 0.7.1, which produced a project that would not compile on Mac. **`modImport.bas` makes the current total nineteen** — re-check SETUP.md's list stays complete whenever a module is added.

**A build must never mutate its own input.** The project lives in OneDrive, where AutoSave is on by default; opening the source `.xlsx` in place and modifying it (importing modules, running setup, stamping the version) lets AutoSave commit those changes back over the source before `SaveAs` ever runs, with `DisplayAlerts = $false` swallowing the warning that would otherwise stop it. This happened once and cost a 410KB `.xlsx` with an embedded VBA project that Excel refused to open at all, recovered from OneDrive version history. Four guards now stand between the build and a repeat: it never opens the source (works on a `%TEMP%` copy); it refuses to start if the `.xlsx` already contains a VBA project; it compares the source's timestamp before and after and warns loudly if it changed; `$wb.AutoSaveOn = $false` on the working copy.

**Excel also saves to `%TEMP%`, and the result is copied into place** — saving a few-hundred-KB workbook directly into an actively-syncing OneDrive folder was refused every time with an error that reads like a missing method rather than a contested destination.

**Backups are pruned to five.** They are build outputs, regenerable from the `.xlsx` plus the VBA source.

Two things the automation buys beyond convenience: a **compile check** (running a macro over COM forces the whole project to compile, so a syntax error fails the build rather than surfacing on a user's first click), and **safe testability** (quiet mode makes the test scripts non-destructive by construction, not by care).

**A compile error does not fail cleanly — it hangs.** Unlike a runtime error (caught by `InitialiseWorkbook`'s own `On Error`, logged quietly, build continues), a genuine VBA compile error is a blocking VBE dialog raised the moment the first `$xl.Run(...)` forces the project to compile — before any of the workbook's own error handling exists to catch it. `build.ps1` (and every `test-*.ps1`) just hangs, unattended, with no output and no COM exception to catch, because Excel is waiting on a modal dialog nothing in the automation can see or dismiss. Diagnosed by reading the dialog's Win32 window text directly (`EnumWindows`/`EnumChildWindows` against the `excel.exe` process — the "Compile error: ..." message lives in a child `Static` control) rather than trusting anything COM reports, since COM itself never gets a chance to return.

**A PowerShell `[char]` passed to an Excel COM indexer is not the same as a one-letter string, and the failure looks exactly like COM flakiness.** `$letter = [char](64 + $n); $ws.Columns($letter)` silently indexes by the `Char`'s **ordinal value** (82 for `'R'`), not the letter — a completely different, unrelated column, which then (correctly) reads back as not-hidden. This cost real time in testing (2026-09-22, `test-reports.ps1`'s Job ID column check) because the symptom is indistinguishable from genuine session-level COM unreliability: it "sometimes" fails depending on which column the wrong index happens to land on and what's there, so chasing it as a timing/settling issue looks entirely plausible right up until you compare `Columns($charVar)` against the literal `Columns('R')` side by side. The fix is `.ToString()` on the `Char` before handing it to any COM call: `([char](64+$n)).ToString()`.

**A caution about verifying over COM.** `$wb.BuiltinDocumentProperties('Title')` does not resolve as a parameterized COM property from PowerShell — it returns blank rather than raising, indistinguishable from a workbook whose properties were never written. Likewise Excel's `Names` collection reports hidden `_xlfn.*` placeholders it invents at run time and never saves. Both verification scripts now read the OOXML package directly. **When a check and the thing it checks disagree, suspect the check.**

### 13.2 Supporting scripts

| Script | Does |
|---|---|
| `probe.ps1` | Structure of a workbook, read-only. Whole workbook, or `-Sheet <name>` cell by cell. Flags defined names not stored in the file |
| `verify.ps1` | Opens the built `.xlsm` **read-only** and reports version, document properties, tables, registry, consolidated range and button bindings |
| `test-duplicate.ps1` | AT-13 — duplicates a print room in VBA, refreshes, reports |
| `test-reports.ps1` | Summary and Reports end to end: criteria cases, filters, sort, Export names, minimum-columns view, Summary buttons, Export report snapshot |
| `test-layout.ps1` | Shipped column order, outline/hidden state, header block through a reduced-view toggle |
| `test-manualhide.ps1` | Manual column hide/unhide: button relocation and side-panel following via the selection-change handler |
| `test-export.ps1` | Export, the CSV's contents, and the fingerprint's response to an edit |
| `test-validation.ps1` | AT-11, AT-14, AT-16 |
| `test-nextid.ps1` | The persisted Job ID high-water mark: deleting the top row must not reissue its ID; the mark must survive `RefreshLocations` |
| `test-catalogids.ps1` | Site-prefixed catalogue IDs: every named row has a unique well-formed ID; Add row and typed rows are allocated the next number and a deleted ID is never reissued; restoring another site's catalogue adds rows (renaming clashing names to `Name (SITE)`), overwrites nothing, is idempotent, and does not wind the counters back |
| `test-import.ps1` | Export All Locations, and Import restoring into origin and into a different room |
| `test-deletereports.ps1` | Export report (the static-value `.xlsx` snapshot) and the Reports-page bulk delete, including the audit log entry |
| `prune-backups.ps1` | Keeps the N most recent backups |

**Every test script drives a copy in `%TEMP%`, never `src\PrintCosts.xlsm` directly** — required because `Workbook_Open` does real work on every open (`ProtectAll`, `Invalidate`, `HealButtons`), and because this file is held through a **cloud-backed handle**: AutoSave commits those changes immediately, so `$wb.Saved` reads `True` on the line right after `Open` despite every sheet having just been modified, and the bytes land at `Close`/`Quit` regardless of `Close($false)` or a late `AutoSaveOn = $false`. This is not the sync client — it reproduces with syncing paused and settled. `probe.ps1` and `verify.ps1` sidestep it entirely by opening **read-only**.

**Do not kill Excel while it holds the workbook.** Observed once, not fully explained: an orphaned Excel process being killed was followed by the next open writing a stale cached session over a freshly rebuilt file, reverting it by several builds — suspected to involve the Office document cache. Let scripts quit Excel themselves.

Run the four `test-*.ps1` regression scripts **one at a time**, not in a tight loop — Excel can share one process between scripts, and a later script then reads an earlier one's in-memory state as though it were the build.

---

## 14. Open items

**Resolved:** O1–O6, O8 (delivery automated), O9 (export-based mitigation, §4.3/§10.4), O10 (closed — not a defect: `_xlfn.` prefixes in stored formulas are required, and Excel's runtime-only `_xlfn.*` placeholder names in its `Names` collection are never actually saved).

**Deferred:**

| # | Item |
|---|---|
| **O7** | Aggregation granularity for a future *automated* master — job-level, summary-level, or job-level without names. See §12.2–§12.4: a manual, job-level-with-names route already exists via Export All Locations + Import, which somewhat pre-empts the urgency of this decision but does not resolve it for automation |

**Tracked in §16 rather than here:** the open implementation debt (punch list in §16.2) is not given numbered "O" items, since it is not open design questions. The phase-8 visual-polish gaps were closed in 0.8.1 and the stale in-code version stamp in 0.8.0 (§16.1); their history is in `docs/HISTORY.md`.

---

## 15. Traceability

| Acceptance test | Design element | Verified |
|---|---|---|
| AT-01 | §5.1 Paper Cost formula; `S_UnitCost` for sheet stock | Phase 10 |
| AT-02 | §5.1 Area formula, roll branch, blank print width | Phase 10 |
| AT-03 | §5.1 — print width in Area only, never in Paper Cost | By construction |
| AT-04 | §7.3 width check | Phase 10 |
| AT-05 | §7.2 dependent list; §7.3 re-edit check | Phase 10 |
| AT-06 | §7.2 valid printer list from `LOC_Printers` | Phase 10 |
| AT-07 | §9 AddPrintJob seeds flags from location defaults | Phase 10 |
| AT-08 | §5 flags are plain row values with no link to the defaults | By construction |
| AT-09 | §6 snapshot block; no live link from job rows to config | By construction |
| AT-10 | §8.3 live filter — **within this workbook** (D17) | **Yes — `test-reports.ps1`** |
| AT-11 | §7.3 student consistency scan (D1, I5) — warns, never blocks | **Yes — `test-validation.ps1`** |
| AT-12 | §8.3 criteria seeded to array shape; blanks collapse to TRUE | **Yes — `test-reports.ps1`** |
| AT-13 | §4.2 marker detection, table parking and renaming, code assignment, formula rewrite | **Yes — `test-duplicate.ps1`** |
| AT-14 | §10.2 confirmation content; `modUtils.DateSerialOf` | **Yes — `test-validation.ps1`** |
| AT-15 | §9.3 cross-platform strategy; picker render check on Mac | Phase 10 |
| AT-16 | §3.2 stable IDs; §6.3 inactive-record behaviour | **Yes — `test-validation.ps1`** |
| — | §10.4 export completeness and fingerprint | **Yes — `test-export.ps1`** |
| — | §3.2 Job ID high-water mark survives deletion and refresh | **Yes — `test-nextid.ps1`** |
| — | §10.5 Export All Locations; Import (origin and cross-room) | **Yes — `test-import.ps1`** |
| — | §10.6 Export report snapshot; Reports-page bulk delete | **Yes — `test-deletereports.ps1`** |

Several acceptance tests are about what happens as a person types, which is worth testing as a person rather than only as a script, and AT-15 needs a Mac — these still need the phase 10 run. Export/Import/NextId/Reports-delete have no numbered acceptance test in the original spec, being design-led additions; they are covered by their own scripts instead.

---

## 16. Known gaps — read this before starting the next phase

### 16.1 The in-code version stamp — resolved 2026-09-21

`modVersion.APP_VERSION` had stayed at `"0.7.1"` through thirteen commits of real work (the NextId high-water-mark fix, Export All Locations, Export report snapshot, Import, the Reports rework, the Mac OneDrive export-folder fix, and the seventeen-item 2026-09-21 snag list), none of which were reflected in its changelog comment. **Bumped to `0.8.0`**, with the comment block above `APP_VERSION` rewritten to list everything in that batch, and a note that phase 8's original scope (§16.2) is explicitly *not* part of the bump. **Bumped again to `0.8.1`** once phase 8's original scope itself landed (§16.2) — a revision within phase 8, not a new phase digit, since this is that phase's own scope completing rather than new work being scoped in. This is a source-code change — it only reaches the workbook's About block and document properties the next time `build.ps1` runs.

### 16.2 Open items and history

Sections 16.2–16.5 (phase 8 scope, other loose ends, the phase 9 snag list, the printer/paper compatibility rework) were completed-work logs and now live in `docs/HISTORY.md`, numbering preserved. Still open from them:

**Punch list (updated 2026-10-02).** Carried forward:

- **Button arrangement** has had no overall pass across every sheet's buttons. 0.10.12 to 0.10.14 moved the column-view and cost-detail controls to row 2 and shrank the side panel to six, so this is partly eased, but the Settings page is still untouched (HISTORY §16.3).
- **Phase 10 manual run.** The acceptance tests in §15 marked "Phase 10" (AT-01 to AT-09, AT-15) have never been exercised by a person typing; AT-15 needs a Mac.
- **O7**, aggregation granularity for an automated master (§14). Still deferred.

Added in this pass:

- ~~**Docs drift**~~ **Closed 2026-10-02.** HISTORY §16.3's `ConvertQtyIfCentimetres` entry is marked superseded by 0.10.11, §14's stale phase-8 paragraph is rewritten, and SETUP.md is brought up to 0.10.x (twenty modules including `modBackup`, Add/Remove print room, column views, roll unit, Reports Paid editing, the current test scripts and `run-tests.ps1`, and the phase 10 status). SETUP.md should still be re-read whenever a user-facing feature ships.
- ~~**`ReorderJobColumns` validation corruption**~~ **Closed 2026-10-01.** `ReorderJobColumns` itself was deleted on 2026-09-29 (layout now ships in the .xlsx). What remained was stale hand-placed validation on rows 28-2010 *below* the table in the shipped template (e.g. a Yes/No list on what had become Sheet size), which `EnsureJobColumnValidation` never cleared. `modInit.ClearBelowTableValidation` now clears it once per sheet (setup and Refresh Locations; not from `BindColumns`, which runs on every row added and made Excel reject the next COM call), and every rule is bound by header name, so a future reorder/insert cannot misplace one. Guarded by `test-jobvalidation.ps1`, which also moves columns with the old Cut + Insert and checks one `BindColumns` repairs them.
- **Clean up `src\*.bak.xlsm`.** Five backups from 2026-10-01 are sitting beside `PrintJob.xlsm`; run `prune-backups.ps1` and consider having `build.ps1` prune on success.
- ~~**Column-view edge cases**~~ **Closed 2026-10-02**, no code change needed. The view is workbook-wide (one `SET_LOC_REDUCED_VIEW`), so the original wording was off: a new room starts in the *current* mode, not always All, and follows later changes with every other room. `test-viewedge.ps1` pins down: Refresh Locations and a full `InitialiseWorkbook` keep Minimal; Add print room in Minimal and in All; Import into a sheet with hidden columns fills them and leaves them hidden; a Settings list naming a nonexistent header, empty entries or a blank list fails soft (bad entries skipped, blank falls back to the defaults).
- **Reports date filters** have no calendar picker (platform limit, HISTORY §16.3). Revisit if a future Excel adds one.
- ~~**`LOC_RollUnit` is per-sheet, not per-row**~~ **Closed 2026-10-02, working as intended.** The roll length unit is a deliberate per-sheet display setting (0.10.11): Qty is held in the sheet's unit and converted to metres for `_Data`. A location is not meant to mix centimetre and metre entry job-by-job, so no per-row unit (and no schema bump) is planned.
- ~~**Mixed-unit import**~~ **Closed 2026-10-02**, no code change needed. `test-importunits.ps1` (Example Print Room on Metres, the Annexe fixture on Centimetres) covers a metres export into a cm room (roll Qty x100, Unit "cm"), a cm export into a metres room (/100, Unit "metres"), Area m2 and Paper Cost unchanged both ways, sheet stock never converted, a cm round trip landing back on the original metres, and a same-unit control.


---

## Sources

- [Worksheet.Protect method (Excel) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protect) — `UserInterfaceOnly` is not persisted across save/reopen
- [Excel for Mac: No ActiveX Support, OLE is Limited — SumProduct](https://sumproduct.com/blog/excel-for-mac-no-activex-support-on-mac-object-linking-and-embedding-ole-is-limited/)
- [Run-time error 429: ActiveX component can't create object on Mac — Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5407644/im-getting-a-run-time-error-429-activex-component)
- [Create dependent drop-downs with spill ranges — Excel University](https://www.excel-university.com/create-dependent-drop-downs-with-spill-ranges/)
- [Dynamic data validation lists with spill ranges — BrainBell](https://brainbell.com/excel/dynamic-list-with-spill-ranges.html)
- [VSTACK: combine multiple sheets with one formula — Xelplus](https://www.xelplus.com/excel-vstack-function/)
