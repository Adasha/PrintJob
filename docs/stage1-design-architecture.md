# University Printing Cost Management Workbook
## Stage 1 — Architecture & Design Document

**Status:** Phases 1–7 built and verified, plus a 0.7.1 review pass. Phases 8–9 outstanding.
**Covers:** spec §18 Stage 1 deliverables — workbook architecture, data model, entity relationships, calculation rules, location-sheet architecture, VBA architecture, cross-platform control strategy, data-validation strategy, historical-data strategy.
**Version:** 1.8 — 20 September 2026
**Workbook version at time of writing:** 0.7.1, data schema 1.0

> **Changes in 1.8 — review pass before phase 8 (0.7.1).** A read of the whole source against this document, before building on it. **No change to the data schema, and no new behaviour.** Six code fixes and two documentation ones:
>
> - **`modReports.SheetNamed` moved a sheet later using `Move Before:`, which is a no-op.** Every build shipped Cost Calculations on tab 1 and Summary on tab 2 — the reverse of what the code asks for — because the `.xlsx` ships neither report sheet, so `build.ps1` always takes the first-run path. Moving later needs `After:`.
> - **`Val()` on numeric cells replaced with `modUtils.NumOf`** (§8.2). `Val` takes a `String`, so a number round-trips through the machine’s decimal separator and is parsed back expecting a US point: on any comma-decimal locale every paper cost, consumable rate and stock dimension reads as **zero**, silently. §8.2 stated the rule for date cells because that is where it was first caught; the cause is the locale round-trip, not the date. Six call sites in `modCatalog` and `modValidation`. AT-15 runs on a Mac, which is where this would have surfaced.
> - **`modSnapshot.LogAudit` and `modExport.ExportSig` no longer wrap a batch of independent operations in one `On Error Resume Next`** (§8.2, the `StampProperties` rule). `LogAudit` could leave `_Audit` unprotected and a truncated entry indistinguishable from a complete one; `ExportSig` could drop a fingerprint term and report *no changes since* for a location that had changed.
> - **`modRegistry.WriteRegistry` no longer raises on a workbook with no print rooms** — `DataBodyRange` is `Nothing` once every registry row is deleted, which is exactly the §3.3 case where Refresh Locations most needs to finish and name what went missing.
> - **`modSettings.Money` made locale-safe**, and given a `CurrencySymbol` helper that reads `SET_CURRENCY`. Both are still unreferenced: the symbol is hardcoded as `ChrW(163)` in twelve places, and §10’s “GBP formatting driven by the global setting” is **not yet true**. Carried into phase 8.
> - **`SETUP.md` listed fourteen modules of eighteen**, omitting `modRegistry`, `modReports`, `modExport` and `modVersion` — the phase 4–7 additions. The manual route is the only route on Mac, so following it produced a project that would not compile. This blocked AT-15.
> - §4.1 wrote `SET_RoundDP`; the defined name is `SET_ROUND_DP`.
>
> **Two further defects, found by the 0.7.1 regression run rather than by reading.**
>
> - **The export wrote nothing and reported success.** `ThisWorkbook.Path` is not reliably a filesystem path: for a OneDrive-backed workbook Excel may report the service URL, and this project lives in OneDrive. `ExportPath` concatenated that with `Application.PathSeparator`, `SaveAs` did not raise, and the export announced *“5 print jobs exported”* having written nothing the user could find — then stamped the location as exported. The §3.3 warning that stands between a sheet deletion and the records said the opposite of the truth. `modExport` now resolves the URL back to the local OneDrive root, refuses with an explanation when it cannot, and **verifies the file exists before stamping anything**. That last check is the durable part: the status must never be optimistic.
> - **The test scripts were rewriting the workbook they tested.** `Workbook_Open` does real work on every open — `ProtectAll` unprotects and reprotects all fourteen sheets, plus `Invalidate` and `HealButtons` — so the workbook is dirty the instant it opens. Excel holds the file through its **cloud-backed handle** (`Workbook.Path` returns a `d.docs.live.net` URL), so AutoSave is on and **commits that immediately**: `$wb.Saved` reads **True** on the next line, despite every sheet having just been modified. The bytes land when the handle closes — timestamp and SHA-256 both move at `Close`/`Quit`, not at `Open`. That is precisely why `Close($false)` cannot help: it discards *unsaved* changes, and by then there are none, so what is written at close is committed state. `AutoSaveOn = False` after `Open` is too late for the same reason, and `Open` offers no earlier hook — it fails *by construction* rather than by not taking effect. **This is not the sync client**: it reproduces with syncing paused and settled, and it is not §12.1 repeating. §8.2’s own rule applies — when a check and the thing it checks disagree, suspect the check; here the check was `Saved`, and it was telling the truth about a save that had already happened. `verify.ps1` has documented the first half since phase 5, which is why it and `probe.ps1` open **read-only**. The four `test-*` scripts need read-write to drive mutating macros, so they now drive a copy in `%TEMP%` — `build.ps1`’s §12.1 rule applied to the tests — as does `diag.ps1`. `test-export.ps1` also stops deleting the real exports in `src\` to find its own.
>
> **Regression run.** All four scripts pass against the 0.7.1 build, with figures identical to the phase 5–7 baselines, and each verified by SHA-256 to have left `src\PrintCosts.xlsm` byte-identical.
>
> Also confirmed against the built file rather than assumed: the §4.1 formulas, the §7.1 consolidated formula, the 21 + 12 column layout, the marker cell, and §10’s freeze panes and filter buttons. What §10 still lacks is the Summary legend, the warning and error conditional formats, and currency from the setting — which is phase 8’s scope. One item is deferred to phase 8 rather than fixed here: `modExport.ExportLocation` never sets `Application.DisplayAlerts`, and every test script sets it `$false` on the Application before driving the workbook, so Excel’s CSV prompt cannot appear under test and will appear for a user.
>
> **Changes in 1.7 — phase 7 swept.** AT-11, AT-14 and AT-16 verified by `test-validation.ps1` (§14), and **no defect was found**: the mechanisms were built in earlier phases and this confirmed them. AT-14's outstanding re-test after the `Val()` date fix is now done. One note added to §2.5 about a version increment that carries no behaviour change, and backup retention added to §12.1.
>
> **Changes in 1.6 — phase 6 built.** Per-location export verified (§9.4): a CSV carrying every column including the snapshot block, a derived fingerprint, and a `LOC_Export` status surfaced on each sheet and by *Check workbook* (§6.3) and *Refresh Locations* (§3.2). One new convention in §8.2, from a defect the test caught.
>
> **Changes in 1.5 — phase 5 built.** Summary and Cost Calculations verified (§7.2, §7.3). `modReports` is a real module, but only as a **builder** (§8.1). Four conventions added to §8.2. §12.1 records the build's near-loss of the source `.xlsx` to OneDrive AutoSave.
>
> **Changes in 1.4 — scope settled.** Multi-workbook aggregation became a **planned feature, out of scope for v1** (§11, D17); O7 deferred on GDPR grounds. **O9 resolved** by export-based mitigation (§9.4, D18). **O10 closed — not a defect.** Derived-over-tracked recorded as a standing principle (§8.2).
>
> **Changes in 1.3:** version identity — two independent version numbers (§2.5, D16).
>
> **Changes in 1.2 — the document catching up with the build.** Delivery automated (§12.1), closing O8. Three divergences recorded as built: consolidated range on `_Data`; location detection by marker *cell*; dependent dropdowns as staging columns.
>
> **Changes in 1.1:** Microsoft 365 only, permitting dynamic arrays; reports rebuilt on a consolidated range; multi-site hooks added; O1–O6 resolved.

---

## 1. Decisions taken at clarification

| # | Question | Decision |
|---|---|---|
| D1 | Student name/number consistency (§10.2, AT-11) with no master list | No Students register. On entry, cross-check the pair against all existing rows in all locations; warn (never block) on conflict |
| D2 | Mechanism for historical cost integrity (§12.6) | Snapshot the resolved rates onto each row at creation; visible costs remain live formulas over those snapshot values |
| D3 | Contents of "global settings" (§5.1) | Currency format + rounding rule; organisation/report header metadata; site identity. No VAT, no overhead markup |
| D4 | Per-row *Now* and *Remove* controls (§10.1, §10.12) | Toolbar of Form Control buttons above each table, acting on the selected row; plus double-click to stamp date/time |
| D5 | How reports discover location sheets after duplication (§17) | Marker cell on each location sheet; a **Refresh Locations** command rescans and rebuilds a hidden registry |
| D6 | Papers "Cost" means £/sheet or £/metre (§8.4) | One `Cost` column, with an adjacent calculated `Cost unit` label driven by the stock's family |
| D7 | Units for transaction print width (§10.6) | Millimetres, matching stock width |
| D8 | Presentation of disregarded costs (§12.5, §13) | Gross, Disregarded and Chargeable as separate columns; consumption unaffected throughout |
| D9 | Custom sheet sizes (§8.2) | Standard A-sizes remain a picker; explicit custom width × height in mm also permitted |
| D10 | Multi-select for printer families and location printers (§7.1, §9.1) | Worksheet-based picker with a multi-select list, writing a delimited string to a single cell |
| D11 | Deployment / file-sharing model | Single workbook is the product. Multi-workbook aggregation is a planned feature — D17, §11 |
| D12 | Minimum Excel version | **Microsoft 365 builds only**, Windows and Mac. Dynamic array formulas may be relied upon |
| D13 | Report construction | Reports are live formulas over a consolidated spill range, not VBA-generated output (§7) |
| D14 | Multi-site collation mechanism | Superseded by D17/D18: collation, when built, reads the **exported CSV files** (§11.2) |
| D15 | `.xlsm` delivery (was O8) | A PowerShell build script drives Excel over COM. Manual import route retained for Mac (§12.1) |
| D16 | Version identity | **Two independent version numbers** — a workbook version for the build, and a data schema version for the shape of the records (§2.5) |
| D17 | Scope of v1 | **Multi-workbook aggregation is a planned feature, not in v1.** v1 delivers a single workbook complete in itself, preserving the invariants an aggregator needs (§11.0). O7 deferred with it |
| D18 | Deleted location sheets (was O9) | **Export-based mitigation** rather than a shadow copy: per-location CSV, a visible per-sheet status, surfaced by *Check workbook* and *Refresh Locations* (§9.4, §3.3) |

### 1.1 Interpretations — confirmed at sign-off

- **I1 — Sheet-stock consumable area scales with quantity.** §10.9 says "for sheet stock, printed area = full sheet area", without reference to quantity; taken literally, fifty A1 prints would cost the same in consumables as one. **Confirmed:** area is `sheet area × number of sheets`.
- **I2 — Paper families carry a measurement basis.** The spec treats the family list as user-extensible (§5.4) while deriving sheet-versus-roll behaviour from it (§12.1). **Confirmed:** each family carries an explicit **Measurement basis** of `Sheet` or `Roll`, and behaviour keys off that attribute, never off the family's name.
- **I3 — Sheet-stock quantities are whole numbers.** Roll quantities are decimal metres; sheet quantities integers ≥ 1. **Confirmed.**
- **I4 — Rounding is applied per cost component at row level**, so a printed report's rows add up to its totals. **Confirmed.**
- **I5 — AT-11 is a warning, not a block.** With no authoritative student list the first entry of any student is unverifiable, and legitimate name changes would trigger it. **Confirmed**, and verified in phase 7.

---

## 2. Data model

### 2.1 Entities and relationships

```
Settings (global key/value)
PaperType ──┐
            ├──< PaperStock >──── PaperFamily ──< PrinterFamily >── Printer ──── ConsumableType
StandardSize┘         │                                               │
                      │                                               │
                      └───────────┐                   ┌───────────────┘
                                  │                   │
Technician ──────────────────────>PrintJob<───────────┘
                                  │
                              Location (one sheet each)
                                  │
                              Site (one workbook each)
```

- A **Printer** supports one-or-more **PaperFamily** (delimited list on the printer row).
- A **PaperStock** belongs to exactly one **PaperFamily** and one **PaperType**.
- A **PrintJob** references one Technician, one Printer and one PaperStock, and belongs to one Location.
- A stock is valid for a printer only when the stock's family appears in that printer's supported families.
- A **Location** permits a subset of printers.
- A **Site** is one workbook file. In v1 there is one site per workbook and nothing collates them.

### 2.2 Identity and referential integrity

Every configuration record carries a **stable ID** — `TEC-001`, `PRN-001`, `STK-001` — assigned once and **never re-issued**, even after deletion. Job rows store the ID in their snapshot block, so renaming a printer or technician cannot orphan or silently re-point a historical record. IDs come from the highest ever issued, never the row count.

The visible cell on a job row holds the **display name**, because that is what a technician recognises in a dropdown. Configuration names are therefore unique within their table, enforced on entry.

**Job IDs are `<SITE>-<LOC>-00001`.** The site component exists so records from several workbook copies can never be conflated. It costs nothing in v1 and is one of the invariants aggregation depends on (§11.0).

Because a location's code is baked into its job IDs, **an existing code is never reassigned** (§3.2).

> **Integrity rule:** configuration rows must be deactivated, not deleted (§16). Deleting a row history references is blocked with a message naming the number of dependent records.

### 2.3 Configuration tables

| Table | Columns |
|---|---|
| `tblSettings` | Key, Setting, Value, Notes |
| `tblPaperTypes` | Paper type, Active |
| `tblStandardSizes` | Size name, Width mm, Height mm |
| `tblPaperFamilies` | Family, **Measurement basis** (`Sheet`/`Roll`), Notes |
| `tblConsumables` | Consumable type, Active |

Each setting is exposed as a workbook-scoped defined name so formulas and VBA reference meaning rather than cell addresses: `SET_SITE_ID`, `SET_SITE_NAME`, `SET_ORG`, `SET_DEPT`, `SET_CURRENCY`, `SET_ROUND_DP`, `SET_FOOTER`, `SET_FY_START`, plus read-only `SET_SCHEMA`, `SET_LASTREF`, `SET_APP_VER`, `SET_BUILT`, `SET_BUILT_BY`.

`modSettings` addresses every setting as `SET_<KEY>`, so the defined name — not the row — is what makes an entry a real setting.

**Print Technicians** — `tblTechnicians`: TechID, Name, Department, Active.
**Printers** — `tblPrinters`: PrinterID, Model, Consumable type, **Cost per m2**, Supported families, Active.
**Papers** — `tblPapers`: StockID, Description, Paper type, Family, *Measure* (calc), Size mode, Standard size, **Width mm**, **Height mm**, **Cost**, *Cost unit* (calc), Active.

All dimensions are stored in millimetres; metres appear only as a transaction quantity for roll stock.

### 2.4 Hidden working sheets

`xlSheetVeryHidden`, so they cannot be unhidden from the tab context menu:

| Sheet | Purpose |
|---|---|
| `_Registry` | `tblLocations`: SheetName, Code, Name, Department, Rows, First date, Last date, State, **Last export**, **Export sig** |
| `_Work` | **Validation staging only** — §6.2 |
| `_Data` | The consolidated job range (§7.1) |
| `_Audit` | `tblAudit`: When, User, Action, Location, Detail (§9.3) |
| `_Export` | Reserved. The export writes to disk (§9.4) |
| `_Picker` | Backing sheet for the multi-select picker (§8.3) |

**Why `_Data` is separate from `_Work`.** `modLists` claims `_Work` one column at a time, tagging each in row 1 and scanning up to 200 for a free slot, and will overwrite anything else parked there.

**Why the export stamp lives in the registry.** `Last export` and `Export sig` (§9.4) are keyed by sheet name, not stored in the location's config block. A sheet-scoped cell would copy with a duplicated sheet, and the copy would inherit "already exported" while holding records never exported anywhere. A registry row means a duplicate gets a fresh row and correctly reads as unexported; a renamed sheet likewise. Both are the conservative direction.

The shipped `.xlsx` predates the export feature, so those two columns are added to the existing table by `EnsureColumn` rather than being part of its creation.

### 2.5 Version identity (D16)

| | Meaning | Changes when | Held in |
|---|---|---|---|
| **Workbook version** | This build — code, layout, formulas | Every rebuild worth distinguishing | `modVersion.APP_VERSION` → `SET_APP_VER` |
| **Data schema version** | The shape of the stored records | Only on a genuine shape change | `modUtils.SCHEMA_VER` → `SET_SCHEMA`, stamped on every job row as `S_SchemaVer` |

The schema version is load-bearing: it is what an importer or aggregator checks before reading records (§9.4, §11), and what a job row carries so its origin is recoverable. Conflating the two would force a schema bump on every cosmetic change and make that check meaningless.

**What bumps the schema version.** Adding, removing or renaming a column bumps it. **Reordering columns does not** — exports carry a header row and are read by column *name*. The case the number cannot protect against is a column whose *meaning* changes while its name stays the same; that is a discipline, not a mechanism.

**Numbering is `0.<phase>.<revision>`**, tracking §12: `0.7.x` means phases 1–7 are built. It reaches `1.0.0` when acceptance passes.

*A note on 0.7.0 specifically.* Phase 7 changed nothing in the workbook's behaviour — it swept the validation against the acceptance tests and found no defects — so that build differs from 0.6.0 only in its stamp. The increment still earns its place: the scheme records which phases a file has been through, and whether a workbook's validation has been swept is a fact about the artefact even when it is not a fact about its code.

**Where the version is visible**, in ascending order of effort: the file's **document properties** (Explorer, Finder — without opening the file); the **About block** on Settings, which reads the settings cells by formula so it cannot drift; and the **About button**.

**What stamps what.** `EnsureVersionSettings` writes `APP_VER` from the constant on every setup run, because the code in the workbook *is* its version. `BUILT` and `BUILT_BY` are written only by `StampBuild`, called by the build script **after** `InitialiseWorkbook` — users re-run setup whenever they add a print room, and stamping from there would reset the build date to whenever someone last duplicated a sheet.

---

## 3. Location-sheet architecture

### 3.1 Layout

```
Rows 1–8    Configuration block
Row 10      Toolbar (Form Control buttons)
Row 12      Table header
Row 13+     tblJobs_<CODE> body
```

| Cell | Name | Content |
|---|---|---|
| B1 | `LOC_Name` | Room name |
| B2 | `LOC_Dept` | Department |
| B3 | `LOC_Code` | Location code (auto-assigned, read-only) |
| B4 | `LOC_DefDisPaper` | Default disregard paper cost (Yes/No) |
| B5 | `LOC_DefDisCons` | Default disregard consumable cost (Yes/No) |
| B6 | `LOC_Printers` | Permitted printers, delimited, read-only |
| B7 | `LOC_Status` | Validation summary for the sheet |
| B8 | `LOC_Export` | Export status (§9.4) |
| AZ1 | — | Marker cell, `PRINTLOC/v1`. Hidden column, never edited |

`LOC_Export` is deliberately separate from `LOC_Status`: validation and export state are independent, and a sheet can easily be valid and unexported at once. Both are derived and rewritten, never typed. `LOC_Export`'s name is created by `modExport` on every refresh rather than shipped in the `.xlsx`, so a duplicated or renamed sheet gets a correct one without hand surgery.

**Why sheet-scoped names matter:** Excel copies them with the worksheet. Workbook-scoped names would collide on copy and silently become `LOC_Name1`, `LOC_Name2`.

**The marker is the exception, and is a plain cell.** It copies with a duplicated sheet exactly as a scoped name does, Excel cannot quietly rename it, and it is one moving part instead of two.

### 3.2 Detection and registration (D5)

A sheet is a location if and only if `AZ1` reads `PRINTLOC/v1`. **Refresh Locations** walks every worksheet, tests the marker, and rebuilds `_Registry.tblLocations`. In order:

1. **Parks every job table on a temporary name.** Excel auto-renames a duplicated `ListObject` — in testing, `tblJobs_ANNEX` became `tblJobs_ANNEX14`. Renaming straight to the final name fails whenever that name is still held by a table the pass has not reached, which is exactly the situation after a duplication.
2. **Assigns a location code** where `LOC_Code` is blank or duplicated. Sheets are visited in workbook order and Excel inserts a duplicate immediately after its original, so the original is reached first and keeps its code — required, because its job IDs already carry it.
3. **Renames each job table** to `tblJobs_` + its code (AT-13).
4. **Re-points Form Control buttons** whose `OnAction` acquired a workbook-file prefix on copy.
5. **Rebuilds the dependent dropdowns** and **re-applies protection**.
6. **Rewrites the consolidated range formula** (§7.1).
7. **Recomputes each location's export status** and writes `LOC_Export` — after the registry is written, because the status is read out of it (§9.4).
8. **Stamps `SET_LASTREF`.**

Any location with unexported changes is named in the refresh message.

**Button bindings need healing on open, too.** Excel re-qualifies every button's `OnAction` with the workbook's file name when it saves, so the shipped file reads `PrintCosts.xlsm!btnAddPrintJob` — which breaks if the file is renamed. `Workbook_Open` calls `HealButtons`, which rewrites only on a mismatch, so merely opening an untouched workbook never marks it dirty and prompts a save.

### 3.3 Deleted locations (D18, was O9)

**Excel provides no `BeforeDeleteSheet` event.** Deleting a location sheet cannot be intercepted, undone or audited. What changes is whether the records were somewhere else first.

- **A recovery path.** Each location exports to CSV holding every column, including the snapshot block (§9.4).
- **A visible warning where the user is.** `LOC_Export` states how many records have changed since the last export, and *Check workbook* and *Refresh Locations* surface the same line per location.
- **Self-healing afterwards.** The consolidated formula is wrapped in `IFERROR`, so a deleted sheet degrades to an empty result rather than cascading `#REF!`. Registry entries are compared against the sheets present on each refresh, and anything vanished is named.

**This is a sign on the cliff, not a fence**, and the design says so rather than implying otherwise. The rejected alternative was a mirrored shadow copy of every location's rows — a permanent cost on every write, for an edge case a working export covers better and which has independent value.

### 3.4 Documented procedure for adding a location

1. Right-click any location tab → *Move or Copy* → tick *Create a copy*.
2. Rename the new tab.
3. Click **Clear All** to discard the copied records.
4. Enter room name, department, defaults; click *Select printers…*.
5. Click **Refresh Locations**.

---

## 4. Transaction table

| # | Column | Type | Notes |
|---|---|---|---|
| 1 | Status | calc | `OK` or a warning; conditionally formatted |
| 2 | Job ID | VBA | `<SITE>-<LOC>-00001`, never re-used |
| 3 | Date/Time | input | Required. Toolbar *Now*, or double-click |
| 4 | Student Name | input | At least one of name/number required |
| 5 | Student No | input | " |
| 6 | Technician | input | Dropdown, active only |
| 7 | Printer | input | Dropdown, active ∩ permitted here |
| 8 | Paper Stock | input | Dropdown, active ∩ compatible with printer |
| 9 | Unit | calc | `sheets` or `metres` |
| 10 | Quantity | input | > 0; whole number for sheet stock (I3) |
| 11 | Print Width mm | input | Roll only; blank = full stock width; ≤ stock width |
| 12 | Disregard Paper | input | Yes/No, seeded from location default |
| 13 | Disregard Consumable | input | Yes/No, seeded from location default |
| 14 | Area m2 | calc | Printed area |
| 15–19 | Paper Cost, Consumable Cost, Gross Cost, Disregarded, Chargeable Cost | calc | §4.1 |
| 20 | Notes | input | Optional |
| 21 | H_Issues | calc | Hidden working column behind Status |

**Snapshot block (22–33)** — locked, grey, collapsed group headed *Historical record — do not edit*: `S_PrinterID`, `S_StockID`, `S_TechID`, `S_Family`, `S_Measure`, `S_UnitCost`, `S_StockWidth_mm`, `S_SheetHeight_mm`, `S_ConsRate`, `S_StampedAt`, `S_StampedBy`, `S_SchemaVer`.

Column *order* is not load-bearing anywhere: formulas use structured references, VBA resolves columns by header name, reports resolve `_Data` by header name (§8.2), and exports carry a header row read by name (§9.4). Reordering is free and does not bump the schema version.

### 4.1 Formulas

```
Unit             =IF([@S_Measure]="Sheet","sheets","metres")

Area m2          =IF([@Quantity]="","",
                   LET(w, IF([@S_Measure]="Sheet",
                             [@S_StockWidth_mm],
                             IF([@[Print Width mm]]="", [@S_StockWidth_mm], [@[Print Width mm]])),
                       h, IF([@S_Measure]="Sheet", [@S_SheetHeight_mm]/1000, [@Quantity]),
                       n, IF([@S_Measure]="Sheet", [@Quantity], 1),
                       (w/1000) * h * n))

Paper Cost       =IF([@Quantity]="","",ROUND([@Quantity]*[@S_UnitCost],SET_ROUND_DP))
Consumable Cost  =IF([@[Area m2]]="","",ROUND([@[Area m2]]*[@S_ConsRate],SET_ROUND_DP))
Gross Cost       =IF([@[Paper Cost]]="","",[@[Paper Cost]]+[@[Consumable Cost]])
Chargeable Cost  =IF([@[Paper Cost]]="","",
                    IF([@[Disregard Paper]]="Yes",0,[@[Paper Cost]])
                  + IF([@[Disregard Consumable]]="Yes",0,[@[Consumable Cost]]))
Disregarded      =IF([@[Gross Cost]]="","",[@[Gross Cost]]-[@[Chargeable Cost]])
```

The formulas read **only snapshot columns and the row's own inputs**. No formula on a job row reaches into `tblPapers` or `tblPrinters` — the property that satisfies §12.6 and AT-09, because a price change has no live link to propagate along.

`Print Width` appears in the area formula and nowhere in paper cost: AT-02 and AT-03 made structural rather than procedural.

---

## 5. Historical-data strategy (D2)

### 5.1 Stamping

When **Add Print Job** creates a row, and whenever printer or paper stock changes, `modSnapshot.StampRow` resolves the current configuration and writes the twelve snapshot values. Costs are computed by the row's own formulas from those values — *live formulas over frozen rates*, so the arithmetic stays auditable and a mis-typed quantity recalculates immediately.

### 5.2 Re-stamping

A **Re-stamp prices** command exists for the genuine correction case. Deliberately not on the location toolbar: it lives on Settings, confirms while naming the affected rows, and writes an `_Audit` entry.

### 5.3 What survives configuration change

| Change | Effect on existing rows |
|---|---|
| Paper cost edited | None — from `S_UnitCost` (AT-09) |
| Consumable £/m² edited | None — from `S_ConsRate` |
| Printer renamed | None — row holds `S_PrinterID` |
| Printer families changed | None — compatibility validated at entry (AT-05) |
| Stock dimensions changed | None — area from the stamped dimensions |
| **Record set Inactive** | **None — remains in reports; excluded only from new dropdowns (AT-16, verified phase 7)** |
| Location defaults changed | None — flags copied into the row at creation (AT-07, AT-08) |
| **Global rounding changed** | **Recalculates existing rows — accepted (O1)**, and labelled as such on Settings |

---

## 6. Data-validation strategy

### 6.1 Layer 1 — Data Validation

Static lists, numeric constraints, date type constraints. Enforced before VBA sees anything.

### 6.2 Layer 2 — Dependent lists on staged ranges

Version 1.1 specified dropdowns on spill references. Two hard limits ruled that out:

- **A Data Validation rule is uniform down a table column**, but each row's candidate stocks depend on *that row's* printer. One spill cannot serve every row, and the workaround — a `SelectionChange` handler writing the active row's printer into a context cell — makes the dropdown depend on the selection. Replacing a cell's validation closes an open dropdown, so the arrow visibly flashed and vanished.
- **A validation list supplied as a literal string is capped at 255 characters**, which a real stock list exceeds.

`modLists` gives **every distinct list its own staging column** on `_Work`, claimed by a tag in row 1 — `PRN|<sheet>`, `TEC`, `STK|<model>`. Consequences:

- Lists rebuild when something **changes**, never on selection, so nothing flashes.
- Rows are grouped by printer before binding, so each distinct list is written once.
- A per-list column was necessary: one shared column per list *type* meant the last sheet bound overwrote the others, and one print room showed another's printers.
- **Paper Stock is locked while Printer is blank**, making the incompatible combination unreachable rather than merely caught.

### 6.3 Layer 3 — Row and workbook validation

`Worksheet_Change` runs `OnCellChanged` for cross-field rules validation lists cannot express: print width > stock width (AT-04); print width against sheet stock; non-integer sheet quantity (I3); neither student name nor number; name/number conflict against history (D1, AT-11); printer/stock combination made incompatible by a later edit (AT-05).

Because a row can be invalidated by an edit elsewhere, *Status* persists the verdict, and **Check workbook** sweeps every location and reports all outstanding problems in one list.

**Check workbook also reports export state, counted separately.** Unexported records are not a defect in the data — they are a risk *to* it, because sheet deletion cannot be intercepted (§3.3). Folding them into the problem count would make a perfectly valid workbook report problems, and a count that cries wolf gets ignored.

Messages state what is wrong, why, and what to do — §15's three-part requirement, enforced by `modUtils.Say` taking the three parts as separate arguments so a caller cannot quietly omit the third.

---

## 7. Reports (D13)

### 7.1 The consolidated range

`Refresh Locations` writes one formula to `_Data!A10`:

```
=LET(raw,
  VSTACK(
    HSTACK(IF(SEQUENCE(ROWS(tblJobs_MAIN[Job ID]))>0,"MAIN"),  tblJobs_MAIN[[Job ID]:[Notes]]),
    HSTACK(IF(SEQUENCE(ROWS(tblJobs_ANNEX[Job ID]))>0,"ANNEX"), tblJobs_ANNEX[[Job ID]:[Notes]])),
  IFERROR(FILTER(raw, INDEX(raw,,2)<>""), ""))
```

- **`IF(SEQUENCE(ROWS(…))>0,"CODE")` rather than a bare `"CODE"`**: `HSTACK` does not broadcast a scalar against a column.
- **`FILTER(raw, INDEX(raw,,2)<>"")` drops blank table rows.**
- **`IFERROR` on the outside** is the §3.3 safety net.

Shape: `Location`, then `Job ID` through `Notes` — 20 columns. Headers are written to row 9 from the first location's actual header row, so they cannot drift.

VBA's entire role in reporting is rewriting that one formula. Everything downstream is a live worksheet formula: reports are never stale, and roughly half of what would have been a reporting engine does not exist and therefore cannot carry a defect into acceptance.

### 7.2 Summary sheet (§13) — **built**

One row per **Location × Paper stock**, from a single `LET` deriving key pairs with `SORT(UNIQUE(HSTACK(...)))` and aggregating with `COUNTIFS`/`SUMIFS` taking the key columns as **array criteria** — which makes the results spill alongside the keys instead of needing one formula per row.

| Location | Paper stock | Type | Family | Unit | Jobs | Quantity | Area m² | Paper cost | Consumable cost | Gross | Disregarded | Chargeable |

Type, Family and Unit are resolved from `tblPapers` by stock description, wrapped in `IFNA` so a stock renamed or removed since shows `(not in Papers)` rather than an error. These are labels only — no cost figure is ever looked up live (§4.1).

Consumption columns count **every** record regardless of disregard flags (§12.5); only money columns split (D8).

**Totals sit above the table, not beneath it.** The detail spills to an unpredictable height, so anything below it is overwritten the moment a job is added.

### 7.3 Cost Calculations sheet (§14) — **built**

A live `FILTER` driven by four criteria cells; results update as criteria are typed.

Student number matches exactly after trimming; name case-insensitively as a "contains" match. Because the column stores date *and time*, the end-date test is `< end + 1` — otherwise a job logged at 16:30 on the closing date falls outside its own range.

Output: the full record list, plus breakdowns by location and by paper stock, and totals for the current selection. **Breakdowns sit to the right of the record list**, for the same reason totals sit above the Summary table. They cannot use `SUMIFS` — its arguments must be ranges, and after `FILTER` these are arrays — so per-key aggregation is `BYROW` with a `LAMBDA`.

§14.1's disjoint-criteria case is a warning above the results when a name and number are both given and never appear together.

**This is the whole of AT-10 for v1.** A student's total across every print *room* in this workbook is in scope and served here; only the cross-*workbook* case is deferred (D17).

**Verified** by `test-reports.ps1`: all four blank-criteria combinations, single-day ranges including times, partial and case-insensitive names, the §14.1 warning, and a date left as text. Breakdowns reconcile exactly to the Summary totals.

### 7.4 Where the commands live

| Sheet | Buttons |
|---|---|
| Summary | Refresh Locations, Check workbook, Go to Settings |
| Settings | Refresh Locations, Check workbook, Re-stamp prices…, About |
| Each location | Add Print Job, Now, Remove Row, Select printers…, Check this sheet, Clear All, Export… |
| Printers | Select families… |

Settings keeps its own copies deliberately: it is where someone lands when configuring, and *Re-stamp prices* and *About* belong nowhere else.

---

## 8. VBA architecture

### 8.1 Modules

| Module | Responsibility | State |
|---|---|---|
| `modMain` | Public entry points bound to buttons. Thin wrappers; names are a stable contract with the Form Controls | Built |
| `modJobs` | AddPrintJob, StampNow, RemoveRow, ClearAll | Built |
| `modSnapshot` | StampRow, ReStampAll, LogAudit | Built |
| `modValidation` | OnCellChanged, CheckSheet, CheckWorkbook, student consistency scan | Built |
| `modLists` | Dependent dropdowns and their staging columns (§6.2) | Built |
| `modRegistry` | RefreshLocations, EnsureSystemSheets, code assignment, table renaming, button healing, consolidated-range formula | Built |
| `modReports` | **Builds** the Summary and Cost Calculations sheets — layout and formulas, once (§7.2, §7.3) | Built |
| `modExport` | Per-location CSV, the export fingerprint, the `LOC_Export` status line (§9.4) | Built |
| `modVersion` | Version constants, settings rows, About block, document properties (§2.5) | Built |
| `modPicker` | The multi-select picker | Built |
| `modInit` | One-time setup: draws the Form Controls, applies protection, hands over to RefreshLocations | Built |
| `modCatalog` | Configuration loaded once per operation | Built |
| `modSettings` | Typed accessors for settings and named ranges | Built |
| `modProtect` | Protect/unprotect wrappers, re-applied on open | Built |
| `modUtils` | Application state, messaging, quiet mode, table and name access, ID generation | Built |

| Class | Responsibility | State |
|---|---|---|
| `clsDict` | Keyed collection — §8.3 | Built |
| `clsStock` / `clsPrinterDef` | One paper stock / one printer | Built |

**`modReports` exists, but 1.1's claim holds in substance.** It writes layout and formulas once and computes nothing at run time. The distinction matters: a defect in `modReports` can only produce a wrong *sheet*, visible immediately, not a wrong *figure* on a sheet that looks right.

`clsCatalog` and `clsLocation` were not needed. `modCatalog` holds configuration in module-level `clsDict` instances with an `Invalidate` flag; location sheets are addressed through `modUtils` helpers.

### 8.2 Conventions

- `Option Explicit` in every module.
- Every public entry point: disable events/screen updating → `On Error GoTo Fail` → work → restore state → exit. `AppOff`/`AppOn` are depth-counted so nesting is safe.
- **`EnableEvents = False` during our own writes is not negotiable.** Handlers write cells, and writing a cell re-enters `Worksheet_Change`. `AddPrintJob` writes five cells on a row that does not exist yet; with events live, validation would judge a half-built row, emit spurious Status warnings, and could clear a legitimate Print Width. `ClearAll` and `ReStampAll` would run a validation pass per cell write, and manual calculation means such validation reads pre-edit values anyway. Anything needing to know that data changed must **derive** it, not listen for it.
- **Derive state; do not track it.** Applied four times: `IsLocation` reads a marker cell (§3.1); the registry rescans (§3.2); the export fingerprint is computed on demand (§9.4); report figures are formulas rather than stored results (§7). Derived state cannot rot, survives a user editing with macros disabled, and survives an old copy of a sheet being dropped back in — all three defeat a tracked flag silently.
- **Reports address `_Data` columns by header name**: `INDEX(_Data!$A$10#,,MATCH("Quantity",_Data!$A$9:$AZ$9,0))`. Wordier than `INDEX(...,,10)`, and the reason the job table can be reordered without touching a report formula.
- **A criteria expression must be an *array*, not a scalar.** `IF($C$5="",TRUE,…)` returns a bare `TRUE` when the box is empty, and `FILTER(column, TRUE)` is `#CALC!`. With all four criteria blank the product is the scalar `1`, so Cost Calculations failed in its most ordinary state. Every criteria product is seeded with a column-shaped term (`Job ID <> ""`) that fixes its height.
- **Coerce both sides of a date comparison** with `*1`. A user typing `16/09/2026` into an unformatted cell leaves text behind, and `number >= text` is FALSE for every row without raising anything — a filter that silently returns nothing.
- **A staging sheet used to reach Excel's CSV writer must be formatted as Text.** The export formats dates as ISO and the schema version as `1.0`, then writes them into cells; without `NumberFormat = "@"` Excel re-parses both — the date becomes a date value again and is written out in the machine's locale (`9/15/2026` here), and `1.0` becomes the number `1`, so an importer checking for `1.0` sees `1`.
- **Totals go above a spill; secondary tables go beside it.** A spill's height is unknowable, so anything below it is displaced the moment the data grows.
- Formulas written from VBA use `.Formula2`, never `.Formula`. `.Formula2` is array-aware; `.Formula` applies implicit intersection and silently stores a single value where a spill was intended.
- No hard-coded row or column numbers: columns are resolved by header name.
- **Never `Val()` on a date cell.** `Val` takes a `String`, so a date is coerced to text and its leading digits read: `15/09/2026 10:24` comes back as `15`, a date in January 1900. This was a live defect — `ClearAll`'s confirmation, the text AT-14 checks, reported 1900 dates. `modUtils.DateSerialOf` reads `.Value2` and is the only sanctioned route to a date serial.
- **`On Error Resume Next` around a block, never around a batch of independent writes.** `StampProperties` wrapped four document-property assignments in one blanket handler; when they appeared to fail there was no way to tell which or why, and the investigation went the wrong way for two rebuilds (§12.1).
- **Quiet mode.** `modUtils.SetQuiet` switches `Say` from `MsgBox` to collecting messages for `QuietLog`, and makes `Ask` return **False**. A script driving the workbook over COM has nobody to dismiss a dialog, and one `MsgBox` hangs the run indefinitely. `Ask` returning False is deliberate: an unattended run must never confirm a destructive operation on the user's behalf — and it has a second use, since it lets a test read a destructive command's confirmation text while guaranteeing the command aborts (§12.1).

### 8.3 Cross-platform strategy

Version is no longer a constraint (D12); **platform still is**. Excel for Mac has no ActiveX and no COM automation ([SumProduct](https://sumproduct.com/blog/excel-for-mac-no-activex-support-on-mac-object-linking-and-embedding-ole-is-limited/); [Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5407644/im-getting-a-run-time-error-429-activex-component)).

| Unavailable on Mac | Approach taken |
|---|---|
| `Scripting.Dictionary` | `clsDict`, built on `Collection` with error-trapped lookup |
| `Scripting.FileSystemObject` | Not used |
| `ADODB.Stream` | Not used — CSV written via `SaveAs xlCSVUTF8` (§9.4) |
| ActiveX controls | Form Controls only |
| `Application.FileDialog` | Export writes beside the workbook using `ThisWorkbook.Path` (§9.4) |
| UserForm rendering differences | The picker is a **worksheet** pretending to be a dialog (`_Picker`) |
| COM automation | The build script is Windows-only; Mac uses the manual import route |

- **Protection must be re-applied on open.** `UserInterfaceOnly:=True` is not persisted; a workbook saved in that state reopens fully protected and VBA can no longer write to its own sheets ([Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protect)).
- **1900 date system enforced** on open, since an older Mac workbook may carry 1904 and shift every date by four years.

### 8.4 Protection model

| Element | State |
|---|---|
| Input cells | Unlocked |
| Formula, snapshot and header cells | Locked |
| Sheets | `UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True, DrawingObjects:=False` |
| Hidden sheets | `xlSheetVeryHidden` |
| Workbook structure | **Unprotected** — required so users can duplicate location sheets (§17) |
| VBA project | Locked for viewing |

Sheet password is `printlog`. Structural operations route through `modProtect` wrappers that unprotect, act and reprotect within a single error-guarded call, because adding rows to a ListObject on a protected sheet is unreliable even with `UserInterfaceOnly`.

---

## 9. Destructive operations and export

### 9.1 Remove row (§10.12)

Acts on the selected row. Confirms with the job's identifying detail — ID, date, student, printer, cost — never a bare "Are you sure?". Writes an `_Audit` entry, then deletes the row.

### 9.2 Clear All (§11, AT-14) — **verified**

Scoped to the current location sheet. States the location name, record count and date range:

> *Clear All — Main Print Room*
> *This will permanently delete 5 print jobs dated 15/09/2026 to 17/09/2026.*
> *Records on other print room sheets are not affected.*
> *This cannot be undone. Continue?*

The date range depends on `DateSerialOf` (§8.2) — the `Val()` defect that made this text report 1900 dates is fixed, and `test-validation.ps1` re-tests it. Where the location has unexported changes, the confirmation says so.

### 9.3 Audit note

VBA operations clear Excel's undo stack. `_Audit` records what was removed, when and by whom. Sheet deletion is the one destructive act that cannot be caught (§3.3).

### 9.4 Export (D18) — **built**

**Purpose**, three at once: a recovery path for a deleted location sheet (§3.3), a way to get records out for finance, and the transport a future aggregator reads (§11.2).

**Content — the one irreversible decision.** Every column, **including the snapshot block**. An export without `S_UnitCost`, `S_ConsRate` and the rest is permanently unimportable: re-importing would recost every row at today's prices, destroying exactly the integrity §12.6 and AT-09 protect.

`H_Issues` is the single exclusion — an internal working cell behind the Status message, meaningless outside the workbook and recomputed on import. Calculated columns *are* included: they cost nothing and make the file readable by a person rather than only by a machine.

**Format.** CSV, written by copying to a temporary workbook and using `SaveAs FileFormat:=xlCSVUTF8`. Excel's own CSV writer, so quoting and comma escaping are its problem, and `xlCSVUTF8` gets `£` and any non-ASCII student name right. The staging sheet is formatted as Text throughout — see §8.2, which this feature is the reason for.

A header block — schema version, site ID, location code and name, generated-at, row count — then a header row, then records. Dates are written `yyyy-mm-dd hh:nn:ss`: a serial would be unreadable and a locale-formatted date ambiguous, since `06/07` is two different days depending on who opens it.

**Column order is not part of the contract.** The header row names the columns and any importer maps by name, so the job table may be reordered freely without invalidating older exports or bumping the schema version (§2.5). The file uses a canonical order fixed in `modExport`.

**Import is explicitly a later revision.** v1 exports only; the format is designed so import stays possible, which costs nothing now and cannot be retrofitted.

**Naming.** `PrintCosts-<SITE>-<LOC>-yyyymmdd-hhmm.csv`, beside the workbook via `ThisWorkbook.Path`.

**Knowing what is unexported — derived, not tracked.** A per-location fingerprint computed on demand:

```
row count + CountA(data range) + Sum(Quantity) + Sum(Chargeable Cost) + Max(Date/Time)
```

compared against `tblLocations.Export sig` from the last export. Nothing hooks any event and no mutator must remember anything — which matters because `EnableEvents` is off during our own writes (§8.2), because a user can type into a sheet with macros disabled, and because restoring an old copy of a sheet brings a stale flag with it.

*The limitation, plainly:* a collision needs an edit leaving count, cell count, both sums and the latest date unchanged — two compensating changes in one location between exports. Constructible deliberately, vanishingly unlikely by accident.

**Verified** by `test-export.ps1`: status transitions from *never exported* → *no changes since* → *CHANGED* after a single quantity edit; the CSV carries all 32 columns with every snapshot column present; dates and schema version survive as written.

**Data protection.** Export multiplies the places student names and numbers live, in files outside the workbook's protection. Not an objection, but the reason the destination is beside the workbook rather than a chooser, and a point for whoever writes the operating procedure. It connects directly to the deferred O7 (§11.1).

---

## 10. Visual design (§3)

| Cell role | Treatment |
|---|---|
| User input | White fill, blue left border, unlocked |
| Calculated | Light grey fill, italic, locked |
| Configuration | Pale blue fill, unlocked |
| Read-only reference | Light grey fill, grey text, locked |
| Snapshot / historical | Darker grey, collapsed group, locked |
| Warning state | Amber fill via conditional formatting |
| Error state | Red text via conditional formatting |

A legend on the Summary sheet explains the four everyday states (phase 8). Transaction tables use freeze panes below the header, filter buttons, banded rows, and GBP formatting driven by the global setting.

---

## 11. Multi-workbook aggregation — planned feature (D17)

**Not in v1. Nothing here is built, and v1 does not depend on any of it.**

### 11.0 What v1 must preserve

| Invariant | Why it already holds |
|---|---|
| Job IDs globally unique — `UNI-MAIN-00001` | §2.2 |
| Site ID is a real setting | `SET_SITE_ID` |
| Every job row stamps `S_SchemaVer` | §4 |
| All records consolidated into one range | `_Data`, built for the local reports anyway |
| Location codes deterministic, never reassigned | §3.2 |
| A complete per-location export exists | §9.4 |

v1 does not have to *do* anything for aggregation; it has to avoid breaking these. That makes aggregator-readiness testable at acceptance rather than aspirational.

### 11.1 Deferred decision — O7

**Granularity**: does a master read job-level records, or each site's Summary?

- **Summary-level** aggregates by Location × Paper stock, discarding the student. A master built from summaries could never produce a cross-site student total — the entire purpose of §14 and AT-10.
- **Job-level** preserves it, and costs nothing extra since the export already exists.
- **Job-level with student number but no names** keeps cross-site totals working while narrowing what leaves each site.

**Deferred deliberately on data-protection grounds** — a decision about lawful basis and data handling rather than about software. Recorded here so it is re-decided from these three options rather than re-derived.

### 11.2 The transport, when it is built

Collation reads **the exported CSV files** (§9.4), not the source workbooks. This supersedes D14 and is simpler three ways: no Excel automation of other people's files, no `GrantAccessToMultipleFiles` sandbox handling on Mac, and no reading a spill range out of a closed workbook — which does not reliably work at all.

### 11.3 The master workbook

A separate file: a **Sources** sheet listing folders or files, a **Collate** command stacking their records into a consolidated range identical in shape to §7.1, then the *same* Summary and Cost Calculations formulas over the merged range. Failures are per-source and non-fatal.

---

## 12. Build sequence

| Phase | Content | Verifies | State |
|---|---|---|---|
| 1 | Settings, Technicians, Printers, Papers; configuration tables, IDs, validation | — | Built |
| 2 | Location sheet: layout, table, formulas, snapshot block, protection | AT-01, AT-02, AT-03 | Built |
| 3 | `modJobs`, `modSnapshot` — add/stamp/remove; dependent lists | AT-04…AT-09 | Built |
| 4 | `modRegistry` — marker, refresh, table renaming, consolidated range, duplication | AT-13 | Built and verified |
| 5 | Summary and Cost Calculations | AT-10, AT-12 | Built and verified |
| 6 | `modExport` — per-location CSV, export fingerprint, `LOC_Export` status | D18 / §9.4 | Built and verified |
| 7 | Validation sweep against the acceptance tests | AT-11, AT-14, AT-16 | **Swept — no defects** |
| 8 | Visual polish, legend, user notes | — | **Next** |
| 9 | Acceptance testing on Windows and Mac; test report | AT-15, all | Outstanding |
| — | Multi-workbook aggregation | — | Planned, post-v1 (§11) |

**Phase 7 changed no code.** Its three tests passed on the first run, which is the outcome the earlier phases were built for rather than a surprise: the mechanisms — `Say`'s three-part messages, `Ask`'s detailed confirmations, the student scan, the `Active` flag's separation of dropdowns from reports — were all in place by phase 6. What phase 7 adds is evidence, and a regression test that will notice if any of it is broken later.

### 12.1 Delivery (D15, was O8)

A `.xlsm` cannot be assembled outside Excel: the VBA project is a binary structure no library on the build machine can author. Two artefacts and a script:

- `src\PrintCosts.xlsx` — everything a file can carry.
- `src\PrintCosts-VBA\*.bas`, `*.cls` — the complete VBA source, the authoritative copy.
- `build.ps1` — validates the source, backs up the existing `.xlsm`, copies the source to `%TEMP%`, imports every module into the copy, pastes `ThisWorkbook.cls` into the existing document module, runs `InitialiseWorkbook` then `StampBuild` in quiet mode, saves to a temp `.xlsm`, copies that into `src\`, and prunes old backups to the five most recent.

Requires **Trust access to the VBA project object model**, once. Windows only.

**A build must never mutate its own input — and this one did.** The project lives in OneDrive, where AutoSave is on by default. The build opened the source `.xlsx` in place and necessarily left it heavily modified — modules imported, setup run, version stamped — between `Open` and `SaveAs`. AutoSave wrote all of that back over the source, and `DisplayAlerts = $false` swallowed the "you can't save macros in a macro-free workbook" prompt that would have stopped it. The result was a 410KB `.xlsx` containing a VBA project, which Excel then refused to open at all. Recovered from OneDrive version history.

Four guards now stand between the build and a repeat:

1. **It never opens the source.** It works on a copy in `%TEMP%`, so AutoSave has nothing of the project's within reach. The other three are safety nets; this is the fix.
2. **It refuses to start** if the `.xlsx` contains a VBA project, naming the problem rather than failing later with Excel's "the file format or file extension is not valid".
3. **It compares the source's timestamp before and after** and warns loudly if the build changed it.
4. `$wb.AutoSaveOn = $false` on the working copy.

**Excel also saves to `%TEMP%`, and the result is copied into place.** Saving a 400KB workbook directly into an actively-syncing OneDrive folder was refused every time, and the error — *"Unable to get the SaveAs property of the Workbook class"* — reads like a missing method rather than a contested destination, which sent the investigation after Excel's state instead. A plain file copy has none of that trouble.

**Backups are pruned to five.** They are build outputs, regenerable from the `.xlsx` plus the VBA source, so a deep history earns nothing and syncs hundreds of megabytes over time.

Two things the automation buys beyond convenience:

- **A compile check.** Running a macro over COM forces the whole project to compile, so a syntax error fails the build rather than surfacing on a user's first click.
- **Testability, and safely.** The test scripts drive the workbook and close without saving. More than that, quiet mode makes them **non-destructive by construction**: with `Ask` returning False, `test-validation.ps1` reads Clear All's confirmation text in full while the deletion it guards cannot happen.

**A caution about verifying over COM.** `$wb.BuiltinDocumentProperties('Title')` does not resolve as a parameterized COM property from PowerShell: it returns a blank rather than raising, indistinguishable from a workbook whose properties were never written — and cost two rebuilds chasing a defect that did not exist. Likewise Excel's `Names` collection reports hidden `_xlfn.*` placeholders it invents at run time and never saves, which produced the false O10. Both verification scripts now read the OOXML package directly. **When a check and the thing it checks disagree, suspect the check.**

**Supporting scripts**

| Script | Does |
|---|---|
| `probe.ps1` | Structure of a workbook, read-only. Whole workbook, or `-Sheet <name>` cell by cell. Flags defined names not stored in the file |
| `verify.ps1` | Opens the built `.xlsm` **read-only** and reports version, document properties, tables, registry, consolidated range, button bindings |
| `test-duplicate.ps1` | AT-13 — duplicates a print room in VBA, refreshes, reports |
| `test-reports.ps1` | AT-10 and AT-12, and the Summary/breakdown figures |
| `test-export.ps1` | Export, the CSV's contents, and the fingerprint's response to an edit |
| `test-validation.ps1` | AT-11, AT-14, AT-16 |
| `filecheck.ps1` | Integrity of the `.xlsx`, `.xlsm` and every backup |
| `lockcheck.ps1` | Who is holding the `.xlsm` |
| `xlfnscan.ps1` | Every `_xlfn.` / `_xlws.` occurrence in the package |
| `prune-backups.ps1` | Keeps the N most recent backups |

---

## 13. Open items

Resolved in 1.1: **O1**–**O6**. Resolved in 1.2: **O8** (delivery automated).

Resolved in 1.4:

- **O9 — resolved** by export-based mitigation rather than a shadow copy (D18, §9.4, §3.3); built and verified as of 1.6.
- **O10 — closed, not a defect.** `_xlfn.IFERROR` is not in the file. Excel synthesises a hidden `=#NAME?` placeholder for every `_xlfn.`-prefixed token it meets and never saves them. The prefixes in stored formulas are **required** — `_xlfn.LET`, `_xlfn.VSTACK`, `_xlfn.HSTACK`, `_xlfn.SEQUENCE`, `_xlfn._xlws.FILTER` — and stripping them would produce the `#NAME?` the item feared. The genuine residual risk is that `_Data` shows `#NAME?` on any Excel without dynamic arrays, which D12 places out of scope: a deployment constraint to enforce, not a defect to fix.

**Deferred:**

| # | Item |
|---|---|
| **O7** | Aggregation granularity — job-level, summary-level, or job-level without names. Deferred with the feature on data-protection grounds; three options in §11.1 |

---

## 14. Traceability

| Acceptance test | Design element | Verified |
|---|---|---|
| AT-01 | §4.1 Paper Cost formula; `S_UnitCost` for sheet families | Phase 9 |
| AT-02 | §4.1 Area formula, roll branch, blank print width | Phase 9 |
| AT-03 | §4.1 — print width in Area only, never in Paper Cost | By construction |
| AT-04 | §6.3 width check | Phase 9 |
| AT-05 | §6.2 dependent list; §6.3 re-edit check | Phase 9 |
| AT-06 | §6.2 valid printer list from `LOC_Printers` | Phase 9 |
| AT-07 | §8 AddPrintJob seeds flags from location defaults | Phase 9 |
| AT-08 | §4 flags are plain row values with no link to the defaults | By construction |
| AT-09 | §5 snapshot block; no live link from job rows to config | By construction |
| AT-10 | §7.3 live filter — **within this workbook** (D17) | **Yes — `test-reports.ps1`** |
| AT-11 | §6.3 student consistency scan (D1, I5) — warns, never blocks | **Yes — `test-validation.ps1`** |
| AT-12 | §7.3 criteria seeded to array shape; blanks collapse to TRUE | **Yes — `test-reports.ps1`** |
| AT-13 | §3.2 marker detection, table parking and renaming, code assignment, formula rewrite | **Yes — `test-duplicate.ps1`** |
| AT-14 | §9.2 confirmation content; §8.2 `DateSerialOf` | **Yes — `test-validation.ps1`**, including the date-range re-test |
| AT-15 | §8.3 cross-platform strategy; picker render check on Mac | Phase 9 |
| AT-16 | §2.2 stable IDs; §5.3 inactive-record behaviour | **Yes — `test-validation.ps1`** |
| — | §9.4 export completeness and fingerprint | **Yes — `test-export.ps1`** |

Five of the sixteen acceptance tests are verified by script, three hold by construction, and the rest need the phase 9 run — several of them (AT-01, AT-02, AT-04…AT-07) because they are about what happens as a person types, which is worth testing as a person rather than only as a script. AT-15 needs a Mac.

Export has no acceptance test in the spec, being design-led. The end-to-end case — export, delete the sheet, confirm the records are recoverable from the file — belongs in phase 9.

---

## Sources

- [Worksheet.Protect method (Excel) — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protect) — `UserInterfaceOnly` is not persisted across save/reopen
- [Excel for Mac: No ActiveX Support, OLE is Limited — SumProduct](https://sumproduct.com/blog/excel-for-mac-no-activex-support-on-mac-object-linking-and-embedding-ole-is-limited/)
- [Run-time error 429: ActiveX component can't create object on Mac — Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5407644/im-getting-a-run-time-error-429-activex-component)
- [Create dependent drop-downs with spill ranges — Excel University](https://www.excel-university.com/create-dependent-drop-downs-with-spill-ranges/)
- [Dynamic data validation lists with spill ranges — BrainBell](https://brainbell.com/excel/dynamic-list-with-spill-ranges.html)
- [VSTACK: combine multiple sheets with one formula — Xelplus](https://www.xelplus.com/excel-vstack-function/)