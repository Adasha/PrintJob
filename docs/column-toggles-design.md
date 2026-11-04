# Column toggles on print room sheets — design notes

Status: design agreed 2026-11-03. Being built on `feat/column-toggles` as 0.12.0; tested on the Mac only. The plan below notes where the build departed from it.

Replaces the view row described in `column-view-presets-design.md`, including its
layered model (decision 8), and ARCHITECTURE §4.1's All / Reduced / Minimal text once built.

## Why

The row 2 controls are one drop-down and two buttons, and they confuse people:

- The controls overlap. The drop-down cannot touch the pass columns, and the cost toggle
  works on top of it, so "All" does not always mean all.
- Reduced hides `Status`, the column that reports job problems.
- The pass toggle mixes columns technicians need (`Passes`, `Colour`) with plumbing
  (`Row Type`, `Parent`, `Pass`).
- Everything is workbook-wide, though printers are permitted per room.

## Decisions

1. Row 2 becomes a "Show:" label and a line of six checkboxes. Ticked means shown. The
   drop-down and both toggle buttons go.
2. The checkboxes are per room. Clicking one changes only that sheet.
3. Internal columns are one workbook-wide Yes/No on the Settings sheet, not a checkbox.
   They are off by default.
4. The groups are fixed in code. The editable column lists on Settings go.
5. Colour passes and Printer have auto defaults. A click always beats the rule.
6. Technician is a plain toggle with no auto rule. An auto-hide rule tied to the default
   technician was considered and dropped as confusing (Adam, 2026-11-03). Revisit after
   usability testing.
7. Existing views reset to the new defaults once, on upgrade to this version. The old
   settings are not converted, because the old and new groups do not map cleanly. The
   changelog says so.

## Groups

| Checkbox | Columns | Default |
|---|---|---|
| Colour passes | Passes, Colour | Auto (below) |
| Cost detail | Paper Cost, Consumable Cost, Set-up Cost, Gross Cost, Disregarded, Area m2 | On |
| Disregards | Disregard Paper, Disregard Consumable | On |
| Print size | Print Width mm, Sheet size | On |
| Technician | Technician | On |
| Printer | Printer | Auto (below) |

- **Internal (Settings):** Job ID, Row Type, Parent, Pass, S_SchemaVer. Pass rows are read
  by their order under the job, so the `Pass` number does not need to be seen.
- **Always visible:** Date/Time, Student Name, Student No, Paper Stock, Unit, Qty,
  Chargeable Cost, Paid, Status, Notes.
- **Always hidden, unchanged:** the `H_` columns and the snapshot block, apart from
  S_SchemaVer (the last table column, which the template leaves visible).

Hiding only changes what is seen. Nothing on the location sheets reads visible cells only,
and reports and exports read values whatever the column state.

The room's default cells are in column B, so no toggle can hide them.

## Auto rules

- **Colour passes:** on if the room permits a multi-pass printer or already has pass rows.
- **Printer:** on unless the room permits exactly one active printer.

A checkbox follows its rule until someone clicks it. The click then sticks until the
rule's answer changes, at which point the click is dropped and the rule takes over again.
So adding a RISO to a room shows its pass columns even if someone had hidden them, and
there is no third checkbox state or reset button.

To do this without events on every path that can change a printer, a click on an auto
group stores the rule's answer at the time of the click, and the state is worked out
whenever the room's columns are applied. If the stored answer differs from the current
one, the click is discarded.

## Showing a column when it is needed

A job must not get an error about a column nobody can see. These put the column back and
tick its box, stored as a click:

| Event | Column shown |
|---|---|
| Add pass | Colour passes |
| New or Repeat job, default technician blank | Technician |
| New or Repeat job, printer not filled | Printer |
| Student-supplied stock chosen | Print size |

`BindColumns`' one-choice auto-fill already fills `Printer` on a new job when the room
permits exactly one active printer, even with no default set, so the auto rule's hidden
state never leaves it blank. The paper list already offers only compatible stock in that case, so a stock pick cannot clear
the printer. Job-planner jobs (`modPlanner`) follow the same rules as New job.

Hiding Printer by hand in a room with several printers remains possible. A stock pick
cannot blank the printer: when the two clash, `OnStockChanged` clears the stock instead.

`Parent job not found` and `Pass number out of sequence` refer to Internal columns. They
should not occur in normal use, so they stay as Status messages; the Settings switch shows
the columns for investigation.

## Storage

- **Per room:** one new name, `LOC_Columns`, in the location settings block (AX11, label in
  AW11), holding only what has been clicked. Example:
  `passes=off@on;cost=off`, where `@on` is the rule's answer at the time of the click.
  Empty means all defaults. Adding a later group needs no new cell.
- **Workbook:** `SET_SHOW_INTERNAL_COLS` (Yes/No, default No). Setup adds the row and name
  (`modColumns.EnsureColumnSettings`), as `EnsureSubsidySetting` does, so the template is
  not edited.
- **Retired:** `SET_LOC_REDUCED_VIEW`, `SET_LOC_REDUCED_COLUMNS`,
  `SET_LOC_MINIMAL_COLUMNS`, `SET_COST_COLS_HIDDEN`, `SET_MULTIPASS_COLS_HIDDEN`. Setup
  removes the rows and names; Restore skips them in older backups.
- **Schema:** `SCHEMA_VER` stays as it is, since it only tracks job-table columns.

## Layout

The controls still have to follow column changes. What changes is how much code does it.

Today `RepositionViewButtons` hunts for a visible column per control, and rewrites the
label into a cell on every run, which unprotects the sheet. The new row is one rigid strip
with fixed widths and gaps, like `RepositionToolbar` since 0.11.0. Only its start is worked
out: the left edge of column D, which stays correct even when Technician (D) is hidden, and
columns A to C cannot be hidden. The label is a shape rather than a cell, so repositioning
needs no cell write.

It runs at the same moments as now (activation, selection change, after a toggle). Excel
has no column-width event, so after a width change the strip catches up at the next
selection or activation, as the toolbar does.

The strip could reach the side panel if every group is off and the columns are narrow. The
layout test covers that case.

## Implementation plan

Announce each Excel batch first (CLAUDE.md). Do not edit the template or VBA, or build,
while another session's test run is in progress.

1. **Trial (Excel, on its own).** Through `xlmac.builder_helper`, trapped, on a throwaway
   sheet: a forms checkbox with `OnAction` on a protected sheet, `Application.Caller`
   naming it, and setting `.Value` from code after `UnlockSheet`. Existing checkboxes are
   only read when the picker's OK is pressed, so none of this has been exercised. Adam
   then clicks one by hand, because tests cannot click. If it fails, stop and rethink the
   control before going further.
2. **Settings rows.** Done in code at setup rather than in the template (see Storage).
3. **Groups and state** (`modInit`, `modPasses`). Group definitions, `LOC_Columns` read and
   write, rule evaluation, and one `ApplyRoomColumns ws` replacing `ApplyReducedView`,
   `ApplyCostColumnsVisibility` and `ApplyMultiPassVisibility`. It changes a column only if
   its state differs, as it will run on activation. It keeps using header names through
   `ApplyColumnVisibility` and never does a blanket reset.
4. **Row 2.** Replace `DrawViewDropDown`, the two `DrawOne` calls and
   `RepositionViewButtons` with the strip. Replace `ddViewMode`, `btnToggleCostColumns` and
   `btnToggleMultiPass` with one checkbox macro in `modMain`. Remove `btnToggleMultiPass`
   from the greying in `modPasses`. Old controls must be deleted from existing sheets on
   redraw. The pass buttons keep their current greying.
5. **Reveals and auto-fill** (`modJobs`, `modPasses`, `modValidation`, `modPlanner`), per
   the table above.
6. **Settings switch.** In `Workbook_SheetChange`'s Settings case, re-apply every room when
   that cell changes, and on `Refresh Locations` as now (`modRegistry`).
7. **Rooms and backup.** Add `LOC_Columns` to both lists in `modBackup` (`RoomSettingNames`
   and `ApplyRoomSettings`), to the reset in `CreatePrintRoom` (`modRegistry`), and to
   `EnsureLocName` for existing rooms.
8. **Older backups.** `ApplyCatalogRows` would have appended the five retired rows from a
   pre-0.12.0 backup as dead rows, so it now skips them (`IsRetiredSetting`).
9. **Docs and release.** `APP_VERSION` to 0.12.0 (the middle number has so far tracked
   build phases; see `modVersion`), changelog entry including the one-off reset, SETUP.md's
   row 2 paragraph and test table, ARCHITECTURE §4.1 and the §17 multi-pass button text, the
   wiki plan's status, and `mac/docs-screens.py`.

## Tests

A changed test gets both versions, `.ps1` and `mac/*.py`, with the same checks and
messages. Windows has not been verified since 0.11.1, so results will be Mac-only until a
Windows run happens.

Affected: `test-reducedview`, `test-costcolumns`, `test-multipass-views`, `test-viewedge`,
`test-layout`, `test-settingsrows`, `test-paid`. The first two are probably best replaced
by one new `test-columntoggles`.

New checks:

- Each checkbox hides and shows exactly its columns, on that room only.
- `Status`, `Qty` and the other always-visible columns survive every combination.
- Both auto rules, including pass rows in a room with no multi-pass printer.
- A click sticks, then is dropped when the rule's answer changes.
- Each reveal in the table above, with its box ticked afterwards.
- Single-printer auto-fill on New job.
- The Internal setting applies to every room and leaves the `H_` columns hidden.
- `LOC_Columns` carries over when restoring from a copy of the workbook (a CSV backup
  holds no room settings, so a room made from one starts at the defaults), and resets on
  Add print room.
- Layout with every group off: the strip is clear of the side panel.

## Manual checks

Tests can call macros but cannot click or see, so Adam checks by hand on the Mac:

- Clicking each checkbox on a protected sheet, and how quickly it responds.
- That the strip reads well, and the state of the row after hiding columns.
- Usability testing with technicians, including whether Technician should gain an auto
  rule (decision 6) and whether Area m2 belongs under Cost detail.

## Out of scope

- Per-person views. Views are per room, so two technicians in one room share a view.
- Greying the pass buttons per room. `MultiPassPrinterExists` stays workbook-wide.
