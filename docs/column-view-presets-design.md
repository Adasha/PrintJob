# Column view presets — design notes (companion to multipass-costing-design.md, decision 14)

Status: design agreed 2026-10-06 (all questions answered). Nothing implemented.

## Current behaviour (read from src/PrintCosts-VBA/modInit.bas, master)
- View row on each location sheet: "Show columns:" label, an All / Reduced / Minimal drop-down (forms control, works on Mac), and a Hide/Show cost detail button.
- Reduced hides the columns listed in setting LOC_REDUCED_COLUMNS (default "Status;Job ID;Area m2;S_SchemaVer"). Minimal hides those PLUS LOC_MINIMAL_COLUMNS (default "Printer;Disregard Paper;Disregard Consumable;Print Width mm;Sheet size"). Current mode is stored in setting LOC_REDUCED_VIEW.
- Location-sheet buttons are forms-control Buttons (`ws.Buttons.Add`, modInit.DrawOne), free-floating. `ThisWorkbook.Workbook_SheetSelectionChange` already runs on every selection change on a location sheet (calls OnSelection, RepositionLocationButtons, RefreshStatusNotes), so a selection-driven button state update has a place to hook in.

## Decisions
1. (Adam, 2026-10-06) The multi-pass columns (colour and Passes) can be toggled as a GROUP, with their own independent toggle like the cost-columns button, rather than via extra presets. The hidden helper columns stay hidden regardless.
2. (Adam, 2026-10-06, "keep it with the other cost columns") Set-up Cost stays in the existing cost-columns group and is shown/hidden by the cost toggle only, NOT by the multi-pass toggle.
3. Earlier (main doc decision 5): colour column hidden by Reduced for now.
4. Earlier (main doc decision 14): prefer independent toggles over a long list of presets; presets must cover single and multi-pass use.
5. (Adam, 2026-10-06) The multi-pass toggle is on ALL location sheets for consistency. If no multi-pass printer is set up, it is displayed in a lighter shade as a visual hint. Adam then said "yes shade everything": the pass buttons (Add pass, Remove pass, Toggle passes, Toggle all passes) are shaded too when no multi-pass printer exists.
6. (Adam, 2026-10-06) Add pass / Remove pass self-disable and re-enable depending on the selected row: clickable only when a multi-pass job row (or one of its pass rows) is selected. Greyed via caption font colour (see test), plus a guard inside the macro.
7. (Adam, 2026-10-06, "give a short message") Clicking a greyed button shows a short message (e.g. "Select a multi-pass job first") rather than doing nothing. Exact wording not yet set.
8. (Adam, 2026-10-06, "keep it layered - the drop down is a 'convenience' level declutterer, while the toggles target specific workflows") Layered model: the All/Reduced/Minimal drop-down keeps its own hidden-column lists (Settings); the group toggles (cost columns, multi-pass columns) act on top. A column is hidden if either the view or its toggle hides it. The colour/Passes columns are in the Reduced list (decision 3).
   - "All" with a toggle that is off (Adam, 2026-10-06, "yes go with a"): option A — "All" shows everything EXCEPT groups that are toggled off; the toggles always win. Choosing "All" does NOT switch the groups back on.
9. (Adam, 2026-10-06, "maybe we can look at the terminology used to make it clearer later on") The wording of the view drop-down and toggles (All/Reduced/Minimal, "convenience" level vs toggles) may be reworked later to make the layering clearer. Not in scope now; keep labels easy to change.

## Test results (Windows, Excel 16.0.20430, throwaway workbook, deleted afterwards)
- `Button.Enabled = False` can be set and re-set from code, including on a sheet protected with UserInterfaceOnly (ProtectDrawingObjects True). It is NOT visually greyed: screenshot shows an Enabled=False forms button looking identical to a normal one. Whether it also blocks clicks was NOT tested (no click test possible from this session).
- `Button.Characters().Font.Color = grey` DOES grey the caption (screenshot confirmed) on an unprotected sheet, but FAILS on a protected sheet ("You cannot use this command on a protected sheet"). So a visual grey-out needs unprotect, set colour, re-protect (existing UnlockSheet pattern) — acceptable cost per selection change only when the state actually changes.
- Not tested: Mac behaviour of Button.Enabled / font colour; real click behaviour of Enabled=False; flicker/perf of unprotect/re-protect inside SelectionChange.
- Design implied: grey caption via font colour (visual) + the macro itself guards and shows the short message if clicked in the wrong state (safe even if Enabled does not block clicks).
