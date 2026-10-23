# Feature-to-page map

Every sheet, button, control and setting in the workbook, with the wiki page
that documents it and how it behaves on each platform. A feature with no row
here is undocumented. Wiki contributors and reviewers use this file; it is not
a page for workbook users, and the publishing Action should skip it.

Inventoried from the 0.11.2 build (`src/PrintJob.xlsm`, built from 5dd31d2) on
2026-10-23 by reading the file's sheets, defined names and drawing parts, then
following each button's macro in `modMain.bas` to the routine that does the
work.

**Platform column.** `Same` = checked on both, no difference. `Mac:` or
`Windows:` = a known difference, described. `Check` = not yet looked at by a
person (the phase 1 Mac checks below). Nothing has been run on Windows since
0.11.1, so no row is `Same` yet.

**Tests column.** The Mac test (`mac/test-<name>.py`) and its Windows twin
(`test-<name>.ps1`) that cover the feature. Blank = no automated test.

## Sheets

| Sheet | Visibility | Purpose | Page |
| --- | --- | --- | --- |
| Summary | Visible | Totals by room, printer and paper; the job planner; By department | Reports-Summary |
| Reports | Visible | Filtered list of every job, Paid tracking, report export | Reports-Filtering, Reports-Export, Jobs-Paid-Tracking |
| Example Print Room | Visible | One sheet per print room; the sample room ships with 12 jobs | Jobs-Recording, Setup-Print-Rooms |
| Print Technicians | Visible | `tblTechnicians` | Setup-Print-Technicians |
| Printers | Visible | `tblPrinters` | Setup-Printers |
| Papers | Visible | `tblPapers` | Setup-Papers |
| Consumables | Visible | `tblColours`, the colours multi-pass printers use | Setup-Consumables |
| Departments | Visible | `tblDepartments`, free departments and aliases | Setup-Departments |
| Settings | Visible | Global settings, lookup lists, print room and data commands | Setup-Settings, Setup-Lookup-Lists |
| `_Registry`, `_Work`, `_Export`, `_Audit`, `_Picker`, `_Data` | Very hidden | Working sheets; cannot be unhidden from the tab menu | Admin-Protection-and-Hidden-Sheets |

## Buttons and controls

### Print room sheet

| Control | Macro | Code | Page | Dialogs | Tests | Platform |
| --- | --- | --- | --- | --- | --- | --- |
| **New** (Print jobs panel) | `btnAddPrintJob` | `modJobs.AddPrintJob` | Jobs-Recording | none | defaults, newpaperrow | Check |
| **Repeat** | `btnRepeatJob` | `modJobs.RepeatJob` | Jobs-Recording | none | multipass | Check |
| **Delete** | `btnRemoveRow` | `modJobs.RemoveRow` | Jobs-Recording | Yes/No | nextid, removelastrow | Check |
| **Add** (Colour passes panel) | `btnAddPass` | `modPasses.AddPass` | Jobs-Colour-Passes | none | multipass | Check |
| **Remove** (Colour passes panel) | `btnRemovePass` | `modPasses.RemovePass` | Jobs-Colour-Passes | Yes/No | multipass | Check |
| **Toggle passes** | `btnTogglePasses` | `modPasses.TogglePasses` | Jobs-Colour-Passes | none | multipass | Check |
| **Toggle all passes** | `btnToggleAllPasses` | `modPasses.ToggleAllPasses` | Jobs-Colour-Passes | none | multipass | Check |
| **Now** | `btnNow` | `modJobs.StampNow` | Jobs-Recording | none | | Check |
| Double-click a blank Date/Time cell | event | `ThisWorkbook.Workbook_SheetBeforeDoubleClick` | Jobs-Recording | none | | Check |
| **Show columns:** drop-down (All, Reduced, Minimal) | `ddViewMode` | `modInit.ViewDropDownChanged` | Jobs-Views-and-Units | none | reducedview, viewedge, manualhide | Check |
| **Hide cost detail** / **Show cost detail** | `btnToggleCostColumns` | `modInit.ToggleCostColumns` | Jobs-Views-and-Units | none | costcolumns | Check |
| **Hide pass columns** / **Show pass columns** | `btnToggleMultiPass` | `modPasses.ToggleMultiPassColumns` | Jobs-Colour-Passes | none | multipass-views | Check |
| **Clear defaults** | `btnClearDefaults` | `modJobs.ClearDefaults` | Jobs-Recording | none | defaults | Check |
| **Select printers...** | `btnSelectPrinters` | `modPicker.PickPrinters` | Setup-Print-Rooms | picker sheet (OK, Cancel) | | Check |
| **Check this sheet** | `btnCheckSheet` | `modCatalog.LoadCatalog`, `modValidation` | Jobs-Validation-and-Status | message | validation, jobvalidation, statusnotes | Check |
| **Clear All** | `btnClearAll` | `modJobs.ClearAll` | Data-Deleting-and-Clearing | Yes/No | clearformulas | Check |
| **Export...** | `btnExport` | `modExport.ExportLocation` | Data-Export-Room | names Yes/No | export, exportnames, macsave | Mac: writes beside the workbook; a OneDrive workbook reports a URL (plan, known differences) |
| **Import...** | `btnImportLocation` | `modImport.ImportIntoLocation` | Data-Import | file picker, Yes/No | import, importunits | Mac: the picker falls back through several methods (`modUtils.PickCsvFile`) |

### Summary

| Control | Macro | Code | Page | Dialogs | Tests | Platform |
| --- | --- | --- | --- | --- | --- | --- |
| **Refresh Locations** | `btnRefreshLocations` | `modRegistry.RefreshLocations` | Setup-Print-Rooms | message | duplicate, nextid, restorereuse | Check |
| **Check workbook** | `btnCheckWorkbook` | `modValidation.CheckWorkbook` | Jobs-Validation-and-Status | message | validation | Check |
| **Hide settings sheets** / **Show settings sheets** | `btnToggleConfigSheets` | `modInit.ToggleConfigSheets` | Admin-Protection-and-Hidden-Sheets | none | reports | Check |
| Plan a print job (Printer, Paper, Roll length, Print width, Sheet size) | cell edits | `modPlanner.PlannerChanged` | Reports-Summary | none | jobplanner | Check |
| **Add to print room** | `btnPlannerAdd` | `modPlanner.AddPlannedJob` | Reports-Summary | room drop-down; message if the plan is incomplete | jobplanner | Check |
| By department block | formulas | `modReports.BuildSummary` | Reports-Summary | none | departments, summaryformat | Check |

### Reports

| Control | Macro | Code | Page | Dialogs | Tests | Platform |
| --- | --- | --- | --- | --- | --- | --- |
| Filters (see below) | cell edits | live formulas over `_Data` | Reports-Filtering | none | reportsfilters | Check |
| **Clear all filters** | `btnClearFilters` | `modReports.ClearReportFilters` | Reports-Filtering | none | reportsfilters | Check |
| **Toggle Paid** (Selected record) | `btnTogglePaid` | `modReports.TogglePaidSelected` | Jobs-Paid-Tracking | Yes/No | togglepaid | Check |
| **Go to record** (Selected record) | `btnGoToRecord` | `modReports.GoToReportRecord` | Reports-Filtering | none | reportsnav | Check |
| **Paid** / **Unpaid** (Mark all as...) | `btnMarkPaid`, `btnMarkUnpaid` | `modReports.MarkVisibleReports` | Jobs-Paid-Tracking | Yes/No | markpaid | Check |
| **Export report...** | `btnExportReport` | `modExport.ExportReportSnapshot` | Reports-Export | names Yes/No | export, exportnames | Mac: as Export... above |
| **Delete visible records...** | `btnDeleteVisible` | `modReports.DeleteVisibleReports` | Data-Deleting-and-Clearing | two Yes/No | deletereports | Check |
| Paid column (Yes/No drop-down, read-only otherwise) | cell edit | `modReports` | Jobs-Paid-Tracking | none | paid | Check |
| Captions **Selected record:**, **Mark all as...** | text boxes | `modReports` | Reports-Filtering | none | | Same since 0.11.2 snags (plan) |

Filters, in sheet order: Student/dept. name, Student number, Department, From
date, To date, Min cost, Max cost, Paid, Has a problem, Disregarded (column A);
Location (print room), Technician, Printer model, Printer type, Paper stock,
Paper type, Student-supplied paper, Has notes (column D); Sort by, Sort
direction. Totals row: Jobs, Gross, Disregarded, Chargeable, Paid, Unpaid, and
After subsidy when `SUBSIDY_PCT` is set. Side blocks: By print room, By paper
stock, By department.

### Settings

| Control | Macro | Code | Page | Dialogs | Tests | Platform |
| --- | --- | --- | --- | --- | --- | --- |
| **Add print room...** | `btnAddPrintRoom` | `modRegistry.AddPrintRoom` | Setup-Print-Rooms | typed name, code | addroom, duplicate | Check |
| **Rename print room...** | `btnRenamePrintRoom` | `modRegistry.RenamePrintRoom` | Setup-Print-Rooms | typed code, name | renameroom | Check |
| **Remove print room...** | `btnRemovePrintRoom` | `modRegistry.RemovePrintRoom` | Setup-Print-Rooms | typed code, Yes/No, typed name | removeroom | Mac: prompts kept under 255 characters since 0.11.2 |
| **Refresh Locations** | `btnRefreshLocations` | `modRegistry.RefreshLocations` | Setup-Print-Rooms | message | duplicate, nextid | Check |
| **Export All Locations...** | `btnExportAll` | `modExport.ExportAllLocations` | Data-Export-Room | names Yes/No | export | Mac: as Export... above |
| **Import (choose room)...** | `btnImportGlobal` | `modImport.ImportGlobal` | Data-Import | file picker, typed code, Yes/No | import | Mac: as Import... above |
| **Backup workbook...** | `btnBackupWorkbook` | `modBackup.BackupAll` | Data-Backup-and-Restore | message | backup, macsave | Mac: before 0.11.1 could delete the workbook (fixed) |
| **Restore workbook...** | `btnRestoreWorkbook` | `modBackup.RestoreWorkbook` | Data-Backup-and-Restore, Admin-Upgrading | file picker, Yes/No | restoreworkbook, restoreexample, restorerooms, restorereuse | Mac: Excel asks about the older workbook's macros; choose Disable Macros |
| **Check workbook** | `btnCheckWorkbook` | `modValidation.CheckWorkbook` | Jobs-Validation-and-Status | message | validation | Check |
| **Re-stamp prices...** | `btnReStamp` | `modSnapshot.ReStampAll` | Setup-Settings | Yes/No | | Check |
| **About** | `btnAbout` | `modVersion.ShowAbout` | Getting-Started | message | | Check |
| **+** / **-** beside Paper stock types | `btnAddRowPaperTypes`, `btnRemoveRowPaperTypes` | `modCatalog.AddCatalogRow`, `RemoveCatalogRow` | Setup-Lookup-Lists | Yes/No on remove | settingsrows | Check |
| **+** / **-** beside Standard paper sizes | `btnAddRowSize`, `btnRemoveRowSize` | as above | Setup-Lookup-Lists | Yes/No on remove | settingsrows | Check |
| **+** / **-** beside Consumable types | `btnAddRowConsumable`, `btnRemoveRowConsumable` | as above | Setup-Lookup-Lists | Yes/No on remove | settingsrows | Check |
| Layout of the button rows (Print rooms, Data, Workbook) | | `modInit` | Setup-Settings | | settingslayout | Check |

### Catalogue sheets

Print Technicians, Printers, Papers, Consumables and Departments each have the
same three buttons at row 4, calling `modCatalog` with their table name.

| Control | Macro (per sheet) | Code | Page | Dialogs | Tests | Platform |
| --- | --- | --- | --- | --- | --- | --- |
| **Add row** | `btnAddRow<Table>` | `modCatalog.AddCatalogRow` | the sheet's Setup page | none | catalogids, newpaperrow | Check |
| **Remove row** | `btnRemoveRow<Table>` | `modCatalog.RemoveCatalogRow` | the sheet's Setup page | Yes/No | | Check |
| **Clear table** | `btnClear<Table>` | `modCatalog.ClearCatalogTable` | Data-Deleting-and-Clearing | Yes/No | clearcatalog | Check |
| Typing a row under the table (gets its ID) | cell edit | `modCatalog.OnCatalogEdited`, `OnPaperEdited`, `modDepartments.OnDepartmentEdited`, `modConsumables.OnColourEdited` | the sheet's Setup page | duplicate-name warning | catalogids, departments | Check |
| Papers: Std. size fills Width and Height | cell edit | `modCatalog.OnPaperEdited` | Setup-Papers | none | suppliedstock | Check |

## Job row behaviour

Things that happen when a cell on a print room's job table is edited
(`modValidation.OnCellChanged` unless named).

| Column | What happens | Page | Tests | Platform |
| --- | --- | --- | --- | --- |
| Printer, Paper Stock | Drop-downs filter each other to compatible choices; incompatible pairs flagged | Jobs-Validation-and-Status | dropdowns, validation | Check |
| Qty | Whole sheets for sheet stock; metres or centimetres for roll (room's Roll length unit) | Jobs-Views-and-Units | rollunit | Check |
| Print Width mm, Sheet size | Checked against the roll or the printer's max sheet | Jobs-Validation-and-Status | validation | Check |
| Student Name, Student No | Name/number pair cross-checked against every room; warns, never blocks | Jobs-Validation-and-Status | validation | Check |
| Student Name matching a Free department | Sets both Disregard columns to Yes (`modDepartments.StampDepartmentDisregards`) | Setup-Departments | departments | Check |
| Technician | From the active technician list | Jobs-Recording | dropdowns | Check |
| Colour (pass rows) | Colours of the printer's consumable type | Jobs-Colour-Passes | multipass | Check |
| Status, hover note | Problems listed in Status with a matching cell note | Jobs-Validation-and-Status | statusnotes | Check |
| Paid | Yes/No | Jobs-Paid-Tracking | paid | Check |

## Print room settings (top of each room sheet)

| Cell | Name | Label | Page |
| --- | --- | --- | --- |
| B1 | `LOC_Name` | Print room (locked once set; use Rename print room) | Setup-Print-Rooms |
| B2 | `LOC_Dept` | Department | Setup-Print-Rooms |
| B4 | `LOC_DefTech` | Default: technician | Jobs-Recording |
| B5 | `LOC_DefPrinter` | Default: printer | Jobs-Recording |
| B6 | `LOC_DefPaper` | Default: paper | Jobs-Recording |
| B8 | `LOC_DefDisPaper` | Disregard paper | Jobs-Recording |
| B9 | `LOC_DefDisCons` | Disregard consumable | Jobs-Recording |
| B11 | `LOC_RollUnit` | Roll length unit | Jobs-Views-and-Units |
| B12 | | Printers at this location (from `LOC_Printers`) | Setup-Print-Rooms |
| AX2 | `LOC_Code` | Location code (side panel) | Setup-Print-Rooms |
| AX6 | `LOC_Status` | Sheet status (side panel) | Jobs-Validation-and-Status |
| AX8 | `LOC_Export` | Export (side panel; at AX6 over Sheet status before 0.11.3, finding 1) | Data-Export-Room |
| AX9 | `LOC_Printers` | Permitted printers (raw list) | Setup-Print-Rooms |

## Settings keys

All on Settings, `tblSettings`, each with a `SET_<KEY>` name. Read-only rows
are locked and skipped by Restore.

| Key | Label | Editable | Page |
| --- | --- | --- | --- |
| `SITE_ID` | Site ID | Yes | Setup-Settings |
| `SITE_NAME` | Site name | Yes | Setup-Settings |
| `ORG` | Organisation | Yes | Setup-Settings |
| `DEPT` | Department | Yes | Setup-Settings |
| `CURRENCY` | Currency symbol | Yes | Setup-Settings |
| `SUBSIDY_PCT` | University subsidy | Yes | Setup-Settings, Reports-Filtering |
| `ROUND_DP` | Currency decimal places | Yes (recalculates existing records) | Setup-Settings |
| `EXPORT_FOLDER` | Export folder | Yes | Data-Export-Room |
| `LOC_REDUCED_VIEW` | Location column view | Set by the Show columns drop-down | Jobs-Views-and-Units |
| `LOC_REDUCED_COLUMNS` | Reduced view - hidden columns | Yes | Jobs-Views-and-Units |
| `LOC_MINIMAL_COLUMNS` | Minimal view - extra hidden columns | Yes | Jobs-Views-and-Units |
| `COST_COLS_HIDDEN` | Hide cost detail columns | Set by the button | Jobs-Views-and-Units |
| `MULTIPASS_COLS_HIDDEN` | Hide multi-pass columns | Set by the button | Jobs-Colour-Passes |
| `REPORT_HEADING` | Report heading | Yes | Reports-Export |
| `SCHEMA` | Schema version | Read-only | Admin-Upgrading |
| `APP_VER` | Workbook version | Read-only | Admin-Upgrading |
| `BUILT`, `BUILT_BY` | Built, Built by | Read-only | Admin-Upgrading |
| `LASTREF` | Last locations refresh | Read-only | Setup-Print-Rooms |
| `TECH_ID_HWM`, `PRINTER_ID_HWM`, `STOCK_ID_HWM`, `DEPT_ID_HWM`, `COLOUR_ID_HWM` | Last ... ID number | Read-only | Glossary (IDs) |

## Catalogue tables

| Table | Sheet | Columns | Page |
| --- | --- | --- | --- |
| `tblTechnicians` | Print Technicians | TechID, Name, Department, Active | Setup-Print-Technicians |
| `tblPrinters` | Printers | PrinterID, Model, Consumable type, Cost per m2, Colour mode, Template cost, Max roll width mm, Max sheet size, Active | Setup-Printers |
| `tblPapers` | Papers | StockID, Description, Paper type, Measure, Size mode, Std. size, Width mm, Height mm, Cost, Cost unit, Active, Supplied by student | Setup-Papers |
| `tblColours` | Consumables | ColourID, Consumable type, Colour, Cost per m2, Active | Setup-Consumables |
| `tblDepartments` | Departments | DeptID, Name, Aliases, Free, Active, Notes, Match list (hidden), Dis Paper %, Dis Cons %, Allowance, Allowance period (last four reserved) | Setup-Departments |
| `tblPaperTypes`, `tblStandardSizes`, `tblConsumables` | Settings | Paper type / Size name, Width mm, Height mm / Consumable type; each with Active except sizes | Setup-Lookup-Lists |
| `tblJobs_<CODE>` | each print room | 45 columns, A:AS; `H_` and `S_` columns hidden | Jobs-Recording |

## Findings from the inventory

1. **Sheet status was never shown (fixed in 0.11.3).** The template puts
   Sheet status at AX6 and Export at AX8, but `modExport.EXPORT_CELL` was
   `$AX$6`, so every refresh wrote the export line and an "Export" label over
   Sheet status. 0.11.3 writes Export to AX8 (branch `fix/sheet-status-cell`,
   checked by test-statusnotes). Jobs-Validation-and-Status documents the
   0.11.3 behaviour.
2. **No test** for Now, the double-click stamp, Select printers..., Re-stamp
   prices..., About or Remove row on a catalogue sheet. They are short, but
   phase 1 should look at them by hand.

## Phase 1 Mac checks

What the test suite cannot see: anything a person reads or clicks. Each check
fills in the Platform column above. Run on Excel for Mac 16.114 with the
built workbook in Excel's container, macros enabled.

- [ ] Every prompt above appears, reads correctly, and Cancel leaves nothing changed (Delete, Remove pass, Clear All, Remove row, Clear table, Re-stamp, Toggle Paid, Mark all as, Delete visible records, Restore)
- [ ] Typed prompts: Add, Rename and Remove print room, Import (choose room)
- [ ] Add to print room: the room drop-down, and the messages when the plan is incomplete
- [ ] File pickers: Import..., Import (choose room)..., Restore workbook... (CSV and older workbook)
- [ ] Export, Export All, Export report and Backup: where the files land, and the message shown
- [ ] Select printers... picker sheet: OK and Cancel
- [ ] Layout: room sheet panels and side panel at All, Reduced and Minimal; Reports with the filter group collapsed and expanded; Settings button rows
- [ ] About text and version
- [ ] Double-click Date/Time and the Now button
