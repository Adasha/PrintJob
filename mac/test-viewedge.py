# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-viewedge.ps1 - same checks, same messages, same order; see
# that script for where the column views meet other features.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, sheet, ExcelError, col_letters  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
ANNEX, ATBL = 'Annexe', 'tblJobs_ANNEX'
REDUCED = ['Status', 'Job ID', 'Area m2']
MINIMAL_EXTRA = ['Printer', 'Disregard Paper', 'Disregard Consumable', 'Print Width mm', 'Sheet size']
ALL_LISTED = REDUCED + MINIMAL_EXTRA

ctx = open_copy()
with ctx as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def is_hidden(ws, table, header):
        c = wb.table_col(ws, table, header)
        if c < 1:
            raise RuntimeError(f"column lookup failed for '{header}'")
        return wb.column_hidden(ws, c)

    def dd_index(ws):
        d = next((d for d in wb.dropdowns(ws) if d['name'].startswith('pcb_ddViewMode')), None)
        return int(d['list_index']) if d else 0

    def view_setting():
        return wb.wb_name_text('SET_LOC_REDUCED_VIEW')

    def settle():
        time.sleep(0.3)

    def view(mode):
        wb.run('SetViewMode', mode)
        settle()

    def jobs_table(ws):
        return next((t for t in wb.table_names(ws) if t.startswith('tblJobs_')), None)

    # ------------------------------------------------- Refresh Locations
    section('Refresh Locations keeps Minimal', first=True)
    view('Minimal')
    wb.run('btnRefreshLocations')
    wb.calculate_full_rebuild()
    time.sleep(1)
    check(view_setting() == 'Minimal', 'setting still reads Minimal after Refresh Locations')
    for h in ALL_LISTED:
        check(is_hidden(MAIN, TBL, h), f'{h} still hidden on Example Print Room')
        check(is_hidden(ANNEX, ATBL, h), f'{h} still hidden on Annexe')
    check(dd_index(MAIN) == 3, 'drop-down still shows Minimal')

    # ------------------------------------------------------- Add print room
    section('Add print room while Minimal: new room takes the current view')
    wb.call('CreatePrintRoom', 'View Edge Room', 'Test', 'VEDGE')
    settle()
    wb.calculate_full_rebuild()
    new_ws = 'View Edge Room'
    new_lo = jobs_table(new_ws)
    check(new_lo is not None, 'new room has a jobs table')
    for h in ALL_LISTED:
        check(is_hidden(new_ws, new_lo, h), f'{h} hidden on the new room (inherits Minimal)')
    check(view_setting() == 'Minimal', 'creating a room did not reset the workbook view')
    check(dd_index(new_ws) == 3, "new room's drop-down shows Minimal")

    view('All')
    for h in ALL_LISTED:
        check(not is_hidden(new_ws, new_lo, h), f'{h} visible on the new room after switching to All')

    section('Add print room while All: new room starts with everything visible')
    wb.call('CreatePrintRoom', 'View Edge Two', '', 'VEDG2')
    settle()
    wb.calculate_full_rebuild()
    ws2 = 'View Edge Two'
    lo2 = jobs_table(ws2)
    for h in ALL_LISTED:
        check(not is_hidden(ws2, lo2, h), f'{h} visible on a room created in All')
    check(dd_index(ws2) == 1, "new room's drop-down shows All")

    # ------------------------------------------- Import into hidden columns
    section('Import into a sheet whose columns are hidden')
    wb.run('btnExportAll')
    found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-EXAMPLE-*.csv')), key=os.path.getmtime)
    main_csv = found[-1] if found else None
    check(main_csv is not None, 'export produced a CSV for Example Print Room')

    view('Minimal')
    printer_col = wb.table_col(ANNEX, ATBL, 'Printer')
    id_col = wb.table_col(ANNEX, ATBL, 'Job ID')
    # Annexe is a copy of Example Print Room, so its Job IDs already match the
    # export (import skips known IDs): remove two rows so there is something to bring back.
    wb.unprotect(ANNEX)
    wb.delete_table_row(ANNEX, ATBL, 2)
    wb.delete_table_row(ANNEX, ATBL, 1)
    wb.protect(ANNEX)
    settle()
    before = wb.table_row_count(ANNEX, ATBL)
    if main_csv:
        rows = wb.call('ReadImportRows', main_csv)
        wb.call('ApplyImportConfirmed', sheet(ANNEX), rows)
    wb.calculate_full_rebuild()
    settle()
    after = wb.table_row_count(ANNEX, ATBL)
    check(after >= before + 2, f'the two deleted rows were imported back ({before} -> {after})')
    first = wb.table_first_row(ANNEX, ATBL) + 1
    last = wb.table_last_row(ANNEX, ATBL)
    p = col_letters(printer_col)
    j = col_letters(id_col)
    filled = sum(1 for (v,) in wb.texts(ANNEX, f'{p}{first}:{p}{last}') if v != '')
    ids = sum(1 for (v,) in wb.texts(ANNEX, f'{j}{first}:{j}{last}') if v != '')
    check(filled >= after - before, f'Printer (hidden) holds a value on imported rows ({filled} filled)')
    check(ids >= after - before, f'Job ID (hidden) holds a value on imported rows ({ids} filled)')
    for h in ALL_LISTED:
        check(is_hidden(ANNEX, ATBL, h), f'{h} still hidden on Annexe after import')
    check(view_setting() == 'Minimal', 'import did not change the workbook view')
    view('All')

    # ------------------------------------- full setup pass in a hidden view
    section('InitialiseWorkbook (setup) in Minimal keeps the columns hidden')
    view('Minimal')
    wb.run('InitialiseWorkbook')
    wb.calculate_full_rebuild()
    time.sleep(1)
    for h in ALL_LISTED:
        check(is_hidden(MAIN, TBL, h), f'{h} still hidden on Example Print Room after setup')
    view('All')

    # ------------------------------------------------- bad Settings entries
    section('A Settings list naming a nonexistent header fails soft')
    red_cell = wb.wb_name('SET_LOC_REDUCED_COLUMNS')
    min_cell = wb.wb_name('SET_LOC_MINIMAL_COLUMNS')
    orig_list = wb.text(*red_cell)
    orig_extra = wb.text(*min_cell)
    wb.set_value(*red_cell, 'Status; No Such Column ;;Job ID;')
    wb.set_value(*min_cell, 'Bogus Header;Printer')

    def try_view(mode):
        try:
            wb.run('SetViewMode', mode)
            return True
        except ExcelError as e:
            print(f'  {e}')
            return False

    ok = try_view('Reduced')
    settle()
    check(ok, 'Reduced with a bad header in the list does not raise')
    check(is_hidden(MAIN, TBL, 'Status') and is_hidden(MAIN, TBL, 'Job ID'),
          'the real headers in the list are still hidden (bad entry skipped, not fatal)')
    check(not is_hidden(MAIN, TBL, 'Area m2'), 'a column not in the edited list stays visible')
    ok = try_view('Minimal')
    settle()
    check(ok, 'Minimal with a bad header in the extra list does not raise')
    check(is_hidden(MAIN, TBL, 'Printer'), 'Printer hidden in Minimal despite Bogus Header beside it')
    check(view_setting() == 'Minimal', 'setting reads Minimal')
    wb.run('btnRefreshLocations')
    settle()
    check(is_hidden(MAIN, TBL, 'Printer') and is_hidden(MAIN, TBL, 'Status'), 'Refresh Locations with a bad list still applies the view')

    section('Blank lists fall back to the defaults')
    wb.clear_contents(*red_cell)
    wb.clear_contents(*min_cell)
    view('All')
    view('Minimal')
    for h in ALL_LISTED:
        check(is_hidden(MAIN, TBL, h), f'{h} hidden in Minimal from the built-in default lists')

    wb.set_value(*red_cell, orig_list)
    wb.set_value(*min_cell, orig_extra)
    wb.run('SetViewMode', 'All')
    wb.run('SetQuiet', False)

finish()
