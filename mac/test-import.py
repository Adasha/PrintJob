# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-import.ps1 - same steps, same output lines (OK:/FAIL: and
# PASS, not Check); see that script for why it drives ReadImportRows +
# ApplyImportConfirmed rather than the file pickers. The rows ReadImportRows
# returns come back as a modScript handle and go straight into
# ApplyImportConfirmed. The CSV is imported exactly where and as Excel wrote
# it: Excel's sandbox will not read a copy.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself.

import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, addr, col_letters, sheet, spill  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

MAIN, MTBL = 'Example Print Room', 'tblJobs_EXAMPLE'
ANNEX, ATBL = 'Annexe', 'tblJobs_ANNEX'


def fail(msg):
    print(f'FAIL: {msg}')
    sys.exit(1)


ctx = open_copy()
with ctx as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def col(ws, table, name):
        c = wb.table_col(ws, table, name)
        if c == 0:
            raise RuntimeError(f"column '{name}' not found")
        return c

    id_col, chg_col, unit_cost_col = col(MAIN, MTBL, 'Job ID'), col(MAIN, MTBL, 'Chargeable Cost'), col(MAIN, MTBL, 'S_UnitCost')

    def ids(ws, table):
        return wb.table_column_texts(ws, table, 'Job ID')

    # --- Export All Locations -------------------------------------------
    print('=== Export All Locations ===')
    wb.run('btnExportAll')
    print(wb.run('QuietLog'))
    found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-*-EXAMPLE-*.csv')), key=os.path.getmtime)
    if not found:
        fail('no CSV produced for Example Print Room')
    main_csv = found[-1]
    print(f'  {os.path.basename(main_csv)}')

    # --- capture the row we are about to delete, then delete it ----------
    victim = wb.table_first_row(MAIN, MTBL) + 1
    victim_id = wb.cell_text(MAIN, victim, id_col)
    victim_chg = wb.number(MAIN, addr(victim, chg_col))
    victim_unit_cost = wb.number(MAIN, addr(victim, unit_cost_col))
    print()
    print(f'victim row: {victim_id}  chargeable={victim_chg}  S_UnitCost={victim_unit_cost}')

    wb.unprotect(MAIN)
    wb.delete_table_row(MAIN, MTBL, 1)
    wb.protect(MAIN)
    if victim_id in ids(MAIN, MTBL):
        raise RuntimeError('victim row was not actually deleted')
    print('deleted from Example Print Room')

    # --- restore it via import into the SAME room -------------------------
    print()
    print('=== Import: restore into origin ===')
    rows = wb.call('ReadImportRows', main_csv)
    wb.call('ApplyImportConfirmed', sheet(MAIN), rows)
    print(wb.run('QuietLog'))
    wb.calculate_full_rebuild()
    time.sleep(0.3)

    now = ids(MAIN, MTBL)
    if victim_id not in now:
        fail(f'{victim_id} was not restored')
    found_row = now.index(victim_id) + 1
    wb.calculate_full_rebuild()
    sheet_row = wb.table_first_row(MAIN, MTBL) + found_row
    restored_chg = wb.number(MAIN, addr(sheet_row, chg_col))
    restored_unit_cost = wb.number(MAIN, addr(sheet_row, unit_cost_col))
    print(f'restored row {found_row}: {victim_id}  chargeable={restored_chg}  S_UnitCost={restored_unit_cost}')
    if abs(restored_chg - victim_chg) > 0.001:
        fail(f'chargeable cost changed on restore ({victim_chg} -> {restored_chg}) - it was recosted instead of using the snapshot')
    if abs(restored_unit_cost - victim_unit_cost) > 0.001:
        fail(f'S_UnitCost changed on restore ({victim_unit_cost} -> {restored_unit_cost})')
    print('OK: restored row matches the original snapshot exactly (not recosted)')

    # --- import the SAME export into a DIFFERENT room ----------------------
    print()
    print('=== Import: into a different room ===')
    rows2 = wb.call('ReadImportRows', main_csv)
    wb.call('ApplyImportConfirmed', sheet(ANNEX), rows2)
    print(wb.run('QuietLog'))
    headers = wb.table_headers(ANNEX, ATBL)
    print(f'DIAG: annexe.Name={ANNEX} loAnnexe.Name={ATBL} ListColumns.Count={len(headers)} '
          f'ListRows.Count={wb.table_row_count(ANNEX, ATBL)}')
    for i, h in enumerate(headers, 1):
        print(f"  col {i}: '{h}'")
    if victim_id not in ids(ANNEX, ATBL):
        fail(f'{victim_id} was not imported into Annexe Print Room')
    if '-EXAMPLE-' not in victim_id:
        fail('test assumption broken, victim ID does not carry an EXAMPLE prefix')
    print(f"OK: '{victim_id}' landed in Annexe's table with its ORIGINAL Job ID (prefix untouched)")

    # --- confirm it now reports under Annexe's location in _Data ----------
    wb.run('btnRefreshLocations')
    wb.calculate_full_rebuild()
    heads = wb.texts('_Data', 'A9:AD9')[0]
    job_id_col = heads.index('Job ID') + 1 if 'Job ID' in heads else 0
    n = int(wb.prop(spill('_Data', 'A10'), 'Rows.Count'))
    locs = [r[0] for r in wb.texts('_Data', f'A10:A{9 + n}')]
    jc = col_letters(job_id_col)
    jids = [r[0] for r in wb.texts('_Data', f'{jc}10:{jc}{9 + n}')]
    # The LAST match, as the .ps1's loop leaves it: the job is in both rooms
    # now (restored into Main, imported into Annexe), and Annexe's comes after.
    matches = [lc for lc, j in zip(locs, jids) if j == victim_id]
    found_loc = matches[-1] if matches else None
    print(f"_Data reports {victim_id} under Location = '{found_loc}'")
    if found_loc != 'ANNEX':
        fail(f"expected Location 'ANNEX', got '{found_loc}'")
    print("OK: consolidated range reports the imported job under its NEW sheet's location")

    wb.run('SetQuiet', False)

print()
print('PASS')
