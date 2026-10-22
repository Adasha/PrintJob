# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-validation.ps1 - the same steps and the same printed report
# (AT-11, AT-14, AT-16). Like the .ps1 it reports rather than asserts: read the
# "correct" / "BUG" verdicts. Quiet mode makes AT-14 non-destructive: Ask()
# declines and logs its prompt.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, addr, col_letters, sheet, spill  # noqa: E402

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'

with open_copy() as wb:
    before = wb.table_row_count(MAIN, TBL)

    # ---------------------------------------------------------------- AT-14
    print('=== AT-14: Clear All confirmation ===')
    wb.run('SetQuiet', True)
    wb.activate(MAIN)
    wb.run('btnClearAll')
    log = wb.run('QuietLog')
    print(log)
    after = wb.table_row_count(MAIN, TBL)
    verdict = 'NOT DELETED - Ask returned False as designed' if before == after else 'DELETED - BUG'
    print(f'  rows before={before} after={after}  {verdict}')
    for need in ('Example Print Room', '5 print job', '15/09/2026', '17/09/2026', 'cannot be undone'):
        print(f"  contains '{need:<18}' {'yes' if need in log else 'NO'}")

    # ---------------------------------------------------------------- AT-11
    print()
    print('=== AT-11: student name/number conflict ===')
    wb.run('SetQuiet', True)
    wb.unprotect(MAIN)
    name_col, no_col = wb.table_col(MAIN, TBL, 'Student Name'), wb.table_col(MAIN, TBL, 'Student No')
    r1 = wb.table_first_row(MAIN, TBL) + 1
    print(f'  row1: {wb.cell_text(MAIN, r1, name_col)} / {wb.cell_text(MAIN, r1, no_col)}')
    print(f'  row2: {wb.cell_text(MAIN, r1 + 1, name_col)} / {wb.cell_text(MAIN, r1 + 1, no_col)}')

    wb.set_value(MAIN, addr(r1 + 1, name_col), 'Jane Smyth')
    time.sleep(0.3)
    log = wb.run('QuietLog')
    print('  --- warning raised ---')
    print(log)
    now = wb.cell_text(MAIN, r1 + 1, name_col)
    print(f"  value kept as typed: '{now}'  {'WARNED not blocked - correct' if now == 'Jane Smyth' else 'BLOCKED - BUG'}")

    # ---------------------------------------------------------------- AT-16
    print()
    print('=== AT-16: inactive records ===')
    tech, tt = 'Print Technicians', 'tblTechnicians'
    wb.unprotect(tech)
    t_first = wb.table_first_row(tech, tt) + 1
    victim = wb.cell_text(tech, t_first, wb.table_col(tech, tt, 'Name'))
    print(f'  deactivating technician: {victim}')
    wb.set_value(tech, addr(t_first, wb.table_col(tech, tt, 'Active')), 'No')
    wb.run('Invalidate')
    wb.call('BindColumns', sheet(MAIN))

    # The dropdown lists live in staging columns on _Work, tagged in row 1.
    heads = wb.texts('_Work', 'A1:GR1')[0]
    listed = []
    if 'TEC' in heads:
        c = col_letters(heads.index('TEC') + 1)
        for (v,) in wb.texts('_Work', f'{c}2:{c}60'):
            if v == '':
                break
            listed.append(v)
    print(f"  technician dropdown now: {', '.join(listed)}")
    print(f"  '{victim}' excluded from dropdown: {'yes - correct' if victim not in listed else 'NO - BUG'}")

    # ...but their historical rows must remain in the reports.
    wb.calculate_full_rebuild()
    dheads = wb.texts('_Data', 'A9:AD9')[0]
    tc = col_letters(dheads.index('Technician') + 1) if 'Technician' in dheads else 'A'
    n = int(wb.prop(spill('_Data', 'A10'), 'Rows.Count'))
    still = sum(1 for (v,) in wb.texts('_Data', f'{tc}10:{tc}{9 + n}') if v == victim)
    print(f"  rows in the consolidated range still crediting '{victim}': {still}  "
          f"{'retained - correct' if still > 0 else 'LOST - BUG'}")

    wb.run('SetQuiet', False)
