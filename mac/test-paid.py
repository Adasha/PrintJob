# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-paid.ps1 - same checks, same messages, same order; see that
# script for the history behind each one. Differences:
#
#   - Validation.Type comes back as Excel's enum name, 'validate list', where
#     COM gives 3.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, SUM_TOT_ROW  # noqa: E402

MAIN = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    # --------------------------------------------------------------- schema
    section('Schema version bumped', first=True)
    # 1.4 since multi-pass costing (0.11.0); see test-paid.ps1 for the history.
    schema = wb.wb_name_text('SET_SCHEMA')
    check(schema == '1.4', f"SET_SCHEMA reads 1.4 (got '{schema}')")

    # ----------------------------------------------------------- column exists
    section('Paid column')
    # Sheet columns throughout; the .ps1 compares table-relative indexes, and
    # the offsets are the same for every column, so the comparisons agree.
    sheet_paid_col = wb.table_col(MAIN, TBL, 'Paid')
    sheet_chg_col = wb.table_col(MAIN, TBL, 'Chargeable Cost')
    notes_col = wb.table_col(MAIN, TBL, 'Notes')
    check(sheet_paid_col > 0, 'Paid column exists')
    check(sheet_paid_col == sheet_chg_col + 1,
          f'Paid sits immediately after Chargeable Cost (chg={sheet_chg_col} paid={sheet_paid_col})')
    check(notes_col > sheet_paid_col, f'Notes still sits after Paid (paid={sheet_paid_col} notes={notes_col})')

    first_data_row = wb.table_first_row(MAIN, TBL) + 1
    dv = wb.validation_type(MAIN, addr(first_data_row, sheet_paid_col))
    check(dv == 'validate list', 'Paid cells carry list (Yes/No) validation')

    # ---------------------------------------------------- existing rows blank
    section('Pre-existing rows are left blank, not force-set')
    row1_paid = wb.cell_text(MAIN, first_data_row, sheet_paid_col)
    check(row1_paid == '', f"row 1 (pre-existing, predates the column) has a blank Paid (got '{row1_paid}')")

    # ------------------------------------------------------------- new job
    section('New jobs default to No')
    wb.activate(MAIN)
    wb.run('btnAddPrintJob')
    new_row = wb.table_last_row(MAIN, TBL)
    new_paid = wb.cell_text(MAIN, new_row, sheet_paid_col)
    check(new_paid == 'No', f"new job's Paid defaults to No (got '{new_paid}')")

    # Mark row 1 (blank/unpaid) - leave one Yes for the totals check below.
    wb.set_value(MAIN, addr(first_data_row, sheet_paid_col), 'Yes')
    time.sleep(0.2)

    # ---------------------------------------------------------- Summary totals
    section('Summary totals split by paid status and reconcile')
    wb.calculate_full_rebuild()
    chargeable = wb.number('Summary', f'H{SUM_TOT_ROW}')
    paid = wb.number('Summary', f'J{SUM_TOT_ROW}')
    unpaid = wb.number('Summary', f'L{SUM_TOT_ROW}')
    check(paid > 0, f'Paid total is non-zero (got {paid})')
    check(abs((paid + unpaid) - chargeable) < 0.01,
          f'Paid + Unpaid reconciles to Chargeable ({paid} + {unpaid} = {paid + unpaid}, Chargeable={chargeable})')

    # ---------------------------------------------------------- Reports totals
    section('Reports "Matching" totals split by paid status and reconcile')
    r_chargeable = wb.number('Reports', 'N9')
    r_paid = wb.number('Reports', 'W9')
    r_unpaid = wb.number('Reports', 'X9')
    check(abs((r_paid + r_unpaid) - r_chargeable) < 0.01,
          f'Paid + Unpaid reconciles to Chargeable on Reports ({r_paid} + {r_unpaid} = {r_paid + r_unpaid}, Chargeable={r_chargeable})')

    # ------------------------------------------------------- cost visibility
    section('Chargeable Cost and Paid are never hidden by the cost-columns toggle')

    def hidden(c):
        return wb.column_hidden(MAIN, c)

    h = hidden(sheet_chg_col)
    check(not h, f'Chargeable Cost is visible by default (got hidden={h})')
    h = hidden(sheet_paid_col)
    check(not h, f'Paid is visible by default (got hidden={h})')
    wb.run('btnToggleCostColumns')
    time.sleep(0.3)
    h = hidden(sheet_chg_col)
    check(not h, f'Chargeable Cost stays visible once cost detail is hidden (got hidden={h})')
    h = hidden(sheet_paid_col)
    check(not h, f'Paid stays visible once cost detail is hidden (got hidden={h})')
    wb.run('btnToggleCostColumns')
    time.sleep(0.3)

    wb.run('SetQuiet', False)

finish()
