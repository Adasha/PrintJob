# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-layout.ps1 - same checks, same messages, same order; see
# that script for the column layout and view behaviour it guards.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, SUM_TOT_ROW  # noqa: E402

MAIN = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    def sheet_col(header):
        return wb.table_col(MAIN, TBL, header)

    def level(col):
        return wb.column_outline_level(MAIN, col)

    def hidden(col):
        return wb.column_hidden(MAIN, col)

    # ------------------------------------------------------- outline / hidden
    section('No outline group; Notes visible; H_Issues/snapshot hidden', first=True)
    for h in ('Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded', 'Chargeable Cost', 'Paid'):
        lv = level(sheet_col(h))
        check(lv == 1, f'{h} is at outline level 1, no group (got {lv})')
    notes_col = sheet_col('Notes')
    lv = level(notes_col)
    check(lv == 1, f'Notes is at outline level 1 (got {lv})')
    check(not hidden(notes_col), 'Notes is NOT hidden')
    for h in ('H_Issues', 'S_PrinterID', 'S_StampedBy'):
        c = sheet_col(h)
        lv = level(c)
        check(lv == 1, f'{h} is at outline level 1, no group (got {lv})')
        check(hidden(c), f'{h} is hidden via .Hidden, not a group')
    names = wb.table_headers(MAIN, TBL)
    col0 = sheet_col(names[0])
    max_level = max(level(col0 + i) for i in range(len(names)))
    check(max_level == 1, f'max outline level across the whole table is 1 (got {max_level})')
    print()
    # ------------------------------------------------------------ new order
    print('=== Column order ===')
    print('  ' + ', '.join(names))
    check(names[0] == 'Date/Time', f"column 1 is Date/Time (got '{names[0]}')")

    def idx(n):
        return names.index(n) + 1 if n in names else 0

    status_idx, job_id_idx, paid_idx = idx('Status'), idx('Job ID'), idx('Paid')
    notes_idx, schema_idx, pass_idx = idx('Notes'), idx('S_SchemaVer'), idx('Pass')
    check(status_idx == paid_idx + 1, f'Status sits right after Paid (paid={paid_idx} status={status_idx})')
    check(job_id_idx == status_idx + 1, f'Job ID sits right after Status (status={status_idx} jobId={job_id_idx})')
    check(pass_idx == job_id_idx + 3 and names[job_id_idx] == 'Row Type' and names[job_id_idx + 1] == 'Parent',
          f'Row Type, Parent and Pass follow Job ID (jobId={job_id_idx} pass={pass_idx})')
    check(notes_idx == pass_idx + 1, f'Notes sits right after Pass (pass={pass_idx} notes={notes_idx})')
    check(schema_idx == len(names), f'S_SchemaVer is still the last column (idx={schema_idx} of {len(names)})')

    # ------------------------------------------------------------- header block
    section('Header block survives before toggling')
    check(wb.text(MAIN, 'B1') == 'Example Print Room', 'B1 (room name) reads correctly before toggle')
    check(wb.text(MAIN, 'B2') == 'Print Services', 'B2 (department) reads correctly before toggle')

    section('Reduce clutter ON: header block still visible, button still there')
    wb.run('SetViewMode', 'Reduced')
    time.sleep(0.3)
    check(not hidden(1), 'column A is NOT hidden (header block lives there)')
    check(not hidden(2), 'column B is NOT hidden (header block lives there)')
    check(wb.text(MAIN, 'B1') == 'Example Print Room', 'B1 (room name) still reads correctly with reduced view ON')
    check(wb.text(MAIN, 'B2') == 'Print Services', 'B2 (department) still reads correctly with reduced view ON')

    sheet_status_col = col0 + status_idx - 1
    sheet_job_id_col = col0 + job_id_idx - 1
    check(hidden(sheet_status_col), "Status's own (new) column IS hidden")
    check(hidden(sheet_job_id_col), "Job ID's own (new) column IS hidden")

    found = any(n.startswith('pcb_ddViewMode') for n in wb.dropdown_names(MAIN))
    check(found, 'the view drop-down is still present after switching to Reduced')

    section('Reduce clutter OFF again: everything restored')
    wb.run('SetViewMode', 'All')
    time.sleep(0.3)
    check(not hidden(sheet_status_col), "Status's column visible again")
    check(not hidden(sheet_job_id_col), "Job ID's column visible again")

    # ---------------------------------------------------------------- sanity
    section('Downstream sanity: Summary/Reports still reconcile')
    wb.calculate_full_rebuild()
    chargeable = wb.number('Summary', f'H{SUM_TOT_ROW}')
    paid = wb.number('Summary', f'J{SUM_TOT_ROW}')
    unpaid = wb.number('Summary', f'L{SUM_TOT_ROW}')
    check(abs((paid + unpaid) - chargeable) < 0.01,
          f'Summary Paid+Unpaid still reconciles to Chargeable ({paid} + {unpaid} = {paid + unpaid}, Chargeable={chargeable})')
    spill_rows = int(wb.prop('@@spill:_Data!A10', 'Rows.Count'))
    check(spill_rows >= 5, f'consolidated range still has rows (got {spill_rows})')

    wb.run('SetQuiet', False)

finish()
