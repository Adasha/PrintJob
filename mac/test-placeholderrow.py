# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-placeholderrow.ps1 - same checks, same messages, same order.
# An empty print room keeps one blank row. With a single printer permitted,
# re-binding the drop-downs (Select printers... OK, Refresh Locations) wrote
# that printer into the blank row, which then counted as a job: "1 print job -
# never exported", and New could not reuse the row (found in the 2026-10-27
# Mac checks; fixed in 0.11.4, modLists.BindColumns). A real job still gets
# the single printer filled in.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, table, check, section, finish  # noqa: E402

ROOM, PRINTER = 'One Printer Room', 'Xerox Versant 180'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    ws = wb.call('CreatePrintRoom', ROOM, '', 'ONEPRN')
    check(ws.startswith('@@h:'), f"CreatePrintRoom made '{ROOM}'")
    lo = table('tblJobs_ONEPRN')

    def col(h):
        return wb.table_column_texts(ROOM, 'tblJobs_ONEPRN', h)

    # ------------------------------------------------------------ blank row
    section('A single permitted printer is not written into the blank row', first=True)
    wb.call('UnlockSheet', ws)
    wb.set_value(ROOM, wb.name_addr(ROOM, 'LOC_Printers'), PRINTER)
    wb.call('RelockSheet', ws)
    wb.call('BindColumns', ws)
    check(col('Printer') == [''], f"the blank row's Printer is still empty (got {col('Printer')})")
    got = wb.call('RowCount', lo)
    check(got == '0', f'the room still has no jobs (got {got!r})')
    got = wb.call('ExportStatusText', ws)
    check(got == 'No print jobs recorded.', f'the export line says there is nothing to export (got {got!r})')

    # ------------------------------------------------------------- real job
    section('New still fills in the single printer, in the blank row')
    wb.call('AddPrintJob', ws)
    got = wb.prop(lo, 'ListRows.Count')
    check(got == '1', f'New reused the blank row (rows: {got})')
    check(col('Printer') == [PRINTER], f"the new job's Printer is {PRINTER} (got {col('Printer')})")
    check(col('Job ID')[0].startswith('UNI-ONEPRN-'), f"the new job has a Job ID (got {col('Job ID')})")
    wb.call('BindColumns', ws)
    got = wb.call('RowCount', lo)
    check(got == '1', f're-binding leaves one job (got {got!r})')

    wb.run('SetQuiet', False)

finish()
