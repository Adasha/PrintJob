# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-nextid.ps1 - same steps, same output lines; see that script
# for why the high-water mark is checked across a deleted top row and a
# RefreshLocations rebuild. Like the .ps1 it reports with OK:/FAIL: lines and
# PASS rather than Check.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy  # noqa: E402

WS = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    id_col = wb.table_col(WS, TBL, 'Job ID')
    if id_col == 0:
        raise RuntimeError('Job ID column not found')
    first = wb.table_first_row(WS, TBL)

    def last_row_id():
        r = wb.table_last_row(WS, TBL)
        return r - first, wb.cell_text(WS, r, id_col)

    wb.activate(WS)
    wb.run('btnAddPrintJob')
    n1, id1 = last_row_id()
    print(f'added row {n1}: {id1}')
    if not id1:
        raise RuntimeError('first Add Print Job produced no Job ID')

    # Deleted directly, as the .ps1 does - RemoveRow's Ask() declines under SetQuiet.
    wb.unprotect(WS)
    wb.delete_table_row(WS, TBL, n1)
    wb.protect(WS)
    print(f'deleted row {n1} ({id1})')

    wb.run('btnAddPrintJob')
    n2, id2 = last_row_id()
    print(f'added row {n2}: {id2}')

    if id2 == id1:
        print()
        print(f"FAIL: '{id2}' reissued the deleted ID '{id1}'")
        sys.exit(1)
    print(f"OK: new ID '{id2}' does not reuse deleted ID '{id1}'")

    # The HWM must survive WriteRegistry's delete-and-rebuild of every
    # registry row, which RefreshLocations performs on every call.
    wb.run('btnRefreshLocations')
    wb.run('btnAddPrintJob')
    n3, id3 = last_row_id()
    print(f'after RefreshLocations, added row {n3}: {id3}')
    if id3 in (id1, id2):
        print()
        print(f"FAIL: '{id3}' collides with an earlier ID after RefreshLocations")
        sys.exit(1)
    print('OK: high-water mark survived RefreshLocations')

    wb.run('SetQuiet', False)

print()
print('PASS')
