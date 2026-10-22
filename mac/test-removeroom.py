# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-removeroom.ps1 - the same steps and the same printed report
# (it reports rather than asserts). The .ps1 injects a TestRemoveRoom()
# function that strings these calls together; here they are made one at a
# time through modScript.ScriptRun, with the new room's Worksheet held as a
# handle, so nothing is injected. DeletePrintRoom is called directly, not
# RemovePrintRoom, for the InputBox reason the .ps1 gives.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, ExcelError  # noqa: E402


def table_lines(wb, ws, t):
    heads = wb.table_headers(ws, t)
    cols = [wb.table_column_texts(ws, t, h) for h in heads]
    return [' | '.join(c[i] for c in cols) + ' | ' for i in range(len(cols[0]) if cols else 0)]


with open_copy() as wb:
    r = ''
    try:
        before = len(wb.sheet_names())
        r += f'sheets before: {before}\n'

        wb.run('SetQuiet', True)
        room2 = wb.call('CreatePrintRoom', 'Test Room To Remove', 'Facilities', 'REMOVE')
        wb.run('SetQuiet', False)
        if not room2.startswith('@@h:'):
            print(r + 'FAIL: CreatePrintRoom returned Nothing')
            sys.exit(1)
        removed = wb.prop(room2, 'Name')
        r += f"second room created: '{removed}' LOC_Code={wb.call('LocValue', room2, 'LOC_Code')}\n"
        r += f'sheets with second room: {len(wb.sheet_names())}\n'

        wb.run('SetQuiet', True)
        wb.call('DeletePrintRoom', room2)
        r += 'Refresh Locations said:\n' + wb.run('QuietLog')
        wb.run('SetQuiet', False)

        names = wb.sheet_names()
        r += f'sheets after removal: {len(names)}\n'
        r += f'removed sheet still present: {removed in names}\n'
        reg_ws = next((s for s in names if 'tblLocations' in wb.table_names(s)), None)
        r += 'registry after removal:\n' + (''.join(f'   {ln}\n' for ln in table_lines(wb, reg_ws, 'tblLocations'))
                                           if reg_ws else '   (no tblLocations)\n')
        audit = table_lines(wb, '_Audit', 'tblAudit')
        r += f"last audit row: {audit[-1] if audit else '(no audit rows)'}\n"
    except ExcelError as e:
        r += f'ERROR {e}\n'
    print(r)
