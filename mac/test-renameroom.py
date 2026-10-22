# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-renameroom.ps1 - same checks, same messages, same order.
# The name is typed into the room's protected sheet the way a person types it
# (an AppleScript cell write fires Worksheet_Change, as a COM write does).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, sheet, table, check, section, finish  # noqa: E402

FIRST, RENAMED, TAKEN = 'Rename Me', 'Renamed Room', 'Example Print Room'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    ws = wb.call('CreatePrintRoom', FIRST, 'Facilities', 'RENAME')
    check(ws.startswith('@@h:'), f"CreatePrintRoom made '{FIRST}'")
    reg_ws = next((s for s in wb.sheet_names() if 'tblLocations' in wb.table_names(s)), None)

    def reg_row(sheet_name):
        """The registry row for a sheet, as {header: text}, or None."""
        heads = wb.table_headers(reg_ws, 'tblLocations')
        cols = {h: wb.table_column_texts(reg_ws, 'tblLocations', h) for h in heads}
        for i, s in enumerate(cols['SheetName']):
            if s == sheet_name:
                return {h: cols[h][i] for h in heads}
        return None

    def audit_last():
        heads = wb.table_headers('_Audit', 'tblAudit')
        return [wb.table_column_texts('_Audit', 'tblAudit', h)[-1] for h in heads]

    # ------------------------------------------------------------- rename
    section('Typing a new room name renames the tab', first=True)
    # Two job IDs handed out, so the room has a Job ID high-water mark to keep.
    for _ in range(2):
        wb.call('NextJobId', sheet(FIRST), table('tblJobs_RENAME'))
    hwm = (reg_row(FIRST) or {}).get('Job ID HWM', '')
    check(hwm == '2', f'the room has a Job ID high-water mark of 2 (got {hwm!r})')
    name_a1 = wb.name_addr(FIRST, 'LOC_Name')
    check(wb.protected(FIRST), 'the sheet is protected')
    check(wb.prop(f'@@range:{FIRST}!{name_a1}', 'Locked') == 'False', 'the room name cell is not locked')
    wb.set_value(FIRST, name_a1, RENAMED)
    names = wb.sheet_names()
    check(RENAMED in names and FIRST not in names, f"the tab now reads '{RENAMED}' (sheets: {', '.join(names)})")
    check(wb.protected(RENAMED), 'the sheet is still protected')
    row = reg_row(RENAMED)
    check(row is not None and row['Name'] == RENAMED, f'the registry row follows the new tab name ({row})')
    check(row is not None and row['Code'] == 'RENAME', 'the room keeps its code')
    check(row is not None and row['Job ID HWM'] == '2', 'the Job ID high-water mark came across')
    check(reg_row(FIRST) is None, 'no registry row is left under the old tab name')
    last = audit_last()
    check('Rename print room' in last and any(FIRST in c for c in last), f'the rename is in the audit log ({last})')

    # ---------------------------------------------------------- name taken
    section('A name another sheet holds gets a suffix')
    wb.set_value(RENAMED, name_a1, TAKEN)
    names = wb.sheet_names()
    check(f'{TAKEN} 2' in names and RENAMED not in names, f"the tab reads '{TAKEN} 2' (sheets: {', '.join(names)})")
    check(wb.text(TAKEN, wb.name_addr(TAKEN, 'LOC_Name')) == TAKEN, f"'{TAKEN}' itself is untouched")

    # --------------------------------------------------------------- blank
    section('A blank name leaves the tab alone')
    wb.set_value(f'{TAKEN} 2', name_a1, '')
    check(f'{TAKEN} 2' in wb.sheet_names(), f"the tab still reads '{TAKEN} 2'")

    wb.run('SetQuiet', False)

finish()
