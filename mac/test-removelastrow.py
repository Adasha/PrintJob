# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-removelastrow.ps1 - same checks, same messages, same order;
# see that script for the header-formatting bug it guards. A cell's look is
# read through modScript.ScriptProp, and ListRows.Add is called through it
# too ("ListRows.Add.Index"), since neither is in AppleScript's dictionary.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, sheet, table, rng  # noqa: E402

CASES = [
    {'Sheet': 'Print Technicians', 'Table': 'tblTechnicians', 'Jobs': False},
    {'Sheet': 'Printers', 'Table': 'tblPrinters', 'Jobs': False},
    {'Sheet': 'Papers', 'Table': 'tblPapers', 'Jobs': False},
    {'Sheet': 'Departments', 'Table': 'tblDepartments', 'Jobs': False},
    {'Sheet': 'Consumables', 'Table': 'tblColours', 'Jobs': False},
    {'Sheet': 'Example Print Room', 'Table': 'tblJobs_EXAMPLE', 'Jobs': True},
]

with open_copy() as wb:
    wb.run('SetQuiet', True)

    def look(ws, a1):
        """How a cell looks, as one comparable string."""
        r = rng(ws, a1)
        p = [wb.prop(r, x) for x in ('Interior.Color', 'Interior.Pattern', 'Font.Bold', 'Font.Color', 'WrapText')]
        return f'fill={p[0]}/{p[1]} bold={p[2]} font={p[3]} wrap={p[4]}'

    def row_look(ws, t, n):
        """Looks of the first two columns of table row n."""
        row = wb.table_first_row(ws, t) + n
        c0 = wb.table_col(ws, t, wb.table_headers(ws, t)[0])
        return look(ws, addr(row, c0)) + ' | ' + look(ws, addr(row, c0 + 1))

    for i, c in enumerate(CASES):
        section(f"{c['Sheet']} / {c['Table']}", first=(i == 0))
        ws, t = c['Sheet'], c['Table']
        wb.call('UnlockSheet', sheet(ws))
        hdr = wb.table_first_row(ws, t)
        hdr_look = look(ws, addr(hdr, wb.table_col(ws, t, wb.table_headers(ws, t)[0])))

        # Down to one row (bypassing the sole-row guard, which only the last delete hits).
        while wb.table_row_count(ws, t) > 1:
            last = wb.table_last_row(ws, t)
            wb._do('delete entire row of range (item 3 of argv) of worksheet (item 2 of argv) of wb', ws, f'A{last}')
        want = row_look(ws, t, 1)
        check(hdr_look not in want, 'the remaining row does not look like the header to begin with')

        if c['Jobs']:
            wb.call('DeleteJobGroupAt', sheet(ws), table(t), 1, 'test')
        else:
            wb.call('DoRemoveCatalogRow', table(t), 1)
        check(wb.table_row_count(ws, t) == 1, 'removing the last row leaves one blank row, not zero')
        try:
            has_body = int(wb.prop(table(t), 'DataBodyRange.Rows.Count')) > 0
        except Exception:
            has_body = False
        check(has_body, 'the table still has a data body')
        check(int(wb.call('RowCount', table(t))) == 0, 'RowCount reports zero rows')
        check(row_look(ws, t, 1) == want, 'the blank row keeps its own formatting')

        # Add a row the old way (ListRows.Add) - it copies the row above.
        wb.call('UnlockSheet', sheet(ws))
        new_index = int(wb.prop(table(t), 'ListRows.Add.Index'))
        check(row_look(ws, t, new_index) == want, 'a row added afterwards is not dressed like the header')
        check(hdr_look not in row_look(ws, t, new_index), 'new row does not carry the header fill')

    wb.run('SetQuiet', False)

finish()
