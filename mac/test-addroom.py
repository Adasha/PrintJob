# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-addroom.ps1 - the same steps and the same printed report
# (it reports rather than asserts). The .ps1 injects a TestAddRoom() function
# that strings these calls together; here they are made one at a time through
# modScript.ScriptRun, with each new room's Worksheet held as a handle, so
# nothing is injected. CreatePrintRoom is called directly, not AddPrintRoom,
# so there is no InputBox to answer.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, sheet, spill, ExcelError  # noqa: E402

SRC = 'Example Print Room'


def table_lines(wb, ws, t):
    heads = wb.table_headers(ws, t)
    cols = [wb.table_column_texts(ws, t, h) for h in heads]
    return [' | '.join(c[i] for c in cols) + ' | ' for i in range(len(cols[0]) if cols else 0)]


with open_copy() as wb:
    r = ''
    try:
        def lv(ws, name):
            return wb.call('LocValue', ws, name)

        def rows(ws_token):
            return wb.prop(wb.call('JobsTable', ws_token), 'ListRows.Count')

        src = sheet(SRC)
        # Give the template some non-default state first, so clearing it is
        # actually exercised rather than trivially true.
        wb.call('UnlockSheet', src)
        wb.set_value(SRC, wb.name_addr(SRC, 'LOC_DefDisPaper'), 'Yes')
        wb.set_value(SRC, wb.name_addr(SRC, 'LOC_RollUnit'), 'Centimetres')
        wb.call('RelockSheet', src)

        r = (f"template before: Rows={rows(src)} DefDisPaper={lv(src, 'LOC_DefDisPaper')} "
             f"RollUnit={lv(src, 'LOC_RollUnit')} Printers={lv(src, 'LOC_Printers')}\n")

        wb.run('SetQuiet', True)
        ws = wb.call('CreatePrintRoom', 'Test Room #1 / Annex:2', 'Facilities')
        r += 'Refresh Locations said:\n' + wb.run('QuietLog')
        wb.run('SetQuiet', False)
        if not ws.startswith('@@h:'):
            print(r + 'FAIL: CreatePrintRoom returned Nothing')
            sys.exit(1)

        r += f"new sheet name: '{wb.prop(ws, 'Name')}'\n"
        r += (f"new sheet: Rows={rows(ws)} LOC_Name={lv(ws, 'LOC_Name')} LOC_Dept={lv(ws, 'LOC_Dept')} "
              f"LOC_Code={lv(ws, 'LOC_Code')} DefTech=[{lv(ws, 'LOC_DefTech')}] "
              f"DefDisPaper=[{lv(ws, 'LOC_DefDisPaper')}] RollUnit={lv(ws, 'LOC_RollUnit')} "
              f"Printers=[{lv(ws, 'LOC_Printers')}]\n")
        r += f"template unchanged: Rows={rows(src)} LOC_Code={lv(src, 'LOC_Code')}\n"

        # A second room with an explicit, deliberately messy code - see the .ps1.
        wb.run('SetQuiet', True)
        ws2 = wb.call('CreatePrintRoom', 'Photo Lab', 'Media Services', 'photo lab 2026!!')
        wb.run('SetQuiet', False)
        r += f"second room: LOC_Name={lv(ws2, 'LOC_Name')} LOC_Code={lv(ws2, 'LOC_Code')}\n"

        reg_ws = next((s for s in wb.sheet_names() if 'tblLocations' in wb.table_names(s)), None)
        r += 'registry:\n' + (''.join(f'   {ln}\n' for ln in table_lines(wb, reg_ws, 'tblLocations'))
                              if reg_ws else '   (no tblLocations)\n')
        try:
            d = spill('_Data', 'A10')
            r += f"consolidated: {wb.prop(d, 'Rows.Count')} rows x {wb.prop(d, 'Columns.Count')} cols\n"
        except ExcelError:
            r += 'consolidated: NO SPILL\n'
    except ExcelError as e:
        r += f'ERROR {e}\n'
    print(r)
