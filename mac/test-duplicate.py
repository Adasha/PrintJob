# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-duplicate.ps1 (AT-13) - the same steps and the same printed
# report (it reports rather than asserts): duplicate a print room sheet, run
# Refresh Locations, and show the job tables, the registry and the
# consolidated range. The .ps1 injects a TestDuplicate() function to do the
# copy in VBA; here the sheet is copied through AppleScript, as the Annexe
# fixture does, and the rest goes through modScript.ScriptRun, so nothing is
# injected.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, spill, ExcelError  # noqa: E402

SRC = 'Example Print Room'


def table_lines(wb, ws, t):
    heads = wb.table_headers(ws, t)
    cols = [wb.table_column_texts(ws, t, h) for h in heads]
    return [' | '.join(c[i] for c in cols) + ' | ' for i in range(len(cols[0]) if cols else 0)]


def table_list(wb):
    r = ''
    for ws in wb.sheet_names():
        for t in wb.table_names(ws):
            if t.startswith('tblJobs_'):
                code_cell = wb.name_addr(ws, 'LOC_Code')
                code = wb.text(ws, code_cell) if code_cell else ''
                r += f"   {t}  on '{ws}'  LOC_Code={code}\n"
    return r


with open_copy() as wb:
    r = ''
    try:
        r = 'job tables before:\n' + table_list(wb)
        new = wb._do('''set src to worksheet (item 2 of argv) of wb
set oldNames to name of every worksheet of wb
copy worksheet src after src
set newNames to name of every worksheet of wb
repeat with n in newNames
  if oldNames does not contain (n as text) then return (n as text)
end repeat
return ""''', SRC)
        r += f"duplicated '{SRC}' -> '{new}'\n"
        r += 'job tables straight after the copy:\n' + table_list(wb)

        wb.run('SetQuiet', True)
        wb.run('RefreshLocations')
        r += 'Refresh Locations said:\n' + wb.run('QuietLog')
        wb.run('SetQuiet', False)

        r += 'job tables after refresh:\n' + table_list(wb)
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
