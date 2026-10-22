# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-clearformulas.ps1 - same checks, same messages. The .ps1
# injects a MakeRoom() wrapper only to call CreatePrintRoom and get the new
# sheet's name; modScript.ScriptRun does the same without injecting anything
# (the returned Worksheet comes back as a handle, read with ScriptProp).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, finish, addr  # noqa: E402

with open_copy() as wb:
    wb.run('SetQuiet', True)
    ws = wb.prop(wb.call('CreatePrintRoom', 'Formula Room', 'Facilities'), 'Name')
    wb.run('SetQuiet', True)
    t = wb.table_names(ws)[0]
    row = wb.table_first_row(ws, t) + 1
    for h in ('Unit', 'Area m2', 'Paper Cost', 'Consumable Cost', 'Status'):
        check(wb.has_formula(ws, addr(row, wb.table_col(ws, t, h))), f"new room: '{h}' keeps its formula")
    check(wb.table_row_count(ws, t) == 1, 'new room keeps one blank row')

finish()
