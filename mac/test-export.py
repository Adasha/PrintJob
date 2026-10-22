# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-export.ps1 - the same steps and the same printed report:
# export a print room, read the CSV back, edit a row and show the fingerprint
# notice. Like the .ps1 it fails only when no CSV is written; the rest is for
# reading.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import csv
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, addr, sheet  # noqa: E402

WS, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'

ctx = open_copy()
with ctx as wb:
    wb.run('SetQuiet', True)
    export_cell = wb.name_addr(WS, 'LOC_Export')

    print('=== status before export ===')
    print('  Export: ' + wb.text(WS, export_cell))

    wb.activate(WS)
    wb.run('btnExport')
    print()
    print('=== what the export reported ===')
    print(wb.run('QuietLog'))

    print('=== status after export ===')
    print('  Export: ' + wb.text(WS, export_cell))

    # Edit a quantity. The fingerprint must notice.
    cell = addr(wb.table_first_row(WS, TBL) + 1, wb.table_col(WS, TBL, 'Qty'))
    was = wb.number(WS, cell)
    wb.unprotect(WS)
    wb.set_value(WS, cell, was + 7)
    wb.calculate_full_rebuild()
    wb.call('RefreshExportStatus', sheet(WS))
    print()
    print(f'=== after changing a quantity ({was:g} -> {wb.number(WS, cell):g}) ===')
    print('  Export: ' + wb.text(WS, export_cell))

    wb.run('SetQuiet', False)

print()
print('=== the CSV ===')
found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-*.csv')), key=os.path.getmtime)
if not found:
    print('  NO CSV WRITTEN')
    sys.exit(1)
path = found[-1]
print(f'  {os.path.basename(path)}  ({os.path.getsize(path)} bytes)')
print()
with open(path, encoding='utf-8-sig') as f:
    lines = f.read().splitlines()
print('  --- header block ---')
for ln in lines[0:8]:
    print('  ' + ln)
print('  --- column headers ---')
print('  ' + lines[8])
print('  --- first record ---')
print('  ' + lines[9])
print()
print(f'  total lines: {len(lines)}')
hdrs = next(csv.reader([lines[8]]))
print(f'  columns: {len(hdrs)}')
for need in ('S_UnitCost', 'S_ConsRate', 'S_StockWidth_mm', 'S_SheetHeight_mm', 'S_SchemaVer'):
    print(f"  snapshot column {need:<18} {'present' if need in hdrs else 'MISSING'}")
