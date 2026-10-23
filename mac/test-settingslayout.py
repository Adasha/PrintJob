# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-settingslayout.ps1 - same checks, same messages, same
# order; see that script for the button arrangement it guards.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, finish  # noqa: E402

SET = 'Settings'
GROUPS = {
    'Print rooms': ['btnAddPrintRoom', 'btnRenamePrintRoom', 'btnRemovePrintRoom', 'btnRefreshLocations'],
    'Data': ['btnExportAll', 'btnImportGlobal', 'btnBackupWorkbook', 'btnRestoreWorkbook'],
    'Workbook': ['btnCheckWorkbook', 'btnReStamp', 'btnAbout'],
}
SMALL = ['btnAddRowPaperTypes', 'btnRemoveRowPaperTypes', 'btnAddRowSize', 'btnRemoveRowSize',
         'btnAddRowConsumable', 'btnRemoveRowConsumable']


def find(buttons, macro):
    return [b for b in buttons if b['name'].startswith(f'pcb_{macro}_')]


def check_layout(wb, tag):
    buttons = wb.buttons(SET)
    hdr = wb.table_first_row(SET, 'tblSettings')
    check(hdr == 8, f'{tag}: tblSettings header is on row 8 (was {hdr})')
    for t in ('tblPaperTypes', 'tblStandardSizes', 'tblConsumables'):
        check(wb.table_first_row(SET, t) == 8, f'{tag}: {t} header is on row 8 too (tables stay aligned)')
    check(wb.text(SET, 'P2') != 'Commands', f"{tag}: the stale 'Commands' heading is gone")
    check(not wb.text(SET, 'P6').startswith('Buttons appear'), f'{tag}: the stale placeholder note is gone')

    limit = wb.column_lefts(SET, 5)[4]   # right edge of column D
    subtitle_top = wb.row_top_height(SET, 7)[0]
    row = 3
    for g, macros in GROUPS.items():
        check(wb.text(SET, f'A{row}') == g, f"{tag}: A{row} is the '{g}' label")
        row_top, row_h = wb.row_top_height(SET, row)
        row_bottom = row_top + row_h
        placed = []
        for m in macros:
            hit = find(buttons, m)
            check(len(hit) == 1, f'{tag}: exactly one {m}')
            if len(hit) == 1:
                b = hit[0]
                check(b['top'] >= row_top and (b['top'] + b['height']) <= row_bottom + 0.5, f'{tag}: {m} sits inside its group row ({row})')
                check((b['top'] + b['height']) <= subtitle_top, f'{tag}: {m} is above the table titles')
                placed.append(b)
        placed.sort(key=lambda b: b['left'])
        for i in range(1, len(placed)):
            gap = placed[i]['left'] - (placed[i - 1]['left'] + placed[i - 1]['width'])
            check(0 <= gap <= 10, f'{tag}: {g} buttons {i}/{i + 1} are clustered (gap {gap:.1f}pt)')
        if placed:
            last = placed[-1]
            check((last['left'] + last['width']) <= limit, f'{tag}: the {g} row ends inside columns A:D')
        row += 1
    row6_top = wb.row_top_height(SET, 6)[0]
    for m in SMALL:
        hit = find(buttons, m)
        check(len(hit) == 1, f'{tag}: exactly one {m}')
        if len(hit) == 1:
            check(hit[0]['top'] >= row6_top and (hit[0]['top'] + hit[0]['height']) <= subtitle_top + 2,
                  f"{tag}: {m} sits on row 6, just above its table's title")
    all_macros = [m for ms in GROUPS.values() for m in ms] + SMALL
    extra = len([b for b in buttons if b['name'].startswith('pcb_')]) - len(all_macros)
    check(extra == 0, f'{tag}: no stray buttons on Settings (extra: {extra})')


with open_copy() as wb:
    wb.run('SetQuiet', True)

    check_layout(wb, 'built')

    # A second setup run must not change the layout.
    wb.run('InitialiseWorkbook')
    time.sleep(0.5)
    check_layout(wb, 'after second setup')

finish()
