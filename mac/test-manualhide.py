# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-manualhide.ps1 - same checks, same messages, same order;
# see that script for the Workbook_SheetSelectionChange self-heal it guards.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish  # noqa: E402

MAIN = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'
SIDE_PANEL = ['pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation']

with open_copy() as wb:
    wb.run('SetQuiet', True)

    def button_left(prefix):
        for b in wb.buttons(MAIN):
            if b['name'].startswith(prefix):
                return float(b['left'])
        return -1

    def button_column_visible(prefix):
        """First VISIBLE column whose Left matches the button (a hidden
        column shares its Left with the visible one after it)."""
        left = button_left(prefix)
        if left < 0:
            return 0
        lefts = wb.column_lefts(MAIN, 60)
        for c in range(1, 61):
            if abs(lefts[c - 1] - left) < 0.5 and not wb.column_hidden(MAIN, c):
                return c
        return -1

    def nudge():
        """Fire Workbook_SheetSelectionChange by moving the selection."""
        wb.select(MAIN, 'B3')
        time.sleep(0.2)
        wb.select(MAIN, 'B4')
        time.sleep(0.3)

    def col47():
        return wb.column_lefts(MAIN, 47)[46]

    wb.activate(MAIN)
    tech_col = wb.table_col(MAIN, TBL, 'Technician')

    section('Baseline', first=True)
    check(button_column_visible('pcb_btnClearDefaults') == tech_col, f'btnClearDefaults starts on Technician (col {tech_col})')
    col36_before = col47()
    for p in SIDE_PANEL:
        check(button_left(p) == col36_before, f'{p} on column 47 (Left={col36_before})')

    section('Manually hide Technician, stay on the sheet, click a cell')
    wb.set_column_hidden(MAIN, tech_col, True)
    nudge()
    moved = button_column_visible('pcb_btnClearDefaults')
    check(moved > 0 and moved != tech_col, f'btnClearDefaults left the hidden Technician column (now col {moved})')
    check(moved > 0 and not wb.column_hidden(MAIN, moved), 'its new column is visible')

    section('Reveal Technician, click a cell')
    wb.set_column_hidden(MAIN, tech_col, False)
    nudge()
    check(button_column_visible('pcb_btnClearDefaults') == tech_col, 'btnClearDefaults back on Technician')

    section('Manually hide a column left of 47, click a cell: side panel follows')
    hide_col = wb.table_col(MAIN, TBL, 'Printer')
    wb.set_column_hidden(MAIN, hide_col, True)
    nudge()
    col36_hidden = col47()
    check(col36_hidden != col36_before, f'hiding column {hide_col} moved column 47 (was {col36_before}, now {col36_hidden})')
    for p in SIDE_PANEL:
        check(button_left(p) == col36_hidden, f'{p} followed column 47 (Left={col36_hidden})')
    wb.set_column_hidden(MAIN, hide_col, False)
    nudge()
    for p in SIDE_PANEL:
        check(button_left(p) == col36_before, f'{p} back on column 47 after reveal')

    wb.run('SetQuiet', False)

finish()
