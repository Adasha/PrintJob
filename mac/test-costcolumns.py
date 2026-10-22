# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-costcolumns.ps1 - same checks, same messages, same order;
# see that script for the "button lag" fix it guards. Like the .ps1, nothing
# waits between the toggle and the checks that matter: the repositioning has
# to have happened inside the same macro call.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
ANNEX, ATBL = 'Annexe', 'tblJobs_ANNEX'
SIDE_PANEL = ['pcb_btnSelectPrinters', 'pcb_btnCheckSheet', 'pcb_btnClearAll', 'pcb_btnExport', 'pcb_btnImportLocation']
SPAN = ['Paper Cost', 'Consumable Cost', 'Gross Cost', 'Disregarded']

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def is_hidden(sheet, table, header):
        return wb.column_hidden(sheet, wb.table_col(sheet, table, header))

    def button(sheet, prefix):
        for b in wb.buttons(sheet):
            if b['name'].startswith(prefix):
                return b
        return None

    def caption(sheet, prefix):
        b = button(sheet, prefix)
        return b['text'] if b else None

    def side_panel_left(prefix):
        b = button(MAIN, prefix)
        return float(b['left']) if b else -1

    def col47():
        return wb.column_lefts(MAIN, 47)[46]

    # -------------------------------------------------------------- setting
    section('Setting self-provisioned', first=True)
    on_off = wb.wb_name_text('SET_COST_COLS_HIDDEN')
    check(on_off == 'No', f"SET_COST_COLS_HIDDEN defaults to No (got '{on_off}')")

    # ---------------------------------------------------------- initial state
    section('Initially visible: nothing in the cost-detail span is hidden')
    for h in SPAN:
        check(not is_hidden(MAIN, TBL, h), f'{h} is visible on Example Print Room')
    check(not is_hidden(MAIN, TBL, 'Chargeable Cost'), 'Chargeable Cost is visible')
    check(not is_hidden(MAIN, TBL, 'Paid'), 'Paid is visible')

    cost_btn = button(MAIN, 'pcb_btnToggleCostColumns')
    check(cost_btn is not None, 'cost toggle exists')
    check(cost_btn is not None and cost_btn['top'] < wb.row_top_height(MAIN, 3)[0],
          'cost toggle is on row 2 (view row), not in the side panel')

    col36_before = col47()
    for p in SIDE_PANEL:
        check(side_panel_left(p) == col36_before, f'{p} sits on column 47 before hiding (Left={col36_before})')

    # -------------------------------------------------------------- toggle on
    section('Toggling ON hides the cost-detail span on every location sheet')
    wb.run('btnToggleCostColumns')
    for h in SPAN:
        check(is_hidden(MAIN, TBL, h), f'{h} is hidden on Example Print Room')
        check(is_hidden(ANNEX, ATBL, h), f'{h} is hidden on Annexe too (workbook-wide, not per-sheet)')
    check(not is_hidden(MAIN, TBL, 'Chargeable Cost'), 'Chargeable Cost stays visible (never hidden by this toggle)')
    check(not is_hidden(MAIN, TBL, 'Paid'), 'Paid stays visible (never hidden by this toggle)')

    on_off_after = wb.wb_name_text('SET_COST_COLS_HIDDEN')
    check(on_off_after == 'Yes', 'SET_COST_COLS_HIDDEN now reads Yes')

    btn_caption = caption(MAIN, 'pcb_btnToggleCostColumns')
    check(btn_caption == 'Show cost detail', f"button caption flipped to 'Show cost detail' (got '{btn_caption}')")
    annex_caption = caption(ANNEX, 'pcb_btnToggleCostColumns')
    check(annex_caption == 'Show cost detail', f"Annexe's toggle relabelled too (got '{annex_caption}')")

    # ------------------------------------------ THE FIX: no lag, same click
    section('No lag: side panel already repositioned in the SAME click that hid the columns')
    col36_hidden = col47()
    check(col36_hidden != col36_before,
          f'hiding the span actually moved column 47 (was {col36_before}, now {col36_hidden}) - otherwise this check proves nothing')
    for p in SIDE_PANEL:
        check(side_panel_left(p) == col36_hidden,
              f"{p} already sits on column 47's NEW position, no separate click needed (Left={col36_hidden})")

    # ------------------------------------------------------------- toggle off
    section('Toggling OFF shows the span again, same click, no lag')
    wb.run('btnToggleCostColumns')
    for h in SPAN:
        check(not is_hidden(MAIN, TBL, h), f'{h} is visible again on Example Print Room')
    btn_caption2 = caption(MAIN, 'pcb_btnToggleCostColumns')
    check(btn_caption2 == 'Hide cost detail', f"button caption flipped back to 'Hide cost detail' (got '{btn_caption2}')")
    for p in SIDE_PANEL:
        check(side_panel_left(p) == col36_before,
              f"{p} is back on column 47's original position, same click (Left={col36_before})")

    wb.run('SetQuiet', False)

finish()
