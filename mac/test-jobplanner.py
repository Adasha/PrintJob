# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-jobplanner.ps1 - same checks, same messages, same order;
# see that script for the job planner box on Summary.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, sheet, SUM_HDR_ROW, ExcelError  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

SUM, WS, TBL = 'Summary', 'Example Print Room', 'tblJobs_EXAMPLE'

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def list_items(a1):
        """The items behind a list validation: a range, or a name for one."""
        f = wb.validation_formula1(SUM, a1).lstrip('=')
        if not f:
            return []
        if '!' not in f:
            try:
                f = wb._do('return references of named item (item 2 of argv) of wb', f).lstrip('=')
            except ExcelError:
                return []
        sh, _, ref = f.rpartition('!')
        sh = sh.strip("'").replace("''", "'")
        return [c for row in wb.texts(sh, ref.replace('$', '')) for c in row if c != '']

    def t(a1):
        return wb.text(SUM, a1)

    def recalc():
        wb.call('PlannerRecalc', sheet(SUM))

    def last(col):
        return wb.value(WS, addr(wb.table_last_row(WS, TBL), wb.table_col(WS, TBL, col)))

    section('layout', first=True)
    check(t('A5') == 'Plan a print job', 'box title in A5')
    check(t('A7') == 'Printer' and t('A8') == 'Paper', 'Printer and Paper labels')
    check(t('A23') == 'Totals', 'Totals label moved below the box (A23)')
    check(wb.cell_text(SUM, SUM_HDR_ROW, 1) == 'Location', f'Summary table header is on row {SUM_HDR_ROW}')
    caps = [b['text'] for b in wb.buttons(SUM)]
    check('Add to print room' in caps, 'Add to print room button exists')
    check('Refresh Locations' in caps, 'existing Summary buttons are still there')

    section('dropdowns')
    prn_items, stk_items, room_items = list_items('B7'), list_items('B8'), list_items('B19')
    check(len(prn_items) >= 1, f'printers offered ({len(prn_items)})')
    check(len(stk_items) >= 1, f'papers offered ({len(stk_items)})')
    check('Annexe' in room_items, 'room list includes the second room (Annexe)')
    check(len(room_items) >= 2, f'room list has every room ({len(room_items)})')

    section('incomplete input')
    wb.clear_contents(SUM, 'B7:B11')
    recalc()
    check(t('A17') == 'Choose a printer.', f"no printer -> prompt (got '{t('A17')}')")
    check(t('B16') == '', 'no estimate without inputs')

    # Use the printer and paper of the shipped sample job: known to be a valid pair.
    first = wb.table_first_row(WS, TBL) + 1
    prn = wb.cell_text(WS, first, wb.table_col(WS, TBL, 'Printer'))
    stk = wb.cell_text(WS, first, wb.table_col(WS, TBL, 'Paper Stock'))
    check(prn != '' and stk != '', f'sample job gives a printer/paper pair ({prn} / {stk})')

    wb.set_value(SUM, 'B7', prn)
    recalc()
    check(t('A17') == 'Choose a paper.', 'printer only -> asks for paper')
    wb.set_value(SUM, 'B8', stk)
    recalc()
    check(t('B16') == '', 'no amount -> no estimate')

    section('estimate equals the job row')
    is_sheet = t('A9') == 'Number of sheets'
    wb.set_value(SUM, 'B9', 7)
    recalc()
    total, paper, ink = wb.number(SUM, 'B16'), wb.number(SUM, 'B14'), wb.number(SUM, 'B15')
    check(total > 0, f'an estimate is produced ({total})')
    check(abs(paper + ink - total) < 0.0001, f'paper + ink = overall ({paper} + {ink})')

    room_name = wb.call('LocValue', sheet(WS), 'LOC_Name')
    wb.set_value(SUM, 'B19', room_name)
    before = wb.table_row_count(WS, TBL)
    wb.run('btnPlannerAdd')
    time.sleep(0.5)
    after = wb.table_row_count(WS, TBL)
    check(after == before + 1, f'a row was added ({before} -> {after})')

    gross = float(last('Gross Cost') or 0)
    check(abs(gross - total) < 0.005, f"new row Gross Cost ({gross}) equals the planner's Overall ({total})")

    section('the added row is prefilled and stamped')
    check(last('Printer') == prn, 'Printer filled in')
    check(last('Paper Stock') == stk, 'Paper Stock filled in')
    unit = wb.call('RollUnitOf', sheet(WS))
    want_qty = 700 if (not is_sheet and unit == 'Centimetres') else 7
    check(float(last('Qty') or 0) == want_qty, f'Qty is {want_qty} (room unit {unit})')
    check(last('Paid') == 'No', 'Paid is No')
    check(last('Job ID') not in (None, ''), 'Job ID allocated')
    check(last('S_StampedAt') not in (None, ''), 'row is stamped (S_StampedAt)')

    section('nothing added when it should not be')
    wb.clear_contents(SUM, 'B19')
    before = wb.table_row_count(WS, TBL)
    wb.run('btnPlannerAdd')
    time.sleep(0.3)
    check(wb.table_row_count(WS, TBL) == before, 'no room chosen -> no row added')

    wb.set_value(SUM, 'B19', room_name)
    wb.clear_contents(SUM, 'B9')
    recalc()
    wb.run('btnPlannerAdd')
    time.sleep(0.3)
    check(wb.table_row_count(WS, TBL) == before, 'incomplete plan -> no row added')

    wb.run('SetQuiet', False)

finish()
