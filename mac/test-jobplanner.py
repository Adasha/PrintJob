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

    section('a multi-pass printer')
    # Added here, not with the Annexe: everything above is the single-pass box.
    CWS, CTBL, PWS, PTBL = 'Consumables', 'tblColours', 'Printers', 'tblPrinters'
    wb.activate(CWS)
    for name, typ, rate in (('Black', 'Risograph', 0.30), ('Blue', 'Risograph', 0.40), ('Cyan', 'Ink', 0.10)):
        wb.run('AddCatalogRow', CTBL)
        crow = wb.table_last_row(CWS, CTBL)
        wb.call('UnlockSheet', sheet(CWS))
        wb.set_value(CWS, addr(crow, wb.table_col(CWS, CTBL, 'Colour')), name)
        wb.set_value(CWS, addr(crow, wb.table_col(CWS, CTBL, 'Consumable type')), typ)
        wb.set_value(CWS, addr(crow, wb.table_col(CWS, CTBL, 'Cost per m2')), rate)
    wb.activate(PWS)
    wb.run('AddCatalogRow', PTBL)
    prow = wb.table_last_row(PWS, PTBL)
    wb.call('UnlockSheet', sheet(PWS))
    for h, v in (('Model', 'Test RISO'), ('Consumable type', 'Risograph'), ('Colour mode', 'multi-pass'),
                 ('Template cost', 3.0), ('Max sheet size', 'A1'), ('Active', 'Yes')):
        wb.set_value(PWS, addr(prow, wb.table_col(PWS, PTBL, h)), v)
    wb.call('UnlockSheet', sheet(WS))
    loc = wb.name_addr(WS, 'LOC_Printers')
    wb.set_value(WS, loc, wb.text(WS, loc) + ';Test RISO')
    wb.run('RefreshSummaryPlanner')
    wb.activate(SUM)

    check('Test RISO' in list_items('B7'), 'the multi-pass printer is offered')
    wb.clear_contents(SUM, 'B7:B11')
    wb.set_value(SUM, 'B7', 'Test RISO')
    wb.set_value(SUM, 'B8', 'Gloss 200gsm SRA3 sheet')
    wb.set_value(SUM, 'B9', 100)
    recalc()
    check(t('A10') == 'Colours (one per pass)', f"row 10 asks for colours (got '{t('A10')}')")
    check(t('A15') == 'Ink and set-up', 'the ink line includes set-up')
    b_items, f_items = list_items('B10'), list_items('F10')
    check(b_items == ['Black', 'Blue'], f"first slot lists this printer's colours (got {b_items})")
    check(f_items == ['Black', 'Blue'], f"fifth slot lists the same colours (got {f_items})")
    check(t('C10') == '', 'the width hint is gone')
    check(t('A17') == 'Choose at least one colour.', f"no colour -> prompt (got '{t('A17')}')")
    check(t('B16') == '', 'no colour -> no estimate')

    wb.set_value(SUM, 'B19', room_name)
    before = wb.table_row_count(WS, TBL)
    wb.run('btnPlannerAdd')
    time.sleep(0.3)
    check(wb.table_row_count(WS, TBL) == before, 'no colour -> no row added')

    wb.set_value(SUM, 'B10', 'Cyan')
    recalc()
    check(t('A17') == "'Cyan' is not a colour Test RISO can use.", f"another type's colour -> message (got '{t('A17')}')")

    section('colours are priced one by one and can repeat')
    wb.set_value(SUM, 'B10', 'Black')
    wb.set_value(SUM, 'D10', 'Blue')          # a gap in the strip is not a pass
    wb.set_value(SUM, 'F10', 'Black')
    recalc()
    total, paper, ink_setup = wb.number(SUM, 'B16'), wb.number(SUM, 'B14'), wb.number(SUM, 'B15')
    check(total > 0, f'an estimate is produced ({total})')
    check(abs(paper + ink_setup - total) < 0.0001, f'paper + ink and set-up = overall ({paper} + {ink_setup})')
    check(t('C15').endswith('set-up (3 passes)'), f"the breakdown names the passes (got '{t('C15')}')")

    wb.run('btnPlannerAdd')
    time.sleep(0.5)
    after = wb.table_row_count(WS, TBL)
    check(after == before + 4, f'a job and three passes were added ({before} -> {after})')

    section('the added job and its passes')
    first_data = wb.table_first_row(WS, TBL)
    job = before + 1

    def cell(row_no, col):
        return wb.value(WS, addr(first_data + row_no, wb.table_col(WS, TBL, col)))

    gross = float(cell(job, 'Gross Cost') or 0)
    cons, setup = float(cell(job, 'Consumable Cost') or 0), float(cell(job, 'Set-up Cost') or 0)
    check(abs(gross - total) < 0.005, f"the job's Gross Cost ({gross}) equals the planner's Overall ({total})")
    check(abs(cons + setup - ink_setup) < 0.005, f"its ink and set-up ({cons} + {setup}) equal the planner's ({ink_setup})")
    check(abs(setup - 9) < 0.005, f'set-up is 3 passes at 3.00 ({setup})')
    check(int(float(cell(job, 'Passes') or 0)) == 3, 'Passes is 3')
    check(wb.row_outline_level(WS, first_data + job) == 1, 'the job row is outline level 1')
    job_id = cell(job, 'Job ID')
    for i, (colour, rate) in enumerate((('Black', 0.30), ('Blue', 0.40), ('Black', 0.30)), start=1):
        p = job + i
        check(cell(p, 'Row Type') == 'Pass' and cell(p, 'Parent') == job_id and int(float(cell(p, 'Pass'))) == i,
              f'pass {i} belongs to the job')
        check(cell(p, 'Colour') == colour and abs(float(cell(p, 'S_ConsRate') or 0) - rate) < 0.0001,
              f'pass {i} is {colour} at {rate}')
        check(wb.row_outline_level(WS, first_data + p) == 2, f'pass {i} is outline level 2')

    wb.activate(SUM)
    wb.clear_contents(SUM, 'D10')
    wb.clear_contents(SUM, 'F10')
    recalc()
    check(t('C15').endswith('set-up (1 pass)'), f"one colour reads as one pass (got '{t('C15')}')")
    before = wb.table_row_count(WS, TBL)
    wb.run('btnPlannerAdd')
    time.sleep(0.5)
    wb.activate(SUM)
    check(wb.table_row_count(WS, TBL) == before + 2, 'a second job with one pass was added')
    check(wb.row_outline_level(WS, first_data + before + 1) == 1, 'added below a pass, the job row is still level 1')

    section('row 10 goes back for a single-pass printer')
    wb.set_value(SUM, 'C10', 'Blue')
    wb.set_value(SUM, 'B7', prn)
    wb.set_value(SUM, 'B8', stk)
    wb.set_value(SUM, 'B9', 7)
    recalc()
    check(t('A10') == 'Print width mm', 'the Print width label is back')
    check(t('A15') == 'Ink (consumable)', 'the ink label is back')
    check(t('B10') == '', 'the colour in B10 is gone')
    check(t('C10') != 'Blue' and t('D10') == '' and t('F10') == '', 'and the other slots')
    check(t('C15') == '', 'no set-up breakdown')
    check(wb.number(SUM, 'B16') > 0, 'the estimate still works')
    wb.set_value(SUM, 'B7', 'Test RISO')
    recalc()
    check(t('B10') == '' and t('C10') == '', 'back on the multi-pass printer, the strip starts empty')
    wb.set_value(SUM, 'B8', 'Matte 180gsm 1067mm roll')
    wb.set_value(SUM, 'B10', 'Black')
    recalc()
    check('cannot take' in t('A17'), f"a roll paper on the multi-pass printer -> message (got '{t('A17')}')")
    check(t('B16') == '', 'and no estimate')

    wb.run('SetQuiet', False)

finish()
