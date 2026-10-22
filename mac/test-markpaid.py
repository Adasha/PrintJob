# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-markpaid.ps1 - same checks, same messages, same order;
# see that script for the "Mark all as..." cluster and the filter safeguard.
# The Worksheet and spilled Range the Confirmed routines take go in as
# modScript tokens (the spill token is resolved at each call, as the .ps1
# re-reads VisibleInfo); the caption label is read through a shape token.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself.

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, col_letters, sheet, spill, shape  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

REP, MAIN, ANNEX = 'Reports', 'Example Print Room', 'Annexe'
ROOMS = [(MAIN, 'tblJobs_EXAMPLE'), (ANNEX, 'tblJobs_ANNEX')]
FILTER_CELLS = ['B4', 'B5', 'B8', 'B9', 'B11', 'B12', 'B13', 'E4', 'E5', 'E7', 'E10', 'E11', 'E12', 'B15', 'E15', 'I7', 'K7']
CLEARABLE = ['B4', 'B5', 'B8', 'B9', 'B11', 'B12', 'B13', 'E4', 'E5', 'E7', 'E10', 'E11', 'E12', 'B15', 'E15']
OA_2020_01_01 = 43831.0   # (Get-Date '2020-01-01').ToOADate()

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)
    wb.activate(REP)

    def snapshot():
        h = {}
        for ws, t in ROOMS:
            for i, p in zip(wb.table_column_texts(ws, t, 'Job ID'), wb.table_column_texts(ws, t, 'Paid')):
                h[f'{t}|{i}'] = p.strip()
        return h

    def total_rows():
        return sum(wb.table_row_count(ws, t) for ws, t in ROOMS)

    def visible_info():
        heads = wb.texts(REP, 'H10:AB10')[0]
        loc_col = max((8 + i for i, h in enumerate(heads) if h == 'Location'), default=0)
        job_col = max((8 + i for i, h in enumerate(heads) if h == 'Job ID'), default=0)
        n = int(wb.prop(spill(REP, 'H11'), 'Rows.Count'))
        locs = [r[0] for r in wb.texts(REP, f'{col_letters(loc_col)}11:{col_letters(loc_col)}{10 + n}')]
        jobs = [r[0] for r in wb.texts(REP, f'{col_letters(job_col)}11:{col_letters(job_col)}{10 + n}')]
        return {'LocCol': loc_col, 'JobCol': job_col, 'N': n,
                'VKeys': [f'tblJobs_{lc}|{jb}' for lc, jb in zip(locs, jobs)]}

    def filter_state():
        return ';'.join(f'{c}={wb.text(REP, c)}' for c in FILTER_CELLS)

    def clear_filters():
        for a in CLEARABLE:
            wb.clear_contents(REP, a)
        wb.calculate_full_rebuild()

    def has_filter():
        return wb.call('HasActiveFilter', sheet(REP)) == 'True'

    def mark(vis, value):
        wb.call('MarkVisibleReportsConfirmed', sheet(REP), spill(REP, 'H11'), vis['LocCol'], vis['JobCol'], vis['N'], value)

    def delete(vis):
        wb.call('DeleteVisibleReportsConfirmed', sheet(REP), spill(REP, 'H11'), vis['LocCol'], vis['JobCol'], vis['N'])

    def log():
        return wb.run('QuietLog')

    # ------------------------------------------------------------ the cluster
    section('The cluster on the Reports header', first=True)
    lbl = shape(REP, 'pcb_lblMarkAll')
    try:
        cap = wb.prop(lbl, 'Caption')
        has_lbl = True
    except Exception:
        has_lbl = False
    check(has_lbl, "the 'Mark all as...' caption exists")
    if has_lbl:
        check(cap == 'Mark all as...', f"caption reads 'Mark all as...' (got '{cap}')")
    check(wb.text(REP, 'O1') == '', 'O1 holds no caption text (the caption is a label now)')
    buttons = wb.buttons(REP)
    paid_btn = next((b for b in buttons if b['name'].startswith('pcb_btnMarkPaid_')), None)
    unpaid_btn = next((b for b in buttons if b['name'].startswith('pcb_btnMarkUnpaid_')), None)
    check(paid_btn is not None, 'Paid button exists')
    check(unpaid_btn is not None, 'Unpaid button exists')
    if paid_btn and unpaid_btn:
        check(paid_btn['text'] == 'Paid', "Paid button caption is 'Paid'")
        check(unpaid_btn['text'] == 'Unpaid', "Unpaid button caption is 'Unpaid'")
        check(paid_btn['on_action'].endswith('btnMarkPaid'), f"Paid button runs btnMarkPaid (got '{paid_btn['on_action']}')")
        check(unpaid_btn['on_action'].endswith('btnMarkUnpaid'), f"Unpaid button runs btnMarkUnpaid (got '{unpaid_btn['on_action']}')")
        t4, t5, t6 = (wb.row_top_height(REP, r)[0] for r in (4, 5, 6))
        check(paid_btn['top'] >= t4 and paid_btn['top'] + paid_btn['height'] <= t5, 'Paid button sits inside row 4')
        check(unpaid_btn['top'] >= t5 and unpaid_btn['top'] + unpaid_btn['height'] <= t6, 'Unpaid button sits inside row 5')
        check(abs(paid_btn['left'] - unpaid_btn['left']) < 1, 'Paid and Unpaid are left-aligned with each other')
        if has_lbl:
            l_left, l_top, l_h = (float(wb.prop(lbl, p)) for p in ('Left', 'Top', 'Height'))
            check(abs(l_left - paid_btn['left']) < 1.5, 'the caption is left-aligned over the buttons')
            check(l_top + l_h <= paid_btn['top'] + 1, 'the caption sits above the Paid button')
        check(paid_btn['left'] <= 900 and unpaid_btn['left'] <= 900,
              f"both buttons are within the first screenful (Left <= 900pt; Paid at {int(paid_btn['left'])})")
        check(paid_btn['top'] < unpaid_btn['top'], 'Paid is the top button, Unpaid the bottom')

    # ------------------------------------------------- HasActiveFilter rules
    section('HasActiveFilter mirrors Criteria()')
    clear_filters()
    check(not has_filter(), 'no filter set -> false')
    wb.set_value(REP, 'B4', '   ')
    check(not has_filter(), 'spaces only in the name box -> false')
    wb.clear_contents(REP, 'B4')
    wb.set_value(REP, 'B8', 'not a date')
    check(not has_filter(), 'From date text that will not coerce (ignored by Criteria) -> false')
    wb.clear_contents(REP, 'B8')
    wb.set_value(REP, 'I7', 'Qty')
    check(not has_filter(), 'Sort by set, no filter -> false (sort is not a filter)')
    wb.clear_contents(REP, 'I7')
    wb.set_value(REP, 'B8', OA_2020_01_01)
    check(has_filter(), 'a real From date -> true')
    wb.clear_contents(REP, 'B8')
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    check(has_filter(), 'Printer set -> true')
    clear_filters()
    wb.set_value(REP, 'B11', 'abc')
    check(not has_filter(), 'Min cost text that is not a number -> false')
    clear_filters()
    wb.set_value(REP, 'B11', 1)
    check(has_filter(), 'a numeric Min cost -> true')
    clear_filters()
    wb.set_value(REP, 'B12', 100)
    check(has_filter(), 'a numeric Max cost -> true')
    clear_filters()
    wb.set_value(REP, 'E15', 'Yes')
    check(has_filter(), 'Has notes set -> true')
    clear_filters()

    # ---------------------------------------------- refused with no filter
    section('No filter: everything refuses and changes nothing')
    wb.calculate_full_rebuild()
    wb.activate(REP)
    everything = visible_info()
    print(f"  unfiltered sheet shows {everything['N']} records")
    before = snapshot()
    rows_before = total_rows()

    wb.run('SetQuiet', True)
    mark(everything, 'Yes')
    check('Set at least one filter before you mark records as Paid' in log(), 'Mark Paid said to set a filter first')
    wb.run('SetQuiet', True)
    wb.run('MarkVisibleReports', 'No')
    lg = log()
    check('Set at least one filter before you mark records as Unpaid' in lg, 'interactive Mark Unpaid said to set a filter first')
    check('DECLINED' not in lg, 'interactive path refused BEFORE reaching the confirmation prompt')
    wb.run('SetQuiet', True)
    delete(everything)
    check('Set at least one filter before you delete records' in log(), 'Delete visible said to set a filter first')
    wb.run('SetQuiet', True)
    wb.run('DeleteVisibleReports')
    lg = log()
    check('Set at least one filter before you delete records' in lg and 'DECLINED' not in lg,
          'interactive Delete refused BEFORE its confirmation prompt')

    after_refused = snapshot()
    rows_after = total_rows()
    check(rows_after == rows_before, f'no rows deleted ({rows_before} -> {rows_after})')
    check(sum(1 for k in before if before[k] != after_refused.get(k)) == 0, 'no Paid value changed')

    # ------------------------------------------------- filtered: Paid = Yes
    section('With a filter: only the visible rows change')
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    wb.calculate_full_rebuild()
    vis = visible_info()
    print(f"  filtered to Printer = 'Epson SureColor P9500': {vis['N']} of {len(before)} records visible")
    if vis['N'] == 0:
        print('FAIL: expected at least one visible record')
        sys.exit(1)
    if vis['N'] >= len(before):
        print('FAIL: the filter must leave some records hidden for this test to mean anything')
        sys.exit(1)
    vis_set = set(vis['VKeys'])
    filters_before = filter_state()

    # Start from a known state: every visible row Unpaid so a stray write is visible.
    wb.run('SetQuiet', True)
    mark(vis, 'No')
    baseline = snapshot()
    check(all(baseline.get(k) == 'No' for k in vis['VKeys']), 'visible rows now all read No (baseline)')

    wb.run('SetQuiet', True)
    vis = visible_info()
    mark(vis, 'Yes')
    lg = log()
    print(lg)
    after = snapshot()
    check(all(after.get(k) == 'Yes' for k in vis['VKeys']), f"all {vis['N']} visible rows are now Yes")
    check(sum(1 for k in after if k not in vis_set and after[k] != baseline.get(k)) == 0, 'no hidden row was touched')
    check(filter_state() == filters_before, 'filter and sort boxes are exactly as they were')
    check(wb.text(REP, 'E7') == 'Epson SureColor P9500', 'Printer filter still set')
    check('marked Paid' in lg, 'confirmation message names Paid')
    check(wb.protected(MAIN) and wb.protected(ANNEX), 'room sheets are protected again afterwards')

    # ---------------------------------------------- filtered: Paid -> Unpaid
    section('Unpaid, and already-at-target rows are left alone')
    wb.calculate_full_rebuild()
    vis = visible_info()
    wb.run('SetQuiet', True)
    mark(vis, 'No')
    after2 = snapshot()
    check(all(after2.get(k) == 'No' for k in vis['VKeys']), 'all visible rows are now No')
    check(sum(1 for k in after2 if k not in vis_set and after2[k] != baseline.get(k)) == 0, 'no hidden row was touched')
    wb.run('SetQuiet', True)
    mark(vis, 'No')
    lg = log()
    check(re.search(r'0 records marked Unpaid.*already (were|was) Unpaid', lg, re.S) is not None,
          'a repeat run reports 0 changed, all already Unpaid')

    # ---------------------------------------------- interactive still asks
    section('The interactive path still confirms first')
    wb.run('SetQuiet', True)
    wb.run('MarkVisibleReports', 'Yes')
    lg = log()
    check(re.search(r'DECLINED \(quiet mode\).*Mark visible records as Paid', lg, re.S) is not None, 'Ask() was reached for Paid')
    check(re.search(rf"Mark {vis['N']} visible record.*as Paid\?", lg, re.S) is not None, 'prompt states the record count')
    after3 = snapshot()
    check(sum(1 for k in after3 if after3[k] != after2.get(k)) == 0, 'declining changed nothing')

    # ----------------------------------------------------------- audit log
    section('Audit log')
    actions = wb.table_column_texts('_Audit', 'tblAudit', wb.table_headers('_Audit', 'tblAudit')[2])
    check('Mark visible Paid (Reports)' in actions, "audit has 'Mark visible Paid (Reports)'")
    check('Mark visible Unpaid (Reports)' in actions, "audit has 'Mark visible Unpaid (Reports)'")

    # ------------------------------------- delete still works once filtered
    section('Delete visible still works with a filter set')
    wb.calculate_full_rebuild()
    vis = visible_info()
    rows_before = total_rows()
    wb.run('SetQuiet', True)
    delete(vis)
    rows_after = total_rows()
    check(rows_before - rows_after == vis['N'], f"exactly {vis['N']} rows deleted ({rows_before} -> {rows_after})")

    wb.run('SetQuiet', False)

finish()
