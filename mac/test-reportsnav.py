# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-reportsnav.ps1 - same checks, same messages, same order;
# see that script for the 0.10.25 Reports additions it guards. Differences:
#
#   - "Buttons only" is the sheet's Form Controls buttons (xlmac.buttons), the
#     set the .ps1 picks out of Shapes by Type 8 / FormControlType 0.
#   - The two captions are Form Controls labels, found through a shape token.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import fnmatch
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, addr, col_letters, sheet, spill, shape, ExcelError  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

HDR, FIRST = 10, 11   # modReports.REP_HDR_ROW, REP_FIRST_ROW
REP, MAIN, ANNEX = 'Reports', 'Example Print Room', 'Annexe'
ROOMS = [(MAIN, 'tblJobs_EXAMPLE'), (ANNEX, 'tblJobs_ANNEX')]
BOXES = ['B4', 'B5', 'B8', 'B9', 'B11', 'B12', 'B13', 'E4', 'E5', 'E7', 'E10', 'E11', 'E12', 'B15', 'E15']
OA_2020_01_01 = 43831.0

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.calculate_full_rebuild()
    time.sleep(3)
    wb.run('SetQuiet', True)
    wb.activate(REP)

    def rep_col(pattern):
        heads = wb.texts(REP, f'H{HDR}:AV{HDR}')[0]   # columns 8..48
        return next((8 + i for i, h in enumerate(heads) if fnmatch.fnmatchcase(h, pattern)), 0)

    def jobs():
        wb.calculate_full_rebuild()
        return int(wb.text(REP, 'I9'))

    def spill_ok():
        return not wb.cell_text(REP, FIRST, 8).startswith('#SPILL')

    def clear_all_boxes():
        for a in BOXES + ['I7', 'K7']:
            wb.clear_contents(REP, a)
        wb.calculate_full_rebuild()

    def spill_rows():
        return int(wb.prop(spill(REP, f'H{FIRST}'), 'Rows.Count'))

    def spill_col(col, kind='text'):
        n = spill_rows()
        c = col_letters(col)
        if kind == 'text':
            return [r[0] for r in wb.texts(REP, f'{c}{FIRST}:{c}{FIRST + n - 1}')]
        return [r[0] for r in wb.values(REP, f'{c}{FIRST}:{c}{FIRST + n - 1}')]

    def find_shape(macro):
        return next((b for b in wb.buttons(REP) if b['name'].startswith(f'pcb_{macro}_')), None)

    def label(name):
        try:
            return {p: float(wb.prop(shape(REP, name), p)) for p in ('Left', 'Top', 'Height')}
        except ExcelError:
            return None

    def has_filter():
        return wb.call('HasActiveFilter', sheet(REP)) == 'True'

    def body(ws, table, header):
        c = col_letters(wb.table_col(ws, table, header))
        return f'{c}{wb.table_first_row(ws, table) + 1}:{c}{wb.table_last_row(ws, table)}'

    total = jobs()
    check(total >= 5, f'unfiltered Reports shows {total} records')

    # ------------------------------------------------------------ Has notes
    section('Has notes filter')
    # The sample data already carries notes on most rows, so start from none.
    for ws, t in ROOMS:
        wb.clear_contents(ws, body(ws, t, 'Notes'))
    first_main = wb.table_first_row(MAIN, 'tblJobs_EXAMPLE') + 1
    notes_col = wb.table_col(MAIN, 'tblJobs_EXAMPLE', 'Notes')
    wb.set_value(MAIN, addr(first_main, notes_col), 'check this one')
    wb.set_value(MAIN, addr(first_main + 1, notes_col), '   ')   # spaces only: counts as no notes
    total = jobs()
    wb.set_value(REP, 'E15', 'Yes')
    check(jobs() == 1, 'Has notes = Yes shows only the record with real notes (spaces-only is not a note)')
    wb.set_value(REP, 'E15', 'No')
    check(jobs() == total - 1, 'Has notes = No shows every other record')
    wb.clear_contents(REP, 'E15')
    check(jobs() == total, 'Has notes blank shows everything')
    rep_notes = rep_col('Notes')
    if rep_notes > 0:
        check(sum(1 for v in spill_col(rep_notes) if v == '0') == 0, 'no blank note is displayed as 0')

    # ---------------------------------------------------------- cost range
    section('Min / Max chargeable cost')
    cc = rep_col('Chargeable*')
    check(cc > 0, f'found the Chargeable cost column ({cc})')
    vals = [round(float(v or 0), 4) for v in spill_col(cc, 'value')]
    srt = sorted(vals)
    mid = round(srt[int(len(srt) / 2)], 4)
    wb.set_value(REP, 'B11', mid)
    exp_min = sum(1 for v in vals if v >= mid)
    check(jobs() == exp_min, f'Min cost {mid} shows {exp_min} records (>=, inclusive)')
    wb.clear_contents(REP, 'B11')
    wb.set_value(REP, 'B12', mid)
    exp_max = sum(1 for v in vals if v <= mid)
    check(jobs() == exp_max, f'Max cost {mid} shows {exp_max} records (<=, inclusive)')
    wb.set_value(REP, 'B11', mid)
    exp_both = sum(1 for v in vals if v == mid)
    check(jobs() == exp_both, f'Min = Max = {mid} shows the {exp_both} records at exactly that cost')
    check(has_filter(), 'a cost range counts as a filter')
    clear_all_boxes()
    wb.set_value(REP, 'B11', 'abc')
    check(jobs() == total, 'non-numeric Min cost is ignored')
    check(not has_filter(), 'non-numeric Min cost is not a filter')
    clear_all_boxes()

    # -------------------------------------------------------- clear filters
    section('Clear all filters')
    wb.set_value(REP, 'B4', 'Patel')
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    wb.set_value(REP, 'B11', 1)
    wb.set_value(REP, 'B12', 9999)
    wb.set_value(REP, 'E15', 'No')
    wb.set_value(REP, 'B13', 'No')
    wb.set_value(REP, 'B8', OA_2020_01_01)
    wb.set_value(REP, 'I7', 'Qty')
    wb.set_value(REP, 'K7', 'Descending')
    wb.run('btnClearFilters')
    left = [a for a in BOXES if wb.formula(REP, a) != '']
    check(not left, f"every filter box is empty afterwards (still set: {', '.join(left)})")
    check(wb.text(REP, 'I7') == 'Qty' and wb.text(REP, 'K7') == 'Descending', 'Sort by / Sort direction are left alone')
    check(jobs() == total, 'all records are back')
    check(wb.validation_type(REP, 'E7') == 'validate list', 'a cleared box keeps its dropdown')
    clear_all_boxes()

    # ---------------------------------------------------------- go to record
    section('Go to record')
    loc_col, job_col = rep_col('Location'), rep_col('Job ID')
    codes = spill_col(loc_col)
    uniq = list(dict.fromkeys(codes))
    check(len(uniq) >= 2, f"results span both rooms ({', '.join(uniq)})")
    for pick in (codes[0], next((c for c in codes if c != codes[0]), None)):
        if not pick:
            continue
        row = FIRST + codes.index(pick)
        want_job = wb.cell_text(REP, row, job_col).strip()
        wb.activate(REP)
        wb.select(REP, addr(row, 8))
        wb.run('btnGoToRecord')
        act = wb.active_sheet()
        table = f'tblJobs_{pick}'
        found = table in wb.table_names(act)
        check(found, f"room {pick} : went to its own sheet ('{act}')")
        if found:
            s_row, s_col, s_rows = wb.selection()
            job_idx = wb.table_headers(act, table).index('Job ID') + 1
            got_job = wb.cell_text(act, s_row, s_col + job_idx - 1).strip()
            check(got_job == want_job, f"room {pick} : selected the record {want_job} (got '{got_job}')")
            check(s_rows == 1, f'room {pick} : the whole table row is selected')
            v_row, v_n = wb.visible_rows()
            check(v_row <= s_row <= v_row + v_n - 1, f'room {pick} : the row is in view')
    # Nothing sensible selected: stays put and says so (quiet mode logs it).
    wb.activate(REP)
    wb.run('SetQuiet', True)
    wb.select(REP, 'A5')
    wb.run('btnGoToRecord')
    check(wb.active_sheet() == 'Reports', 'a cell above the results does not leave the Reports sheet')
    check('Select a record' in wb.run('QuietLog'), 'and asks for a record to be selected')

    # ------------------------------------------------------- button placement
    section('Reports buttons stay in place as columns change')
    names = ['btnClearFilters', 'btnTogglePaid', 'btnGoToRecord', 'btnMarkPaid', 'btnMarkUnpaid', 'btnExportReport', 'btnDeleteVisible']
    for n in names:
        check(find_shape(n) is not None, f'button {n} exists')
    check(label('pcb_lblSelected') is not None, "the 'Selected record:' caption exists")
    check(label('pcb_lblMarkAll') is not None, "the 'Mark all as...' caption exists")
    before = {n: find_shape(n)['left'] for n in names}
    lbl_before = {n: label(n)['Left'] for n in ('pcb_lblSelected', 'pcb_lblMarkAll')}

    def overlaps():
        a, b, m, e = (find_shape(x) for x in ('btnClearFilters', 'btnTogglePaid', 'btnMarkPaid', 'btnExportReport'))
        return (a['left'] + a['width'] > b['left'] + 0.5 or b['left'] + b['width'] > m['left'] + 0.5
                or m['left'] + m['width'] > e['left'] + 0.5)

    def captions_ok():
        d1 = abs(label('pcb_lblSelected')['Left'] - find_shape('btnTogglePaid')['left'])
        d2 = abs(label('pcb_lblMarkAll')['Left'] - find_shape('btnMarkPaid')['left'])
        return d1 < 1.5 and d2 < 1.5

    check(not overlaps(), 'default columns: the four blocks do not overlap')
    check(captions_ok(), 'default columns: each caption sits over its buttons')
    check(find_shape('btnGoToRecord')['left'] == find_shape('btnTogglePaid')['left'], 'Go to record is directly under Toggle Paid')
    check(find_shape('btnMarkUnpaid')['left'] == find_shape('btnMarkPaid')['left'], 'Unpaid is directly under Paid')
    wb.activate(REP)
    wb.set_column_hidden(REP, 8, True)
    wb.select(REP, 'A5')
    wb.select(REP, 'A6')
    cl = find_shape('btnClearFilters')['left']
    check(cl >= wb.column_lefts(REP, 9)[8] - 0.5, f'H hidden: the strip moved off the hidden column (Left {int(cl)})')
    check(not overlaps(), 'H hidden: the blocks do not overlap')
    check(captions_ok(), 'H hidden: each caption still sits over its buttons')
    wb.set_column_hidden(REP, 8, False)
    wb.select(REP, 'A5')
    wb.select(REP, 'A6')
    same = True
    for n in names:
        now = find_shape(n)['left']
        if abs(now - before[n]) > 1:
            same = False
            print(f'    {n} at {int(now)}, was {int(before[n])}')
    for n in ('pcb_lblSelected', 'pcb_lblMarkAll'):
        now = label(n)['Left']
        if abs(now - lbl_before[n]) > 1:
            same = False
            print(f'    {n} at {int(now)}, was {int(lbl_before[n])}')
    check(same, 'N shown again: every button and caption is back where it was')

    # ------------------------------------------------- button names / macros
    section('Button names agree with their macros')
    bad = []
    for ws in wb.sheet_names():
        for b in wb.buttons(ws):
            if b['name'].startswith('pcb_'):
                macro = b['on_action'][b['on_action'].find('!') + 1:]
                if len(b['name']) > 31 or not b['name'].startswith(f'pcb_{macro}_'):
                    bad.append(f"{ws}: {b['name']} -> {macro}")
    check(not bad, f"every button name is <= 31 characters and holds its macro {'; '.join(bad)}")
    d = find_shape('btnDeleteVisible')
    check(d is not None and d['on_action'].endswith('btnDeleteVisible'), 'Delete visible records runs btnDeleteVisible')
    wb.set_on_action(REP, d['name'], 'OlderName.xlsm!btnDeleteVisible')
    wb.run('HealButtons')
    healed = find_shape('btnDeleteVisible')['on_action']
    check(not healed.startswith('OlderName') and healed.endswith('btnDeleteVisible'),
          f"after a file rename HealButtons restores the macro (got '{healed}')")

    # --------------------------------------- Toggle Paid under a Paid = No view
    section('Toggle Paid while filtered to Paid = No')
    paid_col = rep_col('Paid')
    clear_all_boxes()
    for ws, t in ROOMS:
        wb.set_value(ws, body(ws, t, 'Paid'), 'No')
    wb.set_value(REP, 'B13', 'No')
    unpaid = jobs()
    check(unpaid >= 5, f'filtered to Paid = No: {unpaid} records')
    wb.activate(REP)
    wb.select(REP, addr(FIRST, paid_col))
    wb.run('btnTogglePaid')
    wb.calculate_full_rebuild()
    check(spill_ok(), 'the report is not left showing #SPILL!')
    check(jobs() == unpaid - 1, 'the record that became Yes drops out of the Paid = No view')
    yes = sum(wb.table_column_texts(ws, t, 'Paid').count('Yes') for ws, t in ROOMS)
    check(yes == 1, f'exactly one record is Yes in the data (found {yes})')
    clear_all_boxes()
    check(jobs() == total, 'clearing the filter shows every record again')

if xlmac._any_fail:
    print('FAIL: see above')
    sys.exit(1)
print('OK: Reports navigation, filters and buttons')
