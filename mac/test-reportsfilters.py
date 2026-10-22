# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-reportsfilters.ps1 - same checks, same messages, same
# order; see that script for the 0.10.27 Reports filter groups it guards.
# Expected counts are read straight off the job tables, not off Reports.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, addr, col_letters, sheet  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

REP, MAIN, ANNEX = 'Reports', 'Example Print Room', 'Annexe'
ROOMS = [(MAIN, 'tblJobs_EXAMPLE'), (ANNEX, 'tblJobs_ANNEX')]
BOXES = ['B4', 'B5', 'B8', 'B9', 'B11', 'B12', 'B13', 'E4', 'E5', 'E7', 'E10', 'E11', 'E12', 'B15', 'B16', 'E15', 'I7', 'K7']
MACROS = ['btnClearFilters', 'btnTogglePaid', 'btnGoToRecord', 'btnMarkPaid', 'btnMarkUnpaid', 'btnExportReport', 'btnDeleteVisible']

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.calculate_full_rebuild()
    time.sleep(3)
    wb.run('SetQuiet', True)
    wb.activate(REP)

    def jobs():
        wb.calculate_full_rebuild()
        return int(wb.text(REP, 'I9'))

    def clear_all_boxes():
        for a in BOXES:
            wb.clear_contents(REP, a)
        wb.calculate_full_rebuild()

    def records():
        """One record per job-table row that has a Date/Time."""
        out = []
        for ws, t in ROOMS:
            cols = {h: wb.table_column_texts(ws, t, h) for h in ('Date/Time', 'Status', 'Disregard Paper', 'Disregard Consumable', 'Paper Stock')}
            for i, dt in enumerate(cols['Date/Time']):
                if dt == '':
                    continue
                out.append({'Status': cols['Status'][i], 'DPaper': cols['Disregard Paper'][i],
                            'DCons': cols['Disregard Consumable'][i], 'Stock': cols['Paper Stock'][i]})
        return out

    def has_filter():
        return wb.call('HasActiveFilter', sheet(REP)) == 'True'

    def first_body(ws, t, h):
        return addr(wb.table_first_row(ws, t) + 1, wb.table_col(ws, t, h))

    def btn_left(macro):
        b = next((b for b in wb.buttons(REP) if b['name'].startswith(f'pcb_{macro}_')), None)
        return float(b['left']) if b else -1

    def col_left(c):
        return wb.column_lefts(REP, c)[c - 1]

    total = jobs()
    check(total >= 5, f'unfiltered Reports shows {total} records')

    # ---------------------------------------------------------- the group
    section('More filters: a closed nested group')
    check(not wb.row_hidden(REP, 14) and not wb.row_hidden(REP, 15),
          'rows 14-15 (the less-used filters) are visible: the row groups are gone (0.10.28)')
    check(wb.row_outline_level(REP, 4) == 1 and wb.row_outline_level(REP, 14) == 1, 'no row groups on the Reports sheet')
    lv = [wb.column_outline_level(REP, c) for c in (1, 6, 7, 8)]
    check(lv == [2, 2, 1, 1], 'filter columns A:F are one column group; G (gap) and H are not in it')
    check(wb.text(REP, 'A15') == 'Has a problem' and wb.text(REP, 'D15') == 'Has notes' and wb.text(REP, 'A16') == 'Disregarded',
          'the closed group holds Has a problem, Has notes and Disregarded')
    check(wb.text(REP, 'H7') == 'Sort by', 'Sort by sits at H7 above the table')
    for a in ('B15', 'B16', 'E15', 'E12'):
        check(wb.validation_type(REP, a) == 'validate list', f'{a} is a dropdown')

    # --------------------------------------------- counter on the summary row
    section('The summary row counts hidden filters that are set')
    wb.set_value(REP, 'E15', 'Yes')
    wb.calculate()
    wb.set_value(REP, 'B16', 'Paper')
    wb.calculate()
    clear_all_boxes()

    # ------------------------------------ Has a problem / Disregarded
    section('Has a problem and Disregarded')
    wb.set_value(MAIN, first_body(MAIN, 'tblJobs_EXAMPLE', 'Disregard Paper'), 'Yes')
    wb.set_value(ANNEX, first_body(ANNEX, 'tblJobs_ANNEX', 'Disregard Consumable'), 'Yes')
    wb.set_value(MAIN, first_body(MAIN, 'tblJobs_EXAMPLE', 'Disregard Consumable'), 'Yes')   # row 1 now has both flags
    wb.activate(MAIN)
    wb.run('btnAddPrintJob')          # an incomplete row: its Status lists issues
    wb.call('RefreshStatusNotes', sheet(MAIN))
    wb.activate(REP)
    recs = records()
    want_paper = sum(1 for r in recs if r['DPaper'] == 'Yes')
    want_cons = sum(1 for r in recs if r['DCons'] == 'Yes')
    want_prob = sum(1 for r in recs if r['Status'] not in ('OK', ''))
    print(f'  expected from the tables: paper disregarded {want_paper}, consumable disregarded {want_cons}, '
          f'problems {want_prob} of {len(recs)}')
    check(want_paper >= 1 and want_cons >= 1, 'fixture has both disregard flags set')
    want_both = sum(1 for r in recs if r['DPaper'] == 'Yes' and r['DCons'] == 'Yes')
    want_ok = sum(1 for r in recs if r['Status'] == 'OK')
    check(want_both >= 1, 'fixture has a record with both flags')
    wb.set_value(REP, 'B16', 'Paper')
    check(jobs() == want_paper, 'Disregarded = Paper matches the records with Disregard Paper = Yes')
    wb.set_value(REP, 'B16', 'Consumable')
    check(jobs() == want_cons, 'Disregarded = Consumable matches the records with Disregard Consumable = Yes')
    wb.set_value(REP, 'B16', 'Both')
    check(jobs() == want_both, f'Disregarded = Both matches the records with both flags ({want_both})')
    want_neither = sum(1 for r in recs if r['DPaper'] != 'Yes' and r['DCons'] != 'Yes')
    wb.set_value(REP, 'B16', 'Neither')
    check(jobs() == want_neither, f'Disregarded = Neither matches the records with no flag set ({want_neither})')
    check(1 <= want_neither < len(recs), 'fixture has records both with and without a disregard flag')
    clear_all_boxes()
    wb.set_value(REP, 'B15', 'Yes')
    check(jobs() == want_prob, 'Has a problem = Yes matches the records whose Status is not OK')
    check(want_prob >= 1, 'fixture has at least one problem record')
    wb.set_value(REP, 'B15', 'No')
    check(jobs() == want_ok, f'Has a problem = No matches the records whose Status is OK ({want_ok})')
    clear_all_boxes()
    check(jobs() == len(recs), 'blank filters show every record')

    # ----------------------------------------------- hidden filters still act
    section('A filter inside the closed row still applies and still counts')
    check(not has_filter(), 'no filter set: HasActiveFilter is false')
    wb.set_value(REP, 'B16', 'Paper')
    check(has_filter(), 'Disregarded alone satisfies the at-least-one-filter safeguard')
    clear_all_boxes()
    wb.set_value(REP, 'E15', 'Yes')
    check(has_filter(), 'Has notes (now in the closed row) alone satisfies the safeguard')
    wb.set_value(REP, 'B15', 'Yes')
    wb.set_value(REP, 'B16', 'Both')
    wb.run('btnClearFilters')
    check(wb.formula(REP, 'B15') == '' and wb.formula(REP, 'B16') == '' and wb.formula(REP, 'E15') == '',
          'Clear all filters empties the closed row too')

    # ------------------------------------------------ Student-supplied paper
    section('Student-supplied paper')
    clear_all_boxes()
    stock_col = wb.table_col(MAIN, 'tblJobs_EXAMPLE', 'Paper Stock')
    stock_name = wb.cell_text(MAIN, wb.table_first_row(MAIN, 'tblJobs_EXAMPLE') + 1, stock_col)
    descs = wb.table_column_texts('Papers', 'tblPapers', 'Description')
    pr = descs.index(stock_name) + 1 if stock_name in descs else 0
    check(pr > 0, f"the first job's stock '{stock_name}' is a Papers row")
    sup_cell = addr(wb.table_first_row('Papers', 'tblPapers') + pr, wb.table_col('Papers', 'tblPapers', 'Supplied by student'))
    wb.set_value('Papers', sup_cell, 'Yes')
    recs = records()
    want_sup = sum(1 for r in recs if r['Stock'] == stock_name)
    print(f"  expected: {want_sup} of {len(recs)} records on '{stock_name}'")
    check(want_sup >= 1, 'at least one record uses the flagged stock')
    wb.set_value(REP, 'E12', 'Yes')
    check(jobs() == want_sup, 'Yes shows exactly the records on stock marked Supplied by student')
    wb.set_value(REP, 'E12', 'No')
    check(jobs() == len(recs) - want_sup, 'No shows every other record')
    wb.clear_contents(REP, 'E12')
    check(jobs() == len(recs), 'blank shows every record')
    wb.set_value('Papers', sup_cell, 'No')
    wb.set_value(REP, 'E12', 'Yes')
    check(jobs() == 0, 'with no stock marked Supplied by student, Yes shows nothing')
    # The two built-in stocks are not Papers rows; put one on a record and count again.
    row2 = addr(wb.table_first_row(MAIN, 'tblJobs_EXAMPLE') + 2, stock_col)
    wb.set_value(MAIN, row2, 'Supplied (Roll)')
    if wb.text(MAIN, row2) == 'Supplied (Roll)':
        recs = records()
        want_builtin = sum(1 for r in recs if r['Stock'] in ('Supplied (Roll)', 'Supplied (Sheet)'))
        check(want_builtin >= 1, 'a record is on a built-in Supplied stock')
        check(jobs() == want_builtin, 'Yes also shows records on the built-in Supplied stocks')
    else:
        print("  NOTE: the sheet would not accept 'Supplied (Roll)' on that row; built-in stock case not exercised")
    clear_all_boxes()
    check(not has_filter(), 'all clear again')

    # --------------------------------------------- a rebuild does not stack levels
    section('Rebuilding the workbook leaves the outline as it was')
    wb.run('InitialiseWorkbook')
    wb.run('InitialiseWorkbook')
    check(wb.column_outline_level(REP, 1) == 2 and wb.column_outline_level(REP, 8) == 1 and wb.row_outline_level(REP, 4) == 1,
          'after two rebuilds the column group is still one level')
    check(not wb.row_hidden(REP, 5), 'the main filters are expanded after a rebuild')
    check(not wb.column_hidden(REP, 1), 'the filter columns are expanded after a rebuild')

    # ------------------------------------- freeze panes and the button strip
    section('No freeze panes; the button strip is one rigid piece (0.11.0 refinement)')
    wb.activate(REP)
    wb.scroll_to(1, 1)
    check(not wb.freeze_panes(), 'Reports has no freeze panes')
    open_left = btn_left('btnClearFilters')
    check(open_left >= col_left(8) - 1, f'expanded: the strip starts at column H (Left {int(open_left)}, H at {int(col_left(8))})')
    lefts0 = [btn_left(m) for m in MACROS]

    wb.show_levels(REP, 0, 1)
    check(wb.column_hidden(REP, 1) and wb.column_hidden(REP, 6) and not wb.column_hidden(REP, 7),
          'collapsed: A:F hidden, the gap column G stays')
    wb.call('RepositionReportsButtons', sheet(REP))
    shut = btn_left('btnClearFilters')
    check(abs(shut - col_left(8)) < 1.5, f'collapsed: the strip follows the table to column H (Left {int(shut)})')
    check(shut < open_left - 1, 'collapsed: the strip has moved left of its expanded position')
    wb.show_levels(REP, 0, 2)
    wb.call('RepositionReportsButtons', sheet(REP))
    check(abs(btn_left('btnClearFilters') - open_left) < 1.5, 'expanded again: the strip is back where it was, clear of the filter boxes')
    g_left = col_left(7)
    g_width = wb.column_lefts(REP, 8)[7] - g_left
    check(btn_left('btnClearFilters') >= g_left + g_width - 1, 'the strip starts right of the filter columns')

    # Resizing a results column under the strip moves none of it.
    w0 = wb.column_width(REP, 11)
    wb.set_column_width(REP, 11, w0 + 20)
    lefts1 = [btn_left(m) for m in MACROS]
    moved = sum(1 for a, b in zip(lefts0, lefts1) if abs(a - b) > 0.5)
    check(moved == 0, f'widening a results column leaves the strip alone ({moved} of {len(MACROS)} buttons moved)')
    wb.call('RepositionReportsButtons', sheet(REP))
    lefts2 = [btn_left(m) for m in MACROS]
    moved = sum(1 for a, b in zip(lefts0, lefts2) if abs(a - b) > 0.5)
    check(moved == 0, f'...and re-settling keeps it where it was ({moved} moved)')
    wb.set_column_width(REP, 11, w0)

if xlmac._any_fail:
    print('FAIL: see above')
    sys.exit(1)
print('OK: Reports filter groups')
