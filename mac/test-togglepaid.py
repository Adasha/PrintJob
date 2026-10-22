# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-togglepaid.ps1 - same checks, same messages, same order;
# see that script for the Toggle Paid button it guards. Differences:
#
#   - The "Selected record:" caption is a Form Controls label, read through
#     modScript.ScriptProp on a shape token; the Reports spill through a
#     spill token. MarkVisibleReportsConfirmed and TogglePaidSelected go
#     through Workbook.call (objects, and an Optional parameter).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, col_letters, sheet, spill, shape  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

REP, MAIN, ANNEX = 'Reports', 'Example Print Room', 'Annexe'
ROOMS = [(MAIN, 'tblJobs_EXAMPLE'), (ANNEX, 'tblJobs_ANNEX')]
FILTER_CELLS = ['B4', 'B5', 'B8', 'B9', 'B11', 'B12', 'B13', 'E4', 'E5', 'E7', 'E10', 'E11', 'E12', 'B15', 'E15', 'I7', 'K7']

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    def snapshot():
        h = {}
        for ws, t in ROOMS:
            ids = wb.table_column_texts(ws, t, 'Job ID')
            paid = wb.table_column_texts(ws, t, 'Paid')
            for i, p in zip(ids, paid):
                h[f'{t}|{i}'] = p.strip()
        return h

    def find_cols():
        heads = wb.texts(REP, 'H10:AB10')[0]   # columns 8..28
        o = {'Loc': 0, 'Job': 0, 'Paid': 0}
        for i, hd in enumerate(heads):
            for k, name in (('Loc', 'Location'), ('Job', 'Job ID'), ('Paid', 'Paid')):
                if hd == name:
                    o[k] = 8 + i
        return o

    def record_at(row):
        return f"tblJobs_{wb.cell_text(REP, row, cols['Loc'])}|{wb.cell_text(REP, row, cols['Job'])}"

    def spill_count():
        return int(wb.prop(spill(REP, 'H11'), 'Rows.Count'))

    def spill_intact():
        return not wb.text(REP, 'H11').startswith('#SPILL') and spill_count() == n

    def filter_state():
        return ';'.join(f'{c}={wb.text(REP, c)}' for c in FILTER_CELLS)

    def audit_actions():
        return wb.table_column_texts('_Audit', 'tblAudit', wb.table_headers('_Audit', 'tblAudit')[2])

    def shown(row):
        return wb.cell_text(REP, row, cols['Paid']).strip()

    def other(s):
        return 'No' if s == 'Yes' else 'Yes'

    def select_cell(row, col=None):
        wb.select(REP, addr(row, col or cols['Paid']))

    def changed(before, after):
        return [k for k in after if after[k] != before.get(k)]

    def changed_detail(before, after, skip):
        d = [f"{k}: {repr(before[k]) if k in before else '(absent)'} -> {after[k]!r}"
             for k in after if k != skip and after[k] != before.get(k)]
        return '' if not d else ' [' + '; '.join(d) + f'; keys before/after {len(before)}/{len(after)}]'

    wb.activate(REP)
    wb.calculate_full_rebuild()
    cols = find_cols()
    n = spill_count()
    print(f"  unfiltered Reports shows {n} records; Paid is column {cols['Paid']}")
    filters_before = filter_state()

    # --------------------------------------------------------- the column
    section('Paid cells are read-only: locked, no dropdown')
    vt = wb.validation_type(REP, addr(11, cols['Paid']))
    check(vt == '', f"no validation (no dropdown) on the Paid cells (type: '{vt}')")
    check(wb.cell_format(REP, addr(11, cols['Paid']))['locked'], 'Paid cell is locked')
    check(wb.cell_format(REP, addr(2000, cols['Paid']))['locked'], 'locked down to row 2000')
    check(wb.protected(REP), 'Reports sheet is protected')

    # ----------------------------------------------------------- the button
    section('The Toggle Paid button')
    buttons = wb.buttons(REP)
    btn = next((b for b in buttons if b['name'].startswith('pcb_btnTogglePaid_')), None)
    go_btn = next((b for b in buttons if b['name'].startswith('pcb_btnGoToRecord_')), None)
    check(btn is not None, 'Toggle Paid button exists')
    if btn:
        check(btn['on_action'].endswith('btnTogglePaid'), f"button runs btnTogglePaid (got '{btn['on_action']}')")
        check(len(btn['name']) <= 31, f"shape name within 31 characters ({btn['name']})")
        if go_btn:
            check(abs(btn['left'] - go_btn['left']) < 1.5, 'sits in the same block as Go to record (same Left)')
            check(go_btn['top'] > btn['top'], 'Go to record is below Toggle Paid')
        lbl = shape(REP, 'pcb_lblSelected')
        try:
            cap = wb.prop(lbl, 'Caption')
            has_lbl = True
        except Exception:
            has_lbl = False
        check(has_lbl, "the 'Selected record:' caption exists")
        if has_lbl:
            check(cap == 'Selected record:', f"caption reads 'Selected record:' (got '{cap}')")
            l_left, l_top, l_h = (float(wb.prop(lbl, p)) for p in ('Left', 'Top', 'Height'))
            check(abs(l_left - btn['left']) < 1.5, 'the caption is left-aligned over Toggle Paid')
            check(l_top + l_h <= btn['top'] + 1, 'the caption sits above Toggle Paid')
        r4_top = wb.row_top_height(REP, 4)[0]
        r5_top = wb.row_top_height(REP, 5)[0]
        check(btn['top'] >= r4_top - 1 and btn['top'] + btn['height'] <= r5_top + 1, 'sits inside row 4')

    # ------------------------------------------------ toggle, both rooms
    section('Toggle Paid lands on the right record (both rooms)')
    done = set()
    for row in range(11, n + 11):
        key = record_at(row)
        room = key.split('|')[0]
        if room in done:
            continue
        done.add(room)
        s = shown(row)
        target = other(s)
        before = snapshot()
        select_cell(row)
        wb.run('btnTogglePaid')
        after = snapshot()
        check(after.get(key) == target, f"{room} : {key} changed '{s}' -> '{target}'")
        check(len(changed(before, after)) == 1, f'{room} : no other record changed{changed_detail(before, after, key)}')
        check(spill_intact(), f'{room} : results spill is intact ({spill_count()} rows)')
        check(shown(row) == target, f"{room} : the row now displays '{target}'")
        check(wb.protected(REP) and wb.protected(MAIN) and wb.protected(ANNEX), f'{room} : sheets re-protected')
        # And back again.
        wb.run('btnTogglePaid')
        back = snapshot()
        check(back.get(key) == s or (s == '' and back.get(key) == 'Yes') or back.get(key) == other(target),
              f'{room} : toggling again flips it back')
    check(filter_state() == filters_before, 'filters untouched')
    check('Paid edited (Reports)' in audit_actions(), "audit has 'Paid edited (Reports)'")

    # ------------------------------------------- any cell of the row works
    section('Any cell of the row selects the record')
    row = 12
    key = record_at(row)
    s = shown(row)
    before = snapshot()
    select_cell(row, 8)
    wb.run('btnTogglePaid')
    after = snapshot()
    check(after.get(key) == other(s), f"record changed from a cell in column A ('{s}' -> '{other(s)}')")
    check(len(changed(before, after)) == 1, f'only that record changed{changed_detail(before, after, key)}')

    # ----------------------------------------------------- several rows
    section('Several rows: all take the value the active row flips to')
    r1, r2, r3 = 13, 14, 15
    select_cell(r1)
    s1 = shown(r1)
    keys = [record_at(r1), record_at(r2), record_at(r3)]
    expect = other(s1)
    before = snapshot()
    wb.run('SetQuiet', True)
    pc = col_letters(cols['Paid'])
    wb.select(REP, f'{pc}{r1}:{pc}{r3}')
    wb.run('btnTogglePaid')    # unconfirmed under SetQuiet: Ask declines
    after = snapshot()
    check(len(changed(before, after)) == 0, 'declined confirmation: nothing changed')
    wb.call('TogglePaidSelected', True)
    after = snapshot()
    for k in keys:
        check(after.get(k) == expect, f"{k} is '{expect}'")
    check(len(changed(before, after)) <= 3, 'no record outside the selection changed')
    check(spill_intact(), 'spill intact')
    check('Paid edited (Reports)' in audit_actions(), 'audited')

    # ------------------------------------------------ nothing to change
    section('A selection with no record is refused')
    for label, r in (('header row', 10), ('below the results', 11 + n + 3)):
        wb.run('SetQuiet', True)
        before = snapshot()
        select_cell(r)
        wb.run('btnTogglePaid')
        log = wb.run('QuietLog')
        check('Select a record first' in log, f'{label}: said to select a record')
        check(len(changed(before, snapshot())) == 0, f'{label}: no record changed')
    check(spill_intact(), 'spill intact')

    # ---------------------------------- bulk buttons still fine afterwards
    section('Mark all as... still works')
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    wb.calculate_full_rebuild()
    nn = spill_count()
    wb.run('SetQuiet', True)
    wb.call('MarkVisibleReportsConfirmed', sheet(REP), spill(REP, 'H11'), cols['Loc'], cols['Job'], nn, 'Yes')
    after = snapshot()
    locs = [r[0] for r in wb.texts(REP, f"{col_letters(cols['Loc'])}11:{col_letters(cols['Loc'])}{10 + nn}")]
    jobs = [r[0] for r in wb.texts(REP, f"{col_letters(cols['Job'])}11:{col_letters(cols['Job'])}{10 + nn}")]
    marked = sum(1 for lc, jb in zip(locs, jobs) if after.get(f'tblJobs_{lc}|{jb}') == 'Yes')
    check(marked == nn, f'all {nn} visible records marked Yes')

    wb.run('SetQuiet', False)

finish()
