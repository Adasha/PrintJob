# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-statusnotes.ps1 - same checks, same output lines (it
# reports with OK:/FAIL: lines, not Check); see that script for the Status
# hover notes and warning fill it guards. A cell's note, its displayed fill
# and a conditional format's fill are not in Excel's AppleScript dictionary,
# so they are read through modScript.ScriptProp on a range token.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, addr, col_letters, sheet, rng, ExcelError  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

WS, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
failed = 0

with open_copy() as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)
    wb.activate(WS)

    sc = wb.table_col(WS, TBL, 'Status')
    px = wb.column_width(WS, sc) * 7 + 5
    if px > 110:
        print(f'FAIL: Status is {px} px wide, expected about 90')
        failed += 1
    else:
        print(f'OK: Status is about {int(px)} px wide')

    # A new row is incomplete, so its Status lists issues.
    wb.run('btnAddPrintJob')
    wb.call('RefreshStatusNotes', sheet(WS))

    def note(row):
        """The cell's note text, or None if it has none."""
        try:
            return wb.prop(rng(WS, addr(row, sc)), 'Comment.Text')
        except ExcelError:
            return None

    def check_notes(label):
        bad = with_note = 0
        first = wb.table_first_row(WS, TBL) + 1
        statuses = wb.table_column_texts(WS, TBL, 'Status')
        for i, txt in enumerate(statuses, 1):
            cm = note(first + i - 1)
            want = txt not in ('', 'OK')
            if want:
                if cm is None:
                    print(f"FAIL ({label}): row {i} status '{txt}' has no note")
                    bad += 1
                else:
                    with_note += 1
                    flat = cm.replace('• ', '').replace('\r\n', '\n').replace('\r', '\n').replace('\n', '; ').strip()
                    if flat != txt:
                        print(f"FAIL ({label}): row {i} note '{flat}' <> status '{txt}'")
                        bad += 1
            elif cm is not None:
                print(f"FAIL ({label}): row {i} status '{txt}' should have no note")
                bad += 1
        print(f'{label}: {len(statuses)} rows, {with_note} with notes, {bad} problems')
        return bad, with_note

    b, w = check_notes('after add')
    failed += b
    if w < 1:
        print('FAIL: expected at least one row with a note after adding an incomplete job')
        failed += 1

    wb.run('btnRefreshLocations')
    b, _ = check_notes('after refresh')
    failed += b

    # ---- warning fill: every non-"OK" Status cell, on every location sheet
    print()
    print('=== Warning state: Status column conditional formatting ===')

    def status_body(ws, table):
        c = col_letters(wb.table_col(ws, table, 'Status'))
        return f'{c}{wb.table_first_row(ws, table) + 1}:{c}{wb.table_last_row(ws, table)}'

    for ws, table in ((WS, TBL), ('Annexe', 'tblJobs_ANNEX')):
        body = rng(ws, status_body(ws, table))
        cnt = int(wb.prop(body, 'FormatConditions.Count'))
        if cnt < 1:
            print(f'FAIL: {ws} Status has no conditional format')
            failed += 1
            continue
        f1 = wb.prop(body, 'FormatConditions.Item(1).Formula1')
        fill = int(float(wb.prop(body, 'FormatConditions.Item(1).Interior.Color')))
        if not ('<>' in f1 and 'OK' in f1.split('<>', 1)[1]):
            print(f'FAIL: {ws} rule is not a <> OK test ({f1})')
            failed += 1
        elif fill != 49407:
            print(f'FAIL: {ws} fill is {fill}, expected amber 49407')
            failed += 1
        else:
            print(f'OK: {ws} Status warns in amber when not OK')

    # The incomplete row added above reads something other than OK, so the
    # rule must actually fire on it, not just exist.
    last = wb.table_last_row(WS, TBL)
    last_text = wb.cell_text(WS, last, sc)
    shown = int(float(wb.prop(rng(WS, addr(last, sc)), 'DisplayFormat.Interior.Color')))
    if last_text == 'OK':
        print('FAIL: the added incomplete row reads OK')
        failed += 1
    elif shown != 49407:
        print(f'FAIL: incomplete row Status is not amber (fill {shown})')
        failed += 1
    else:
        print(f"OK: the incomplete row's Status ('{last_text}') is shown amber")

    # Re-running setup must not stack a second rule (see the .ps1 for why the
    # count is compared, not required to be 1).
    cnt_before = int(wb.prop(rng(WS, status_body(WS, TBL)), 'FormatConditions.Count'))
    if cnt_before < 2:
        print(f'FAIL: Status has only {cnt_before} rules (expected problem + notice + shading)')
        failed += 1
    wb.run('InitialiseWorkbook')
    cnt2 = int(wb.prop(rng(WS, status_body(WS, TBL)), 'FormatConditions.Count'))
    if cnt2 != cnt_before:
        print(f'FAIL: Status has {cnt2} rules after a second InitialiseWorkbook (expected {cnt_before})')
        failed += 1
    else:
        print(f'OK: re-running InitialiseWorkbook leaves the same {cnt2} Status rules')

if failed > 0:
    print(f'FAIL: {failed} problem(s)')
    sys.exit(1)
print('OK: status notes match')
