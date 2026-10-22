# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-deletereports.ps1 - same steps, same output lines (it
# reports with OK:/FAIL: lines and PASS, not Check); see that script for why
# DeleteVisibleReportsConfirmed is called directly. The exported .xlsx is read
# back in the same Excel rather than a second instance.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself.

import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, Workbook, open_in_excel, osa, sheet, spill  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

REP, MAIN, ANNEX = 'Reports', 'Example Print Room', 'Annexe'


def fail(msg):
    print(f'FAIL: {msg}')
    sys.exit(1)


ctx = open_copy()
with ctx as wb:
    add_annexe_fixture(wb)
    wb.run('SetQuiet', True)

    print('=== Export report ===')
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    wb.calculate_full_rebuild()
    before_jobs = int(wb.text(REP, 'I9'))
    print(f"  filtered to Printer = 'Epson SureColor P9500': {before_jobs} jobs")
    if before_jobs == 0:
        fail('expected at least one job for this filter')

    wb.activate(REP)
    wb.run('btnExportReport')
    print(wb.run('QuietLog'))

    found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-Report-*.xlsx')), key=os.path.getmtime)
    if not found:
        fail('no report .xlsx was produced')
    xlsx = found[-1]
    print(f'  OK: {os.path.basename(xlsx)} ({os.path.getsize(xlsx)} bytes)')

    sig = wb.call('ReportsExportSig', sheet(REP))
    print(f"  stamped signature: '{sig}'")
    if not sig:
        fail('ReportsExportSig was not stamped')

    # Peek inside the exported file: a real, static-value snapshot, not a link back.
    out = Workbook(os.path.basename(xlsx))
    open_in_excel(xlsx)
    try:
        ws1 = osa(f'tell application "Microsoft Excel" to return name of worksheet 1 of workbook "{out.name}"')
        print(f"  snapshot header row: {out.value(ws1, 'A1')}")
        a2_formula = out.has_formula(ws1, 'A2')
        print(f'  snapshot row2 formula type: {a2_formula}')
    finally:
        osa(f'tell application "Microsoft Excel" to close workbook "{out.name}" saving no')
    if a2_formula:
        fail('snapshot should be static values, not formulas')
    print('  OK: snapshot holds static values, no live formulas')

    print()
    print('=== Delete visible records (staleness check should now be clean) ===')
    before_main = wb.table_row_count(MAIN, 'tblJobs_EXAMPLE')
    before_annexe = wb.table_row_count(ANNEX, 'tblJobs_ANNEX')
    print(f'  before: Main={before_main} rows, Annexe={before_annexe} rows')

    wb.activate(REP)
    n = int(wb.prop(spill(REP, 'H11'), 'Rows.Count'))
    heads = wb.texts(REP, 'H10:AB10')[0]
    loc_col = max((8 + i for i, h in enumerate(heads) if h == 'Location'), default=0)
    job_col = max((8 + i for i, h in enumerate(heads) if h == 'Job ID'), default=0)
    print(f'  deleting {n} visible rows (locCol={loc_col}, jobCol={job_col})')
    wb.call('DeleteVisibleReportsConfirmed', sheet(REP), spill(REP, 'H11'), loc_col, job_col, n)
    print(wb.run('QuietLog'))

    after_main = wb.table_row_count(MAIN, 'tblJobs_EXAMPLE')
    after_annexe = wb.table_row_count(ANNEX, 'tblJobs_ANNEX')
    print(f'  after: Main={after_main} rows, Annexe={after_annexe} rows')
    if (before_main - after_main) + (before_annexe - after_annexe) != n:
        fail(f'expected exactly {n} rows removed across both sheets')
    print(f'  OK: exactly {n} rows removed, from the correct source sheets')

    heads = wb.table_headers('_Audit', 'tblAudit')
    last_action = (wb.table_column_texts('_Audit', 'tblAudit', heads[2]) or [''])[-1]
    last_detail = (wb.table_column_texts('_Audit', 'tblAudit', heads[4]) or [''])[-1]
    print(f"  audit log: action='{last_action}' detail='{last_detail}'")
    if last_action != 'Delete visible (Reports)':
        fail('audit log entry missing/wrong')
    print('  OK: audit entry recorded')

    wb.run('SetQuiet', False)

print()
print('PASS')
