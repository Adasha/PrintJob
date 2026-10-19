# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-subsidy.ps1 - same checks, same messages, same order; see
# that script for what the University subsidy is. Differences:
#
#   - Validation.Type comes back as Excel's enum name: 'validate decimal' is
#     COM's 2.
#   - The exported .xlsx is read back by opening it in the same Excel (there is
#     no second instance to start on a Mac), then closed.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, Workbook, open_in_excel, osa, sheet, addr  # noqa: E402

SET, TBL, REP = 'Settings', 'tblSettings', 'Reports'

ctx = open_copy()
with ctx as wb:
    wb.run('SetQuiet', True)
    pct_cell = wb.wb_name('SET_SUBSIDY_PCT')

    def match_col(label):
        """Sheet column of a label in the Matching row, row 8."""
        row = wb.texts(REP, 'H8:AN8')[0]   # columns 8..40
        return row.index(label) + 8 if label in row else 0

    def set_pct(v):
        wb.call('UnlockSheet', sheet(SET))
        if v is None:
            wb.clear_contents(*pct_cell)
        else:
            wb.set_value(*pct_cell, float(v))
        wb.calculate()

    def key_rows():
        return [i for i, k in enumerate(wb.table_column_texts(SET, TBL, 'Key'), 1) if k == 'SUBSIDY_PCT']

    section('The Settings row', first=True)
    rows = key_rows()
    check(len(rows) == 1, 'exactly one SUBSIDY_PCT row')
    check(wb.number(*pct_cell) == 0, 'defaults to 0')
    fmt = wb.number_format(*pct_cell)
    check(fmt.endswith('%'), f'value cell is formatted as a percentage ({fmt})')
    check(wb.cell_format(*pct_cell)['locked'] is False, 'value cell is unlocked (editable)')
    check(wb.validation_type(*pct_cell) == 'validate decimal', 'value cell has decimal validation')
    notes = wb.table_column_texts(SET, TBL, 'Notes')[rows[0] - 1] if rows else ''
    check(not notes.startswith('Read-only'), 'row is not marked Read-only (so a restore carries it)')

    set_pct(0.25)
    wb.run('EnsureSubsidySetting')
    wb.run('EnsureSubsidySetting')
    check(len(key_rows()) == 1, 're-running setup does not duplicate the row')
    check(abs(wb.number(*pct_cell) - 0.25) < 1e-9, 're-running setup keeps a value the user set')
    set_pct(0)

    section('Reports totals')
    c_ch, c_pd, c_un = match_col('Chargeable'), match_col('Paid'), match_col('Unpaid')
    check(c_ch > 0 and c_pd > 0 and c_un > 0, 'Chargeable / Paid / Unpaid totals found in the Matching row')
    check(not wb.column_hidden(REP, c_un), 'the Unpaid column is visible (so the subsidy total above it is)')

    def chg():
        return wb.number(REP, addr(9, c_ch))

    def bill():
        return wb.number(REP, addr(7, c_un))

    def lab():
        return wb.cell_text(REP, 6, c_un)

    check(chg() > 0, f'there is a chargeable total to test with ({chg()})')
    check(abs(bill() - chg()) < 0.005, f'0%: billable equals chargeable ({bill()})')
    check(lab() == 'After 0% subsidy', f"0%: label reads '{lab()}'")

    set_pct(0.25)
    full = chg()
    expect = round(full * 0.75, 2)
    check(abs(bill() - expect) < 0.005, f'25%: billable is chargeable x 0.75 ({bill()} vs {expect})')
    check(lab() == 'After 25% subsidy', f"25%: label reads '{lab()}'")
    check(abs(chg() - full) < 0.005, 'Chargeable total is still the full amount')
    check(abs(wb.number(REP, addr(9, c_un)) - full) < 0.005, 'Unpaid is still the full amount (nothing paid yet)')
    bfmt = wb.number_format(REP, addr(7, c_un))
    check(bfmt.endswith('#,##0.00'), f'billable is formatted as money ({bfmt})')

    set_pct(0.125)
    check(lab() == 'After 12.5% subsidy', f"12.5%: label reads '{lab()}'")
    check(abs(bill() - round(full * 0.875, 2)) < 0.005, '12.5%: billable is chargeable x 0.875')

    # Per-job amounts are untouched: the results' Chargeable column still sums to the full total.
    heads = wb.texts(REP, 'H10:AN10')[0]
    hdr_col = max((i + 8 for i, h in enumerate(heads) if h == 'Chargeable'), default=0)
    col = xlmac.col_letters(hdr_col)
    total = sum(float(r[0]) for r in wb.values(REP, f'{col}11:{col}400') if isinstance(r[0], float))
    check(abs(total - full) < 0.005, f'per-job Chargeable column still sums to the full amount ({total})')

    set_pct(1.5)
    check(abs(bill()) < 0.005, 'over 100% reads as 100% (billable 0), not an error')
    set_pct(-0.1)
    check(abs(bill() - full) < 0.005, 'a negative value reads as no subsidy')
    set_pct(None)
    check(abs(bill() - full) < 0.005, 'a blank value reads as no subsidy')

    section('Follows the filters')
    set_pct(0.5)
    wb.set_value(REP, 'E7', 'Epson SureColor P9500')
    wb.calculate_full_rebuild()
    filtered = chg()
    check(0 < filtered < full, f'the filter narrows Chargeable ({filtered} of {full})')
    check(abs(bill() - round(filtered * 0.5, 2)) < 0.011, f'billable follows the filtered total ({bill()})')

    section('Export report')
    set_pct(0.25)
    wb.run('SetExportNames', False)
    wb.calculate_full_rebuild()
    exp_bill, exp_chg = bill(), chg()
    wb.activate(REP)
    wb.run('btnExportReport')
    time.sleep(0.3)
    found = sorted(glob.glob(os.path.join(ctx.dir, 'PrintCosts-Report-*.xlsx')), key=os.path.getmtime)
    xlsx = found[-1] if found else None
    check(xlsx is not None, 'an .xlsx was produced')
    if xlsx:
        out = Workbook(os.path.basename(xlsx))
        open_in_excel(xlsx)
        try:
            w = osa('tell application "Microsoft Excel" to return name of worksheet 1 of workbook "'
                    + out.name + '"')
            col_a = [r[0] for r in out.texts(w, 'A1:A40')]
            r = {}
            for i, k in enumerate(col_a[:20], 1):
                if k and k not in r:
                    r[k] = i
            key = 'Billable after 25% subsidy'
            check(key in r, 'a "Billable after 25% subsidy" line is present')
            if key in r:
                row = r[key]
                check(row == r.get('Still owed', -9) + 1, 'it follows Still owed')
                got = out.number(w, f'B{row}')
                check(abs(got - exp_bill) < 0.005, f'it carries the Reports value ({got} vs {exp_bill})')
                check(abs(out.number(w, f"B{r.get('Total chargeable', 1)}") - exp_chg) < 0.005,
                      'Total chargeable is still the full amount')
                check(out.number_format(w, f'B{row}').endswith('#,##0.00'), 'it is formatted as money')
                check(not str(out.value(w, f'A{row}')).startswith('='), 'the label is plain text')
            tbl = next((i for i, k in enumerate(col_a, 1) if k == 'Date/Time'), 0)
            check(tbl > 0, f'the results table still follows the header block (row {tbl})')
        finally:
            osa(f'tell application "Microsoft Excel" to close workbook "{out.name}" saving no')

if xlmac._any_fail:
    print('FAILED')
    sys.exit(1)
print('All subsidy checks passed.')
