# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-summaryformat.ps1 - same checks, same messages, same order;
# see that script for the history of each one, including why NumberFormat is
# re-read a few times after InitialiseWorkbook.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlmac  # noqa: E402
from xlmac import open_copy, check, section, SUM_FIRST_ROW, SUM_HDR_ROW, SUM_TOT_ROW  # noqa: E402
from fixture_annexe import add_annexe_fixture  # noqa: E402

with open_copy() as wb:
    add_annexe_fixture(wb)

    # -------------------------------------------------- error: Summary ---
    section('Error state: Summary Type conditional formatting', first=True)
    err_rng = f'E{SUM_FIRST_ROW}:E2000'
    fcs = wb.format_conditions('Summary', err_rng)
    n = len(fcs)
    check(n >= 1, f'Summary Type column has a conditional format (count {n})')
    if n >= 1:
        f1, font_color = fcs[0]
        check('(not in Papers)' in f1, f"rule tests for '(not in Papers)' (formula: {f1})")
        check(font_color == 2097328, f'rule font is red RGB(176,0,32) (got {font_color})')

    # -------------------------------------------------------- legend ---
    section('Summary legend (column R)')
    expected = ['Type your own values here', 'Calculated automatically', 'Workbook configuration',
                'Reference information (read-only)']
    for i, want in enumerate(expected):
        r = 10 + i
        text = wb.cell_text('Summary', r, 18)
        check(text == want, f"R{r} reads '{want}' (got '{text}')")

    # Chargeable is the table's last column, O: until 0.11.5 the legend
    # narrowed O to 4 and its figures showed as ####.
    section('Summary table: Chargeable (column O) is readable')
    head = wb.cell_text('Summary', SUM_HDR_ROW, 15)
    check(head == 'Chargeable', f"O{SUM_HDR_ROW} is the Chargeable header (got '{head}')")
    width = wb.column_width('Summary', 15)
    check(width >= 10, f'column O is at least 10 wide (got {width})')
    first = wb.cell_text('Summary', SUM_FIRST_ROW, 15)
    check('#' not in first, f"O{SUM_FIRST_ROW} shows a figure, not #### (got '{first}')")

    # ------------------------------------------------------- currency ---
    section('Currency wiring: SET_CURRENCY -> NumberFormat')
    rep_addr = 'N9'   # Matching: Chargeable total (modReports.REP_MATCH_VAL_ROW)
    wb.set_value(*wb.wb_name('SET_CURRENCY'), '$')
    wb.run('SetQuiet', True)
    wb.run('InitialiseWorkbook')
    wb.run('SetQuiet', False)

    def read_fmt(sheet, a1):
        for _ in range(6):
            v = wb.number_format(sheet, a1)
            if v:
                return v
            time.sleep(0.5)
        return ''

    sum_fmt = read_fmt('Summary', f'D{SUM_TOT_ROW}')
    rep_fmt = read_fmt('Reports', rep_addr)
    check(sum_fmt == '[$$]#,##0.00', f"Summary D6 follows SET_CURRENCY (got '{sum_fmt}')")
    check(rep_fmt == '[$$]#,##0.00', f"Reports {rep_addr} follows SET_CURRENCY (got '{rep_fmt}')")

    # Re-running InitialiseWorkbook must not have stacked duplicate rules.
    section('Idempotency: re-run does not stack rules')
    n2 = len(wb.format_conditions('Summary', err_rng))
    check(n2 == 1, f'Summary Type rule count after a second InitialiseWorkbook: {n2} (expected 1)')

if xlmac._any_fail:
    print('FAIL: see above')
    sys.exit(1)
print('OK: summary formatting and currency wiring')
