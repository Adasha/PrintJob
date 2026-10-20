# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-jobvalidation.ps1 - same checks, same messages, same order;
# see that script for why validation is bound by header name. Differences:
#
#   - Validation types are Excel's enum names ('validate list' is COM's 3,
#     'validate decimal' 2, 'validated date' 4); messages show the name.
#   - The old Cut + Insert column move is done as two method calls through
#     modScript.ScriptProp ("Cut", then "Insert -4161" = xlShiftToRight).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, col_letters, sheet, rng  # noqa: E402

WS, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
LIST, DECIMAL, DATE = 'validate list', 'validate decimal', 'validated date'

# Header name -> expected validation type, and the list text where it is fixed.
RULES = {
    'Date/Time': {'Type': DATE},
    'Technician': {'Type': LIST},
    'Printer': {'Type': LIST},
    'Paper Stock': {'Type': LIST},
    'Qty': {'Type': DECIMAL},
    'Print Width mm': {'Type': DECIMAL},
    'Sheet size': {'Type': LIST, 'F1': '=RNG_STD_SIZES'},
    'Disregard Paper': {'Type': LIST, 'F1': 'Yes,No'},
    'Disregard Consumable': {'Type': LIST, 'F1': 'Yes,No'},
    'Paid': {'Type': LIST, 'F1': 'Yes,No'},
}

with open_copy() as wb:

    def vtype(a1):
        return wb.validation_type(WS, a1) or '(none)'

    def check_row(row_no, label):
        headers = wb.table_headers(WS, TBL)
        row = wb.table_first_row(WS, TBL) + row_no
        c0 = wb.table_col(WS, TBL, headers[0])
        for ci, name in enumerate(headers):
            a1 = addr(row, c0 + ci)
            t = vtype(a1)
            f1 = wb.validation_formula1(WS, a1)
            if name in RULES:
                want = RULES[name]
                ok = t == want['Type']
                if ok and 'F1' in want:
                    ok = f1 == want['F1']
                check(ok, f"{label}: '{name}' rule ok (type {t}, '{f1}')")
            else:
                check(t == '(none)', f"{label}: '{name}' has no validation")

    def check_below(label):
        first = wb.table_last_row(WS, TBL) + 1
        n = len(wb.table_headers(WS, TBL))
        c0 = wb.table_col(WS, TBL, wb.table_headers(WS, TBL)[0])
        bad = [addr(r, c0 + ci) for r in (first, first + 4, 2010) for ci in range(n) if vtype(addr(r, c0 + ci)) != '(none)']
        check(not bad, f"{label}: no validation below the table" + (' - stray on ' + ', '.join(bad) if bad else ''))

    section('as shipped, after setup', first=True)
    check_row(1, 'first row')
    check_row(wb.table_row_count(WS, TBL), 'last row')
    check_below('shipped')

    section('a row added to the table')
    wb.run('SetQuiet', True)
    wb.activate(WS)
    before = wb.table_row_count(WS, TBL)
    wb.run('btnAddPrintJob')
    time.sleep(0.5)
    after = wb.table_row_count(WS, TBL)
    check(after == before + 1, f'row added ({before} -> {after})')
    check_row(after, 'added row')
    check_below('after add')

    section('columns moved with Cut + Insert, then one BindColumns')
    wb.unprotect(WS)
    # Paid in front of Qty, Disregard Paper in front of Student Name - see the .ps1.
    for src, dst in (('Paid', 'Qty'), ('Disregard Paper', 'Student Name')):
        first, last = wb.table_first_row(WS, TBL), wb.table_last_row(WS, TBL)
        s = col_letters(wb.table_col(WS, TBL, src))
        d = col_letters(wb.table_col(WS, TBL, dst))
        wb.prop(rng(WS, f'{s}{first}:{s}{last}'), 'Cut')
        wb.prop(rng(WS, f'{d}{first}:{d}{last}'), 'Insert -4161')
    time.sleep(0.5)
    check(wb.table_col(WS, TBL, 'Paid') < wb.table_col(WS, TBL, 'Qty'), 'Paid now sits before Qty')
    wb.call('BindColumns', sheet(WS))
    time.sleep(0.5)
    check_row(1, 'reordered, first row')
    check_row(wb.table_row_count(WS, TBL), 'reordered, last row')
    check_below('reordered')

    wb.run('SetQuiet', False)

finish()
