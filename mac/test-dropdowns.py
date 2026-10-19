# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-dropdowns.ps1 - same checks, same messages, same order; see
# that script for the catalogue combinations each check relies on.
#
# Like the .ps1, this writes the cell value directly (Range.Value2), which
# fires Worksheet_Change. It does NOT cover picking from the in-cell dropdown
# by hand: on Excel for Mac that runs no VBA at all (CHANGELOG 0.10.25), and
# nothing scripted can reproduce it.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr, col_letters  # noqa: E402

MAIN = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    prn_col = wb.table_col(MAIN, TBL, 'Printer')
    stk_col = wb.table_col(MAIN, TBL, 'Paper Stock')

    def new_row():
        wb.activate(MAIN)
        wb.run('btnAddPrintJob')
        return wb.table_last_row(MAIN, TBL)

    def put(row, col, v):
        wb.set_value(MAIN, addr(row, col), v)
        time.sleep(0.3)

    def get(row, col):
        return wb.cell_text(MAIN, row, col)

    def staged(tag):
        """The list _Work stages for a dropdown: the column headed tag, down to its first blank."""
        heads = wb.texts('_Work', 'A1:GR1')[0]   # 200 columns, as the .ps1 scans
        if tag not in heads:
            return []
        c = col_letters(heads.index(tag) + 1)
        items = []
        for (v,) in wb.texts('_Work', f'{c}2:{c}60'):
            if v == '':
                break
            items.append(v)
        return items

    # ------------------------------------------------------- stock -> printer
    section('Paper Stock chosen first narrows/auto-fills Printer', first=True)
    r1 = new_row()
    put(r1, stk_col, 'Gloss 200gsm SRA3 sheet')
    got = get(r1, prn_col)
    check(got == 'Xerox Versant 180', f"Sheet stock auto-fills the one compatible printer (got '{got}')")

    r2 = new_row()
    put(r2, stk_col, 'Canvas 914mm roll')
    got = get(r2, prn_col)
    check(got == 'Epson SureColor P9500', f"Long Roll stock auto-fills the one compatible printer (got '{got}')")

    r3 = new_row()
    put(r3, stk_col, 'Satin photo 610mm roll')
    got = get(r3, prn_col)
    check(got == '', f"Short Roll stock (2 compatible printers) leaves Printer blank, not auto-filled (got '{got}')")
    s = staged('PRN|Example Print Room|Satin photo 610mm roll')
    joined = ', '.join(s)
    check(len(s) == 3, f'Printer list still shows all 3 permitted printers, not narrowed (got {len(s)}: {joined})')
    check('Xerox Versant 180 (unavailable)' in s,
          f'the incompatible printer (Sheet only) is listed but marked unavailable (got: {joined})')
    check('Epson SureColor P9500' in s, f'the compatible Epson is listed, unmarked (got: {joined})')
    check('HP DesignJet Z9+' in s, f'the compatible HP DesignJet is listed, unmarked (got: {joined})')

    # ------------------------------------------------- picking a marked item
    section('Picking a marked "(unavailable)" dropdown entry re-filters, AT-05 style')
    put(r3, prn_col, 'Xerox Versant 180 (unavailable)')
    prn_after = get(r3, prn_col)
    stk_after = get(r3, stk_col)
    check(prn_after == 'Xerox Versant 180', f"the marker is stripped off the Printer cell itself (got '{prn_after}')")
    check(stk_after == '', f"the now-incompatible Paper Stock is cleared (got '{stk_after}')")
    log = wb.run('QuietLog')
    check('cannot be used on' in log, 'picking the marked entry still raised the AT-05 incompatibility warning')

    # ------------------------------------------------------- printer -> stock
    section('Printer chosen first still narrows/auto-fills Paper Stock')
    r4 = new_row()
    put(r4, prn_col, 'HP DesignJet Z9+')
    got = get(r4, stk_col)
    check(got == '',
          f"HP DesignJet Z9+ (2 compatible stocks: Satin photo roll + Supplied (Roll)) leaves Paper Stock blank (got '{got}')")

    r5 = new_row()
    put(r5, prn_col, 'Xerox Versant 180')
    got = get(r5, stk_col)
    check(got == '', f"Xerox Versant 180 (5 compatible sheet stocks) leaves Paper Stock blank (got '{got}')")

    # --------------------------------------------------------- incompatible
    section('Incompatible combination still caught (AT-05, both directions)')
    r6 = new_row()
    put(r6, prn_col, 'Xerox Versant 180')
    put(r6, stk_col, 'Canvas 914mm roll')   # Long Roll - incompatible with Xerox (Sheet only)
    stk_after = get(r6, stk_col)
    prn_after = get(r6, prn_col)
    check(stk_after == '', f"the just-typed incompatible Paper Stock is cleared (got '{stk_after}')")
    check(prn_after == 'Xerox Versant 180', f"the existing Printer is left alone (got '{prn_after}')")
    log = wb.run('QuietLog')
    check('cannot be used on' in log, 'a warning naming the incompatibility was raised')

    r7 = new_row()
    put(r7, stk_col, 'Canvas 914mm roll')
    put(r7, prn_col, 'Xerox Versant 180')   # Sheet only - incompatible with Canvas (Long Roll)
    stk_after = get(r7, stk_col)
    prn_after = get(r7, prn_col)
    check(stk_after == '', f"changing to an incompatible Printer clears the existing Paper Stock (got '{stk_after}')")
    check(prn_after == 'Xerox Versant 180', f"the just-typed Printer itself is kept (got '{prn_after}')")

    wb.run('SetQuiet', False)

finish()
