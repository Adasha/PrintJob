# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-newpaperrow.ps1 - same checks, same messages, same order;
# see that script for the bug it guards (a new Papers row never reaching the
# print rooms' Paper Stock dropdowns).
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr  # noqa: E402

PAPERS, PTBL = 'Papers', 'tblPapers'
MAIN, JTBL = 'Example Print Room', 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    def pcol(name):
        c = wb.table_col(PAPERS, PTBL, name)
        if c == 0:
            raise RuntimeError(f"column '{name}' not found")
        return c

    def prow(n):
        """Worksheet row of tblPapers' data row n."""
        return wb.table_first_row(PAPERS, PTBL) + n

    def put(n, col, v):
        wb.set_value(PAPERS, addr(prow(n), pcol(col)), v)

    def val(n, col):
        return wb.value(PAPERS, addr(prow(n), pcol(col)))

    def num(n, col):
        v = val(n, col)
        return 0.0 if v is None else float(v)

    def stock_items(row, col):
        """The items a Paper Stock cell's dropdown offers, read back through
        its validation source on the _Work staging sheet."""
        src = wb.validation_formula1(MAIN, addr(row, col))
        m = re.match(r"^='?_Work'?!(.+)$", src)
        if not m:
            raise RuntimeError(f"unexpected validation source '{src}'")
        return [c for r in wb.texts('_Work', m.group(1).replace('$', '')) for c in r]

    def new_paper(desc, measure, mode, std, w, h):
        wb.activate(PAPERS)
        wb.run('AddCatalogRow', PTBL)
        n = wb.table_last_row(PAPERS, PTBL) - wb.table_first_row(PAPERS, PTBL)
        put(n, 'StockID', f'STK-TEST-{n}')
        put(n, 'Description', desc)
        put(n, 'Paper type', 'Matte')
        put(n, 'Measure', measure)
        put(n, 'Size mode', mode)
        if std:
            put(n, 'Std. size', std)
        if w is not None:
            put(n, 'Width mm', float(w))
        if h is not None:
            put(n, 'Height mm', float(h))
        put(n, 'Cost', 0.5)
        time.sleep(0.3)
        return n

    def job_stock_cell():
        wb.activate(MAIN)
        wb.run('btnAddPrintJob')
        return wb.table_last_row(MAIN, JTBL), wb.table_col(MAIN, JTBL, 'Paper Stock')

    def compatible(model, desc):
        # A Boolean through AppleScript comes back as 'true'/'false'.
        return wb.run('Compatible', model, desc).lower() == 'true'

    # ------------------------------------------------ standard-size sheet stock
    section('New sheet stock, Size mode Standard, Std. size A3, no width/height typed', first=True)
    desc = 'TEST Matte 150gsm A3 sheet'
    row = new_paper(desc, 'Sheet', 'Standard', 'A3', None, None)
    w, h = num(row, 'Width mm'), num(row, 'Height mm')
    active = str(val(row, 'Active') or '')
    check(w == 297 and h == 420, f'Width/Height filled from A3 (got {w:g} x {h:g})')
    check(active == 'Yes', f"Active defaulted to Yes (got '{active}')")
    check(compatible('Xerox Versant 180', desc), 'compatible with Xerox Versant 180 (max sheet A1)')
    check(not compatible('HP DesignJet Z9+', desc), 'not compatible with HP DesignJet Z9+ (roll only)')

    section('It reaches the print room dropdown after leaving Papers')
    items = stock_items(*job_stock_cell())   # activating the print room rebinds every dropdown
    check(desc in items, 'new stock is listed in Paper Stock')

    # -------------------------------------------- changing the standard size
    section('Changing Std. size re-fills Width/Height')
    wb.activate(PAPERS)
    put(row, 'Std. size', 'A2')
    time.sleep(0.3)
    w2, h2 = num(row, 'Width mm'), num(row, 'Height mm')
    check(w2 == 420 and h2 == 594, f'Width/Height follow the new size (got {w2:g} x {h2:g})')

    # ---------------------------------- a hand-typed custom size is not touched
    section('Custom sheet: typed Width/Height are kept')
    desc2 = 'TEST Custom 200x300 sheet'
    row2 = new_paper(desc2, 'Sheet', 'Custom sheet', None, 200, 300)
    check(num(row2, 'Width mm') == 200 and num(row2, 'Height mm') == 300, 'typed size unchanged')
    check(compatible('Xerox Versant 180', desc2), 'compatible with Xerox Versant 180')

    # ------------------------------------------------ explicit Active = No kept
    section('An explicit Active = No is never overwritten')
    put(row2, 'Active', 'No')
    time.sleep(0.3)
    put(row2, 'Description', desc2 + ' x')
    time.sleep(0.3)
    check(val(row2, 'Active') == 'No', 'Active stays No after a further edit')

    wb.run('SetQuiet', False)

finish()
