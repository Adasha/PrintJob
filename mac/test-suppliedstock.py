# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-suppliedstock.ps1 - same checks, same messages, same order;
# see that script for the "Supplied by student" stocks and the catalogue data
# each check relies on.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, addr  # noqa: E402

MAIN, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'
PAPERS, PTBL = 'Papers', 'tblPapers'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    def col(name, sheet=MAIN, table=TBL):
        c = wb.table_col(sheet, table, name)
        if c == 0:
            raise RuntimeError(f"column '{name}' not found in table '{table}' (columns: {', '.join(wb.table_headers(sheet, table))})")
        return c

    prn_col, stk_col, qty_col = col('Printer'), col('Paper Stock'), col('Qty')
    width_col, size_col, area_col = col('Print Width mm'), col('Sheet size'), col('Area m2')
    paper_cost_col, cons_cost_col, status_col = col('Paper Cost'), col('Consumable Cost'), col('Status')

    def new_row():
        wb.activate(MAIN)
        wb.run('btnAddPrintJob')
        return wb.table_last_row(MAIN, TBL)

    def put(row, c, v):
        wb.set_value(MAIN, addr(row, c), v)

    def num(row, c):
        v = wb.value(MAIN, addr(row, c))
        return 0.0 if v in (None, '') else float(v)

    def txt(row, c):
        v = wb.value(MAIN, addr(row, c))
        return '' if v is None else str(v)

    def log():
        return wb.run('QuietLog')

    # ---------------------------------------------------- roll, zero paper cost
    section('Supplied (Roll): zero paper cost, normal consumable cost', first=True)
    r1 = new_row()
    put(r1, prn_col, 'Epson SureColor P9500')
    put(r1, stk_col, 'Supplied (Roll)')
    time.sleep(0.3)
    put(r1, width_col, 900)
    put(r1, qty_col, 5)
    time.sleep(0.3)
    paper_cost, cons_cost, area = num(r1, paper_cost_col), num(r1, cons_cost_col), num(r1, area_col)
    check(paper_cost == 0, f'Paper Cost is 0 (got {paper_cost})')
    check(cons_cost > 0, f'Consumable Cost is charged normally (got {cons_cost})')
    check(abs(area - 0.9 * 5) < 0.001, f'Area m2 uses the entered Print Width, not a catalogue width (got {area})')

    # ---------------------------------------------------------- roll, required
    section('Supplied (Roll): Print Width mm required')
    r2 = new_row()
    put(r2, prn_col, 'Epson SureColor P9500')
    put(r2, stk_col, 'Supplied (Roll)')
    put(r2, qty_col, 3)
    time.sleep(0.3)
    status = txt(r2, status_col)
    check('Print width required for student-supplied roll stock' in status, f"Status flags the missing width (got '{status}')")

    # ----------------------------------------------- roll, exceeds printer max
    section('Supplied (Roll): width exceeding the printer max is rejected')
    r3 = new_row()
    put(r3, prn_col, 'HP DesignJet Z9+')   # Max roll width 610mm
    put(r3, stk_col, 'Supplied (Roll)')
    time.sleep(0.3)
    put(r3, width_col, 900)
    time.sleep(0.3)
    width_after = txt(r3, width_col)
    check(width_after == '', f"900mm on a 610mm-max printer is cleared, not accepted (got '{width_after}')")
    check('maximum roll width' in log(), "a warning naming the printer's maximum roll width was raised")

    # --------------------------------------------------- sheet, zero paper cost
    section('Supplied (Sheet): zero paper cost, normal consumable cost')
    r4 = new_row()
    put(r4, prn_col, 'Xerox Versant 180')   # Max sheet size A1
    put(r4, stk_col, 'Supplied (Sheet)')
    time.sleep(0.3)
    put(r4, size_col, 'A3')
    put(r4, qty_col, 10)
    time.sleep(0.3)
    paper_cost4, cons_cost4 = num(r4, paper_cost_col), num(r4, cons_cost_col)
    check(paper_cost4 == 0, f'Paper Cost is 0 (got {paper_cost4})')
    check(cons_cost4 > 0, f'Consumable Cost is charged normally (got {cons_cost4})')

    # -------------------------------------------------------- sheet, required
    section('Supplied (Sheet): Sheet size required')
    r5 = new_row()
    put(r5, prn_col, 'Xerox Versant 180')
    put(r5, stk_col, 'Supplied (Sheet)')
    put(r5, qty_col, 4)
    time.sleep(0.3)
    status5 = txt(r5, status_col)
    check('Sheet size required for student-supplied sheet stock' in status5, f"Status flags the missing size (got '{status5}')")

    # -------------------------------------- sheet, exceeds a smaller printer max
    section('Supplied (Sheet): size exceeding a smaller printer max is rejected')
    max_size_col = col('Max sheet size', 'Printers', 'tblPrinters')
    p_first = wb.table_first_row('Printers', 'tblPrinters')
    p_col2 = wb.table_col('Printers', 'tblPrinters', wb.table_headers('Printers', 'tblPrinters')[1])
    models = [r[0] for r in wb.texts('Printers', f'{addr(p_first + 1, p_col2)}:{addr(p_first + wb.table_row_count("Printers", "tblPrinters"), p_col2)}')]
    xerox_row = p_first + 1 + models.index('Xerox Versant 180')
    wb.unprotect('Printers')
    wb.set_value('Printers', addr(xerox_row, max_size_col), 'SRA3')
    wb.protect('Printers')
    wb.run('Invalidate')

    r6 = new_row()
    put(r6, prn_col, 'Xerox Versant 180')
    put(r6, stk_col, 'Supplied (Sheet)')
    time.sleep(0.3)
    put(r6, size_col, 'A1')   # bigger than the now-reduced SRA3 max
    time.sleep(0.3)
    size_after = txt(r6, size_col)
    check(size_after == '', f"A1 on an SRA3-max printer is cleared, not accepted (got '{size_after}')")
    check('maximum sheet size' in log(), "a warning naming the printer's maximum sheet size was raised")

    # ------------------------------------------------------- Disregard Consumable still works
    section('Disregard Consumable still waives ink cost on a supplied-stock row')
    put(r1, col('Disregard Consumable'), 'Yes')
    time.sleep(0.3)
    chargeable1 = num(r1, col('Chargeable Cost'))
    check(chargeable1 == 0, f'Chargeable Cost drops to 0 once Disregard Consumable is set (got {chargeable1}, Paper Cost was already 0)')

    # ------------------------------------- built in, not rows of tblPapers
    section('Supplied (Roll)/(Sheet) are built in, not tblPapers rows')

    def pap_row(n):
        return wb.table_first_row(PAPERS, PTBL) + n

    def set_pap(n, name, v):
        wb.set_value(PAPERS, addr(pap_row(n), col(name, PAPERS, PTBL)), float(v) if isinstance(v, int) else v)

    def get_pap(n, name):
        v = wb.value(PAPERS, addr(pap_row(n), col(name, PAPERS, PTBL)))
        return '' if v is None else v

    def desc_count(desc):
        return wb.table_column_texts(PAPERS, PTBL, 'Description').count(desc)

    def compatible(model, desc):
        return wb.run('Compatible', model, desc).lower() == 'true'

    # The shipped template has no such rows (PrintCosts.xlsx), so a build has none.
    check(desc_count('Supplied (Roll)') == 0, 'no Supplied (Roll) row in tblPapers')
    check(desc_count('Supplied (Sheet)') == 0, 'no Supplied (Sheet) row in tblPapers')
    check(compatible('Epson SureColor P9500', 'Supplied (Roll)'), 'Supplied (Roll) is still usable on a roll printer')
    check(not compatible('HP DesignJet Z9+', 'Supplied (Sheet)'), 'Supplied (Sheet) is refused by a printer with no sheet capacity')

    section('Typing a reserved name into the Papers table is rejected')
    wb.activate(PAPERS)
    wb.run('AddCatalogRow', PTBL)
    n = wb.table_row_count(PAPERS, PTBL)
    check(str(get_pap(n, 'Active')) == 'Yes', 'new row defaults Active to Yes')
    sbs = str(get_pap(n, 'Supplied by student'))
    check(sbs == 'No', f"Add Row defaults Supplied by student to No (got '{sbs}')")
    wb.unprotect(PAPERS)
    set_pap(n, 'Description', 'Supplied (Sheet)')
    time.sleep(0.3)
    check(str(get_pap(n, 'Description')) == '', 'the reserved name is cleared, not accepted')
    check('built-in paper stock' in log(), 'a warning explained why')

    # ------------------------------------------- Supplied by student = Yes rows
    section('A new stock defaults to Supplied by student = No')
    wb.unprotect(PAPERS)
    set_pap(n, 'Description', 'Bulk test roll')
    time.sleep(0.3)
    sbs = str(get_pap(n, 'Supplied by student'))
    check(sbs == 'No', f"Supplied by student defaults to No (got '{sbs}')")
    wb.unprotect(PAPERS)
    set_pap(n, 'Measure', 'Roll')
    set_pap(n, 'Size mode', 'Roll')
    set_pap(n, 'Width mm', 610)
    set_pap(n, 'Cost', 7)
    time.sleep(0.3)

    def stock_job(printer, row=None):
        r = new_row()
        put(r, prn_col, printer)
        put(r, stk_col, 'Bulk test roll')
        put(r, qty_col, 4)
        time.sleep(0.3)
        return r

    section('Supplied by student = No: paper is charged')
    r7 = stock_job('Epson SureColor P9500')
    paper7 = num(r7, paper_cost_col)
    check(paper7 > 0, f'Paper Cost is charged for an ordinary stock (got {paper7})')

    section('Supplied by student = Yes: cost forced to zero')
    wb.unprotect(PAPERS)
    set_pap(n, 'Supplied by student', 'Yes')
    time.sleep(0.3)
    cost = get_pap(n, 'Cost')
    check(float(cost or 0) == 0, f'the Cost cell is reset to 0 (got {cost})')
    check('not charged for' in log(), 'a warning explained why')
    r8 = stock_job('Epson SureColor P9500')
    paper8, cons8 = num(r8, paper_cost_col), num(r8, cons_cost_col)
    check(paper8 == 0, f'Paper Cost is 0 (got {paper8})')
    check(cons8 > 0, f'Consumable Cost is charged normally (got {cons8})')

    section('The flag decides, not the Cost cell (a cost written behind the events)')
    wb.unprotect(PAPERS)
    wb.set_enable_events(False)
    set_pap(n, 'Cost', 7)
    wb.set_enable_events(True)
    wb.run('Invalidate')
    r9 = stock_job('Epson SureColor P9500')
    paper9 = num(r9, paper_cost_col)
    check(paper9 == 0, f'Paper Cost is still 0 with Cost = 7 and Supplied by student = Yes (got {paper9})')

    section('Such a stock keeps its catalogue size (no per-job size)')
    r10 = new_row()
    put(r10, prn_col, 'HP DesignJet Z9+')   # Max roll width 610mm - fits exactly
    put(r10, stk_col, 'Bulk test roll')
    time.sleep(0.3)
    check(txt(r10, stk_col) == 'Bulk test roll', 'a 610mm bulk roll is accepted on a 610mm-max printer')
    status10 = txt(r10, status_col)
    check('Print width required' not in status10, f"Print width is NOT demanded for it (status '{status10}')")
    wb.run('SetQuiet', False)

finish()
