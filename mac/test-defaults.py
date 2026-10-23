# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-defaults.ps1 - same checks, same messages, same order; see
# that script for what each section is guarding. Differences:
#
#   - ClearDefaults takes a Worksheet, which cannot cross from AppleScript, so
#     this calls btnClearDefaults (ClearDefaults ActiveSheet) after activating
#     the sheet.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish  # noqa: E402

MAIN = 'Example Print Room'
TBL = 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)

    tech_col = wb.table_col(MAIN, TBL, 'Technician')
    prn_col = wb.table_col(MAIN, TBL, 'Printer')
    stk_col = wb.table_col(MAIN, TBL, 'Paper Stock')

    def new_row():
        wb.activate(MAIN)
        wb.run('btnAddPrintJob')
        return wb.table_last_row(MAIN, TBL)

    def t(a1):
        return wb.text(MAIN, a1)

    # ------------------------------------------------------------ named cells
    section('Named cells exist and are laid out one per row at A3:B5', first=True)
    expected_rows = {'LOC_DefTech': 4, 'LOC_DefPrinter': 5, 'LOC_DefPaper': 6}
    for nm in ('LOC_DefTech', 'LOC_DefPrinter', 'LOC_DefPaper'):
        r = wb.name_ref(MAIN, nm)
        check(r is not None and r[0] == expected_rows[nm], f'{nm} exists and refers to row {expected_rows[nm]}')
    check(t('A4') == 'Default: technician', f"A4 label reads correctly (got '{t('A4')}')")

    # Disregard-cost defaults sit under Default: paper (2026-09-29), row 8 is a blank spacer,
    # and the settings below moved down to rows 9-11.
    layout = {'LOC_DefDisPaper': 8, 'LOC_DefDisCons': 9, 'LOC_RollUnit': 11}
    for nm, row in layout.items():
        r = wb.name_ref(MAIN, nm)
        check(r is not None and r == (row, 2), f'{nm} refers to B{row}')
    check(t('A8') == 'Disregard paper', f"A8 label reads correctly (got '{t('A8')}')")
    check(t('A9') == 'Disregard consumable', f"A9 label reads correctly (got '{t('A9')}')")
    check(t('A10') == '' and t('B10') == '', 'Row 10 is a blank spacer')
    check(t('A11') == 'Roll length unit', f"A11 label reads correctly (got '{t('A11')}')")
    check(t('A13') == 'Print jobs', f"A13 label reads correctly (got '{t('A13')}')")

    # ------------------------------------------------------------- pre-fill
    section('Setting defaults pre-fills subsequently added jobs')
    t_cell = wb.name_addr(MAIN, 'LOC_DefTech')
    wb.set_value(MAIN, t_cell, 'A. Shailer')
    p_cell = wb.name_addr(MAIN, 'LOC_DefPrinter')
    wb.set_value(MAIN, p_cell, 'Xerox Versant 180')
    time.sleep(0.3)
    # Xerox is Sheet-only -> Paper should auto-fill only if singleton; Main has
    # 5 sheet stocks, so it should stay blank, not auto-filled.
    s_cell = wb.name_addr(MAIN, 'LOC_DefPaper')
    check(t(s_cell) == '', 'Default Paper stays blank (5 compatible sheet stocks, not a singleton)')
    wb.set_value(MAIN, s_cell, 'Gloss 200gsm SRA3 sheet')
    time.sleep(0.3)

    r1 = new_row()
    got_tech = wb.cell_text(MAIN, r1, tech_col)
    got_prn = wb.cell_text(MAIN, r1, prn_col)
    got_stk = wb.cell_text(MAIN, r1, stk_col)
    check(got_tech == 'A. Shailer', f"new job's Technician pre-filled (got '{got_tech}')")
    check(got_prn == 'Xerox Versant 180', f"new job's Printer pre-filled (got '{got_prn}')")
    check(got_stk == 'Gloss 200gsm SRA3 sheet', f"new job's Paper Stock pre-filled (got '{got_stk}')")
    got_unit = wb.cell_text(MAIN, r1, wb.table_col(MAIN, TBL, 'Unit'))
    check(got_unit == 'sheets', f"new job's Unit fills from the pre-filled Paper Stock (got '{got_unit}')")

    r2 = new_row()
    got_tech2 = wb.cell_text(MAIN, r2, tech_col)
    check(got_tech2 == 'A. Shailer', f"a second job also pre-fills from the same defaults (got '{got_tech2}')")

    # ---------------------------------------------- changing default is safe
    section('Changing a default afterwards leaves earlier jobs alone (AT-07/AT-08)')
    wb.set_value(MAIN, t_cell, 'J. Okonkwo')
    time.sleep(0.3)
    still_a = wb.cell_text(MAIN, r1, tech_col)
    check(still_a == 'A. Shailer', f"row 1's Technician is untouched by the later default change (got '{still_a}')")

    # ------------------------------------------------------- autofill/narrow
    section('Defaults implement the same bidirectional filtering as table cells (1a)')
    wb.activate(MAIN)
    wb.run('btnClearDefaults')
    time.sleep(0.2)
    wb.set_value(MAIN, s_cell, 'Canvas 914mm roll')   # Long Roll -> only Epson fits at this location
    time.sleep(0.3)
    got_prn3 = t(p_cell)
    check(got_prn3 == 'Epson SureColor P9500',
          f"Long Roll default paper auto-fills the one compatible default printer (got '{got_prn3}')")

    # ------------------------------------------------------------------- clear
    section('Clear defaults empties all three')
    wb.activate(MAIN)
    wb.run('btnClearDefaults')
    time.sleep(0.3)
    all_blank = t(t_cell) == '' and t(p_cell) == '' and t(s_cell) == ''
    check(all_blank, 'all three default cells are blank after Clear defaults')

    r3 = new_row()
    got_tech3 = wb.cell_text(MAIN, r3, tech_col)
    check(got_tech3 == '', f"a job added after Clear defaults gets no pre-fill (got '{got_tech3}')")

    wb.run('SetQuiet', False)

finish()
