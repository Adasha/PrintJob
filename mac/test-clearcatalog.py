# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-clearcatalog.ps1 - same checks, same messages, same order;
# see that script for what each one guards.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, table  # noqa: E402

CASES = [
    {'Sheet': 'Print Technicians', 'Table': 'tblTechnicians', 'Macro': 'btnClearTechnicians'},
    {'Sheet': 'Printers', 'Table': 'tblPrinters', 'Macro': 'btnClearPrinters'},
    {'Sheet': 'Papers', 'Table': 'tblPapers', 'Macro': 'btnClearPapers'},
]

with open_copy() as wb:
    wb.run('SetQuiet', True)

    def rows(t):
        return int(wb.call('RowCount', table(t)))

    def list_rows(c):
        return wb.table_last_row(c['Sheet'], c['Table']) - wb.table_first_row(c['Sheet'], c['Table'])

    def has_button(sheet, macro):
        return any(b['on_action'].endswith(macro) for b in wb.buttons(sheet))

    for i, c in enumerate(CASES):
        section(c['Sheet'], first=(i == 0))
        check(has_button(c['Sheet'], c['Macro']), f"'Clear table' button ({c['Macro']}) is on the sheet")

    for c in CASES:
        section(f"Clear {c['Table']} only")
        before = {o['Table']: rows(o['Table']) for o in CASES}
        check(before[c['Table']] > 0, f"{c['Table']} starts with rows (got {before[c['Table']]})")

        # Quiet mode: Ask declines, so the prompting path must delete nothing.
        wb.run('ClearCatalogTable', c['Table'])
        check(rows(c['Table']) == before[c['Table']], 'declined confirmation deletes nothing')

        wb.run('DoClearCatalogTable', c['Table'])
        check(rows(c['Table']) == 0, 'table reports zero rows after clear')
        check(list_rows(c) == 1, 'one blank templated row is left')
        for o in CASES:
            if o['Table'] != c['Table']:
                check(rows(o['Table']) == before[o['Table']], f"{o['Table']} untouched ({before[o['Table']]} rows)")

        # An empty table clears as a no-op.
        wb.run('DoClearCatalogTable', c['Table'])
        check(list_rows(c) == 1, 'clearing an empty table is a no-op')

        # The blank row is reusable and gets its default Active.
        wb.activate(c['Sheet'])
        wb.run('AddCatalogRow', c['Table'])
        check(list_rows(c) == 1, 'AddCatalogRow reuses the blank row rather than adding a second')
        active = wb.table_col(c['Sheet'], c['Table'], 'Active')
        if active:
            got = wb.cell_text(c['Sheet'], wb.table_first_row(c['Sheet'], c['Table']) + 1, active)
            check(got == 'Yes', 'Active defaults to Yes on the reused row')

    wb.run('SetQuiet', False)

finish()
