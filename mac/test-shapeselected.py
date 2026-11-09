# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-shapeselected.ps1 - same checks, same messages, same order;
# see that script for the history (0.12.2: a button left selected).
#
# Selecting the shape goes through modScript.ScriptProp ("Select" is called as
# a method on the Form Controls button). Drives a COPY in a temp folder, never
# src/PrintJob.xlsm itself. Closes WITHOUT saving.

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, section, finish, sheet, table, shape, ExcelError  # noqa: E402

WS, TBL = 'Example Print Room', 'tblJobs_EXAMPLE'

with open_copy() as wb:
    wb.run('SetQuiet', True)
    wb.activate(WS)
    btn = wb.buttons(WS)[0]['name']

    def select_button():
        # Select returns Null, which ScriptProp cannot turn into text (error 94):
        # the shape is selected all the same, and the check below proves it.
        try:
            wb.prop(shape(WS, btn), 'Select')
        except ExcelError as e:
            if 'Invalid use of Null' not in str(e):
                raise

    def selected_type():
        return wb._do('return (class of selection) as text')

    section('A button is selected')
    select_button()
    check('range' not in selected_type().lower(), 'the button really is the selection')

    before = wb.table_row_count(WS, TBL)
    stamp = wb.table_column_texts(WS, TBL, 'Date/Time')
    lo = table(TBL)

    check(wb.call('SelectedRow', sheet(WS), lo) == '0', 'SelectedRow answers 0 instead of raising 438')
    check(wb.call('SelectedTableRow', sheet(WS), lo) == '0', 'SelectedTableRow answers 0')

    for name, args in (('StampNow', (sheet(WS),)), ('RemoveRow', (sheet(WS),)), ('RepeatJob', (sheet(WS),)),
                       ('AddPass', (sheet(WS),)), ('RemovePass', (sheet(WS),)),
                       ('RemoveCatalogRow', ('tblColours',))):
        select_button()
        try:
            wb.call(name, *args)
            ok = True
        except Exception as e:  # noqa: BLE001
            ok = False
            print(f'        {e}')
        check(ok, f'{name} with a button selected raises no error')

    check(wb.table_row_count(WS, TBL) == before, 'no print job row was added or removed')
    check(wb.table_column_texts(WS, TBL, 'Date/Time') == stamp, 'no Date/Time was stamped')
    check(wb.table_row_count('Consumables', 'tblColours') > 0, 'no consumable was removed')

    wb.run('SetQuiet', False)

finish()
