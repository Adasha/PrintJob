# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Mac port of test-settingsrows.ps1 - same checks, same messages, same order.
# Differences:
#
#   - A border's LineStyle comes back as Excel's enum name: 'line style none'
#     is COM's -4142 and 'continuous' its 1. Colours are converted to COM's
#     BGR integer (xlmac.cell_format), so 14606046 and 16777215 still apply.
#
# Drives a COPY in a temp folder, never src/PrintJob.xlsm itself. Closes
# WITHOUT saving.

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlmac import open_copy, check, finish, addr  # noqa: E402

SET = 'Settings'
TBL = 'tblSettings'


def row_of(keys, key):
    return keys.index(key) + 1 if key in keys else 0


def check_order(wb, tag):
    keys = wb.table_column_texts(SET, TBL, 'Key')
    red = row_of(keys, 'LOC_REDUCED_COLUMNS')
    mn = row_of(keys, 'LOC_MINIMAL_COLUMNS')
    check(red > 0 and mn > 0, f'{tag}: both rows exist (reduced={red} minimal={mn})')
    check(mn == red + 1, f'{tag}: LOC_MINIMAL_COLUMNS is directly below LOC_REDUCED_COLUMNS (rows {red}, {mn})')
    _, cell = wb.wb_name('SET_LOC_MINIMAL_COLUMNS')
    name_row = int(''.join(c for c in cell if c.isdigit()))
    check(name_row == wb.table_first_row(SET, TBL) + mn, f'{tag}: SET_LOC_MINIMAL_COLUMNS points at the moved row')
    check(len(keys) == len(set(keys)), f'{tag}: no duplicate keys')


def check_style(wb, tag):
    keys = wb.table_column_texts(SET, TBL, 'Key')
    notes = wb.table_column_texts(SET, TBL, 'Notes')
    vc = wb.table_col(SET, TBL, 'Value')
    first = wb.table_first_row(SET, TBL) + 1
    locked = opened = 0
    for i, key in enumerate(keys):
        v = wb.cell_format(SET, addr(first + i, vc))
        is_ro = notes[i].strip().startswith('Read-only')
        check(v['locked'] == is_ro, f'{tag}: {key} lock state follows its Notes')
        if is_ro:
            locked += 1
            check(v['color'] == 14606046 and v['italic'] is True, f'{tag}: {key} (locked) is grey italic')
            check(v['left_border'] == 'line style none', f'{tag}: {key} (locked) has no border')
        else:
            opened += 1
            check(v['color'] == 16777215 and v['italic'] is False, f'{tag}: {key} (editable) is white, upright')
            check(v['left_border'] == 'continuous', f'{tag}: {key} (editable) has a border')
    check(locked > 0 and opened > 0, f'{tag}: table has both kinds ({locked} locked, {opened} editable)')


with open_copy() as wb:
    wb.run('SetQuiet', True)
    check_order(wb, 'built')
    check_style(wb, 'built')

    wb.run('InitialiseWorkbook')
    time.sleep(0.5)
    check_order(wb, 'second setup')
    check_style(wb, 'second setup')

finish()
